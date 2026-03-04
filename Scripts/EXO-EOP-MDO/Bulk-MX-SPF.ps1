############################################################################
#This sample script is not supported under any Microsoft standard support program or service.
#This sample script is provided AS IS without warranty of any kind.
#Microsoft further disclaims all implied warranties including, without limitation, any implied
#warranties of merchantability or of fitness for a particular purpose. The entire risk arising
#out of the use or performance of the sample script and documentation remains with you. In no
#event shall Microsoft, its authors, or anyone else involved in the creation, production, or
#delivery of the scripts be liable for any damages whatsoever (including, without limitation,
#damages for loss of business profits, business interruption, loss of business information,
#or other pecuniary loss) arising out of the use of or inability to use the sample script or
#documentation, even if Microsoft has been advised of the possibility of such damages.
############################################################################

<#
.SYNOPSIS
    Bulk domain health checker — MX, SPF, DKIM, DMARC, and MTA-STS analysis.

.DESCRIPTION
    Reads a CSV of domains, resolves DNS records, checks email authentication
    posture, and generates a Bootstrap-styled HTML report.

.PARAMETER CsvPath
    Path to the input CSV file containing domain names.

.PARAMETER OutputFolder
    Folder where the HTML report will be saved. Created automatically if missing.

.AUTHOR
    Ernesto Cobos Roqueñí

.VERSION
    5.3.4 - 02/18/2026
#>

#Requires -RunAsAdministrator

param(
    [Parameter(Mandatory = $false)]
    [string]$CsvPath = "C:\Scripts\MDO\AcceptedDomains.csv",

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = "C:\Scripts\MDO"
)

# 1. Clear DNS Cache
Clear-DnsClientCache

# 2. Path Settings
if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}
$reportPath = Join-Path -Path $OutputFolder -ChildPath "Bulk_HealthCheck_$(Get-Date -Format 'MMddyy_HHmm').html"

# 3. Module Verification
Write-Host "--- Reviewing Requirements ---" -f Yellow
$modules = @('DomainHealthChecker', 'EmailAuthChecker')
foreach ($mod in $modules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Installing $mod..." -f Cyan
        Install-Module $mod -Force -Confirm:$false -Scope CurrentUser
    }
    if (-not (Import-Module $mod -ErrorAction SilentlyContinue -PassThru)) {
        Write-Host "  WARNING: Could not import $mod. Some checks may be skipped." -f Yellow
    }
}

# 4. Import and Initialization
if (-not (Test-Path $CsvPath)) {
    Write-Host "CRITICAL ERROR: CSV file not found at $CsvPath" -f Red
    return
}
$firstLine = Get-Content $CsvPath -TotalCount 1
$detectedDelimiter = ","
if ($firstLine -like "*;*") { $detectedDelimiter = ";" }
$csvData = Import-Csv $CsvPath -Delimiter $detectedDelimiter

$results = [System.Collections.Generic.List[PSObject]]::new()
$totalDomains = 0
$spfIssues = 0 ; $dkimIssues = 0 ; $dmarcIssues = 0 ; $mtaIssues = 0

function Get-TtlClass {
    param($ttlValue)
    if ([string]::IsNullOrWhiteSpace($ttlValue) -or $ttlValue -eq "N/A") { return "bg-danger text-white px-2 rounded fw-bold" }
    $val = 0
    if ([int]::TryParse($ttlValue, [ref]$val)) {
        if ($val -eq 3600) { return "bg-success text-white px-2 rounded fw-bold" }
    }
    return "bg-danger text-white px-2 rounded fw-bold"
}

# 5. Processing
foreach ($row in $csvData) {
    $domain = ""
    foreach ($prop in $row.psobject.Properties) {
        if ($prop.Value -like "*.*") { $domain = [string]$prop.Value ; break }
    }
    $domain = [string]$domain.Trim()
    if ([string]::IsNullOrWhiteSpace($domain) -or $domain.Length -lt 4) { continue }
    $totalDomains++
    Write-Host "Analyzing: $domain" -f Cyan
    
    try {
        $DHC = Invoke-SpfDkimDmarc -Name $domain -ErrorAction Stop
        
        # --- MX Record ---
        $MXRecords = @(Resolve-DnsName -Name $domain -Type MX -ErrorAction SilentlyContinue)
        $MXList = @()
        foreach ($record in $MXRecords) {
            $rawTTL = 0
            if ($record.psobject.Properties.Name -contains "TTL") { $rawTTL = $record.TTL }
            elseif ($record.psobject.Properties.Name -contains "TimeToLive") { $rawTTL = $record.TimeToLive }
            
            $MXList += [PSCustomObject]@{
                Hostname = $record.NameExchange.TrimEnd('.')
                Pref     = $record.Preference
                TTL      = "$rawTTL"
                TClass   = Get-TtlClass $rawTTL
            }
        }

        # --- SPF (String error safe) ---
        $spfTTLVal = "N/A"
        $txtRecords = @(Resolve-DnsName -Name $domain -Type TXT -ErrorAction SilentlyContinue)
        foreach ($txt in $txtRecords) {
            if ($txt.psobject.Properties.Name -contains "Strings") {
                $content = $txt.Strings -join ''
                if ($content -match '^\s*v=spf1\b') { $spfTTLVal = $txt.TTL ; break }
            }
        }
        
        # SPF Conditional Formatting
        $spfLenClass = if ([int]$DHC.SPFRecordLength -gt 255) { "bg-danger text-white" } else { "bg-success text-white" }
        $lookupText = [string]$DHC.SPFRecordDnsLookupCount
        $lookupClass = "bg-secondary text-white"
        if ($lookupText -match "OK") { $lookupClass = "bg-success text-white" }
        if ($lookupText -match "but maximum DNS Lookups reached!") { $lookupClass = "bg-danger text-white" }

        # --- DKIM (String error safe) ---
        $dkimResults = @()
        $selectors = @("selector1","selector2","google","google1","google2","s1","s2","key1","key2","k1","k2","eversrv","mxvault","dkim","default")
        foreach ($s in $selectors) {
            $dkRecords = @(Resolve-DnsName -Name "$s._domainkey.$domain" -Type ANY -ErrorAction SilentlyContinue)
            foreach ($dk in $dkRecords) {
                if ($dk.Type -in @("CNAME","TXT")) {
                    $rawKey = "N/A"
                    if ($dk.psobject.Properties.Name -contains "Strings") { $rawKey = $dk.Strings -join "" }
                    elseif ($dk.psobject.Properties.Name -contains "NameHost") { $rawKey = $dk.NameHost }
                    
                    $dkimResults += [PSCustomObject]@{
                        Selector = $s
                        TTL      = $dk.TTL
                        TClass   = Get-TtlClass $dk.TTL
                        Key      = $rawKey
                    }
                }
            }
        }

        # --- MTA-STS ---
        $mtaAuth = $null
        $mtaInfo = [PSCustomObject]@{ DnsTtl="N/A"; Version="N/A"; Mode="N/A"; MaxAge="N/A"; MX="N/A" }
        if (Get-Command -Name Get-MailPolicyAuth -ErrorAction SilentlyContinue) {
            $mtaAuth = Get-MailPolicyAuth -Domain $domain -ErrorAction SilentlyContinue
        } else {
            Write-Host "  [!] Get-MailPolicyAuth not available. Skipping MTA-STS for $domain" -f Yellow
        }
        if ($null -ne $mtaAuth -and $null -ne $mtaAuth.MTA_STS) {
            if($mtaAuth.MTA_STS.DnsTtl){$mtaInfo.DnsTtl = $mtaAuth.MTA_STS.DnsTtl}
            if($mtaAuth.MTA_STS.Version){$mtaInfo.Version = $mtaAuth.MTA_STS.Version}
            if($mtaAuth.MTA_STS.Mode){$mtaInfo.Mode = $mtaAuth.MTA_STS.Mode}
            if($mtaAuth.MTA_STS.Max_Age){$mtaInfo.MaxAge = $mtaAuth.MTA_STS.Max_Age}
            if($mtaAuth.MTA_STS.MX){$mtaInfo.MX = $mtaAuth.MTA_STS.MX -join ", "}
        }

        # --- Dashboard ---
        $failKeys = @("not sufficiently strict","missing","not found","characters","prevent abuse","doesn't exist","couldn't find","does not have")
        function Check-IsFail([string]$val) {
            if ([string]::IsNullOrWhiteSpace($val) -or $val -eq "Not found") { return $true }
            foreach ($k in $failKeys) { if ($val -match $k) { return $true } }
            return $false
        }

        $spfSt = if (Check-IsFail $DHC.SpfAdvisory) { "bg-danger text-white" } else { "bg-success text-white" }
        $dkSt  = if (Check-IsFail $DHC.DkimAdvisory) { "bg-danger text-white" } else { "bg-success text-white" }
        $dmSt  = if (Check-IsFail $DHC.DmarcAdvisory) { "bg-danger text-white" } else { "bg-success text-white" }
        $mtSt  = if (Check-IsFail $DHC.MtaAdvisory) { "bg-danger text-white" } else { "bg-success text-white" }

        if ($spfSt -match "danger") { $spfIssues++ }
        if ($dkSt -match "danger")  { $dkimIssues++ }
        if ($dmSt -match "danger")  { $dmarcIssues++ }
        if ($mtSt -match "danger")  { $mtaIssues++ }

        $results += [PSCustomObject]@{
            Domain = $domain ; MXData = $MXList ; SpfRecord = if($DHC.SpfRecord){$DHC.SpfRecord}else{"Not found"}
            SpfAdvisory = $DHC.SpfAdvisory ; SpfStatusClass = $spfSt ; SpfTTL = $spfTTLVal ; SpfTtlClass = Get-TtlClass $spfTTLVal
            SpfLen = $DHC.SPFRecordLength ; SpfLenClass = $spfLenClass ; SpfLookups = $lookupText ; SpfLookupClass = $lookupClass
            DkimDetails = $dkimResults ; DkimAdvisory = $DHC.DkimAdvisory ; DkimStatusClass = $dkSt
            DmarcRecord = if($DHC.DmarcRecord){$DHC.DmarcRecord}else{"Not found"} ; DmarcAdvisory = $DHC.DmarcAdvisory ; DmarcStatusClass = $dmSt
            DmarcTTL = if(@(Resolve-DnsName -Name "_dmarc.$domain" -Type TXT -ErrorAction SilentlyContinue).Count -gt 0){(Resolve-DnsName -Name "_dmarc.$domain" -Type TXT)[0].TTL}else{"N/A"}
            MtaAdvisory = $DHC.MtaAdvisory ; MtaStatusClass = $mtSt ; MtaInfo = $mtaInfo
        }
    } catch { Write-Host "  [!] Error on $domain : $($_.Exception.Message)" -f Red }
}

# 6. HTML CONSTRUCTION
$date = Get-Date -Format "MM/dd/yyyy HH:mm"
$htmlHeader = @"
<!DOCTYPE html>
<html lang="en">
<head>
   <meta charset="UTF-8">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f4f7f9; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .hero { background-color: #0078d4; color: white; padding: 30px; margin-bottom: 30px; border-bottom: 4px solid #005a9e; }
        .logo-img { max-height: 35px; filter: brightness(0) invert(1); }        
        .dashboard-card { border-radius: 12px; border: none; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
        .stat-label { font-size: 1.1rem; font-weight: bold; text-transform: uppercase; opacity: 0.9; }
        .stat-number { font-size: 2.8rem; font-weight: 800; line-height: 1; }
        .table-card { background: white; border: 1px solid #e1e4e8; border-radius: 8px; padding: 25px; margin-bottom: 35px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        .tech-label { font-weight: bold; color: #333; width: 220px; display: inline-block; font-size: 1.1rem; }
        .record-box { background-color: #f8f9fa; border: 1px solid #ddd; padding: 10px; border-radius: 4px; display: block; margin-top: 5px; word-break: break-all; font-family: 'Consolas', monospace; }
        .badge-advisory { display: block; padding: 10px; border-radius: 5px; margin-top: 5px; font-weight: 500; font-family: 'Consolas', monospace; }
        .section-divider { border-bottom: 2px solid #0078d4; color: #0078d4; font-weight: bold; font-size: 1.2rem; margin: 25px 0 15px 0; padding-bottom: 5px; }
        .mx-table { width: 100%; margin-top: 10px; font-size: 0.95rem; }
        .mx-table th { background-color: #f8f9fa; color: #555; padding: 8px; border: 1px solid #dee2e6; }
        .mx-table td { padding: 8px; border: 1px solid #dee2e6; font-family: 'Consolas', monospace; }
    </style>
</head>
<body>
    <div class="hero text-center">
        <img src="https://dco.microsoft.com/Images/microsoft-white-logo.png" alt="Microsoft" class="logo-img mb-3">
        <h1>Domain Health Checker Report</h1>
        <span class="badge bg-light text-dark">Date: $date</span>
    </div>
    <div class="container-fluid px-4 py-4">
        <div class="row g-3 mb-5 text-center">
            <div class="col"><div class="card p-4 bg-primary text-white"><div>Total Domains</div><div class="stat-number">$totalDomains</div></div></div>
            <div class="col"><div class="card p-4 $(if($spfIssues -gt 0){'bg-danger'}else{'bg-success'}) text-white"><div>SPF Issues</div><div class="stat-number">$spfIssues</div></div></div>
            <div class="col"><div class="card p-4 $(if($dkimIssues -gt 0){'bg-danger'}else{'bg-success'}) text-white"><div>DKIM Issues</div><div class="stat-number">$dkimIssues</div></div></div>
            <div class="col"><div class="card p-4 $(if($dmarcIssues -gt 0){'bg-danger'}else{'bg-success'}) text-white"><div>DMARC Issues</div><div class="stat-number">$dmarcIssues</div></div></div>
            <div class="col"><div class="card p-4 $(if($mtaIssues -gt 0){'bg-danger'}else{'bg-success'}) text-white"><div>MTA-STS Issues</div><div class="stat-number">$mtaIssues</div></div></div>
        </div>
"@

$htmlBody = ""
foreach ($r in $results) {
    $mxLines = ""
    foreach ($mx in $r.MXData) { $mxLines += "<tr><td>$($mx.Hostname)</td><td>$($mx.Pref)</td><td><span class='$($mx.TClass)'>$($mx.TTL)</span></td></tr>" }
    $dkLines = "<li>No common selectors found.</li>"
    if ($r.DkimDetails.Count -gt 0) { $dkLines = "" ; foreach($dk in $r.DkimDetails) { $dkLines += "<li><strong>Selector:</strong> <code>$($dk.Selector)</code> | <strong>TTL:</strong> <span class='$($dk.TClass)'>$($dk.TTL)</span><br><strong>Public Key:</strong> <code class='record-box'>$($dk.Key)</code></li>" } }

    $htmlBody += @"
    <div class="table-card">
        <h2 style="color: #0078d4;">🌐 Domain: $($r.Domain)</h2>
        <div class="section-divider">📫 1. MX Record</div>
        <table class="table table-sm ">
            <thead class="table-light"><tr><th>Hostname</th><th>Priority</th><th>TTL</th></tr></thead>
            <tbody>$mxLines</tbody>
        </table>

        <div class="section-divider">🛡️ 2. SPF Record</div>
        <p><strong>Published Record:</strong> <code class="record-box">$($r.SpfRecord)</code></p>
        <p><strong>SPF Advisory:</strong> <span class="badge-advisory $($r.SpfStatusClass)">$($r.SpfAdvisory)</span></p>
        <div class="mt-2">
            <strong>TTL:</strong> <span class="$($r.SpfTtlClass)">$($r.SpfTTL)</span> |
            <strong>SPF Length:</strong> <span class="cond-badge $($r.SpfLenClass)">$($r.SpfLen)</span> |
            <strong>DNS Lookups:</strong> <span class="cond-badge $($r.SpfLookupClass)">$($r.SpfLookups)</span>
        </div>

        <div class="section-divider">🔑 3. DKIM Details</div>
        <ul>$dkLines</ul>
        <p><strong>DKIM Advisory:</strong> <span class="badge-advisory $($r.DkimStatusClass)">$($r.DkimAdvisory)</span></p>

        <div class="section-divider">🚦 4. DMARC Policy</div>
        <p><strong>Published Record:</strong> <code class="record-box">$($r.DmarcRecord)</code></p>
        <p><strong>DMARC Advisory:</strong> <span class="badge-advisory $($r.DmarcStatusClass)">$($r.DmarcAdvisory)</span></p>
        <p><strong>DMARC TTL:</strong> <span class="$(Get-TtlClass $r.DmarcTTL)">$($r.DmarcTTL)</span></p>

        <div class="section-divider">🌐 5. MTA-STS Policy Details</div>
        <p><strong>MTA-STS Advisory:</strong> <span class="badge-advisory $($r.MtaStatusClass)">$($r.MtaAdvisory)</span></p>
        <div class="card p-3 bg-light border-0">
            <strong>Policy details:</strong>
            <ul class="mb-0 small">
                <li><strong>DNS TTL:</strong> <span class="$(Get-TtlClass $r.MtaInfo.DnsTtl)">$($r.MtaInfo.DnsTtl)</span></li>
                <li><strong>Version:</strong> $($r.MtaInfo.Version)</li>
                <li><strong>Mode:</strong> $($r.MtaInfo.Mode)</li>
                <li><strong>Max Age:</strong> $($r.MtaInfo.MaxAge)</li>
                <li><strong>MX Authorized:</strong> $($r.MtaInfo.MX)</li>
            </ul>
        </div>
    </div>
"@
}

$htmlFooter = @"
<!-- Note -->
        <div class="alert alert-secondary text-center mt-4">
            &#128161; <strong>Note:</strong> Security advisory issues and character limits (SPF Length &gt; 255) are highlighted in <span class="text-danger fw-bold">Red</span>.
        </div>

        <!-- Action Items & Microsoft Recommendations -->
        <div class="card shadow-sm border-primary mb-4">
            <div class="card-header text-white" style="background-color: #0078d4;">&#128221; Action Items &amp; Microsoft Recommendations</div>
            <div class="card-body">
                <div class="list-group">
                    <div class="list-group-item task-link">
                        <strong>&#128279; Reduce SPF record length (max 255 chars)</strong><br>
                        <small><a href="https://www.rfc-editor.org/rfc/rfc7208" target="_blank" class="link-docs">&#128218; RFC 7208</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; SPF Record Syntax</strong><br>
                        <small><a href="http://www.open-spf.org/SPF_Record_Syntax/" target="_blank" class="link-docs">&#128218; open-spf.org</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; Implement DKIM record</strong><br>
                        <small><a href="https://dkim.org/" target="_blank" class="link-docs">&#128218; dkim.org</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; List of DKIM selectors</strong><br>
                        <small><a href="https://www.syskeo.com/en/resources/dkim" target="_blank" class="link-docs">&#128218; syskeo.com</a></small>
                    </div>                    
                    <div class="list-group-item task-link">
                        <strong>&#128279; Upgrade DMARC policy from 'none' to 'reject'</strong><br>
                        <small><a href="https://www.rfc-editor.org/rfc/rfc7489.html" target="_blank" class="link-docs">&#128218; RFC 7489</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; DMARC Record Syntax: Every Tag and Parameter Explained</strong><br>
                        <small><a href="https://dmarccreator.com/resources/dmarc-record-syntax-tags" target="_blank" class="link-docs">&#128218; dmarccreator.com</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; SMTP MTA Strict Transport Security (MTA-STS)</strong><br>
                        <small><a href="https://www.rfc-editor.org/rfc/rfc8461" target="_blank" class="link-docs">&#128218; RFC 8461</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; Implement MTA-STS</strong><br>
                        <small><a href="https://learn.microsoft.com/en-us/purview/enhancing-mail-flow-with-mta-sts" target="_blank" class="link-docs">&#128218; Microsoft Configuration Guide</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; DNS Propagation and TTL Explained</strong><br>
                        <small><a href="https://www.whatsmyiplive.com/blog/dns-propagation-and-ttl.html" target="_blank" class="link-docs">&#128218; whatsmyiplive.com</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; Double check with EmailAuthChecker: Start-EmailAuthChecker</strong><br>
                        <small><a href="https://www.linkedin.com/posts/abdullah-al-zmaili-57496128_i-am-excited-to-share-that-i-have-developed-activity-7358838297034407936-tM70" target="_blank" class="link-docs">&#128218; Introducing EmailAuthChecker</a></small>
                    </div>
                </div>
            </div>
        </div>

        <!-- Microsoft Official Documentation -->
        <div class="card shadow-sm border-info mb-4">
            <div class="card-header text-white" style="background-color: #0078d4;">&#128218; Microsoft Official Documentation</div>
            <div class="card-body">
                <div class="list-group">
                    <div class="list-group-item task-link">
                        <strong>&#128279; SPF Setup Guide</strong><br>
                        <small><a href="https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/email-authentication-spf-configure" target="_blank" class="link-docs">&#128218; Documentaci&oacute;n</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; SPF Setup Parked Domains</strong><br>
                        <small><a href="https://learn.microsoft.com/en-us/defender-office-365/email-authentication-spf-configure#scenario-parked-domains" target="_blank" class="link-docs">&#128218; Documentaci&oacute;n</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; Connect your domain by adding DNS records</strong><br>
                        <small><a href="https://learn.microsoft.com/en-us/microsoft-365/admin/get-help-with-domains/create-dns-records-at-any-dns-hosting-provider?view=o365-worldwide&tabs=domain-connect" target="_blank" class="link-docs">&#128218; Documentaci&oacute;n</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; DKIM Setup Guide</strong><br>
                        <small><a href="https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/email-authentication-dkim-configure" target="_blank" class="link-docs">&#128218; Documentaci&oacute;n</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; DMARC Setup Guide</strong><br>
                        <small><a href="https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/email-authentication-dmarc-configure" target="_blank" class="link-docs">&#128218; Documentaci&oacute;n</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; MTA-STS Enhancing mail flow</strong><br>
                        <small><a href="https://learn.microsoft.com/en-us/purview/enhancing-mail-flow-with-mta-sts" target="_blank" class="link-docs">&#128218; Documentaci&oacute;n</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; Configure trusted ARC sealers</strong><br>
                        <small><a href="https://learn.microsoft.com/en-us/defender-office-365/email-authentication-arc-configure" target="_blank" class="link-docs">&#128218; Documentaci&oacute;n</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; DMARC TXT records for *.onmicrosoft.com</strong><br>
                        <small><a href="https://learn.microsoft.com/en-us/defender-office-365/email-authentication-dmarc-configure#use-the-microsoft-365-admin-center-to-add-dmarc-txt-records-for-onmicrosoftcom-domains-in-microsoft-365" target="_blank" class="link-docs">&#128218; Documentaci&oacute;n</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; DMARC TXT records for parked domains</strong><br>
                        <small><a href="https://learn.microsoft.com/en-us/defender-office-365/email-authentication-dmarc-configure#dmarc-txt-records-for-parked-domains-in-microsoft-365" target="_blank" class="link-docs">&#128218; Documentaci&oacute;n</a></small>
                    </div>
                    <div class="list-group-item task-link">
                        <strong>&#128279; Message Header Analyzer</strong><br>
                        <small><a href="https://mha.azurewebsites.net/" target="_blank" class="link-docs">&#128218; Documentaci&oacute;n</a></small>
                    </div>
                </div>
            </div>
        </div>        
      
        <!-- Footer -->
        <div class="text-center py-4"><p>chiringuito365.com&reg; | Internal Tools 2026</p></div>

    </div><!-- /container -->
</body>
</html>
"@

$finalHtml = $htmlHeader + $htmlBody + $htmlFooter
$finalHtml | Out-File -FilePath $reportPath -Encoding utf8 -Force
Write-Host "--- REPORT FINISHED v5.3.4 ---" -f Green
Invoke-Item $reportPath