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
    Retrieves accepted domains from Exchange Online and checks DNS/email authentication health.

.DESCRIPTION
    Connects to Exchange Online, enumerates accepted domains, and for each domain resolves
    MX records and checks SPF, DKIM, DMARC, MTA-STS, and ARC configuration.
    Results are exported to a timestamped CSV file.

.AUTHOR
    Ernesto Cobos Roqueñí

.VERSION
    2.1 -  2/Mar/2026 - Fix bugs
    2.0 - 20/Mar/2025 - Added DomainHealthChecker integration
    1.0 - 12/Mar/2025 - Initial version

.NOTES
    Requires administrator privileges and internet connectivity.
#>
#Requires -RunAsAdministrator

# --- Module prerequisites ---
$requiredModules = @(
    @{ Name = 'ExchangeOnlineManagement' },
    @{ Name = 'DomainHealthChecker' }
)

foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod.Name)) {
        Write-Host "$($mod.Name) module does not exist. " -ForegroundColor Red -NoNewline
        Write-Host "Downloading and installing now..." -ForegroundColor Yellow
        Install-Module -Name $mod.Name -Force -Confirm:$false
        Write-Host "$($mod.Name) module installed successfully." -ForegroundColor Green
    }
    else {
        Write-Host "$($mod.Name) module found." -ForegroundColor Green
    }
}

# --- Output file setup (removes previous run to avoid duplicates) ---
$outputPath = Join-Path -Path 'C:\Users\ecobos\OneDrive - Microsoft\ecobos\Documents\05 Scripts\Sender Authentication Get AcceptedDomains' -ChildPath "AcceptedDomains_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# --- Connect to Exchange Online ---
try {
    Connect-ExchangeOnline -ErrorAction Stop
}
catch {
    Write-Host "Failed to connect to Exchange Online: $_" -ForegroundColor Red
    return
}

try {
    $ARC = $null
    $arcSealers = 'N/A'
    $arcModified = 'N/A'
    try {
        $ARC = Get-ArcConfig -ErrorAction Stop
        if ($ARC) {
            # Discover the correct property names dynamically
            $arcProps = $ARC | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
            $sealerProp = $arcProps | Where-Object { $_ -match 'Sealer|Trusted' } | Select-Object -First 1
            $modifiedProp = $arcProps | Where-Object { $_ -match 'Modified|Changed|Date|Time' } | Select-Object -First 1
            if ($sealerProp)  { $arcSealers = $ARC.$sealerProp }
            if ($modifiedProp) { $arcModified = $ARC.$modifiedProp }
            Write-Host "ARC config loaded (properties: $($arcProps -join ', '))" -ForegroundColor Green
        }
    } catch { Write-Host "Could not retrieve ARC config: $_" -ForegroundColor Yellow }
    $domains = @(Get-AcceptedDomain)
    $totalDomains = $domains.Count
    $currentIndex = 0

    foreach ($entry in $domains) {
        $currentIndex++
        $domain = $entry.DomainName
        Write-Progress -Activity "Checking domain health" -Status "$domain ($currentIndex of $totalDomains)" -PercentComplete (($currentIndex / $totalDomains) * 100)

        # Resolve MX records with error handling
        $MX = $null
        try {
            $MX = Resolve-DnsName -Name $domain -Type MX -ErrorAction Stop
        }
        catch {
            Write-Host "  Could not resolve MX for $domain" -ForegroundColor Yellow
        }

        if ($null -eq $MX) {
            $MXHostName   = "DNS name does not exist"
            $MXType       = "N/A"
            $MXPreference = "N/A"
            $MXTTL        = "N/A"
        }
        else {
            $MXHostName   = $MX.NameExchange -join ' | '
            $MXType       = $MX.Type -join ' | '
            $MXPreference = $MX.Preference -join ' | '
            $MXTTL        = $MX.TTL -join ' | '
        }

        # SPF / DKIM / DMARC / MTA-STS check
        $DHC = Invoke-SpfDkimDmarc -Name $domain

        $SenderAuth = [PSCustomObject]@{
            Domain_Name      = $domain
            RecordType       = $MXType
            Host_Name        = $MXHostName
            TimeToLive       = $MXTTL
            Preference       = $MXPreference
            SPFAdvisory      = $DHC.SpfAdvisory
            SpfRecord        = $DHC.SpfRecord
            SPFLength        = $DHC.SPFRecordLength
            DmarcAdvisory    = $DHC.DmarcAdvisory
            DmarcRecord      = $DHC.DmarcRecord
            ARCTrustedSealers = $arcSealers
            LastModified     = $arcModified
            MtaAdvisory      = $DHC.MtaAdvisory
            MtaRecord        = $DHC.MtaRecord
            DkimAdvisory     = $DHC.DkimAdvisory
            DkimSelector     = $DHC.DkimSelector
            DkimRecord       = $DHC.DkimRecord
        }

        $SenderAuth | Export-Csv -Path $outputPath -Append -NoTypeInformation
    }

    Write-Progress -Activity "Checking domain health" -Completed
    Write-Host "`nResults exported to: $outputPath" -ForegroundColor Cyan
    Invoke-Item $outputPath
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false
}