##############################################################################################
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
##############################################################################################
<#
.SYNOPSIS
    Obtiene un reporte detallado de todas las reglas de flujo de correo (Transport Rules) en Exchange Online.

.DESCRIPTION
    Este script recopila la mayor cantidad de detalles de cada Transport Rule configurada en Exchange Online:
    - Información general (nombre, estado, prioridad, modo, comentarios)
    - Condiciones (remitente, destinatario, dominios, palabras clave, encabezados, etc.)
    - Excepciones
    - Acciones (rechazar, redirigir, agregar encabezados, disclaimers, SCL, etc.)
    - Fechas de activación/expiración
    - Auditoría y metadatos (última modificación, creador, etc.)

    Genera dos salidas:
    1. Reporte en consola con formato visual
    2. Exportación a CSV con todos los campos relevantes

.NOTES
    Requiere conexión previa a Exchange Online:
        Connect-ExchangeOnline

    Autor  : ecobos
    Fecha  : 2026-03-03
    Versión: 1.1
#>

# ─────────────────────────────────────────────
# Validación de módulo y carpeta de reportes
# ─────────────────────────────────────────────
if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
    Write-Host "Módulo ExchangeOnlineManagement instalado correctamente." -ForegroundColor DarkGray
}
else {
    Write-Host "[X] Módulo ExchangeOnlineManagement no encontrado. " -ForegroundColor Red -NoNewline
    Write-Host "Descargando e instalando..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser
}

# ─────────────────────────────────────────────
# Conexión a Exchange Online
# ─────────────────────────────────────────────
try {
    $null = Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' }
    if (-not $_) { throw "No conectado" }
    Write-Host "Ya existe una sesión activa de Exchange Online." -ForegroundColor DarkGray
}
catch {
    Write-Host "Conectando a Exchange Online..." -ForegroundColor Yellow
    try {
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Host "Conexión establecida exitosamente." -ForegroundColor Green
    }
    catch {
        Write-Host "[ERROR] No se pudo conectar a Exchange Online." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        return
    }
}

$reportDir = "C:\Scripts\TransportRules"
if (-not (Test-Path $reportDir)) {
    New-Item -Path $reportDir -ItemType Directory -Force | Out-Null
    Write-Host "Carpeta creada: $reportDir" -ForegroundColor DarkGray
}
else {
    Write-Host "Carpeta de reportes existe: $reportDir" -ForegroundColor DarkGray
}

$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$tenantName = (Get-OrganizationConfig).DisplayName
$csvPath    = Join-Path $reportDir "TransportRules_$timestamp.csv"
$htmlPath   = Join-Path $reportDir "TransportRules_$timestamp.html"

# ─────────────────────────────────────────────
# Función auxiliar para convertir arrays a string
# ─────────────────────────────────────────────
function ConvertTo-FlatString {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        return ($Value | ForEach-Object { $_.ToString() }) -join "; "
    }
    return $Value.ToString()
}

# ─────────────────────────────────────────────
# Obtener todas las reglas de transporte
# ─────────────────────────────────────────────


try {
    $rules = Get-TransportRule -ResultSize Unlimited | Sort-Object Priority
}
catch {
    Write-Host "[ERROR] No se pudieron obtener las reglas. ¿Está conectado a Exchange Online?" -ForegroundColor Red
    Write-Host "Ejecute: Connect-ExchangeOnline" -ForegroundColor Yellow
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}

$totalRules    = ($rules | Measure-Object).Count
$enabledRules  = ($rules | Where-Object { $_.State -eq 'Enabled' } | Measure-Object).Count
$disabledRules = ($rules | Where-Object { $_.State -eq 'Disabled' } | Measure-Object).Count

if ($totalRules -eq 0) {
    Write-Host "No se encontraron reglas de transporte." -ForegroundColor Yellow
    return
}

# ─────────────────────────────────────────────
# Recorrer cada regla y construir datos de exportación
# ─────────────────────────────────────────────
$reportData = @()

foreach ($rule in $rules) {

    # ── Construir objeto para exportación CSV ──
    $reportData += [PSCustomObject]@{
        # General
        Name                                 = $rule.Name
        State                                = $rule.State
        Priority                             = $rule.Priority
        Mode                                 = $rule.Mode
        Guid                                 = $rule.Guid
        Comments                             = $rule.Comments
        Description                          = $rule.Description
        WhenCreated                          = $rule.WhenCreated
        WhenChanged                          = $rule.WhenChanged
        CreatedBy                            = $rule.CreatedBy
        LastModifiedBy                       = $rule.LastModifiedBy
        ActivationDate                       = $rule.ActivationDate
        ExpiryDate                           = $rule.ExpiryDate
        RuleType                             = $rule.RuleType
        DlpPolicy                            = $rule.DlpPolicy
        ManuallyModified                     = $rule.ManuallyModified
        SenderAddressLocation                = $rule.SenderAddressLocation

        # Condiciones - Remitente
        From                                 = ConvertTo-FlatString $rule.From
        FromAddressContainsWords             = ConvertTo-FlatString $rule.FromAddressContainsWords
        FromAddressMatchesPatterns           = ConvertTo-FlatString $rule.FromAddressMatchesPatterns
        FromMemberOf                         = ConvertTo-FlatString $rule.FromMemberOf
        FromScope                            = ConvertTo-FlatString $rule.FromScope
        SenderDomainIs                       = ConvertTo-FlatString $rule.SenderDomainIs
        SenderIpRanges                       = ConvertTo-FlatString $rule.SenderIpRanges
        SenderADAttributeContainsWords       = ConvertTo-FlatString $rule.SenderADAttributeContainsWords
        SenderADAttributeMatchesPatterns     = ConvertTo-FlatString $rule.SenderADAttributeMatchesPatterns

        # Condiciones - Destinatario
        SentTo                               = ConvertTo-FlatString $rule.SentTo
        SentToMemberOf                       = ConvertTo-FlatString $rule.SentToMemberOf
        SentToScope                          = ConvertTo-FlatString $rule.SentToScope
        RecipientDomainIs                    = ConvertTo-FlatString $rule.RecipientDomainIs
        RecipientAddressContainsWords        = ConvertTo-FlatString $rule.RecipientAddressContainsWords
        RecipientAddressMatchesPatterns      = ConvertTo-FlatString $rule.RecipientAddressMatchesPatterns
        AnyOfToHeader                        = ConvertTo-FlatString $rule.AnyOfToHeader
        AnyOfToHeaderMemberOf                = ConvertTo-FlatString $rule.AnyOfToHeaderMemberOf
        AnyOfCcHeader                        = ConvertTo-FlatString $rule.AnyOfCcHeader
        AnyOfCcHeaderMemberOf                = ConvertTo-FlatString $rule.AnyOfCcHeaderMemberOf
        AnyOfToCcHeader                      = ConvertTo-FlatString $rule.AnyOfToCcHeader
        AnyOfToCcHeaderMemberOf              = ConvertTo-FlatString $rule.AnyOfToCcHeaderMemberOf
        BetweenMemberOf1                     = ConvertTo-FlatString $rule.BetweenMemberOf1
        BetweenMemberOf2                     = ConvertTo-FlatString $rule.BetweenMemberOf2

        # Condiciones - Contenido
        SubjectContainsWords                 = ConvertTo-FlatString $rule.SubjectContainsWords
        SubjectMatchesPatterns               = ConvertTo-FlatString $rule.SubjectMatchesPatterns
        SubjectOrBodyContainsWords           = ConvertTo-FlatString $rule.SubjectOrBodyContainsWords
        SubjectOrBodyMatchesPatterns         = ConvertTo-FlatString $rule.SubjectOrBodyMatchesPatterns
        HeaderContainsMessageHeader          = ConvertTo-FlatString $rule.HeaderContainsMessageHeader
        HeaderContainsWords                  = ConvertTo-FlatString $rule.HeaderContainsWords
        HeaderMatchesMessageHeader           = ConvertTo-FlatString $rule.HeaderMatchesMessageHeader
        HeaderMatchesPatterns                = ConvertTo-FlatString $rule.HeaderMatchesPatterns
        MessageTypeMatches                   = ConvertTo-FlatString $rule.MessageTypeMatches
        HasClassification                    = ConvertTo-FlatString $rule.HasClassification
        HasNoClassification                  = ConvertTo-FlatString $rule.HasNoClassification
        HasSenderOverride                    = ConvertTo-FlatString $rule.HasSenderOverride
        MessageSizeOver                      = ConvertTo-FlatString $rule.MessageSizeOver
        SCLOver                              = ConvertTo-FlatString $rule.SCLOver
        WithImportance                       = ConvertTo-FlatString $rule.WithImportance

        # Condiciones - Adjuntos
        AttachmentSizeOver                   = ConvertTo-FlatString $rule.AttachmentSizeOver
        AttachmentIsUnsupported              = ConvertTo-FlatString $rule.AttachmentIsUnsupported
        AttachmentProcessingLimitExceeded    = ConvertTo-FlatString $rule.AttachmentProcessingLimitExceeded
        AttachmentHasExecutableContent       = ConvertTo-FlatString $rule.AttachmentHasExecutableContent
        AttachmentIsPasswordProtected        = ConvertTo-FlatString $rule.AttachmentIsPasswordProtected
        AttachmentContainsWords              = ConvertTo-FlatString $rule.AttachmentContainsWords
        AttachmentMatchesPatterns            = ConvertTo-FlatString $rule.AttachmentMatchesPatterns
        AttachmentNameMatchesPatterns        = ConvertTo-FlatString $rule.AttachmentNameMatchesPatterns
        AttachmentExtensionMatchesWords      = ConvertTo-FlatString $rule.AttachmentExtensionMatchesWords
        AttachmentPropertyContainsWords      = ConvertTo-FlatString $rule.AttachmentPropertyContainsWords
        ContentCharacterSetContainsWords     = ConvertTo-FlatString $rule.ContentCharacterSetContainsWords

        # Excepciones principales
        ExceptIfFrom                                 = ConvertTo-FlatString $rule.ExceptIfFrom
        ExceptIfFromAddressContainsWords             = ConvertTo-FlatString $rule.ExceptIfFromAddressContainsWords
        ExceptIfFromAddressMatchesPatterns           = ConvertTo-FlatString $rule.ExceptIfFromAddressMatchesPatterns
        ExceptIfFromMemberOf                         = ConvertTo-FlatString $rule.ExceptIfFromMemberOf
        ExceptIfFromScope                            = ConvertTo-FlatString $rule.ExceptIfFromScope
        ExceptIfSenderDomainIs                       = ConvertTo-FlatString $rule.ExceptIfSenderDomainIs
        ExceptIfSenderIpRanges                       = ConvertTo-FlatString $rule.ExceptIfSenderIpRanges
        ExceptIfSentTo                               = ConvertTo-FlatString $rule.ExceptIfSentTo
        ExceptIfSentToMemberOf                       = ConvertTo-FlatString $rule.ExceptIfSentToMemberOf
        ExceptIfRecipientDomainIs                    = ConvertTo-FlatString $rule.ExceptIfRecipientDomainIs
        ExceptIfSubjectContainsWords                 = ConvertTo-FlatString $rule.ExceptIfSubjectContainsWords
        ExceptIfSubjectOrBodyContainsWords           = ConvertTo-FlatString $rule.ExceptIfSubjectOrBodyContainsWords
        ExceptIfHeaderContainsMessageHeader          = ConvertTo-FlatString $rule.ExceptIfHeaderContainsMessageHeader
        ExceptIfHeaderContainsWords                  = ConvertTo-FlatString $rule.ExceptIfHeaderContainsWords
        ExceptIfAttachmentNameMatchesPatterns        = ConvertTo-FlatString $rule.ExceptIfAttachmentNameMatchesPatterns
        ExceptIfAttachmentExtensionMatchesWords      = ConvertTo-FlatString $rule.ExceptIfAttachmentExtensionMatchesWords

        # Acciones
        RejectMessageReasonText              = ConvertTo-FlatString $rule.RejectMessageReasonText
        RejectMessageEnhancedStatusCode      = ConvertTo-FlatString $rule.RejectMessageEnhancedStatusCode
        DeleteMessage                        = $rule.DeleteMessage
        Disconnect                           = $rule.Disconnect
        Quarantine                           = $rule.Quarantine
        RedirectMessageTo                    = ConvertTo-FlatString $rule.RedirectMessageTo
        AddToRecipients                      = ConvertTo-FlatString $rule.AddToRecipients
        CopyTo                               = ConvertTo-FlatString $rule.CopyTo
        BlindCopyTo                          = ConvertTo-FlatString $rule.BlindCopyTo
        ModerateMessageByUser                = ConvertTo-FlatString $rule.ModerateMessageByUser
        ModerateMessageByManager             = $rule.ModerateMessageByManager
        AddManagerAsRecipientType            = ConvertTo-FlatString $rule.AddManagerAsRecipientType
        PrependSubject                       = ConvertTo-FlatString $rule.PrependSubject
        SetHeaderName                        = ConvertTo-FlatString $rule.SetHeaderName
        SetHeaderValue                       = ConvertTo-FlatString $rule.SetHeaderValue
        RemoveHeader                         = ConvertTo-FlatString $rule.RemoveHeader
        ApplyHtmlDisclaimerLocation          = ConvertTo-FlatString $rule.ApplyHtmlDisclaimerLocation
        ApplyHtmlDisclaimerText              = ConvertTo-FlatString $rule.ApplyHtmlDisclaimerText
        ApplyHtmlDisclaimerFallbackAction    = ConvertTo-FlatString $rule.ApplyHtmlDisclaimerFallbackAction
        SetSCL                               = ConvertTo-FlatString $rule.SetSCL
        ApplyClassification                  = ConvertTo-FlatString $rule.ApplyClassification
        ApplyRightsProtectionTemplate        = ConvertTo-FlatString $rule.ApplyRightsProtectionTemplate
        SetAuditSeverity                     = ConvertTo-FlatString $rule.SetAuditSeverity
        GenerateIncidentReport               = ConvertTo-FlatString $rule.GenerateIncidentReport
        IncidentReportContent                = ConvertTo-FlatString $rule.IncidentReportContent
        GenerateNotification                 = ConvertTo-FlatString $rule.GenerateNotification
        RouteMessageOutboundConnector        = ConvertTo-FlatString $rule.RouteMessageOutboundConnector
        RouteMessageOutboundRequireTls       = $rule.RouteMessageOutboundRequireTls
        ApplyOME                             = $rule.ApplyOME
        RemoveOME                            = $rule.RemoveOME
        RemoveOMEv2                          = $rule.RemoveOMEv2
        RemoveRMSAttachmentEncryption        = $rule.RemoveRMSAttachmentEncryption
        StopRuleProcessing                   = $rule.StopRuleProcessing
        SenderNotificationType               = ConvertTo-FlatString $rule.SenderNotificationType
        SmtpRejectMessageRejectText          = ConvertTo-FlatString $rule.SmtpRejectMessageRejectText
        SmtpRejectMessageRejectStatusCode    = ConvertTo-FlatString $rule.SmtpRejectMessageRejectStatusCode
    }
}

# ─────────────────────────────────────────────
# Exportar a CSV (UTF-8 con BOM para compatibilidad con Excel y caracteres especiales)
# ─────────────────────────────────────────────
$utf8Bom = New-Object System.Text.UTF8Encoding $true
$csvContent = ($reportData | ConvertTo-Csv -NoTypeInformation) -join "`r`n"
[System.IO.File]::WriteAllText($csvPath, $csvContent, $utf8Bom)

# ─────────────────────────────────────────────
# Generar reporte HTML
# ─────────────────────────────────────────────
$htmlHead = @"
<style>
    body   { font-family: 'Segoe UI', Tahoma, sans-serif; margin: 20px; background: #f5f5f5; color: #333; }
    h1     { color: #0078d4; border-bottom: 2px solid #0078d4; padding-bottom: 8px; }
    h2     { color: #005a9e; margin-top: 30px; }
    h3     { color: #333; }
    table  { border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 13px; }
    th     { background: #0078d4; color: #fff; padding: 10px; text-align: left; }
    td     { border: 1px solid #ddd; padding: 8px; color: #333; }
    tr:nth-child(even) { background: #e9e9e9; }
    tr:nth-child(odd)  { background: #fff; }
    .enabled  { color: #107c10; font-weight: bold; }
    .disabled { color: #d13438; font-weight: bold; }
    .summary  { background: #0078d4; color: #fff; padding: 12px 20px; border-radius: 6px; display: inline-block; margin: 5px; }
</style>
"@

$htmlBody = '<h1>Reporte de Reglas de Flujo de Correo - Exchange Online <em style="font-size: 0.75em; font-weight: normal; margin-left: 80px;">&ldquo;La tecnolog&iacute;a habilita la seguridad, pero es la disciplina la que garantiza su efectividad&rdquo;</em></h1>'
$htmlBody += "<p><strong>Tenant:</strong> $tenantName | <strong>Generado:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>"
$htmlBody += "<div>"
$htmlBody += "<span class='summary'>Total: $totalRules</span>"
$htmlBody += "<span class='summary'>Habilitadas: $enabledRules</span>"
$htmlBody += "<span class='summary'>Deshabilitadas: $disabledRules</span>"
$htmlBody += "</div>"

$htmlBody += "<h2>Resumen de Reglas</h2>"
$htmlBody += "<table><tr><th>Prioridad</th><th>Nombre</th><th>Estado</th><th>Modo</th><th>Última Modificación</th><th>Acción Principal</th></tr>"

foreach ($rule in $rules) {
    $stateClass = if ($rule.State -eq 'Enabled') { 'enabled' } else { 'disabled' }

    # Determinar acción principal
    $mainAction = @()
    if ($rule.RejectMessageReasonText)       { $mainAction += "Rechazar" }
    if ($rule.DeleteMessage -eq $true)       { $mainAction += "Eliminar" }
    if ($rule.Quarantine -eq $true)          { $mainAction += "Cuarentena" }
    if ($rule.RedirectMessageTo)             { $mainAction += "Redirigir" }
    if ($rule.PrependSubject)                { $mainAction += "Prefijo asunto" }
    if ($rule.SetHeaderName)                 { $mainAction += "Encabezado" }
    if ($rule.ApplyHtmlDisclaimerText)       { $mainAction += "Disclaimer" }
    if ($rule.SetSCL)                        { $mainAction += "SCL=$($rule.SetSCL)" }
    if ($rule.CopyTo)                        { $mainAction += "Copiar" }
    if ($rule.BlindCopyTo)                   { $mainAction += "CCO" }
    if ($rule.ModerateMessageByUser)         { $mainAction += "Moderar" }
    if ($rule.ApplyOME -eq $true)            { $mainAction += "Cifrar (OME)" }
    if ($rule.StopRuleProcessing -eq $true)  { $mainAction += "Detener procesamiento" }
    if ($rule.ApplyRightsProtectionTemplate) { $mainAction += "RMS/IRM" }
    if ($mainAction.Count -eq 0)             { $mainAction += "—" }

    $htmlBody += "<tr>"
    $htmlBody += "<td style='text-align:center'>$($rule.Priority)</td>"
    $htmlBody += "<td>$($rule.Name)</td>"
    $htmlBody += "<td class='$stateClass'>$($rule.State)</td>"
    $htmlBody += "<td>$($rule.Mode)</td>"
    $htmlBody += "<td>$($rule.WhenChanged)</td>"
    $htmlBody += "<td>$($mainAction -join ', ')</td>"
    $htmlBody += "</tr>"
}
$htmlBody += "</table>"

$htmlFooter = '<footer style="text-align: center; margin-top: 40px; padding: 15px 0; border-top: 2px solid #0078d4; color: #555; font-size: 13px;">chiringuito365.com&reg; | Internal Tools 2026</footer>'

$htmlReport = ConvertTo-Html -Head $htmlHead -Body ($htmlBody + $htmlFooter) -Title "Transport Rules Report"
$htmlReport | Out-File -FilePath $htmlPath -Encoding utf8

# ─────────────────────────────────────────────
# Resumen final
# ─────────────────────────────────────────────
Write-Host ""
Write-Host " Reportes generados exitosamente:" -ForegroundColor Green
Write-Host "   CSV  : $csvPath" -ForegroundColor Cyan
Write-Host "   HTML : $htmlPath" -ForegroundColor Cyan
Write-Host ""
Write-Host " Total de propiedades exportadas por regla: ~100+" -ForegroundColor DarkGray
Write-Host ("═" * 100) -ForegroundColor DarkCyan

# Abrir el reporte HTML en el navegador predeterminado
Invoke-Item $htmlPath
Write-Host ""
Write-Host "chiringuito365.com® | Internal Tools 2026" -ForegroundColor DarkGray
Write-Host ""