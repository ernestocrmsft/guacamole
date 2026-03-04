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
    Versión: 1.0
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

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath   = Join-Path $reportDir "TransportRules_$timestamp.csv"
$htmlPath  = Join-Path $reportDir "TransportRules_$timestamp.html"

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
# Función para colorear el estado de la regla
# ─────────────────────────────────────────────
function Write-RuleHeader {
    param(
        [string]$Name,
        [string]$State,
        [int]$Priority,
        [string]$Mode
    )
    $stateColor = switch ($State) {
        'Enabled'  { 'Green'  }
        'Disabled' { 'Red'    }
        default    { 'Yellow' }
    }
    $modeColor = switch ($Mode) {
        'Enforce'  { 'Green'  }
        'Audit'    { 'Yellow' }
        default    { 'Cyan'   }
    }
    Write-Host ""
    Write-Host ("═" * 100) -ForegroundColor DarkCyan
    Write-Host " Regla: " -NoNewline -ForegroundColor White
    Write-Host "$Name" -ForegroundColor Cyan
    Write-Host " Estado: " -NoNewline -ForegroundColor White
    Write-Host "$State" -ForegroundColor $stateColor -NoNewline
    Write-Host "  |  Prioridad: " -NoNewline -ForegroundColor White
    Write-Host "$Priority" -ForegroundColor Yellow -NoNewline
    Write-Host "  |  Modo: " -NoNewline -ForegroundColor White
    Write-Host "$Mode" -ForegroundColor $modeColor
    Write-Host ("─" * 100) -ForegroundColor DarkGray
}

function Write-Detail {
    param(
        [string]$Label,
        [string]$Value,
        [string]$Color = "White"
    )
    if (-not [string]::IsNullOrWhiteSpace($Value)) {
        Write-Host "   $($Label.PadRight(45))" -NoNewline -ForegroundColor Gray
        Write-Host "$Value" -ForegroundColor $Color
    }
}

# ─────────────────────────────────────────────
# Obtener todas las reglas de transporte
# ─────────────────────────────────────────────
Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║   Reporte Detallado de Reglas de Flujo de Correo (Transport)   ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "Obteniendo reglas de transporte..." -ForegroundColor Yellow

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

Write-Host ""
Write-Host " Resumen rápido:" -ForegroundColor White
Write-Host "   Total de reglas:        $totalRules" -ForegroundColor Cyan
Write-Host "   Habilitadas:            $enabledRules" -ForegroundColor Green
Write-Host "   Deshabilitadas:         $disabledRules" -ForegroundColor Red
Write-Host ""

if ($totalRules -eq 0) {
    Write-Host "No se encontraron reglas de transporte." -ForegroundColor Yellow
    return
}

# ─────────────────────────────────────────────
# Recorrer cada regla y mostrar detalles
# ─────────────────────────────────────────────
$reportData = @()

foreach ($rule in $rules) {

    Write-RuleHeader -Name $rule.Name -State $rule.State -Priority $rule.Priority -Mode $rule.Mode

    # ── Información General ──
    Write-Host "   [Información General]" -ForegroundColor DarkYellow
    Write-Detail -Label "GUID"                        -Value (ConvertTo-FlatString $rule.Guid)
    Write-Detail -Label "Comentarios"                 -Value (ConvertTo-FlatString $rule.Comments)
    Write-Detail -Label "Descripción"                 -Value (ConvertTo-FlatString $rule.Description)
    Write-Detail -Label "Fecha de creación"           -Value (ConvertTo-FlatString $rule.WhenCreated)
    Write-Detail -Label "Última modificación"         -Value (ConvertTo-FlatString $rule.WhenChanged)
    Write-Detail -Label "Creado por"                  -Value (ConvertTo-FlatString $rule.CreatedBy)
    Write-Detail -Label "Última modificación por"     -Value (ConvertTo-FlatString $rule.LastModifiedBy)
    Write-Detail -Label "Fecha de activación"         -Value (ConvertTo-FlatString $rule.ActivationDate)
    Write-Detail -Label "Fecha de expiración"         -Value (ConvertTo-FlatString $rule.ExpiryDate)
    Write-Detail -Label "Nivel de severidad"          -Value (ConvertTo-FlatString $rule.SetAuditSeverity)
    Write-Detail -Label "Tipo de regla"               -Value (ConvertTo-FlatString $rule.RuleType)
    Write-Detail -Label "Manual de usuario"           -Value (ConvertTo-FlatString $rule.ManuallyModified)
    Write-Detail -Label "Tiene DLP Policy"            -Value (ConvertTo-FlatString $rule.DlpPolicy)

    # ── Condiciones ──
    Write-Host "   [Condiciones]" -ForegroundColor DarkYellow
    Write-Detail -Label "From (Remitente)"                   -Value (ConvertTo-FlatString $rule.From)
    Write-Detail -Label "FromAddressContainsWords"           -Value (ConvertTo-FlatString $rule.FromAddressContainsWords)
    Write-Detail -Label "FromAddressMatchesPatterns"         -Value (ConvertTo-FlatString $rule.FromAddressMatchesPatterns)
    Write-Detail -Label "FromMemberOf"                       -Value (ConvertTo-FlatString $rule.FromMemberOf)
    Write-Detail -Label "FromScope"                          -Value (ConvertTo-FlatString $rule.FromScope)
    Write-Detail -Label "SenderDomainIs"                     -Value (ConvertTo-FlatString $rule.SenderDomainIs)
    Write-Detail -Label "SenderAddressLocation"              -Value (ConvertTo-FlatString $rule.SenderAddressLocation)
    Write-Detail -Label "SenderIpRanges"                     -Value (ConvertTo-FlatString $rule.SenderIpRanges)
    Write-Detail -Label "SenderADAttributeContainsWords"     -Value (ConvertTo-FlatString $rule.SenderADAttributeContainsWords)
    Write-Detail -Label "SenderADAttributeMatchesPatterns"   -Value (ConvertTo-FlatString $rule.SenderADAttributeMatchesPatterns)
    Write-Detail -Label "SentTo"                             -Value (ConvertTo-FlatString $rule.SentTo)
    Write-Detail -Label "SentToMemberOf"                     -Value (ConvertTo-FlatString $rule.SentToMemberOf)
    Write-Detail -Label "SentToScope"                        -Value (ConvertTo-FlatString $rule.SentToScope)
    Write-Detail -Label "RecipientDomainIs"                  -Value (ConvertTo-FlatString $rule.RecipientDomainIs)
    Write-Detail -Label "RecipientAddressContainsWords"      -Value (ConvertTo-FlatString $rule.RecipientAddressContainsWords)
    Write-Detail -Label "RecipientAddressMatchesPatterns"    -Value (ConvertTo-FlatString $rule.RecipientAddressMatchesPatterns)
    Write-Detail -Label "RecipientADAttributeContainsWords"  -Value (ConvertTo-FlatString $rule.RecipientADAttributeContainsWords)
    Write-Detail -Label "RecipientADAttributeMatchesPatterns" -Value (ConvertTo-FlatString $rule.RecipientADAttributeMatchesPatterns)
    Write-Detail -Label "AnyOfToHeader"                      -Value (ConvertTo-FlatString $rule.AnyOfToHeader)
    Write-Detail -Label "AnyOfToHeaderMemberOf"              -Value (ConvertTo-FlatString $rule.AnyOfToHeaderMemberOf)
    Write-Detail -Label "AnyOfCcHeader"                      -Value (ConvertTo-FlatString $rule.AnyOfCcHeader)
    Write-Detail -Label "AnyOfCcHeaderMemberOf"              -Value (ConvertTo-FlatString $rule.AnyOfCcHeaderMemberOf)
    Write-Detail -Label "AnyOfToCcHeader"                    -Value (ConvertTo-FlatString $rule.AnyOfToCcHeader)
    Write-Detail -Label "AnyOfToCcHeaderMemberOf"            -Value (ConvertTo-FlatString $rule.AnyOfToCcHeaderMemberOf)
    Write-Detail -Label "BetweenMemberOf1"                   -Value (ConvertTo-FlatString $rule.BetweenMemberOf1)
    Write-Detail -Label "BetweenMemberOf2"                   -Value (ConvertTo-FlatString $rule.BetweenMemberOf2)
    Write-Detail -Label "SubjectContainsWords"               -Value (ConvertTo-FlatString $rule.SubjectContainsWords)
    Write-Detail -Label "SubjectMatchesPatterns"             -Value (ConvertTo-FlatString $rule.SubjectMatchesPatterns)
    Write-Detail -Label "SubjectOrBodyContainsWords"         -Value (ConvertTo-FlatString $rule.SubjectOrBodyContainsWords)
    Write-Detail -Label "SubjectOrBodyMatchesPatterns"       -Value (ConvertTo-FlatString $rule.SubjectOrBodyMatchesPatterns)
    Write-Detail -Label "HeaderContainsMessageHeader"        -Value (ConvertTo-FlatString $rule.HeaderContainsMessageHeader)
    Write-Detail -Label "HeaderContainsWords"                -Value (ConvertTo-FlatString $rule.HeaderContainsWords)
    Write-Detail -Label "HeaderMatchesMessageHeader"         -Value (ConvertTo-FlatString $rule.HeaderMatchesMessageHeader)
    Write-Detail -Label "HeaderMatchesPatterns"              -Value (ConvertTo-FlatString $rule.HeaderMatchesPatterns)
    Write-Detail -Label "MessageTypeMatches"                 -Value (ConvertTo-FlatString $rule.MessageTypeMatches)
    Write-Detail -Label "HasClassification"                  -Value (ConvertTo-FlatString $rule.HasClassification)
    Write-Detail -Label "HasSenderOverride"                  -Value (ConvertTo-FlatString $rule.HasSenderOverride)
    Write-Detail -Label "MessageSizeOver"                    -Value (ConvertTo-FlatString $rule.MessageSizeOver)
    Write-Detail -Label "AttachmentSizeOver"                 -Value (ConvertTo-FlatString $rule.AttachmentSizeOver)
    Write-Detail -Label "AttachmentIsUnsupported"            -Value (ConvertTo-FlatString $rule.AttachmentIsUnsupported)
    Write-Detail -Label "AttachmentProcessingLimitExceeded"  -Value (ConvertTo-FlatString $rule.AttachmentProcessingLimitExceeded)
    Write-Detail -Label "AttachmentHasExecutableContent"     -Value (ConvertTo-FlatString $rule.AttachmentHasExecutableContent)
    Write-Detail -Label "AttachmentIsPasswordProtected"      -Value (ConvertTo-FlatString $rule.AttachmentIsPasswordProtected)
    Write-Detail -Label "AttachmentContainsWords"            -Value (ConvertTo-FlatString $rule.AttachmentContainsWords)
    Write-Detail -Label "AttachmentMatchesPatterns"          -Value (ConvertTo-FlatString $rule.AttachmentMatchesPatterns)
    Write-Detail -Label "AttachmentNameMatchesPatterns"      -Value (ConvertTo-FlatString $rule.AttachmentNameMatchesPatterns)
    Write-Detail -Label "AttachmentExtensionMatchesWords"    -Value (ConvertTo-FlatString $rule.AttachmentExtensionMatchesWords)
    Write-Detail -Label "AttachmentPropertyContainsWords"    -Value (ConvertTo-FlatString $rule.AttachmentPropertyContainsWords)
    Write-Detail -Label "ContentCharacterSetContainsWords"   -Value (ConvertTo-FlatString $rule.ContentCharacterSetContainsWords)
    Write-Detail -Label "HasNoClassification"                -Value (ConvertTo-FlatString $rule.HasNoClassification)
    Write-Detail -Label "SCLOver"                            -Value (ConvertTo-FlatString $rule.SCLOver)
    Write-Detail -Label "WithImportance"                     -Value (ConvertTo-FlatString $rule.WithImportance)
    Write-Detail -Label "ManagerAddresses"                   -Value (ConvertTo-FlatString $rule.ManagerAddresses)
    Write-Detail -Label "ManagerForEvaluatedUser"            -Value (ConvertTo-FlatString $rule.ManagerForEvaluatedUser)
    Write-Detail -Label "ADComparisonAttribute"              -Value (ConvertTo-FlatString $rule.ADComparisonAttribute)
    Write-Detail -Label "ADComparisonOperator"               -Value (ConvertTo-FlatString $rule.ADComparisonOperator)
    Write-Detail -Label "SenderManagementRelationship"       -Value (ConvertTo-FlatString $rule.SenderManagementRelationship)

    # ── Excepciones ──
    Write-Host "   [Excepciones]" -ForegroundColor DarkYellow
    Write-Detail -Label "ExceptIfFrom"                              -Value (ConvertTo-FlatString $rule.ExceptIfFrom)
    Write-Detail -Label "ExceptIfFromAddressContainsWords"          -Value (ConvertTo-FlatString $rule.ExceptIfFromAddressContainsWords)
    Write-Detail -Label "ExceptIfFromAddressMatchesPatterns"        -Value (ConvertTo-FlatString $rule.ExceptIfFromAddressMatchesPatterns)
    Write-Detail -Label "ExceptIfFromMemberOf"                      -Value (ConvertTo-FlatString $rule.ExceptIfFromMemberOf)
    Write-Detail -Label "ExceptIfFromScope"                         -Value (ConvertTo-FlatString $rule.ExceptIfFromScope)
    Write-Detail -Label "ExceptIfSenderDomainIs"                    -Value (ConvertTo-FlatString $rule.ExceptIfSenderDomainIs)
    Write-Detail -Label "ExceptIfSenderIpRanges"                    -Value (ConvertTo-FlatString $rule.ExceptIfSenderIpRanges)
    Write-Detail -Label "ExceptIfSentTo"                            -Value (ConvertTo-FlatString $rule.ExceptIfSentTo)
    Write-Detail -Label "ExceptIfSentToMemberOf"                    -Value (ConvertTo-FlatString $rule.ExceptIfSentToMemberOf)
    Write-Detail -Label "ExceptIfRecipientDomainIs"                 -Value (ConvertTo-FlatString $rule.ExceptIfRecipientDomainIs)
    Write-Detail -Label "ExceptIfRecipientAddressContainsWords"     -Value (ConvertTo-FlatString $rule.ExceptIfRecipientAddressContainsWords)
    Write-Detail -Label "ExceptIfRecipientAddressMatchesPatterns"   -Value (ConvertTo-FlatString $rule.ExceptIfRecipientAddressMatchesPatterns)
    Write-Detail -Label "ExceptIfSubjectContainsWords"              -Value (ConvertTo-FlatString $rule.ExceptIfSubjectContainsWords)
    Write-Detail -Label "ExceptIfSubjectMatchesPatterns"            -Value (ConvertTo-FlatString $rule.ExceptIfSubjectMatchesPatterns)
    Write-Detail -Label "ExceptIfSubjectOrBodyContainsWords"        -Value (ConvertTo-FlatString $rule.ExceptIfSubjectOrBodyContainsWords)
    Write-Detail -Label "ExceptIfSubjectOrBodyMatchesPatterns"      -Value (ConvertTo-FlatString $rule.ExceptIfSubjectOrBodyMatchesPatterns)
    Write-Detail -Label "ExceptIfHeaderContainsMessageHeader"       -Value (ConvertTo-FlatString $rule.ExceptIfHeaderContainsMessageHeader)
    Write-Detail -Label "ExceptIfHeaderContainsWords"               -Value (ConvertTo-FlatString $rule.ExceptIfHeaderContainsWords)
    Write-Detail -Label "ExceptIfHeaderMatchesMessageHeader"        -Value (ConvertTo-FlatString $rule.ExceptIfHeaderMatchesMessageHeader)
    Write-Detail -Label "ExceptIfHeaderMatchesPatterns"             -Value (ConvertTo-FlatString $rule.ExceptIfHeaderMatchesPatterns)
    Write-Detail -Label "ExceptIfMessageTypeMatches"                -Value (ConvertTo-FlatString $rule.ExceptIfMessageTypeMatches)
    Write-Detail -Label "ExceptIfMessageSizeOver"                   -Value (ConvertTo-FlatString $rule.ExceptIfMessageSizeOver)
    Write-Detail -Label "ExceptIfAttachmentNameMatchesPatterns"     -Value (ConvertTo-FlatString $rule.ExceptIfAttachmentNameMatchesPatterns)
    Write-Detail -Label "ExceptIfAttachmentExtensionMatchesWords"   -Value (ConvertTo-FlatString $rule.ExceptIfAttachmentExtensionMatchesWords)
    Write-Detail -Label "ExceptIfAttachmentContainsWords"           -Value (ConvertTo-FlatString $rule.ExceptIfAttachmentContainsWords)
    Write-Detail -Label "ExceptIfAttachmentMatchesPatterns"         -Value (ConvertTo-FlatString $rule.ExceptIfAttachmentMatchesPatterns)
    Write-Detail -Label "ExceptIfAttachmentIsUnsupported"           -Value (ConvertTo-FlatString $rule.ExceptIfAttachmentIsUnsupported)
    Write-Detail -Label "ExceptIfAttachmentIsPasswordProtected"     -Value (ConvertTo-FlatString $rule.ExceptIfAttachmentIsPasswordProtected)
    Write-Detail -Label "ExceptIfAttachmentHasExecutableContent"    -Value (ConvertTo-FlatString $rule.ExceptIfAttachmentHasExecutableContent)
    Write-Detail -Label "ExceptIfAttachmentPropertyContainsWords"   -Value (ConvertTo-FlatString $rule.ExceptIfAttachmentPropertyContainsWords)
    Write-Detail -Label "ExceptIfSCLOver"                           -Value (ConvertTo-FlatString $rule.ExceptIfSCLOver)
    Write-Detail -Label "ExceptIfHasClassification"                 -Value (ConvertTo-FlatString $rule.ExceptIfHasClassification)
    Write-Detail -Label "ExceptIfHasNoClassification"               -Value (ConvertTo-FlatString $rule.ExceptIfHasNoClassification)
    Write-Detail -Label "ExceptIfAnyOfToHeader"                     -Value (ConvertTo-FlatString $rule.ExceptIfAnyOfToHeader)
    Write-Detail -Label "ExceptIfAnyOfCcHeader"                     -Value (ConvertTo-FlatString $rule.ExceptIfAnyOfCcHeader)
    Write-Detail -Label "ExceptIfAnyOfToCcHeader"                   -Value (ConvertTo-FlatString $rule.ExceptIfAnyOfToCcHeader)
    Write-Detail -Label "ExceptIfManagerAddresses"                  -Value (ConvertTo-FlatString $rule.ExceptIfManagerAddresses)
    Write-Detail -Label "ExceptIfWithImportance"                    -Value (ConvertTo-FlatString $rule.ExceptIfWithImportance)

    # ── Acciones ──
    Write-Host "   [Acciones]" -ForegroundColor DarkYellow
    Write-Detail -Label "RejectMessageReasonText"            -Value (ConvertTo-FlatString $rule.RejectMessageReasonText)        -Color "Red"
    Write-Detail -Label "RejectMessageEnhancedStatusCode"    -Value (ConvertTo-FlatString $rule.RejectMessageEnhancedStatusCode) -Color "Red"
    Write-Detail -Label "DeleteMessage"                      -Value (ConvertTo-FlatString $rule.DeleteMessage)                   -Color "Red"
    Write-Detail -Label "Disconnect"                         -Value (ConvertTo-FlatString $rule.Disconnect)                      -Color "Red"
    Write-Detail -Label "Quarantine"                         -Value (ConvertTo-FlatString $rule.Quarantine)                      -Color "Yellow"
    Write-Detail -Label "RedirectMessageTo"                  -Value (ConvertTo-FlatString $rule.RedirectMessageTo)               -Color "Yellow"
    Write-Detail -Label "AddToRecipients"                    -Value (ConvertTo-FlatString $rule.AddToRecipients)
    Write-Detail -Label "CopyTo"                             -Value (ConvertTo-FlatString $rule.CopyTo)
    Write-Detail -Label "BlindCopyTo"                        -Value (ConvertTo-FlatString $rule.BlindCopyTo)
    Write-Detail -Label "ModerateMessageByUser"              -Value (ConvertTo-FlatString $rule.ModerateMessageByUser)
    Write-Detail -Label "ModerateMessageByManager"           -Value (ConvertTo-FlatString $rule.ModerateMessageByManager)
    Write-Detail -Label "AddManagerAsRecipientType"          -Value (ConvertTo-FlatString $rule.AddManagerAsRecipientType)
    Write-Detail -Label "PrependSubject"                     -Value (ConvertTo-FlatString $rule.PrependSubject)                  -Color "Yellow"
    Write-Detail -Label "SetHeaderName"                      -Value (ConvertTo-FlatString $rule.SetHeaderName)
    Write-Detail -Label "SetHeaderValue"                     -Value (ConvertTo-FlatString $rule.SetHeaderValue)
    Write-Detail -Label "RemoveHeader"                       -Value (ConvertTo-FlatString $rule.RemoveHeader)
    Write-Detail -Label "ApplyHtmlDisclaimerLocation"        -Value (ConvertTo-FlatString $rule.ApplyHtmlDisclaimerLocation)
    Write-Detail -Label "ApplyHtmlDisclaimerText"            -Value (ConvertTo-FlatString $rule.ApplyHtmlDisclaimerText)
    Write-Detail -Label "ApplyHtmlDisclaimerFallbackAction"  -Value (ConvertTo-FlatString $rule.ApplyHtmlDisclaimerFallbackAction)
    Write-Detail -Label "SetSCL"                             -Value (ConvertTo-FlatString $rule.SetSCL)                         -Color "Yellow"
    Write-Detail -Label "ApplyClassification"                -Value (ConvertTo-FlatString $rule.ApplyClassification)
    Write-Detail -Label "ApplyRightsProtectionTemplate"      -Value (ConvertTo-FlatString $rule.ApplyRightsProtectionTemplate)
    Write-Detail -Label "SetAuditSeverity"                   -Value (ConvertTo-FlatString $rule.SetAuditSeverity)
    Write-Detail -Label "GenerateIncidentReport"             -Value (ConvertTo-FlatString $rule.GenerateIncidentReport)
    Write-Detail -Label "IncidentReportContent"              -Value (ConvertTo-FlatString $rule.IncidentReportContent)
    Write-Detail -Label "GenerateNotification"               -Value (ConvertTo-FlatString $rule.GenerateNotification)
    Write-Detail -Label "RouteMessageOutboundConnector"      -Value (ConvertTo-FlatString $rule.RouteMessageOutboundConnector)
    Write-Detail -Label "RouteMessageOutboundRequireTls"     -Value (ConvertTo-FlatString $rule.RouteMessageOutboundRequireTls)
    Write-Detail -Label "ApplyOME"                           -Value (ConvertTo-FlatString $rule.ApplyOME)
    Write-Detail -Label "RemoveOME"                          -Value (ConvertTo-FlatString $rule.RemoveOME)
    Write-Detail -Label "RemoveOMEv2"                        -Value (ConvertTo-FlatString $rule.RemoveOMEv2)
    Write-Detail -Label "RemoveRMSAttachmentEncryption"      -Value (ConvertTo-FlatString $rule.RemoveRMSAttachmentEncryption)
    Write-Detail -Label "StopRuleProcessing"                 -Value (ConvertTo-FlatString $rule.StopRuleProcessing)             -Color "Magenta"
    Write-Detail -Label "SenderNotificationType"             -Value (ConvertTo-FlatString $rule.SenderNotificationType)
    Write-Detail -Label "SmtpRejectMessageRejectText"        -Value (ConvertTo-FlatString $rule.SmtpRejectMessageRejectText)
    Write-Detail -Label "SmtpRejectMessageRejectStatusCode"  -Value (ConvertTo-FlatString $rule.SmtpRejectMessageRejectStatusCode)

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
# Exportar a CSV
# ─────────────────────────────────────────────
$reportData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host ("═" * 100) -ForegroundColor DarkCyan

# ─────────────────────────────────────────────
# Generar reporte HTML
# ─────────────────────────────────────────────
$htmlHead = @"
<style>
    body   { font-family: Segoe UI, Calibri, sans-serif; margin: 20px; background: #1e1e1e; color: #d4d4d4; }
    h1     { color: #569cd6; border-bottom: 2px solid #569cd6; padding-bottom: 8px; }
    h2     { color: #4ec9b0; margin-top: 30px; }
    h3     { color: #ce9178; }
    table  { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
    th     { background: #264f78; color: #fff; padding: 8px 12px; text-align: left; border: 1px solid #3c3c3c; }
    td     { padding: 6px 12px; border: 1px solid #3c3c3c; background: #252526; }
    tr:hover td { background: #2d2d30; }
    .enabled  { color: #6a9955; font-weight: bold; }
    .disabled { color: #f44747; font-weight: bold; }
    .summary  { display: flex; gap: 20px; margin: 15px 0; }
    .card     { background: #252526; border: 1px solid #3c3c3c; border-radius: 8px; padding: 15px 25px; text-align: center; }
    .card h3  { margin: 0; font-size: 2em; }
    .card p   { margin: 5px 0 0 0; color: #808080; }
</style>
"@

$htmlBody = "<h1>Reporte de Reglas de Flujo de Correo - Exchange Online</h1>"
$htmlBody += "<p>Generado: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>"
$htmlBody += "<div class='summary'>"
$htmlBody += "<div class='card'><h3>$totalRules</h3><p>Total</p></div>"
$htmlBody += "<div class='card'><h3 class='enabled'>$enabledRules</h3><p>Habilitadas</p></div>"
$htmlBody += "<div class='card'><h3 class='disabled'>$disabledRules</h3><p>Deshabilitadas</p></div>"
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

$htmlReport = ConvertTo-Html -Head $htmlHead -Body $htmlBody -Title "Transport Rules Report"
$htmlReport | Out-File -FilePath $htmlPath -Encoding UTF8

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
Write-Host ""
