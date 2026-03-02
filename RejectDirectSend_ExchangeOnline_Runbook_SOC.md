# Control de Seguridad: RejectDirectSend en Exchange Online

**Audiencia:** Arquitectura, Messaging, SOC, SecOps, CISO

**Nivel:** Técnico / Operativo (Enterprise)

**Marco:** Zero Trust – Mail Flow Security

---

## 1. Resumen ejecutivo

`RejectDirectSend` es un control nativo de Exchange Online que **bloquea el uso de Direct Send**, un método SMTP anónimo históricamente utilizado por dispositivos y aplicaciones legacy. Este control elimina una de las **principales superficies de ataque para spoofing interno**, phishing lateral y fraude corporativo.

Al habilitarlo, **Exchange Online rechaza el mensaje en tiempo SMTP**, antes de que ingrese al pipeline de antispam o a Defender for Office 365.

---

## 2. ¿Qué es Direct Send?

**Direct Send** permite enviar correos a buzones internos del tenant usando:

- SMTP puerto **25**
- Destino: `tenant.mail.protection.outlook.com`
- **Sin autenticación** (anónimo)
- Dominio del remitente (**P1 MAIL FROM**) pertenece a un *accepted domain*

Diseñado para:
- Impresoras
- Scanners
- Aplicaciones legacy on‑prem

### Riesgo inherente

- No requiere compromiso de cuenta
- Permite **suplantación interna creíble** (CEO, Finanzas, RRHH)
- Depende de SPF / DKIM / DMARC (controles posteriores, no preventivos)

---

## 3. ¿Qué hace RejectDirectSend?

```powershell
Set-OrganizationConfig -RejectDirectSend $true
```

Exchange Online rechaza mensajes SMTP anónimos cuyo **P1 MAIL FROM** pertenece a un dominio aceptado y no está asociado a un Mail Flow Connector autenticado.

---

# RUNBOOK SOC – Direct Send / RejectDirectSend

## Objetivo

Detectar y responder a intentos de spoofing interno vía Direct Send.

## Detección

```kql
EmailEvents
| where SenderFromDomain == RecipientEmailDomain
| where isempty(ConnectorId)
| where isempty(AuthenticationDetails)
```

## Respuesta

1. Confirmar intento
2. Clasificar origen
3. Migrar app legítima a connector autenticado
4. Bloquear origen desconocido

---

## KQL – Intentos bloqueados

```kql
EmailEvents
| where ActionType == "Reject"
| where ErrorCode has "5.7.68"
```

---

**Zero Trust Mail Flow – Enterprise Security**
