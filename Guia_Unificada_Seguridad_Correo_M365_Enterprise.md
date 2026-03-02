# Seguridad de Correo en Microsoft 365 – Guía Unificada Enterprise

**Audiencia:** Arquitectura, Messaging, SOC, SecOps, CISO  
**Nivel:** Técnico / Operativo (Enterprise)  
**Marco:** Zero Trust – Mail Flow Security

---

## Índice

1. [Resumen ejecutivo](#1-resumen-ejecutivo)  
2. [Reglas básicas de flujo de correo – Microsoft 365](#2-reglas-básicas-de-flujo-de-correo--microsoft-365)  
3. [Control de Seguridad: RejectDirectSend en Exchange Online](#3-control-de-seguridad-rejectdirectsend-en-exchange-online)  
4. [SPF, DKIM, DMARC y MTA-STS](#4-spf-dkim-dmarc-y-mta-sts)  
5. [Runbook SOC – Detección y Respuesta](#5-runbook-soc--detección-y-respuesta)  
6. [Recomendaciones finales Enterprise](#6-recomendaciones-finales-enterprise)

---

## 1. Resumen ejecutivo

Una postura de seguridad sólida en Microsoft 365 requiere **controles preventivos en múltiples capas**. Esta guía unifica:

- Reglas de flujo de correo (mail flow rules)
- El control nativo `RejectDirectSend`
- Autenticación y protección de dominio (SPF, DKIM, DMARC, MTA-STS)

En conjunto, estos controles eliminan **spoofing interno**, reducen **phishing**, previenen **fraude corporativo** y alinean Exchange Online a un modelo **Zero Trust Mail Flow**.

---

## 2. Reglas básicas de flujo de correo – Microsoft 365

Las reglas de flujo de correo permiten bloquear vectores comunes de abuso y errores de configuración.

### Objetivos

- Bloquear correos enviados a dominios técnicos (`*.onmicrosoft.com`)
- Enviar a cuarentena mensajes que no pueden ser inspeccionados

> Estas reglas reducen superficie de ataque y errores operativos comunes.

---

## 3. Control de Seguridad: RejectDirectSend en Exchange Online

`RejectDirectSend` es un control nativo de Exchange Online que **bloquea el uso de Direct Send**, un método SMTP anónimo históricamente utilizado por dispositivos y aplicaciones legacy.

### Direct Send – Riesgo

- SMTP puerto 25 hacia `tenant.mail.protection.outlook.com`
- Sin autenticación
- Dominio remitente (P1 MAIL FROM) = dominio aceptado

Permite **suplantación interna** (CEO, Finanzas, RRHH) sin compromiso de identidad.

### Funcionamiento de RejectDirectSend

```powershell
Set-OrganizationConfig -RejectDirectSend $true
```

Exchange Online rechaza en tiempo SMTP mensajes:

- Anónimos
- No asociados a Mail Flow Connector
- P1 MAIL FROM = dominio del tenant

**Error típico:**
```
550 5.7.68 TenantInboundAttribution; Direct Send not allowed for this organization
```

No entra al pipeline antispam ni evalúa SPF/DKIM/DMARC.

---

## 4. SPF, DKIM, DMARC y MTA-STS

Estos mecanismos protegen la **identidad del dominio** y la **entregabilidad**.

### SPF

Define qué servidores pueden enviar correo por el dominio.

Ejemplo:
```
v=spf1 include:spf.protection.outlook.com -all
```

### DKIM

Firma criptográficamente los mensajes para asegurar integridad y autenticidad del dominio.

### DMARC

Orquesta SPF y DKIM, define políticas (`none`, `quarantine`, `reject`) y genera reportes.

Ejemplo recomendado:
```
v=DMARC1; p=reject; adkim=s; aspf=s; pct=100
```

### MTA-STS

Fuerza TLS en tránsito entre MTAs y evita ataques MITM.

---

## 5. Runbook SOC – Detección y Respuesta

### Detección Direct Send

```kql
EmailEvents
| where SenderFromDomain == RecipientEmailDomain
| where isempty(ConnectorId)
| where isempty(AuthenticationDetails)
```

### Intentos bloqueados RejectDirectSend

```kql
EmailEvents
| where ActionType == "Reject"
| where ErrorCode has "5.7.68"
```

### Respuesta

1. Confirmar evento (Message Trace / Advanced Hunting)
2. Clasificar origen (IP, app, dispositivo)
3. Migrar apps legítimas a Connector autenticado
4. Bloquear origen no autorizado
5. Actualizar inventario y controles

---

## 6. Recomendaciones finales Enterprise

✔ Habilitar `RejectDirectSend` en todos los tenants  
✔ Usar Mail Flow Connectors autenticados  
✔ Implementar SPF, DKIM y DMARC en modo `reject`  
✔ Habilitar MTA-STS y TLS-RPT  
✔ Monitorear continuamente desde SOC

---

**Exchange Online + Zero Trust Mail Flow = Eliminación efectiva del spoofing interno.**
