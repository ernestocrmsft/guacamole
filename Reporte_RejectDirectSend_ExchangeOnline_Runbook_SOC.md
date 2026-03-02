# Control de Seguridad: RejectDirectSend en Exchange Online

**Audiencia:** Arquitectura, Messaging, SOC, SecOps, CISO  
**Nivel:** Técnico / Operativo (Enterprise)  
**Marco:** Zero Trust – Mail Flow Security

---

## 1. Resumen ejecutivo

`RejectDirectSend` es un control nativo de Exchange Online que bloquea el uso de **Direct Send**, un método SMTP anónimo históricamente utilizado por dispositivos y aplicaciones legacy. Este control elimina una de las principales superficies de ataque para **spoofing interno**, phishing lateral y fraude corporativo.

Al habilitarlo, Exchange Online rechaza el mensaje **en tiempo SMTP**, antes de que ingrese al pipeline de antispam o a Defender for Office 365.

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
- Permite suplantación interna creíble (CEO, Finanzas, RRHH)  
- Depende de SPF / DKIM / DMARC (controles posteriores, no preventivos)

---

## 3. ¿Qué hace RejectDirectSend?

```powershell
Set-OrganizationConfig -RejectDirectSend $true
```

### Lógica de evaluación

Exchange Online **rechaza el mensaje** cuando:

1. El correo llega de forma **anónima**  
2. No está asociado a un **Mail Flow Connector autenticado**  
3. El **P1 MAIL FROM** pertenece a un dominio aceptado del tenant  
4. El destinatario es un buzón interno

### Resultado

- ❌ No entra al pipeline antispam  
- ❌ No se evalúa SPF / DKIM / DMARC  
- ✅ Rechazo inmediato en SMTP

**Error típico:**

```
550 5.7.68 TenantInboundAttribution; Direct Send not allowed for this organization
```

---

## 4. Qué NO hace este control

- No valida el **P2 From header**  
- No analiza reputación  
- No depende de DMARC  
- No aplica heurística

Es un **control determinístico**, no probabilístico.

---

## 5. Impacto en seguridad (SOC view)

### Sin RejectDirectSend

- Phishing interno sin compromiso de identidad  
- Correos spoofeados pueden llegar a Inbox / Junk  
- Alto riesgo de fraude financiero

### Con RejectDirectSend

- Bloqueo total de spoofing interno por SMTP  
- Reducción inmediata de superficie de ataque  
- Control alineado a Zero Trust

---

## 6. Impacto operativo en aplicaciones

### Flujos que se rompen

- Impresoras / scanners  
- ERPs / HR legacy  
- Scripts SMTP antiguos  
- SaaS mal configurados

### Alternativas soportadas

- ✅ Mail Flow Connector autenticado por **certificado** (recomendado)  
- ✅ Mail Flow Connector por **IP fija**  
- ✅ SMTP AUTH con cuenta dedicada (último recurso)

---

## 7. Estado del control

| Propiedad | Valor |
|---------|------|
| Default | false |
| GA | Septiembre 2025 |
| Propagación | ~30 minutos |

Verificación:

```powershell
Get-OrganizationConfig | Select RejectDirectSend
```

---

# RUNBOOK SOC – Direct Send / RejectDirectSend

## Objetivo

Detectar y responder a intentos de uso de Direct Send y validar que el control esté bloqueando correctamente intentos de spoofing interno.

---

## Detección – Qué buscar

### Indicadores clave

- Errores SMTP `5.7.68 TenantInboundAttribution`  
- Correos internos con:
  - `SenderFromDomain` = dominio corporativo  
  - `AuthenticationDetails` = vacío  
  - `ConnectorId` = null

---

## Respuesta – Playbook

1. **Confirmar intento**  
   Revisar Message Trace / Advanced Hunting
2. **Clasificar origen**  
   IP, dispositivo, aplicación
3. **Decisión**  
   ✅ App legítima → Migrar a Connector autenticado  
   ❌ Origen desconocido → Bloqueo permanente
4. **Acción correctiva**  
   Crear Mail Flow Connector y documentar excepción
5. **Lección aprendida**  
   Actualizar inventario de apps y revisar SPF / DKIM / DMARC

---

# KQL – Detección histórica de Direct Send

## 1. Correos internos anónimos (indicador Direct Send)

```kql
EmailEvents
| where SenderFromDomain == RecipientEmailDomain
| where isempty(ConnectorId)
| where isempty(AuthenticationDetails)
| project Timestamp, NetworkMessageId, SenderFromAddress, RecipientEmailAddress, SenderIPv4, Subject
```

## 2. Intentos bloqueados por RejectDirectSend

```kql
EmailEvents
| where ActionType == "Reject"
| where ErrorCode has "5.7.68"
| project Timestamp, SenderFromAddress, RecipientEmailAddress, SenderIPv4, ErrorCode
```

## 3. Top IPs intentando Direct Send

```kql
EmailEvents
| where SenderFromDomain == RecipientEmailDomain
| where isempty(ConnectorId)
| summarize Attempts=count() by SenderIPv4
| order by Attempts desc
```

---

## Recomendación final enterprise

✔ Habilitar `RejectDirectSend` en todos los tenants  
✔ Migrar aplicaciones a conectores autenticados  
✔ Complementar con SPF estricto, DKIM y DMARC `p=reject`  
✔ Monitorear continuamente desde SOC

---

**Este control convierte Exchange Online en un modelo de correo interno Zero Trust por diseño.**
