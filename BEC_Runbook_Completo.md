# Protección contra Business Email Compromise (BEC)

## Índice
1. Resumen Ejecutivo
2. ¿Qué es BEC?
3. Estrategia de Protección por Capas
4. Runbook SOC – Detección
5. Runbook SOC – Respuesta
6. KQL – Advanced Hunting

---

## 1. Resumen Ejecutivo
Business Email Compromise (BEC) es una de las principales causas de fraude financiero en organizaciones modernas. La mitigación efectiva requiere controles técnicos, identidad fuerte y disciplina operativa.

---

## 2. ¿Qué es BEC?
Ataque dirigido que utiliza correo electrónico y suplantación de identidad para provocar acciones de negocio con impacto financiero.

---

## 3. Estrategia de Protección por Capas
- SPF, DKIM, DMARC (reject)
- Microsoft Defender for Office 365
- MFA y Conditional Access
- Controles de proceso
- SOC y concientización

---

## 4. Runbook SOC – Detección
Indicadores:
- Inbox rules sospechosas
- Respuestas a hilos antiguos
- Cambios de comportamiento del remitente

---

## 5. Runbook SOC – Respuesta
1. Contener cuenta
2. Reset de credenciales
3. Eliminar reglas
4. Análisis de impacto
5. Comunicación a negocio

---

## 6. KQL – Advanced Hunting
```kql
EmailEvents
| where ThreatTypes has "BEC" or DetectionMethods has "Impersonation"
```
