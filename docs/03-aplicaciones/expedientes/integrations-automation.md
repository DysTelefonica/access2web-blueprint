# Expedientes — integraciones y automatización

## Contratos observados

| Sistema | Uso | Evidencia |
|---|---|---|
| Lanzadera | usuario, permisos e ID de aplicación 19 | `EVE`, `constructor.getUsuario`, `TbUsuariosAplicaciones` vinculado |
| AGEDYS | DPD/proyectos y datos de expediente | `ExpedienteAGEDYS`, `constructor`, tablas enlazadas documentadas en ERD |
| HPS | usuarios y relaciones HPS | `getdbHPS`, alta HPS |
| Riesgos | proyectos por expediente/código | `GestionRiesgos`, `getGestionRiesgos` |
| No conformidades | NC por expediente/código S4H | helpers de operaciones |
| Correos | avisos y errores administrativos | `CorreoOperaciones`, `TbCorreosEnviados`/contrato externo |
| SharePoint/ficheros | enlaces de documentación | `AccesoSharepoint`, anexos |
| JSON/Excel | intercambio y reporting | exportadores y selector de consultas |

## Automatización y tareas

`PintarTareas` mantiene seis categorías operativas y actualiza menús, contadores y bandeja. `Entorno` usa carga diferida y puede trabajar con datos en memoria. La exportación E2E añade selección por sesión, configuración de destino por usuario, batch, detalle e historial.

Para procesos batch y automatizaciones sanitizados, enlazar [procesos-batch-y-automatizaciones.md](../../04-integraciones-y-operacion/procesos-batch-y-automatizaciones.md); no se duplica aquí.
