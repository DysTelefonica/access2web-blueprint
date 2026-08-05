# Gestion_Riesgos — integraciones y automatización

## Contratos observados

| Sistema | Uso | Evidencia |
|---|---|---|
| **Lanzadera** (acoplamiento directo) | identidad, permisos y entorno (`getdbLanzadera()` consulta la BD de Lanzadera desde `Constructor.bas`) | `src/modules/Constructor.bas` (callers de `getdbLanzadera`); `src/modules/Variables Globales.bas:485` (`getdb`) |
| Expedientes | vínculo por código de riesgo / expediente | `Expediente.cls`, `ExpedienteEntidad.cls`, `ExpedienteResponsable.cls` (clases compartidas con Lanzadera) |
| HPS | pedidos vinculados | `Pedido.cls` (presumido) |
| No Conformidades | NC por código S4H / expediente | `NC.cls` (presumido) |
| AGEDYS | DPD/proyectos (vía código compartido) | `ExpedienteAGEDYS.cls` y entidades compartidas |
| Correos | avisos de error al administrador, publicabilidad | `Correo.cls`, `EdicionCorreoRevision.cls` |
| SharePoint/ficheros | enlaces de documentación | referencias en anexos |
| Informes | generación HTML/PDF/Excel | `InformeRiesgoPDFServicio.cls`, `Form_FormInformeTipoSalida.cls` |

## Acoplamiento declarado con Lanzadera

La aplicación consulta **directamente** la base de datos de Lanzadera desde `Constructor.getUsuario` y posiblemente otros callers de `getdbLanzadera()`. Este acoplamiento NO se resuelve con un servicio web; es un acceso JDBC/ODBC directo a `Lanzadera_Datos.accdb`.

Implicaciones para la nueva plataforma:

- La identidad debe resolverse vía el **adaptador unificado de autenticación** (D9–D10), no vía un DAO directo a Lanzadera.
- Los permisos por aplicación se cargan desde `UsuarioAplicacionPermisos` con `IDAplicacion = "5"` (producción) o `"51"` (pruebas).
- `getdbLanzadera()` se reemplaza por una llamada al servicio de identidad / permisos expuesto por la plataforma.

## Automatización y tareas

`PintarTareas` mantiene worklists de calidad y técnico con seis categorías operativas aproximadas (aceptado/retirado/visado, materializado, retipificación, detalle de edición). `Entorno` usa carga diferida y puede trabajar con datos en memoria.

En la nueva plataforma:

- El patrón de actualización de contadores pendientes vía polling HTMX con `hx-trigger="every 30s"` (D69) aplica también a Gestion_Riesgos.
- Las tareas de calidad y técnico se traducen a queries indexadas con sus propios endpoints.
- El envío de avisos (correos) pasa por el servicio unificado de notificaciones (D11–D13).

Para procesos batch y automatizaciones sanitizados, enlazar [procesos-batch-y-automatizaciones.md](../../04-integraciones-y-operacion/procesos-batch-y-automatizaciones.md); no se duplica aquí.

## Configuración y flags de operación

Flags activos hoy en producción (ver `Variables Globales.bas:247-272`):

| Flag | Valor actual | Significado |
|---|---|---|
| `CadenaJerarquicaModelo` | `"nuevo"` (línea 248) | modelo de carga de cadena jerárquica; el valor `"antiguo"` queda en código comentado (línea 249) |
| `JPMesesAvisoEntreEdiciones` | `3` | meses entre ediciones para disparar aviso |
| `JPDiasPreviosParaElAviso` | `15` | días previos al aviso |
| `CalDiaInicialMesAviso` | `2` | día del mes en que se publica el calendario |
| `Publicabilidad_Usar_Cache` | `"No"` | deshabilitado en producción (línea 254; valor `"Sí"` comentado) |
| `DatosEnLocal` | `"No"` | modo local deshabilitado |
| `EnDesarrollo` | `"No"` | deshabilitado en producción |
| `EnPruebas` | `"No"` | pruebas deshabilitadas; cuando `"Sí"`, IDAplicacion pasa a `"51"` |

En la nueva plataforma estos flags se mueven a **configuración del módulo** (no TempVars); algunos se reinterpretan:

- `CadenaJerarquicaModelo` se descarta (D88): el modelo único es CTE recursivo.
- `JPMesesAvisoEntreEdiciones`, `JPDiasPreviosParaElAviso`, `CalDiaInicialMesAviso` se conservan como configuración del scheduler de avisos.
- `Publicabilidad_Usar_Cache`, `DatosEnLocal`, `EnDesarrollo`, `EnPruebas` se descartan o se mueven a configuración del módulo (no TempVars).

Para rutas UNC, hosts y nombres de máquina concretos: **NO se reproducen en este artefacto** (regla de evidencia). Las rutas operativas se mantienen en `04-integraciones-y-operacion/rutas-entornos-y-contingencia.md` con criterio de privacidad.