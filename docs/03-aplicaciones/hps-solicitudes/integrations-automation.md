# HPS_Solicitudes — integraciones y automatización

## Contratos observados

| Sistema | Uso | Evidencia |
|---|---|---|
| **Lanzadera** (identidad y permisos) | identidad, permisos, entorno (vía `getdbLanzadera()` en `constructor.bas:358`) | `src/modules/constructor.bas:309` (`getUsuario`); `src/modules/Variables Globales.bas:458` (`getdb`) |
| **HPS** (gestión de usuarios HPS) | `IDUsuarioHPS` (FK conceptual) + flag `RegistroEnHPS = "Sí"` | `TempVars("RegistroEnHPS")` en `EVE`; `IDUsuarioHPS` en `TbSolicitudes` |
| **ONS** (Organismo Notificador de Seguridad) | traspasos de solicitudes a ONS con adjuntos | `Form_FormAdjuntaTraspasoONS.cls`, `URLAdjuntoEnvioONS` en `TbSolicitudes` |
| **Microsoft Excel** | el solicitante adjunta Excel con datos de la solicitud | `Form_FormAdjuntarExcelSolicitante.cls` |
| **Microsoft Word** (presumido vía plantillas HTML) | plantillas con datos mergeados | `Form_FormPlantillasHTML.cls` |
| AGEDYS | presumiblemente vía código compartido (no inspeccionado en detalle) | — |
| Correos | envío de correos automáticos con plantillas | `Correo.cls`, `CorreoOperaciones.cls`, `CorreoServicio.cls`, `TbCorreosEnviados` |
| SharePoint/ficheros | enlaces a documentación externa | `URLAdjunto`, `URLAdjuntoEnvioONS` (Memo) |
| Expedientes (Lanzadera) | `IDExpediente` (FK conceptual) | `Form_FormExpedientesBusqueda.cls`, `Form_FormExpedienteDetalle.cls` |

## Flags de operación (TempVars en `EVE`)

Flags activos hoy en producción (inferidos de `Variables Globales.bas:263-270`):

| Flag | Valor | Significado |
|---|---|---|
| `EnDesarrollo` | `"No"` (línea 257) | Modo desarrollo desactivado |
| `DatosEnLocal` | `"No"` (línea 259) | Modo local desactivado |
| `EnPruebas` | `"No"` (línea 262) | Modo pruebas desactivado (a diferencia de otras apps) |
| `ConCorreoCopiaGestor` | `"Sí"` (línea 263) | El gestor recibe copia de los correos |
| `ActivadoCorreoAutomatico` | `"Sí"` (línea 264) | Envío automático de correos |
| `RegistroEnHPS` | `"Sí"` (línea 266) | Las solicitudes se registran automáticamente en HPS |
| `ExpedienteUnificado` | `"No"` (línea 269) | Unificación de expediente desactivada (puede ser un feature futuro) |

En la nueva plataforma estos flags se mueven a **configuración del módulo** (no TempVars).

- `EnDesarrollo`, `DatosEnLocal`, `EnPruebas` → variables de entorno del runner.
- `ConCorreoCopiaGestor`, `ActivadoCorreoAutomatico`, `RegistroEnHPS`, `ExpedienteUnificado` → config del módulo (boolean).

## Integración con ONS (D101)

`Form_FormAdjuntaTraspasoONS.cls` y la columna `URLAdjuntoEnvioONS` (Memo en `TbSolicitudes`) son el **contrato de integración con ONS** (Organismo Notificador de Seguridad). El flujo:

1. La solicitud se crea en HPS_Solicitudes.
2. Se marca como pendiente de traspaso a ONS.
3. Se adjuntan los ficheros relevantes.
4. Se envía el traspaso a ONS (vía API o cola).

En la nueva plataforma, esto se traduce a un **adaptador de salida** que serializa la solicitud + adjuntos y los envía al servicio ONS. El `URLAdjuntoEnvioONS` se reemplaza por una URL firmada del object storage (D16) que ONS puede descargar.

## Integración con HPS (gestión de usuarios)

`RegistroEnHPS = "Sí"` indica que **toda solicitud crea automáticamente un registro en HPS** (la app de gestión de usuarios HPS, no HPS_Solicitudes). La columna `IDUsuarioHPS` en `TbSolicitudes` referencia al usuario HPS asociado.

En la nueva plataforma, esto se traduce a una **llamada al adaptador de HPS** (D9) que crea el usuario si no existe. La FK `IDUsuarioHPS` se mantiene como referencia conceptual o se formaliza como FK numérica.

## Integración con Lanzadera (Expedientes)

`IDExpediente` (FK conceptual) y los forms `Form_FormExpedientesBusqueda.cls` / `Form_FormExpedienteDetalle.cls` son la **integración con Lanzadera** para vincular solicitudes a expedientes. En la nueva plataforma, la integración se hace vía el **adaptador de Lanzadera** (D9-D10) o vía los endpoints que Lanzadera exponga.

## Integración con Microsoft Excel (adjuntos)

`Form_FormAdjuntarExcelSolicitante.cls` permite al solicitante **adjuntar un Excel** con datos adicionales de la solicitud. El Excel probablemente se parsea server-side (no inspeccionado en detalle) para extraer datos y validarlos.

En la nueva plataforma, esto se traduce a un **endpoint de upload** con parsing server-side (probablemente `openpyxl` en Python o `pandas.read_excel`).

## Integración con plantillas HTML (vistas web)

`Form_FormPlantillasHTML.cls` y `Form_FormWeb.cls` sugieren que HPS_Solicitudes también tiene un **sistema de plantillas HTML** (similar a Condor con `WebVisorCacheServicio` pero en versión más simple). Las plantillas se generan con datos de la solicitud.

En la nueva plataforma, esto se traduce a **Jinja2 templates** server-side (D66) con auto-escape para evitar XSS.

## Automatización

`AutomatizacionRepositorio.bas` y `Automiatizacion.bas` (typo en el nombre: "Automiatizacion" en lugar de "Automatizacion") sugieren que hay un sistema de **automatización de procesos** (probablemente traspasos automáticos, renovaciones, etc.). Migrar a la nueva plataforma como **jobs del scheduler unificado** (D59) con `kill switch` (D91).

## Configuración y flags de operación

Flags activos hoy en producción (inferidos de `Variables Globales.bas` y `getConfiguracion`):

| Flag / Campo | Valor por defecto | Significado |
|---|---|---|
| `IDAplicacion` | `"22"` (producción) | ID en `TbAplicaciones` |
| `BackendActivo` | `PROD` (producción) / `LOCAL` / `SANDBOX` | Cuál backend se usa |
| `ConCorreoCopiaGestor` | `True` (TempVar) | El gestor recibe copia de los correos |
| `ActivadoCorreoAutomatico` | `True` (TempVar) | Envío automático de correos |
| `RegistroEnHPS` | `True` (TempVar) | Las solicitudes se registran automáticamente en HPS |
| `ExpedienteUnificado` | `False` (TempVar) | Unificación de expediente desactivada |
| `PasswordBackend` | texto (presumido INI) | Contraseña del backend (⚠️ D93 cross-cutting) |

En la nueva plataforma estos flags se mueven a **configuración del módulo**:

- `BackendActivo`, `PasswordBackend` → config del puerto de persistencia / secret manager.
- `ConCorreoCopiaGestor`, `ActivadoCorreoAutomatico`, `RegistroEnHPS`, `ExpedienteUnificado` → config del módulo (boolean).
- `EnDesarrollo`, `DatosEnLocal`, `EnPruebas` → variables de entorno del runner.

Para rutas UNC, hosts y nombres de máquina concretos: **NO se reproducen en este artefacto** (regla de evidencia).