# HPS_Solicitudes — seguridad y reglas

## Autorización

La matriz común se mantiene en [06-autorizacion-legacy-matriz.md](../../06-autorizacion-legacy-matriz.md). Aquí solo se conserva el comportamiento específico de HPS_Solicitudes:

- **Dos roles** (Administrador / Técnico) — más simple que las otras apps.
- Si no es Administrador (`EsAdministradorCalculado`), se evalúa `EsUsuarioTecnicoCalculado`.
- Si no es Técnico, error 1000 con mensaje "Usuario no autorizado" (`Variables Globales.bas:329-331`).
- El flujo de roles es **fail-fast** con `Err.Raise 1000`.

## Permisos por aplicación

Los permisos efectivos se cargan desde `UsuarioAplicacionPermisos` (clase compartida con Lanzadera). El acceso a HPS_Solicitudes está condicionado por `IDAplicacion = "22"` (producción). En la nueva plataforma esto se reemplaza por el **catálogo de capabilities** (D45-D46) declarado por el módulo.

## Reglas de negocio críticas

- **Validación crítica fail-fast en `EVE`**: si falla la lectura de `TbConfiguracion`, error 1000 (línea 408-411 de `Variables Globales.bas`).
- **Datos personales en `TbSolicitudes`**: 245 filas con DNI, nombres, fechas, email, teléfono (ver D98). Mismo riesgo que HPS (D92).
- **FK por email (`TbResponsables.Correo → TbSolicitudes.emailResponsable`)**: ⚠️ si el email cambia en `TbResponsables`, la FK lógica se rompe. Ver D99.
- **Traspasos a ONS**: workflow regulatorio con adjuntos (`URLAdjuntoEnvioONS`). Ver D101.
- **Registro en HPS automático**: si `RegistroEnHPS = "Sí"`, cada solicitud crea un usuario en HPS. Cross-app.
- **Flags de operación**: 7 TempVars activos en `EVE` que se traducen a config del módulo en la nueva plataforma.
- **Expediente unificado desactivado**: `ExpedienteUnificado = "No"` indica que la unificación de expediente está desactivada (posible feature futuro).
- **Patrón `.cls + .form.txt`**: igual que en Lanzadera/Expedientes/Gestion_Riesgos/NoConformidades/HPS. La nueva plataforma no usa este patrón.

## Roles y funciones diferenciadas

HPS_Solicitudes separa explícitamente **Administrador** de **Técnico** (no hay nivel "Calidad" en esta app). La diferencia con otras apps:

- **Administrador**: gestiona configuración, responsables, motivos, datos maestros. Acceso completo.
- **Técnico**: gestiona solicitudes específicas que le corresponden. Acceso limitado.

Las pantallas de administración están en `Form_Form0BDOpciones` con acceso diferenciado.

## Plantillas HTML y vistas web (similar a Condor)

`Form_FormPlantillasHTML.cls` y `Form_FormWeb.cls` sugieren que HPS_Solicitudes también tiene un **sistema de plantillas HTML** (similar al de Condor con `WebVisorCacheServicio`). Las plantillas se generan con datos de la solicitud.

En la nueva plataforma, esto se traduce a **Jinja2 templates** server-side (D66) con auto-escape para evitar XSS.

## D98 · Datos personales en `TbSolicitudes` (245 filas)

**Estado**: PROPUESTO. Detalle completo en [Matriz de migración § D98](migration-matrix.md#d98--datos-personales-en-tbsolicitudes-245-filas).

Resumen:

- `DNI`, `Nombre`, `Apellido1`, `Apellido2`, `FNacimiento`, `LugarNacimiento`, `email`, `Telefono` en `TbSolicitudes` → **245 filas con datos personales completos**.
- Riesgo análogo a HPS (D92).
- Recomendaciones: preservar sin transformaciones, enmascarar en logs, documentar en capabilities, evaluar encriptación en reposo.
- Acción operativa: **auditar el `.gitignore` del repo** para asegurar que `Solicitudes_HPS.accdb` no está siendo commiteado.

## D99 · FKs conceptuales con data integrity gaps

**Estado**: PROPUESTO. Detalle completo en [Matriz de migración § D99](migration-matrix.md#d99--fks-conceptuales-con-data-integrity-gaps).

Resumen:

- `TbResponsables.Correo → TbSolicitudes.emailResponsable`: FK por **texto email**, no por ID.
- `TbJustificaciones.idjustificacion → TbSolicitudes.idjustificacion`: dirección de FK confusa.
- Recomendaciones: agregar FK numérica en migración; normalizar dirección de FK.

## D100 · `TbLogs` vacía (presumible desuso)

**Estado**: PROPUESTO. Ver [Matriz de migración § D100](migration-matrix.md#d100--tblogs-vacía--presumible-desuso).

`TbLogs` tiene **0 filas** mientras que `TbLogsGeneral` tiene **2058 filas**. Esto sugiere que `TbLogs` está en desuso. Migrar `TbLogsGeneral` como tabla de logs principal.

## D101 · Integración con ONS

**Estado**: PROPUESTO. Ver [Matriz de migración § D101](migration-matrix.md#d101--integración-con-ons-organismo-notificador-de-seguridad).

Resumen:

- Sistema de traspasos a ONS con adjuntos (`URLAdjuntoEnvioONS`).
- Migrar como adaptador de salida (D16).
- Disponer de endpoint "traspaso a ONS" que serialice la solicitud + adjuntos.

## Patrón de testing

HPS_Solicitudes tiene **cobertura limitada** (2 archivos: `Test.bas` y `TestParametrosParser.bas`) comparado con HPS (9), NoConformidades (7) o Condor (cobertura alta). Esto indica una disciplina TDD menos madura.

**Recomendación**: en la nueva plataforma, ampliar la cobertura con **pytest** (D66) + **httpx** para integración + **Playwright** para flujos críticos (alta/renovación/traspaso a ONS).

## Edge WebView embebido (presumido)

`Form_FormWeb.cls` sugiere que HPS_Solicitudes también tiene un **control Edge embebido** similar al de Condor. La nueva plataforma web **absorbe** este patrón: las vistas son páginas nativas, no embebidas en Access.

## Riesgos de privacidad/migración

- **Datos personales** (DNI, nombres, fechas, correos, teléfonos) — 245 filas. Mismo riesgo que HPS (D92).
- **FK por email** (`emailResponsable`) — si el email cambia, la FK se rompe. Ver D99.
- **Adjuntos ONS** (`URLAdjuntoEnvioONS`) — pueden contener datos sensibles del solicitante. Manejo vía object storage (D16) con autorización.
- **Adjuntos Excel** — pueden contener datos personales parseados. Manejo cuidadoso en el parsing server-side.
- **Plantillas HTML generadas dinámicamente** — riesgo de XSS si los datos del solicitante no se escapan. La nueva plataforma usa auto-escape de Jinja2.
- **APAP y APAP_WEB** — no aparecen ni se mencionan (proyecto personal del desarrollador; regla transversal del blueprint).