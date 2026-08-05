# Condor — seguridad y reglas

## Autorización

La matriz común se mantiene en [06-autorizacion-legacy-matriz.md](../../06-autorizacion-legacy-matriz.md). Aquí solo se conserva el comportamiento específico de Condor:

- **Tres roles** (Administrador, Calidad, Técnico), calculados vía `rolUsuario` y `rolUsuarioReal`. La separación `rolUsuario` vs `rolUsuarioReal` indica **preparación para impersonación** (preparación actual: variables declaradas pero no implementadas).
- `JerarquiaRolesHelper.bas`: lógica de jerarquía de roles.
- `CondorError.cls`: clase de error tipada con `.Create`, `.Raise` — patrón de error handling estructurado.
- Las pantallas técnicas son de solo lectura; el acceso a workflows específicos está condicionado por capacidades del rol.

## Permisos por aplicación

Los permisos efectivos se cargan desde `UsuarioAplicacionPermisos` (clase compartida con Lanzadera). El acceso a Condor está condicionado por `IDAplicacion = "23"` (producción). En la nueva plataforma esto se reemplaza por el **catálogo de capabilities** (D45-D46) declarado por el módulo.

## Reglas de negocio críticas

- **Validación crítica fail-fast en `EVE`**: si falla la conexión con el backend, `Application.Quit` (línea 151 de `Variables Globales.bas`). Si falla la validación del usuario, `Err.Raise 513`. Patrón defensivo robusto.
- **4 tipos de Solicitud** (`PC`, `CD_CA`, `CD_CA_SUB`, `PC_SUB`): cada uno con su propio Servicio, Repositorio y ViewModel. La nueva plataforma debe mantener la distinción como **4 dominios dentro del módulo**, no como módulos separados.
- **Vinculación con NoConformidades**: `Form_frmGestionSolicitud.cls:1908` verifica `ncServ.getNoConformidadPorCodigoCondor` antes de eliminar. La eliminación está condicionada por la existencia de NCs vinculadas.
- **Vinculación con Expedientes**: `idExpediente` (FK conceptual). Sin constraint físico.
- **Transacciones de workflow**: `WorkflowServicio.cls` orquesta transiciones de estado con log en `tbLogEstados`.
- **Validación de calidad**: `revisionCalidadEstado` + `revisionCalidadComentarios` en `tbSolicitudes`. `ValidacionRevisionServicio.cls` orquesta el flujo.
- **Rechazos con historial**: `tbRechazos` + `tbHistorialRechazos`. Trazabilidad completa.
- **Snapshot para WebView**: `SnapshotServicio.cls` toma snapshot del estado actual. Migración: traducir a endpoint de snapshot server-side.
- **Datos por tipo**: `DatosPC`, `DatosCDCA`, `DatosCDCASUB`, `DatosPCSUB`. Cada uno tiene su propia tabla y servicios.

## Roles y funciones diferenciadas

Condor separa explícitamente **Calidad** de **Técnico**:

- **Calidad**: revisión de calidad, validación de workflow, gestión de rechazos. Acceso a `Form_frmGestionSolicitud` para `revisionCalidadEstado` y `revisionCalidadComentarios`.
- **Técnico**: alta/edición de Solicitudes, subida de adjuntos. Acceso a `Form_frm0PpalTecnico` + `Form_frmBuscarSolicitudes`.

## Transaccionalidad

Las operaciones de Solicitud (alta, edición, cambio de estado) usan transacciones DAO implícitas vía `UsuarioServicio` (análogo a las otras apps). La nueva plataforma usa transacciones SQLAlchemy `AsyncSession.begin()` (D66, D82).

## D93 · Password hardcoded como fallback en `GetPasswordDB`

**Estado**: PROPUESTO. **CRÍTICO** ⚠️⚠️⚠️. Detalle completo en [Matriz de migración § D93](migration-matrix.md#d93--password-hardcodeado-como-fallback-en-getpassworddb).

Resumen:

- `FUNCIONES UTILES.bas:150` tiene `GetPasswordDB = "dpddpd"` como fallback hardcoded.
- Si la lectura del INI falla, **cae al literal `"dpddpd"`**.
- Riesgo: bypass de rotación de contraseña, exposición en código fuente.
- **Recomendaciones**: eliminar fallback, fallar explícitamente; rotar contraseña; auditar otras apps (cross-cutting).
- **Regla HR-3**: cero secretos en código fuente, cero secretos en repos.

## D94 · FKs conceptuales sin constraint

**Estado**: PROPUESTO. Detalle completo en [Matriz de migración § D94](migration-matrix.md#d94--fks-conceptuales-sin-constraint).

Resumen:

- `tbSolicitudes.idEstadoInterno` debería tener FK a `tbEstados` pero **NO la tiene**. Data integrity gap.
- `tbSolicitudes.idExpediente` (Lanzadera), `idNCAsociada` (NoConformidades) son FKs conceptuales cross-app.
- `tbTransiciones`, `tbLogEstados` deberían tener FK a `tbSolicitudes` y `tbEstados` pero **NO las tienen**.

Recomendación: formalizar las FKs intra-app en PostgreSQL; mantener las FKs cross-app como referencias conceptuales (D86/D87).

## D95 · Vinculación con NoConformidades

**Estado**: PROPUESTO. Detalle completo en [Matriz de migración § D95](migration-matrix.md#d95--vinculación-con-noconformidades).

Resumen:

- `Form_frmGestionSolicitud.cls:1908` verifica NC vinculada antes de eliminar.
- `idNCAsociada` (FK conceptual) + `NoConformidadServicio.getNoConformidadPorCodigoCondor`.
- Relación por código, no por ID directa.
- **Recomendación**: mantener como referencia conceptual; orden de migración estricto (NoConformidades antes que Condor).

## Edge WebView embebido — implicaciones de seguridad

`Me.webInfo.Navigate rutaTemporal` carga HTML generado dinámicamente. Esto es un vector de XSS potencial si el HTML incluye datos del usuario sin escape. La nueva plataforma usa HTMX + Jinja2 con auto-escape (D67) — la seguridad mejora nativamente.

`SnapshotServicio.cls` escribe ficheros `.html` en disco. Si los snapshots no se sanitizan, **pueden exponer datos sensibles**. La nueva plataforma usa endpoints server-side con autorización (D27, D45) — los snapshots se sirven autenticados.

## Patrón de testing sandbox (preservar como referencia)

`FUNCIONES UTILES.bas:84-120` implementa un patrón maduro de testing sandbox:

- `m_TestingMode=True` enruta a `m_BackendSandboxURL`.
- Cache safety: cierre y reapertura si la conexión apunta a otra URL.
- Fail-fast: `Err.Raise 513 "TESTS BLOCKED"` si el sandbox no está configurado.
- `m_TestOnly*` overrides para tests.

**Se preserva como referencia del puerto de testing** de la nueva plataforma (FastAPI + pytest + fixtures con BD aislada).

## Riesgos de privacidad/migración

- **Datos personales**: `usuarioCreacion` y `usuarioModificacion` en `tbSolicitudes` son nombres de usuario (no datos personales sensibles pero sí auditables).
- **Comentarios de revisión**: `revisionCalidadComentarios` (Memo) puede contener datos sensibles. Revisar antes de portar.
- **Adjuntos**: rutas externas a ficheros. Manejo vía object storage (D16).
- **APAP y APAP_WEB**: no aparecen ni se mencionan (proyecto personal del desarrollador; regla transversal del blueprint).