# Capacidad: Seguridad, administración e infraestructura

## §0 Identidad

- **ID de capacidad**: CAP-010
- **Tier**: critical
- **Estado**: active con deuda de identificación de usuarios externos y pruebas focales de selección de backend
- **Source**: hybrid
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: `dysflow.verify_binary` sobre seguridad, administración, catálogos, mapeos, globals y `frm0OtrosAdmin`: `actionableOk=true`. `dysflow.get_schema` confirmó que `TbConfiguracionBackends` vive en el frontend `CONDOR.accdb`, no en `condor_datos.accdb`, y documentó su esquema real.
- **Confianza global**: `Verified-runtime` parcial. La jerarquía de roles y la selección de backend están razonablemente implementadas y fuente↔binario está sincronizado, pero la identidad del usuario se basa en `Wscript.Network.UserName` (no en Kerberos/SAML) y no hay manifest focal. La promoción a `Verified-runtime` depende de pruebas con seams.
- **Deuda TDD v2.4.2 transversal**: ver [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md). Seguridad/administración no debe promocionarse a `Verified-runtime` sin seam de `getdb*()` y suite focal.

**Contrato TDD vigente**: las pruebas de seguridad/administración deben migrarse a `access-vba-tdd` v2.4.2 — `Public Function` con retorno JSON canónico, fixture propio, schema-first, `DAO.Database` inyectado, manifests atómicos, cero mutación de `TbConfiguracionBackends` por tests.

**Justificación del nivel**: crítico, porque gobierna la identidad, los permisos por rol, la selección de backend (PRD / sandbox), el mapa de plantillas documentales y la infraestructura de pruebas.

## §1 Intención de negocio — POR QUÉ

- **Propósito**: autenticar al usuario, calcular su rol efectivo (Administrador/Calidad/Técnico/otros), exponer la matriz de permisos por aplicación, gestionar el catálogo de No Conformidades, el catálogo de Suministradores, el mapa de campos para plantillas Word y la configuración de backend (`TbConfiguracionBackends`).
- **Usuarios / perfiles**: administrador (configuración, alta/baja de usuarios), calidad, técnicos, RAC, secretaría, economía y todo el catálogo de `TbUsuariosAplicacionesPermisos`.
- **Problema que resuelve**: centraliza la identidad, los permisos y la infraestructura de backends/plantillas; permite que el resto de capacidades (CAP-001/002/003/004/005/006/007/008/009) se apoyen en una sola fuente de verdad para la sesión.
- **Valor de negocio**: seguridad por rol, configuración externa sin tocar binario, gemelos (PC, PCSUB, CD/CA, CDCASUB) que comparten la misma infraestructura de identidad.
- **No-objetivos**: este documento no cubre el detalle de cada capacidad de negocio. Sí las vincula.
- **Origen de la intención**: código actual, PRD `00_PRD_CONDOR` y SDD históricos de la feature de backend seleccionable.
- **Referencia de tracker de origen**: Pendiente de confirmación.

## §2 Contrato de comportamiento — QUÉ

### Escenarios principales

- **DADO** un usuario que abre la app **CUANDO** se llama `UsuarioServicio.getUsuarioConectadoConPermisos` **ENTONCES** se obtiene `Wscript.Network.UserName`, se aplica el mapeo de alias (`local1`/`adm1` → `adm`) y se compone un objeto `usuario` con `ColAplicacionesPermisos`, `Permisos` específicos para `IDAplicacion` y `rol` calculado por `DeterminarRol`. **Estado**: `Verified-static`.
- **DADO** un usuario `EsAdministrador = "Sí"` **CUANDO** se evalúa su rol **ENTONCES** se devuelve `rol.Administrador`. Si no, si `Permisos.EsUsuarioCalidad = "Sí"`, `rol.Calidad`. En caso contrario, `rol.Tecnico`. **Estado**: `Verified-static`.
- **DADO** un usuario en cualquier rol **CUANDO** se quiere obtener la lista de responsables técnicos **ENTONCES** `getResponsablesTecnicos` consulta `TbExpedientesResponsables` y `TbUsuariosAplicaciones` con `EsJefeProyecto = "Sí"`, `Pecal = "Sí"` y `FechaBaja Is Null` desde `getdbExpedientes()`. **Estado**: `Verified-static`.
- **DADO** un usuario **CUANDO** se quiere obtener la lista de responsables de calidad **ENTONCES** `getResponsablesCalidad` consulta `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos` con `IDAplicacion = IDAplicacion` y `EsUsuarioCalidad = "Sí"` desde `getdbLanzadera()`. **Estado**: `Verified-static`.
- **DADO** un expediente y un suministrador **CUANDO** se necesita un mapeo de campos Word **ENTONCES** `MapeoServicio.getMapeoPC/CDCA/CDCASUB/PCSUB` consulta `tbMapeoCampos` por nombre de plantilla a través de `getdb()`. **Estado**: `Verified-static`; ver CAP-002 §3.
- **DADO** una No Conformidad externa **CUANDO** se necesita su catálogo **ENTONCES** `NoConformidadServicio` lee desde `getdbNoConformidades()` (BD externa, no CONDOR). **Estado**: `Verified-static`; el seam entre BD externa y CONDOR es un riesgo arquitectónico.
- **DADO** la necesidad de conmutar entre backends (PRD, sandbox, etc.) **CUANDO** se modifica `TbConfiguracionBackends.BackendActivo` **ENTONCES** `getdb()` resuelve a la nueva ruta. **Estado**: `Verified-static`; el mecanismo exacto de conmutación vive en `FUNCIONES UTILES.bas` y es leído por `getdb()`.
- **DADO** la aplicación **CUANDO** se crea una entidad del dominio **ENTONCES** `Factoria.CreateEntity(nombre)` devuelve la instancia adecuada. **Estado**: `Verified-static`; lista cerrada de tipos en `Factoria.bas`.
- **DADO** un usuario con rol Administrador **CUANDO** abre `frm0OtrosAdmin` **ENTONCES** el formulario expone opciones de configuración. **Estado**: `Verified-static`; el contenido exacto no se ha auditado en este pase.

### Reglas de negocio

| ID regla | Enunciado | Autoridad | ¿Aplicada en código? | Prueba / evidencia | Confianza |
|---|---|---|---|---|---|
| BR-001 | El usuario se identifica por `Wscript.Network.UserName` con mapeo de alias (`local1`/`adm1` → `adm`). | Código | Sí: `UsuarioServicio.getUsuarioConectadoConPermisos` líneas 53-77. | Pendiente. | Verified-static |
| BR-002 | La jerarquía de roles efectiva es: `EsAdministrador = "Sí"` → `rol.Administrador`; si no, `EsUsuarioCalidad = "Sí"` → `rol.Calidad`; resto → `rol.Tecnico`. | Código | Sí: `UsuarioServicio.DeterminarRol` líneas 79-109. | Pendiente. | Verified-static |
| BR-003 | Los responsables técnicos se leen de `getdbExpedientes()`; los responsables de calidad de `getdbLanzadera()`. | Código | Sí: `UsuarioServicio.getResponsablesTecnicos` y `getResponsablesCalidad`. | Pendiente. | Verified-static |
| BR-004 | `MapeoServicio.getMapeo*` lee `tbMapeoCampos` por nombre de plantilla (`PC`, `CDCA`, `CDCASUB`, `PCSUB`) usando `getdb()`. | Código | Sí: `MapeoServicio.cls` líneas 17-59. | Pendiente. | Verified-static |
| BR-005 | `NoConformidadRepositorio` lee y escribe contra `getdbNoConformidades()` (BD externa). | Código | Sí: `NoConformidadRepositorio.bas` (líneas 15-102). | Pendiente. | Verified-static |
| BR-006 | `SuministradorServicio` ofrece `getSuministradorPorID` y `getSuministradoresPorExpediente` para alimentar `getNombresSuministradores` de CAP-006. | Código | Sí: `SuministradorServicio.cls`. | Pendiente. | Verified-static |
| BR-007 | `Factoria.CreateEntity` soporta 19+ nombres de clase; cualquier nombre no listado devuelve `Nothing`. | Código | Sí: `Factoria.bas`. | Pendiente. | Verified-static |
| BR-008 | `getdb()`/`getdbLanzadera()`/`getdbExpedientes()`/`getdbNoConformidades()`/`getdbCorreo()` son singletons de proceso; cualquier cambio en `TbConfiguracionBackends.BackendActivo` se aplica en el siguiente acceso. | Código + AGENTS | Sí: `FUNCIONES UTILES.bas` líneas 64-199. | Pendiente. | Verified-static |
| BR-009 | `Variables Globales.bas` declara los globals canónicos: `m_ObjUsuarioConectado`, `m_ObjEntorno`, `m_ObjUsuarioActivo`, `m_ObjUsuarioReal`, `rolUsuario`, `g_blnImpersonando`, `g_objLastError`, `g_dbLanzadera`, `g_dbExpedientes`, `g_dbNoConformidades`, `g_dbCorreo`, `g_dbCondor`. | Código + AGENTS | Sí: `Variables Globales.bas` y `Entorno.cls`. | Pendiente. | Verified-static |
| BR-010 | `modDevTools.bas` y `modActualizaciones.bas` son utilidades de administración/migración. | Código | Sí: `modDevTools.bas`, `modActualizaciones.bas`. | Pendiente. | Verified-static |
| BR-011 | `IDAplicacion` es la constante canónica para el identificador de la aplicación CONDOR en `TbUsuariosAplicacionesPermisos`. | Código | Sí: referenciada por `UsuarioServicio.getResponsablesCalidad` y otros. | Pendiente. | Verified-static |

### Validaciones observadas

- `getUsuarioConectadoConPermisos` lanza `Err.Raise 513` si el objeto `usuario` es `Nothing` al determinar el rol.
- `Factoria.CreateEntity` devuelve `Nothing` para nombres no listados; los llamadores deben comprobarlo.
- `MapeoRepositorio` usa `On Error GoTo Errores` y expone `p_Error` para capturar fallo de `getdb()`.

### Transiciones de estado y navegación

- `TbConfiguracionBackends` no se modifica por transiciones de workflow; su cambio es administrativo.
- `Usuario` y `Permisos` son read-only durante una sesión: cualquier cambio requiere alta/baja por el administrador.

### Casos límite y hallazgos

- La identidad se basa en `Wscript.Network.UserName` (alias de Windows). No hay integración con Kerberos/SAML/OAuth. Migrar a web requerirá un IdP.
- `MapeoServicio.SembrarBookmarksDesdePlantilla` interactúa con Word COM (no es transaccional ni testeable sin Word). Conviene un seed-offline.
- `NoConformidadRepositorio` opera sobre una BD externa no versionada en este repo. Cualquier cambio de esquema allí requiere coordinación con el equipo de NCs.
- `Factoria.CreateEntity` tiene una lista cerrada; añadir un nuevo tipo requiere editar este módulo. Es un sealamiento central que conviene sustituir por un registro declarativo.
- `TbConfiguracionBackends` está en el frontend `CONDOR.accdb`, no en el backend `condor_datos.accdb`. Esto es deliberado pero importante: la selección de backend es configuración local de frontend, no dato de negocio del backend activo.

### Señales de aceptación / presencia

- Existen `Usuario.cls`, `UsuarioServicio.cls`, `UsuarioAplicacionPermisos.cls`, `UsuarioRepositorio.bas`.
- Existen `Entorno.cls`, `Factoria.bas`, `Variables Globales.bas`, `FUNCIONES UTILES.bas`, `modDevTools.bas`, `modActualizaciones.bas`.
- Existen `NoConformidad.cls`, `NoConformidadServicio.cls`, `NoConformidadRepositorio.bas`.
- Existen `Suministrador.cls`, `SuministradorServicio.cls`, `SuministradorRepositorio.bas`.
- Existen `MapeoCampos.cls`, `MapeoServicio.cls`, `MapeoRepositorio.bas` (vinculados a CAP-002).
- `Form_frm0OtrosAdmin.cls` y `Form_frm0OtrosAdmin.form.txt` están sincronizados fuente↔binario.

## §3 Mapa de implementación — CÓMO

- **Puntos de entrada de UI**:
  - `Form_frm0OtrosAdmin.Form_Load` y opciones administrativas.
  - `Form_frmAltaSolicitud` y otros formularios consumen `m_ObjUsuarioActivo` para `usuarioCreacion`/`usuarioModificacion` y `rolUsuario` para habilitar UI.
  - Los formularios de búsqueda/alta (`frmBuscarSolicitudes`, `frmAltaSolicitud`) consumen `UsuarioServicio.getResponsablesTecnicos`/`getResponsablesCalidad` para los combos.
- **Puntos de entrada de código**:
  - `UsuarioServicio.getUsuarioConectadoConPermisos`, `getUsuarioConPermisos`, `DeterminarRol`, `getResponsablesTecnicos`, `getResponsablesCalidad`.
  - `MapeoServicio.getMapeoPC/CDCA/CDCASUB/PCSUB`, `SembrarBookmarksDesdePlantilla`.
  - `NoConformidadServicio.getNoConformidades`, `getNoConformidadesPorExpediente`, `getNoConformidadPorID`, `getNoConformidadPorCodigoCondor`, `estaRegistradaEnBaseDatosExternaa`, `actualizarCodigoConcesionEnTransaccion`, `limpiarCodigoConcesionEnTransaccion`.
  - `SuministradorServicio.getSuministradorPorID`, `getSuministradoresPorExpediente`.
  - `Entorno.cls` (`m_ObjEntorno`) con `URLDirectorioDocumentacion`, `DirUTE`, `NombreJefeCalidad`, `estados`, etc.
  - `Factoria.CreateEntity`.
  - `FUNCIONES UTILES.bas` con `getdb`, `getdbLanzadera`, `getdbExpedientes`, `getdbNoConformidades`, `getdbCorreo`.
- **Datos afectados**:
  - `TbConfiguracionBackends` (tabla local del frontend): lectura desde `getdb()`/rutinas de backend para resolver backend activo.
  - `TbUsuariosAplicaciones`/`TbUsuariosAplicacionesPermisos`: lectura desde `getdbLanzadera()`.
  - `TbExpedientes`/`TbExpedientesResponsables`/`TbExpedientesSuministradores`: lectura desde `getdbExpedientes()`.
  - `TbNoConformidades`: lectura/escritura desde `getdbNoConformidades()`.
  - `TbSuministradores`: lectura desde `getdb()`.
  - `tbMapeoCampos`: lectura desde `getdb()`.
- **Dependencias**:
  - `Wscript.Network` (COM) para `getUsuarioConectadoConPermisos`.
  - `DAO.Database` (Access) para todas las conexiones.
  - `Word.Application` COM en `SembrarBookmarksDesdePlantilla`.
- **Sincronización fuente↔binario**: si se modifican `Usuario*`, `Suministrador*`, `NoConformidad*`, `Mapeo*`, `Entorno`, `Factoria`, `FUNCIONES UTILES`, `Variables Globales`, basta `dysflow.import_code`. Si se modifica `frm0OtrosAdmin`, `import-form`.
- **Valoración de diseño (tal-como-está vs ideal)**: la identidad es simple y suficiente para Access monopuesto, pero no es apta para una migración web. La separación de BDs (`getdb*`) es pragmática, pero el acoplamiento con `TbConfiguracionBackends` debería ser explícito y testeable. La deuda principal está en (a) identidad local sin IdP, (b) `Factoria.CreateEntity` con lista cerrada, (c) `NoConformidadRepositorio` con `getdbNoConformidades()` sin seam, (d) `MapeoServicio.SembrarBookmarksDesdePlantilla` con Word COM. La pieza está bien hecha para el estado actual del producto; no se recomienda `Verified-runtime` sin seam y suite focal.

## §4 Receta de reconstrucción — REPRODUCIBILIDAD

1. Restaurar `Usuario.cls`, `UsuarioServicio.cls`, `UsuarioAplicacionPermisos.cls`, `UsuarioRepositorio.bas`.
2. Restaurar `Suministrador.cls`, `SuministradorServicio.cls`, `SuministradorRepositorio.bas`.
3. Restaurar `NoConformidad.cls`, `NoConformidadServicio.cls`, `NoConformidadRepositorio.bas`.
4. Restaurar `MapeoCampos.cls`, `MapeoServicio.cls`, `MapeoRepositorio.bas` (vinculados a CAP-002).
5. Restaurar `Entorno.cls`, `Factoria.bas`, `Variables Globales.bas`, `FUNCIONES UTILES.bas`, `modDevTools.bas`, `modActualizaciones.bas`.
6. Confirmar `TbConfiguracionBackends` en el frontend `CONDOR.accdb` con su esquema real: `Id`, `Habilitado`, `BackendActivo`, `BackendProduccion`, `BackendSandbox`, `BackendTest`, `PasswordBackend`, `IDAplicacion`, `NotificacionesTipo`, `RutaDirectorioAplicacion_PROD`, `RutaDirectorioAplicacion_LOCAL`.
7. Restaurar `Form_frm0OtrosAdmin.cls`.
8. Importar con `dysflow.import_modules` y compilar con `dysflow.compile_vba`. Verificar binario con `dysflow.verify_binary`.
9. Demostrar los escenarios de §2 con un manifest atómico `tests/tests.security.json` que cubra identidad, rol, selección de backend, mapa de campos y catálogo de NCs/suministradores. Mientras no exista, esta capacidad queda en `Verified-static`.

## §5 Evidencia y trazabilidad

- **Evidencia estática consultada**:
  - `src/classes/Usuario.cls`, `UsuarioServicio.cls` (150 líneas), `UsuarioAplicacionPermisos.cls`.
  - `src/classes/Suministrador.cls`, `SuministradorServicio.cls`.
  - `src/classes/NoConformidad.cls`, `NoConformidadServicio.cls`.
  - `src/classes/MapeoCampos.cls`, `MapeoServicio.cls` (98 líneas).
  - `src/classes/Entorno.cls`, `src/modules/Factoria.bas`, `Variables Globales.bas`, `FUNCIONES UTILES.bas`, `modDevTools.bas`, `modActualizaciones.bas`.
  - `src/forms/Form_frm0OtrosAdmin.cls`.
  - `docs/ERD/condor_datos.md` (notar que `TbConfiguracionBackends`, `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`, `TbSuministradores`, `TbNoConformidades`, `TbRACS`, `TbExpedientesRACS`, `TbCorreosEnviados` aparecen como tablas vinculadas no alcanzadas).
- **Evidencia Dysflow incorporada**:
  - `dysflow.get_schema` contra `condor_datos.accdb` para `TbConfiguracionBackends`: falla con tabla no encontrada.
  - `dysflow.get_schema` contra `CONDOR.accdb` para `TbConfiguracionBackends`: confirma columnas `Id`, `Habilitado`, `BackendActivo`, `BackendProduccion`, `BackendSandbox`, `BackendTest`, `PasswordBackend`, `IDAplicacion`, `NotificacionesTipo`, `RutaDirectorioAplicacion_PROD`, `RutaDirectorioAplicacion_LOCAL`.
  - `dysflow.list_tables` contra `CONDOR.accdb`: tablas locales `TbConfiguracionBackends`, `tbSandboxLog`.
  - `dysflow.list_linked_tables`: `TbCorreosEnviados`, `TbExpedientes`, `TbExpedientesRACS`, `TbExpedientesResponsables`, `TbExpedientesSuministradores`, `TbNoConformidades`, `TbRACS`, `TbSuministradores`, `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`.
  - `dysflow.verify_binary`: `actionableOk=true`; `Factoria`, `FUNCIONES UTILES`, `Mapeo*`, `NoConformidad*`, `Suministrador*`, `Usuario*`, `modActualizaciones`, `Form_frm0OtrosAdmin.cls` y `.form.txt` matched; `Entorno`, `modDevTools`, `UsuarioRepositorio`, `UsuarioServicio`, `Variables Globales` solo diferencias no accionables (`whitespaceOnly`/`caseOnly`).
- **Tests existentes**: manifest atómico `tests/testsSecurity.json` (commit `fc025e1`, 2026-06-15, Slice B6). Cubre `UsuarioServicio.DeterminarRol` (pure logic, no DB): jerarquía Administrador > Calidad > Tecnico, fallback a Tecnico cuando `Permisos=Nothing`, y raise de CondorError 513 cuando `Usuario=Nothing`. 4/4 átomos verdes.
- **Evidencia runtime Dysflow**:
  - `Test_Security_Strict_DeterminarRol_ReturnsAdministrador_WhenEsAdministradorSi`: **VERDE** 2.7 s, `EsAdministrador='Sí'` → retorna `1` (= `rol.Administrador`).
  - `Test_Security_Strict_DeterminarRol_ReturnsCalidad_WhenEsUsuarioCalidadSi`: **VERDE** 2.5 s, `EsAdministrador='No'` + `Permisos.EsUsuarioCalidad='Sí'` → retorna `2` (= `rol.Calidad`).
  - `Test_Security_Strict_DeterminarRol_ReturnsTecnico_ByDefault`: **VERDE** 2.7 s, `EsAdministrador='No'` + `Permisos=Nothing` → retorna `3` (= `rol.Tecnico`, fallback).
  - `Test_Security_Strict_DeterminarRol_RaisesError_WhenUsuarioIsNothing`: **VERDE** 2.6 s, sad path, captura error 513 con descripción "Se intentó determinar el rol de un objeto de usuario Nulo." y pila de llamadas.
- **SDD/intención consultada**: PRD `00_PRD_CONDOR` (característica de backend seleccionable), código actual.

### Diagnóstico de regresión

| Síntoma | Causa probable | Comprobación (Dysflow) | Ancla |
|---|---|---|---|
| `getUsuarioConectadoConPermisos` falla con `Wscript.Network` | COM no disponible o permisos del shell. | prueba con stub. | §2 BR-001 |
| `DeterminarRol` devuelve `Tecnico` en lugar de `Calidad` para un usuario calidad | `EsUsuarioCalidad` en BD externa no es `"Sí"`. | auditoría de BD externa. | §2 BR-002 |
| `getdb()` resuelve a backend incorrecto | `TbConfiguracionBackends.BackendActivo` cambiado o cache de `g_dbCondor` stale. | `dysflow.list_tables` + reinicio de la app. | §2 BR-008 |
| `MapeoServicio.getMapeoPC` devuelve `Nothing` | `tbMapeoCampos` no sembrada para `PC`. | sembrar con `SembrarBookmarksDesdePlantilla`. | §2 BR-004 |
| `NoConformidadServicio.getNoConformidades` falla | BD externa no accesible. | verificar `getdbNoConformidades()`. | §2 BR-005 |
| `Factoria.CreateEntity` devuelve `Nothing` | Nombre de entidad nuevo no añadido. | añadir a `Factoria.bas`. | §2 BR-007 |

### Trazabilidad de release

| Elemento | Ref. tracker | Versión de staging (UAT) | Estado UAT | Release de producción | Fecha en producción | Nota |
|---|---|---|---|---|---|---|
| Seguridad, administración e infraestructura | Pendiente | Pendiente de confirmación | pending | Pendiente | Pendiente | Fuente↔binario sincronizado y esquema de `TbConfiguracionBackends` auditado; pendiente manifest atómico. |

## §6 Notas de migración web

- **Conservar**: jerarquía de roles efectiva, mapa de campos por tipo de solicitud, separación de catálogos de NCs/suministradores, configuración de backend.
- **Transformar**: `Wscript.Network` COM a un IdP (Kerberos/SAML/OAuth); `getdb*()` singletons a contextos inyectables; `Factoria.CreateEntity` a un registro declarativo; `MapeoServicio.SembrarBookmarksDesdePlantilla` a una herramienta offline de administración de plantillas.
- **NO copiar**: identidad local sin IdP, `TbConfiguracionBackends` mutable en caliente sin seam, `NoConformidadRepositorio` con `getdbNoConformidades()` directo, dependencia de Word COM en el sembrado de bookmarks.
- **Preguntas abiertas**: ¿La identidad debe migrar a un IdP corporativo? (responsable de producto). ¿`TbConfiguracionBackends` se conserva como modelo de configuración runtime o se sustituye por un servicio de configuración central? (equipo técnico). ¿`MapeoCampos` debe versionarse en git o mantenerse como configuración externa? (responsable de producto + equipo técnico).

## §7 Registro de confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| La identidad se basa en `Wscript.Network.UserName` con mapeo de alias. | Verified-static | `UsuarioServicio.cls` líneas 53-77. | 2026-06-15 |
| La jerarquía de roles efectiva es `Administrador` ⊇ `Calidad` ⊇ `Tecnico`. | Verified-static | `UsuarioServicio.cls` líneas 79-109. | 2026-06-15 |
| Los responsables técnicos se leen de `getdbExpedientes()`; los de calidad de `getdbLanzadera()`. | Verified-static | `UsuarioServicio.cls` líneas 113-149. | 2026-06-15 |
| `MapeoServicio.getMapeo*` lee `tbMapeoCampos` por nombre de plantilla. | Verified-static | `MapeoServicio.cls` líneas 17-59. | 2026-06-15 |
| `NoConformidadRepositorio` opera contra `getdbNoConformidades()`. | Verified-static | `NoConformidadRepositorio.bas` (múltiples). | 2026-06-15 |
| `Factoria.CreateEntity` lista 19+ nombres de clase y devuelve `Nothing` para el resto. | Verified-static | `Factoria.bas`. | 2026-06-15 |
| `getdb*()` son singletons de proceso; cualquier cambio en `TbConfiguracionBackends.BackendActivo` se aplica en el siguiente acceso. | Verified-static | `FUNCIONES UTILES.bas` líneas 64-199. | 2026-06-15 |
| `TbConfiguracionBackends` vive en `CONDOR.accdb`, no en `condor_datos.accdb`. | Verified-static | `dysflow.get_schema`: backend falla; frontend confirma esquema. | 2026-06-15 |
| `TbConfiguracionBackends` tiene columnas `Id`, `Habilitado`, `BackendActivo`, `BackendProduccion`, `BackendSandbox`, `BackendTest`, `PasswordBackend`, `IDAplicacion`, `NotificacionesTipo`, `RutaDirectorioAplicacion_PROD`, `RutaDirectorioAplicacion_LOCAL`. | Verified-static | `dysflow.get_schema` contra `CONDOR.accdb`. | 2026-06-15 |
| Fuente↔binario de seguridad/administración no tiene diferencias accionables. | Verified-static | `dysflow.verify_binary` `actionableOk=true`; solo `whitespaceOnly`/`caseOnly` no accionables. | 2026-06-15 |
| `MapeoServicio.SembrarBookmarksDesdePlantilla` interactúa con Word COM. | Verified-static / deuda de testabilidad | `MapeoServicio.cls` líneas 61-94. | 2026-06-15 |
| Existe un manifest atómico de pruebas de seguridad que cumpla `access-vba-tdd` v2.4.2. | Verified-runtime (parcial) / pendiente expansión a `getUsuario*` y `getResponsables*` | `tests/testsSecurity.json` con cuatro átomos verdes (commit `fc025e1`): `DeterminarRol` para los 3 roles + sad path Nothing. Pendientes: `getUsuarioConPermisos` (requiere `TbUsuariosAplicacionesPermisos`), `getResponsablesTecnicos`/`Calidad` (requiere `TbExpedientesResponsables`). | | 2026-06-15 |

**Divergencias pendientes de revisión humana**:

- BR-008: el esquema de `TbConfiguracionBackends` ya está auditado; falta una prueba focal que demuestre la conmutación de backend sin mutar configuración productiva.
- BR-001: la identidad local no es apta para web. Planificar la integración con un IdP corporativo.
- BR-007: `Factoria.CreateEntity` con lista cerrada es un acoplamiento; considerar un registro declarativo.
- `MapeoServicio.SembrarBookmarksDesdePlantilla`: separar en una herramienta offline con seam Word explícito.
