[← Back to DOCS](../../../DOCS.md)

# 03 · Condor

## Propósito

Evidencia de descubrimiento de **Condor**, aplicación Access/VBA que gestiona **solicitudes de aprobación de calidad** vinculadas a Expedientes. Cada solicitud tiene un tipo (`PC`, `CD_CA`, `CD_CA_SUB`, `PC_SUB`), pasa por un workflow con estados, validaciones de calidad, posibles rechazos y adjuntos. Condor es la **app más activa del ecosistema**: está en plena evolución hacia una arquitectura hexagonal con ViewModels, repositorios y un Edge WebView embebido para mostrar vistas web.

## Estado

- **Fase:** descubrimiento completo, Batch 6 + inventario Dysflow real sobre staging.
- **Fecha de evidencia:** 2026-08-05.
- **Repositorio:** `C:\00repos\codigo\00_CONDOR\staging` (seleccionado por el user para este lote). CodeGraph-VBA inspeccionado; inventario Dysflow completo.
- **Main de comparación:** `C:\00repos\codigo\00_CONDOR\00_main` (release publicado); no inspeccionado en esta pasada.
- **Frontend:** `CONDOR.accdb` en staging (44 MB; el frontend más grande de las 8 aplicaciones). Backup `CONDOR.accdb.bak-20260626113218` (49 MB) presente en staging.
- **Backend autoritativo:** `C:\00repos\datos\condor_datos.accdb` (5 MB). **Duplicado en staging local** (`staging/condor_datos.accdb`).
- **Dysflow:** `staging/.dysflow/project.json` válido (`projectId: condor`). Inventario backend completo: **15 tablas**, **5 FKs**, **volumen real del backend autoritativo** (`C:\00repos\datos\condor_datos.accdb`): 1 solicitud, 9 estados, 0 rechazos, 0 adjuntos, 5 log cambios, 3 log errores, **183 mapeo campos**, 0 log estados, 0 validaciones. **Staging lee del backend autoritativo** vía `TbConfiguracionBackends` (mismo volumen).
- **APAP y APAP_WEB** no aparecen (proyecto personal del desarrollador).

## Lote asociado

Lote 5 del plan de discovery (siguiente a NoConformidades). Posicionado por ser la app con la evolución arquitectónica más avanzada del ecosistema.

## Entregables

1. [Capacidades](capabilities.md)
2. [Formularios y call paths](forms.md)
3. [Modelo físico y diccionario](data-model.md)
4. [Matriz de migración](migration-matrix.md)
5. [Integraciones y automatización](integrations-automation.md)
6. [Seguridad y reglas](security-rules.md)

## Fuentes de autoridad

1. `C:\00repos\codigo\00_CONDOR\staging\src` (clases, forms, módulos) — codegraph-vba.
2. `C:\00repos\codigo\00_CONDOR\staging\tests\` (tests VBA) — codegraph-vba.
3. `C:\00repos\documentacion\OPENSPEC\00_CONDOR` (documentación previa).
4. Dysflow read-only sobre `CONDOR.accdb` (frontend) y `condor_datos.accdb` (backend).
5. Engram como contexto histórico.

## Reglas de evidencia

- No se han realizado imports, exports, sync, tests, compile, cleanup ni escrituras.
- Se excluyen valores personales, correos, credenciales, hashes, hosts y nombres de máquina.
- Las rutas UNC y hosts no se reproducen (regla de evidencia transversal).
- APAP y APAP_WEB no se mencionan.
- ⚠️ **Hallazgo de seguridad crítico**: el módulo `FUNCIONES UTILES.bas` tiene un fallback hardcoded `GetPasswordDB = "dpddpd"` cuando el INI no tiene la contraseña (línea 150). **Esta contraseña está expuesta en código fuente**. Ver [Seguridad § D93](security-rules.md#d93--password-hardcodeado-como-fallback-en-getpassworddb).

## Hallazgos críticos del lote

1. **⚠️ Password hardcoded "dpddpd"** como fallback en `FUNCIONES UTILES.bas:150` (`GetPasswordDB`). Si el INI no tiene la contraseña, **cae al literal `"dpddpd"`**. Riesgo de seguridad grave. Disposición D93 propuesta.

2. **Solo 1 caller de `getdb()`** y **2 callers de `getUsuario()`**. Condor centraliza las operaciones de BD en un único punto (`FUNCIONES UTILES.bas:84` y `UsuarioRepositorio.bas:9`). Patrón diferente a las otras apps (más DAO-direct). Es una **re-factorización reciente**.

3. **`m_TestingMode` con sandbox seguro** (`FUNCIONES UTILES.bas:84-120`): cuando `m_TestingMode=True`, getdb() enruta a `m_BackendSandboxURL` con `m_BackendSandboxPassword`. Hay validación de cache safety (Spec-008) y comentario explícito "TESTS BLOCKED" si el sandbox no está configurado. **Patrón maduro** — referencia para la nueva plataforma.

4. **52+ clases de dominio** vs 27-47 de las otras apps. Condor tiene la **mayor superficie de dominio**. Estructura:
   - **ViewModels**: `SolicitudViewModel`, `SolicitudBusquedaViewModel`, `AdjuntoViewModel`, `DatosCDCA/CDCASUB/PC/PCSUBViewModel`, `ExpedienteViewModel`, `FiltrosSolicitud`. **7 ViewModels** — patrón hexagonal en VBA.
   - **Servicios**: `SolicitudServicio`, `AdjuntosServicio`, `DatosCDCAServicio`, `DatosCDCASUBServicio`, `DatosPCServicio`, `DatosPCSUBServicio`, `DocumentoServicio`, `EstadoServicio`, `ExpedienteServicio`, `LogCambioServicio`, `NoConformidadServicio`, `NotificacionServicio`, `RechazoServicio`, `RevisionServicio`, `SuministradorServicio`, `UsuarioServicio`, `ValidacionRevisionServicio`, `WebVisorCacheServicio`, `WorkflowServicio`, `MapeoServicio`, `SnapshotServicio`. **21 Servicios** — lógica de negocio separada de UI y DAO.
   - **Repositorios** (en `modules/`): `AdjuntoRepositorio`, `AplicacionRepositorio`, `CorreoRepositorio`, `DatosCDCARepositorio`, `DatosCDCASUBRepositorio`, `DatosPCRepositorio`, `DatosPCSUBRepositorio`, `EstadoRepositorio`, `ExpedienteRepositorio`, `LogCambioRepositorio`, `LogErrorRepositorio`, `LogEstadoRepositorio`, `MapeoRepositorio`. **13 Repositorios**.
   - **Domain entities**: `Solicitud`, `Estado`, `Adjunto`, `Rechazo`, `Expediente`, `Suministrador`, `NoConformidad`, `Usuario`, `Entorno`, `LogCambio`, `LogError`, `LogEstado`, `MapeoCampos`, `ValidacionRevision`, `DatosCDCA`, `DatosCDCASUB`, `DatosPC`, `DatosPCSUB`.
   - **Sandbox**: `SandboxConfig`, `SandboxGestor`, `SandboxValidator`.
   - **Mocks**: `MockNotifServ`.
   - **Errores**: `CondorError`.

5. **Edge WebView embebido**: `Form_frmGestionSolicitud.cls` usa `Me.webInfo.Navigate rutaNavegacion` con `WebVisorCacheServicio` + `SnapshotServicio`. **Condor ya tiene vistas web embebidas en Access** — un patrón híbrido Access+web. La nueva plataforma web absorbe esto.

6. **4 tipos de Solicitud** con datos y servicios propios: `PC`, `CD_CA`, `CD_CA_SUB`, `PC_SUB`. Cada uno tiene:
   - Clase de datos: `DatosPC`, `DatosCDCA`, `DatosCDCASUB`, `DatosPCSUB`.
   - Servicio: `DatosPCServicio`, `DatosCDCAServicio`, `DatosCDCASUBServicio`, `DatosPCSUBServicio`.
   - Repositorio: `DatosPCRepositorio`, `DatosCDCARepositorio`, `DatosCDCASUBRepositorio`, `DatosPCSUBRepositorio`.
   - ViewModel: `DatosPCViewModel`, `DatosCDCAViewModel`, `DatosCDCASUBViewModel`, `DatosPCSUBViewModel`.

7. **Vinculación con NoConformidades**: `Form_frmGestionSolicitud.cls:1908` verifica `ncServ.getNoConformidadPorCodigoCondor(m_ViewModel.Solicitud.codigoSolicitud)`. Hay FK conceptual `idNCAsociada` en `tbSolicitudes`. Cruce directo con el Lote 5 (NoConformidades).

8. **WorkflowServicio + SnapshotServicio**: motor de workflow + sistema de snapshots. Indica que Condor tiene un flujo de aprobación complejo con historial.

9. **15 tablas en staging** vs ~42 de NoConformidades. La diferencia es que Condor tiene **menos superficie de datos** y **más lógica de aplicación** (ViewModels, Servicios, Repositorios). Es una app "delgada" en datos y "gruesa" en comportamiento.

10. **9 estados en `tbEstados`** (volumen bajo, staging de prueba). Catálogo que define el ciclo de vida de una Solicitud.

11. **`tbLogErrores` con 3 filas**: hay logging de errores estructurado. Coherente con D27 (logs estructurados canónicos).

12. **`tbTransiciones` y `tbValidacionRevision`**: tablas de soporte que documentan las transiciones de estado y la validación de calidad. Patrón workflow maduro.

## Checklist

- [x] Inventario funcional, formularios, clases y módulos documentados vía codegraph-vba.
- [x] Inventario real Dysflow del backend (15 tablas, 5 FKs, staging con 1 solicitud).
- [x] **52+ clases** con desglose por capas (ViewModels, Servicios, Repositorios, Domain).
- [x] D93 propuesto para password hardcoded (riesgo de seguridad crítico).
- [x] APAP y APAP_WEB no aparecen en esta evidencia.
- [ ] Épica + tickets + matriz de migración de datos para Condor.

## Siguiente paso

Cruzar el inventario con la documentación previa en `OPENSPEC/00_CONDOR`. Cerrar D93 (password hardcoded) y D94 (FKs conceptuales sin constraint). Generar la **épica + tickets accionables + matriz de migración de datos** para Condor (alcance expandido). Continuar después con Brass (Lote 6) e HPS_Solicitudes (Lote 7).

## Core invariants

- **⚠️ D93 password hardcoded `"dpddpd"` como fallback**: `FUNCIONES UTILES.bas:150` define `GetPasswordDB = "dpddpd"` cuando el INI no tiene la contraseña. **NO es la contraseña directa**, pero igualmente expone el backend si la configuración se queda vacía. Riesgo CRÍTICO; remediación operativa separada (saneamiento del código + rotación de la contraseña).
- **`m_TestingMode` con sandbox seguro**: `FUNCIONES UTILES.bas:84-120` enruta a `m_BackendSandboxURL` con `m_BackendSandboxPassword` cuando `m_TestingMode=True`. Hay validación de cache safety (Spec-008) y comentario explícito "TESTS BLOCKED" si el sandbox no está configurado. **Patrón de referencia** para los tests de la nueva plataforma.
- **52+ clases (mayor superficie de dominio del ecosistema)**: ViewModels (7) + Servicios (21) + Repositorios (13) + Domain entities (18) + Sandbox (3) + Mocks + Errores. La nueva plataforma absorbe esta forma hexagonal con `app/src/modules/<app>/{domain,ports,application,adapters,di,delivery}` (D8, DA-1).
- **Edge WebView embebido**: `Form_frmGestionSolicitud.cls` usa `Me.webInfo.Navigate rutaNavegacion` con `WebVisorCacheServicio` + `SnapshotServicio`. La nueva plataforma web absorbe esto: las rutas HTMX reemplazan el WebView, pero el patrón de cache de vistas se preserva como puerto.
- **Vinculación con NoConformidades (`idNCAsociada` en `tbSolicitudes`)**: FK conceptual (D94). La nueva plataforma formaliza con constraint o mantiene como referencia documentada, según la decisión del SDD.
- **`TbConfiguracionBackends` vive en frontend**: el módulo de configuración con `BackendActivo`, `BackendProduccion`, `BackendSandbox`, `BackendTest`, `IDAplicacion`, `PasswordBackend` está en el frontend legacy. La nueva plataforma lo migra al backend con el adapter `ConfigPort`.

## Contributor checklist

- [ ] El cambio respeta las 6 reglas de §Core invariants; el `ci / quality` check pasa verde.
- [ ] Si el PR toca `FUNCIONES UTILES.bas`, NO se reintroduce ninguna contraseña hardcodeada (D93, D104); el acceso a backend se hace vía `SecretManagerPort` (D9-D10) o variable de entorno.
- [ ] Si el cambio añade una migración Alembic, sigue `expand_and_contract` (D82): añadir columnas o tablas, sin `DROP` ni `ALTER` destructivos en la misma release.
- [ ] Si el cambio toca `m_TestingMode` o el sandbox, se conserva como patrón de referencia y se documenta en `tests/` con un test de smoke que pinea el comportamiento.
- [ ] Si el cambio introduce una FK hacia NoConformidades, la decisión D94 sigue aplicando: la FK se formaliza con constraint o se documenta como referencia conceptual.
- [ ] El PR es ≤ 400 líneas (`additions + deletions`); si no, partir por unidad de trabajo o encadenar.

## Navigation

Previous: [hps-solicitudes](../hps-solicitudes/README.md) | Next: [expedientes](../expedientes/README.md)