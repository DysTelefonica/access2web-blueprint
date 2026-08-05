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
- **Dysflow:** `staging/.dysflow/project.json` válido (`projectId: condor`). Inventario backend completo: **15 tablas**, **5 FKs**, staging con volumen bajo (1 solicitud, 9 estados, 3 errores, 0 rechazos, 0 adjuntos).
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