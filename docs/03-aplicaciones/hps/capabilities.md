# HPS — capacidades observadas

## Resultado

La aplicación cubre un agregado de **gestión de personal HPS**: alta/baja/renovación de usuarios HPS, seguimiento de cursos obligatorios, observaciones, histórico y anexos. Es **read-heavy**: la mayoría de las operaciones son consultas sobre el catálogo de usuarios, con escrituras limitadas a actualizaciones de HPS, observaciones y carga inicial. La paridad futura debe incluir como mínimo las capacidades siguientes; ninguna se marca como retirada.

| Dominio | Capacidades evidenciadas | Evidencia principal |
|---|---|---|
| Arranque e identidad | `EVE` con `getUsuario` (`Constructor.bas:1415`), `getdb` (`VariablesEntorno.bas:587`), IDAplicacion `17` (producción) / `51` (pruebas) | `src/modules/VariablesEntorno.bas`; `src/modules/Constructor.bas` |
| Roles | tres niveles (Administrador / Calidad / Técnico), análogo al resto del ecosistema | `src/forms/Form_FormUsuario02HPS.cls`, `Form_FormInicial*.cls` |
| Configuración de backend | `TbConfiguracionBackends` con `BackendActivo` (PROD/LOCAL/SANDBOX), `IDAplicacion`, `EnPruebas`/`EnDesarrollo`, `PasswordBackend` (⚠️ requiere D92) | `src/modules/VariablesEntorno.bas` |
| Usuarios HPS | alta, baja, renovación, modificación de datos personales, cálculo de `APuntoDeCaducar`, `Caducado`, `Solicitado`, `PendienteRenovacion` | `src/classes/HPS.cls`, `UsuarioHPS.cls` |
| Histórico de usuarios | `TbUsuariosHistoricos` con 242 filas (auditoría de cambios) | `UsuarioHistorico.cls`, `AnexoUsuarioHistorico.cls` |
| Integración SICA | `TbUsuariosSICA` + `TbAnexosUsuariosSICA` (vinculación con sistema SICA externo) | `UsuarioSICA.cls`, `AnexoUsuarioSICA.cls` |
| HPS propiamente | `TbHPS` con 1280 registros (relación usuario-curso) + `TbHPSEquivalencia` + `TbHPSGrado` | `HPS.cls`, `clsIndicadoresBus.cls` |
| Motivo HPS | `TbMotivoHPS` con motivos de alta/baja | `HPS.cls`, `modConfiguracionHPS.bas` |
| Cursos | `TbAuxCursos` con cursos obligatorios | `DatosLocal.cls`, `modIndicadores.bas` |
| Observaciones | `TbObservaciones` (330 filas) + `TbObservacionesHistoricas` (histórico de cambios en observaciones) | `Observacion.cls`, `ObservacionHistorica.cls` |
| Anexos por usuario | `TbAnexosUsuariosHPS`, `TbAnexosUsuariosHistoricos`, `TbAnexosUsuariosSICA`, `TbUsuarioAnexos` (4 tablas para 4 fuentes de anexos) | `AnexoUsuarioHPS.cls`, `AnexoUsuarioHistorico.cls`, `AnexoUsuarioSICA.cls`, `Aspecto.cls` |
| Indicadores en tiempo real | servicio de indicadores con coherencia verificada | `Indicador.cls`, `clsIndicadoresBus.cls`, `modIndicadores.bas`, `Test_RealTimeIndicatorCoherence.bas` |
| Transacciones de selección | `AnexoSelectionTransactionCoordinator`, `UsuarioLifecycleTransactionCoordinator` (transaccionalidad de operaciones multi-tabla) | `*.cls` correspondientes, `Test_AnexoSelectionTransaction.bas` |
| Consultas pre-armadas | `Form_FormInicial02ConsultasPrincipal` + 13 formularios de consulta por criterio (consultas por contratistas, datos, empresa, estado, grado, jurídica, motivo HPS, nombre, proyecto asignado) | `src/forms/Form_FormInicial0[1-5]*.cls` + `.form.txt` |
| Consultas SQL externas | `TbConsultas` (almacena queries) + `Consulta.cls` (ejecutor) | `Consulta.cls` |
| Histórico con anexos | `Mod_Sincronizacion_Historico.bas` (sincronización histórico ↔ anexos) | `Mod_Sincronizacion_Historico.bas`, `Test_HistoricoAdjuntosTransactionWrapper.bas` |
| Catálogos | `TbJuridicasContratacion` con jurídicas de contratación | `Form_FormInicial03ConsultasJuridicaTramitacion.cls` |
| Integraciones | Lanzadera (identidad), SICA (usuarios SICA), HPS_Solicitudes (vinculación por `IDSolicitud`), AGEDYS (presumido), correos | `Constructor.bas` (vía `getdbLanzadera` o tablas compartidas) |
| Acoplamiento DAO directo | `getdb()` con **107 callers** (el más bajo de las 8 aplicaciones; HPS es read-heavy) | `src/modules/VariablesEntorno.bas:587` (`getdb`) |
| Capa de caché propia | `cacheUsuario.bas`, `cacheSuministrador.bas`, `Mod_Cache_Core.bas`, `Mod_StartupCacheInitialization.bas`, `CacheConsistencyAudit.bas`, `Mod_Sincronizacion_Historico.bas` (3+ módulos de caché) | confirma D91 cross-app |
| Tests VBA | 9 archivos `Test_*.bas`: Test_AnexoSelectionTransaction, Test_CacheConsistencyAudit, Test_HistoricoAdjuntosTransactionWrapper, Test_HPSConfig, Test_HPSEntorno, Test_LocalReadAccessAuthorization, Test_PerAnexoMove, Test_RealTimeIndicatorCoherence, Test_StartupCacheInitialization | confirma D87 (tests VBA como evidencia de comportamiento) |

## Reglas de conservación

- La identidad se carga vía `Constructor.getUsuario` (análogo a Lanzadera/Expedientes/Gestion_Riesgos/NoConformidades). El acceso directo a la BD de Lanzadera debe **romperse** (D86) en favor del adaptador unificado de autenticación.
- `IDAplicacion = "17"` (producción) / `"51"` (pruebas) se asigna por la TempVar `EnPruebas`; en la nueva plataforma viene de la configuración (D9).
- El **histórico de usuarios** (`TbUsuariosHistoricos` + `TbObservacionesHistoricas` + `TbAnexosUsuariosHistoricos`) preserva trazabilidad de cambios; **la retención debe ser al menos la misma que la de auditoría** (D29).
- Los **booleanos como Text(2)** (`CursoEnVigor`, `Requiere_Curso`, `RequiereComunicacionConcesion`) deben **estandarizarse** a `BOOLEAN` en PostgreSQL con regla de migración explícita (ver matriz de migración § D92).
- Los **indicadores en tiempo real** (`Indicador.cls`, `Test_RealTimeIndicatorCoherence.bas`) tienen **kill switch atómico** que se preserva como referencia del puerto de observabilidad de la nueva plataforma.
- Las **transacciones de selección** (`AnexoSelectionTransactionCoordinator`) y **transacciones de lifecycle** (`UsuarioLifecycleTransactionCoordinator`) son **transaccionales**; se traducen a SQLAlchemy `AsyncSession.begin()` (D66, D82).

## Evidencia previa

Se han cosechado PRD, Discovery Map, Architecture Overview, ERD, OpenSpec CAP-001..055, UAT y releases antes de inspeccionar staging. CodeGraph-VBA sobre `00_main` se consultó primero; el inventario Dysflow real se ejecutó después (22 tablas, 6 FKs, 27 columnas en `TbUsuarios`, 493 filas principales). Las afirmaciones divergentes entre esos documentos y el código quedan abiertas, no resueltas por intención.