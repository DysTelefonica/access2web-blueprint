# HPS — formularios, navegación y call paths

## Navegación principal

```text
Form_frmSplash
  -> EVE
  -> Form_FormInicial (opciones generales)
      ├─> Form_FormInicial00Principal (entrada)
      ├─> Form_FormInicial01Lateral (menú lateral)
      ├─> Form_FormInicial02ConsultasPrincipal (consultas)
      │    └─> Form_FormInicial03Consultas<PorCampo>:
      │         - Contratistas
      │         - Datos
      │         - EmpresaUsuario
      │         - Estado
      │         - Grado
      │         - JuridicaTramitacion
      │         - MotivoHPS
      │         - Nombre
      │         - ProyectoAsignado
      └─> Form_FormInicial05Configuracion (configuración)
  -> Form_FormExpedientesBusqueda (búsqueda de expedientes vinculados)
  -> Form_FormUsuario02HPS (detalle de un usuario HPS)
```

## Call paths críticos

| Capacidad | Camino observado | Persistencia / efecto |
|---|---|---|
| Inicio | `frmSplash.Form_Timer → EVE → Constructor.getUsuario → getdb()` | TempVars, sesión, caché de usuarios |
| Búsqueda de usuario | `Form_FormUsuario02HPS → HPS.Usuario → getUsuarioHPS` (lazy load) | consulta `TbUsuarios` por `idUsuario` |
| Cálculo de estado | `HPS.APuntoDeCaducar / Caducado / Solicitado / PendienteRenovacion` (propiedades calculadas vía `DatosHPSCalculados`) | derivado de `TbHPS` + `F_Curso` + `Requiere_Curso` |
| Renovación | `HPS.UsuarioHPS.Registrar / Modificar` (con `UsuarioLifecycleTransactionCoordinator`) | transacción DAO sobre `TbUsuarios` + `TbHPS` + `TbObservaciones` |
| Histórico | `UsuarioHistorico.cls` + `ObservacionHistorica.cls` | snapshot en `TbUsuariosHistoricos` + `TbObservacionesHistoricas` |
| Anexos | `AnexoUsuarioHPS.AnexoUsuario → AnexoSelectionTransactionCoordinator` (transaccional) | `TbAnexosUsuariosHPS` + binding a fichero |
| Indicadores | `Indicador.cls + clsIndicadoresBus.cls` + `Mod_StartupCacheInitialization.bas` (cache coherente) | tiempo real verificado por `Test_RealTimeIndicatorCoherence.bas` |
| Consultas | `Form_FormInicial02ConsultasPrincipal → 8 Form_FormInicial03Consultas<PorCampo>` | lee `TbUsuarios` con filtros; usa índices |
| Vinculación Expedientes | `Form_FormExpedientesBusqueda → Expediente (clase compartida Lanzadera)` | conceptual, sin FK física |
| Cache | `Mod_Cache_Core.bas` + `cacheUsuario.bas` + `cacheSuministrador.bas` | `TbConfiguracion.CacheHabilitada` |
| Auditoría | `AnexoSelectionTransactionCoordinator` + `UsuarioLifecycleTransactionCoordinator` | garantiza atomicidad |

## Inventario normalizado

- **Formularios**: ~30 archivos `Form_*.cls` cada uno con su `.form.txt` compañero. Cada par = código + layout exportado.
- **Patrón `.cls + .form.txt`**: idéntico al de NoConformidades.
- **Clases**: 29 en `src/classes/`:
  - **Dominio**: `HPS.cls`, `Usuario.cls`, `UsuarioHPS.cls`, `UsuarioHistorico.cls`, `UsuarioSICA.cls`, `UsuarioServicio.cls`, `Expediente.cls`, `Suministrador.cls`, `Consulta.cls`, `Correo.cls`, `Entorno.cls`, `Aspecto.cls`, `DatosLocal.cls`, `Indicador.cls`, `Observacion.cls`, `ObservacionHistorica.cls`.
  - **Tests Doubles**: `clsTestDoubleForm.cls`, `clsTestDoubleHpsEditListener.cls`, `clsTestDoubleHpsEditPublisher.cls`, `clsTestDoubleIndicador.cls`, `clsTestDoubleIndicadorConsumer.cls` (5 archivos, indica TDD maduro).
  - **Coordinadores transaccionales**: `AnexoSelectionTransactionCoordinator.cls`, `UsuarioLifecycleTransactionCoordinator.cls`.
  - **Anexos**: `AnexoUsuarioHPS.cls`, `AnexoUsuarioHistorico.cls`, `AnexoUsuarioSICA.cls`.
  - **Compartidas con Lanzadera**: `UsuarioAplicacionPermisos.cls`, `Expediente.cls` (FK conceptual).
  - **Indicadores**: `Indicador.cls`, `clsIndicadoresBus.cls`.
- **Módulos**: 30 en `src/modules/`:
  - **Bootstrap/factory/DAO**: `VariablesEntorno.bas` (con `EVE` + `getdb`), `Constructor.bas` (con `getUsuario`), `Funciones Generales.bas`, `Instalador.bas`.
  - **Caché (3+ módulos)**: `cacheUsuario.bas`, `cacheSuministrador.bas`, `Mod_Cache_Core.bas`, `Mod_StartupCacheInitialization.bas`, `CacheConsistencyAudit.bas`, `Mod_Sincronizacion_Historico.bas`.
  - **Indicadores**: `modIndicadores.bas`.
  - **Configuración**: `modConfiguracionHPS.bas`.
  - **Filtros**: `Filtro.bas`.
  - **JSON**: `JsonConverter.bas`.
  - **Datos dummy/Fixtures**: `DummyBig.bas`, `DummyBig2.bas`, `DummyHatw.bas`, `DummyHatw2.bas`, `DummySmall.bas`, `DummySpecial.bas`, `Módulo1.bas` — fixtures de pruebas para `clsTestDouble*`.
  - **Tests VBA**: 9 archivos `Test_*.bas` (ver capabilities.md).
- **Queries**: `TbConsultas` (queries SQL externas pre-armadas).
- **Reports/macros/queries**: no se inspeccionaron queries exportadas en esta pasada; macros embebidas requieren revisión del binario (no hecha en este lote).

## Nota de evidencia

CodeGraph-VBA se consultó primero sobre `00_main` y devolvió call paths dinámicos con 31 símbolos en 1 archivo para la query inicial. Dysflow read-only se ejecutó después de `register_worktree` + `migrate_project_config` (T18 caps-block) + `accessPath` absoluto explícito: 22 tablas, 6 FKs, 27 columnas en `TbUsuarios`. La inspección de UI se mantiene read-only y no se han alterado formularios.