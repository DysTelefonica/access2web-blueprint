# PRD — HPS

> **Estado:** fases 0 (fijación) y 1 (inventario) cerradas; fase 2 (formularios) con 25 de 38.

| Campo | Valor |
|---|---|
| Repositorio | DysTelefonica/HPS |
| Commit fijado | `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` (`main`, 2026-10-08, «chore(github): add bug report issue form (#28)») |
| Índice CodeGraph-VBA | 137 ficheros indexados, 13 638 nodos, versión 2.1.1, sobre clon de solo lectura de ese commit |
| Estudio previo | [epic.md](../epic.md), [forms.md](../forms.md), [capabilities.md](../capabilities.md), [data-model.md](../data-model.md), [security-rules.md](../security-rules.md), [integrations-automation.md](../integrations-automation.md), [migration-matrix.md](../migration-matrix.md) (DRAFT), walkthroughs `walkthrough-*.json`, [ERD/schema.sql](../ERD/schema.sql) |
| Snapshot comparado | `data/staging/hps/src` (2026-08-06) |
| Issue | #1076 (método y plantilla de #1069) |

## Alcance actual

Fija el commit y concilia el inventario. Los recuentos de líneas, controles y procedimientos son mecánicos (ver [Método del recuento](#método-del-recuento)) y no sustituyen al análisis.

## Inventario reconciliado

### Recuentos por tipo

| Tipo | `5e64cf9` | Snapshot | Diferencia |
|---|---|---|---|
| Formularios (`src/forms/Form_*.cls`, con su `.form.txt`) | 38 | 38 | +0 |
| Módulos estándar (`src/modules/*.bas`) | 31 | 31 | +0 |
| Clases (`src/classes/*.cls`) | 28 | 28 | +0 |

La comparación normaliza las líneas `Attribute VB_`, los comentarios, las líneas en blanco y las diferencias de espacios y mayúsculas. Diferencia textual cero no demuestra que el estudio previo sea correcto: solo que el código no ha cambiado desde el snapshot.

### Delta de código frente al snapshot

- **Distintos del snapshot:** 0 formularios, 0 módulos y 0 clases.
- **Nuevos en el commit (sin equivalente en el snapshot):** 0 formularios, 0 módulos y 0 clases.
- **Del snapshot ausentes en el commit:** 0 formularios, 0 módulos y 0 clases.

### Formularios

Líneas: `.cls` del formulario. Controles: bloques `Begin` del `.form.txt`. Procedimientos: `Sub`, `Function` y `Property` del `.cls`. «Frente al snapshot»: **igual** si el `.cls` coincide tras la normalización.

| Objeto | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|
| Form_FormExpedientesBusqueda | src/forms/Form_FormExpedientesBusqueda.cls | 279 | 24 | 8 | igual | pendiente | — |
| Form_FormInicial | src/forms/Form_FormInicial.cls | 338 | 26 | 9 | igual | pendiente | — |
| Form_FormInicial00Principal | src/forms/Form_FormInicial00Principal.cls | 150 | 22 | 5 | igual | analizada | [Form_FormInicial00Principal.md](formularios/Form_FormInicial00Principal.md) |
| Form_FormInicial01Lateral | src/forms/Form_FormInicial01Lateral.cls | 495 | 27 | 18 | igual | pendiente | — |
| Form_FormInicial02ConsultasPrincipal | src/forms/Form_FormInicial02ConsultasPrincipal.cls | 318 | 20 | 9 | igual | pendiente | — |
| Form_FormInicial03ConsultasContratistas | src/forms/Form_FormInicial03ConsultasContratistas.cls | 97 | 15 | 6 | igual | analizada | [Form_FormInicial03ConsultasContratistas.md](formularios/Form_FormInicial03ConsultasContratistas.md) |
| Form_FormInicial03ConsultasDatos | src/forms/Form_FormInicial03ConsultasDatos.cls | 35 | 317 | 2 | igual | analizada | [Form_FormInicial03ConsultasDatos.md](formularios/Form_FormInicial03ConsultasDatos.md) |
| Form_FormInicial03ConsultasEmpresaUsuario | src/forms/Form_FormInicial03ConsultasEmpresaUsuario.cls | 189 | 16 | 7 | igual | analizada | [Form_FormInicial03ConsultasEmpresaUsuario.md](formularios/Form_FormInicial03ConsultasEmpresaUsuario.md) |
| Form_FormInicial03ConsultasEstado | src/forms/Form_FormInicial03ConsultasEstado.cls | 56 | 13 | 3 | igual | analizada | [Form_FormInicial03ConsultasEstado.md](formularios/Form_FormInicial03ConsultasEstado.md) |
| Form_FormInicial03ConsultasGrado | src/forms/Form_FormInicial03ConsultasGrado.cls | 256 | 30 | 15 | igual | analizada | [Form_FormInicial03ConsultasGrado.md](formularios/Form_FormInicial03ConsultasGrado.md) |
| Form_FormInicial03ConsultasJuridicaTramitacion | src/forms/Form_FormInicial03ConsultasJuridicaTramitacion.cls | 102 | 16 | 6 | igual | analizada | [Form_FormInicial03ConsultasJuridicaTramitacion.md](formularios/Form_FormInicial03ConsultasJuridicaTramitacion.md) |
| Form_FormInicial03ConsultasMotivoHPS | src/forms/Form_FormInicial03ConsultasMotivoHPS.cls | 93 | 16 | 6 | igual | analizada | [Form_FormInicial03ConsultasMotivoHPS.md](formularios/Form_FormInicial03ConsultasMotivoHPS.md) |
| Form_FormInicial03ConsultasNombre | src/forms/Form_FormInicial03ConsultasNombre.cls | 124 | 19 | 8 | igual | analizada | [Form_FormInicial03ConsultasNombre.md](formularios/Form_FormInicial03ConsultasNombre.md) |
| Form_FormInicial03ConsultasProyectoAsignado | src/forms/Form_FormInicial03ConsultasProyectoAsignado.cls | 193 | 15 | 7 | igual | analizada | [Form_FormInicial03ConsultasProyectoAsignado.md](formularios/Form_FormInicial03ConsultasProyectoAsignado.md) |
| Form_FormInicial05Configuracion | src/forms/Form_FormInicial05Configuracion.cls | 79 | 22 | 2 | igual | analizada | [Form_FormInicial05Configuracion.md](formularios/Form_FormInicial05Configuracion.md) |
| Form_FormInicial06BusquedaGlobal | src/forms/Form_FormInicial06BusquedaGlobal.cls | 504 | 22 | 16 | igual | pendiente | — |
| Form_FormInicial07UsuariosSICA | src/forms/Form_FormInicial07UsuariosSICA.cls | 236 | 24 | 9 | igual | analizada | [Form_FormInicial07UsuariosSICA.md](formularios/Form_FormInicial07UsuariosSICA.md) |
| Form_FormInicial07UsuariosSICADatos | src/forms/Form_FormInicial07UsuariosSICADatos.cls | 40 | 48 | 3 | igual | analizada | [Form_FormInicial07UsuariosSICADatos.md](formularios/Form_FormInicial07UsuariosSICADatos.md) |
| Form_FormInicial08UsuariosHistoricos | src/forms/Form_FormInicial08UsuariosHistoricos.cls | 225 | 20 | 8 | igual | pendiente | — |
| Form_FormInicial08UsuariosHistoricosDatos | src/forms/Form_FormInicial08UsuariosHistoricosDatos.cls | 47 | 70 | 4 | igual | analizada | [Form_FormInicial08UsuariosHistoricosDatos.md](formularios/Form_FormInicial08UsuariosHistoricosDatos.md) |
| Form_FormInicial09Indicadores | src/forms/Form_FormInicial09Indicadores.cls | 712 | 28 | 22 | igual | pendiente | — |
| Form_FormInicial09IndicadoresDatosLocalDatos | src/forms/Form_FormInicial09IndicadoresDatosLocalDatos.cls | 42 | 311 | 2 | igual | analizada | [Form_FormInicial09IndicadoresDatosLocalDatos.md](formularios/Form_FormInicial09IndicadoresDatosLocalDatos.md) |
| Form_FormInicial09IndicadoresUsuariosSICADatos | src/forms/Form_FormInicial09IndicadoresUsuariosSICADatos.cls | 47 | 61 | 2 | igual | analizada | [Form_FormInicial09IndicadoresUsuariosSICADatos.md](formularios/Form_FormInicial09IndicadoresUsuariosSICADatos.md) |
| Form_FormUsuario00Principal | src/forms/Form_FormUsuario00Principal.cls | 63 | 18 | 2 | igual | analizada | [Form_FormUsuario00Principal.md](formularios/Form_FormUsuario00Principal.md) |
| Form_FormUsuario01Lateral | src/forms/Form_FormUsuario01Lateral.cls | 167 | 14 | 5 | igual | analizada | [Form_FormUsuario01Lateral.md](formularios/Form_FormUsuario01Lateral.md) |
| Form_FormUsuario02Alta | src/forms/Form_FormUsuario02Alta.cls | 424 | 64 | 11 | igual | pendiente | — |
| Form_FormUsuario02Anexos | src/forms/Form_FormUsuario02Anexos.cls | 590 | 16 | 13 | igual | pendiente | — |
| Form_FormUsuario02DatosObservacion | src/forms/Form_FormUsuario02DatosObservacion.cls | 194 | 23 | 3 | igual | analizada | [Form_FormUsuario02DatosObservacion.md](formularios/Form_FormUsuario02DatosObservacion.md) |
| Form_FormUsuario02DatosObservacionHistorico | src/forms/Form_FormUsuario02DatosObservacionHistorico.cls | 60 | 22 | 2 | igual | analizada | [Form_FormUsuario02DatosObservacionHistorico.md](formularios/Form_FormUsuario02DatosObservacionHistorico.md) |
| Form_FormUsuario02Edicion | src/forms/Form_FormUsuario02Edicion.cls | 977 | 70 | 24 | igual | pendiente | — |
| Form_FormUsuario02HPS | src/forms/Form_FormUsuario02HPS.cls | 988 | 90 | 16 | igual | pendiente | — |
| Form_FormUsuario07SICAAlta | src/forms/Form_FormUsuario07SICAAlta.cls | 516 | 56 | 10 | igual | pendiente | — |
| Form_FormUsuario07SICAEdicion | src/forms/Form_FormUsuario07SICAEdicion.cls | 911 | 57 | 14 | igual | pendiente | — |
| Form_FormUsuarioHistorico00Principal | src/forms/Form_FormUsuarioHistorico00Principal.cls | 98 | 19 | 3 | igual | analizada | [Form_FormUsuarioHistorico00Principal.md](formularios/Form_FormUsuarioHistorico00Principal.md) |
| Form_FormUsuarioHistorico01Lateral | src/forms/Form_FormUsuarioHistorico01Lateral.cls | 130 | 13 | 4 | igual | analizada | [Form_FormUsuarioHistorico01Lateral.md](formularios/Form_FormUsuarioHistorico01Lateral.md) |
| Form_FormUsuarioHistorico02Anexos | src/forms/Form_FormUsuarioHistorico02Anexos.cls | 156 | 8 | 4 | igual | analizada | [Form_FormUsuarioHistorico02Anexos.md](formularios/Form_FormUsuarioHistorico02Anexos.md) |
| Form_FormUsuarioHistorico02Edicion | src/forms/Form_FormUsuarioHistorico02Edicion.cls | 235 | 54 | 5 | igual | analizada | [Form_FormUsuarioHistorico02Edicion.md](formularios/Form_FormUsuarioHistorico02Edicion.md) |
| Form_Formulario1 | src/forms/Form_Formulario1.cls | 28 | 24 | 1 | igual | analizada | [Form_Formulario1.md](formularios/Form_Formulario1.md) |

### Módulos estándar

| Objeto | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|
| CacheConsistencyAudit | src/modules/CacheConsistencyAudit.bas | 373 | — | 18 | igual | pendiente | — |
| Constructor | src/modules/Constructor.bas | 3399 | — | 54 | igual | pendiente | — |
| DummyBig | src/modules/DummyBig.bas | 6002 | — | 2000 | igual | pendiente | — |
| DummyBig2 | src/modules/DummyBig2.bas | 6002 | — | 2000 | igual | pendiente | — |
| DummyHatw | src/modules/DummyHatw.bas | 2035 | — | 74 | igual | pendiente | — |
| DummyHatw2 | src/modules/DummyHatw2.bas | 2033 | — | 74 | igual | pendiente | — |
| DummySmall | src/modules/DummySmall.bas | 3002 | — | 1000 | igual | pendiente | — |
| DummySpecial | src/modules/DummySpecial.bas | 1662 | — | 20 | igual | pendiente | — |
| Filtro | src/modules/Filtro.bas | 353 | — | 4 | igual | pendiente | — |
| Funciones Generales | src/modules/Funciones Generales.bas | 5382 | — | 106 | igual | pendiente | — |
| Instalador | src/modules/Instalador.bas | 288 | — | 6 | igual | pendiente | — |
| JsonConverter | src/modules/JsonConverter.bas | 1403 | — | 31 | igual | pendiente | — |
| Mod_Cache_Core | src/modules/Mod_Cache_Core.bas | 235 | — | 4 | igual | pendiente | — |
| Mod_Sincronizacion_Historico | src/modules/Mod_Sincronizacion_Historico.bas | 246 | — | 3 | igual | pendiente | — |
| Mod_StartupCacheInitialization | src/modules/Mod_StartupCacheInitialization.bas | 414 | — | 15 | igual | pendiente | — |
| Módulo1 | src/modules/Módulo1.bas | 223 | — | 4 | igual | pendiente | — |
| Test_AnexoSelectionTransaction | src/modules/Test_AnexoSelectionTransaction.bas | 1256 | — | 39 | igual | pendiente | — |
| Test_CacheConsistencyAudit | src/modules/Test_CacheConsistencyAudit.bas | 545 | — | 35 | igual | pendiente | — |
| Test_HPSConfig | src/modules/Test_HPSConfig.bas | 1351 | — | 22 | igual | pendiente | — |
| Test_HPSEntorno | src/modules/Test_HPSEntorno.bas | 1054 | — | 21 | igual | pendiente | — |
| Test_HistoricoAdjuntosTransactionWrapper | src/modules/Test_HistoricoAdjuntosTransactionWrapper.bas | 3538 | — | 94 | igual | pendiente | — |
| Test_LocalReadAccessAuthorization | src/modules/Test_LocalReadAccessAuthorization.bas | 391 | — | 18 | igual | pendiente | — |
| Test_PerAnexoMove | src/modules/Test_PerAnexoMove.bas | 479 | — | 15 | igual | pendiente | — |
| Test_RealTimeIndicatorCoherence | src/modules/Test_RealTimeIndicatorCoherence.bas | 403 | — | 15 | igual | pendiente | — |
| Test_StartupCacheInitialization | src/modules/Test_StartupCacheInitialization.bas | 227 | — | 11 | igual | pendiente | — |
| Tests | src/modules/Tests.bas | 126 | — | 5 | igual | pendiente | — |
| VariablesEntorno | src/modules/VariablesEntorno.bas | 760 | — | 10 | igual | pendiente | — |
| cacheSuministrador | src/modules/cacheSuministrador.bas | 170 | — | 3 | igual | pendiente | — |
| cacheUsuario | src/modules/cacheUsuario.bas | 260 | — | 3 | igual | pendiente | — |
| modConfiguracionHPS | src/modules/modConfiguracionHPS.bas | 663 | — | 19 | igual | pendiente | — |
| modIndicadores | src/modules/modIndicadores.bas | 15 | — | 1 | igual | pendiente | — |

### Clases

| Objeto | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|
| AnexoSelectionTransactionCoordinator | src/classes/AnexoSelectionTransactionCoordinator.cls | 511 | — | 6 | igual | pendiente | — |
| AnexoUsuarioHPS | src/classes/AnexoUsuarioHPS.cls | 593 | — | 15 | igual | pendiente | — |
| AnexoUsuarioHistorico | src/classes/AnexoUsuarioHistorico.cls | 353 | — | 9 | igual | pendiente | — |
| AnexoUsuarioSICA | src/classes/AnexoUsuarioSICA.cls | 406 | — | 11 | igual | pendiente | — |
| Aspecto | src/classes/Aspecto.cls | 214 | — | 27 | igual | pendiente | — |
| Consulta | src/classes/Consulta.cls | 88 | — | 3 | igual | pendiente | — |
| Correo | src/classes/Correo.cls | 324 | — | 7 | igual | pendiente | — |
| DatosLocal | src/classes/DatosLocal.cls | 1393 | — | 10 | igual | pendiente | — |
| Entorno | src/classes/Entorno.cls | 3595 | — | 107 | igual | pendiente | — |
| Expediente | src/classes/Expediente.cls | 593 | — | 6 | igual | pendiente | — |
| HPS | src/classes/HPS.cls | 853 | — | 30 | igual | pendiente | — |
| Indicador | src/classes/Indicador.cls | 404 | — | 17 | igual | pendiente | — |
| Observacion | src/classes/Observacion.cls | 347 | — | 9 | igual | pendiente | — |
| ObservacionHistorica | src/classes/ObservacionHistorica.cls | 349 | — | 9 | igual | pendiente | — |
| Suministrador | src/classes/Suministrador.cls | 350 | — | 12 | igual | pendiente | — |
| Usuario | src/classes/Usuario.cls | 449 | — | 8 | igual | pendiente | — |
| UsuarioAplicacionPermisos | src/classes/UsuarioAplicacionPermisos.cls | 400 | — | 19 | igual | pendiente | — |
| UsuarioHPS | src/classes/UsuarioHPS.cls | 2207 | — | 52 | igual | pendiente | — |
| UsuarioHistorico | src/classes/UsuarioHistorico.cls | 1147 | — | 25 | igual | pendiente | — |
| UsuarioLifecycleTransactionCoordinator | src/classes/UsuarioLifecycleTransactionCoordinator.cls | 1265 | — | 48 | igual | pendiente | — |
| UsuarioSICA | src/classes/UsuarioSICA.cls | 1119 | — | 22 | igual | pendiente | — |
| UsuarioServicio | src/classes/UsuarioServicio.cls | 192 | — | 12 | igual | pendiente | — |
| clsIndicadoresBus | src/classes/clsIndicadoresBus.cls | 24 | — | 1 | igual | pendiente | — |
| clsTestDoubleForm | src/classes/clsTestDoubleForm.cls | 23 | — | 1 | igual | pendiente | — |
| clsTestDoubleHpsEditListener | src/classes/clsTestDoubleHpsEditListener.cls | 25 | — | 2 | igual | pendiente | — |
| clsTestDoubleHpsEditPublisher | src/classes/clsTestDoubleHpsEditPublisher.cls | 21 | — | 1 | igual | pendiente | — |
| clsTestDoubleIndicador | src/classes/clsTestDoubleIndicador.cls | 31 | — | 2 | igual | pendiente | — |
| clsTestDoubleIndicadorConsumer | src/classes/clsTestDoubleIndicadorConsumer.cls | 33 | — | 3 | igual | pendiente | — |

Totales: 0 analizados · 0 no aplicables · 97 pendientes.

## Discrepancias abiertas

| ID | Discrepancia | Evidencia | Estado |
|---|---|---|---|
| D-HPS-1 | ERD de origen: falta la extracción de solo lectura del `.accdb` (workflow `dysflow-erd`, #1144) | `ERD/` del legacy | pendiente del propietario |

## Método del recuento

Clon de solo lectura bajo `.atl/sources/hps/` (ruta ignorada) con HEAD desacoplado en el commit fijado e índice CodeGraph-VBA (`codegraph init -i`). Los recuentos salen de un script que cuenta líneas, bloques `Begin` y declaraciones de procedimiento; la comparación con el snapshot usa la normalización descrita arriba.
