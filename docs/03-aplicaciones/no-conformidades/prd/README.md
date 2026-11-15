# PRD — No Conformidades

> **Estado:** en curso · fases 0, 1 y 2 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) completas (48 de 48 fichas), con [matriz de campos](matriz-campos.md), [coexistencia](coexistencia.md), [capacidades](capacidades.md) e [issues propuestas](issues-propuestas.md). Falta el [ERD de origen](erd-origen.md), pendiente del propietario.

| Campo | Valor |
|---|---|
| Repositorio | DysTelefonica/No_conformidades |
| Commit fijado | `1d7a2f826a1e3c582874e4e29886a6183efb9067` (`main`, 2026-09-15, «docs(uat, roadmap): advance release-2026-009 to uat_handoff per HR-11») |
| Índice CodeGraph-VBA | 162 ficheros indexados, 8 335 nodos, versión 2.1.1, sobre clon de solo lectura de ese commit |
| Estudio previo | [epic.md](../epic.md) (DRAFT), [forms.md](../forms.md), [capabilities.md](../capabilities.md), [data-model.md](../data-model.md), [integrations-automation.md](../integrations-automation.md), [security-rules.md](../security-rules.md), [migration-matrix.md](../migration-matrix.md), [walkthroughs G1 a G5](../walkthrough-G1.json), [ERD/schema.sql](../ERD/schema.sql) |
| Línea base | `main` del repositorio legacy (producción), por decisión del propietario registrada en [#1080](https://github.com/DysTelefonica/access2web-blueprint/issues/1080#issuecomment-6322559873) |
| Snapshot comparado | `data/staging/no-conformidades/src` (2026-08-06), tomado de la rama `staging`; solo se usa para clasificar el delta |
| Issue | #1080 (método y plantilla de #1069) |

## Alcance actual

Fija el commit, concilia el inventario y enlaza las 48 fichas de formulario. Incluye la matriz de campos, la coexistencia, las capacidades y las issues propuestas; el ERD de origen queda pendiente del propietario. Los recuentos de líneas, controles y procedimientos son mecánicos (véase [Método del recuento](#método-del-recuento)) y no sustituyen al análisis.

## Inventario reconciliado

### Recuentos por tipo

| Tipo | `1d7a2f8` | Snapshot 2026-08-06 | Diferencia |
|---|---|---|---|
| Formularios (`src/forms/*.cls`, con su `.form.txt`) | 48 | 48 | 0 (mismos nombres); 25 con código distinto |
| Módulos estándar (`src/modules/*.bas`) | 16 | 105 | −90 retirados, +1 añadido, 7 modificados |
| Clases (`src/classes/*.cls`) | 41 | 48 | −8 retiradas, +1 añadida, 13 modificadas |
| Consultas guardadas (`src/queries/`) | 3 (`Consulta1` a `Consulta3`) | sin recuento | — |

La comparación normaliza las líneas `Attribute VB_`, los comentarios, las líneas en blanco y las diferencias de espacios y mayúsculas. **Un delta de 0 formularios no significa que el código coincida:** 25 de los 48 formularios difieren del snapshot. La causa está en el origen del snapshot (véase D-NC-1).

### Formularios

Líneas: `.cls` del formulario. Controles: bloques `Begin` del `.form.txt`. Procedimientos: `Sub` y `Function` del `.cls`. «Frente al snapshot»: **igual** si el `.cls` coincide con el del snapshot tras la normalización anterior; **modificado** si difiere.

| Objeto | Tipo | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|---|
| Form_Form0BDOpciones | formulario | src/forms/Form_Form0BDOpciones.cls | 205 | 22 | 6 | igual | analizado | [formularios/Form_Form0BDOpciones.md](formularios/Form_Form0BDOpciones.md) |
| Form_Form0BDOpcionesAuditorias | formulario | src/forms/Form_Form0BDOpcionesAuditorias.cls | 290 | 31 | 10 | modificado | analizado | [formularios/Form_Form0BDOpcionesAuditorias.md](formularios/Form_Form0BDOpcionesAuditorias.md) |
| Form_Form0BDOpcionesParteProyectos | formulario | src/forms/Form_Form0BDOpcionesParteProyectos.cls | 369 | 31 | 13 | modificado | analizado | [formularios/Form_Form0BDOpcionesParteProyectos.md](formularios/Form_Form0BDOpcionesParteProyectos.md) |
| Form_Form0BDTecnicos | formulario | src/forms/Form_Form0BDTecnicos.cls | 191 | 23 | 6 | igual | analizado | [formularios/Form_Form0BDTecnicos.md](formularios/Form_Form0BDTecnicos.md) |
| Form_FormARAuditoriaDocumentos | formulario | src/forms/Form_FormARAuditoriaDocumentos.cls | 462 | 23 | 11 | modificado | analizado | [formularios/Form_FormARAuditoriaDocumentos.md](formularios/Form_FormARAuditoriaDocumentos.md) |
| Form_FormARProyectoDocumentos | formulario | src/forms/Form_FormARProyectoDocumentos.cls | 517 | 23 | 12 | modificado | analizado | [formularios/Form_FormARProyectoDocumentos.md](formularios/Form_FormARProyectoDocumentos.md) |
| Form_FormAuditoria | formulario | src/forms/Form_FormAuditoria.cls | 187 | 25 | 5 | igual | analizado | [formularios/Form_FormAuditoria.md](formularios/Form_FormAuditoria.md) |
| Form_FormAuditoriaDocumentos | formulario | src/forms/Form_FormAuditoriaDocumentos.cls | 580 | 23 | 12 | modificado | analizado | [formularios/Form_FormAuditoriaDocumentos.md](formularios/Form_FormAuditoriaDocumentos.md) |
| Form_FormAuditoriaSeleccion | formulario | src/forms/Form_FormAuditoriaSeleccion.cls | 116 | 18 | 4 | igual | analizado | [formularios/Form_FormAuditoriaSeleccion.md](formularios/Form_FormAuditoriaSeleccion.md) |
| Form_FormAuditoriasGestion | formulario | src/forms/Form_FormAuditoriasGestion.cls | 741 | 30 | 19 | modificado | analizado | [formularios/Form_FormAuditoriasGestion.md](formularios/Form_FormAuditoriasGestion.md) |
| Form_FormCorreo | formulario | src/forms/Form_FormCorreo.cls | 156 | 24 | 4 | igual | analizado | [formularios/Form_FormCorreo.md](formularios/Form_FormCorreo.md) |
| Form_FormExpedientesBusqueda | formulario | src/forms/Form_FormExpedientesBusqueda.cls | 379 | 24 | 11 | igual | analizado | [formularios/Form_FormExpedientesBusqueda.md](formularios/Form_FormExpedientesBusqueda.md) |
| Form_FormIndicadores | formulario | src/forms/Form_FormIndicadores.cls | 155 | 40 | 10 | igual | analizado | [formularios/Form_FormIndicadores.md](formularios/Form_FormIndicadores.md) |
| Form_FormMotivosNoRequiereControlEficacia | formulario | src/forms/Form_FormMotivosNoRequiereControlEficacia.cls | 28 | 19 | 3 | igual | analizado | [formularios/Form_FormMotivosNoRequiereControlEficacia.md](formularios/Form_FormMotivosNoRequiereControlEficacia.md) |
| Form_FormNCAuditoria | formulario | src/forms/Form_FormNCAuditoria.cls | 219 | 24 | 5 | igual | analizado | [formularios/Form_FormNCAuditoria.md](formularios/Form_FormNCAuditoria.md) |
| Form_FormNCAuditoriaAC | formulario | src/forms/Form_FormNCAuditoriaAC.cls | 328 | 34 | 7 | igual | analizado | [formularios/Form_FormNCAuditoriaAC.md](formularios/Form_FormNCAuditoriaAC.md) |
| Form_FormNCAuditoriaAR | formulario | src/forms/Form_FormNCAuditoriaAR.cls | 654 | 36 | 16 | modificado | analizado | [formularios/Form_FormNCAuditoriaAR.md](formularios/Form_FormNCAuditoriaAR.md) |
| Form_FormNCAuditoriaAcciones | formulario | src/forms/Form_FormNCAuditoriaAcciones.cls | 1398 | 32 | 30 | modificado | analizado | [formularios/Form_FormNCAuditoriaAcciones.md](formularios/Form_FormNCAuditoriaAcciones.md) |
| Form_FormNCAuditoriaControlEficacia | formulario | src/forms/Form_FormNCAuditoriaControlEficacia.cls | 416 | 23 | 13 | modificado | analizado | [formularios/Form_FormNCAuditoriaControlEficacia.md](formularios/Form_FormNCAuditoriaControlEficacia.md) |
| Form_FormNCAuditoriaControlEficaciaAlta | formulario | src/forms/Form_FormNCAuditoriaControlEficaciaAlta.cls | 239 | 23 | 9 | modificado | analizado | [formularios/Form_FormNCAuditoriaControlEficaciaAlta.md](formularios/Form_FormNCAuditoriaControlEficaciaAlta.md) |
| Form_FormNCAuditoriaDocumentos | formulario | src/forms/Form_FormNCAuditoriaDocumentos.cls | 511 | 19 | 11 | modificado | analizado | [formularios/Form_FormNCAuditoriaDocumentos.md](formularios/Form_FormNCAuditoriaDocumentos.md) |
| Form_FormNCAuditoriaGeneral | formulario | src/forms/Form_FormNCAuditoriaGeneral.cls | 648 | 43 | 18 | modificado | analizado | [formularios/Form_FormNCAuditoriaGeneral.md](formularios/Form_FormNCAuditoriaGeneral.md) |
| Form_FormNCAuditoriaGestion | formulario | src/forms/Form_FormNCAuditoriaGestion.cls | 1338 | 52 | 40 | modificado | analizado | [formularios/Form_FormNCAuditoriaGestion.md](formularios/Form_FormNCAuditoriaGestion.md) |
| Form_FormNCAuditoriaMotivoEliminado | formulario | src/forms/Form_FormNCAuditoriaMotivoEliminado.cls | 211 | 13 | 7 | igual | analizado | [formularios/Form_FormNCAuditoriaMotivoEliminado.md](formularios/Form_FormNCAuditoriaMotivoEliminado.md) |
| Form_FormNCAuditoriaNota | formulario | src/forms/Form_FormNCAuditoriaNota.cls | 241 | 14 | 6 | igual | analizado | [formularios/Form_FormNCAuditoriaNota.md](formularios/Form_FormNCAuditoriaNota.md) |
| Form_FormNCAuditoriaReplanificaciones | formulario | src/forms/Form_FormNCAuditoriaReplanificaciones.cls | 177 | 18 | 4 | igual | analizado | [formularios/Form_FormNCAuditoriaReplanificaciones.md](formularios/Form_FormNCAuditoriaReplanificaciones.md) |
| Form_FormNCAuditoriaSeguimiento | formulario | src/forms/Form_FormNCAuditoriaSeguimiento.cls | 69 | 25 | 3 | modificado | analizado | [formularios/Form_FormNCAuditoriaSeguimiento.md](formularios/Form_FormNCAuditoriaSeguimiento.md) |
| Form_FormNCAuditoriaSeguimientoNC | formulario | src/forms/Form_FormNCAuditoriaSeguimientoNC.cls | 358 | 24 | 13 | modificado | analizado | [formularios/Form_FormNCAuditoriaSeguimientoNC.md](formularios/Form_FormNCAuditoriaSeguimientoNC.md) |
| Form_FormNCAuditoriaSeguimientoTareas | formulario | src/forms/Form_FormNCAuditoriaSeguimientoTareas.cls | 363 | 24 | 13 | modificado | analizado | [formularios/Form_FormNCAuditoriaSeguimientoTareas.md](formularios/Form_FormNCAuditoriaSeguimientoTareas.md) |
| Form_FormNCProyecto | formulario | src/forms/Form_FormNCProyecto.cls | 212 | 24 | 5 | igual | analizado | [formularios/Form_FormNCProyecto.md](formularios/Form_FormNCProyecto.md) |
| Form_FormNCProyectoAC | formulario | src/forms/Form_FormNCProyectoAC.cls | 339 | 34 | 7 | igual | analizado | [formularios/Form_FormNCProyectoAC.md](formularios/Form_FormNCProyectoAC.md) |
| Form_FormNCProyectoAR | formulario | src/forms/Form_FormNCProyectoAR.cls | 674 | 36 | 16 | modificado | analizado | [formularios/Form_FormNCProyectoAR.md](formularios/Form_FormNCProyectoAR.md) |
| Form_FormNCProyectoAcciones | formulario | src/forms/Form_FormNCProyectoAcciones.cls | 1458 | 32 | 30 | modificado | analizado | [formularios/Form_FormNCProyectoAcciones.md](formularios/Form_FormNCProyectoAcciones.md) |
| Form_FormNCProyectoControlEficacia | formulario | src/forms/Form_FormNCProyectoControlEficacia.cls | 449 | 23 | 13 | modificado | analizado | [formularios/Form_FormNCProyectoControlEficacia.md](formularios/Form_FormNCProyectoControlEficacia.md) |
| Form_FormNCProyectoControlEficaciaAlta | formulario | src/forms/Form_FormNCProyectoControlEficaciaAlta.cls | 243 | 23 | 9 | igual | analizado | [formularios/Form_FormNCProyectoControlEficaciaAlta.md](formularios/Form_FormNCProyectoControlEficaciaAlta.md) |
| Form_FormNCProyectoDocumentos | formulario | src/forms/Form_FormNCProyectoDocumentos.cls | 535 | 19 | 11 | modificado | analizado | [formularios/Form_FormNCProyectoDocumentos.md](formularios/Form_FormNCProyectoDocumentos.md) |
| Form_FormNCProyectoGeneral | formulario | src/forms/Form_FormNCProyectoGeneral.cls | 1229 | 58 | 29 | modificado | analizado | [formularios/Form_FormNCProyectoGeneral.md](formularios/Form_FormNCProyectoGeneral.md) |
| Form_FormNCProyectoGeneralConVinculoNC | formulario | src/forms/Form_FormNCProyectoGeneralConVinculoNC.cls | 454 | 37 | 13 | igual | analizado | [formularios/Form_FormNCProyectoGeneralConVinculoNC.md](formularios/Form_FormNCProyectoGeneralConVinculoNC.md) |
| Form_FormNCProyectoGestion | formulario | src/forms/Form_FormNCProyectoGestion.cls | 1827 | 73 | 56 | modificado | analizado | [formularios/Form_FormNCProyectoGestion.md](formularios/Form_FormNCProyectoGestion.md) |
| Form_FormNCProyectoMotivoEliminado | formulario | src/forms/Form_FormNCProyectoMotivoEliminado.cls | 224 | 13 | 7 | igual | analizado | [formularios/Form_FormNCProyectoMotivoEliminado.md](formularios/Form_FormNCProyectoMotivoEliminado.md) |
| Form_FormNCProyectoNota | formulario | src/forms/Form_FormNCProyectoNota.cls | 270 | 14 | 6 | igual | analizado | [formularios/Form_FormNCProyectoNota.md](formularios/Form_FormNCProyectoNota.md) |
| Form_FormNCProyectoReplanificaciones | formulario | src/forms/Form_FormNCProyectoReplanificaciones.cls | 181 | 18 | 4 | igual | analizado | [formularios/Form_FormNCProyectoReplanificaciones.md](formularios/Form_FormNCProyectoReplanificaciones.md) |
| Form_FormNCProyectoSeguimiento | formulario | src/forms/Form_FormNCProyectoSeguimiento.cls | 75 | 26 | 3 | modificado | analizado | [formularios/Form_FormNCProyectoSeguimiento.md](formularios/Form_FormNCProyectoSeguimiento.md) |
| Form_FormNCProyectoSeguimientoNC | formulario | src/forms/Form_FormNCProyectoSeguimientoNC.cls | 550 | 32 | 18 | modificado | analizado | [formularios/Form_FormNCProyectoSeguimientoNC.md](formularios/Form_FormNCProyectoSeguimientoNC.md) |
| Form_FormNCProyectoSeguimientoTareas | formulario | src/forms/Form_FormNCProyectoSeguimientoTareas.cls | 599 | 33 | 19 | modificado | analizado | [formularios/Form_FormNCProyectoSeguimientoTareas.md](formularios/Form_FormNCProyectoSeguimientoTareas.md) |
| Form_FormNCProyectoTipologiaGestion | formulario | src/forms/Form_FormNCProyectoTipologiaGestion.cls | 362 | 21 | 10 | igual | analizado | [formularios/Form_FormNCProyectoTipologiaGestion.md](formularios/Form_FormNCProyectoTipologiaGestion.md) |
| Form_FormTipologiaNCProyecto | formulario | src/forms/Form_FormTipologiaNCProyecto.cls | 245 | 19 | 6 | igual | analizado | [formularios/Form_FormTipologiaNCProyecto.md](formularios/Form_FormTipologiaNCProyecto.md) |
| Form_formRiesgosSeleccion | formulario | src/forms/Form_formRiesgosSeleccion.cls | 206 | 21 | 7 | igual | analizado | [formularios/Form_formRiesgosSeleccion.md](formularios/Form_formRiesgosSeleccion.md) |

Totales (formularios): 48 analizados · 0 no aplicables · 0 pendientes. «Analizado» significa ficha con citas `fichero:línea` comprobadas; los pendientes de cada ficha (referencias dinámicas sin resolver) se listan en su apartado «Pendientes» y no cuentan como resueltos.
Totales mecánicos: 1 316 bloques `Begin` · 590 procedimientos · 21 678 líneas de `.cls`.

### Módulos y clases

Pendiente de análisis. Módulos (16): `CacheNCProyecto.bas`, `Funciones Generales.bas`, `HTML.bas`, `IndicadorRepositorio.bas`, `InicializadorCache.bas`, `Instalador.bas`, `JSONHelper.bas`, `JsonConverter.bas`, `Módulo1.bas`, `RiesgoRepositorio.bas`, `RiesgoServicio1.bas`, `TestCacheNC.bas`, `Variables Globales.bas`, `constantes.bas`, `constructor.bas`, `mdlCursor.bas`. Clases (41): las de dominio de No Conformidad (`NCAuditoria`, `NCProyecto` y sus `*Operaciones`), acciones correctivas y preventivas (`ACAuditoria`, `ACProyecto`, `ARAuditoria`, `ARProyecto` y sus `*Operaciones`), documentos, replanificaciones, seguimientos (`SegNC*`, `SegTareas*`), registros (`LogNC*`), `Auditoria`, `Entorno`, `Usuario`, `UsuarioAplicacionPermisos`, `Correo`, `Informe`, `InformeNCAuditorias`, `Riesgo`, `RiesgoServicio`, `Expediente`, `ExpedienteResponsable`, `IndicadorServicio`, `Juridica` y `TipologiaNCProyectos`.

## Delta frente al estudio previo

El estudio previo describe la rama `staging`; el commit fijado es `main`, que el historial del repositorio marca como línea base de producción. El detalle por capacidad queda para `capacidades.md`; el delta mecánico de objetos es:

| Ámbito | Clasificación mecánica | Evidencia |
|---|---|---|
| Formularios | 23 sin cambios · 25 modificados · 0 añadidos · 0 retirados | recuento de la tabla anterior |
| Módulos | 90 retirados (capa de caché de indicadores y de listados, ayudantes `modFrm*Helper`, bancos `Test_*`) · 1 añadido (`RiesgoServicio1.bas`) · 7 modificados | `src/modules/` frente al snapshot |
| Clases | 8 retiradas (`CacheNCCacheRepositorio`, `CacheNCService`, `CacheNCCrud`, `CierreNCService`, `ReplanificacionesService`, `DocumentoService`, `NCProyectoDetailVM`, `NCProyectoListItemVM`) · 1 añadida (`InformeNCAuditorias.cls`) · 13 modificadas | `src/classes/` frente al snapshot |

## Capacidades

45 capacidades (CAP-NC-1 a CAP-NC-45) clasificadas frente al estudio previo en [capacidades.md](capacidades.md): 8 sin cambios, 1 mixta, 30 modificadas, 1 añadida y 5 retiradas. Las propuestas de implementación están en [issues-propuestas.md](issues-propuestas.md) (sin publicar). Las capacidades del [estudio previo](../capabilities.md) son contexto, no evidencia.

## Discrepancias abiertas

| Id | Discrepancia | Evidencia | Estado |
|---|---|---|---|
| D-NC-1 | El snapshot del estudio previo procede de la rama `staging` (248 commits por delante de `main`, último 2026-09-16), mientras que el commit fijado es `main`. | `git rev-list --count 1d7a2f8..origin/staging` = 248; `git rev-list --count origin/staging..1d7a2f8` = 9 | Resuelta: la línea base es `main` (producción). Lo que solo existe en `staging` va al [anexo de cambios no liberados](#anexo-cambios-no-liberados), sin fichas propias; los módulos y clases ausentes de `main` se listan como retirados.  |
| D-NC-2 | #1080 cita como fecha del commit de partida 2026-10-21; el commit `1d7a2f8` es del 2026-09-15 (el 2026-10-21 es el día del recuento). | `git log -1 1d7a2f8` | Resuelta: se cita la fecha real. |
| D-NC-3 | El ERD de origen no se ha extraído: `dysflow` no está disponible en el entorno de trabajo y los backends `.accdb` no son legibles con herramientas de solo lectura. | véase `erd-origen.md` cuando llegue | Abierta, **pendiente del propietario** (runner Windows `dysflow-erd`). |
| D-NC-4 | El volcado `ERD/schema.sql` declara 49 tablas y su DDL contiene 42. | `-- Tables: 49` en la cabecera del volcado frente a 42 `CREATE TABLE` | Abierta, **pendiente del propietario**: la extracciÃ³n de [erd-origen.md](erd-origen.md) lo resuelve. |

## Anexo: cambios no liberados

Cambios de la rama `staging` del repositorio legacy que **no** están en producción (`main`). Son contexto: no tienen ficha ni forman parte del alcance de migración hasta que se liberen. Datos mecánicos medidos con `git` sobre el clon de solo lectura.

| Medida | Valor |
|---|---|
| Commits en `staging` que no están en el commit fijado | 248 (2026-05-28 a 2026-09-16) |
| Commits del commit fijado que no están en `staging` | 9 |
| Ficheros de `src` distintos entre ambos | 216 (+47 890 / −2 290 líneas) |
| Formularios con código distinto del snapshot | 25 de 48 |

Líneas de trabajo no liberadas, según el tema de los 18 commits `feat` (sin analizar su código): capa de caché de indicadores y de listados de NC (issues 18 y siguientes), invalidación de caché tras escrituras, informe y cola de correo programados sobre la caché, ayudantes de prueba de formularios, validación de `IDAplicacion` al arranque (con su reversión posterior) y aplazamiento de `FechaPrevistaControlEficacia` hasta el cierre de la NC.

### Formularios con código distinto del snapshot (25)

`Form_Form0BDOpcionesAuditorias`, `Form_Form0BDOpcionesParteProyectos`, `Form_FormARAuditoriaDocumentos`, `Form_FormARProyectoDocumentos`, `Form_FormAuditoriaDocumentos`, `Form_FormAuditoriasGestion`, `Form_FormNCAuditoriaAR`, `Form_FormNCAuditoriaAcciones`, `Form_FormNCAuditoriaControlEficacia`, `Form_FormNCAuditoriaControlEficaciaAlta`, `Form_FormNCAuditoriaDocumentos`, `Form_FormNCAuditoriaGeneral`, `Form_FormNCAuditoriaGestion`, `Form_FormNCAuditoriaSeguimiento`, `Form_FormNCAuditoriaSeguimientoNC`, `Form_FormNCAuditoriaSeguimientoTareas`, `Form_FormNCProyectoAR`, `Form_FormNCProyectoAcciones`, `Form_FormNCProyectoControlEficacia`, `Form_FormNCProyectoDocumentos`, `Form_FormNCProyectoGeneral`, `Form_FormNCProyectoGestion`, `Form_FormNCProyectoSeguimiento`, `Form_FormNCProyectoSeguimientoNC`, `Form_FormNCProyectoSeguimientoTareas`.

## Objetos retirados frente al estudio previo

Existen en el snapshot y no en `main`. Se listan sin ficha.

**Módulos (90):** `CacheTrustDiagnostics.bas`, `EstadoCatalogoBootstrap.bas`, `ModuloCacheIndicadores.bas`, `ModuloCacheIndicadoresIssue18.bas`, `ModuloCacheWarmupWrapper.bas`, `ModuloMigracionIssue18.bas`, `NCAuditoriaGestionListadoHelper.bas`, `NCAuditoriaListadoCache.bas`, `NCAuditoriaSeguimientoHelper.bas`, `NCProyectoGestionListadoHelper.bas`, `NCProyectoSeguimientoHelper.bas`, `TestHelper.bas`, `Test_BackendConfigPaths.bas`, `Test_CAT_MaestrosCatalogos.bas`, `Test_COM_ComunicacionesReportes.bas`, `Test_CacheListadoEstados.bas`, `Test_CacheListadoNC_ACAR.bas`, `Test_CacheListadoNC_Parity.bas`, `Test_CacheTrustDiagnostics.bas`, `Test_CacheWarmupWrapper.bas`, `Test_CacheWriteInvalidation.bas`, `Test_Cache_PrepararStagingConCaches.bas`, `Test_E2E_BateriaNC.bas`, `Test_EXP_ExpedientesRiesgosResponsables.bas`, `Test_EstadoCatalogoBootstrap.bas`, `Test_FormHelper_Coverage.bas`, `Test_IndicadoresCaracterizacion.bas`, `Test_IndicadoresTelemetry.bas`, `Test_Issue19_CEGating.bas`, `Test_KillSwitch.bas`, `Test_NCA_AccionesSeguimiento.bas`, `Test_NCAuditoriaGestionListadoHelper.bas`, `Test_NCProyectoGestionListadoHelper.bas`, `Test_NCProyectoSeguimientoHelper_FormOpenConfig.bas`, `Test_ProyectoGestionHelper_CacheFirstContract.bas`, `Test_Spec007b_General.bas`, `Test_UPN_UsuariosPermisos.bas`, `Test_modFormInteractionHelper.bas`, `Test_modFrmARAuditoriaDocumentosHelper.bas`, `Test_modFrmARProyectoDocumentosHelper.bas`, `Test_modFrmAuditoriaDocumentosHelper.bas`, `Test_modFrmAuditoriasGestionHelper.bas`, `Test_modFrmNCAuditoriaARHelper.bas`, `Test_modFrmNCAuditoriaAccionesHelper.bas`, `Test_modFrmNCAuditoriaControlEficaciaAltaHelper.bas`, `Test_modFrmNCAuditoriaControlEficaciaHelper.bas`, `Test_modFrmNCAuditoriaDocumentosHelper.bas`, `Test_modFrmNCAuditoriaGeneralHelper.bas`, `Test_modFrmNCAuditoriaGestionHelper.bas`, `Test_modFrmNCAuditoriaSeguimientoNCHelper.bas`, `Test_modFrmNCAuditoriaSeguimientoTareasHelper.bas`, `Test_modFrmNCProyectoARHelper.bas`, `Test_modFrmNCProyectoAccionesHelper.bas`, `Test_modFrmNCProyectoControlEficaciaAltaHelper.bas`, `Test_modFrmNCProyectoControlEficaciaHelper.bas`, `Test_modFrmNCProyectoDocumentosHelper.bas`, `Test_modFrmNCProyectoGeneralHelper.bas`, `Test_modFrmNCProyectoGestionHelper.bas`, `Test_modFrmNCProyectoSeguimientoNCHelper.bas`, `Test_modFrmNCProyectoSeguimientoTareasHelper.bas`, `Test_modFrmNCProyectoSeguimientoTareasListadoHelper.bas`, `modFormInteractionHelper.bas`, `modFrmARAuditoriaDocumentosHelper.bas`, `modFrmARProyectoDocumentosHelper.bas`, `modFrmAuditoriaDocumentosHelper.bas`, `modFrmAuditoriasGestionHelper.bas`, `modFrmNCAuditoriaARHelper.bas`, `modFrmNCAuditoriaAccionesHelper.bas`, `modFrmNCAuditoriaControlEficaciaAltaHelper.bas`, `modFrmNCAuditoriaControlEficaciaHelper.bas`, `modFrmNCAuditoriaDocumentosHelper.bas`, `modFrmNCAuditoriaGeneralHelper.bas`, `modFrmNCAuditoriaGestionHelper.bas`, `modFrmNCAuditoriaSeguimientoNCHelper.bas`, `modFrmNCAuditoriaSeguimientoTareasHelper.bas`, `modFrmNCProyectoARHelper.bas`, `modFrmNCProyectoAccionesHelper.bas`, `modFrmNCProyectoControlEficaciaAltaHelper.bas`, `modFrmNCProyectoControlEficaciaHelper.bas`, `modFrmNCProyectoDocumentosHelper.bas`, `modFrmNCProyectoGeneralHelper.bas`, `modFrmNCProyectoGestionHelper.bas`, `modFrmNCProyectoSeguimientoNCHelper.bas`, `modFrmNCProyectoSeguimientoTareasHelper.bas`, `modFrmNCProyectoSeguimientoTareasListadoHelper.bas`, `modPipeFlatten.bas`, `modSegTareasHydration.bas`, `modSeguimientoTareasHydration.bas`, `modStub_RebuildNCProyectoListadoCache.bas`, `modTestingCoreHelper.bas`.

**Clases (8):** `CacheNCCacheRepositorio.cls`, `CacheNCCrud.cls`, `CacheNCService.cls`, `CierreNCService.cls`, `DocumentoService.cls`, `NCProyectoDetailVM.cls`, `NCProyectoListItemVM.cls`, `ReplanificacionesService.cls`.

## Método del recuento

- El clon vive en `.atl/sources/nc-1d7a2f8/` (ruta ignorada, solo lectura) y está en `HEAD` desacoplado en el commit fijado.
- Controles: líneas `^\s*Begin ` del `.form.txt` anteriores al código embebido. Procedimientos: líneas `(Private |Public |Friend )?(Static )?(Sub|Function) ` del `.cls`.
- Comparación con el snapshot: normalización descrita arriba, sobre `.cls` y `.bas`. No se comparan los `.form.txt`.
- Recuento de CodeGraph: tabla `nodes` de `.codegraph/codegraph.db` del clon.
