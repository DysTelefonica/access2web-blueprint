# PRD — No Conformidades

> **Estado:** en curso · fases 0 (fijación) y 1 (inventario) del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md). Ningún formulario está analizado todavía.

| Campo | Valor |
|---|---|
| Repositorio | DysTelefonica/No_conformidades |
| Commit fijado | `1d7a2f826a1e3c582874e4e29886a6183efb9067` (`main`, 2026-09-15, «docs(uat, roadmap): advance release-2026-009 to uat_handoff per HR-11») |
| Índice CodeGraph-VBA | 162 ficheros indexados, 8 335 nodos, versión 2.1.1, sobre clon de solo lectura de ese commit |
| Estudio previo | [epic.md](../epic.md) (DRAFT), [forms.md](../forms.md), [capabilities.md](../capabilities.md), [data-model.md](../data-model.md), [integrations-automation.md](../integrations-automation.md), [security-rules.md](../security-rules.md), [migration-matrix.md](../migration-matrix.md), [walkthroughs G1 a G5](../walkthrough-G1.json), [ERD/schema.sql](../ERD/schema.sql) |
| Snapshot comparado | `data/staging/no-conformidades/src` (2026-08-06), tomado de la rama `staging` del repositorio legacy |
| Issue | #1080 (método y plantilla de #1069) |

## Alcance actual

Fija el commit, concilia el inventario y deja todos los formularios como **pendientes**. No contiene fichas, ERD, matriz, coexistencia ni capacidades: llegan en las tandas siguientes de la cadena. Los recuentos de líneas, controles y procedimientos son mecánicos (véase [Método del recuento](#método-del-recuento)) y no sustituyen al análisis.

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
| Form_Form0BDOpciones | formulario | src/forms/Form_Form0BDOpciones.cls | 205 | 22 | 6 | igual | pendiente | sin ficha todavía |
| Form_Form0BDOpcionesAuditorias | formulario | src/forms/Form_Form0BDOpcionesAuditorias.cls | 290 | 31 | 10 | modificado | pendiente | sin ficha todavía |
| Form_Form0BDOpcionesParteProyectos | formulario | src/forms/Form_Form0BDOpcionesParteProyectos.cls | 369 | 31 | 13 | modificado | pendiente | sin ficha todavía |
| Form_Form0BDTecnicos | formulario | src/forms/Form_Form0BDTecnicos.cls | 191 | 23 | 6 | igual | pendiente | sin ficha todavía |
| Form_FormARAuditoriaDocumentos | formulario | src/forms/Form_FormARAuditoriaDocumentos.cls | 462 | 23 | 11 | modificado | pendiente | sin ficha todavía |
| Form_FormARProyectoDocumentos | formulario | src/forms/Form_FormARProyectoDocumentos.cls | 517 | 23 | 12 | modificado | pendiente | sin ficha todavía |
| Form_FormAuditoria | formulario | src/forms/Form_FormAuditoria.cls | 187 | 25 | 5 | igual | pendiente | sin ficha todavía |
| Form_FormAuditoriaDocumentos | formulario | src/forms/Form_FormAuditoriaDocumentos.cls | 580 | 23 | 12 | modificado | pendiente | sin ficha todavía |
| Form_FormAuditoriaSeleccion | formulario | src/forms/Form_FormAuditoriaSeleccion.cls | 116 | 18 | 4 | igual | pendiente | sin ficha todavía |
| Form_FormAuditoriasGestion | formulario | src/forms/Form_FormAuditoriasGestion.cls | 741 | 30 | 19 | modificado | pendiente | sin ficha todavía |
| Form_FormCorreo | formulario | src/forms/Form_FormCorreo.cls | 156 | 24 | 4 | igual | pendiente | sin ficha todavía |
| Form_FormExpedientesBusqueda | formulario | src/forms/Form_FormExpedientesBusqueda.cls | 379 | 24 | 11 | igual | pendiente | sin ficha todavía |
| Form_FormIndicadores | formulario | src/forms/Form_FormIndicadores.cls | 155 | 40 | 10 | igual | pendiente | sin ficha todavía |
| Form_FormMotivosNoRequiereControlEficacia | formulario | src/forms/Form_FormMotivosNoRequiereControlEficacia.cls | 28 | 19 | 3 | igual | pendiente | sin ficha todavía |
| Form_FormNCAuditoria | formulario | src/forms/Form_FormNCAuditoria.cls | 219 | 24 | 5 | igual | pendiente | sin ficha todavía |
| Form_FormNCAuditoriaAC | formulario | src/forms/Form_FormNCAuditoriaAC.cls | 328 | 34 | 7 | igual | pendiente | sin ficha todavía |
| Form_FormNCAuditoriaAR | formulario | src/forms/Form_FormNCAuditoriaAR.cls | 654 | 36 | 16 | modificado | pendiente | sin ficha todavía |
| Form_FormNCAuditoriaAcciones | formulario | src/forms/Form_FormNCAuditoriaAcciones.cls | 1398 | 32 | 30 | modificado | pendiente | sin ficha todavía |
| Form_FormNCAuditoriaControlEficacia | formulario | src/forms/Form_FormNCAuditoriaControlEficacia.cls | 416 | 23 | 13 | modificado | pendiente | sin ficha todavía |
| Form_FormNCAuditoriaControlEficaciaAlta | formulario | src/forms/Form_FormNCAuditoriaControlEficaciaAlta.cls | 239 | 23 | 9 | modificado | pendiente | sin ficha todavía |
| Form_FormNCAuditoriaDocumentos | formulario | src/forms/Form_FormNCAuditoriaDocumentos.cls | 511 | 19 | 11 | modificado | pendiente | sin ficha todavía |
| Form_FormNCAuditoriaGeneral | formulario | src/forms/Form_FormNCAuditoriaGeneral.cls | 648 | 43 | 18 | modificado | pendiente | sin ficha todavía |
| Form_FormNCAuditoriaGestion | formulario | src/forms/Form_FormNCAuditoriaGestion.cls | 1338 | 52 | 40 | modificado | pendiente | sin ficha todavía |
| Form_FormNCAuditoriaMotivoEliminado | formulario | src/forms/Form_FormNCAuditoriaMotivoEliminado.cls | 211 | 13 | 7 | igual | pendiente | sin ficha todavía |
| Form_FormNCAuditoriaNota | formulario | src/forms/Form_FormNCAuditoriaNota.cls | 241 | 14 | 6 | igual | pendiente | sin ficha todavía |
| Form_FormNCAuditoriaReplanificaciones | formulario | src/forms/Form_FormNCAuditoriaReplanificaciones.cls | 177 | 18 | 4 | igual | pendiente | sin ficha todavía |
| Form_FormNCAuditoriaSeguimiento | formulario | src/forms/Form_FormNCAuditoriaSeguimiento.cls | 69 | 25 | 3 | modificado | pendiente | sin ficha todavía |
| Form_FormNCAuditoriaSeguimientoNC | formulario | src/forms/Form_FormNCAuditoriaSeguimientoNC.cls | 358 | 24 | 13 | modificado | pendiente | sin ficha todavía |
| Form_FormNCAuditoriaSeguimientoTareas | formulario | src/forms/Form_FormNCAuditoriaSeguimientoTareas.cls | 363 | 24 | 13 | modificado | pendiente | sin ficha todavía |
| Form_FormNCProyecto | formulario | src/forms/Form_FormNCProyecto.cls | 212 | 24 | 5 | igual | pendiente | sin ficha todavía |
| Form_FormNCProyectoAC | formulario | src/forms/Form_FormNCProyectoAC.cls | 339 | 34 | 7 | igual | pendiente | sin ficha todavía |
| Form_FormNCProyectoAR | formulario | src/forms/Form_FormNCProyectoAR.cls | 674 | 36 | 16 | modificado | pendiente | sin ficha todavía |
| Form_FormNCProyectoAcciones | formulario | src/forms/Form_FormNCProyectoAcciones.cls | 1458 | 32 | 30 | modificado | pendiente | sin ficha todavía |
| Form_FormNCProyectoControlEficacia | formulario | src/forms/Form_FormNCProyectoControlEficacia.cls | 449 | 23 | 13 | modificado | pendiente | sin ficha todavía |
| Form_FormNCProyectoControlEficaciaAlta | formulario | src/forms/Form_FormNCProyectoControlEficaciaAlta.cls | 243 | 23 | 9 | igual | pendiente | sin ficha todavía |
| Form_FormNCProyectoDocumentos | formulario | src/forms/Form_FormNCProyectoDocumentos.cls | 535 | 19 | 11 | modificado | pendiente | sin ficha todavía |
| Form_FormNCProyectoGeneral | formulario | src/forms/Form_FormNCProyectoGeneral.cls | 1229 | 58 | 29 | modificado | pendiente | sin ficha todavía |
| Form_FormNCProyectoGeneralConVinculoNC | formulario | src/forms/Form_FormNCProyectoGeneralConVinculoNC.cls | 454 | 37 | 13 | igual | pendiente | sin ficha todavía |
| Form_FormNCProyectoGestion | formulario | src/forms/Form_FormNCProyectoGestion.cls | 1827 | 73 | 56 | modificado | pendiente | sin ficha todavía |
| Form_FormNCProyectoMotivoEliminado | formulario | src/forms/Form_FormNCProyectoMotivoEliminado.cls | 224 | 13 | 7 | igual | pendiente | sin ficha todavía |
| Form_FormNCProyectoNota | formulario | src/forms/Form_FormNCProyectoNota.cls | 270 | 14 | 6 | igual | pendiente | sin ficha todavía |
| Form_FormNCProyectoReplanificaciones | formulario | src/forms/Form_FormNCProyectoReplanificaciones.cls | 181 | 18 | 4 | igual | pendiente | sin ficha todavía |
| Form_FormNCProyectoSeguimiento | formulario | src/forms/Form_FormNCProyectoSeguimiento.cls | 75 | 26 | 3 | modificado | pendiente | sin ficha todavía |
| Form_FormNCProyectoSeguimientoNC | formulario | src/forms/Form_FormNCProyectoSeguimientoNC.cls | 550 | 32 | 18 | modificado | pendiente | sin ficha todavía |
| Form_FormNCProyectoSeguimientoTareas | formulario | src/forms/Form_FormNCProyectoSeguimientoTareas.cls | 599 | 33 | 19 | modificado | pendiente | sin ficha todavía |
| Form_FormNCProyectoTipologiaGestion | formulario | src/forms/Form_FormNCProyectoTipologiaGestion.cls | 362 | 21 | 10 | igual | pendiente | sin ficha todavía |
| Form_FormTipologiaNCProyecto | formulario | src/forms/Form_FormTipologiaNCProyecto.cls | 245 | 19 | 6 | igual | pendiente | sin ficha todavía |
| Form_formRiesgosSeleccion | formulario | src/forms/Form_formRiesgosSeleccion.cls | 206 | 21 | 7 | igual | pendiente | sin ficha todavía |

Totales (formularios): 0 analizados · 0 no aplicables · 48 pendientes.
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

Pendiente. Las capacidades del [estudio previo](../capabilities.md) son contexto, no evidencia.

## Discrepancias abiertas

| Id | Discrepancia | Evidencia | Estado |
|---|---|---|---|
| D-NC-1 | El snapshot del estudio previo procede de la rama `staging` (248 commits por delante de `main`, último 2026-09-16), mientras que #1080 fija `main`. `main` carece de la capa de caché de indicadores y de listados y de los ayudantes `modFrm*Helper` que el estudio previo documenta como capacidad madura. | `git rev-list --count 1d7a2f8..origin/staging` = 248; `git rev-list --count origin/staging..1d7a2f8` = 9; recuentos de módulos | Abierta, **pendiente de decisión del propietario**: confirmar si el PRD debe fijar `main` (producción) o `staging`/una release candidate. Hasta entonces el PRD describe `main` y clasifica la caché como retirada frente al estudio previo. |
| D-NC-2 | #1080 cita como fecha del commit de partida 2026-10-21; el commit `1d7a2f8` es del 2026-09-15 (el 2026-10-21 es el día del recuento). | `git log -1 1d7a2f8` | Resuelta: se cita la fecha real. |
| D-NC-3 | El ERD de origen no se ha extraído: `dysflow` no está disponible en el entorno de trabajo y los backends `.accdb` no son legibles con herramientas de solo lectura. | véase `erd-origen.md` cuando llegue | Abierta, **pendiente del propietario** (runner Windows `dysflow-erd`). |

## Método del recuento

- El clon vive en `.atl/sources/nc-1d7a2f8/` (ruta ignorada, solo lectura) y está en `HEAD` desacoplado en el commit fijado.
- Controles: líneas `^\s*Begin ` del `.form.txt` anteriores al código embebido. Procedimientos: líneas `(Private |Public |Friend )?(Static )?(Sub|Function) ` del `.cls`.
- Comparación con el snapshot: normalización descrita arriba, sobre `.cls` y `.bas`. No se comparan los `.form.txt`.
- Recuento de CodeGraph: tabla `nodes` de `.codegraph/codegraph.db` del clon.
