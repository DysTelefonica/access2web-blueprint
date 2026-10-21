# PRD — Gestión de Riesgos

> **Estado:** fases 0, 1, 3, 4, 5 y 6 hechas. **4 de 63 formularios** tienen ficha; los **59 restantes quedan pendientes** y llegan en las tandas siguientes de la cadena. El ERD de origen queda pendiente del propietario (D-GR-4).
>
> **Qué hay en este PRD ahora mismo:** [matriz campo a campo](matriz-campos.md) de las 85 tablas · [coexistencia](coexistencia.md) · [capacidades](capacidades.md) · [issues propuestas](issues-propuestas.md) · 4 fichas en [formularios/](formularios/).

| Campo | Valor |
|---|---|
| Repositorio | `DysTelefonica/GESTION_RIESGOS` |
| Commit fijado | `36fcf60161fa5ec84c3956b1308537ddf2579256` (`main`, 2026-07-23, «fix(release): bump VersionAplicacion to 2026-007.5») |
| Índice CodeGraph-VBA | 196 ficheros indexados, 13 397 nodos, 23 837 aristas, versión 2.1.1, sobre clon de solo lectura de ese commit |
| Estudio previo | [epic.md](../epic.md) (DRAFT v0.1), [forms.md](../forms.md), [capabilities.md](../capabilities.md), [data-model.md](../data-model.md), [security-rules.md](../security-rules.md), [integrations-automation.md](../integrations-automation.md), [migration-matrix.md](../migration-matrix.md), [walkthroughs G1 a G5](../walkthrough-G1.json), [ERD/schema.sql](../ERD/schema.sql) |
| Snapshot comparado | `data/staging/gestion-riesgos/src` de `origin/main` (2026-08-06) |
| ERD propio del legacy | `ERD/Estructura_Datos.md` del commit fijado (2026-02-13): 81 tablas, 532 campos, 15 tablas vinculadas |
| Web existente | no hay módulo `app/src/modules/gestion_riesgos`; este PRD no parte de código web |
| Issue | #1079 (método y plantilla de #1069) |

## Alcance actual

Fija el commit, concilia el inventario contra el snapshot y deja los 63 formularios **pendientes** de ficha. No contiene fichas, matriz campo a campo, coexistencia, capacidades ni issues propuestas: llegan en las tandas siguientes. Los recuentos de líneas, controles y procedimientos son mecánicos (ver [Método del recuento](#método-del-recuento)) y no sustituyen al análisis.

## Inventario reconciliado

### Recuentos por tipo

| Tipo | Commit fijado `36fcf60` | Snapshot 2026-08-06 | Diferencia |
|---|---|---|---|
| Formularios (`src/forms/*.cls`, con su `.form.txt`) | 63 | 62 | +1 (2 añadidos, 1 retirado) |
| Módulos estándar (`src/modules/*.bas`) | 19 | 136 | −117 |
| Clases (`src/classes/*.cls`) | 47 | 65 | −18 |
| Consultas guardadas (`src/queries/`) | no existe la carpeta | no existe la carpeta | — |

La diferencia de módulos y clases **no es una retirada**: el snapshot contiene 112 módulos `Test_*` y ayudantes `mod*Helper` que el commit fijado no tiene, más 28 clases y módulos de aplicación. El anexo [cambios no liberados](#anexo-cambios-no-liberados) los enumera. La comparación se hizo con el snapshot de `origin/main`, no con el de la copia local, que está desalineada respecto de `origin`.

Sobre consultas guardadas: no hay carpeta `src/queries/` ni fichero `queries.json` en ninguno de los dos lados. Hay tres llamadas a `CreateQueryDef` con nombre vacío (consultas efímeras, no guardadas) en `src/forms/Form_formIndicadorProyectos.cls:243` y en el código embebido de sus dos `.form.txt`.

### Formularios

Líneas: `.cls` del formulario. Controles: bloques `Begin` del `.form.txt` anteriores al código embebido. Procedimientos: `Sub` y `Function` del `.cls`. «Frente al snapshot»: **igual** si el `.cls` coincide con el del snapshot tras quitar `Attribute VB_`, comentarios, blancos, finales de línea y diferencias de mayúsculas y espacios; **distinto** en caso contrario, sin prejuzgar si el cambio es cosmético o funcional.

| Objeto | Tipo | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha |
|---|---|---|---|---|---|---|---|---|
| Form_Form0BDOpciones | formulario | src/forms/Form_Form0BDOpciones.cls | 499 | 44 | 16 | distinto | pendiente | Form_Form0BDOpciones (ficha por llegar) |
| Form_Form0BDOpcionesTecnico | formulario | src/forms/Form_Form0BDOpcionesTecnico.cls | 314 | 36 | 10 | distinto | pendiente | Form_Form0BDOpcionesTecnico (ficha por llegar) |
| Form_FormAnexos | formulario | src/forms/Form_FormAnexos.cls | 492 | 33 | 12 | distinto | pendiente | Form_FormAnexos (ficha por llegar) |
| Form_FormAnexos1 | formulario | src/forms/Form_FormAnexos1.cls | 613 | 29 | 10 | distinto | pendiente | Form_FormAnexos1 (ficha por llegar) |
| Form_FormCalidadRiesgoAceptadoRetiradoVisado | formulario | src/forms/Form_FormCalidadRiesgoAceptadoRetiradoVisado.cls | 348 | 29 | 5 | distinto | pendiente | Form_FormCalidadRiesgoAceptadoRetiradoVisado (ficha por llegar) |
| Form_FormCalidadRiesgoMaterializaciones | formulario | src/forms/Form_FormCalidadRiesgoMaterializaciones.cls | 716 | 27 | 15 | distinto | pendiente | Form_FormCalidadRiesgoMaterializaciones (ficha por llegar) |
| Form_FormCalidadTareaExplicacion | formulario | src/forms/Form_FormCalidadTareaExplicacion.cls | 69 | 13 | 1 | distinto | pendiente | Form_FormCalidadTareaExplicacion (ficha por llegar) |
| Form_FormCalidadTareaRiesgosAceptadosRetirados | formulario | src/forms/Form_FormCalidadTareaRiesgosAceptadosRetirados.cls | 391 | 45 | 8 | distinto | pendiente | Form_FormCalidadTareaRiesgosAceptadosRetirados (ficha por llegar) |
| Form_FormCalidadTareaRiesgosMaterializadosPorDecidir | formulario | src/forms/Form_FormCalidadTareaRiesgosMaterializadosPorDecidir.cls | 202 | 32 | 4 | distinto | pendiente | Form_FormCalidadTareaRiesgosMaterializadosPorDecidir (ficha por llegar) |
| Form_FormCalidadTareaRiesgosRetipificacion | formulario | src/forms/Form_FormCalidadTareaRiesgosRetipificacion.cls | 198 | 36 | 4 | distinto | pendiente | Form_FormCalidadTareaRiesgosRetipificacion (ficha por llegar) |
| Form_FormCalidadTareas | formulario | src/forms/Form_FormCalidadTareas.cls | 542 | 21 | 15 | distinto | pendiente | Form_FormCalidadTareas (ficha por llegar) |
| Form_FormCalidadTareasDetalleEdicion | formulario | src/forms/Form_FormCalidadTareasDetalleEdicion.cls | 361 | 52 | 8 | distinto | pendiente | Form_FormCalidadTareasDetalleEdicion (ficha por llegar) |
| Form_FormControlCambiosGestion | formulario | src/forms/Form_FormControlCambiosGestion.cls | 454 | 31 | 14 | distinto | pendiente | Form_FormControlCambiosGestion (ficha por llegar) |
| Form_FormExpedientesBusqueda | formulario | src/forms/Form_FormExpedientesBusqueda.cls | 421 | 31 | 12 | distinto | pendiente | Form_FormExpedientesBusqueda (ficha por llegar) |
| Form_FormGestionRiesgos | formulario | src/forms/Form_FormGestionRiesgos.cls | 422 | 30 | 8 | distinto | pendiente | Form_FormGestionRiesgos (ficha por llegar) |
| Form_FormGestionRiesgosAutorizados | formulario | src/forms/Form_FormGestionRiesgosAutorizados.cls | 92 | 12 | 2 | distinto | pendiente | Form_FormGestionRiesgosAutorizados (ficha por llegar) |
| Form_FormGestionRiesgosDatosGenerales | formulario | src/forms/Form_FormGestionRiesgosDatosGenerales.cls | 645 | 47 | 16 | distinto | pendiente | Form_FormGestionRiesgosDatosGenerales (ficha por llegar) |
| Form_FormGestionRiesgosRiesgosOferta | formulario | src/forms/Form_FormGestionRiesgosRiesgosOferta.cls | 738 | 29 | 17 | distinto | pendiente | Form_FormGestionRiesgosRiesgosOferta (ficha por llegar) |
| Form_FormGestionRiesgosSuministradores | formulario | src/forms/Form_FormGestionRiesgosSuministradores.cls | 552 | 22 | 13 | distinto | pendiente | Form_FormGestionRiesgosSuministradores (ficha por llegar) |
| Form_FormIndicador | formulario | src/forms/Form_FormIndicador.cls | 316 | 50 | 17 | igual | pendiente | Form_FormIndicador (ficha por llegar) |
| Form_formIndicadorProyectos | formulario | src/forms/Form_formIndicadorProyectos.cls | 277 | 25 | 11 | distinto | pendiente | Form_formIndicadorProyectos (ficha por llegar) |
| Form_FormInformeTipoSalida | formulario | src/forms/Form_FormInformeTipoSalida.cls | 102 | 19 | 3 | anadido | analizado | [formularios/Form_FormInformeTipoSalida.md](formularios/Form_FormInformeTipoSalida.md) |
| Form_FormPlanAcciones | formulario | src/forms/Form_FormPlanAcciones.cls | 378 | 36 | 18 | distinto | pendiente | Form_FormPlanAcciones (ficha por llegar) |
| Form_FormPlanPrincipal | formulario | src/forms/Form_FormPlanPrincipal.cls | 196 | 27 | 6 | distinto | pendiente | Form_FormPlanPrincipal (ficha por llegar) |
| Form_FormProyectosGestion | formulario | src/forms/Form_FormProyectosGestion.cls | 955 | 43 | 29 | distinto | pendiente | Form_FormProyectosGestion (ficha por llegar) |
| Form_FormPublicacionCalidad | formulario | src/forms/Form_FormPublicacionCalidad.cls | 147 | 28 | 4 | distinto | pendiente | Form_FormPublicacionCalidad (ficha por llegar) |
| Form_FormPublicacionCalidadNotas | formulario | src/forms/Form_FormPublicacionCalidadNotas.cls | 116 | 17 | 3 | distinto | pendiente | Form_FormPublicacionCalidadNotas (ficha por llegar) |
| Form_FormPublicacionCalidadPublicar | formulario | src/forms/Form_FormPublicacionCalidadPublicar.cls | 472 | 28 | 11 | distinto | pendiente | Form_FormPublicacionCalidadPublicar (ficha por llegar) |
| Form_FormPublicacionCalidadPublicarEjecutar | formulario | src/forms/Form_FormPublicacionCalidadPublicarEjecutar.cls | 321 | 33 | 7 | distinto | pendiente | Form_FormPublicacionCalidadPublicarEjecutar (ficha por llegar) |
| Form_FormPublicacionCalidadRechazoObservaciones | formulario | src/forms/Form_FormPublicacionCalidadRechazoObservaciones.cls | 194 | 25 | 5 | distinto | pendiente | Form_FormPublicacionCalidadRechazoObservaciones (ficha por llegar) |
| Form_FormPublicacionDatosGenerales | formulario | src/forms/Form_FormPublicacionDatosGenerales.cls | 81 | 28 | 2 | distinto | pendiente | Form_FormPublicacionDatosGenerales (ficha por llegar) |
| Form_FormPublicacionSuministradores | formulario | src/forms/Form_FormPublicacionSuministradores.cls | 378 | 19 | 7 | distinto | pendiente | Form_FormPublicacionSuministradores (ficha por llegar) |
| Form_FormPublicacionTecnico | formulario | src/forms/Form_FormPublicacionTecnico.cls | 142 | 26 | 4 | distinto | pendiente | Form_FormPublicacionTecnico (ficha por llegar) |
| Form_FormPublicacionTecnicoPropuesta | formulario | src/forms/Form_FormPublicacionTecnicoPropuesta.cls | 421 | 26 | 7 | distinto | pendiente | Form_FormPublicacionTecnicoPropuesta (ficha por llegar) |
| Form_FormPublicacionUTE | formulario | src/forms/Form_FormPublicacionUTE.cls | 225 | 18 | 5 | distinto | pendiente | Form_FormPublicacionUTE (ficha por llegar) |
| Form_FormRiesgo | formulario | src/forms/Form_FormRiesgo.cls | 576 | 32 | 14 | distinto | pendiente | Form_FormRiesgo (ficha por llegar) |
| Form_FormRiesgoBiblioteca | formulario | src/forms/Form_FormRiesgoBiblioteca.cls | 299 | 25 | 8 | distinto | pendiente | Form_FormRiesgoBiblioteca (ficha por llegar) |
| Form_FormRiesgoDefinicion | formulario | src/forms/Form_FormRiesgoDefinicion.cls | 702 | 33 | 17 | distinto | pendiente | Form_FormRiesgoDefinicion (ficha por llegar) |
| Form_FormRiesgoDefinicionNoBiblioteca | formulario | src/forms/Form_FormRiesgoDefinicionNoBiblioteca.cls | 317 | 23 | 9 | distinto | pendiente | Form_FormRiesgoDefinicionNoBiblioteca (ficha por llegar) |
| Form_FormRiesgoExternoDetalle | formulario | src/forms/Form_FormRiesgoExternoDetalle.cls | 416 | 34 | 11 | distinto | pendiente | Form_FormRiesgoExternoDetalle (ficha por llegar) |
| Form_FormRiesgoMaterializado | formulario | src/forms/Form_FormRiesgoMaterializado.cls | 302 | 19 | 6 | distinto | pendiente | Form_FormRiesgoMaterializado (ficha por llegar) |
| Form_FormRiesgoMitigacion | formulario | src/forms/Form_FormRiesgoMitigacion.cls | 594 | 25 | 12 | distinto | pendiente | Form_FormRiesgoMitigacion (ficha por llegar) |
| Form_FormRiesgoNC | formulario | src/forms/Form_FormRiesgoNC.cls | 469 | 50 | 10 | distinto | pendiente | Form_FormRiesgoNC (ficha por llegar) |
| Form_FormRiesgoPlazoCosteCalidad | formulario | src/forms/Form_FormRiesgoPlazoCosteCalidad.cls | 230 | 19 | 7 | distinto | pendiente | Form_FormRiesgoPlazoCosteCalidad (ficha por llegar) |
| Form_FormRiesgoPosiblesModificacionesAceptarRetirar | formulario | src/forms/Form_FormRiesgoPosiblesModificacionesAceptarRetirar.cls | 227 | 27 | 4 | distinto | pendiente | Form_FormRiesgoPosiblesModificacionesAceptarRetirar (ficha por llegar) |
| Form_FormRiesgoPosiblesModificacionesQuitarAceptar | formulario | src/forms/Form_FormRiesgoPosiblesModificacionesQuitarAceptar.cls | 181 | 20 | 6 | distinto | pendiente | Form_FormRiesgoPosiblesModificacionesQuitarAceptar (ficha por llegar) |
| Form_FormRiesgoRetirado | formulario | src/forms/Form_FormRiesgoRetirado.cls | 571 | 27 | 11 | distinto | pendiente | Form_FormRiesgoRetirado (ficha por llegar) |
| Form_FormRiesgosBibliotecaGestion | formulario | src/forms/Form_FormRiesgosBibliotecaGestion.cls | 822 | 37 | 22 | distinto | pendiente | Form_FormRiesgosBibliotecaGestion (ficha por llegar) |
| Form_FormRiesgosEstablecerPrioridades | formulario | src/forms/Form_FormRiesgosEstablecerPrioridades.cls | 667 | 26 | 17 | distinto | pendiente | Form_FormRiesgosEstablecerPrioridades (ficha por llegar) |
| Form_FormRiesgosGestion | formulario | src/forms/Form_FormRiesgosGestion.cls | 1783 | 38 | 30 | distinto | pendiente | Form_FormRiesgosGestion (ficha por llegar) |
| Form_FormRiesgosGestionEdicion | formulario | src/forms/Form_FormRiesgosGestionEdicion.cls | 552 | 41 | 10 | distinto | pendiente | Form_FormRiesgosGestionEdicion (ficha por llegar) |
| Form_FormRiesgosGestionPlanAcciones | formulario | src/forms/Form_FormRiesgosGestionPlanAcciones.cls | 563 | 25 | 16 | distinto | pendiente | Form_FormRiesgosGestionPlanAcciones (ficha por llegar) |
| Form_FormRiesgosGestionPlanPrincipal | formulario | src/forms/Form_FormRiesgosGestionPlanPrincipal.cls | 400 | 22 | 9 | distinto | pendiente | Form_FormRiesgosGestionPlanPrincipal (ficha por llegar) |
| Form_FormRiesgosGestionRiesgo | formulario | src/forms/Form_FormRiesgosGestionRiesgo.cls | 371 | 47 | 9 | distinto | pendiente | Form_FormRiesgosGestionRiesgo (ficha por llegar) |
| Form_FormRiesgoVulnerabilidad | formulario | src/forms/Form_FormRiesgoVulnerabilidad.cls | 98 | 13 | 3 | distinto | pendiente | Form_FormRiesgoVulnerabilidad (ficha por llegar) |
| Form_FormSuministrador | formulario | src/forms/Form_FormSuministrador.cls | 111 | 28 | 4 | distinto | pendiente | Form_FormSuministrador (ficha por llegar) |
| Form_FormTecnicoTareaExplicacion | formulario | src/forms/Form_FormTecnicoTareaExplicacion.cls | 68 | 13 | 1 | distinto | pendiente | Form_FormTecnicoTareaExplicacion (ficha por llegar) |
| Form_FormTecnicoTareaRiesgosAceptadosRetirados | formulario | src/forms/Form_FormTecnicoTareaRiesgosAceptadosRetirados.cls | 194 | 44 | 3 | distinto | pendiente | Form_FormTecnicoTareaRiesgosAceptadosRetirados (ficha por llegar) |
| Form_FormTecnicoTareas | formulario | src/forms/Form_FormTecnicoTareas.cls | 387 | 19 | 12 | distinto | pendiente | Form_FormTecnicoTareas (ficha por llegar) |
| Form_FormTecnicoTareasDetalleEdicion | formulario | src/forms/Form_FormTecnicoTareasDetalleEdicion.cls | 223 | 43 | 5 | distinto | pendiente | Form_FormTecnicoTareasDetalleEdicion (ficha por llegar) |
| Form_Formulario1 | formulario | src/forms/Form_Formulario1.cls | 15 | 11 | 1 | anadido | analizado | [formularios/Form_Formulario1.md](formularios/Form_Formulario1.md) |
| Form_FormWeb | formulario | src/forms/Form_FormWeb.cls | 68 | 17 | 2 | distinto | analizado | [formularios/Form_FormWeb.md](formularios/Form_FormWeb.md) |
| Form_frmSplash | formulario | src/forms/Form_frmSplash.cls | 102 | 16 | 3 | distinto | analizado | [formularios/Form_frmSplash.md](formularios/Form_frmSplash.md) |

El relevo es una hipótesis por nombre, no un hecho: no se confirma sin leer el código de las fases siguientes.

### Módulos y clases

19 módulos y 47 clases en el commit fijado. Módulos: `Constantes`, `Constructor`, `ControlCambios`, `ExcelInforme`, `Funciones Generales`, `IndicadorRiesgosRepositorio`, `InformeRiesgoHTML`, `Instalacion`, `JsonConverter`, `Módulo1`, `PublicabilidadEdicion`, `PublicabilidadRiesgo`, `Variables Globales`, `mIndicador`, `mIndicadorProyectosState`, `mdlCursor`, `modIndicadorDashboard`, `modIndicadorOpenArgsParser`, `modMigracionAux`, `testing`. No hay módulos `Test_*` en `main`.

Clases: 47, entre ellas el agregado de dominio (`Proyecto`, `Edicion`, `Riesgo`, `RiesgoBiblioteca`, `RiesgoExterno`, `RiesgoMaterializacion`, `MitigacionValor`, `AreaImpacto`, `RAC`, `Pedido`, `Juridica`, `NC`, `OrganoContratacion`, `Suministrador`, `Usuario`, `Entorno`, `Correo`, `HTML`, `Documento`, `Anexo`, `AnexoAntiguo`), los planes (`PM`, `PC`, `PMAccion`, `PCAccion`, `PMAccionReversa`, `PCAccionReversa`), el control de cambios (`CCVersion`, `CCCambio`, `CCDocumentoCambio`), las tareas (`TareasCalidad`, `TareasTecnico`, `ArbolTareasCalidad`, `ArbolTareasTecnico`), las evidencias (`SuministradorParaEvidencias`) y los servicios `InformeRiesgoPDFServicio`, `ProyectoSuministrador` y `EdicionSuministrador`. El inventario completo se reconcilia en `capacidades.md`.

## Delta frente al estudio previo

En el código de `main` no hay ningún formulario retirado de lo que describe el estudio previo: los 63 formularios cubren las mismas áreas que [forms.md](../forms.md). Lo que cambia es la precisión de los recuentos que el estudio previo da como contexto, no el conjunto de objetos: ver [Discrepancias abiertas](#discrepancias-abiertas). La clasificación completa está en [capacidades.md](capacidades.md): 14 sin cambios, 9 modificadas, 1 añadida y **0 retiradas**.

## Capacidades

El [estudio previo](../capabilities.md) enumera las capacidades observadas. La clasificación de cada una frente al commit fijado (añadida, modificada, retirada o sin cambios, con evidencia) llega en `capacidades.md`.

## Discrepancias abiertas

| Id | Discrepancia | Evidencia | Estado |
|---|---|---|---|
| D-GR-1 | #1079 fecha el commit de partida el 2026-10-21; el commit `36fcf60` es del 2026-07-23. | `git log -1 36fcf60` | Resuelta: el 2026-10-21 es el día del recuento. `origin/main` del legacy está en ese mismo commit, así que el pin sigue vigente. |
| D-GR-2 | El recuento de llamadas a `getdb()` difiere entre fuentes: 308 en [capabilities.md](../capabilities.md), 372 en [epic.md](../epic.md) (D88) y 362 medidos aquí en ficheros de producción. | Medición sobre el commit fijado, excluyendo `Test_*`: 196 en `src/modules` (12 ficheros), 163 en `src/classes` (27), 3 en `src/forms` (3) | Abierta: la diferencia es de alcance (se incluye o no la definición de `getdb`, los módulos de prueba y las llamadas con otro nombre). El dato de este PRD es el medido, con su alcance declarado. |
| D-GR-3 | El snapshot tiene 136 módulos y 65 clases; el commit fijado tiene 19 y 47. | Anexo [cambios no liberados](#anexo-cambios-no-liberados) | Explicada en lo mecánico: el snapshot es una rama de trabajo posterior con 112 módulos `Test_*` y ayudantes más 28 objetos de aplicación. Ninguno está retirado en `main`; quedan sin analizar mientras no se liberen. |
| D-GR-4 | El ERD de origen no se ha extraído con `dysflow`: el backend `Gestion_Riesgos_Datos.accdb` no está en el clon (el clon trae el frontend, `Gestion_Riesgos.accdb`) y `mdb-tables` rechaza el fichero que sí está. Las dos fuentes disponibles (`ERD/Estructura_Datos.md` y [ERD/schema.sql](../ERD/schema.sql)) no tienen claves primarias, índices, relaciones ni nulabilidad. | `ERD/Estructura_Datos.md` (2026-02-13); `mdb-tables Gestion_Riesgos.accdb` | Abierta, **pendiente del propietario**: extraer en su máquina Windows, con `dysflow`, el backend `Gestion_Riesgos_Datos.accdb` y en solo lectura. Bloqueada además por la vía de extracción: el workflow `dysflow-erd` está en el PR #1144, **abierto y no mergeado**, así que hoy no hay forma de dispararlo desde el repositorio. |
| D-GR-5 | `ERD/Estructura_Datos.md` del legacy declara `TbProyectos1` como tabla **vinculada** a `Gestion_Riesgos_Datos.accdb`, que es el propio backend de la aplicación. | `ERD/Estructura_Datos.md` (sección `TbProyectos1`) | Abierta: o el documento está desactualizado (es de 2026-02-13 y el commit fijado es de julio) o hay un vínculo a sí mismo que no puede ser. Verificar contra el binario antes de usar esa tabla como evidencia de vínculo. |

## Hallazgos del legacy

18 Hallazgos del legacy (H-GR-1 a H-GR-18), todos «a verificar». Ninguno se corrige aquí.

| Rango | Ficha | Qué agrupa |
|---|---|---|
| H-GR-1 a H-GR-8 | [Form_frmSplash](formularios/Form_frmSplash.md) | El arranque: error que cierra sin abrir menú, `Err.Raise` sin descripción, salto `/dev` en producción |
| H-GR-9 a H-GR-14 | [Form_FormWeb](formularios/Form_FormWeb.md) | El visor: cierre en silencio si el fichero no existe, apertura en segundo plano, ruta de una máquina concreta en un comentario |
| — | [Form_FormInformeTipoSalida](formularios/Form_FormInformeTipoSalida.md) | Opción deshabilitada que sigue activa, dos defaults distintos, `ShellExecuteA` sobre una URL generada |
| — | [Form_Formulario1](formularios/Form_Formulario1.md) | Huérfano de otra aplicación, con botones que no hacen nada |

El listado completo, con la evidencia `fichero:línea` de cada uno, está en las fichas. Los que condicionan el diseño de la migración —concurrencia sin transacciones, `Max + 1`, comportamiento de error silenciado— se tratan en [coexistencia](coexistencia.md) y en [issues-propuestas](issues-propuestas.md), no aquí.

## Anexo: cambios no liberados

Objetos que están en el snapshot `data/staging/gestion-riesgos/src` y **no** en el commit fijado. No forman parte de la línea base de este PRD: no se analizan ni se inventarían aquí, pero se dejan registrados para que la migración no los pierda si se liberan.

**Módulos y ayudantes de prueba y soporte (112):** `Test_*` y `mod*Helper` del snapshot. No son functionality de producción y no se migran.

**Objetos de aplicación ausentes en `main` (28):**

`CalidadPublicacionResultado`, `CalidadPublicacionService`, `GestionRiesgosChangeDetector`, `GestionRiesgosEdicionDeleteDecisionService`, `GestionRiesgosSaveDecisionService`, `GestionRiesgosViewState`, `GestionRiesgosViewStateBuilder`, `MaterializacionPlanEligibilityService`, `PublicacionCalidadExecutionService`, `PublicacionCalidadStatusPresenter`, `PublicacionCalidadStatusViewState`, `RiesgoChangeDetector`, `RiesgoEstadoCalidadViewState`, `RiesgoEstadoCalidadViewStateBuilder`, `RiesgoPlanActionChangeDetector`, `RiesgoPriorizacionService`, `RiskTreeDataLoader`, `RiskTreeRefreshService`, `RiskTreeSelectionResolver`, `SuministradorParaEvidencias` (clases); `MigracionCorreosEnviadosIDEdicionLong`, `MigracionMaterializacionPlanContingencia`, `MigracionRiesgosPriorizacionLong`, `MigracionTbCambiosParaPublicacionEdicionLong`, `PriorizacionRepository`, `RiesgoMaterializacionContingenciaService`, `SearchComboCache`, `SuministradoresHelper` (módulos).

Varios de ellos (`PriorizacionRepository`, `RiesgoMaterializacionContingenciaService`, `MigracionMaterializacionPlanContingencia`, `MaterializacionPlanEligibilityService`) son precisamente la materialización de la contingency que `Form_FormMaterializacionPlanContingencia` daba por Cubierta. Si se liberan, hay que reabrir el análisis de los formularios de plan.

**Objetos en `main` que no están en el snapshot (3):** `EdicionSuministrador` (clase), `ExcelInforme` (módulo), `ProyectoSuministrador` (clase). Están en producción y no llegaron al snapshot.

## Método del recuento

- El clon vive en `~/agent-wt/GESTION_RIESGOS/gr-1079-src/` (ruta ignorada, solo lectura) y está en `HEAD` desacoplado en el commit fijado.
- Controles: líneas `^\s*Begin ` del `.form.txt` anteriores al bloque `Begin VB.Module` embebido. Procedimientos: líneas `^\s*(Private |Public )?(Sub|Function) \w+` del `.cls`.
- Comparación con el snapshot: `.cls` con CR eliminados, `Attribute VB_*` fuera, comentarios, blancos, espacios y mayúsculas normalizados. No se comparan los `.form.txt`.
- Recuento de CodeGraph: salida de `codegraph init -i` (2.1.1) sobre el clon.
- Las llamadas a `getdb()`, las transacciones y los `On Error Resume Next` se cuentan solo en ficheros de producción (excluidos los módulos `Test_*`).
