# PRD — Expedientes

> **Estado:** en curso · fases 0 (fijación) y 1 (inventario) del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) cerradas; fase 2 (formularios) iniciada con 13 de 45.

| Campo | Valor |
|---|---|
| Repositorio | DysTelefonica/Expedientes |
| Commit fijado | `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` (`release/2026-008.2.rc1`, 2026-10-20, «fix(cache): hard-cutover p_db obligatorio en RealinearCacheExpedientes y RealinearTodaLaCache (#244)») |
| Índice CodeGraph-VBA | 233 ficheros indexados, 11 784 nodos, versión 2.1.1, sobre clon de solo lectura de ese commit |
| Estudio previo | [epic.md](../epic.md) (DRAFT), [forms.md](../forms.md), [walkthroughs G1 a G5](../walkthrough-G1.json), [capabilities/](../capabilities/), [data-model.md](../data-model.md), [ERD/schema.sql](../ERD/schema.sql), [migration-matrix.md](../migration-matrix.md) |
| Snapshot comparado | `data/staging/expedientes/src` (2026-08-06) |
| Issue | #1073 (método y plantilla de #1069) |

## Alcance de esta tanda

Fija el commit, concilia el inventario de objetos y deja todos los formularios como **pendientes**. No contiene fichas, ERD, matriz ni coexistencia: llegan en las tandas siguientes de la cadena. Los recuentos de líneas, controles y procedimientos son mecánicos (ver [Método del recuento](#método-del-recuento)) y no sustituyen al análisis.

## Inventario reconciliado

### Recuentos por tipo

| Tipo | RC1 `ec6a8c3` | Snapshot 2026-08-06 | Diferencia |
|---|---|---|---|
| Formularios (`src/forms/*.cls`, con su `.form.txt`) | 45 | 46 | −1 (3 añadidos, 4 retirados) |
| Módulos estándar (`src/modules/*.bas`) | 82 | 143 | −61 |
| Clases (`src/classes/*.cls`) | 60 | 66 | −6 (1 añadida, 7 retiradas) |
| Consultas guardadas (`src/queries/`) | 1 (`Consulta1.sql`) y `queries.json` | sin recuento | pendiente |

### Formularios

Líneas: `.cls` del formulario. Controles: bloques `Begin` del `.form.txt`. Procedimientos: `Sub` y `Function` del `.cls`. «Frente al snapshot»: **añadido** si no existe en el snapshot; **distinto** si existe y el `.cls` difiere (ignorando CRLF); la clasificación semántica (modificada o sin cambios) se hace en la ficha.

| Objeto | Tipo | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|---|
| Form_Form0BDGestorEntidades | formulario | src/forms/Form_Form0BDGestorEntidades.cls | 187 | 14 | 7 | distinto | pendiente | sin ficha todavía |
| Form_Form0BDOpciones | formulario | src/forms/Form_Form0BDOpciones.cls | 356 | 26 | 14 | distinto | pendiente | sin ficha todavía |
| Form_Form0BDOpcionesTecnicos | formulario | src/forms/Form_Form0BDOpcionesTecnicos.cls | 166 | 20 | 7 | distinto | pendiente | sin ficha todavía |
| Form_FormCPV | formulario | src/forms/Form_FormCPV.cls | 199 | 23 | 6 | distinto | pendiente | sin ficha todavía |
| Form_FormCPVsGestion | formulario | src/forms/Form_FormCPVsGestion.cls | 416 | 23 | 14 | distinto | pendiente | sin ficha todavía |
| Form_FormComercial | formulario | src/forms/Form_FormComercial.cls | 199 | 23 | 6 | distinto | pendiente | sin ficha todavía |
| Form_FormComercialesGestion | formulario | src/forms/Form_FormComercialesGestion.cls | 415 | 23 | 14 | distinto | pendiente | sin ficha todavía |
| Form_FormE2ECola | formulario | src/forms/Form_FormE2ECola.cls | 336 | 45 | 13 | **añadido** | pendiente | sin ficha todavía |
| Form_FormE2EGestionEnlaces | formulario | src/forms/Form_FormE2EGestionEnlaces.cls | 261 | 26 | 11 | **añadido** | pendiente | sin ficha todavía |
| Form_FormEjercito | formulario | src/forms/Form_FormEjercito.cls | 199 | 23 | 6 | distinto | pendiente | sin ficha todavía |
| Form_FormEjercitosGestion | formulario | src/forms/Form_FormEjercitosGestion.cls | 414 | 23 | 14 | distinto | pendiente | sin ficha todavía |
| Form_FormEleccionTipoConsulta | formulario | src/forms/Form_FormEleccionTipoConsulta.cls | 390 | 16 | 10 | distinto | pendiente | sin ficha todavía |
| Form_FormExpediente | formulario | src/forms/Form_FormExpediente.cls | 597 | 27 | 13 | distinto | analizado | [formularios/Form_FormExpediente.md](formularios/Form_FormExpediente.md) |
| Form_FormExpedienteAltaParaHPS | formulario | src/forms/Form_FormExpedienteAltaParaHPS.cls | 453 | 47 | 23 | distinto | analizado | [formularios/Form_FormExpedienteAltaParaHPS.md](formularios/Form_FormExpedienteAltaParaHPS.md) |
| Form_FormExpedienteAltaTipo | formulario | src/forms/Form_FormExpedienteAltaTipo.cls | 363 | 14 | 8 | distinto | analizado | [formularios/Form_FormExpedienteAltaTipo.md](formularios/Form_FormExpedienteAltaTipo.md) |
| Form_FormExpedienteCambioTipo | formulario | src/forms/Form_FormExpedienteCambioTipo.cls | 336 | 21 | 8 | distinto | analizado | [formularios/Form_FormExpedienteCambioTipo.md](formularios/Form_FormExpedienteCambioTipo.md) |
| Form_FormExpedienteDocumentacion | formulario | src/forms/Form_FormExpedienteDocumentacion.cls | 329 | 14 | 8 | distinto | analizado | [formularios/Form_FormExpedienteDocumentacion.md](formularios/Form_FormExpedienteDocumentacion.md) |
| Form_FormExpedienteEntidades | formulario | src/forms/Form_FormExpedienteEntidades.cls | 1974 | 61 | 49 | distinto | analizado | [formularios/Form_FormExpedienteEntidades.md](formularios/Form_FormExpedienteEntidades.md) |
| Form_FormExpedienteFechas | formulario | src/forms/Form_FormExpedienteFechas.cls | 327 | 44 | 15 | distinto | analizado | [formularios/Form_FormExpedienteFechas.md](formularios/Form_FormExpedienteFechas.md) |
| Form_FormExpedienteGeneral | formulario | src/forms/Form_FormExpedienteGeneral.cls | 905 | 73 | 19 | distinto | analizado | [formularios/Form_FormExpedienteGeneral.md](formularios/Form_FormExpedienteGeneral.md) |
| Form_FormExpedienteHitos | formulario | src/forms/Form_FormExpedienteHitos.cls | 296 | 21 | 6 | distinto | analizado | [formularios/Form_FormExpedienteHitos.md](formularios/Form_FormExpedienteHitos.md) |
| Form_FormExpedienteModificados | formulario | src/forms/Form_FormExpedienteModificados.cls | 344 | 22 | 8 | distinto | analizado | [formularios/Form_FormExpedienteModificados.md](formularios/Form_FormExpedienteModificados.md) |
| Form_FormExpedienteSuministradores | formulario | src/forms/Form_FormExpedienteSuministradores.cls | 961 | 25 | 38 | distinto | analizado | [formularios/Form_FormExpedienteSuministradores.md](formularios/Form_FormExpedienteSuministradores.md) |
| Form_FormExpedientesGestion | formulario | src/forms/Form_FormExpedientesGestion.cls | 3963 | 124 | 99 | distinto | pendiente | sin ficha todavía |
| Form_FormExpedientesGestionTecnica | formulario | src/forms/Form_FormExpedientesGestionTecnica.cls | 773 | 43 | 24 | distinto | pendiente | sin ficha todavía |
| Form_FormExpedientesParaCambioTipo | formulario | src/forms/Form_FormExpedientesParaCambioTipo.cls | 321 | 22 | 10 | distinto | analizado | [formularios/Form_FormExpedientesParaCambioTipo.md](formularios/Form_FormExpedientesParaCambioTipo.md) |
| Form_FormExpedientesResponsables | formulario | src/forms/Form_FormExpedientesResponsables.cls | 187 | 25 | 4 | **añadido** | analizado | [formularios/Form_FormExpedientesResponsables.md](formularios/Form_FormExpedientesResponsables.md) |
| Form_FormGradoClasificacion | formulario | src/forms/Form_FormGradoClasificacion.cls | 197 | 23 | 6 | distinto | pendiente | sin ficha todavía |
| Form_FormGradosClasificacionGestion | formulario | src/forms/Form_FormGradosClasificacionGestion.cls | 425 | 23 | 14 | distinto | pendiente | sin ficha todavía |
| Form_FormLugarEjecucion | formulario | src/forms/Form_FormLugarEjecucion.cls | 197 | 23 | 6 | distinto | pendiente | sin ficha todavía |
| Form_FormLugarEjecucionGestion | formulario | src/forms/Form_FormLugarEjecucionGestion.cls | 416 | 23 | 14 | distinto | pendiente | sin ficha todavía |
| Form_FormModificado | formulario | src/forms/Form_FormModificado.cls | 211 | 27 | 6 | distinto | pendiente | sin ficha todavía |
| Form_FormOficinaPrograma | formulario | src/forms/Form_FormOficinaPrograma.cls | 197 | 23 | 6 | distinto | pendiente | sin ficha todavía |
| Form_FormOficinasProgramaGestion | formulario | src/forms/Form_FormOficinasProgramaGestion.cls | 425 | 23 | 14 | distinto | pendiente | sin ficha todavía |
| Form_FormOrganoContratacion | formulario | src/forms/Form_FormOrganoContratacion.cls | 205 | 23 | 6 | distinto | pendiente | sin ficha todavía |
| Form_FormOrganoContratacionGestion | formulario | src/forms/Form_FormOrganoContratacionGestion.cls | 427 | 23 | 14 | distinto | pendiente | sin ficha todavía |
| Form_FormPECAL | formulario | src/forms/Form_FormPECAL.cls | 197 | 23 | 6 | distinto | pendiente | sin ficha todavía |
| Form_FormPECALESGestion | formulario | src/forms/Form_FormPECALESGestion.cls | 424 | 23 | 14 | distinto | pendiente | sin ficha todavía |
| Form_FormRAC | formulario | src/forms/Form_FormRAC.cls | 205 | 25 | 6 | distinto | pendiente | sin ficha todavía |
| Form_FormRACSGestion | formulario | src/forms/Form_FormRACSGestion.cls | 414 | 23 | 14 | distinto | pendiente | sin ficha todavía |
| Form_FormSuministrador | formulario | src/forms/Form_FormSuministrador.cls | 236 | 33 | 6 | distinto | pendiente | sin ficha todavía |
| Form_FormSuministradoresGestion | formulario | src/forms/Form_FormSuministradoresGestion.cls | 458 | 24 | 15 | distinto | pendiente | sin ficha todavía |
| Form_FormTareas | formulario | src/forms/Form_FormTareas.cls | 470 | 24 | 10 | distinto | pendiente | sin ficha todavía |
| Form_Formulario1 | formulario | src/forms/Form_Formulario1.cls | 89 | 9 | 5 | distinto | pendiente | sin ficha todavía |
| Form_frmSplash | formulario | src/forms/Form_frmSplash.cls | 63 | 13 | 2 | distinto | pendiente | sin ficha todavía |

Totales (formularios): 13 analizados · 0 no aplicables · 32 pendientes.
Totales (controles `Begin`): 1 276 · procedimientos: 618 · líneas de `.cls`: 20 923.

### Formularios retirados respecto al snapshot

Existen en el snapshot y no en RC1. El estudio debe averiguar si desaparecen, se renombran o se absorben en otro formulario; hasta entonces su destino es **pendiente**.

| Objeto | Ruta en el snapshot | Posible relevo (a verificar) |
|---|---|---|
| Form_FormE2EGestionBatch | `data/staging/expedientes/src/forms/Form_FormE2EGestionBatch.cls` | `Form_FormE2EGestionEnlaces` / `Form_FormE2ECola` (pendiente) |
| Form_FormExpedienteAlta | `data/staging/expedientes/src/forms/Form_FormExpedienteAlta.cls` | `Form_FormExpedienteAltaTipo` (pendiente) |
| Form_FormUsuariosGestion | `data/staging/expedientes/src/forms/Form_FormUsuariosGestion.cls` | pendiente |
| Form_frmBusy | `data/staging/expedientes/src/forms/Form_frmBusy.cls` | pendiente |

Los relevos son hipótesis por nombre, no hechos: no se confirman sin leer el código.

### Módulos y clases

La mayoría de las diferencias son módulos `Test_*` y ayudantes `mod*Helper` del snapshot que RC1 no tiene. El estudio los clasifica en la fase 2 al analizar el formulario que los usa. Hasta entonces, el inventario de módulos y clases queda **pendiente** y no se enumera objeto a objeto en esta tanda.

## Delta frente al estudio previo

Todo **pendiente**. Solo se afirma lo mecánico: 3 formularios añadidos (`Form_FormE2ECola`, `Form_FormE2EGestionEnlaces`, `Form_FormExpedientesResponsables`), 4 retirados (tabla anterior) y 42 comunes cuyo `.cls` difiere del snapshot. La clasificación de capacidades (añadida, modificada, retirada, sin cambios) no se ha hecho.

## Capacidades

Pendiente. Las 26 fichas `CAP-*` existentes en [capabilities/](../capabilities/) son contexto, no evidencia; #807 cubre las 37 fichas que faltan.

## Discrepancias abiertas

| Id | Discrepancia | Evidencia | Estado |
|---|---|---|---|
| D-EXP-1 | El roadmap y #1073 citan 142 módulos y clases en RC1 frente a 74 en `main`. | `main` (`13af0f6`, 2026-09-15) solo tiene 2 commits que RC1 no tiene (plantillas de issues de GitHub); RC1 tiene 244 que `main` no tiene; base común `495b965`. | Explicada en lo mecánico: `main` está muy atrasada respecto a RC1. Falta comprobar que ninguno de los 74 objetos de `main` se pierde en RC1. |
| D-EXP-2 | El snapshot tiene 46 formularios y RC1 45. | Cuatro retirados y tres añadidos (tabla anterior). | Abierta: destino de los cuatro retirados sin verificar. |
| D-EXP-3 | Los 42 formularios comunes difieren del snapshot (nombres de handlers en distinta capitalización, `Attribute VB_Name` ausente, cambios de contenido). | Comparación textual `.cls`. | Abierta: la diferencia mezcla cambios cosméticos y funcionales; se separa en cada ficha. |

## Hallazgos del legacy

Comportamientos sospechosos detectados al analizar el código de RC1. Son hallazgos del estudio: no se corrigen, no generan issues en el repositorio de Expedientes y cada uno queda **a verificar** en ejecución. El PRD los conserva para decidir, caso por caso, si la web los reproduce, los corrige o los descarta.

| Id | Hallazgo | Evidencia | Ficha | Estado |
|---|---|---|---|---|
| H-EXP-1 | Tres manejadores de error escriben `Err.Numbr` en lugar de `Err.Number`: el mensaje de error no llevaría el número. | `Form_FormExpediente.cls:29, 47, 102` | [Form_FormExpediente](formularios/Form_FormExpediente.md) | a verificar |
| H-EXP-2 | `HaHabidoCambios` no compara `ContratistaPrincipal`, `EsAM`, `EsLote`, `EsBasado` ni `EsExpediente`, que `Registrar` sí persiste, y compara `ObjetoContrato` dos veces. | `Form_FormExpediente.cls:216-379`; `ExpedienteOperaciones.cls:161-181, 387` | [Form_FormExpediente](formularios/Form_FormExpediente.md) | a verificar |
| H-EXP-3 | Con un único tipo de alta permitido, el selector abre el alta y se cierra justo después; el evento `Alta` puede quedar sin quien lo escuche. `ComandoAlta_Click` se invoca sin comprobar `EsAdministrador`. | `Form_FormExpedienteAltaTipo.cls:52-56` | [Form_FormExpedienteAltaTipo](formularios/Form_FormExpedienteAltaTipo.md) | a verificar |
| H-EXP-4 | `HaHabidoCambios` y el aviso «Ningún cambio que guardar» ignoran los cambios en hijos distintos del árbol de suministradores. | `Form_FormExpediente.cls:71-76` | [Form_FormExpediente](formularios/Form_FormExpediente.md) | a verificar |
| H-EXP-5 | `HPSAplica` no pasa al DTO en ninguno de los dos `RellenarDTO`: un cambio manual no se guardaría; solo cuenta la regla «ámbito HPS ⇒ Sí» de `Registrar`. | `Form_FormExpedienteGeneral.cls:29, 299-300, 846-902`; `ExpedienteOperaciones.cls:66` | [Form_FormExpedienteGeneral](formularios/Form_FormExpedienteGeneral.md) | a verificar |
| H-EXP-6 | El `RellenarDTO` propio de General (el que corre en `Form_Unload`) no copia `POSTAGEDO` (control `EnListas`) ni `AplicaTareaS4H`; el del contenedor sí. | `Form_FormExpedienteGeneral.cls:409-417, 580-587, 846-902`; `Form_FormExpediente.cls:545-546` | [Form_FormExpedienteGeneral](formularios/Form_FormExpedienteGeneral.md) | a verificar |
| H-EXP-7 | Cada recarga de combos añade otra vez los ordinales 01-20 sin vaciar antes la lista, y el literal «Lote sin Acuero Marco» tiene un posible error tipográfico que impediría fijar `EsAM`, `EsLote` y `EsExpediente` para ese tipo. | `Form_FormExpedienteGeneral.cls:565-568, 741` | [Form_FormExpedienteGeneral](formularios/Form_FormExpedienteGeneral.md) | a verificar |
| H-EXP-8 | `ComandoFECHACERTIFICACION_Click` comprueba `IsNumeric` donde cabría `IsDate`, de modo que puede no recalcular el fin de garantía. | `Form_FormExpedienteFechas.cls:15-19` | [Form_FormExpedienteFechas](formularios/Form_FormExpedienteFechas.md) | a verificar |
| H-EXP-9 | `Registrar` calcula `ESTADO` antes de recalcular `FechaFinGarantia`: el estado usaría la garantía anterior. | `ExpedienteOperaciones.cls:68-69` | [Form_FormExpedienteFechas](formularios/Form_FormExpedienteFechas.md) | a verificar |
| H-EXP-10 | `Form_Unload` de Fechas ignora el error devuelto por `RellenarDTO`; `GARANTIAMESES` admite decimales y negativos y no se valida el orden cronológico entre fechas. | `Form_FormExpedienteFechas.cls:238-255` | [Form_FormExpedienteFechas](formularios/Form_FormExpedienteFechas.md) | a verificar |
| H-EXP-11 | En el alta de un expediente, `ColHitos` se rellena en memoria pero `Registrar` no recorre esa colección: los hitos de un expediente nuevo no se persistirían. | `Form_FormExpedienteHitos.cls` (ficha); `ExpedienteOperaciones.cls:402-495`; `ExpedienteDTO.cls:26` | [Form_FormExpedienteHitos](formularios/Form_FormExpedienteHitos.md) | a verificar |
| H-EXP-12 | Registrar un hito con fecha ya existente sustituye el anterior sin avisar en un expediente existente; en uno nuevo el comportamiento difiere. `EliminarHito` lanza `Err.Raise 1000` sin asignar `p_Error`, así que sus validaciones fallan en silencio. | `ExpedienteOperaciones.cls:1850-1851` y ficha | [Form_FormExpedienteHitos](formularios/Form_FormExpedienteHitos.md) | a verificar |
| H-EXP-13 | `RegistrarModificado` y `EliminarModificado` pierden en silencio sus errores de validación (`Err.Raise 1000` sin `p_Error`); dar de alta un número de modificado existente sobrescribe el registro; Eliminar no vacía la selección y el segundo borrado puede caer sobre el registro anterior. | `ExpedienteOperaciones.cls:1899-1911, 1940-1942, 1967-1969`; `Form_FormExpedienteModificados.cls:143-170` | [Form_FormExpedienteModificados](formularios/Form_FormExpedienteModificados.md) | a verificar |
| H-EXP-14 | En el alta de un expediente, `ColAnexos` usa el nombre del fichero como clave pero `Registrar` y `MotivoNoOK` la usan como ruta: guardar un expediente nuevo con anexos fallaría con «El anexo no es accesible». | `Form_FormExpedienteDocumentacion.cls:41`; `ExpedienteOperaciones.cls:463, 1049` | [Form_FormExpedienteDocumentacion](formularios/Form_FormExpedienteDocumentacion.md) | a verificar |
| H-EXP-15 | `EliminarAnexo` usa `Execute` sin `dbFailOnError` y borra el fichero del recurso compartido sin vuelta atrás; un nombre de fichero con comilla simple rompe el SQL concatenado. | `ExpedienteOperaciones.cls:1557, 2662-2663` | [Form_FormExpedienteDocumentacion](formularios/Form_FormExpedienteDocumentacion.md) | a verificar |
| H-EXP-16 | Los IDs de hitos, modificados y anexos se calculan como `Max + 1` en VBA (`DameID`): riesgo de colisión con dos usuarios simultáneos. | `FUNCIONES UTILES.bas:2836-2864` | [Form_FormExpedienteHitos](formularios/Form_FormExpedienteHitos.md) | a verificar |
| H-EXP-17 | `PersistirArbolDesdeMemoria` está vacío (cuerpo comentado): el árbol de suministradores de un alta no se guardaría, aunque `Registrar` lo llama. | `ExpedienteSuministradorServicio.cls:312-316`; `ExpedienteOperaciones.cls:492` | [Form_FormExpedienteSuministradores](formularios/Form_FormExpedienteSuministradores.md) | a verificar |
| H-EXP-18 | El servicio llama al repositorio sin `p_Error` y el repositorio se traga sus errores: un `INSERT`/`UPDATE`/`DELETE` fallido acaba en `CommitTrans`. | `ExpedienteSuministradorServicio.cls:70, 126, 260, 308`; `ExpedienteSuministradorRepositorio.bas:74-77` | [Form_FormExpedienteSuministradores](formularios/Form_FormExpedienteSuministradores.md) | a verificar |
| H-EXP-19 | «Añadir Aquí…» lee `TempVars("SelectedIDSuministrador")`, que ningún fichero de `src` asigna; mover solo impide soltar sobre un hijo directo, así que soltar sobre un nieto crea un ciclo. | `Form_FormExpedienteSuministradores.cls:560-564, 521-523` | [Form_FormExpedienteSuministradores](formularios/Form_FormExpedienteSuministradores.md) | a verificar |
| H-EXP-20 | Tras quitar una empresa no se recalcula `ContratistaPrincipal` en el DTO y `Registrar` escribe el valor anterior; en edición, las operaciones no actualizan `ColArbolSuministradores`, que el contenedor usa para «Ningún cambio que guardar». | `Form_FormExpedienteSuministradores.cls:874-878`; `ExpedienteOperaciones.cls:387`; `Form_FormExpediente.cls:71-76` | [Form_FormExpedienteSuministradores](formularios/Form_FormExpedienteSuministradores.md) | a verificar |
| H-EXP-21 | El guardado de responsables desde el modal llama a `ProcesarEdit` con `p_db = Nothing`: en un expediente ya guardado el cambio solo llega a memoria, pero se muestra «Responsable guardado correctamente.». | `Form_FormExpedientesResponsables.cls:133, 142`; `ExpedienteOperaciones.cls:400-402` | [Form_FormExpedientesResponsables](formularios/Form_FormExpedientesResponsables.md) | a verificar |
| H-EXP-22 | La regla de responsables se valida antes de aplicar el cambio, no sobre el estado posterior; `EsPreventa` se pierde en un alta porque `RegistraResponsable` no lo escribe. | `Form_FormExpedientesResponsables.cls:127, 133`; `ExpedienteOperaciones.cls:1072-1075, 1511-1517`; `modResponsablesEdit.bas:234` | [Form_FormExpedientesResponsables](formularios/Form_FormExpedientesResponsables.md) | a verificar |
| H-EXP-23 | `g_NecesitaRefrescarResponsables` no lo asigna nadie: el refresco por `Form_Activate` de Entidades nunca corre; el guardado del modal se habilita sin mirar `m_AbiertoParaEditar`. | `Form_FormExpedienteEntidades.cls:1372`; `Form_FormExpedientesResponsables.cls:92` | [Form_FormExpedientesResponsables](formularios/Form_FormExpedientesResponsables.md) | a verificar |
| H-EXP-24 | En el alta, `ColAnualidades` guarda el texto `Presupuesto` y `Registrar` hace `Set` sobre él: guardar un alta con anualidades fallaría y revertiría la transacción. Una anualidad con año repetido se ignora en alta pero sustituye la fila en un expediente existente. | `Form_FormExpedienteEntidades.cls:61-62`; `ExpedienteOperaciones.cls:408, 1258-1261` | [Form_FormExpedienteEntidades](formularios/Form_FormExpedienteEntidades.md) | a verificar |
| H-EXP-25 | `RegistrarAnualidad` no tiene manejador de error ni llama a `RegistrarUltimoCambio`; quitar un responsable no comprueba la regla de al menos un JP y un aviso; las altas y bajas de CPV no actualizan derivados ni caché. | `ExpedienteOperaciones.cls:1238-1282, 2562-2592, 1351-1382, 2461-2482` | [Form_FormExpedienteEntidades](formularios/Form_FormExpedienteEntidades.md) | a verificar |
| H-EXP-26 | Ninguna baja de entidades pide confirmación y los botones «Nuevo …» no llevan la etiqueta `EJECUTIVO`, así que siguen activos en solo lectura. | `Form_FormExpedienteEntidades.cls:667-679`; `Form_FormExpedienteEntidades.form.txt:2235-2612` | [Form_FormExpedienteEntidades](formularios/Form_FormExpedienteEntidades.md) | a verificar |
| H-EXP-27 | «Convertir a Lote» sin padre no cambia indicadores pero guarda y cierra como éxito; «Convertir a Lote de Acuerdo Marco» no comprueba que el padre exista ni que sea AM; nada impide indicar el propio expediente como padre. | `Form_FormExpedienteCambioTipo.cls:93-153` | [Form_FormExpedienteCambioTipo](formularios/Form_FormExpedienteCambioTipo.md) | a verificar |
| H-EXP-28 | Los cambios de tipo se aplican al expediente activo en memoria antes de guardar (si falla, queda modificado); `TbExpedientes`, la caché y el último cambio no van en una transacción común; si el expediente no existe, `RegistrarCambioTipo` no escribe ni devuelve error. | `Form_FormExpedienteCambioTipo.cls:84-156`; `ExpedienteOperaciones.cls:2014-2041` | [Form_FormExpedienteCambioTipo](formularios/Form_FormExpedienteCambioTipo.md) | a verificar |
| H-EXP-29 | Gestión no aplica a los basados la comprobación de derivados que sí aplica a Acuerdo Marco y Lote antes de cambiar de tipo. | `Form_FormExpedientesGestion.cls:2656-2667` | [Form_FormExpedienteCambioTipo](formularios/Form_FormExpedienteCambioTipo.md) | a verificar |
| H-EXP-30 | El alta para HPS probablemente no puede guardarse: `ComandoRegistrar` está deshabilitado en el diseño y el código nunca lo habilita; además el formulario no rellena `APLICAESTADO`, por lo que `MotivoNoOK` rechazaría el alta. | `Form_FormExpedienteAltaParaHPS.form.txt:2312, 2320`; `ExpedienteOperaciones.cls:839-842` | [Form_FormExpedienteAltaParaHPS](formularios/Form_FormExpedienteAltaParaHPS.md) | a verificar |
| H-EXP-31 | `MotivoNoOKSoloHPS` (exige al menos una empresa y un lugar) no se llama desde ningún sitio, y el tipo «Expediente Sólo para HPS» se sobrescribe al guardar y queda como expediente individual. | `ExpedienteOperaciones.cls:82, 1131`; `FUNCIONES UTILES.bas:5434-5436` | [Form_FormExpedienteAltaParaHPS](formularios/Form_FormExpedienteAltaParaHPS.md) | a verificar |
| H-EXP-32 | Los mensajes de error están cruzados (al guardar se muestra `m_Error` vacío; al cargar se pierde su texto); la clave temporal usa `Rnd` sin `Randomize`; los botones de lugares de ejecución no tienen gestión de errores. | `Form_FormExpedienteAltaParaHPS.cls:412-414, 36, 43, 283, 304, 340-372` | [Form_FormExpedienteAltaParaHPS](formularios/Form_FormExpedienteAltaParaHPS.md) | a verificar |
| H-EXP-33 | Tras buscar se marca `Selected(2)`; «Ver Todos» vacía la palabra clave y la búsqueda no se ejecuta (lista vacía); la palabra clave se concatena en el SQL sin escapar. | `Form_FormExpedientesParaCambioTipo.cls:43-45, 66, 110-111`; `FUNCIONES UTILES.bas:1228-1230` | [Form_FormExpedientesParaCambioTipo](formularios/Form_FormExpedientesParaCambioTipo.md) | a verificar |
| H-EXP-34 | La lista del selector de padre muestra todos los tipos aunque solo se aceptan AM y Lote; un Lote es aceptado aquí como padre pero «Convertir a Lote» lo rechaza después. | `Form_FormExpedientesParaCambioTipo.cls:177-180`; `Form_FormExpedienteCambioTipo.cls:102-105` | [Form_FormExpedientesParaCambioTipo](formularios/Form_FormExpedientesParaCambioTipo.md) | a verificar |

## Método del recuento

- El clon vive en `.atl/sources/expedientes-rc1/` (ruta ignorada, solo lectura) y está en `HEAD` desacoplado en el commit fijado.
- Controles: líneas `^\s*Begin ` del `.form.txt`. Procedimientos: líneas `(Private |Public )?(Sub|Function) ` del `.cls` (excluye `Property`, `Declare` y eventos declarados).
- Comparación con el snapshot: `.cls` con CR eliminados. No se comparan los `.form.txt`.
- Recuento de CodeGraph: tabla `nodes` de `.codegraph/codegraph.db` del clon.
