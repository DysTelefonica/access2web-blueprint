# Capacidades de No Conformidades

> **Estado:** borrador para revisión · parte de la cadena del PRD de [#1080](README.md). Línea base: `main` del repositorio legacy en el commit `1d7a2f826a1e3c582874e4e29886a6183efb9067`.

Este documento lista las capacidades que hay en `main` y las clasifica frente al [estudio previo](../capabilities.md) (rama `staging`, snapshot de 2026-08-06). Las capacidades del estudio previo son contexto, no evidencia: cada fila cita `fichero:línea` de `src/forms`, `src/classes` o `src/modules` en el commit fijado, y las 48 fichas de [formularios](README.md#formularios) son la evidencia de detalle.

## Cómo leer las tablas

- **Clasificación:** *añadida* (existe en `main` y no en el estudio previo), *modificada* (existe en ambos y el código difiere tras normalizar atributos, comentarios y espacios, según el [método del recuento](README.md#método-del-recuento)), *retirada* (el estudio previo la describe y `main` no la contiene) o *sin cambios*. Un formulario «igual» puede apoyarse en clases modificadas: en ese caso la fila se marca *modificada* y se dice por qué.
- **Issue propuesta:** `P-NC-n` es un identificador provisional, ordenado como las capacidades; el desglose definitivo llega con `issues-propuestas.md`.
- **Criterios de aceptación:** resumen verificable de lo que el código hace hoy. Los comportamientos dudosos van como «a verificar» y no como criterio.
- Las rutas de módulos y clases son las de `src/modules` y `src/classes`; el resto de abreviaturas se evita para que cada cita se pueda comprobar con una búsqueda literal.

## Resumen

| Clasificación | Capacidades |
|---|---|
| Sin cambios | CAP-NC-16, CAP-NC-27, CAP-NC-29, CAP-NC-31, CAP-NC-32, CAP-NC-33, CAP-NC-36, CAP-NC-37 |
| Modificada | CAP-NC-1 a CAP-NC-3, CAP-NC-5 a CAP-NC-15, CAP-NC-17 a CAP-NC-26, CAP-NC-28, CAP-NC-30, CAP-NC-34, CAP-NC-35, CAP-NC-44, CAP-NC-45 |
| Mixta (menú raíz sin cambios, submenús modificados) | CAP-NC-4 |
| Añadida | CAP-NC-38 |
| Retirada | CAP-NC-39 a CAP-NC-43 |

La clasificación de los formularios sale de la columna «Frente al snapshot» del [inventario](README.md#formularios); la de módulos y clases, de una comparación normalizada de `src/modules` y `src/classes` con el snapshot (7 módulos y 13 clases difieren; véase el [delta](README.md#delta-frente-al-estudio-previo)).

## Arranque e identidad

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-1 | Arranque: la portada carga el entorno con `EVE`, calcula el rol y decide entre menú reducido de técnico o menú completo | [Form_Form0BDTecnicos](formularios/Form_Form0BDTecnicos.md), [Form_Form0BDOpciones](formularios/Form_Form0BDOpciones.md) | Modificada: formularios iguales; `Variables Globales.bas` y `Entorno.cls` difieren del snapshot | `Form_Form0BDTecnicos.cls:87-101`, `Form_Form0BDTecnicos.cls:125`, `Variables Globales.bas:294` | Al abrir, se llama a `EVE`; un técnico ve solo NC de proyecto, seguimiento y ayuda; el resto pasa al menú raíz y la portada se cierra | P-NC-1 |
| CAP-NC-2 | Identidad y roles: administrador, calidad y técnico, con exclusión mutua entre ellos; usuario actual resuelto por `getUsuario` | (sin formulario propio) | Modificada: el estudio previo daba tres roles por lookup en `ColUsuariosAdministradores`; `main` los calcula en el objeto de permisos | `Variables Globales.bas:97-99`, `Variables Globales.bas:391-400`, `constructor.bas:119` | Un administrador no es calidad ni técnico, y un usuario de calidad no es administrador ni técnico; el rol condiciona los botones de cada pantalla | P-NC-1 |
| CAP-NC-3 | Configuración de backend e `IDAplicacion` leídos de la configuración local, con distinción de entorno de pruebas | (sin formulario propio) | Modificada: `main` lee `IDAplicacion` de la configuración (`Variables Globales.bas:563`) además del valor fijo por entorno (`Variables Globales.bas:331-333`) | `Variables Globales.bas:523`, `Variables Globales.bas:331-333`, `Variables Globales.bas:563` | La aplicación arranca con el backend activo configurado; `IDAplicacion` tiene un valor por entorno; sin configuración legible se informa del error. Destino de la configuración: a decidir con el ERD global | P-NC-1 |
| CAP-NC-4 | Menús: raíz, submenú de proyectos y de auditorías, con el indicador «Seguimiento usuario / total» | [Form_Form0BDOpciones](formularios/Form_Form0BDOpciones.md), [Form_Form0BDOpcionesParteProyectos](formularios/Form_Form0BDOpcionesParteProyectos.md), [Form_Form0BDOpcionesAuditorias](formularios/Form_Form0BDOpcionesAuditorias.md) | Sin cambios en el menú raíz; modificada en los dos submenús (código distinto del snapshot) | `Form_Form0BDOpciones.cls:7`, `Form_Form0BDOpcionesParteProyectos.cls:16`, `Form_Form0BDOpcionesAuditorias.cls:15`, `Funciones Generales.bas:75` | Cada submenú abre las pantallas de su dominio y repinta el indicador al abrir y al volver (`Form_Form0BDOpcionesParteProyectos.cls:308`, `Form_Form0BDOpcionesAuditorias.cls:270`) | P-NC-2 |

## NC de auditoría

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-5 | Selección de auditoría y alta o edición de una NC de auditoría (datos únicos, requiere acción correctiva y control de eficacia) | [Form_FormAuditoriaSeleccion](formularios/Form_FormAuditoriaSeleccion.md), [Form_FormNCAuditoria](formularios/Form_FormNCAuditoria.md), [Form_FormNCAuditoriaGeneral](formularios/Form_FormNCAuditoriaGeneral.md) | Modificada: `Form_FormNCAuditoriaGeneral` y `NCaUDITORIAOperaciones.cls` difieren del snapshot | `Form_FormAuditoriaSeleccion.cls:33`, `NCaUDITORIAOperaciones.cls:107`, `NCaUDITORIAOperaciones.cls:18` | Se elige auditoría antes de abrir la ficha; `RegistrarDatosUnicos` rechaza combinaciones no válidas (p. ej. requiere control de eficacia sin requerir acción correctiva, `NCaUDITORIAOperaciones.cls:74`) | P-NC-3 |
| CAP-NC-6 | Gestión de NC de auditoría: filtros, selección, alta, borrado y habilitación, historial, correo e informe | [Form_FormNCAuditoriaGestion](formularios/Form_FormNCAuditoriaGestion.md) | Modificada | `Form_FormNCAuditoriaGestion.cls:690`, `Form_FormNCAuditoriaGestion.cls:570`, `Form_FormNCAuditoriaGestion.cls:356`, `Form_FormNCAuditoriaGestion.cls:274` | La lista se filtra por los criterios del formulario; Eliminar pide confirmación y motivo; el historial se muestra en HTML | P-NC-3 |

## NC de proyecto

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-7 | Alta y edición de una NC de proyecto: expediente, descripción, causa raíz, tipología, responsables, control de eficacia y riesgos | [Form_FormNCProyecto](formularios/Form_FormNCProyecto.md), [Form_FormNCProyectoGeneral](formularios/Form_FormNCProyectoGeneral.md) | Modificada: `Form_FormNCProyectoGeneral` y `NCProyectoOperaciones.cls` difieren del snapshot | `NCProyectoOperaciones.cls:324`, `Form_FormNCProyectoGeneral.cls:461`, `NCProyectoOperaciones.cls:446` | Un alta válida crea la NC y deja registro de log «Alta de NC»; una edición deja «Edición de NC/Obs» con el estado previo para comparar | P-NC-4 |
| CAP-NC-8 | Búsqueda y listado de NC de proyecto con filtros combinables y filtrado en memoria; exportación de la lista a Excel y a Word | [Form_FormNCProyectoGestion](formularios/Form_FormNCProyectoGestion.md) | Modificada: formulario con código distinto del snapshot | `Form_FormNCProyectoGestion.cls:833-862`, `Form_FormNCProyectoGestion.cls:1097-1171`, `Form_FormNCProyectoGestion.cls:107`, `Form_FormNCProyectoGestion.cls:388` | Con el estado «Abiertas» solo salen NC abiertas; cada filtro admite limpiar; el doble clic abre el detalle | P-NC-4 |
| CAP-NC-9 | Borrado lógico con motivo y habilitación de NC de proyecto y de auditoría, con recálculo de datos calculados | [Form_FormNCProyectoGestion](formularios/Form_FormNCProyectoGestion.md), [Form_FormNCAuditoriaGestion](formularios/Form_FormNCAuditoriaGestion.md) | Modificada: ambas clases de operaciones difieren del snapshot | `NCProyectoOperaciones.cls:833`, `NCProyectoOperaciones.cls:980`, `NCProyectoOperaciones.cls:1497`, `NCaUDITORIAOperaciones.cls:285`, `NCaUDITORIAOperaciones.cls:428` | Un técnico no puede eliminar ni habilitar (`Form_FormNCProyectoGestion.cls:325-328`); «Eliminar» exige motivo; «Habilitar» revierte | P-NC-4 |
| CAP-NC-10 | Alta de una NC de proyecto vinculada a otra existente y borrado en cascada de las vinculadas | [Form_FormNCProyectoGeneralConVinculoNC](formularios/Form_FormNCProyectoGeneralConVinculoNC.md) | Modificada: formulario igual; `NCProyectoOperaciones.cls` difiere | `NCProyectoOperaciones.cls:594`, `NCProyectoOperaciones.cls:718`, `NCProyectoOperaciones.cls:1418`, `Form_FormNCProyectoGestion.cls:34-48` | La NC nueva parte de expediente, descripción, entidad responsable y detectado por de la origen; el alta deja log de vinculación (`NCProyectoOperaciones.cls:701`) | P-NC-4 |

## Acciones (AC y AR)

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-11 | Acciones correctivas (AC) de proyecto y de auditoría: alta, edición, borrado, responsable, número de acción | [Form_FormNCProyectoAC](formularios/Form_FormNCProyectoAC.md), [Form_FormNCAuditoriaAC](formularios/Form_FormNCAuditoriaAC.md), [Form_FormNCProyectoAcciones](formularios/Form_FormNCProyectoAcciones.md), [Form_FormNCAuditoriaAcciones](formularios/Form_FormNCAuditoriaAcciones.md) | Modificada: `ACProyecto.cls`, `ACProyectoOperaciones.cls`, `ACAuditoriaOperaciones.cls` y ambos formularios de acciones difieren | `ACProyectoOperaciones.cls:111`, `ACProyectoOperaciones.cls:449`, `ACAuditoriaOperaciones.cls:109`, `Form_FormNCProyectoAcciones.cls:17`, `Form_FormNCAuditoriaAcciones.cls:16` | `MotivoNoOK` valida antes de registrar; la lista maestro-detalle muestra AC arriba y sus AR abajo; borrar una AC pide confirmación | P-NC-5 |
| CAP-NC-12 | Tareas o acciones realizadas (AR): alta, detalle, edición, borrado, responsable y fechas de inicio, fin previsto y fin real | [Form_FormNCProyectoAR](formularios/Form_FormNCProyectoAR.md), [Form_FormNCAuditoriaAR](formularios/Form_FormNCAuditoriaAR.md) | Modificada: `ARProyecto.cls`, `ARProyectoOperaciones.cls`, `ARAuditoriaOperaciones.cls` y ambos formularios difieren | `ARProyectoOperaciones.cls:112`, `ARAuditoriaOperaciones.cls:114`, `ARAuditoriaOperaciones.cls:569`, `Form_FormNCProyectoAcciones.cls:349`, `Form_FormNCAuditoriaAcciones.cls:333` | Una AR se registra con su AC activa; las fechas se validan en `Registrar`; el borrado actualiza los datos calculados | P-NC-5 |
| CAP-NC-13 | Historial de AC y de AR, y recálculo de estados de acciones | [Form_FormNCProyectoAcciones](formularios/Form_FormNCProyectoAcciones.md), [Form_FormNCAuditoriaAcciones](formularios/Form_FormNCAuditoriaAcciones.md) | Modificada | `Form_FormNCProyectoAcciones.cls:235`, `Form_FormNCProyectoAcciones.cls:276`, `Form_FormNCAuditoriaAcciones.cls:260`, `ACProyectoOperaciones.cls:359`, `ARProyectoOperaciones.cls:459` | Los botones de actualizar recalculan el estado de las AC y AR de la NC y refrescan la lista | P-NC-5 |

## Control de eficacia

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-14 | Plan de control de eficacia (fecha prevista y descripción) y su alta modal | [Form_FormNCProyectoControlEficaciaAlta](formularios/Form_FormNCProyectoControlEficaciaAlta.md), [Form_FormNCAuditoriaControlEficaciaAlta](formularios/Form_FormNCAuditoriaControlEficaciaAlta.md) | Modificada: el formulario de auditoría difiere; el de proyecto es igual; las operaciones difieren | `NCProyectoOperaciones.cls:1039`, `NCaUDITORIAOperaciones.cls:484`, `Form_FormNCProyectoGeneral.cls:249`, `Form_FormNCAuditoriaGeneral.cls:50` | El alta del plan se lanza desde la pestaña General y se graba con la NC; eliminar el plan es posible (`NCProyectoOperaciones.cls:1112`) | P-NC-6 |
| CAP-NC-15 | Resultado del control de eficacia (fecha, resultado y conformidad) y su eliminación | [Form_FormNCProyectoControlEficacia](formularios/Form_FormNCProyectoControlEficacia.md), [Form_FormNCAuditoriaControlEficacia](formularios/Form_FormNCAuditoriaControlEficacia.md) | Modificada | `NCProyectoOperaciones.cls:1184`, `NCProyectoOperaciones.cls:1285`, `NCaUDITORIAOperaciones.cls:561`, `Form_FormNCProyectoControlEficacia.cls:36`, `Form_FormNCAuditoriaControlEficacia.cls:35` | «Grabar» registra el resultado y la conformidad; solo es posible con un plan previo (a verificar en la ficha) | P-NC-6 |
| CAP-NC-16 | Motivo por el que la NC no requiere control de eficacia | [Form_FormMotivosNoRequiereControlEficacia](formularios/Form_FormMotivosNoRequiereControlEficacia.md) | Sin cambios en el formulario; los formularios que lo abren están modificados | `Form_FormMotivosNoRequiereControlEficacia.cls:18`, `Form_FormNCProyectoGeneral.cls:1163`, `Form_FormNCAuditoriaGeneral.cls:542` | El texto se captura o consulta y se devuelve al formulario de origen por evento; no escribe tablas por sí mismo | P-NC-6 |

## Replanificaciones

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-17 | Replanificación de acciones y tareas con motivo, registrada en la operación de AC y AR | (sin formulario propio de alta) | Modificada: `ACProyectoOperaciones.cls`, `ARProyectoOperaciones.cls` y `ARAuditoriaOperaciones.cls` difieren | `ACProyectoOperaciones.cls:394`, `ARProyectoOperaciones.cls:483`, `ACAuditoriaOperaciones.cls:373`, `ARAuditoriaOperaciones.cls:464`, `ReplanificacionesProyectoOperaciones.cls:107` | `ReplanificarAcciones` y `Replanificar` generan un registro de replanificación validado por `MotivoNoOK`; el alta por formulario no existe (a verificar) | P-NC-7 |
| CAP-NC-18 | Consulta de solo lectura de replanificaciones por NC (acción, tarea y motivo) | [Form_FormNCProyectoReplanificaciones](formularios/Form_FormNCProyectoReplanificaciones.md), [Form_FormNCAuditoriaReplanificaciones](formularios/Form_FormNCAuditoriaReplanificaciones.md) | Sin cambios en los formularios; clasificada como modificada por depender de CAP-NC-17 | `Form_FormNCProyectoReplanificaciones.cls:119`, `ReplanificacionesAuditoriaOperaciones.cls:104`, `ReplanificacionesAuditoriaOperaciones.cls:16` | Al seleccionar una fila se muestran acción, tarea y motivo; la pantalla no escribe datos | P-NC-7 |

## Documentos

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-19 | Documentos de NC: listar, anexar (Examinar y Aceptar), cambiar título, abrir y eliminar | [Form_FormNCProyectoDocumentos](formularios/Form_FormNCProyectoDocumentos.md), [Form_FormNCAuditoriaDocumentos](formularios/Form_FormNCAuditoriaDocumentos.md) | Modificada: ambos formularios y `DocumentoProyectoOperaciones.cls` difieren | `DocumentoProyectoOperaciones.cls:192`, `DocumentoAuditoriaOperaciones.cls:18`, `Form_FormNCProyectoDocumentos.cls:338`, `Form_FormNCAuditoriaDocumentos.cls:205` | Un anexo nuevo exige título válido (`MotivoNoOK`); el cambio de título y el borrado piden la fila seleccionada | P-NC-8 |
| CAP-NC-20 | Documentos de evidencia de AR (proyecto y auditoría) | [Form_FormARProyectoDocumentos](formularios/Form_FormARProyectoDocumentos.md), [Form_FormARAuditoriaDocumentos](formularios/Form_FormARAuditoriaDocumentos.md) | Modificada: ambos formularios difieren | `DocumentoProyectoOperaciones.cls:270`, `DocumentoAuditoriaOperaciones.cls:211`, `Form_FormNCProyectoAcciones.cls:157` | Se gestionan sobre la AR activa; el alta de proyecto es en dos pasos y el de auditoría difiere (véase cada ficha) | P-NC-8 |
| CAP-NC-21 | Documentos de auditoría, con modo memoria (antes del alta) y modo directo | [Form_FormAuditoriaDocumentos](formularios/Form_FormAuditoriaDocumentos.md) | Modificada: formulario con código distinto del snapshot | `Form_FormAuditoria.cls:20`, `AuditoriaOperaciones.cls:189`, `Form_FormAuditoriasGestion.cls:104` | Antes del alta los anexos se acumulan en memoria y se graban al registrar la auditoría; con la auditoría existente se escribe directamente | P-NC-8 |

## Notas y motivos

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-22 | Nota libre de la NC: guardar y borrar; en una NC vinculada se muestra la de la asociada en solo lectura | [Form_FormNCProyectoNota](formularios/Form_FormNCProyectoNota.md), [Form_FormNCAuditoriaNota](formularios/Form_FormNCAuditoriaNota.md) | Modificada por las operaciones; formularios iguales | `NCProyectoOperaciones.cls:1342`, `NCProyectoOperaciones.cls:1386`, `NCaUDITORIAOperaciones.cls:719`, `Form_FormNCProyectoNota.cls:184` | Grabar persiste la nota; borrar la vacía; una nota vacía previa se trata aparte (`NCProyectoOperaciones.cls:1359`) | P-NC-9 |
| CAP-NC-23 | Motivo de borrado de la NC: ver y corregir | [Form_FormNCProyectoMotivoEliminado](formularios/Form_FormNCProyectoMotivoEliminado.md), [Form_FormNCAuditoriaMotivoEliminado](formularios/Form_FormNCAuditoriaMotivoEliminado.md) | Modificada por las operaciones; formularios iguales | `NCProyectoOperaciones.cls:1457`, `NCaUDITORIAOperaciones.cls:795`, `Form_FormNCProyectoMotivoEliminado.cls:171`, `Form_FormNCAuditoriaMotivoEliminado.cls:158` | «Grabar» modifica el motivo de una NC ya borrada | P-NC-9 |

## Seguimientos y tareas

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-24 | Seguimiento de tareas (pestaña «Tareas»): filtros por responsable, estado y expediente o auditoría, apertura de la NC y exportación a Excel | [Form_FormNCProyectoSeguimientoTareas](formularios/Form_FormNCProyectoSeguimientoTareas.md), [Form_FormNCAuditoriaSeguimientoTareas](formularios/Form_FormNCAuditoriaSeguimientoTareas.md), [Form_FormNCProyectoSeguimiento](formularios/Form_FormNCProyectoSeguimiento.md), [Form_FormNCAuditoriaSeguimiento](formularios/Form_FormNCAuditoriaSeguimiento.md) | Modificada: los cuatro formularios y `SegTareasProyecto.cls` difieren | `Form_FormNCProyectoSeguimientoTareas.cls:109`, `Form_FormNCProyectoSeguimientoTareas.cls:219`, `Form_FormNCAuditoriaSeguimientoTareas.cls:34`, `Funciones Generales.bas:380` | La lista sale de las colecciones de seguimiento del entorno; «Ver detalle» abre la NC de la tarea elegida | P-NC-10 |
| CAP-NC-25 | Revisión de estados de NC (pestaña «Rev. Estados») con los mismos filtros | [Form_FormNCProyectoSeguimientoNC](formularios/Form_FormNCProyectoSeguimientoNC.md), [Form_FormNCAuditoriaSeguimientoNC](formularios/Form_FormNCAuditoriaSeguimientoNC.md) | Modificada: ambos formularios difieren | `Form_FormNCProyectoSeguimientoNC.cls:171`, `Form_FormNCAuditoriaSeguimientoNC.cls:64`, `Form_FormNCAuditoriaSeguimientoNC.cls:267` | Filtra por responsable de calidad, estado y expediente o auditoría; abre la NC elegida | P-NC-10 |
| CAP-NC-26 | Log de actividad por NC (proyecto y auditoría) y consulta del historial en HTML | [Form_FormNCProyectoGestion](formularios/Form_FormNCProyectoGestion.md), [Form_FormNCAuditoriaGestion](formularios/Form_FormNCAuditoriaGestion.md) | Modificada: `LogNCProyecto.cls` y `LogNCAuditoria.cls` no cambian, pero `Funciones Generales.bas` (que escribe el log) difiere | `Funciones Generales.bas:993`, `Funciones Generales.bas:1098`, `LogNCProyecto.cls:186`, `Form_FormNCProyectoGestion.cls:146` | Toda alta, edición y vinculación deja un registro de log; el historial se abre en HTML | P-NC-10 |

## Auditorías (gestión)

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-27 | Alta y edición de la cabecera de una auditoría: tipo (interna o externa), fecha de inicio y de fin | [Form_FormAuditoria](formularios/Form_FormAuditoria.md) | Sin cambios en el formulario; clase `AuditoriaOperaciones.cls` igual | `Form_FormAuditoria.cls:55`, `Form_FormAuditoria.cls:154`, `AuditoriaOperaciones.cls:52` | `Registrar` valida y graba; los documentos se adjuntan desde el mismo formulario | P-NC-11 |
| CAP-NC-28 | Catálogo de auditorías: filtrar por año y tipo, cambiar fechas, anexos, eliminar y ver sus NC | [Form_FormAuditoriasGestion](formularios/Form_FormAuditoriasGestion.md) | Modificada: formulario con código distinto del snapshot | `Form_FormAuditoriasGestion.cls:145`, `Form_FormAuditoriasGestion.cls:189`, `Form_FormAuditoriasGestion.cls:397`, `AuditoriaOperaciones.cls:122` | Eliminar pide confirmación; «Ver NC» y «Alta NC» abren las pantallas de NC de auditoría | P-NC-11 |

## Tipologías

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-29 | Catálogo de tipologías de NC de proyecto: alta, edición y eliminación | [Form_FormNCProyectoTipologiaGestion](formularios/Form_FormNCProyectoTipologiaGestion.md), [Form_FormTipologiaNCProyecto](formularios/Form_FormTipologiaNCProyecto.md) | Sin cambios | `TipologiaNCProyectos.cls:125`, `TipologiaNCProyectos.cls:199`, `Form_FormNCProyectoTipologiaGestion.cls:20`, `Form_FormNCProyectoTipologiaGestion.cls:95` | La lista refleja las altas y ediciones; la eliminación pide confirmación | P-NC-12 |

## Vinculación con riesgos y expedientes

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-30 | Asociar riesgos del expediente a una NC de proyecto (multiselección, en memoria durante el alta) | [Form_formRiesgosSeleccion](formularios/Form_formRiesgosSeleccion.md) | Modificada: `RiesgoServicio1.bas` es nuevo respecto al snapshot y está vacío en `main` | `Form_formRiesgosSeleccion.cls:62`, `Form_formRiesgosSeleccion.cls:156`, `RiesgoRepositorio.bas:47`, `NCProyectoOperaciones.cls:1531` | Los candidatos son los riesgos del expediente de la NC; «Aceptar» guarda la selección en la base si la NC existe | P-NC-13 |
| CAP-NC-31 | Búsqueda de expedientes y devolución del elegido por evento | [Form_FormExpedientesBusqueda](formularios/Form_FormExpedientesBusqueda.md) | Sin cambios (formulario igual y `Expediente.cls` sin diferencias) | `Form_FormExpedientesBusqueda.cls:47`, `Form_FormExpedientesBusqueda.cls:168`, `Form_FormNCProyectoGestion.cls:190` | Filtra por palabra clave y responsable de calidad; «Elegir» devuelve el expediente al formulario que lo abrió | P-NC-13 |

## Indicadores

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-32 | Panel de indicadores semestrales de NC de proyecto (abiertas, replanificadas, stock activo y con riesgo) con exportación del detalle a Excel | [Form_FormIndicadores](formularios/Form_FormIndicadores.md) | Sin cambios | `Form_FormIndicadores.cls:74`, `Form_FormIndicadores.cls:141`, `IndicadorServicio.cls:80`, `IndicadorRepositorio.bas:11` | «Calcular» muestra los cuatro recuentos del periodo; cada botón de Excel exporta el detalle de su indicador | P-NC-14 |

## Correo e informes

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-33 | Correo de la NC: componer el informe HTML y registrarlo como correo pendiente (el envío real no es de este formulario) | [Form_FormCorreo](formularios/Form_FormCorreo.md) | Sin cambios | `Form_FormCorreo.cls:14`, `Correo.cls:91`, `HTML.bas:784`, `HTML.bas:865` | «Enviar correo» registra el correo con asunto, destinatarios y copia rellenados | P-NC-15 |
| CAP-NC-34 | Aviso por correo al administrador ante errores | (sin formulario propio) | Modificada: `Funciones Generales.bas` difiere; `Correo.cls` no cambia | `Funciones Generales.bas:1227` | Un error capturado genera un aviso al administrador | P-NC-15 |
| CAP-NC-35 | Informe Word de una NC o de la lista filtrada, y exportación de la lista a Excel | [Form_FormNCProyectoGestion](formularios/Form_FormNCProyectoGestion.md), [Form_FormNCAuditoriaGestion](formularios/Form_FormNCAuditoriaGestion.md) | Modificada: `Funciones Generales.bas` difiere | `Informe.cls:15`, `Funciones Generales.bas:3978`, `Funciones Generales.bas:4367`, `Form_FormNCAuditoriaGestion.cls:531` | El informe se genera desde la plantilla con la NC seleccionada o con todas las filtradas | P-NC-15 |
| CAP-NC-36 | Informe HTML de NC y de historial | [Form_FormNCProyectoGestion](formularios/Form_FormNCProyectoGestion.md) | Sin cambios en `HTML.bas` | `HTML.bas:946` | El informe de una o varias NC se genera en HTML | P-NC-15 |
| CAP-NC-37 | Utilidades de normalización de datos del `Instalador` (estados, tipologías, responsables) | (sin formulario) | Sin cambios | `Instalador.bas:6`, `Instalador.bas:655` | Procesos manuales de relleno de datos calculados; su alcance en la web: a decidir | P-NC-16 |
| CAP-NC-38 | Clase de informe de NC de auditorías (`InformeNCAuditorias`), vacía de procedimientos | (sin formulario) | Añadida: no está en el snapshot | `InformeNCAuditorias.cls:11-12` | Solo declara el error; sin comportamiento que aceptar | P-NC-15 |

## Capacidades del estudio previo ausentes en `main`

Existen en el snapshot y no en el commit fijado; la lista de módulos y clases retirados está en [Objetos retirados frente al estudio previo](README.md#objetos-retirados-frente-al-estudio-previo).

| ID | Capacidad del estudio previo | Formularios | Clasificación | Evidencia de la ausencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-39 | Caché de indicadores (`ModuloCacheIndicadores`) | (ninguno) | Retirada | Módulo ausente de `src/modules`; el único acceso a datos de indicadores es `IndicadorRepositorio.bas:11` | Los indicadores se calculan sobre las tablas, sin caché | P-NC-17 |
| CAP-NC-40 | Caché de listados de NC (`NCAuditoriaListadoCache`, `CacheNCCacheRepositorio`) | (ninguno) | Retirada | Módulos y clase ausentes; el listado filtra en memoria el conjunto de `constructor.bas:1628` | Los listados no dependen de caché | P-NC-17 |
| CAP-NC-41 | Kill switch y bancos de prueba de caché (`Test_KillSwitch`, `Test_Cache*`) | (ninguno) | Retirada | Sin coincidencias de «kill» en `src/modules` ni `src/classes` | No existe interruptor de caché en `main` | P-NC-17 |
| CAP-NC-42 | Ayudantes `modFrm*Helper`, `modFormInteractionHelper` y servicios `CierreNCService`, `DocumentoService`, `ReplanificacionesService` | (ninguno) | Retirada | Módulos y clases ausentes; la lógica vive en los formularios y en las clases `*Operaciones` | La lógica de formulario se prueba a través de los formularios y las operaciones | P-NC-17 |
| CAP-NC-43 | Viewmodels `NCProyectoListItemVM` y `NCProyectoDetailVM` | (ninguno) | Retirada | Clases ausentes de `src/classes` | La web define sus propios DTO de lista y detalle | P-NC-17 |

La caché por NC (`TbCacheNCProyecto`) **no** está retirada: `CacheNCProyecto.bas` e `InicializadorCache.bas` siguen en `main` y difieren del snapshot.

| ID | Capacidad | Formularios | Clasificación | Evidencia | Criterios de aceptación | Issue |
|---|---|---|---|---|---|---|
| CAP-NC-44 | Caché por NC (tabla de caché, JSON de NC, AC, AR, replanificaciones y riesgos), con transacción de caché | (ninguno) | Modificada | `CacheNCProyecto.bas:34-35`, `CacheNCProyecto.bas:42`, `CacheNCProyecto.bas:254`, `CacheNCProyecto.bas:608` | Si la caché es válida se lee de ella; si no, se regenera; la decisión de portarla se toma con el ERD global | P-NC-18 |
| CAP-NC-45 | Mantenimiento de la caché: estadísticas, integridad, regeneración, población masiva y limpieza de logs | (ninguno) | Modificada | `InicializadorCache.bas:287`, `InicializadorCache.bas:439`, `InicializadorCache.bas:910`, `InicializadorCache.bas:969` | Los comandos de mantenimiento listan obsoletos, verifican integridad y limpian logs antiguos | P-NC-18 |

## Capacidades que no duplican otras aplicaciones

Las siguientes capacidades de No Conformidades existen por integración y no se re-especifican aquí: su propietario es otra aplicación.

| Capacidad | Aplicación propietaria | Qué conserva No Conformidades | Evidencia |
|---|---|---|---|
| Identidad y permisos de usuarios (CAP-NC-2) | Lanzadera | Consume el usuario y su rol; no mantiene altas de usuarios | `constructor.bas:119`, `UsuarioAplicacionPermisos.cls:285` |
| Riesgos (CAP-NC-30) | Gestión de Riesgos | Solo el vínculo NC-riesgo y la selección; el riesgo no se edita aquí | `RiesgoRepositorio.bas:7`, `RiesgoServicio.cls:14` |
| Búsqueda de expedientes (CAP-NC-31) | Expedientes | Solo el selector de lectura; los expedientes no se editan aquí | `Expediente.cls:429`, `Form_FormExpedientesBusqueda.cls:47` |

## Pendientes

- Las filas marcadas «a verificar» requieren abrir el binario o el backend; el ERD de origen está pendiente del propietario (D-NC-3).
- La comparación de módulos y clases con el snapshot es normalizada y mecánica; no distingue refactor de cambio de comportamiento.
- Falta decidir el destino web de la caché por NC, del `Instalador` y de la configuración de backend.
