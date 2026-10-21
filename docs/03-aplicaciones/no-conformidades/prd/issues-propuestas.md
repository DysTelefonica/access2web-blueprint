# Issues propuestas: No Conformidades

> **Estado:** propuesta, **no publicada**. Las issues no se crean hasta cerrar el PRD y las aprueba el propietario. Parte de la cadena de [#1080](README.md). Línea base: `main` del repositorio legacy en el commit `1d7a2f826a1e3c582874e4e29886a6183efb9067`.

Cada propuesta agrupa capacidades de [capacidades.md](capacidades.md) en una unidad implementable. La evidencia es la ficha de cada formulario ([formularios/](formularios/Form_Form0BDOpciones.md) y las demás) y las citas `fichero:línea` de la capacidad. Todas llevarían la etiqueta `app/no-conformidades`. Las propuestas de identidad, riesgos y expedientes **no duplican** otras aplicaciones: se enlazan, no se reimplementan.

| Propuesta | Capacidades | Depende de | Criterios de aceptación (resumen verificable) | Etiquetas |
|---|---|---|---|---|
| P-NC-1 Arranque, identidad y roles | CAP-NC-1 a 3 | Contrato de identidad compartido de la Lanzadera (propuesta P-LAN-3, abierta); ERD global | La portada resuelve el usuario y su rol (Administrador, Calidad o Técnico) como `main`; el rol técnico por exclusión queda decidido; la lectura de `IDAplicacion` y backend no expone secretos | `app/no-conformidades`, `identidad` |
| P-NC-2 Menús de entrada | CAP-NC-4 | P-NC-1 | Menú raíz y submenús de proyectos y auditorías con los mismos accesos por rol; los enlaces sin handler en `main` se documentan como defecto, no como requisito | `app/no-conformidades` |
| P-NC-3 NC de auditoría: alta, edición y gestión | CAP-NC-5, 6 | P-NC-11, ERD global | Alta, edición, filtros, borrado y habilitación equivalentes; errores 100/1000 tratados de forma uniforme | `app/no-conformidades` |
| P-NC-4 NC de proyecto: alta, edición, listado, borrado y vinculación | CAP-NC-7 a 10 | P-NC-12, P-NC-13, ERD global | Alta con expediente, filtros combinables, borrado lógico con motivo, alta vinculada; sin fugas de SQL concatenado | `app/no-conformidades` |
| P-NC-5 Acciones correctivas y tareas (AC/AR) | CAP-NC-11 a 13 | P-NC-3, P-NC-4 | Alta, edición, borrado e historial de AC y AR; validación de fechas y evidencias para cerrar una tarea aplicada en el servidor; recálculo de estados | `app/no-conformidades` |
| P-NC-6 Control de eficacia | CAP-NC-14 a 16 | P-NC-3, P-NC-4 | Plan y resultado con las precondiciones de `main`; motivo de no requerirlo; vaciar la fecha prevista no deja datos huérfanos | `app/no-conformidades` |
| P-NC-7 Replanificaciones | CAP-NC-17, 18 | P-NC-5 | Alta con motivo y regla de duplicados (acción realizada y fecha prevista al inicio); consulta de solo lectura | `app/no-conformidades` |
| P-NC-8 Documentos y evidencias | CAP-NC-19 a 21 | P-NC-3, P-NC-4, almacenamiento de ficheros | Anexar, renombrar, abrir y eliminar con fila y fichero atómicos; modo previo al alta de la auditoría | `app/no-conformidades`, `almacenamiento` |
| P-NC-9 Notas y motivos de borrado | CAP-NC-22, 23 | P-NC-3, P-NC-4 | Guardar y borrar nota; ver y corregir motivo; comportamiento en NC vinculada documentado | `app/no-conformidades` |
| P-NC-10 Seguimiento de tareas, revisión de estados y log | CAP-NC-24 a 26 | P-NC-5 | Filtros por responsable, estado y auditoría; exportación a hoja de cálculo; log por NC | `app/no-conformidades` |
| P-NC-11 Catálogo de auditorías | CAP-NC-27, 28 | ERD global | Alta y edición de cabecera (tipo, fechas con validación), filtro por año y tipo, anexos; borrado sin dejar documentos huérfanos | `app/no-conformidades` |
| P-NC-12 Tipologías de NC de proyecto | CAP-NC-29 | ERD global | Alta, edición y eliminación con comprobación de uso previa atómica; identificador generado por la base de datos | `app/no-conformidades` |
| P-NC-13 Vinculación con riesgos y expedientes | CAP-NC-30, 31 | Gestión de Riesgos (tablas `TbRiesgos*`), Expedientes | Selección multiselección de riesgos candidatos de la última edición; asociación sin borrado y reinserción sin transacción; búsqueda de expedientes | `app/no-conformidades` |
| P-NC-14 Indicadores de NC de proyecto | CAP-NC-32 | P-NC-4 | Panel semestral con los mismos contadores; cálculo sin la caché retirada | `app/no-conformidades` |
| P-NC-15 Correo e informes | CAP-NC-33 a 36, 38 | P-NC-3, P-NC-4, servicio de correo | Correo de NC registrado y enviado; informe HTML y Word; sin dirección de copia oculta fija | `app/no-conformidades`, `integraciones` |
| P-NC-16 Utilidades del instalador | CAP-NC-37 | Decisión del propietario | Decidir si se migra o se descarta la carga inicial de datos | `app/no-conformidades` |
| P-NC-17 Capacidades retiradas | CAP-NC-39 a 43 | Decisión del propietario | **No se implementa**: caché de indicadores y de listados, kill switch y ayudantes retirados de `main`; se confirma el descarte | `app/no-conformidades` |
| P-NC-18 Caché por NC | CAP-NC-44, 45 | Decisión del propietario, P-NC-4 | Decidir si la web necesita la caché por NC; en `main` la lectura de la caché está desactivada y solo se invalida | `app/no-conformidades` |
| P-NC-19 Identificadores y atomicidad | Transversal ([coexistencia.md](coexistencia.md)) | ERD global, política de conflictos | Identificadores generados por la base de datos en lugar de máximo + 1; escrituras multi-paso en transacción; cubre todas las propuestas con altas | `app/no-conformidades`, `datos` |
| P-NC-20 Extracción del ERD de origen | Transversal ([erd-origen.md](erd-origen.md)) | PR #1144 (workflow dysflow), ruta del backend | Metadatos de solo lectura de `NoConformidades_Datos.accdb`; filas de la matriz pasan de propuesta a confirmada; cierra D-NC-3 | `app/no-conformidades`, `datos` |

## Decisiones abiertas del propietario

- **Identidad compartida:** el rol técnico se asigna por exclusión en `main`; falta decidir si la web lo conserva.
- **Tablas sin uso en `main`:** 247 filas de la [matriz](matriz-campos.md) están pendientes por este motivo.
- **Corte en la Lanzadera:** reversión por definir ([coexistencia.md](coexistencia.md)).
- **Línea base:** `main` (producción), decidido en [#1080](https://github.com/DysTelefonica/access2web-blueprint/issues/1080#issuecomment-6322559873); los cambios de `staging` no liberados están en el anexo del [README](README.md).
