# Form_FormNCAuditoriaGestion

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoriaGestion.cls` (1338 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pantalla de gestión de las no conformidades (NC) de auditoría: filtra, lista y selecciona NC y desde ella se da de alta, se edita, se marca como borrada o se habilita, se consulta el historial, se envía por correo y se genera el informe. Es un formulario **no enlazado** (sin `RecordSource`); la lista `ListaFiltrados` se rellena por código con una lista de valores (`Form_FormNCAuditoriaGestion.cls:895, 1088`). Las citas son `Form_FormNCAuditoriaGestion.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`, `VG.bas` `modules/Variables Globales.bas` y `Ops` `classes/NCaUDITORIAOperaciones.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form0BDOpcionesAuditorias | Abre el formulario, fija `Estado = "Abiertas"`, opcionalmente el responsable con el usuario conectado, y llama a `ESTADO_AfterUpdate` | `Form_Form0BDOpcionesAuditorias.cls:128-134` |
| FormAuditoriasGestion | Lo cierra y reabre y llama a `ActualizarLista`; también refresca su combo de auditorías | `Form_FormAuditoriasGestion.cls:417-424, 237-239` |
| FormNCAuditoriaAC | Llama a `ActualizarLineaLista` tras guardar | `Form_FormNCAuditoriaAC.cls:185-186, 273-274` |

Abre `FormNCAuditoria` (alta y edición, cerrándolo antes si estaba abierto: `:251-254, 338-341, 584-587`) y `FormCorreo` con `OpenArgs = EnumSino.No` (`:98-101`). Deja la NC en la variable global `m_ObjNCAuditoriaActiva` (`:97, 336`; `VG.bas:127`) y en `m_NCSeleccionada` (pública, `:7`). Se cierra con «Salir» (`:163-166`).

## Controles y eventos

`.form.txt` declara 52 bloques `Begin`: formulario (`.form.txt:4`), 8 de valores por defecto (`.form.txt:420-478`), 3 secciones (`EncabezadoDelFormulario` `:490`, `Detalle` `:602`, `PieDelFormulario` `:2969`), 27 controles con nombre (tabla) y 13 etiquetas, imágenes o líneas sin handler.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaFiltrados` | lista (9 columnas) | `OnClick`, `OnDblClick` | Carga la NC de la fila y habilita botones; doble clic ⇒ equivale a «Detalle» si está habilitado | `:944-1013, 1028-1035`; `.form.txt:609-633` |
| `Auditoria` | combo | `OnChange` | Borra `Estado` y relanza la lista | `:13-17`; `.form.txt:648-673` |
| `Estado` | combo | `AfterUpdate` (`ESTADO_AfterUpdate`) | Filtra por estado; añade la opción «Abiertas» | `:666-688, 1143`; `.form.txt:1025-1054` |
| `Tipo` | combo `NC;OB;OP` | `AfterUpdate` | Filtra por tipo | `:1317-1337`; `.form.txt:2694-2696` |
| `RESPONSABLEIMPLANTACION` | combo | `AfterUpdate` | Filtra por responsable | `:1151-1171`; `.form.txt:2240-2265` |
| `ComboRequiereControlEficacia`, `ComboControlEficaciaRelleno` | combos `Sí/No` | `AfterUpdate` | Filtran por requerimiento y por relleno del control de eficacia | `:613-660`; `.form.txt:1660-1685, 2071-2096` |
| `Descripcion`, `PalabraClave` | cuadros de texto | `KeyDown` | Intro o Tab relanzan la lista | `:1172-1203`; `.form.txt:1273, 2554` |
| `ComandoLimpiarAuditoria`, `ComandoLimiparEstado`, `ComandoLimiparTipo`, `ComandoLimiparDescripcion`, `ComandoLimiparPalabraClave`, `ComandoLimiparRESPONSABLEIMPLANTACION`, `ComandoLimpiaoRequiereControlEficacia`, `ComandoLimpiarControlEficaciaRelleno` | 8 botones «Comando36» | `OnClick` | Vacían su filtro y relanzan la lista | `:268-272, 489-528, 604-609` |
| `ComandoActualizar` («Contador») | botón | `OnClick` | Vacía las cachés `ColAuditorias`/`ColNCsAuditoria` del entorno, recarga combos, recalcula fechas y estado de las NC filtradas y relanza la lista | `:19-44, 45-80`; `.form.txt:527-539` |
| `ComandoAlta` | botón | `OnClick` | Abre `FormNCAuditoria` sin NC (solo no técnicos) | `:570-603` |
| `ComandoAltaDeAuditoria` («Alta para …») | botón | `OnClick` | Alta de NC en la auditoría del filtro o de la NC seleccionada | `:211-266, 1014-1024` |
| `ComandoEditar` («Detalle») | botón | `OnClick` | Abre `FormNCAuditoria` con la NC seleccionada | `:318-353` |
| `ComandoEliminar` («Eliminar»/«Habilitar») | botón | `OnClick` | Marca como borrada con motivo, o habilita (solo no técnicos) | `:356-428`; `.form.txt:2487-2501` |
| `ComandoNCAuditoriaHistorial` | botón | `OnClick` | Genera el historial en texto | `:274-315` |
| `ComandoEnviarCorreoElectronico` | botón | `OnClick` | Abre `FormCorreo` | `:81-117` |
| `ComandoInforme` | botón | `OnClick` | Informe Word de la NC seleccionada | `:531-567` |
| `ComandoInformeLista` (pie) | botón | `OnClick` | Informe Word de todas las NC de la lista | `:430-487`; `.form.txt:2980-2992` |
| `cmdSalir` (pie) | botón | `OnClick` | Cierra sin guardar | `:163-166` |
| Sección `PieDelFormulario` | sección | `OnDblClick` | Genera el HTML de la NC con acciones y lo abre en local | `:120-160`; `.form.txt:2973` |
| `lblTitulo1`, `lblEstado` | etiquetas | — | Título (copiado a `Caption`) y mensajes de progreso, oculto | `:698-699` |

## Flujos

### Apertura (`Form_Open`, `:690-729`)
Ajusta el tamaño (`Ajustar Me`), carga combos (`EstablecerCombos`: responsables, auditorías, estados y «Abiertas»; `:1123-1149`) y deshabilita Alta, Detalle, Eliminar, Correo, Informe, Informe de lista y «Alta para #» (`:706-714`). No carga la lista: aparece vacía hasta aplicar un filtro. Cualquier error cierra el formulario (`:728`).

### Filtrado (`ActualizarLista`, `:862-940`)
1. Deshabilita botones y vacía la selección (`:877-894`); con `p_ManteniendoSeleccionado` recuerda el `IDNoconformidad` para reseleccionarlo (`:878-882, 914-925`).
2. `getNCsAudtoriasFiltrados` (`:731-836`) parte de las NC por palabra clave, o de las abiertas si `Estado = "Abiertas"`, o de todas (`:751-760`) y descarta en memoria las que no cumplen cada filtro informado (`:772-816`).
3. `RellenarListaConCol` (`:1068-1120`) añade una fila por NC con auditoría («Desconocida» si no se resuelve, `:1100`), tipo, número con dos cifras, descripción sin «;», responsable, estado, apertura y cierre, tras `EstadoGrabar` (`:1102`); selecciona la primera fila (`:1111-1114`).
4. Sin resultados, solo se prepara el botón «Alta para …» con la auditoría del filtro (`:900-907`).

### Selección (`ListaFiltrados_Click`, `:944-1013`)
Carga la NC con `constructor.getNCAuditoria` (`:970-979`), habilita Detalle, Eliminar (rótulo «Habilitar» si `Estado = "BORRADA"`, `:990-994`), Historial, Correo e Informe y ajusta «Alta para <auditoría>» (`:996`).

### Borrado y habilitación (`ComandoEliminar_Click`)
Para «Eliminar» pide confirmación (por defecto «No») y motivo por `InputBox` (`:384-387`), asigna `MotivoBorrado` y llama a `Ops.Eliminar EnumSino.No` (`:402-403`); para «Habilitar» confirma y llama a `Ops.Habilitar` (`:389, 405`). Después relanza la lista conservando la selección (`:412-413`).

### Errores
Patrón común: error 1000 ⇒ aviso con `m_Error`; otro error ⇒ mensaje con número y descripción, `CorreoAlAdministrador` y `MsgBox` crítico (p. ej. `:37-43`).

## Reglas, validaciones y cálculos

- **Rol:** Alta, «Alta para», Eliminar/Habilitar comprueban `EsTecnico = EnumSino.Sí` y rechazan con «No tiene autorización para esa acción» (`:220-223, 367-370, 579-582`). Detalle, Historial, Correo, Informe e Informe de lista no comprueban el rol (`:81-117, 274-353, 430-567`).
- **Motivo:** `Ops.Eliminar` exige NC registrada y motivo no vacío (`Ops:310-313`); `Habilitar` borra `Borrado` y `MotivoBorrado` (`Ops:438-454`). El `InputBox` cancelado produce motivo vacío (`:386`).
- **Filtros combinados:** conjunción (AND) en memoria; descripción por `InStr` sobre el texto completo (`:784-788`); «Abiertas» no se compara con el título del estado (`:796`); control de eficacia relleno compara con la cadena vacía (`:808-816`).
- **Estado, fecha de cierre y de previsión de cierre** se recalculan y graban con «Contador» (`:61-67`) y el estado se graba al listar (`:1102`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| NC abiertas | `TbNoConformidadesAuditoria` + `TbNCAuditoriaAccionCorrectivas` (`DISTINCT`, `Borrado=False`) | SQL dinámico (sin parámetros) | `constructor.bas:1938-1975` |
| Todas las NC | `TbNoConformidadesAuditoria` | SQL | `constructor.bas:5762-5800` |
| Palabra clave | `TbNoConformidadesAuditoria` con `LEFT JOIN` a acciones correctivas y `Like '*…*'` | SQL dinámico concatenado | `constructor.bas:5823-5875` |
| Alta/borrado | `TbNoConformidadesAuditoria` | DAO `Edit/Update` (`Borrado`, `MotivoBorrado`) | `Ops:315-333, 438-454` |
| Registro en log | `RegistrarLogAuditoria` / `LogNCProyecto` | llamada de utilidad / `Alta` | `FG.bas:1098`, `Ops:459-464` |
| Estado, fechas | `EstadoGrabar`, `FPREVCIERREGrabar`, `FECHACIERREGrabar` | clase `NCAuditoria` | `:64-67, 1102` |

Sin transacciones. El origen de los combos (`FG.bas:1877, 3075, 3137`) no se detalla aquí.

## Coexistencia

Adaptar a PostgreSQL vinculado: `Like '*…*'` con comodín de Jet (`constructor.bas:5823-5875`), `DISTINCT … .*` con `INNER JOIN` entre paréntesis (`constructor.bas:1938-1975`), `Borrado=False`, comparación de booleanos, y `Edit/Update` sobre `SELECT *` (`Ops:321`). El filtrado y el orden se hacen en VBA, no en base de datos; en una web debe pasar a consultas parametrizadas. Pendiente de la coexistencia (documento aún no redactado).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoInforme_Click` obtiene la NC con `constructor.getNCProyecto` (`:540`) en un formulario de auditoría, con `p_EsDeProyecto:=EnumSino.No` (`:549`).
- `Ops.Habilitar` actualiza con `ActualizarDatosNCProyecto` y registra con `LogNCProyecto` (`Ops:454-464`), objetos de proyecto, para una NC de auditoría.
- `Ops.Eliminar` con `p_SinDejarRastro = No` solo marca la NC (`Ops:315-333`); no consta el registro de log en esa ruta (a verificar).
- `ComandoEnviarCorreoElectronico`, `ComandoEditar`, `ComandoInforme` e `ComandoInformeLista` no comprueban `EsTecnico` (`:81-117, 318-353, 430-487, 531-567`).
- Los bucles `For i = 1 To lst.ListCount - 1` asumen una fila de cabecera (`:448, 915, 1053`).
- Actualizar una fila de la lista con `Replace` sobre el `RowSource` puede alterar otras filas iguales (`:1247, 1306`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtrado y paginación en servidor; gestionar permisos por rol de forma uniforme; sustituir `InputBox` por un diálogo con motivo obligatorio; evitar la mutación masiva de «Contador» durante la consulta.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, la lista está vacía y los botones de acción deshabilitados (`:706-714`).
2. Cada filtro, solo o combinado, restringe la lista y «Limpiar» lo deja vacío (`:772-816`).
3. «Abiertas» muestra solo NC no borradas con acciones correctivas (`constructor.bas:1938-1975`).
4. Un técnico recibe «No tiene autorización para esa acción» en Alta, «Alta para» y Eliminar/Habilitar (`:220, 367, 579`).
5. Eliminar exige confirmar y motivo no vacío; Habilitar restituye `Borrado = False` (`Ops:310-313, 438-454`).
6. Al seleccionar una NC borrada, el botón pasa a «Habilitar» (`:990-992`).
7. «Alta para …» muestra la auditoría de la NC o del filtro (`:1014-1023`).
8. Doble clic en la fila abre la NC (`:1028-1035`).

## Pendientes

- Formulario `FormCorreo`: comportamiento fuera de este lote.
- Origen de los combos (`EstablecerComboEstado`, responsables y auditorías) no detallado: `FG.bas:1877, 3075, 3137`.
- Efecto de `Ajustar Me` y de la propiedad `Estado` establecida desde otro formulario (`Form_Form0BDOpcionesAuditorias.cls:130`).
- Contenido de `HTMLNCAuditoria`, `HTMHistorialAuditoria` y `GenerarWordNoConformidades`: no revisado (`:139, 293, 467, 549`).
