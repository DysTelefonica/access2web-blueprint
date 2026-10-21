# Form_FormTareas

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormTareas.cls` (470 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Bandeja de «Tareas pendientes»: una lista de seis tipos de tarea con su recuento y, al elegir uno, la lista de expedientes que lo cumplen, con acceso a la ficha del expediente y a un informe HTML del tipo. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); los datos llegan de colecciones en caché de `m_ObjEntorno` y no escribe en la base de datos. Las citas son `Form_FormTareas.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpciones` | `DoCmd.OpenForm "FormTareas"` desde las imágenes `ImagenTareasOFF` e `ImagenTareasON` | `Form_Form0BDOpciones.cls:294, 321` |
| `Form_FormExpedientesGestion` | igual, desde sus imágenes de tareas | `Form_FormExpedientesGestion.cls:1823, 1851` |
| `PintarTareas` (cualquier llamador) | si el formulario está abierto, rehace `ListaTipoTareas` y llama a `ListaTipoTareas_Click` | `FUNCIONES UTILES.bas:4047-4062` |

Abre `FormExpediente` sin `OpenArgs` (`:184`), por lo que la ficha entra en modo edición si el usuario es administrador (`Form_FormExpediente.cls:121-125`). Abre el informe HTML del tipo con `ShellExecute` (`HTMLENTXT`, `FUNCIONES UTILES.bas:4233-4234`). Sale con `cmdSalir` (`:12-15`).

Estado público: `m_ExpSeleccionado` y `m_TipoTareaSeleccionada` (`:6-7`), que `PintarTareas` reinicia desde fuera (`FUNCIONES UTILES.bas:4059-4060`). Escribe las globales `m_ObjExpedienteActivo`, `m_ObjExpedienteDTOActivo` y `m_ObjExpedienteAlInicio` antes de abrir la ficha (`:169-176`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaTipoTareas` | lista 2 columnas (ID oculto) | `OnClick` → `ListaTipoTareas_Click` (público) | Muestra la explicación del tipo y filtra si el tipo cambia | `:276-336`, `.form.txt:715-731` |
| `Explicacion` | texto bloqueado | — | Texto de `colExplicacionTipoTarea` del tipo elegido | `:309-310`, `.form.txt:752-766` |
| `ListaFiltrados` | lista 6 columnas con encabezados (ID oculto) | `OnClick` → `ListaFiltrados_Click` | Carga el expediente de la fila y habilita los botones | `:402-469`, `.form.txt:596-625` |
| `ListaFiltrados` | — | `OnDblClick` | Si `ComandoDetalle` está habilitado, lo ejecuta | `:266-274` |
| `ComandoDetalle` («Ver Detalle») | botón | `OnClick` | Prepara las globales del expediente y abre `FormExpediente` | `:151-203` |
| `ComandoVerInforme` (rótulo «Comando36», icono de lupa) | botón | `OnClick` | Genera y abre el informe HTML del tipo seleccionado | `:205-264`, `.form.txt:797-801` |
| `ComandoActualizar` («Ver Todos») | botón | `OnClick` | Vacía cinco cachés de tareas y repinta | `:94-121`, `.form.txt:861-865` |
| `ComandoAyuda` (rótulo «Salir», información «Ayuda») | botón | `OnClick` | `AbrirAyuda` | `:124-149`, `.form.txt:504-509` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `:12-15` |
| `lblListaDeResultados` | etiqueta | — | «Registros filtrados (n)» | `:33, 78` |
| `lblTitulo`, `lblEstado` | etiquetas | — | El título pasa a `Caption` del formulario (`:353`); `lblEstado` está oculta y el código no la usa | `.form.txt:483-484, 960-971` |

Los `RowSource` de diseño de ambas listas (`.form.txt:613-621, 727-728`) se sustituyen en ejecución (`:31`; `FUNCIONES UTILES.bas:4050-4057`). El de `ListaFiltrados` conserva datos reales residuales del diseño; no se reproducen aquí.

## Flujos

### Apertura (`Form_Open`, `:340-400`)

1. `Ajustar Me`, `AllowEdits = False`, oculta `Explicacion` y copia el título (`:350-353`).
2. `PintarTareas EnumSiNo.No` calcula los seis recuentos (cargando las cachés si están vacías), rellena `ListaTipoTareas` con «<rótulo> (<n>)» y llama a `ListaTipoTareas_Click` (`:354`; `FUNCIONES UTILES.bas:4009-4030, 4047-4061`).
3. Preselecciona la primera fila cuya colección no sea `Nothing`, en el orden Estado desconocido → Recepción completa → Recepción de hito → `AdjudicadosTSOLSinCodS4H`, y lanza `ListaTipoTareas_Click` (`:359-377`).
4. Recorre `ListaTipoTareas` y vuelve a lanzar el clic por cada fila seleccionada (`:378-382`).
5. Restablece `AllowEdits = True` (`:387`).

### Selección de tipo (`ListaTipoTareas_Click`, `:276-336`)

1. Si el tipo de la fila es distinto de `m_TipoTareaSeleccionada`, lo guarda y salta a `Ejecutar` (`:294-297`); si es el mismo, no hace nada (`:298-307`). Sin fila numérica también salta a `Ejecutar` (`:290-292`).
2. `Ejecutar`: escribe la explicación y llama a `Filtrar` (`:309-312`).
3. `Filtrar` reinicia la lista con la fila de encabezados, deshabilita `ComandoDetalle` y `ComandoVerInforme` y toma la colección del tipo (`:29-59`). Añade una fila por expediente con `IDExpediente;ESTADOCalculadoTexto;TipoCalculadoTexto;Nemotecnico;CodExp;Titulo` (`:69-77`), actualiza el recuento (`:78`) y, si hay filas, selecciona la primera y ejecuta `ListaFiltrados_Click` (`:79-83`).

### Selección de expediente (`ListaFiltrados_Click`, `:402-469`)

Deshabilita los dos botones; sin ID deja `m_ExpSeleccionado = Nothing` (`:414-425`). Si no hay expediente cargado o su ID difiere, lo relee con `constructor.getExpediente` (`:426-440`). Con expediente, habilita `ComandoDetalle` y `ComandoVerInforme` (`:449-450`).

### Ver detalle (`ComandoDetalle_Click`, `:151-203`)

1. Sin `m_ExpSeleccionado`, lo lee de la fila; si sigue vacío, error «Se ha de seleccionar un Elemento de la lista» (`:159-167`).
2. Fija `m_ObjExpedienteActivo`, construye el DTO con `EstablecerDTODeExpediente` (`p_ParaCopia = No`) y relee de la base el expediente inicial (`:169-179`; `FUNCIONES UTILES.bas:2105`).
3. Cierra `FormExpediente` con `acSaveNo` si estaba abierto, lo abre y se suscribe a sus eventos con `WithEvents` (`:181-187`).

### Ver informe (`ComandoVerInforme_Click`, `:205-264`)

Sin tipo, «Seleccione un tipo de tarea» (`:213-216`). Según el tipo llama a una de las seis `getURLInforme…` (`:217-246`), que componen el HTML desde la misma caché de `Entorno`, lo guardan en un fichero y lo abren (`FUNCIONES UTILES.bas:4347-4527, 4795`). Tipo no previsto: «Tipo de tarea no codificada» (`:247-249`).

### Actualizar (`ComandoActualizar_Click`, `:94-121`)

Pone a `Nothing` cinco colecciones de `m_ObjEntorno` (`:102-108`) y llama a `PintarTareas` (`:109`), que las vuelve a consultar al contarlas.

### Errores

El patrón general acumula el texto en `m_Error`, envía `CorreoAlAdministrador` salvo el error interno 1000 y muestra `MsgBox` crítico (p. ej. `:195-202`). `Form_Open`, además, cierra el formulario (`:399`). `ComandoActualizar_Click` no envía correo (`:115-120`). `ListaTipoTareas_Click` y `ListaFiltrados_DblClick` usan `On Error Resume Next` (`:280, 267`).

## Reglas, validaciones y cálculos

Tipos (`EnumTipoTarea`, `Variables Globales.bas:211-218`), con su consulta en `constructor.bas`, todas sobre `TbExpedientes`:

| Tipo | Condición | Evidencia |
|---|---|---|
| 1 Estado desconocido | `APLICAESTADO = 'Sí'` y `ESTADO = 'Desconocido'` | `constructor.bas:1481-1485` |
| 2 A punto de recepcionar | `FechaFinContrato` entre hoy y hoy + `DiasParaAvisoFinExpediente` (exclusivo) y `EsBasado = 'Sí'` o `EsExpediente = 'Sí'` | `constructor.bas:1584-1590` |
| 3 Hito a punto de recepcionar | `FechaHito` de `TbExpedientesHitos` en la misma ventana | `constructor.bas:1637-1642` |
| 4 Adjudicado sin contrato | sin `FechaInicioContrato`, `GARANTIAMESES` ni `FechaFinContrato`, con `FECHAADJUDICACION` y `APLICAESTADO <> 'No'` | `constructor.bas:1690-1693` |
| 5 Adjudicado sin CodS4H | `CadenaContratistas` de `TbExpedientesConEntidades` igual a un contratista fijo, `Adjudicado = 'Sí'`, `CodS4H` nulo y `AplicaTareaS4H <> 'No'` | `constructor.bas:1742-1748` |
| 6 Oferta sin decidir | con `FECHAOFERTA`, sin `FECHAPERDIDA`, `FECHADESESTIMADA` ni `FECHAADJUDICACION`, y oferta de hace `DiasParaOfertasSinDecision` días o más | `constructor.bas:1797-1802` |

- `DiasParaAvisoFinExpediente` está fijado a «15» en código (`Entorno.cls:151`); `DiasParaOfertasSinDecision` sale de `TempVars` (`Entorno.cls:172`).
- Cada colección es un diccionario por `IDExpediente` sin duplicados (p. ej. `constructor.bas:1611-1613`); una consulta sin filas devuelve `Nothing` (`constructor.bas:1593-1596`).
- Las cachés viven en `Entorno` y solo se recalculan al ponerlas a `Nothing` (`Entorno.cls:1256-1260`): lo hacen `ComandoActualizar` (`:103-107`) y `PintarTareas EnumSiNo.Sí` (`FUNCIONES UTILES.bas:4000-4007`), que `Registrar` invoca solo si cambia el paso por «Desconocido» o `NecesitaCodS4H` (`ExpedienteOperaciones.cls:542-551`).
- Las explicaciones de cada tipo son textos fijos de `Entorno` (`Entorno.cls:1618-1623`).
- No hay comprobación de rol en este formulario; la edición la decide `FormExpediente`.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Colecciones de tareas | `TbExpedientes`, `TbExpedientesHitos`, `TbExpedientesConEntidades`, `TbUsuariosAplicaciones` | SQL dinámico DAO (`getdb().OpenRecordset`) | `constructor.bas:1487, 1591, 1644, 1695, 1750, 1804` |
| Expediente de la fila e inicial | vía `constructor.getExpediente` | DAO (no trazado aquí) | `:176, 427`; `constructor.bas:1134` |
| Informe HTML | fichero temporal en disco | `ADODB.Stream.SaveToFile` | `FUNCIONES UTILES.bas:4220, 4231` |

Sin escrituras en la base de datos.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `DateDiff('d', Date(), …)` y literales `'Sí'` de Jet en todas las consultas de tareas (`constructor.bas:1586-1590, 1641-1642, 1802`); umbrales concatenados en el SQL; la tarea 5 depende de la caché `TbExpedientesConEntidades`, cuya actualización hace `RegistrarExpEntidades`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La preselección comprueba `ColExpedientesAdjudicadosTSOLSinCodS4H` pero selecciona la fila 3, que es «Adjudicados sin contrato» (`:371-372`; orden en `FUNCIONES UTILES.bas:4052-4057`); nunca considera los tipos 4 y 6.
- El bucle `For i = 0 To ListCount` llega a un índice fuera de la lista (`:378`); si `Selected` falla ahí, el manejador de `Form_Open` cerraría el formulario (`:389-399`).
- `ListaTipoTareas_Click` tiene `On Error Resume Next` (`:280`): el `Err.Raise 1000` de `:314` no salta a `errores`, así que los fallos de `Filtrar` se silencian y su `MsgBox` es inalcanzable.
- `ComandoActualizar_Click` no vacía `ColExpedientesFaseOfertaPorMuchoTiempo` (`:103-107`): la tarea 6 no se refresca con «Ver Todos».
- Los errores de `PintarTareas` se pierden: se llama sin `p_Error` (`:109, 354`).
- El mensaje de error de `ComandoVerInforme_Click` dice «ComandoDetalle_Click» (`:258`).
- `ComandoVerInforme` informa de un tipo entero pero solo se habilita con un expediente seleccionado (`:35-36, 449-450`); su rótulo de diseño es «Comando36» y `ComandoAyuda` rotula «Salir» (`.form.txt:505, 798`).
- `m_FormExpediente` se declara `WithEvents` (`:5, 186`) pero no hay handlers: guardar desde la ficha no refresca la bandeja salvo por el `PintarTareas` condicional de `Registrar` (`ExpedienteOperaciones.cls:542-551`); p. ej. completar las fechas de contrato no lo dispara.
- Cuando `PintarTareas` repinta con el formulario abierto, reinicia el tipo seleccionado y deja la lista de expedientes vacía (`FUNCIONES UTILES.bas:4050, 4059-4061`; `:38-40`).
- «Ver Detalle» cierra con `acSaveNo` una ficha de expediente ya abierta (`:181-182`), descartando lo no guardado.
- Las filas se añaden con `AddItem` separando por «;» (`:73`): un título con «;» o comas desplazaría columnas; además, la lista de valores tiene límite de longitud.
- La consulta de hitos hace un `LEFT JOIN TbUsuariosAplicaciones` sin uso (`constructor.bas:1640`); la de ofertas exige fila en `TbExpedientesConEntidades` sin motivo aparente (`constructor.bas:1798-1799`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: recuentos calculados en servidor sin caché de sesión; umbrales configurables en lugar del «15» fijo; informe por tipo independiente de la fila seleccionada.

## Criterios de aceptación

Propuestos, no aprobados:

1. La bandeja muestra los seis tipos con su recuento, que coincide con las condiciones de la tabla de tipos.
2. Elegir un tipo muestra su explicación y la lista de expedientes con estado, tipo, nemotécnico, código y título, y el recuento «Registros filtrados (n)».
3. Con al menos un expediente, el primero queda seleccionado y «Ver Detalle» abre su ficha; el doble clic hace lo mismo.
4. «Ver Detalle» sin selección muestra «Se ha de seleccionar un Elemento de la lista».
5. El informe de un tipo abre un documento con los mismos expedientes que la lista.
6. «Actualizar» recalcula los seis tipos, incluida la oferta sin decidir.

## Pendientes

- Si `FormularioAbierto("FormTareas")` es verdadero durante `Form_Open`, y por tanto si `PintarTareas` rellena la lista en la apertura (`:354`; `FUNCIONES UTILES.bas:2825, 4047`).
- Comportamiento de `Selected(ListCount)` en Access (`:378-379`).
- Clave buscada en `colExplicacionTipoTarea` sin tipo (`m_TipoTareaSeleccionada = 0`, `:310`) y si el diccionario la añade vacía.
- Si un informe sin filas falla con `m_Col.Count` sobre `Nothing` (`FUNCIONES UTILES.bas:4795-4796`).
- Quién asigna `TempVars("DiasParaOfertasSinDecision")` (`Entorno.cls:172`) y el contenido de `Ajustar`, `AbrirAyuda` y `DameUntxtYHtml`.
- Contenido de `EstablecerDTODeExpediente` (`FUNCIONES UTILES.bas:2105`) y de `constructor.getExpediente` (`constructor.bas:1134`).
