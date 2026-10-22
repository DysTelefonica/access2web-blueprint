# Form_FormSolicitudRenovacionDatos

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudRenovacionDatos.cls` (957 líneas) y `.form.txt` (63 controles `Begin`) · **PRD:** [README](../README.md)

Subformulario de la pestaña «Datos» de la renovación: presenta los campos de la solicitud (titular, empresas, expediente, motivo, contacto), detecta si algo ha cambiado y los persiste sobre `TbSolicitudes`. Es el **único** de los cuatro formularios de este lote que escribe en la cabecera de la solicitud. Es **no enlazado** (sin `RecordSource`): los controles son cajas de texto y combos sin origen de datos. Las citas son `Form_FormSolicitudRenovacionDatos.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormSolicitudRenovacion` | `Me.FrmDetalle.SourceObject = "FormSolicitudRenovacionDatos"` al montar la primera pestaña | `Form_FormSolicitudRenovacion.cls:82-83` |
| `FormEnvioCorreo` | El envío de correo llama a su `PintarIndicadores`; no entra aquí | `Form_FormEnvioCorreo.cls:277-289` |

Al terminar de guardar, el subformulario decide qué contenedor refrescar según `Me.Parent.Name`: si es el de renovación, vuelve a invocar `EstablecerDatos` del padre (`:857-860`). No expone eventos propios; usa `WithEvents` para tres hijos (`:5-7`). El contexto llega por globales: `m_ObjSolicitudActiva`, `m_ObjSolicitudAlInicio`, `m_ObjUsuarioHPSActivo`, `m_ObjUsuarioConectado` (`VG.bas:197-205`).

## Controles y eventos

Solo los controles con comportamiento propio; el resto son rótulos y cajas sin evento.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComandoRegistrar` | botón «Editar» | `OnClick` | Guarda los datos y ejecuta la lógica de correos automáticos | `:656-746`, `.form.txt:1475-1477` |
| `PalabraClave` | texto | `OnKeyDown` | Con Enter o Tab dispara la búsqueda de expediente | `:259-276`, `.form.txt:2194` |
| `ComandoBuscarExpediente` | botón | `OnClick` | Resuelve un expediente por identificador o por palabra clave | `:52-150`, `.form.txt:2276-2278` |
| `ComandoIrABusquedaExpedientes` | botón | `OnClick` | Abre el buscador general y le pasa la palabra clave | `:16-50`, `.form.txt:2354-2356` |
| `ComandoVerDetalleExpediente` | botón | `OnClick` | Abre `FormExpedienteDetalle` para el expediente activo | `:183-218`, `.form.txt:2545-2547` |
| `ComandoMotivos` | botón «Editar» | `OnClick` | Da de alta un motivo nuevo y lo añade al combo | `:152-181`, `.form.txt:1656-1658` |
| `ComandoNuevoResponsable` | botón | `OnClick` | Abre `FormResponsable` en alta y recarga el combo | `:298-325`, `.form.txt:1917-1919` |
| `ComandoCambioTipo` | botón | `OnClick` | Abre `FormCambioTipo` y cierra el padre para reabrir con el tipo nuevo | `:907-934`, `.form.txt:2098-2101` |
| `ComandoActualizar` | botón | `OnClick` | Repinta el rótulo de estado calculado | `:278-296`, `.form.txt:2039-2041` |
| `ComandoActualizarar` | botón «Contador» | `OnClick` | Invalida la caché de expedientes del entorno | `:12-14`, `.form.txt:2414-2416` |
| `Gestor` | combo | — | Gestor tramitador; origen de diseño con lista literal de nombres | `.form.txt:1549` |
| `CorreosAutomaticos` | combo | — | Sí/No de envío automático; origen de diseño con dos literales | `.form.txt:826` |
| `IDEmpresaUsuario` / `IDEmpresaTramitadora` | combo | — | Empresa del usuario y empresa tramitadora; listas literales en diseño | `.form.txt:741, 1324` |
| `Motivo_HPS` | combo | — | Motivo de la solicitud; lista literal en diseño | `.form.txt:1620` |
| `emailResponsable` | combo | — | Correo del responsable; se rellena desde `m_ObjEntorno.Responsables` | `:795-806`, `.form.txt:1885` |
| `IDSolicitud` / `IDUsuarioHPS` | texto deshabilitado | — | Identificadores, no editables | `.form.txt:1834, 2640` |
| `TIPO` | texto bloqueado | — | Tipo de la solicitud | `.form.txt:2000` |
| `lblEstadoSolicitud` | etiqueta | — | Estado calculado; oculta en alta en frío | `:443, 511-512`, `.form.txt:1783-1784` |

Las listas de diseño (`Gestor`, `Motivo_HPS`, las dos empresas) contienen valores reales de la instalación. Se citan control y línea; **no se transcriben aquí**.

## Flujos

### Carga (`Form_Load` → `EstablecerDatos`, `:327-357`, `:359-521`)

1. `EstablecerCombos` puebla los cuatro combos desde las cachés del entorno (`:363-365`, `:747-823`).
2. Comprueba que hay solicitud activa (`:367-370`). Si la solicitud tiene `IDSolicitud`, relee la cabecera y la fija como foto de comparación `m_ObjSolicitudAlInicio` (`:371-377`).
3. Sin foto previa (**alta en frío**), copia los datos del usuario HPS seleccionado a los campos —nombre, apellidos, fecha y lugar de nacimiento, correo, teléfono, empresa, expediente y motivo (`:380-443`)— deshabilita el cambio de tipo, oculta el estado y sale.
4. Con foto previa, vuelca los campos desde la solicitud guardada, habilita `ComandoCambioTipo` salvo si el tipo es «Traspaso» (`:446-509`), muestra el estado calculado y memoriza el valor inicial de `CorreosAutomaticos` (`:514`).

### Búsqueda de expediente (`ComandoBuscarExpediente_Click`, `:52-150`)

1. Si la palabra clave está vacía, no hace nada (`:64-69`).
2. Si es numérica, la trata como identificador de expediente y, si existe, la resuelve en un paso (`:71-88`).
3. Si no, busca por texto; si no hay resultados, vacía el expediente y delega en el buscador general (`:90-107`); si hay un único resultado, lo asigna directamente (`:109-122`); si hay varios, abre el buscador con la palabra clave (`:124-138`).
4. Cuando el buscador devuelve una selección, el evento `m_FormExpedientes_Seleccionado` rellena identificador y título, habilita o deshabilita los botones de detalle y limpia la palabra clave (`:220-256`).

### Guardado (`ComandoRegistrar_Click` → `Registrar`, `:656-746`, `:826-871`)

1. `Registrar` llama a `HaHabidoCambios`; si no hay cambios, aborta con «No hay cambios que guardar» (`:836-841`).
2. `RellenarObjeto` vuelca los controles sobre `m_ObjSolicitudActiva` con `Nz` en todos los campos (`:614-654`).
3. `SolicitudOperaciones.RegistrarDatos` valida, calcula el estado y escribe `TbSolicitudes` (`:843-848`; `SolicitudOperaciones.cls:42-118`).
4. Se relee la solicitud guardada y se refresca el padre (`:855-859`).
5. Con la configuración de correos automáticos activa, `ComandoRegistrar_Click` analiza el paso de «No» a «Sí»: si el proceso ya ha empezado (hay fecha de envío de Excel o de alta en Marga) pregunta si hay que reprogramar los avisos (`:664-694`). En el paso de «Sí» a «No» llama a `AnularTodosLosProgramados` (`:697-708`).
6. Avisa de que los datos se han grabado y actualiza la foto de comparación (`:713-736`).

### Errores

Los handlers capturan, envían el aviso al administrador con `CorreoAlAdministrador` cuando el error no es el 1000 interno y muestran `MsgBox` (`:38-49, 130-149, 176-181, 205-218, 243-256, 311-325, 347-357, 862-871, 896-905, 923-934, 946-957`). `ComandoActualizarar_Click` es la excepción: no notifica por correo (`:290-295`).

## Reglas, validaciones y cálculos

- **Guardado condicional:** `Registrar` aborta si `HaHabidoCambios` es falso, salvo que se invoque con `EnumSiNo.No` (`:829-842`).
- **Validación de cabecera:** `MotivoNoOK` exige gestor, nombre, primer apellido, correo con formato válido, correo del responsable con formato válido si viene informado, expediente con contratista principal, indicador de correos automáticos y empresa tramitadora registrada; el DNI, si viene, se valida por formato (`SolicitudOperaciones.cls:343-420`).
- **Estado calculado:** el estado no se teclea, se deriva de los datos antes de escribir (`:846`; `SolicitudOperaciones.cls:62-65`).
- **Identidad del titular:** en alta en frío el titular es el usuario HPS activo; al guardar, `IDUsuarioHPS` se fuerza al del usuario HPS activo (`:640-643`).
- **Detección de cambios:** se comparan uno a uno los dieciséis campos del formulario contra la foto `m_ObjSolicitudAlInicio` (`:524-611`). El título del expediente, el identificador de solicitud y `CorreosAutomaticos` en su valor inicial no participates de esa comparación más allá del caso indicado.
- **Reprogramación condicionada:** pasar de «No» a «Sí» solo ofrece reprogramar si ya existe una fecha de envío de Excel o de alta en Marga (`:670-680`); pasar a «No» limpia siempre la cola (`:699-708`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Alta o actualización de la solicitud | `TbSolicitudes` (SQL concatenado con `IDSolicitud`) | DAO | `SolicitudOperaciones.cls:70-74` |
| Relectura de la cabecera | `TbSolicitudes` (vía `constructor.getSolicitud`) | DAO | `Form_FormSolicitudRenovacionDatos.cls:373, 856` |
| Expediente por identificador | `TbExpedientes` (vía `constructor.getExpediente`) | DAO | `Form_FormSolicitudRenovacionDatos.cls:72` |
| Expedientes por palabra clave | `TbExpedientes` (vía `constructor.getExpedientesBusqueda`) | DAO | `Form_FormSolicitudRenovacionDatos.cls:89` |
| Empresa tramitadora registrada | `TbSuministradores` (vía `constructor.getSuministrador`) | DAO | `SolicitudOperaciones.cls:401` |
| Listas de combos | cachés del entorno (`SumTramitadoras`, `Suministradores`, `Responsables`) y motivos | Caché en memoria | `Form_FormSolicitudRenovacionDatos.cls:747-823` |

## Coexistencia

Pendiente del capítulo `coexistencia.md`, que todavía no existe en este PRD. Lo que ya se ve: el guardado se hace con SQL concatenado (`SolicitudOperaciones.cls:70-74`), patrón que no se traslada tal cual; la comparación de «cambios» por campos es una técnica de formulario, no de servidor; y las listas de empresa, gestor y motivo están embebidas en el diseño del formulario, así que en web deben salir de tablas de referencia.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El guardado escribe concatenando el `IDSolicitud` directamente en el `WHERE` (`SolicitudOperaciones.cls:70-74`). El valor viene de la base de datos, no de un campo editable, así que no es una vía de inyección evidente, pero no hay comprobación de que sea numérico antes de construir la sentencia.
- `RellenarObjeto` fija `IDUsuarioHPS` dos veces seguidas: primero al usuario HPS activo y acto seguido al valor del control, de modo que la primera asignación es código muerto (`:640-643`).
- La rama «sin resultados» de la búsqueda de expediente vacía `IDExpediente` a `Null` y acto seguido lo comprueba contra cadena vacía, así que la condición es siempre cierta y `ComandoActualizar` queda deshabilitado sin motivo real (`:98-104`).
- `ComandoActualizar_Click` pinza `Err.Number` en un literal de texto, igual que el `MsgBox` de `Form_FormSolicitudRenovacion` (`:291`).
- `ComandoActualizarar_Click` es la única exclusión del patrón: no envía correo al administrador cuando falla (`:288-295`), y su nombre parece una errata heredada (`.form.txt:2414`).
- `HaHabidoCambios` compara campo a campo con dieciséis `If` encadenados (`:524-611`): añadir una propiedad al formulario sin tocar esta función hace que el cambio no se detecte, aunque sí se guarde si otro campo lo dispara.
- `Registrar` relee la solicitud recién guardada con `constructor.getSolicitud` sin pasar `p_Error` (`:856`): un fallo de lectura deja la caché con `Nothing` sin avisar.
- `ComandoCambioTipo_Click` cierra el formulario padre (`:930`) y `m_FormCambioTipo_TipoCambiado` vuelve a cerrarlo (`:941`): la segunda llamada se ejecuta sobre un formulario ya cerrado.
- El refresco del padre se decide comparando `Me.Parent.Name` con dos literales (`:857-859`); `Form_FormEnvioCorreo.cls:287` usa el mismo patrón de llamada por nombre de módulo bajo `On Error Resume Next`, de modo que cualquier fallo de resolución de nombre pasa inadvertido.
- `ComandoMotivos_Click` escribe el motivo que el usuario teclee sin hygienicación antes de meterlo en el combo (`:158-162`).
- El `Caption` de diseño (`.form.txt:39`) fija versión y nombre de una persona como metadato residual.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: formulario con detección de cambios por campo modificado en servidor; combos servidos desde tablas de referencia en lugar de listas de diseño; alta de motivo como formulario propio; identificadores de expediente y empresa validados en servidor y no por el cliente.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir sin cambios pendientes y pulsar «Editar», el formulario avisa de que no hay nada que guardar.
2. Un correo con formato inválido impide el guardado y explica el motivo.
3. Un DNI con formato incorrecto impide el guardado.
4. Buscar un expediente por identificador numérico rellena el identificador y el título en un solo paso.
5. Buscar por una palabra clave sin resultados vacía el expediente y abre el buscador general con esa palabra clave.
6. Seleccionar un expediente en el buscador general rellena los campos y habilita la consulta de detalle.
7. Dar de alta un motivo nuevo lo añade al combo y lo deja seleccionado.
8. Cambiar el tipo de la solicitud reabre el contenedor con el tipo nuevo.
9. Pasar de «No» a «Sí» los correos automáticos con el proceso ya iniciado ofrece reprogramar los avisos; sin proceso iniciado no pregunta.
10. Pasar de «Sí» a «No» elimina los envíos programados y lo confirma.
11. Tras guardar, el título de última modificación del contenedor refleja el nuevo dato.

## Pendientes

- Contenido de `constructor.getExpedientesBusqueda`: si filtra por cualquier columna del expediente o solo por una (`:89`).
- Si `MeterUnMotivoEnListaMOtivos` valida duplicados o longitud antes de escribir (`FU.bas:1442`).
- Reglas de `EsDNIValido` y `EsCorreoValido`, y si se aplican también en web (`SolicitudOperaciones.cls:378-381`).
- Qué ocurre con el identificador de solicitud si la escritura de `TbSolicitudes` se resuelve en PostgreSQL sin `IDSolicitudCalculado` (`SolicitudOperaciones.cls:75`).
- Si `ComandoNuevaEmpresaUsuario` (`.form.txt:1386-1387`) tiene handler: el `.cls` no lo declara, igual que `ComandoActualizarar`.