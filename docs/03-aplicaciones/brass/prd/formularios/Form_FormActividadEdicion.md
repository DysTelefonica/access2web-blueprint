# Form_FormActividadEdicion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormActividadEdicion.cls` (803 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Edición de una actividad ya registrada en `TbActividades`: técnico, fecha, horas laborables y extras, tipo de actividad, ubicación, descripción y evento generado. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); lee el registro con DAO en `EstablecerDatos` y lo escribe con DAO en `Edicion`, ambas dentro del propio `.cls`. Se abre en modo lectura si la actividad está facturada o su evento franqueado. Las citas son `Form_FormActividadEdicion.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`. Capacidad de todo el formulario: CAP-4 Actividades y tipos de actividad.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormActividadGestion` (`ComandoEditarActividad_Click`) | Cierra el formulario si estaba abierto y lo abre con `OpenArgs` = ID, franqueado, facturado y URL de anexo separados por `|` | `Form_FormActividadGestion.cls:403-433` |
| `Form_FormEventoGestion` (`ComandoEditarActividad_Click`) | Igual, desde la lista de actividades del evento | `Form_FormEventoGestion.cls:1543-1570` |

`Form_Open` exige `OpenArgs` con exactamente cuatro partes, ID no vacío y facturado igual a «Sí» o «No» (`:363-385`). Solo se leen las tres primeras: `strURLAnexo` nunca recibe valor (`:373-375`, `:4`). Escritura permitida solo si ni facturada ni franqueada (`:386-390`). Salida: refresca `ListaActividades` de `FormEventoGestion` si está abierto (`:81-83`, `:292-294`); no refresca `FormActividadGestion`.

## Controles y eventos

Controles: 34 (42 `Begin` en el `.form.txt`: 34 controles + `Form` + `Section` + 6 plantillas de estilo por defecto).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `IDActividad` | texto bloqueado | — | Muestra el ID; se asigna en la carga | CAP-4 | `:451`, `.form.txt:1338-1368` |
| `ComboTecnico` | combo 3 columnas, lista de valores, solo de la lista | — | Técnico; se carga con `RellenarComboTecnicos(…, Sí)`; obligatorio por `Tag` | CAP-9 | `:402`, `:462`, `.form.txt:1396-1420` |
| `FechaAlta` | texto, fecha corta | — | Fecha de la actividad | CAP-6 | `:468`, `.form.txt:461-490` |
| `HorasLaborables`, `HorasExtras` | textos, número general | — | Horas del día; si no son numéricas se guardan como 0 | CAP-6 | `:470-475`, `:647-652`, `.form.txt:431-460, 491-521` |
| `ComboActividad` | combo 2 columnas, tabla/consulta | `OnEnter` → `Requery` | Tipo de actividad (`TbCodActividad`); obligatorio | CAP-4 | `:343-346`, `:692-695`, `.form.txt:561-600` |
| `ComboUbicacion` | combo, tabla/consulta | `OnEnter` → `Requery` | Ubicación (`TbUbicacion`) | CAP-12 | `:351-354`, `.form.txt:522-560` |
| `ComboIDEventoGenerado` | combo 2 columnas, tabla/consulta | `OnEnter` → `Requery` | Evento generado por la actividad; excluye el evento propio | CAP-3 | `:347-350`, `:517-522`, `.form.txt:601-641` |
| `Descripcion` | texto | — | Descripción; obligatoria | CAP-4 | `:639-642`, `.form.txt:279-307` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` | Valida, guarda y cierra; deshabilitado en solo lectura | CAP-4 | `:223-304`, `:499`, `:514`, `.form.txt:831-868` |
| `cmdSalir` («Salir») | botón | `OnClick` | Si hay cambios y se puede escribir, pregunta si guardar; cierra | CAP-4 | `:51-97`, `.form.txt:391-430` |
| `ComandoAdjuntar` («Adjuntar») | botón | `OnClick` | Guarda si hay cambios y abre `FormAnexos` para la actividad | CAP-11 | `:98-156`, `.form.txt:1028-1079` |
| `ImagenAnexo` | imagen | `OnClick` | Abre el anexo; con varios (`;`) delega en `ComandoAdjuntar` | CAP-11 | `:552-588`, `.form.txt:869-1027` |
| `ComandoVerEvento` (rótulo «Evento <ID>») | botón | `OnClick` | Abre o enfoca `FormEventoGestion` en el evento de la actividad | CAP-3 | `:306-342`, `:455-460`, `.form.txt:1299-1337` |
| `lblFacturado`, `lblFranqueado` | etiquetas, ocultas por diseño | — | Se muestran según los argumentos de apertura | CAP-8 | `:530-539`, `.form.txt:1080-1111, 1157-1188` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título según modo (edición, evento ya franqueado, sin evento, solo lectura); también `Caption` | CAP-4 | `:485-506`, `:406-407` |
| `Etiqueta241`, `Etiqueta209`–`Etiqueta214`, `Etiqueta11` | etiquetas | — | Rótulos de campo | sin asignar | `.form.txt:642-830, 1369-1395` |
| `Etiqueta4`–`Etiqueta8` | etiquetas «*» | — | Marcas de obligatorio (5) | sin asignar | `.form.txt:1189-1298` |
| `Numero_Actividad` | texto bloqueado y deshabilitado | — | Residuo de diseño con `ControlSource` `IDActividad` y `DCount` sobre `Tabla_Parte_Actividad`; no se usa | sin asignar | `.form.txt:308-342` |
| `CuadroSuperior`, `CuadroPie` | rectángulos | — | Decorativos | sin asignar | `.form.txt:343-390` |

## Flujos

### Apertura (`Form_Open`, `:355-421`)

1. `AjustarTamaño`, lectura de `OpenArgs` y reglas de escritura (`:361-390`).
2. `EstablecerDatos` (`:422-551`): lee `TbActividades` por ID con SQL concatenado (`:430-433`); sin fila, «No se han cargado todos los datos» (`:435-438`). Guarda los valores iniciales en variables de módulo (`:439-447`) y vuelca cada uno al control si no está vacío (`:461-484`).
3. En modo escritura habilita `AllowEdits` y `ComandoRegistrar`; en solo lectura los deshabilita (`:485-515`). Reasigna el origen de `ComboIDEventoGenerado` excluyendo el evento propio (`:517-522`) y muestra las etiquetas de facturado y franqueado (`:530-539`).
4. Carga `ComboTecnico` y copia el título (`:402-407`). Cualquier error: `MsgBox` y cierre (`:414-420`).

### Guardado (`ComandoRegistrar_Click`, `:223-304`)

1. Sin ID, «El formulario no se ha abierto con los datos suficientes» (`:231-234`). Sin cambios (`HaHabidoCambios`, `:9-49`): «No ha habido cambios que guarda» (`:240-243`).
2. Si las horas laborables suben respecto a las iniciales, suma las ya registradas ese día para ese técnico (excluida la propia actividad) y rechaza si superan 7,5 (`:252-276`).
3. `Registrar` (`:158-222`) llama a `Edicion` (`:193-195`) y, si va bien, actualiza los valores iniciales (`:205-212`). Se refresca el evento y se cierra (`:292-296`).
4. Error: `MsgBox` crítico y el formulario sigue abierto (`:298-303`).

### Salir y adjuntar

`cmdSalir` pregunta «¿Desea guardar los cambios producidos?» solo con cambios y escritura permitida (`:60-62`). `ComandoAdjuntar` guarda sin preguntar si hay cambios (`:111-122`), cierra `FormAnexos` si estaba abierto (`:137-139`), carga la actividad con `Constructor.getActividad`, la deja en `m_ObjEntidadParaAnexoActiva` (`VG.bas:252`) y abre `FormAnexos` con el nombre de este formulario (`:141-148`).

## Reglas, validaciones y cálculos

- **Descripción:** no vacía (`:639-642`). **Fecha:** válida (`:643-646`). **Tipo de actividad:** obligatorio (`:692-695`). **Evento generado:** distinto del evento de la actividad (`:633-638`).
- **Horas no numéricas:** se convierten a «0» sin avisar (`:647-652`).
- **Tope diario:** 7,5 h laborables por técnico y día, comprobado en el formulario al aumentar horas (`:252-276`); la comprobación equivalente de `Edicion` solo actúa con `strAvisoMaximasHoras = "Sí"` (`:653-673`), y `Registrar` no lo pasa (`:193-195`, `:630-632`).
- **Facturada:** si es «Sí», `Edicion` rechaza cambios de técnico, fecha y horas (`:735-749`); el estado se obtiene con `EstaFacturado("AC", ID)` cuando no llega informado (`:674-691`).
- **Vaciado:** `#` en descripción o evento generado se interpreta como «vaciar» y se guarda `Null` en `IDEVENTOGENERADO` (`:186-187`, `:777-783`).
- **Evento de la actividad:** `Registrar` reenvía el evento inicial (`:195`), de modo que la comprobación de existencia y de fecha mínima (`:707-733`) no llega a ejecutarse desde este formulario.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura inicial | `TbActividades` | SQL dinámico concatenado con el ID | — | `:430-433` |
| Combos | `TbCodActividad`, `TbUbicacion` | SQL fija en el diseño | — | `.form.txt:544, 584` |
| Evento generado | `TbEventos` | SQL dinámico concatenado con el evento inicial | — | `:517-522` |
| Horas del día | `TbActividades` | SQL dinámico con fecha `#mm/dd/yyyy#` y alias (`DameHorasTecnico`) | no | `FG.bas:549-559` |
| Estado de facturación | factura | `EstaFacturado("AC", ID)` | — | `:680` |
| Fecha del evento | `TbEventos` | `Dame` (no alcanzable, véase arriba) | — | `:717` |
| Actualización | `TbActividades` | DAO `Edit`/`Update` sobre registro leído por SQL concatenada | no | `:696-699`, `:734-784` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: literal de fecha `#mm/dd/yyyy#` (`FG.bas:551`), formato de fecha por texto `dd/mm/yyyy` (`:36`, `:453`) que depende de la configuración regional, `Nz`/`IsDate`/`IsNumeric` de VBA sobre valores textuales y `Null` explícito (`:755`, `:781`); en PostgreSQL vinculado, tipos de fecha y numérico decimal de `HorasLaborables`/`HorasExtras`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `strURLAnexo` nunca se asigna (`:4`, `:373-375`, `:523`): `ImagenAnexo` queda siempre oculta aunque el llamador envíe una URL como cuarta parte.
- En `EstablecerDatos` se asigna `Me.ComboTecnico` (`:462`) antes de que `Form_Open` cargue la lista (`:396` frente a `:402`): con un combo de lista de valores sin filas, el técnico podría no mostrarse.
- El mensaje del tope de horas usa `strHorasEscritasAhora`, que nunca se asigna (`:225`, `:268-269`): saldría vacío.
- La comprobación de tope solo se ejecuta si las horas aumentan (`:252-254`); cambiar de fecha o de técnico con las mismas horas no revalida los 7,5 h.
- Dentro de `Edicion`, los parámetros y locales reutilizan nombres de las variables de módulo (`:589-602`, `:624`); `Registrar` pasa `strIDEventoAlInicio` como `strIDEvento` (`:195`), con lo que la rama de validación de evento nunca se activa.
- `Edicion` abre `.Edit` y puede lanzar el error de facturada antes de `.Update` sin `CancelUpdate` (`:734-749`); el cierre del `Recordset` en el manejador lo descarta (`:794-797`).
- Con `strEstaFacturado = "Sí"` se comparan fecha y horas textuales con los valores del registro (`:740-741`), sin normalizar formato: podría rechazar por formato aunque no haya cambio real. Con `AllowEdits = False` (`:506`) el caso no debería darse.
- Ausencia de `ubicación` o `horas` vacías: la ubicación vacía se ignora y conserva la anterior (`:771-773`); las horas vacías se guardan como 0 (`:647-649`, `:765-766`).
- Texto «No ha habido cambios que guarda» con errata (`:241`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: validar el tope diario también al cambiar fecha o técnico; guardar con control optimista de concurrencia; tipos numéricos explícitos para horas; mostrar el anexo realmente asociado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir con una actividad facturada o de evento franqueado deja todo en solo lectura, con `Registrar` deshabilitado y las etiquetas correspondientes visibles.
2. Guardar sin cambios muestra «No ha habido cambios que guarda» y no escribe.
3. Guardar con descripción vacía, fecha no válida o tipo de actividad vacío muestra el mensaje correspondiente y no escribe.
4. Subir las horas laborables de modo que el total del día del técnico supere 7,5 h muestra el aviso y no escribe.
5. Salir con cambios pregunta si guardar; «No» cierra sin escribir.
6. «Adjuntar» con cambios los guarda antes de abrir `FormAnexos`.
7. Tras guardar, la lista de actividades del evento abierto se refresca.

## Pendientes

- Estructura de `TbActividades`, incluidos tipos de `HorasLaborables`, `HorasExtras` y `FechaAlta` (ERD de origen).
- Contenido de `EstaFacturado` (`FG.bas:5758`) y de `Dame` (`FG.bas:290`).
- Efecto de que `m_SQL`, `pregunta`, `flag` y `dato` sean globales compartidas (`VG.bas:203-206`) y se usen aquí (`:71`, `:117`, `:433`): una llamada anidada a otro módulo podría pisarlas.
- Contenido de `AjustarTamaño` (`:361`) y de `AbrirEnLocal` (`:572`).
- Qué guarda `strURLAnexo` en los llamadores (`Form_FormActividadGestion.cls:432`, `Form_FormEventoGestion.cls:1569`).
