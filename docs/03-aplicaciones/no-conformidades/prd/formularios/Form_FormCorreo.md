# Form_FormCorreo

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormCorreo.cls` (156 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana de envío del informe de una NC por correo. Muestra asunto, destinatarios («PARA») y copia («CC.») ya rellenados, y al pulsar «Enviar correo» compone el informe HTML de la NC activa y lo **registra** en la tabla de correos pendientes; el envío real no lo hace este formulario. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); trabaja sobre las variables globales `m_ObjNCProyectoActiva` / `m_ObjNCAuditoriaActiva`. Las citas son `Form_FormCorreo.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_FormNCProyectoGestion | Fija `m_ObjNCProyectoActiva`, cierra un `FormCorreo` previo y abre con `OpenArgs = EnumSino.Sí` | `Form_FormNCProyectoGestion.cls:86-89` |
| Form_FormNCAuditoriaGestion | Ídem con `m_ObjNCAuditoriaActiva` y `OpenArgs = EnumSino.No` | `Form_FormNCAuditoriaGestion.cls:98-101` |

`OpenArgs` es obligatorio: `EnumSino.Sí = 1`, `No = 2` (`constantes.bas:27-30`); vacío ⇒ error de parámetros (`:120-124`). No abre otros formularios; se cierra sin guardar tras enviar o con «Salir» (`:10, 67`).

## Controles y eventos

`.form.txt` declara 24 bloques `Begin`: el formulario (`.form.txt:4`), 11 bloques de valores por defecto (`.form.txt:406-477`), 3 secciones (`EncabezadoDelFormulario` `.form.txt:482`, `Detalle` `.form.txt:511`, `PieDelFormulario` `.form.txt:688`), 3 etiquetas y 3 cuadros de texto de datos, el título y 2 botones.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` («CORREO») | etiqueta | — | Título; se copia a la barra de título | `.form.txt:500-501`, `:99` |
| `Asunto` (etiqueta «ASUNTO») | cuadro de texto | — | Asunto editable, precargado | `.form.txt:533, 564`, `:85-88` |
| `DESTINATARIOS` («PARA») | cuadro de texto | — | Destinatarios editables, precargados | `.form.txt:589, 620`, `:90-97` |
| `DestinatariosConCopia` («CC.») | cuadro de texto | — | Copia editable; se precarga con el usuario conectado | `.form.txt:645, 676`, `:98` |
| `ComandoEnviarCorreo` («Enviar correo») | botón | `OnClick` | Valida, genera HTML y registra el correo | `:14-77`, `.form.txt:812-814` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | `:8-11`, `.form.txt:707-709` |
| Formulario | — | `OnOpen` | Valida parámetros y precarga datos | `:111-156`, `.form.txt:42` |

## Flujos

### Apertura (`Form_Open`, `:111-156`)

1. Reloj de arena, `Ajustar Me` (`:115, 119`; `FG.bas:615`), y lee `OpenArgs` en `m_EsDeProyecto` (`:120`).
2. Error si falta el argumento o si el objeto de NC del tipo indicado es `Nothing` (`:121-135`).
3. `EstablecerDatos` rellena (`:81-108`): **asunto** «Informe de NC de proyecto» + código (`:85`), o «Informe de» + partícula + nombre de auditoría + número a dos cifras (`:87-88`); **PARA**: correo del responsable de Telefónica de la NC de proyecto, si existe (`:91-93`), o, en auditoría, la lista de Calidad de `TempVars("CadenaCorreosCalidadEnPruebas")` (`:96`); **CC.**: correo del usuario conectado (`:98`).
4. Cierra el avance (`AvanceCerrar`, `:141`; `FG.bas:1422-1438`).

### Envío (`ComandoEnviarCorreo_Click`, `:14-77`)

1. Valida asunto no vacío y al menos un destinatario en PARA o CC. (`:30-39`).
2. Genera el cuerpo con `HTMLNCProyecto` o `HTMLNCAuditoria`, con acciones y no solo tablas (`:40-50`; `HTML.bas:784, 865`).
3. Crea un `Correo` con asunto, destinatarios, copia y cuerpo, y llama a `Registrar` (`:51-61`).
4. Aviso «Correo ordenado para su envío» y cierre (`:66-67`).

### Errores

Error 1000 ⇒ solo `MsgBox` con `m_Error`; otro error ⇒ mensaje con número y detalle, `CorreoAlAdministrador` y `MsgBox` crítico (`:70-76`). En `Form_Open` cualquier error muestra el aviso y cierra el formulario (`:147-155`).

## Reglas, validaciones y cálculos

- Asunto obligatorio; PARA y CC. no pueden estar ambos vacíos (`:30-39`); `Correo.Registrar` repite estas comprobaciones y exige cuerpo no vacío (`Correo.cls:103-114`).
- El campo de copia oculta se rellena en `Registrar` con una dirección fija del código cuando llega vacío (`Correo.cls:115-117`); el formulario no lo muestra.
- El originador es el usuario de red conectado y la fecha, `Now` (`Correo.cls:118-119`).
- No valida el formato de las direcciones ni el rol del usuario.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Identificador del correo (`Max + 1`) | `TbCorreosEnviados` | SQL dinámico concatenado (`SELECT Max(...)`) | `Correo.cls:32-51`, `constructor.bas:1143-1165` |
| Alta del correo (aplicación «NoConformidades», cuerpo, asunto, originador, fecha, destinatarios, copia, copia oculta) | `TbCorreosEnviados` | DAO `OpenRecordset` + `AddNew/Update` | `Correo.cls:120-139` |
| Datos del informe | objetos de NC en memoria | lectura de objetos; las consultas las hace `HTML.bas` (ver fichas de NC) | `:41, 46` |

Sin transacciones: el identificador se calcula fuera del `AddNew`.

## Coexistencia

- El identificador `Max + 1` calculado en cliente (`constructor.bas:1165`) se corresponde con una secuencia de PostgreSQL; con varios usuarios puede duplicar claves, y Jet/ACE no ofrece bloqueo.
- Fecha `Now` y alta DAO en tabla vinculada: revisar tipo `timestamp` y valores `Null` en campos opcionales (`Correo.cls:130-138`).
- El envío efectivo lo hace un proceso ajeno a `src` (ver Pendientes).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- En `Form_Open`, si `EstablecerDatos` falla antes de `AvanceCerrar`, el aviso de error se muestra pero el formulario se cierra sin dejar corregir (`:137-155`).
- La copia oculta fija del código (`Correo.cls:116`) envía copia a una dirección personal sin que el usuario lo vea.
- El mensaje de error del envío se titula `ComandoRegistrar_Click`, nombre que no existe en este formulario (`:71`).
- En auditoría el campo PARA se rellena con la variable «en pruebas» (`:96`; `Variables Globales.bas:324`), nombre que sugiere que no es la lista real de producción.
- Doble pulsación de «Enviar correo» podría registrar el correo dos veces: no hay bloqueo del botón (`:14-67`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: editor del cuerpo con vista previa antes de registrar; validación del formato de direcciones; plantilla de destinatarios por tipo de NC.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir desde NC de proyecto se precargan asunto con el código de NC, PARA con el responsable y CC. con el usuario conectado (`:85-98`).
2. Al abrir desde NC de auditoría el asunto incluye partícula, auditoría y número a dos cifras (`:87-88`).
3. Sin asunto, o sin PARA ni CC., no se registra correo y se muestra el motivo (`:30-39`).
4. Un envío correcto crea una fila en la tabla de correos con el cuerpo HTML de la NC y muestra «Correo ordenado para su envío» (`:51-66`).
5. Abierto sin argumento o sin NC activa, avisa y se cierra (`:121-135, 154-155`).

## Pendientes

- Proceso que envía los correos registrados en `TbCorreosEnviados`: no está en `src/forms`; confirmar si lo hace un servicio externo.
- Contenido exacto del informe: pendiente de la lectura de `HTML.bas:784-945` en la ficha de NC.
- Estado de la fila tras el envío (campos de envío): no se escribe aquí.
