# Form_FormExpedienteHitos

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpedienteHitos.cls` (296 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Hitos» de la ficha de expediente: lista los hitos (fecha, descripción, fecha de fin de garantía, importe) y permite añadirlos y borrarlos. Formulario independiente, sin `RecordSource`: los datos llegan por el DTO global `m_ObjExpedienteDTOActivo`. Las citas son `Form_FormExpedienteHitos.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormExpediente` | botón de navegación `NavHitos` con `NavigationTargetName = "FormExpedienteHitos"`, sin handler que vuelque el DTO | `Form_FormExpediente.form.txt:1012` |

No abre otros formularios ni informes. Lee el expediente de `m_ObjExpedienteDTOActivo` (`.Expediente` y `.ColHitos`, `classes/ExpedienteDTO.cls:26`) y el permiso de `Form_FormExpediente.m_AbiertoParaEditar` y `EsAdministrador` (`:175-176`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaHitos` | ListBox, lista de valores, 4 columnas, con encabezados | — | Muestra Fecha, Descripción, Fecha G. e Importe | `.form.txt:587-605`, `.cls:210` |
| `Descripcion` | TextBox independiente | — | Descripción del hito | `.form.txt:635`, `.cls:29` |
| `FechaHito` | TextBox, formato fecha corta | — | Fecha del hito; obligatoria | `.form.txt:695-696`, `.cls:22-26` |
| `FechaGarantiaHito` | TextBox, formato «General Date» (rótulo «F.Fin Garantia») | — | Fecha de fin de garantía | `.form.txt:725-726, 939` |
| `Importe` | TextBox, formato moneda | — | Importe del hito | `.form.txt:664-665` |
| `ComandoRegistrar` (consejo «Alta») | botón | `OnClick` → `ComandoRegistrar_Click` | Añade un hito | `.cls:9-86`, `.form.txt:753-757` |
| `ComandoEliminar` (consejo «Eliminar elemento seleccionado») | botón | `OnClick` → `ComandoEliminar_Click` | Borra el hito seleccionado | `.cls:90-162`, `.form.txt:844-848` |
| Formulario | — | `OnLoad` → `Form_Load` | Permisos y carga de la lista | `.cls:166-198`, `.form.txt:222` |

## Flujos

### Apertura (`Form_Load`, `.cls:166-198`)

1. Parte con `AllowEdits = False`; lo pasa a `True` solo si `m_AbiertoParaEditar` y `EsAdministrador = Sí` (`:174-181`).
2. `EstablecerDatos` exige `m_ObjExpedienteDTOActivo.Expediente`; si falta, «El formulario se ha abierto con parámetros insuficientes» (`:265-268`). Habilita o deshabilita los controles con `Tag = "EJECUTIVO"` con la misma regla (`:269-281`); en el `.form.txt` no hay ningún control con ese `Tag`.
3. `RellenarListas` vacía la lista y llama a `RellenarListaHitos` (`:242-246`), que pone la fila de encabezados (`:210`), carga `ColHitos` desde `Expediente.Hitos` si está vacía (`:211-213`), añade una fila por hito con el importe formateado «#,##0.00 €» y selecciona la fila 0 (`:216-226`).

### Alta de hito (`ComandoRegistrar_Click`, `.cls:9-86`)

1. Valida que `FechaHito` sea fecha: «La fecha del hito es obligatoria» (`:22-26`). Construye un `ExpedienteHito` con los cuatro campos (`:27-33`).
2. **Expediente nuevo** (`IDExpediente = ""`): guarda en `ColHitos`, con la fecha como clave, una cadena `fecha|descripción|garantía|importe` y añade la fila a la lista; si la fecha ya existe, no hace nada (`:38-52`).
3. **Expediente existente:** `ExpedienteOperaciones.RegistrarHito` (`classes/ExpedienteOperaciones.cls:1743-1835`); si devuelve «1», recarga `ColHitos` de la base y la lista (`:54-66`).
4. Selecciona la última fila y vacía los cuatro campos (`:69-73`).

### Baja de hito (`ComandoEliminar_Click`, `.cls:90-162`)

1. Toma la columna 0 de la fila seleccionada; si está vacía, sale sin mensaje (`:103-109`).
2. Exige `EsAdministrador = Sí`: «Operacion no autorizada» (`:110-113`).
3. Expediente nuevo: quita la clave de `ColHitos` (la anula si queda vacía) y la fila de la lista (`:114-134`). Existente: `ExpedienteOperaciones.EliminarHito` por fecha y recarga (`:136-146`).

### Errores

Los errores propios (1000) muestran `m_Error` en un `MsgBox` crítico; los demás, además, envían correo al administrador (`CorreoAlAdministrador`, `:80-85, 156-161, 192-197`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Fecha del hito obligatoria | formulario y operación | `.cls:23-26`; `ExpedienteOperaciones.cls:1782-1785` |
| Descripción obligatoria (solo con expediente existente) | operación | `ExpedienteOperaciones.cls:1778-1781` |
| Un hito por fecha y expediente: registrar con fecha existente **borra** el anterior y crea uno nuevo | operación | `ExpedienteOperaciones.cls:1789-1802` |
| Garantía e importe solo se guardan si son fecha y número | operación | `ExpedienteOperaciones.cls:1809-1814` |
| `IDHitoExpediente` = máximo + 1 de `TbExpedientesHitos` | `ExpedienteHito.IDHitoExpedienteCalculado` | `classes/ExpedienteHito.cls:82`; `FUNCIONES UTILES.bas:2851, 2864` |
| Borrar exige administrador; añadir no lo comprueba en código | formulario | `.cls:110-113` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Leer hitos | `TbExpedientesHitos` por `IDExpediente`, diccionario por `IDHitoExpediente` | SQL dinámico DAO | no | `modules/constructor.bas:885-910` |
| Alta | `TbExpedientesHitos`: `Delete` del de igual fecha + `AddNew` | SQL dinámico DAO, literal `#mm/dd/yyyy#` | no (sin `p_db`) | `ExpedienteOperaciones.cls:1789-1815` |
| Baja | `DELETE * FROM TbExpedientesHitos` por expediente y fecha | SQL dinámico DAO | no | `ExpedienteOperaciones.cls:1853-1857` |
| Tras alta o baja | `RegistrarExpEntidades` (ámbito efectivo `Todo`) y `RegistrarUltimoCambio` | DAO | no | `ExpedienteOperaciones.cls:1819-1821, 1860-1862, 607` |

## Coexistencia

Pendiente de `coexistencia.md`. Ya se ve: literales de fecha Jet `#mm/dd/yyyy#` y `DELETE *` concatenados; identificador `Max + 1` calculado en VBA (carrera entre usuarios); escrituras fuera de transacción; clave lógica «expediente + fecha» sin índice único comprobado.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- En alta de expediente, `ColHitos` guarda cadenas (`:46`), pero `Registrar` no persiste `ColHitos` (`ExpedienteOperaciones.cls:402-495` no la recorre): los hitos de un expediente nuevo se perderían al guardar.
- Al volver a la pestaña en un expediente nuevo, `RellenarListaHitos` hace `Set` sobre una cadena (`:218`); el error quedaría oculto porque `EstablecerDatos` llama a `RellenarListas` sin `p_Error` (`:282`).
- En expediente nuevo, `AddItem` añade cinco valores (clave + cuatro) a una lista de cuatro columnas (`:48`, `.form.txt:593`): filas desplazadas.
- `RellenarListaHitos` no reinicia `m_Importe` en el bucle (`:219-221`): un hito sin importe mostraría el del anterior.
- `ComandoRegistrar` no comprueba permisos; el bloqueo depende de `AllowEdits = False` sobre controles independientes (`:174`).
- `EliminarHito` lanza `Err.Raise 1000, , "<texto>"` sin asignar `p_Error` (`ExpedienteOperaciones.cls:1850-1851`): esas validaciones fallarían en silencio.
- Si `RegistrarExpEntidades` falla tras `Update`, el hito queda guardado pero el formulario muestra error y no recarga la lista (`ExpedienteOperaciones.cls:1815-1819`; `.cls:58-60`).
- Registrar con una fecha ya usada sustituye el hito sin aviso (`ExpedienteOperaciones.cls:1795-1798`); en expediente nuevo, en cambio, lo ignora (`:45`).
- `lst.Selected(0)` (`:226`) apunta a la fila de encabezados.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: edición de un hito existente; aviso de duplicado por fecha; persistir los hitos del alta junto con la cabecera.

## Criterios de aceptación

Propuestos, no aprobados:

1. Añadir un hito sin fecha devuelve «La fecha del hito es obligatoria» y no escribe.
2. Añadir un hito sin descripción a un expediente existente devuelve «La Descripción del hito es obligatoria».
3. Borrar un hito sin ser administrador devuelve «Operacion no autorizada» y no borra.
4. Sin permiso de edición, no se puede añadir ni borrar.
5. Los hitos añadidos durante el alta de un expediente quedan guardados al registrarlo (o la decisión contraria queda documentada).

## Pendientes

- Contenido de `RegistrarUltimoCambio` (`FUNCIONES UTILES.bas:189`) y de `CorreoAlAdministrador`.
- Comportamiento real de `AllowEdits = False` sobre los TextBox independientes y de `AddItem` con cinco valores (comprobación en el binario).
- Esquema de `TbExpedientesHitos` (tipos, índices, unicidad de fecha): `erd-origen.md`.
- Rótulos visibles de los botones (el `.form.txt` declara `Caption = "Comando247"` con imagen).
