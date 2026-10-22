# Form_FormGestionRiesgosDatosGenerales

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormGestionRiesgosDatosGenerales.cls` (650 líneas) y `src/forms/FormGestionRiesgosDatosGenerales.form.txt` · **PRD:** [README](../README.md)

Subformulario de **datos generales de la edición**: identifica el expediente associado, muestra sus fechas y su código documental y registra quién elabora, revisa y aprueba la gestión de riesgos. Es el único de los cuatro que guarda datos de vuelta en el objeto de proyecto. Las citas son `Form_FormGestionRiesgosDatosGenerales.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormGestionRiesgos` | Pestaña «Datos generales»; es además el `SourceObject` por defecto del `FormDetalle` | `Form_FormGestionRiesgos.cls:280-281`, `FormGestionRiesgos.form.txt:449` |
| `Form_FormGestionRiesgos` | Lo llama explícitamente al cambiar de pestaña para volcar los datos al objeto de proyecto | `Form_FormGestionRiesgos.cls:59-60`, `:974-975`, `:1019` |
| `FormExpedientesBusqueda` | Se suscribe a su evento `Seleccionado` y recibe el expediente elegido | `:13`, `:132` |

Es el subformulario con más superficie de entrada y salida de los cuatro. Publica dos funciones que el contenedor invoca: `RellenarDatosAlObjeto` (`:419`) y `RellenarDatosExpediente` (`:329`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Carga los datos; ante error avisa y **cierra el subformulario** | `:199-230` |
| (formulario) | — | `Form_Unload` | Vuelca los campos al objeto de proyecto antes de salir | `:102-130` |
| `PalabraClave` | `TextBox` | `KeyDown` | Con Enter, dispara la búsqueda y anula la tecla | `:168-186` |
| `ComandoBuscarExpediente` | botón | `Click` | Busca por palabra clave y decide entre rellenar, limpiar o abrir la búsqueda | `:235-327` |
| `ComandoIrABusquedaExpedientes` | botón | `Click` | Abre el formulario de búsqueda, opcionalmente filtrado por la palabra clave | `:59-95` |
| `ComandoActualizar` | botón | `Click` | Recupera el expediente por identificador y refresca los datos | `:25-57` |
| `Elaborado`, `Revisado`, `Aprobado` | `ComboBox` | `BeforeUpdate` | Escribe el valor en el objeto de proyecto y repinta los obligatorios | `:20-23`, `:97-100`, `:193-196` |
| `ParaInformeAvisos` | `ComboBox` | `BeforeUpdate` | Igual que los anteriores, con su propio manejador | `:188-191` |
| `IDExpediente`, `Juridica`, `CodExp`, `Nemotecnico`, `Titulo`, `FechaInicioContrato`, `FechaFinContrato`, `OrganoContratacion`, `ResponsableCalidad`, `FechaCierre`, `CodigoDocumento`, `CorreoRAC`, `EnUTE` | campos | — | Solo lectura; los puebla `RellenarDatosExpediente` y `EstablecerDatos` | `:514-1186`, `.form.txt` |

## Flujos

### Apertura (`Form_Load`, `:199-230`)

1. `EstablecerDatos` (`:208`); si devuelve mensaje, `Err.Raise 1000` (`:210`).
2. `AvanceCerrar` (`:213`) — mecanismo común del proyecto para ceder la pantalla al contenedor.
3. En error: mensaje al administrador + `MsgBox` (`:216-226`), y **cierre del subformulario** (`:229`).

### `EstablecerDatos` (`:446-551`)

1. Habilita la búsqueda (`:449-451`).
2. Si no es un alta (`m_EsAlta = EnumSiNo.No`), recupera el proyecto del inicio con `Constructor.getProyecto` (`:453-463`).
3. **Si la edición activa tiene fecha de publicación**, bloquea la búsqueda y su botón (`:466-475`): una edición publicada ya no admite cambiar de expediente.
4. Habilita o deshabilita `ComandoActualizar` según haya identificador de expediente (`:476-482`).
5. Si `m_EsAlta` llegó `Empty`, avisa «No se sabe si es para alta o no» (`:483-486`).
6. Vacía Elaborado/Revisado/Aprobado y puebla los desplegables (`:487-493`).
7. **El permiso de escritura no se decide aquí**: lo toma del contenedor, `Form_FormGestionRiesgos.blnPermitidoEditar`, y lo aplica como `Me.AllowEdits` (`:495-499`).
8. Vuelca los datos del proyecto activo y el expediente a los campos (`:503-540`).

### Búsqueda por palabra clave (`ComandoBuscarExpediente_Click`, `:235-327`)

- **Sin palabra clave**: no hace nada y sale (`:249-255`).
- **Sin resultados** (`m_Col Is Nothing`): limpia **diez** campos, habilita o deshabilita el botón de actualizar y abre la búsqueda de expedientes (`:261-281`).
- **Un solo resultado**: lo carga directamente y **deja el expediente en pantalla** (`:284-305`).
- **Varios resultados**: abre `FormExpedientesBusqueda` con la palabra clave y espera el evento `Seleccionado` (`:306-315`).

### Selección desde la búsqueda (`m_FormExpedientes_Seleccionado`, `:132-166`)

Carga el expediente elegido, habilita el botón de actualizar según el identificador, **vacía la palabra clave**, repinta los obligatorios, pone el foco en `Elaborado` y cede con `AvanceCerrar` (`:138-150`).

### Cierre (`Form_Unload`, `:102-130`)

Vuelca ocho campos al objeto de proyecto mediante `RellenarDatosAlObjeto` (`:110`) y cede con `AvanceCerrar` (`:114`). Si hay error, avisa y cede igualmente (`:119-128`).

### `RellenarDatosAlObjeto` (`:419-444`)

Escribe en el proyecto activo: `IDExpediente`, `CodigoDocumento`, `ParaInformeAvisos`, `EnUTE`, `CorreoRAC`, `Elaborado`, `Revisado` y `Aprobado`, todos con `Nz(...)` (`:428-435`).

## Reglas, validaciones y cálculos

- **El permiso de edición se hereda del contenedor**, no se calcula aquí (`:495-499`).
- **Publicada la edición, el expediente queda congelado** (`:470-475`).
- **Tres campos con derivación de valor.** `Nemotecnico`, `EnUTE` y `CodigoDocumento` prefieren el valor del proyecto **al inicio** si existe, y solo entonces el del expediente (`:353-365`, `:381-394`, `:396-403`). En una edición cargada desde el principio, el valor del expediente no se ve nunca.
- **`ParaInformeAvisos` tiene valor por defecto**: si el proyecto no lo tiene, se fuerza a «Sí» al cargar (`:516-520`).
- **Las fechas solo se escriben si son fechas** (`IsDate`): en caso contrario el control conserva su valor anterior en lugar de vaciarse (`:368-374`).
- **Los desplegables de personas** se rellenan con usuarios técnicos; si se está en alta solo se ofrecen los activos, si no, todos (`:578-595`).

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Expediente por identificador | `Constructor.getExpediente` | `:33` |
| Búsqueda por palabra clave | `Constructor.getExpedientesBusqueda` | `:257` |
| Proyecto del inicio | `Constructor.getProyecto` | `:455` |
| Usuarios de los desplegables | `m_ObjEntorno.ColUsuariosTecnicos(Activos)` | `:624-628` |
| Escritura | Propiedades del objeto de proyecto | `:428-435`, `:21`, `:98`, `:189`, `:194` |

La escritura real en base de datos no se produce aquí: se escriben propiedades del **objeto en memoria**; quién las persiste y cuándo queda fuera de este formulario.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** ficha de la edición con los datos del expediente (identificadores, fechas, órgano de contratación, responsable de calidad, código documental y cadena de responsables de calidad) y tres campos de circuito de aprobación —elaborado, revisado y aprobado— más un interruptor de «para informe de avisos». Permite buscar y elegir expediente por palabra clave, y se bloquea cuando la edición ya tiene fecha de publicación.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-267.** La lista de campos que se limpian cuando la búsqueda no devuelve nada (`:262-271`) tiene **diez entradas**, pero `RellenarDatosExpediente` limpia **doce** (`:334-345`): faltan `Titulo` y `EnUTE`. Si el usuario tenía un expediente cargado y busca una palabra clave sin resultados, esos dos campos conservan el valor anterior.
- **H-GR-268.** `Err.Raise 1000` (`:35, 39, 210`, etc.) no lleva `vbObjectError`; conviene verificar en ejecución que el manejador se ejecuta siempre.
- **H-GR-269.** `ComandoActualizar_Click` y `ComandoBuscarExpediente_Click` **no** llaman a `AvanceCerrar` en el camino correcto, mientras que el resto de manejadores del formulario sí lo hacen (`:43`, `:79`, `:114`, `:150`). La cesión de pantalla es inconsistente entre manejadores.
- **H-GR-270.** El manejador de error de `Form_Unload` (`:119-128`) y el de `m_FormExpedientes_Seleccionado` (`:157-165`) llaman a `AvanceCerrar` **después** de mostrar el `MsgBox`; si el usuario no cierra el diálogo, la cesión se retrasa indefinidamente.
- **H-GR-271.** En `Form_Load`, `AvanceCerrar` (`:213`) se ejecuta también en el camino de error (`:227`), de modo que el subformulario cede la pantalla y luego se cierra (`:229`): comportamiento difícil de reproducir.
- **H-GR-272.** La comparación `If Not m_ObjProyectoAlInicio.EdicionActiva.FechaPublicacion <> ""` (`:470`) usa la forma negada; un valor `Null` hace que la expresión sea distinta a la de una cadena vacía. El congelamiento depende de ese matiz.
- **H-GR-273.** `RellenarDatosAlObjeto` copia `EnUTE` (`:431`) como cadena, pero `RellenarDatosExpediente` lo pinta en pantalla como «Sí»/«No» a partir de la enumeración (`:388-392`). El valor persistido y el mostrado no tienen por qué coincidir.
- **H-GR-274.** Los tres `BeforeUpdate` escriben en el objeto de proyecto **antes** de validar nada, y no cancelan el evento ni devuelven error si la escritura falla: el fallo solo aparecería al abandonar el formulario.
- **H-GR-275.** `EstablecerComboTecnicos` contiene una etiqueta `siguiente:` (`:637`) que no está referenciada por ningún `GoTo`: es código muerto de una depuración anterior.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: modelar la edición como un formulario de cabecera con guardado explícito en lugar del volcado implícito al cerrar; sustituir la búsqueda por texto libre por un buscador con resultados en la misma página; hacer que el circuito elaborado/revisado/aprobado sea un estado con transiciones en lugar de tres campos libres; mover el «para informe de avisos» a una configuración de la edición; unificar la cesión de pantalla en un único mecanismo del contenedor.

## Criterios de aceptación

Propuestos, no aprobados:

1. La ficha muestra los datos del expediente elegido y el valor efectivo deElaborado, Revisado y Aprobado.
2. Buscar por palabra clave con un único resultado lo carga directamente; con varios, ofrece un selector; sin resultados, limpia **todos** los campos y lo indica.
3. Con la edición publicada, el expediente no se puede cambiar y el formulario no admite edición.
4. Al salir de la pestaña, los valores capturados quedan disponibles para el resto de la gestión de riesgos.
5. Si no hay proyecto activo o la edición no se puede determinar, se avisa y no se deja la vista a medio cargar.

## Pendientes

- Qué persiste los valores escritos en el objeto de proyecto: no está en este formulario.
- Qué significa exactamente `AvanceCerrar` y si su omisión en dos manejadores es intencionada.
- Si la cadena de responsables de calidad (`CorreoRAC`) debe mostrarse o solo conservarse.
- El comportamiento esperado de la búsqueda cuando hay más de un resultado, más allá de abrir otro formulario (ver [forms](../../forms.md)).