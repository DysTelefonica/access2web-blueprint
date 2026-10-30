# Form_FormExpedienteDetalle

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormExpedienteDetalle.cls` (79 líneas) y `.form.txt` · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `fichero:línea` del commit fijado.

Ficha de **solo lectura** de un expediente: copia seis campos desde la global `m_ObjExpedienteActivo` a seis cuadros de texto bloqueados. No filtra, no busca y no escribe nada. No declara eventos: solo notifica errores con `MsgBox` y se cierra solo si falla.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormExpedientesBusqueda` | fija `m_ObjExpedienteActivo`, cierra cualquier `FormExpedienteDetalle` previa y abre esta | `Form_FormExpedientesBusqueda.cls:149-153` |
| `Form_FormSolicitudRenovacionDatos` | mismo patrón: fija la global y abre el formulario | `Form_FormSolicitudRenovacionDatos.cls:195-202` |
| `Form_FormSolicitudesAltaTraspasoDatos` | mismo patrón | `Form_FormSolicitudesAltaTraspasoDatos.cls:169-176` |

**Contrato de entrada:** el formulario no recibe argumentos. Exige que la global `m_ObjExpedienteActivo` (`VG.bas:202`) esté informada antes de abrirse; si está vacía, `EstablecerDatos` aborta con «El formulario se ha abierto con parámetros insuficientes» (`:49-52`) y `Form_Open` lo propaga como error interno (`:25-27`) y cierra la ventana (`:41-43`).

No abre ningún otro formulario.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | Único rótulo; `Form_Open` copia su texto al título de la ventana | `:23`, `.form.txt:807-808` |
| `IDExpediente` | texto bloqueado | — | Identificador del expediente; es el único campo que se copia siempre | `:54`, `.form.txt:845-860` |
| `Titulo` | texto bloqueado | — | Título del expediente | `:58-60`, `.form.txt:897-913` |
| `Nemotecnico` | texto bloqueado | — | Nemotécnico | `:55-57`, `.form.txt:931-947` |
| `CodExp` | texto bloqueado | — | Código corto | `:61-63`, `.form.txt:986-1002` |
| `CodExpLargo` | texto bloqueado | — | Código largo | `:64-66`, `.form.txt:1039-1055` |
| `Ordinal` | texto bloqueado | — | Número de expediente | `:67-69`, `.form.txt:1092-1108` |
| `cmdSalir` | botón | `OnClick` | Cierra el formulario | `:7-11`, `.form.txt:1157-1168` |

No hay eventos hacia el contenedor: no declara ningún `Public Event` (`:5-11` declaran solo `m_Error` y los tres procedimientos).

## Flujos

### Apertura (`Form_Open`, `:14-43`)

1. Muestra el reloj de arena, limpia `m_Error` y ajusta las dimensiones con `Ajustar Me` (`:22`; `FU.bas:735-758`).
2. Copia el rótulo al título (`:23`) y llama a `EstablecerDatos` (`:24`).
3. Devuelve el foco y habilita la edición (`:33`).

### Copia de datos (`EstablecerDatos`, `:45-77`)

1. Si `m_ObjExpedienteActivo` está vacía, escribe el error y lanza el error interno 1000 (`:49-52`).
2. Copia `IDExpediente` sin condición (`:54`).
3. Copia los otros cinco campos **solo si la propiedad de origen no está vacía** (`:55-69`): un valor vacío deja intacto lo que hubiera en el cuadro.
4. Ante un error distinto del 1000, deja el texto en `p_Error` (`:72-75`) sin relanzar.

### Errores

`Form_Open` distingue los dos casos: si el error es el interno 1000 ya viene redactado en `m_Error` y no se notifica por correo (`:34-38`); si es cualquier otro, compone el mensaje, avisa a `CorreoAlAdministrador` (`:35-37`), muestra un `MsgBox` crítico (`:42`) y cierra el formulario (`:43`). Devuelve `AllowEdits` y el reloj de arena en ambos caminos (`:33`, `:40-41`).

## Reglas, validaciones y cálculos

- **No hay validación de entrada más allá de la existencia de la global.** Ningún control tiene regla de validación en el `.form.txt`.
- **Todos los campos están bloqueados en diseño** (`.form.txt:846`, `:898`, `:932`, `:987`, `:1040`, `:1093`), pero `Form_Open` pone `AllowEdits = True` (`:33`, `:40`): la propiedad del formulario sugiere edición y los cuadros la impiden. La copia además es **unidireccional**: ningún manejador vuelve a volcar los cuadros sobre `m_ObjExpedienteActivo`, de modo que cualquier cambio del usuario, si lo hubiera, no llegaría a ninguna parte.
- **Escritura nula:** no hay `RecordSource` —el `.form.txt` describe un formulario de detalle sin origen de datos (`.form.txt:4-30`)— ni ninguna llamada de escritura.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Expediente activo | `m_ObjExpedienteActivo`, ya en memoria | lectura de global | `VG.bas:202`; `:53-70` |
| Origen del expediente | `TbExpedientes` | SQL por concatenación, en el formulario que lo cargó | `constructor.bas:2306-2310` |

Sin escrituras: ni en base de datos ni en el sistema de ficheros.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Ya se ve: el paso de datos es una global mutable compartida por tres formularios que la escriben (`Form_FormExpedientesBusqueda.cls:149`, `Form_FormSolicitudRenovacionDatos.cls:195`, `Form_FormSolicitudesAltaTraspasoDatos.cls:169`) en lugar de un identificador por parámetro; la coexistencia entre la ficha y el expediente abierto depende del orden de cierre y reapertura (`Form_FormExpedientesBusqueda.cls:150-153`); y el formulario hereda el patrón de ventana modal por `MsgBox` y reloj de arena.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-9.** `EstablecerDatos` es `Public` (`:45`) pero el único llamador es `Form_Open` (`:24`): la API pública no la usa nadie, ni siquiera los tres formularios que abren la ficha y podrían inyectarla en lugar de escribir la global. En la misma línea, cinco de los seis campos solo se copian si el origen no está vacío (`:55-69`): hoy es inocuo porque Access crea una instancia nueva en cada apertura, pero cualquier reutilización dejaría valores obsoletos en pantalla sin señal alguna. Y `AllowEdits = True` (`:33`, `:40`) contradice los seis cuadros bloqueados en diseño: dos intenciones en el mismo formulario, sin resolución.

**Defectos de interfaz heredados (a verificar):** el título de diseño es `Edición de Suministrador` (`.form.txt:41`), que no corresponde a este formulario; `Form_Open` lo sobrescribe con el texto de `lblTitulo` (`:23`), de modo que el defecto solo es visible en diseño y en la pestaña de la ventana antes de que termine la apertura. Las etiquetas de los cuadros siguen el patrón `Etiqueta<n>` sin renombrar (`.form.txt:835`, `:887`, `:975`, `:1029`, `:1082`, `:1135`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: la ficha como panel embebido en la ficha de solicitud en lugar de ventana aparte; identificador de expediente en la ruta en lugar de global; y un endpoint de lectura de expediente que devuelva los campos ya formateados, en lugar de copiar propiedad a propiedad.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir la ficha con un expediente activo muestra identificador, nemotécnico, título, código corto, código largo y número.
2. Un campo vacío en el origen se muestra vacío en pantalla.
3. Abrir la ficha sin expediente activo avisa al usuario y no deja la ventana abierta.
4. Ningún dato de la ficha se puede modificar.
5. Salir cierra la ventana sin avisos.

## Pendientes

- Si algún flujo legado depende de que `EstablecerDatos` sea público (H-HPS-S-9).
- Si la estructura real de `TbExpedientes` incluye campos que esta ficha no expone: la decisión es de `Expediente.cls` y de los formularios que la abren.
- Si el ancho declarado (`.form.txt:22`) y los límites derecho e inferior (`.form.txt:27-28`) siguen siendo válidos para los títulos largos que llegan desde `Expediente.Titulo`.