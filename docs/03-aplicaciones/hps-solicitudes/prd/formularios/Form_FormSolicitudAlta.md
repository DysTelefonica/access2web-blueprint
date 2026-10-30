# Form_FormSolicitudAlta

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudAlta.cls` (130 líneas) y `.form.txt` · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_X.cls:línea` del commit fijado.

Selector del tipo de solicitud: elige entre alta, traspaso, renovación y elevación de grado y lanza el formulario de datos que corresponde. Es un formulario **no enlazado** (sin `RecordSource`): el único control con contenido es el combo `TIPO`, cuyo `RowSource` es una lista de valores fijada en diseño (`.form.txt:939-940`). No lee ni escribe tablas.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudesGestion` | `ComandoAlta_Click` cierra una instancia previa sin guardar y abre esta | `Form_FormSolicitudesGestion.cls:489-493` |

No declara parámetros. Abre dos formularios según el combo: `FormSolicitudAltaTraspaso` con `OpenArgs:=m_Tipo` para alta y traspaso, y `FormUsuariosHPSBusqueda` para los otros dos tipos (`:61-73`).

Estado compartido: la global `m_ObjSolicitudActiva` (`VG.bas:197`), que aquí se instancia vacía (`:56`), y la global `m_ObjSolicitudAlInicio` (`VG.bas:198`), que se anula (`:57`). También lee `EsTecnico` (`VG.bas:193`), `m_EnOficina` (`VG.bas:191`) y `m_TituloFormulario` (`VG.bas:196`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `TIPO` | combo, `RowSource` de lista de valores | — (sin handler) | Tipo de solicitud a registrar | `:51`, `.form.txt:918-940` |
| `Etiqueta232` | etiqueta | — | Rótulo «TIPO» | `.form.txt:959-969` |
| `ComandoRegistrar` | botón «Registrar» | `OnClick` | Valida el tipo y lanza el formulario de destino | `:41-95`, `.form.txt:978-991` |
| `ComandoAyuda` | botón | `OnClick` | Abre el PDF de ayuda del formulario | `:12-37`, `.form.txt:817-831` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra el formulario | `:7-10`, `.form.txt:1039-1052` |
| `lblTitulo` | etiqueta de encabezado | — | Rótulo «ALTA DE SOLICITUD» | `.form.txt:800-810` |

## Flujos

### Apertura (`Form_Open`, `:97-130`)

1. `Ajustar Me` dimensiona el formulario al formulario y a sus secciones (`:105`; `FU.bas:735-759`).
2. Deshabilita `ComandoRegistrar` y `AllowEdits` (`:106-107`).
3. Fija el título de la ventana con `Me.Caption = m_TituloFormulario` (`:108`).
4. Solo si el usuario **no** es técnico reactiva `ComandoRegistrar` y `AllowEdits` (`:109-112`).
5. `AvanceCerrar` oculta la etiqueta de avance si existe (`:115`; `FU.bas:555-582`) y, al terminar, reactiva `AllowEdits` de forma incondicional (`:119`).

### Selección y registro (`ComandoRegistrar_Click`, `:41-95`)

1. Lee el combo con `Nz(Me.TIPO, "")`; si está vacío, aborta con «Seleccione un tipo» (`:51-55`).
2. Crea una solicitud vacía en `m_ObjSolicitudActiva` y anula la instantánea de inicio (`:56-57`).
3. Si el tipo es alta o traspaso, prepara `FormSolicitudAltaTraspaso` y propaga `TIPO` al objeto solicitud (`:59-61`); si es renovación o elevación de grado, exige `m_EnOficina` y prepara `FormUsuariosHPSBusqueda` (`:62-68`).
4. Si el formulario destino ya está abierto, lo cierra sin guardar (`:69-71`).
5. Abre el destino con `OpenArgs:=m_Tipo` (`:73`) y sale tras `AvanceCerrar` (`:75`, `:81`).

### Alternativos

- Renovación y elevación de grado solo se ofrecen en oficina: fuera de ella el mensaje es «Para poder hacer una renovación se ha de estar en la oficina. No se puede acceder a los datos de HPS» (`:63-67`).
- Cualquier valor del combo que no sea alta o traspaso toma la rama de la búsqueda de usuarios, incluido un texto no previsto (`:61-68`).

### Errores

Los cuatro handlers acumulan el texto en `m_Error`, envían `CorreoAlAdministrador` (`FU.bas:585`) salvo el error interno 1000 —el usado para los avisos previstos— y muestran `MsgBox` crítico. `Form_Open` cierra el formulario tras el aviso (`:121-129`); `ComandoRegistrar_Click` no cierra nada (`:87-94`). `cmdSalir_Click` usa `On Error Resume Next` sin aviso (`:7-10`).

## Reglas, validaciones y cálculos

- **Tipo obligatorio:** única validación de negocio del formulario (`:48-52`).
- **Oficina:** `m_EnOficina` solo se consulta en la rama de renovación y elevación de grado (`:63`); el valor procede de la sesión (`VG.bas:191`).
- **Perfil técnico:** el botón de registro solo se habilita para quien no tiene el perfil técnico (`:109-112`); el resto de formularios no aplican esta comprobación.
- **Sin validaciones de lista:** el combo declara `LimitToList = NotDefault` (`.form.txt:919`), por lo que el valor admitido depende de la propiedad heredada del diseño y no del código.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| — | — | — | El formulario no consulta ni escribe en la base de datos |

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Lo que ya se ve: el formulario solo navega; no aporta reglas de datos que trasladar. La comparación de `OpenArgs` con cadenas literales en lugar de `EnumTipo` (`VG.bas:166-171`) y la apertura de ventanas apiladas en lugar de navegación con subformulario condicionan el diseño de rutas de la versión web.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-14.** `AvanceCerrar` solo oculta la etiqueta de avance y no cierra formularios (`FU.bas:555-582`): tras registrar, `FormSolicitudAlta` permanece abierto detrás de la ventana de destino, en lugar de cederle el paso (`:73-84`).
- **H-HPS-S-15.** `Form_Open` reactiva `AllowEdits` al final sin condición (`:119`), anulando la restricción aplicada al perfil técnico en `:107`; el botón sí queda deshabilitado, pero la propiedad no.
- **H-HPS-S-16.** El destino se cierra con `acSaveNo` si ya estaba abierto (`:69-71`): el trabajo en curso de otro usuario de la aplicación se descarta sin aviso ni confirmación.
- **H-HPS-S-17.** Rótulos heredados de otra aplicación: la ventana arranca como «Edición de Suministrador» (`.form.txt:40`) hasta que `Form_Open` la renombra (`:108`), y el botón de ayuda muestra «Salir» con el texto emergente «Ayuda» (`.form.txt:827-831`).
- `Set m_ObjSolicitudAlInicio = Nothing` (`:57`) descarta la instantánea de sesión que la cadena de registro usa para detectar cambios; el efecto alcanza a los formularios abiertos en el resto de la aplicación.
- `ComandoRegistrar_Click` compara `m_Tipo` con literales en lugar de con `EnumTipo` (`VG.bas:166-171`); un valor no previsto cae en la rama de la búsqueda de usuarios.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: elegir el tipo en una ruta propia (`/solicitudes/nueva?tipo=…`) en lugar de un formulario que se cierra; validar el tipo contra el enumerado en servidor; conservar el aviso de «está en la oficina» como política de autorización y no como mensaje de interfaz.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin tipo seleccionado, el registro no se lanza y se informa «Seleccione un tipo».
2. Con alta o traspaso, se abre el formulario de datos con el tipo ya fijado.
3. Con renovación o elevación de grado fuera de la oficina, el registro no se lanza y se informa del requisito.
4. Con perfil técnico, el botón «Registrar» permanece deshabilitado.
5. El título de la ventana es el de la aplicación, no el del diseño.
6. Tras registrar, no queda una ventana de selección abierta sobre el formulario de destino.

## Pendientes

- Si `m_TituloFormulario` recibe el título correcto en todos los puntos de entrada, y no solo en este.
- El valor heredado de `LimitToList` en el combo `TIPO` (`.form.txt:919`).
- Si `FormUsuariosHPSBusqueda` reenvía a este mismo formulario para los tipos de renovación (rama no cubierta por esta ficha).