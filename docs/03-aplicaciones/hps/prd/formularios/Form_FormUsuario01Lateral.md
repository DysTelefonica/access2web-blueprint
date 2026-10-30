# Form_FormUsuario01Lateral

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormUsuario01Lateral.cls` (167 líneas) y `.form.txt` (996 líneas) · **PRD:** [README](../README.md)

Panel de navegación vertical del usuario **actual**: tres etiquetas-enlace («HPS», «Gral.», «Anexos») que cambian la vista del `SubFormCentral` del contenedor `Form_FormUsuario00Principal`, más tres etiquetas informativas de versión, ubicación y modo de escritura. No está enlazado: no tiene `RecordSource` en el `.form.txt`. Funciona por efecto sobre el contenedor padre, modificando su `SubFormCentral.SourceObject`. Las citas son `Form_FormUsuario01Lateral.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormUsuario00Principal.SubFormLateral` | Carga desde el contenedor en su `Form_Load`; el lateral se activa al pedir el `SourceObject` | `Form_FormUsuario00Principal.cls:50-52` |
| `m_ObjUsuarioActivo` (global) | Lectura: `m_ObjUsuarioActivo` se usa en `SeleccionarEnlace` para conocer el contexto; la rama sin argumento no opera sin él | `:121-123` |

No se abre desde `DoCmd.OpenForm`. Sale al cerrarse el contenedor padre; este formulario no se cierra por sí solo.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblHPS` («HPS») | etiqueta (`OnClick` → `lblHPS_Click`) | `OnClick` | Abre la vista HPS en el central y marca la etiqueta como seleccionada | `:86-110`, `.form.txt:236, 247` |
| `lblDatosGenerales` («Gral.», preseleccionada) | etiqueta (`OnClick` → `lblDatosGenerales_Click`) | `OnClick` | Abre la vista de datos generales y marca la etiqueta como seleccionada; además se invoca desde `Form_Load` | `:57-84`, `.form.txt:690, 702` |
| `lblAnexos` («Anexos») | etiqueta (`OnClick` → `lblAnexos_Click`) | `OnClick` | Abre la vista de anexos y marca la etiqueta como seleccionada | `:32-55`, `.form.txt:719, 730` |
| `lblVersion`, `lblDatosEn`, `lblModoEscritura` | etiquetas informativas | — | Muestran la versión, dónde están los datos y si el sistema permite escritura; se rellenan en `Form_Load` desde `m_ObjEntorno` | `:13-15`, `.form.txt:747, 772, 797` |
| `Imagen7` | imagen | — | Icono del usuario (`UsuarioIcono.png`) | `.form.txt:264, 269` |
| (formulario) | — | `OnLoad` → `Form_Load` | Pinta los rótulos informativos, selecciona «Gral.» y la abre | `:7-29`, `.form.txt:82` |

Todas las etiquetas navegables tienen `HyperlinkAddress ="#"` (`.form.txt:251, 707, 735`) y `Tag = "Elegibles;…Seleccionado"` que `SeleccionarEnlace` actualiza. El `Tag` inicial se fija en diseño: `lblHPS` y `lblAnexos` parten como «NoSeleccionado», y `lblDatosGenerales` como «Seleccionado» (`.form.txt:251, 707, 735`).

## Flujos

### Apertura (`Form_Load`, `:7-29`)

1. Resetea `m_Error` (`:10`) y rellena `lblVersion`, `lblDatosEn` y `lblModoEscritura` con `m_ObjEntorno.VersionAplicacion`, `DatosEn` y `ModoEscritura` (`:13-15`).
2. Llama a `SeleccionarEnlace , m_Error` para devolver a todas las etiquetas a «NoSeleccionado» (`:16`; ver `SeleccionarEnlace`). Si devuelve error, lo eleva como `1000` (`:17-19`).
3. Invoca `lblDatosGenerales_Click`, que es lo que deja el central en `FormUsuario02Edicion` con `lblDatosGenerales` preseleccionada (`:20`).
4. En error: maneja «Al abrir el formulario se ha producido el error: …», envía `CorreoAlAdministrador` si no es un `1000`, muestra `MsgBox` y cierra el formulario (`:22-28`).

### Clic en «HPS» (`lblHPS_Click`, `:86-110`)

1. Si el `SubFormCentral` ya muestra `FormUsuario02HPS`, sale sin tocar nada (`:89-91`).
2. `SeleccionarEnlace "TODOS", m_Error` desmarca las demás y marca `lblHPS` (`:93-94`; ver `SeleccionarEnlace`).
3. `EstablecerControlesformlateral Me, m_Error` aplica el estilo del entorno al panel lateral (`:98`; helper en `FG.bas:963`).
4. Cambia `Me.Parent.Form.SubFormCentral.SourceObject = "FormUsuario02HPS"` (`:102`).
5. En error: compone «Al hacer click en el control se ha producido el error: …», envía `CorreoAlAdministrador` si no es `1000` y muestra `MsgBox` (`:104-109`). **No cierra el formulario lateral**, solo registra el error.

### Clic en «Gral.» (`lblDatosGenerales_Click`, `:57-84`)

1. Si el `SubFormCentral` **no** es `FormUsuario02Edicion`, lo asigna (`:62-64`; usa `Not … = …` en lugar del `… = …` de las otras dos etiquetas, lo que invierte el sentido: la rama abre incluso si ya está cargado, reasignando).
2. `SeleccionarEnlace "", m_Error` repinta los `Tag`; al pasar cadena vacía actúa como reset (`:66`; ver `SeleccionarEnlace`).
3. Marca `lblDatosGenerales.Tag = "Elegibles;Seleccionado"` (`:71`) y aplica estilo vía `EstablecerControlesformlateral Me, m_Error` (`:72`; helper en `FG.bas:963`).

### Clic en «Anexos» (`lblAnexos_Click`, `:32-55`)

Misma estructura que `lblHPS_Click`: optimización de idempotencia (`:35-37`), `SeleccionarEnlace "ANEXO", …` (`:38`), `Me.lblAnexos.Tag = "Elegibles;Seleccionado"` (`:39`), `EstablecerControlesformlateral` (`:43`) y `Me.Parent.Form.SubFormCentral.SourceObject = "FormUsuario02Anexos"` (`:47`).

### `SeleccionarEnlace` (función privada, `:114-167`)

Tres ramas:

1. **`p_TipoHPS = ""`** (`:120-133`): si `m_ObjUsuarioActivo` es `Nothing`, sale en silencio (`:121-123`). En caso contrario, recorre los controles del formulario y, para cada `Label` con `Tag` que contenga «;», lo reescribe a «Elegibles;NoSeleccionado» (`:124-131`). Es el modo «reset».
2. **`p_TipoHPS = "ANEXO"`** (`:134-148`): guarda el nombre del control `lblAnexos` en `m_nombreControlSeleccionado` (`:135`) y lo marca como seleccionado (`:136`); recorre los demás y los pone en «NoSeleccionado» (`:137-147`). Usa `Nz(ctl.Tag, ";")` para tolerar `Tag` nulos (`:141`).
3. **Cualquier otro valor (en la práctica, `"TODOS"`)** (`:150-161`): marca `lblHPS` como seleccionado y desmarca los demás siguiendo el mismo patrón; la variable `m_nombreControlSeleccionado` queda con el valor heredado de la rama anterior — no se reinicia en esta rama (a verificar en runtime si esto afecta a comportamiento no documentado).

El manejador de errores rellena `p_Error` y la función termina silenciosamente (`:163-166`); no eleva la excepción. Las funciones llamadoras son las responsables de propagarlo.

## Reglas, validaciones y cálculos

- **Idempotencia por `SourceObject`.** Los tres `lbl*_Click` salen sin tocar nada si el central ya muestra el formulario pedido (`:35-37, 89-91`); `lblDatosGenerales_Click` es la excepción que reasigna siempre (`:62-64`).
- **Reset por defecto.** `Form_Load` llama a `SeleccionarEnlace` sin argumentos para devolver todas las etiquetas a «NoSeleccionado»; después invoca `lblDatosGenerales_Click` que es el que selecciona «Gral.» (`:16-20`).
- **`Tag` como canal de estado.** La función muta `ctl.Tag` directamente; cualquier handler que lea `Tag` debe asumir el formato `"Elegibles;<estado>"` separado por `;` (`:127, 141, 155`). El `Select Case` se ha reemplazado por `If` con `InStr(…, ";")` y `Nz` (`:127, 141, 155`).
- **Sin captura de errores propia.** `SeleccionarEnlace` no eleva la excepción de error: rellena `p_Error` y devuelve vacío (`:163-166`). Los handlers deben comprobar el `Err.Number` después de cada llamada (`:17-19, 40-42, 67-69, 95-97`).
- **Efecto sobre el contenedor.** Ninguna etiqueta opera sobre sí misma más allá del `Tag`; la carga del subform central se hace siempre sobre el padre (`Me.Parent.Form.SubFormCentral.SourceObject = "FormUsuario02…"`, `:47, 102`). Tres saltos directos a literal cargan un nombre escrito en duro, igual que la cadena usada en `Form_FormUsuario00Principal.cls:51`.
- **`Me.Parent.Form.…` y `Me.Parent.…` son dos accesos distintos** al mismo control del contenedor padre: `lblDatosGenerales_Click` usa `Me.Parent.SubFormCentral.SourceObject` (`:62-63`), mientras que `lblAnexos_Click` y `lblHPS_Click` usan `Me.Parent.Form.SubFormCentral.SourceObject` (`:47, 102`). Conviene verificar que ambos resuelven al mismo control en tiempo de ejecución (a verificar).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Versión, ubicación y modo | `m_ObjEntorno.VersionAplicacion`, `DatosEn`, `ModoEscritura` | propiedades del entorno en memoria | `:13-15` |
| Contexto del usuario | `m_ObjUsuarioActivo` (Nothing check) | propiedad global | `:121-123` |
| Inyección del subform | `Me.Parent.Form.SubFormCentral.SourceObject` | propiedad del subform contenedor | `:47, 102` |
| `Tag` de las etiquetas | en memoria (efecto cosmético) | — | `:128, 142, 156` |

Sin base de datos, sin transacciones, sin correo.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Cosas que ya se ven: el formulario depende de que `SubFormCentral` sea un control concreto del padre — los nombres `FormUsuario02Edicion`, `FormUsuario02HPS`, `FormUsuario02Anexos` son literales que se duplican aquí y en `Form_FormUsuario00Principal` (que solo carga `FormUsuario01Lateral`); `Tag` codifica estado en lugar de propiedad de Access booleana, lo que complica el mapeo a CSS/`aria-pressed`; el triple `SubFormCentral`/`SubFormCentral` se ata a la jerarquía `Parent.Form` con la salvedad de que `lblDatosGenerales_Click` evita un nivel.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `lblDatosGenerales_Click` siempre reasigna `SourceObject` aunque ya esté cargado (`:62-64`); es el opuesto de la guarda idempotente de los otros dos `lbl*_Click` y del propio `lblAnexos_Click`/ `lblHPS_Click`. A verificar si fuerza recargar el subform y, en consecuencia, perder selección del usuario.
- En la rama «TODOS» de `SeleccionarEnlace` (`:149-161`), `m_nombreControlSeleccionado` se queda con el valor de la última iteración de la rama anterior; como solo se usa dentro del `For Each` para comparar, podría no coincidir con ningún control y entonces **ninguna** etiqueta quede desmarcada. A verificar con traza.
- `Me.Parent.…` (`:62-63`) y `Me.Parent.Form.…` (`:47, 102`) son dos formas de alcanzar el mismo destino. `FormUsuario01Lateral` está dentro de `SubFormLateral`, así que `Me.Parent` ya es el contenedor principal; `Me.Parent.Form` añade un salto extra. La verificación del bug depende de la estructura de subform en Access (a verificar).
- `SeleccionarEnlace` no eleva la excepción: un error se rellena en `p_Error` pero no aparece en `Err.Number`. Los tres handlers que llaman comprueban `m_Error` (no `Err.Number`) y propagan con `Err.Raise 1000` (`:17-19, 40-42, 67-69, 95-97`); si el helper escribe `p_Error` y eleva por su cuenta, los `Err.Raise 1000` se anidan como `1000` que el manejador silencia.
- `Form_Load` no es idempotente: si el contenedor pide el lateral dos veces, el segundo `Form_Load` re-marca «Gral.» aunque el usuario haya cambiado a «HPS» (`:20`).
- Las etiquetas `lblHPS`, `lblDatosGenerales` y `lblAnexos` cargan texto en mayúsculas con punto abreviado («Gral.», «Anexos»). En web conviene verificar longitud y capitalización.
- `Imagen7 = "UsuarioIcono.png"` (`.form.txt:267`) depende de un recurso externo al snapshot; no se valida su existencia.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: reemplazar `Tag` por `aria-pressed` o `aria-current`; sustituir la inyección por enrutado (`/usuarios/:id/hps`) con un componente reutilizable; eliminar la duplicación de los nombres `FormUsuario02*` con un registro declarativo; pasar `m_ObjUsuarioActivo` por contexto en lugar de por global.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al entrar, la etiqueta «Gral.» queda preseleccionada y el subform central muestra `FormUsuario02Edicion`.
2. Pulsar «HPS» cambia el subform central a `FormUsuario02HPS`, marca «HPS» como seleccionada y deja las demás sin selección visual.
3. Pulsar «Gral.» o «Anexos» es idempotente si la vista ya está cargada, salvo en el caso conocido de `lblDatosGenerales_Click`.
4. El rótulo `lblVersion` siempre lleva «Versión: » seguido del valor real de `m_ObjEntorno.VersionAplicacion`.
5. Un error de `EstablecerControlesformlateral` se traduce en un `MsgBox` que no cierra el lateral, pero sí impide cambiar el subform central.

## Pendientes

- Firma exacta del setter de `SourceObject` en Access: la doble sintaxis `Me.Parent.…` y `Me.Parent.Form.…` debe documentarse en la coexistencia web.
- `UsuarioIcono.png` no está en el snapshot bajo `src/forms/Form_FormUsuario01Lateral/`.
- `EstablecerControlesformlateral` (`FG.bas:963`) — qué cambia exactamente en este panel y si respeta la separación «Elegibles/NoSeleccionado».
- El `m_ObjUsuarioActivo Is Nothing` (`:121-123`) en `Form_Load` no se da porque `Form_FormUsuario00Principal.Form_Load:32-35` ya garantiza el usuario activo antes de cargar el lateral, pero conviene declararlo como invariante.
