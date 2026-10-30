# Form_FormGestionRiesgosSuministradores

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormGestionRiesgosSuministradores.cls` (553 líneas) y `src/forms/FormGestionRiesgosSuministradores.form.txt` · **PRD:** [README](../README.md)

Subformulario de **suministradores de la edición**: da de alta los suministradores del expediente en la gestión de riesgos, lista los ya registrados y permite cambiarles el indicador de gestión de calidad o retirarlos. Es el único de los cuatro que **elimina** registros. Las citas son `Form_FormGestionRiesgosSuministradores.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormGestionRiesgos` | Pestaña «Suministradores»; el contenedor la habilita o deshabilita según el estado del alta | `Form_FormGestionRiesgos.cls:198`, `:329`, `:344`, `FormGestionRiesgos.form.txt:611` |
| `FormSuministrador` | Se abre para consultar la ficha del suministrador seleccionado | `:361` |

Dos funciones son públicas y reutilizables desde el contenedor u otros formularios: `EstablecerCombo` (`:131`) y `RefrescarSuministradoresExpediente` (`:194`), ambas pensadas para invalidar cachés y repoblar el desplegable.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Carga; ante error avisa y **cierra el subformulario** | `:378-409` |
| `Suministrador` | `ComboBox` | `Change` | Habilita o deshabilita el botón de ver ficha | `:545-552`, `.form.txt:602` |
| `GestionCalidad` | `ComboBox` | — | Valor con el que se da de alta el suministrador | `.form.txt:689` |
| `LIstaSuministradores` | `ListBox` de 3 columnas | `Click` | Selecciona el vínculo de proyecto y habilita los comandos | `:411-463`, `.form.txt:565`, `:571` (`ColumnCount =3`) |
| `ComandoAlta` | botón | `Click` | Registra el suministrador elegido | `:78-122`, `.form.txt:853` |
| `ComandoCalidad` | botón | `Click` | Alterna Sí/No en la gestión de calidad | `:221-282`, `.form.txt:890` |
| `ComandoEliminar` | botón | `Click` | Pide confirmación y elimina el vínculo | `:284-339`, `.form.txt:929` |
| `ComandoVerSuministrador` | botón | `Click` | Abre la ficha del suministrador | `:341-376`, `.form.txt:731` |
| `ComandoActualizar`, `ComandoActualizarSuministradores` | botones | `Click` | Refrescan la lista o el desplegable invalidando cachés | `:20-45`, `:47-76`, `.form.txt:644`, `:967` |

## Flujos

### Apertura (`EstablecerDatos`, `:465-507`)

1. Si falta el proyecto activo **o** el proyecto del inicio, avisa «Es una edición y no se conoce el proyecto activo» (`:472-475`).
2. Calcula `blnPermitidoEscribir`: **los técnicos nunca escriben**; el resto solo si el proyecto está activo (`:476-483`).
3. Pobla el desplegable y la lista (`:484-492`).
4. Habilita `ComandoAlta` solo si puede escribir (`:493-499`).

### Poblado del desplegable (`EstablecerCombo`, `:131-187`)

Sale sin error si no hay proyecto activo, sin expediente o sin identificador de expediente (`:148-158`). Llama a `Constructor.getExpedienteSuministradores_RC` (`:161`) y añade una fila por suministrador con el identificador y el nombre: usa el nemotécnico si lo hay y, si no, el nombre con los puntos y comas sustituidos por dos puntos (`:169-179`). Esa sustitución evita que un nombre con separador desalinee las columnas del desplegable.

### Poblado de la lista (`EstablecerLista`, `:509-543`)

Lee `m_ObjProyectoActivo.ColSuministradores` (`:519`) y propaga el error del objeto (`:520-523`). Añade una fila por vínculo con identificador, nombre limpio y valor de gestión de calidad (`:528-534`), selecciona la primera fila y dispara el manejador de selección (`:535-536`).

### Selección (`LIstaSuministradores_Click`, `:411-463`)

Deshabilita los dos comandos de acción (`:418-419`); si la fila está vacía, limpia la selección y sale (`:420-425`). Si cambia de fila, recarga el vínculo (`:427-439`). Solo con vínculo resuelto y `blnPermitidoEscribir` habilita calidad y eliminar (`:446-449`).

### Alternar calidad (`ComandoCalidad_Click`, `:221-282`)

1. Resuelve el vínculo si no lo tiene (`:235-243`).
2. Calcula el valor contrario a «Sí»/«No»; si el valor actual no es ninguno de los dos, avisa y no sigue (`:247-253`).
3. Llama a `CambiarCalidad` (`:255`).
4. Sustituye la línea de la lista con `Replace` sobre el `RowSource` (`:262`) y **recarga el proyecto completo** desde `Constructor` (`:264`).

### Baja (`ComandoEliminar_Click`, `:284-339`)

Resuelve el vínculo, **apaga el reloj de arena antes de preguntar** (`:303-305`) y muestra una confirmación (`:306`); si no se acepta, sale con el reloj ya apagado (`:307-309`). Si se acepta, lo vuelve a encender (`:310-312`), elimina el vínculo (`:313`), limpia tres cachés y repuebla la lista (`:317-327`).

### Alta (`ComandoAlta_Click`, `:78-122`)

Resuelve el suministrador del desplegable (`:86-90`), avisa si no hay ninguno (`:91-93`), lo registra con el valor de gestión de calidad (`:94-99`), invalida las cachés de proyecto, proyecto del inicio y edición activa (`:100-102`) y repuebla la lista (`:103-112`).

## Reglas, validaciones y cálculos

- **El técnico es solo lectura, siempre** (`:476-477`). No hay ninguna otra condición de rol.
- **La gestión de calidad solo admite «Sí» o «No»**: cualquier otro valor es un error de datos, no un caso válido (`:247-253`).
- **Se exige identificador de expediente** para poblar el desplegable (`:154-158`): sin él, el desplegable queda vacío sin aviso.
- **Los separadores de las cadenas se limpian** con `Replace(..., ";", ":")` en el desplegable (`:175`) y en la lista (`:531`).

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Suministradores del expediente | `Constructor.getExpedienteSuministradores_RC` | `:161` |
| Vínculos del proyecto | `m_ObjProyectoActivo.ColSuministradores` | `:520` |
| Vínculo seleccionado | `Constructor.getSuministradorEnProyecto` | `:293`, `:428`, `:434` |
| Ficha del suministrador | `Constructor.getSuministrador` | `:86`, `:349` |
| Alta | `Proyecto.RegistrarSuministrador` | `:94` |
| Cambio de indicador | `ProyectoSuministrador.CambiarCalidad` | `:255` |
| Baja | `ProyectoSuministrador.Eliminar` | `:313` |

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** alta de los suministradores del expediente en la gestión de riesgos mediante un desplegable filtrado por expediente, con un indicador opcional de gestión de calidad; lista de los ya registrados con tres columnas; alternancia del indicador de calidad y eliminación del vínculo con confirmación. Los técnicos no pueden modificar nada.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-349.** El mensaje de confirmación de la baja (`:306`) está mal construido: termina sin punto y con la coletilla «…de la gestión de riesgoSí», donde el «Sí» es el valor por defecto de la respuesta incrustado sin separador. El texto que ve el usuario es literalmente incorrecto.
- **H-GR-350.** En `ComandoEliminar_Click` el reloj de arena se apaga (`:304`) **antes** del `MsgBox` (`:306`) y solo se vuelve a encender tras aceptar (`:311`). Si el usuario cancela, se sale con el reloj apagado; si el `MsgBox` tarda, la aplicación aparenta estar libre mientras espera una decisión modal.
- **H-GR-351.** `Err.Raise 1000` no lleva `vbObjectError` en ningún punto del módulo.
- **H-GR-352.** La Invalidación de cachés está **duplicada**: `RefrescarSuministradoresExpediente` (`:194-220`) hace casi lo mismo que `ComandoActualizarSuministradores_Click` (`:47-76`) y que el bloque de `ComandoAlta_Click` (`:100-104`), con ligeras diferencias: el alta invalida además `EdicionActiva.ColSuministradores`.
- **H-GR-353.** El comentario de `EstablecerCombo` (`:157`) anuncia una «lógica jerárquica retrocompatible» que selecciona modelo viejo o nuevo según una TempVar, pero el cuerpo de la función **no lee ninguna TempVar**: la lógica parece haber desaparecido del código sin que el comentario se actualice.
- **H-GR-354.** `EstablecerLista` llama a `LIstaSuministradores_Click` sin argumentos y sin `Call` (`:536`), lo que en VBA funciona solo porque no hay parámetros; es frágil ante cualquier cambio de firma.
- **H-GR-355.** `EstablecerLista` fija `p_Error = m_ObjProyectoActivo.Error` (`:520`) leyendo una propiedad de error global del proyecto. Cualquier error pendiente de otra operación se manifiesta aquí como si fuera de esta lista.
- **H-GR-356.** `ComandoCalidad_Click` alterna el valor leyendo el estado actual y calcula el contrario (`:247-253`). Es una alternancia calculada sobre datos persistidos: dos usuarios que pulsen a la vez se pisan, y el botón no refleja el estado, lo invierte a ciegas.
- **H-GR-357.** `ComandoActualizar_Click` (`:20-45`) invalida solo `m_ObjProyectoAlInicio.ColSuministradores` y repuebla la lista, mientras que `ComandoActualizarSuministradores_Click` invalida además el expediente del proyecto activo. Dos botones de nombre casi idéntico con alcances distintos, y ninguno de los dos está etiquetado con claridad en el diseño.
- **H-GR-358.** `blnPermitidoEscribir` se usa para habilitar el botón de alta (`:493-499`) pero los botones de calidad y eliminar se habilitan en el manejador de selección (`:446-449`); si el permiso cambia después de cargar, la vista queda con estados mixtos.
- **H-GR-359.** `Suministrador_Change` (`:545-552`) y el resto del módulo usan `On Error Resume Next`, mientras que los manejadores de acción usan `On Error GoTo`. La política de errores no es uniforme.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: sustituir el desplegable con la limpieza de separadores por un buscador de suministradores con identificador y nombre por separado; modelar la gestión de calidad como un interruptor visible en lugar de un botón que alterna a ciegas; eliminar la invalidación de cachés explícita en favor de una recarga coherente por petición; unificar los dos botones de actualizar en uno solo con alcance declarado; definir la baja como baja lógica en el modelo de web y reflectingla en la lista sin recargar todo el expediente.

## Criterios de aceptación

Propuestos, no aprobados:

1. El desplegable ofrece los suministradores del expediente con identificador y nombre legible.
2. Se puede dar de alta un suministrador con su indicador de gestión de calidad y la lista se actualiza al terminar.
3. Se puede cambiar el indicador de calidad del vínculo seleccionado y el cambio se refleja en la lista.
4. La eliminación pide confirmación, borra el vínculo y actualiza la lista.
5. Un usuario técnico ve la información pero no puede modificarla.
6. Si falta el expediente, la vista lo indica en lugar de mostrar un desplegable vacío.

## Pendientes

- Por qué `Constructor.getExpedienteSuministradores_RC` lleva el sufijo RC y qué variante jerárquica aplica realmente.
- Qué ocurre con el resto de cachés cuando un suministrador cambia en otra sesión.
- Si el botón de ver ficha debería recibir el suministrador seleccionado o el del desplegable: hoy toma el del desplegable (`:349`).
- Cómo se representa la lista en la web: como tabla de tres columnas o como tarjetas (ver [forms](../../forms.md)).