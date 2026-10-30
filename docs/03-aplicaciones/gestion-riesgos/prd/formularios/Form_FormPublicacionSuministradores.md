# Form_FormPublicacionSuministradores

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormPublicacionSuministradores.cls` (378 líneas) y `src/forms/FormPublicacionSuministradores.form.txt` · **PRD:** [README](../README.md)

Pantalla de la pestaña de suministradores dentro de la parte técnica de la publicación. Lista los suministradores dados de alta por Calidad en la edición y permite al técnico adjuntar o retirar el acuerdo de riesgos con cada uno. Es un listado corto: no pagina, no busca y no tiene columna de datos visibles más allá de nombre y si hay anexo. Las citas son `Form_FormPublicacionSuministradores.cls:línea` salvo indicación.

## Entrada y salida

No acepta parámetros ni `OpenArgs`. Toma el contexto de la global `m_ObjEdicionActiva` (`:29`) y **descarta su caché** de suministradores antes de releerla (`:28`). El permiso de escritura no lo calcula: hereda la global `blnPermitidoEscribir`, que fija el contenedor de la pestaña (`:348`, `:356`).

No abre ningún formulario. Para ver un anexo **no navega**: lo abre con el visor local del sistema (`:294`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Carga los datos y devuelve el error al llamador | `:214-242` |
| `lbl1` | etiqueta | — | Rótulo: «Suministradores registrados por Calidad» | `.form.txt:864-865` |
| `LIstaSuministradores` | cuadro de lista | `LIstaSuministradores_Click` | Resuelve la selección y enciende o apaga los botones | `:313-378`, `.form.txt:631, 637` |
| `ComandoAnexar` | botón | `ComandoAnexar_Click` | Selecciona fichero y lo anexa | `:62-140`, `.form.txt:731-734` |
| `ImagenAnexo` | imagen | `ImagenAnexo_Click` | Abre el anexo del seleccionado en local | `:270-311`, `.form.txt:665-667` |
| `ComandoEliminarEvidenciaAcuerdoSuministrador` | botón | `…_Click` | Pide confirmación y retira el anexo | `:142-212`, `.form.txt:797-800` |

Los tres botones tienen `Caption = "Comando47"` en el diseño (`.form.txt:732`, `:798`), es decir, **el rótulo visible es el marcador de Access**, no un texto. La identificación la da el globo.

## Flujos

### Apertura (`Form_Open`, `:214-242`)

1. `m_Error` a vacío, reloj de arena (`:216-222`).
2. `EstablecerDatos m_Error` (`:224`), que a su vez llama a `EstablecerLista` (`:253-256`).
3. No asigna `Cancel`: **el formulario se abre aunque la carga haya fallado**, y el único rastro es el `MsgBox`.

`EstablecerLista` (`:17-60`):

1. Invalida la caché compartida de suministradores de la edición y la vuelve a pedir (`:28-30`).
2. Si la propiedad devuelve error, lo propaga (`:30-33`).
3. Fija el `RowSource` a tres columnas: `ID;Nombre;Anexo` (`:34`). **El diseño trae otro `RowSource` de muestra, con cuatro columnas y datos de personas** (`.form.txt:634-636`), que se sobrescribe aquí.
4. Si la colección es `Nothing`, sale sin poblar (`:35-37`).
5. Recorre la colección y añade una fila por suministrador (`:38-49`). El indicador de anexo se calcula con `IDAnexo = ""` (`:40`), no con el propio objeto `Anexo`.
6. Sustituye los **puntos y coma del nombre por dos puntos** (`:46`) porque el `;` es el separador del `RowSource`.
7. Selecciona la primera fila (`:50`) y dispara a mano el manejador del clic (`:51`).

### Normal: selección (`:313-378`)

1. Oculta la imagen de anexo y el botón de eliminar (`:320-321`).
2. Si la columna 0 está vacía, suelta el objeto y sale (`:324-328`).
3. Resuelve el objeto seleccionado: lo carga si no lo tiene (`:330-333`) o si el identificador **ha cambiado respecto al guardado** (`:335-339`).
4. Si no lo encuentra, sale sin tocar nada más (`:344-347`).
5. `ComandoAnexar` se habilita con `blnPermitidoEscribir` (`:348-352`).
6. Si hay anexo, muestra la imagen y el botón de eliminar, este último solo si además hay permiso (`:353-360`).

### Normal: anexar (`:62-140`)

1. Si no hay objeto seleccionado, lo carga desde la columna 0 con `Nz` (`:79`); si no hay ninguno, avisa «Seleccione un elemento de la lista» (`:83-86`).
2. Abre el diálogo de fichero (`:91`). **Si el usuario cancela, `Exit Sub` sin más** (`:95-98`): es el único camino que sale sin reloj de arena puesto, y es correcto.
3. Captura la fila actual antes de tocar nada (`:101-102`).
4. `AnexarPorTecnico m_URL` (`:104`) y recarga el objeto (`:108`).
5. Refleja el resultado **sustituyendo texto dentro del `RowSource`** (`:112-121`): compone la fila nueva y hace `Replace` de la vieja, y llama a `Requery`.
6. Vuelve a disparar el manejador del clic (`:124`).

### Alternativo: eliminar (`:142-212`)

Mismo arranque que anexar; a partir de ahí pide confirmación con el botón por defecto en **No** (`:169-172`), llama a `EliminarEvidencia` (`:178`), recarga y repite la sustitución de texto (`:193-198`). En esta ruta se hace `SetFocus` en la lista antes de reescribir el `RowSource` (`:193`), cosa que no hace la de anexar.

### Normal: ver anexo (`:270-311`)

Si no hay objeto, lo carga (`:280-283`); si el seleccionado no tiene anexo, avisa (`:288-291`) y `AbrirEnLocal` con la URL del anexo y el `hWnd` del formulario (`:294`).

### Error

Los cuatro manejadores repiten el mismo patrón (`:131-139`, `:204-212`, `:231-240`, `:302-310`): quitan el reloj de arena y, si `Err.Number <> 1000`, **envían correo al administrador** y muestran crítico; si es exactamente 1000, muestran advertencia con el texto ya preparado. Los dos ayudantes (`EstablecerLista`, `EstablecerDatos`) solo describen el error cuando **no** es 1000 (`:56-59`, `:259-262`), para que el texto travel ya puesto no se pierda.

## Reglas, validaciones y cálculos

- El permiso no es de este formulario. `blnPermitidoEscribir` lo decide el contenedor de la pestaña; aquí solo se traduce en `Enabled` de dos botones (`:348`, `:356`). Con el permiso a `False` se puede **ver** el anexo pero no tocarlo.
- La columna 0 es la clave y no se ve: `ColumnWidths = "0;5670;1701"` (`.form.txt:635`). El identificador viaja oculto en el `RowSource`.
- El estado «Sí/No» del anexo se decide por `IDAnexo` vacío o no (`:40`), y después los dos acciones recalculan el mismo valor leyendo el objeto `Anexo` (`:112-115`, `:186-189`). Son dos definiciones del mismo dato en el mismo archivo.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Suministradores de la edición | `m_ObjEdicionActiva.ColSuministradores`, con caché invalidada | `:28-30` |
| Objeto del seleccionado | `Constructor.getSuministradorEnEdicion` | `:79`, `:108`, `:156`, `:182`, `:280`, `:330`, `:336` |
| Anexar evidencia | `EdicionSuministrador.AnexarPorTecnico` | `:104` |
| Retirar evidencia | `EdicionSuministrador.EliminarEvidencia` | `:178` |
| Abrir el anexo | ruta del anexo, en local | `:294` |

## Coexistencia

Pendiente de la capa de coexistencia, sin enlace desde esta ficha. El punto crítico es `AbrirEnLocal` (`:294`): el técnico abre una copia en su equipo, no la vista web del documento. Mientras convivan, el anexo tiene que existir en un sitio que el cliente alcance, lo que ata esta pantalla a la ruta de ficheros del usuario.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-320.** El diseño del cuadro de lista guarda un `RowSource` de muestra con **nombres, identificadores de usuario y correos de personas concretas** (`.form.txt:634-636`), en columnas que ni siquiera coinciden con las tres que fija el código. Dato personal incrustado en el artefacto exportado; no se reproduce aquí. Conviene señalarlo al propietario para decidir si se limpia del binario.
- **H-GR-321.** `Err.Raise 100` en `:296` es el único del archivo que **no** usa el código 1000. Como consecuencia, un fallo banal al abrir un anexo entra en la rama de error real: **manda correo al administrador** y el `MsgBox` crítico muestra número y descripción de Access, **descartando el `m_Error`** que el mensaje preparado sí explicaba.
- **H-GR-322.** El `RowSource` se actualiza con `Replace` sobre el texto completo (`:120`, `:196`). Si dos suministradores comparten nombre, o el identificador aparece dos veces, la sustitución actúa sobre **la primera coincidencia de la cadena**, que puede no ser la fila seleccionada: la lista puede mostrar un estado de anexo atribuido a la persona equivocada.
- **H-GR-323.** Tras `Requery` (`:121`, `:198`) se pierde la selección, y el manejador del clic vuelve a leer `Column(0)` de la fila que quede por defecto (primera). Si el procedimiento actuó sobre una fila distinta, la interfaz salta a la primera sin avisar.
- **H-GR-324.** `EstablecerLista` selecciona la primera fila sin comprobar que la colección tenga elementos (`:50`). Si el diccionario llega vacío pero no `Nothing`, `Selected(1) = True` falla y el error sale por el ayudante como mensaje de método, no como «no hay suministradores».
- **H-GR-325.** `EstablecerLista` y `EstablecerDatos` se declaran `As String` y **nunca devuelven valor**: el resultado viaja por el `ByRef p_Error`. La firma sugiere un contrato que no existe.
- **H-GR-326.** `Set m_ObjEdicionActiva.ColSuministradores = Nothing` (`:28`) invalida una caché **compartida** cada vez que se abre la pestaña: otro formulario que la esté usando ve cómo se reconstruye debajo.
- **H-GR-327.** `Form_Open` no asigna `Cancel` (`:214-242`): si `EstablecerDatos` falla, el formulario se abre vacío y el único aviso es un `MsgBox` que no impide seguir trabajando sobre una lista no cargada.
- **H-GR-328.** Los tres botones tienen `Caption = "Comando47"` (`.form.txt:732`, `:798`), así que en pantalla aparece el nombre interno del botón y no su función.
- **H-GR-329.** El estado del anexo se calcula de dos maneras distintas (`:40` por `IDAnexo`, `:112` y `:186` por el objeto `Anexo`). Si alguna vez divergen, la lista y la base de datos discreparán sin aviso.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: devolver un listado paginado y buscable en lugar de un `RowSource` de texto; actualizar la fila por identificador en el servidor en lugar de rescribir la lista completa; servir el anexo desde una ruta autenticada en lugar de abrirlo en local; marcar como desconocido el estado del anexo cuando no se pueda determinar, en lugar de asumir que no existe.

## Criterios de aceptación

Propuestos, no aprobados:

1. La lista muestra los suministradores de la edición con su nombre y si tienen o no evidencia.
2. Sin permiso de escritura se puede ver la evidencia pero no anexarla ni retirarla.
3. Anexar y retirar deja la lista coherente con lo guardado, sin salirse de la fila accionada.
4. Si no hay suministradores, se dice explícitamente en lugar de mostrar un error.
5. Cancelar el diálogo de fichero no deja la pantalla en estado intermedio.

## Pendientes

- Qué hace exactamente `AnexarPorTecnico` y `EliminarEvidencia`: si anexan a base de datos, a sistema de ficheros o a ambos, y qué ocurre con el fichero original del técnico.
- Si la invalidación de caché (`:28`) es necesaria por un problema conocido de la clase `EdicionSuministrador` o es una precaución heredada.
- Qué permisos tiene el usuario técnico frente al de Calidad sobre esta misma lista, más allá del interruptor global.