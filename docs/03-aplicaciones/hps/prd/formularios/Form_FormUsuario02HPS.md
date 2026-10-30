# Form_FormUsuario02HPS

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormUsuario02HPS.cls` (988 líneas) y `.form.txt` (4127 líneas) · **PRD:** [README](../README.md)

Subpantalla central de la ficha de usuario: gestiona los **cuatro tipos de HPS** (Nacional, OTAN, UE y ESA), cada uno con su grado, su estado, sus fechas y su marca de renovación. Incluye el atajo «pasar al resto», que copia el bloque Nacional sobre los otros tres. Es el formulario con más controles del PRD (90) y el que peor escala de esta migración: su lógica son cuatro bloques casi idénticos escritos a mano. Formulario **no enlazado**: los valores viajan por reflexión con `RellenarHPSConDatosFormularioPorTIpoHPS` (`:653`). Las citas son `Form_FormUsuario02HPS.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormUsuario00Principal` | Se abre con el usuario ya establecido en el global `m_ObjUsuarioActivo`; no recibe argumentos | `:577-581` |
| `Form_FormUsuario02Alta` | Tras dar de alta, el contenedor abre esta pantalla con el usuario ya activo | — |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Exige usuario activo, aplica el modo lectura/escritura, rellena los cuatro combos de grado y decide si habilita «pasar al resto» | `:568-651` |
| `HPS_NAC_Grado` | combo | `Change` → `HPS_NAC_Grado_Change` | Revalida la habilitación de «pasar al resto» cuando cambia el grado Nacional | `:980-987` |
| `HPS_NAC_Grado`, `HPS_OTAN_Grado`, `HPS_UE_Grado`, `HPS_ESA_Grado` | combos | — | Grado de cada tipo; se rellenan con `RellenarComboGrado` | `:611-624`, `FG.bas:1206` |
| `HPS_*_Renovacion` | texto | 4 botones `Click` | Conmutan el valor entre `"X"` y `Null` | `:462-503` |
| `ComandoHPSNACRenovacion`, `ComandoHPSOTANRenovacion`, `ComandoHPSESARenovacion`, `ComandoHPSUERenovacion` | botones | `Click` | Alternan la marca de no renovación de cada tipo | `:462-503` |
| `ComandoPasarAlResto` | botón | `Click` | Copia el bloque Nacional sobre OTAN, UE y ESA | `:505-567` |
| `ComandoGrabar` | botón | `Click` | Detecta cambios por tipo y registra | `:88-329` |
| `ComandoEliminarNacional`, `ComandoEliminarOTAN`, `ComandoEliminarUE`, `ComandoEliminarESA` | botones | `Click` | Baja del bloque de cada tipo | `:330-461` |
| Bloque `HPS_<TIPO>_*` (≈20 campos por tipo) | textos y combos | — | Estado, motivo de irregularidad, fechas de solicitud, baja, concesión y caducidad, especialidad | `.form.txt` |

## Flujos

### Apertura (`Form_Load`, `:568-651`)

1. Pone el reloj de arena (`:573-575`).
2. Exige `m_ObjUsuarioActivo` (`:576-580`).
3. Decide el modo con `Application.TempVars("SoloLectura")`: desactiva `AllowEdits` y nueve botones, o activa `AllowEdits` y los cuatro de renovación (`:582-598`).
4. Aplica la paleta de colores (`:599-603`).
5. Copia el estilo de seis botones (`:604-609`).
6. Rellena los cuatro combos de grado con `RellenarComboGrado` (`:611-626`; `FG.bas:1206`), encadenando error en cada uno.
7. `EstablecerDatos` vuelca el objeto a la pantalla (`:627-630`; `:769-979`).
8. **Vuelve a leer `SoloLectura` y decide «pasar al resto»** según el grado Nacional (`:631-638`).
9. En error: correo al administrador, `MsgBox` crítico y cierre (`:639-651`).

### Cambiar el grado Nacional (`HPS_NAC_Grado_Change`, `:980-987`)

Repite la comprobación del punto 8 con la forma **correcta**: `Nz(valor, "") = "RESERVADO ( R )"` (`:982`). Es el mismo criterio que `Form_Load` aplica en `:632` con otra escritura.

### Marcas de renovación (`:462-503`)

Los cuatro botones hacen lo mismo: si el campo vale `"X"` lo ponen a `Null`, y si no, a `"X"`. **Solo el Nacional tiene manejador de error**: es el único con `On Error GoTo errores` (`:471`) y su bloque `errores:` (`:479-484`). Los otros tres —ESA (`:462-468`), OTAN (`:487-493`) y UE (`:495-503`)— no lo tienen.

### Pasar al resto (`ComandoPasarAlResto_Click`, `:505-567`)

1. Anula estado y motivo de irregularidad de OTAN, UE y ESA (`:512-513`, `:519-520`, y siguientes).
2. **Copia del bloque Nacional** las fechas de solicitud, renovación, baja, concesión, caducidad y especialidad (`:514-518` y equivalentes por tipo).
3. Guarda los cambios.

### Grabar (`ComandoGrabar_Click`, `:88-329`)

1. Llama a `HaHabidoCambiosPorTipoHPS` una vez por tipo (`:123`, `:145`, `:166`, `:187`; `:9-87`), que compara por reflexión los campos de ese bloque.
2. Si ningún tipo tiene cambios, aborta con «No hay cambios que guardar».
3. `RellenarHPSConDatosFormularioPorTIpoHPS` vuelca los controles al bloque de cada tipo (`:653-738`) y `RellenarHPSConDatosFormulario` el resto (`:739-768`).
4. Registra, avisa y vuelve a la pantalla anterior.

### Errores

Repiten `Err.Number <> 1000` ⇒ anexar descripción, `CorreoAlAdministrador` y `MsgBox` crítico. El retorno va al global `pregunta` (`VariablesEntorno.bas:220`).

## Reglas, validaciones y cálculos

- **Cuatro bloques idénticos escritos a mano.** Cada tipo repite estado, motivo, fechas y especialidad. `ComandoPasarAlResto` existe precisamente porque copiar a mano entre bloques es lo que se hacía, y automatiza solo la copia Nacional → resto, no las otras tres direcciones.
- **El modo lectura se decide dos veces**, con `TempVars` para los botones (`:582-598`) y otra vez para «pasar al resto» (`:631`).
- **La comprobación de grado está escrita de dos formas en el mismo fichero.** `Form_Load` usa `Nz(Me.HPS_NAC_Grado = "CONFIDENCIAL ( C )")`, que mete una **comparación booleana** dentro de `Nz` (`:632`), mientras `HPS_NAC_Grado_Change` usa `Nz(Me.HPS_NAC_Grado, "") = "RESERVADO ( R )"`, que es la forma correcta (`:982`). Mismo criterio, dos formas.
- **El botón «pasar al resto» depende solo del grado Nacional** (`:632`, `:982`): el grado de OTAN, UE o ESA no lo activa ni lo desactiva.
- **Todo el estado es global.** `m_ObjUsuarioActivo` y `pregunta` son variables públicas de módulo; el formulario no recibe identificadores.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Detección de cambios por tipo | reflexión sobre los controles de cada bloque | sin consulta | `:123`, `:145`, `:166`, `:187`, `:9-87` |
| Volcado de los cuatro bloques HPS | reflexión sobre el objeto `HPS` | sin consulta | `:653-738`, `:739-768` |
| Catálogo de grados | `RellenarComboGrado` | colección en memoria | `:611-624`, `FG.bas:1206` |
| Paleta y estilos | `EstablecerControlesDatos`, `CopiarEstiloBoton` | presentacional | `:599-609`, `FG.bas:1005`, `FG.bas:2435` |

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Aquí el problema no es el SQL sino la **repetición**: cuatro bloques de campos con la misma forma, comparados y volcados por reflexión uno a uno. En web eso es un array de cuatro objetos con un único render y un único guardado. `DoCmd.Hourglass` aparece en cinco procedimientos (`:574`, `:830`, `:981` y dos más) sin equivalente directo, y el patrón de `Nz(<comparación>)` de `:632` no debe trasladarse tal cual a una condición de plantilla.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **`Nz` con una comparación dentro en `:632`.** `Nz(Me.HPS_NAC_Grado = "CONFIDENCIAL ( C )")` pasa un booleano a `Nz`, que solo tiene sentido con `Null`. Funciona por casualidad (`Nz(False)` es `False`), pero si el grado es `Null` la comparación devuelve `Null` y `Nz` lo convierte en `0`, con lo que el resultado depende de la aritmética de Access. La misma comprobación en `:982` está bien escrita. A unificar.
- **Solo el botón Nacional de renovación tiene manejador de error** (`:471`, `:479-484`), frente a ESA (`:462-468`), OTAN (`:487-493`) y UE (`:495-503`), que no lo tienen. Si `HPS_OTAN_Renovacion`, `HPS_UE_Renovacion` o `HPS_ESA_Renovacion` fallan, no hay aviso ni correo: los tres son idénticos en el resto del código (`:488-492` y `:497-501`, la única diferencia entre ellos es el prefijo del campo).
- `ComandoPasarAlResto` copia Nacional → resto pero **no al revés**: un cambio en el grado de OTAN, UE o ESA no se propaga a Nacional, aunque el botón esté disponible con cualquier grado.
- Los cuatro bloques se recorren con un bloque de código por tipo, no con un bucle (`:123`, `:145`, `:166`, `:187` y los `Rellenar*` de `:653-738`). Un quinto tipo exigiría duplicar el fichero entero.
- `m_Error` es variable de módulo y los helpers la sobrescriben `ByRef` desde fuera.
- El cierre con `DoCmd.Close acForm, Me.Name, acSaveNo` del manejador de error (`:649`) descarta cambios sin preguntar.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: modelar los cuatro tipos como una lista con un único formulario y un único guardado; dejar que el traspaso entre tipos sea una operación simétrica y explícita; y sustituir la comparación por reflexión de `:632` y `:982` por una regla declarativa sobre el grado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin usuario activo, el formulario se abre explicando que faltan los datos y se cierra.
2. En modo solo lectura los campos no admiten escritura y los nueve botones destructivos quedan deshabilitados.
3. Los cuatro combos de grado muestran el catálogo vigente en los cuatro tipos.
4. Marcar «no renovación» en un tipo lo alterna entre activo e inactivo, y un error en cualquiera de los cuatro se comunica.
5. «Pasar al resto» copia estado, motivo y fechas del bloque Nacional a los otros tres, y solo queda disponible con grado RESERVADO, CONFIDENCIAL o SECRETO en Nacional.
6. Grabar sin cambios en ningún tipo no escribe y explica por qué.
7. Grabar con cambios en un solo tipo escribe únicamente ese bloque.

## Pendientes

- Qué decide si los datos bloqueados (`HPS_NAC_Renovacion`) se conservan al pasar a histórico.
- Contenido de `EstablecerDatos` (`:769-979`) y de `RellenarHPSConDatosFormulario` (`:739-768`), que no se ha desglosado en la ficha.
- `HaHabidoCambiosPorTipoHPS` (`:9-87`) compara ocho campos por tipo; cuáles considera no está en el resumen de la tabla.
- Si `Command5` / `Command6` del `.form.txt` tienen handler en el binario.
- Cómo decide `RellenarComboGrado` el catálogo por tipo (`FG.bas:1206`).