# Módulos de apoyo a formularios padre y de registro de cambios en JSON

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/modules/FormulariosPadreAuxiliares.bas` (273 líneas), `src/modules/JsonHelper.bas` (119) · **PRD:** [README](../README.md)

Dos módulos que traducen entre el vocabulario de la interfaz y el de los datos. `FormulariosPadreAuxiliares` resuelve qué subformulario corresponde a una pestaña y cómo se llama cada bloque; `JsonHelper` registra, dentro de un diccionario, qué campo cambió respecto a su valor original, que es la base del rechazo y de su comentario.

Abreviaturas: `FPA`, `modules/FormulariosPadreAuxiliares.bas`; `JH`, `modules/JsonHelper.bas`; `ENUM`, `modules/modEnumeradores.bas`; `CE`, `classes/CondorError.cls`.

## Superficie pública

| Objeto | Tipo | Procedimientos públicos | Qué resuelve |
|---|---|---|---|
| `FormulariosPadreAuxiliares` | módulo | `SonEntidadesIguales`, `ClonarEntidad`, `GetSubformNameFromTabName`, `MapearBloquePorNombreSub`, `GetNombreAmigableBloque`, `HabilitarGuardarEnSubform` (`FPA:25-273`) | Traducción entre pestañas, subformularios y bloques de datos; comparación y copia de entidades |
| `JsonHelper` | módulo | `RegistrarCambio`, `ObtenerJsonRechazo`, `GuardarJsonRechazo` (`JH:12-119`) | Diccionario de cambios por bloque y campo, y su volcado al rechazo |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `FPA` | `SonEntidadesIguales` | Compara dos objetos propiedad a propiedad | — | Sin efectos | `FPA:25-57` |
| `FPA` | `ClonarEntidad` | Devuelve una copia del objeto del tipo indicado | — | Sin efectos | `FPA:58-102` |
| `FPA` | `GetSubformNameFromTabName` | Devuelve el nombre del subformulario según la pestaña y el tipo de formulario | — | Sin efectos | `FPA:103-159`; doble `Select Case` en `FPA:104-113` |
| `FPA` | `MapearBloquePorNombreSub` | Devuelve el bloque de datos a partir del nombre del subformulario | — | Sin efectos | `FPA:160-219` |
| `FPA` | `GetNombreAmigableBloque` | Devuelve la etiqueta visible de un bloque | — | Sin efectos | `FPA:220-267` |
| `FPA` | `HabilitarGuardarEnSubform` | Activa o desactiva el botón de guardar del subformulario | — | Sin efectos | `FPA:268-273` |
| `JH` | `RegistrarCambio` | Crea o actualiza la entrada de un campo con su valor original y su valor final | `ExisteCampo` | Modifica el diccionario en memoria | `JH:12-77`; normalización con `Nz` en `JH:21-22` |
| `JH` | `ObtenerJsonRechazo` | Devuelve el diccionario de cambios de un rechazo | — | Sin escritura | `JH:89-112` |
| `JH` | `GuardarJsonRechazo` | Vuelca el diccionario al JSON del rechazo | — | Escribe mediante el repositorio de rechazo | `JH:113-119` |

## Flujos de trabajo

- **Normal:** el formulario padre resuelve la pestaña activa a subformulario (`FPA:103-159`), carga el bloque y, al guardar, el subformulario recoge los valores. El registro de cambios va por otro camino: cada campo modificado pasa por `RegistrarCambio`, que crea la entrada si no existe o conserva el valor original si ya existe (`JH:12-77`).
- **Alternativos:** `RegistrarCambio` distingue el caso «el usuario volvió al valor original» del caso «no había registro previo» (`JH:24-26`): son dos situaciones que el comentario del propio código describe como ambiguas.
- **Errores:** `FPA` no declara manejadores; el error lo captura el formulario. `JH` envuelve el volcado en la transacción que le pasa el llamador.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Los nombres de subformulario se derivan de la pestaña y del tipo, nunca se escriben a mano | `GetSubformNameFromTabName` | `FPA:103-113` |
| El valor original se fija en la primera anotación y no se sobrescribe | `RegistrarCambio` | `JH:15-16`, `JH:24-26` |
| Los nulos se normalizan a cadena vacía antes de anotar | `RegistrarCambio` | `JH:21-22` |
| La etiqueta visible de un bloque no coincide con su nombre interno | `GetNombreAmigableBloque` | `FPA:220-267` |

## Lecturas y escrituras

| Operación | Recurso | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Sin acceso a datos | — | — | — | `FPA:25-273` |
| Volcado del JSON de rechazo | tabla de rechazos | a través del repositorio de rechazo | La que pase el servicio | `JH:113-119` |

## Coexistencia

El módulo `FPA` describe la estructura de la interfaz: pestañas, subformularios y bloques. En el destino web el equivalente es el estado de navegación de la vista, de modo que este módulo no se traslada: se convierte en la definición de qué bloque se está editando y de cómo se nombra. El nombre del subformulario es hoy la clave que identifica el bloque (`FPA:103-159`), y esa clave —`subfrmDatosPC_Generales` y compañía— es la que aparece en las hojas de los subformularios. `JsonHelper` sí tiene equivalente: un objeto de cambios con valor original y valor final, cuyo destino es una columna JSON o una tabla de líneas de cambio. La decisión pendiente es si el JSON se conserva como columna o si se desnormaliza, porque el rechazo se lee con `ActualizarComentarioRechazoActivo` y `ActualizarJSON` (hoja de repositorios).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual.** `FPA` es una tabla de correspondencias entre pestañas, subformularios y bloques, más dos utilidades de copia y comparación de entidades. `JH` mantiene el diccionario de cambios y su volcado.

**Posibles defectos, a verificar.**
1. La correspondencia está codificada con `Select Case` anidados (`FPA:104-113`): añadir un bloque exige tocar el código.
2. `RegistrarCambio` documenta como ambigua la situación «el valor vuelve al original» (`JH:24-26`): el comportamiento observable depende de una decisión que no está escrita.
3. El diccionario se serializa y se guarda como texto JSON (`JH:113-119`): el esquema de destino tendría que soportar el tipo o desnormalizarlo.

**Propuestas web pendientes de decisión.** Que el bloque en edición sea estado explícito de la ruta y no un nombre de control. Que el conjunto de cambios sea una entidad de primera clase con su propio contrato, en lugar de un diccionario serializado.

## Criterios de aceptación

1. La pestaña activa determina el bloque de datos que se carga y se guarda. Trazado a `FPA:103-159`.
2. Un campo modificado por primera vez conserva su valor original aunque se modifique otra vez. Trazado a `JH:12-77`.
3. Los campos no modificados no aparecen en el registro de cambios. Trazado a `JH:12-77`.
4. El rechazo guarda el conjunto de cambios junto al motivo. Trazado a `JH:113-119`.

## Pendientes

- Confirmar contra el binario qué bloque se edita cuando el usuario vuelve al valor original (`JH:24-26`).
- Confirmar si `SonEntidadesIguales` y `ClonarEntidad` tienen consumidores vivos (`FPA:25-102`).
- Confirmar el tipo de columna donde acaba el JSON del rechazo, porque condiciona la matriz de campos.