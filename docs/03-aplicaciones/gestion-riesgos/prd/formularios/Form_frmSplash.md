# Form_frmSplash

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_frmSplash.cls` (102 líneas) y `src/forms/frmSplash.form.txt` · **PRD:** [README](../README.md)

Pantalla de arranque. Oculta la cinta de opciones, lleva la barra de progreso mientras se calcula el entorno, y decide qué menú abre: el de técnico o el de calidad/administración. No escribe nada. Las citas son `Form_frmSplash.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| **Nadie en `src`** | Ningún `DoCmd.OpenForm` de `src` la abre; solo se comprueba que esté abierta para mover la barra | `Funciones Generales.bas:4891` |

Es, con toda probabilidad, el formulario de arranque de Access (propiedad «Formulario de inicio» del `.accdb`) o lo abre una macro, que no está en el código fuente. **Es la única pieza del arranque que no se puede verificar leyendo el código**: queda pendiente de confirmación en el binario. La contradicción con el comentario de `Funciones Generales.bas:9064` («debe ser llamado desde el evento `Form_Open` del `frmSplash`») no se resuelve aquí.

No acepta parámetros. Lo que sí usa es `Command()`, que lee la línea de comandos de Access (`:35`).

Salidas: abre `Form0BDOpcionesTecnico` y, si tiene tareas, `FormTecnicoTareas` (`:74, 76`), o bien `Form0BDOpciones` y, si hay tareas pendientes, `FormCalidadTareas` (`:83, 85`). En ambos caminos se cierra a sí misma (`:92`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Fija la TempVar `CadenaJerarquicaModelo` a `«nuevo»` | `:22-30` |
| (formulario) | — | `Form_Open` | Oculta la cinta, ajusta el tamaño, pone el contador de pasos a cero y deja la barra a ancho 0 | `:31-46` |
| (formulario) | — | `Form_Timer` | Todo el arranque | `:48-102`, `.form.txt:28` (`TimerInterval =50`) |
| `lblEstado` | etiqueta | — | Texto de progreso («Cargando configuración…», «Analizando y calculando tareas pendientes…») | `:58, 67`, `.form.txt:707` |
| `lblProgresoBarra` | etiqueta | — | Barra de progreso; la ancho son `FormularioAbierto("frmSplash")` y `s_contadorPasos` desde fuera | `:41`, `.form.txt:1037`; `Funciones Generales.bas:4891-4900` |

## Flujos

### Apertura (`Form_Load`, `:22-30`)

`On Error Resume Next`. Si la TempVar `CadenaJerarquicaModelo` no existe la crea con `«nuevo»`; si existe, **la sobrescribe** con `«nuevo»`. No lee el valor anterior.

### Apertura (`Form_Open`, `:31-46`)

1. Si `Command()` **no** contiene `/dev`, oculta la cinta con `GestionarRibbon(False)` (`:35-37`; `Funciones Generales.bas:9068`, que a su vez usa `On Error Resume Next`).
2. `AjustarTamaño Me` (`:39`; `Funciones Generales.bas:1043`).
3. `ResetAvanceCounter` pone `s_contadorPasos` a cero (`:40`; `Funciones Generales.bas:9062`).
4. Deja `lblProgresoBarra` a ancho 0 (`:41`).

El manejador de error muestra un `MsgBox` y **no cancela la apertura** (`Cancel` no se asigna).

### Arranque (`Form_Timer`, `:48-102`)

1. Detiene el temporizador para que no se repita (`:54`).
2. `EVE m_Error` (`:61`; `Variables Globales.bas:204`). Si devuelve error, `Err.Raise 1000`.
3. Recupera el último proyecto del usuario (`:69`).
4. **Ramifica por rol** (`:71`):
   - **Técnico** (`EsTecnico`): `EstablecerTareasTecnico` y abre `Form0BDOpcionesTecnico`; si `TareasTotales <> "0"`, abre además `FormTecnicoTareas` (`:71-77`).
   - **Resto**: guarda `ListaUsuarios`, llama a `EstablecerTareasCalidad` y abre `Form0BDOpciones`; si `HayTareasPendientes = Sí`, abre además `FormCalidadTareas` (`:78-87`).
5. Se cierra (`:92`).

## Reglas, validaciones y cálculos

- **El rol decide el menú, no un permiso.** La condición es `EsTecnico`; no se comprueba ni administrador ni calidad en ningún punto de este formulario. El control de rol ocurre después, en `Form0BDOpciones`.
- **La rama de técnico compara contra un texto:** `m_ObjTareasTecnico.TareasTotales <> "0"` (`:75`), mientras la de calidad compara contra la enumeración (`EnumSiNo.Sí`, `:84`). La misma idea con dos convenciones distintas.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Entorno, usuario, cachés, tareas | TempVars y objetos de sesión | `Variables Globales.bas:204` (`EVE`) |
| Último proyecto del usuario | `Constructor.getUltimoProyecto` | `:69` |
| Lista de usuarios conectados | `m_ObjEntorno.ListaUsuarios` | `:79` |
| Barras y tareas | `s_contadorPasos` (global de módulo) | `Funciones Generales.bas:4891-4900` |

Sin escrituras en base de datos.

## Coexistencia

Pendiente de `../coexistencia.md`. Aquí ya se ve: el arranque depende de `EVE` (`Variables Globales.bas:204`), que es donde se resuelven TempVars, identidad y cachés; y `FormularioAbierto("frmSplash")` funciona porque Access mantiene una colección de formularios abiertos, sin equivalente en una web.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-1.** El manejador de error de `Form_Timer` cierra el formulario sin abrir nada (`:101`), pero el comentario inmediatamente anterior dice «Intentamos abrir el principal de todas formas por seguridad» (`:99`). Si `EVE` falla en producción, el usuario se queda sin menú y sin mensaje que lo explique más allá del `MsgBox`. Conviene comprobar si hay otro camino de arranque.
- **H-GR-2.** `Err.Raise 1000` en `:82` va sin tercer argumento, al revés que las cuatro anteriores (`:62, 70, 73`), que sí pasan `m_Error`: en la rama de calidad, el error raised **pierde la descripción** que el mensaje de `:97-98` intentaba mostrar.
- **H-GR-3.** `Err.Raise 1000` en `:62` no lleva `vbObjectError`, así que Access podría tratarlo como error de usuario y **no ejecutar** el manejador; conviene verificar en ejecución si el splash puede quedarse en blanco.
- **H-GR-4.** `Form_Load` sobrescribe `CadenaJerarquicaModelo` a `«nuevo»` sin mirar el valor previo (`:25-29`). El [estudio previo](../../capabilities.md) documenta que ese flag admite dos modelos de carga del árbol («nuevo» y «antiguo») y que **ambos están activos en producción**; desde este arranque solo se puede alcanzar «nuevo».
- **H-GR-5.** La comparación `TareasTotales <> "0"` (`:75`) depende de que la propiedad devuelva exactamente el texto `0`. Con un entero o un valor vacío, la comparación es distinta.
- **H-GR-6.** El salto por `/dev` (`:35`) desactiva la ocultación de la cinta con cualquier orden que contenga esa cadena. Es un mecanismo de desarrollo embebido en el arranque de producción.
- **H-GR-7.** `GestionarRibbon` usa `On Error Resume Next` (`Funciones Generales.bas:9071`): si `ShowToolbar` falla, la cinta queda visible y el arranque continúa sin avisar.
- **H-GR-8.** `Form_Open` no asigna `Cancel` en su manejador de error (`:44-45`): el formulario se abre aunque `AjustarTamaño` haya fallado.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: barra de progreso real en el arranque en lugar de una etiqueta; derivar el menú de un atributo del usuario resuelto en el servidor en vez de una TempVar; eliminar el concepto de modelo de árbol conmutable y dejar un único modelo.

## Criterios de aceptación

Propuestos, no aprobados:

1. Arrancar muestra el progreso y abre el menú que corresponde al rol.
2. Un usuario con tareas pendientes de su perfil ve además su formulario de tareas.
3. Un usuario sin tareas pendientes no ve ese formulario.
4. Si la inicialización falla, se avisa y no se deja al usuario sin salida.

## Pendientes

- Cómo se abre realmente el formulario: propiedad de inicio de Access o macro, que no están en el código fuente.
- Qué hace `EVE` con exactitud (`Variables Globales.bas:204`) y qué parte del cálculo de tareas se apoya en caché de disco.
- Si `/dev` está activo en producción.