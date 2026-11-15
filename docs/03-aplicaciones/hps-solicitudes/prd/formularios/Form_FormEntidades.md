# Form_FormEntidades

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormEntidades.cls` (170 líneas) y `.form.txt` · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `fichero:línea` del commit fijado.

Menú de acceso a los módulos de configuración del dominio HPS: lista las entidades gestionables y abre el formulario que le corresponde a cada una. Es un formulario **no enlazado** (la lista se puebla con `AddItem`, `.form.txt:643-645`) y **no escribe nada** en la base de datos.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpciones` | `DoCmd.OpenForm "FormEntidades"` solo si no estaba ya abierto | `Form_Form0BDOpciones.cls:137-138` |

Abre el formulario de la entidad elegida con `DoCmd.OpenForm` (`:61`) y se cierra a sí mismo después (`:63`). No declara eventos ni recibe parámetros: todo su contexto llega por la global `m_ObjEntorno` (`VG.bas:190`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaEntidades` | lista 2 columnas | `OnClick` | Habilita el botón de abrir cuando hay fila seleccionada | `:136-162`, `.form.txt:631-650` |
| `ListaEntidades` | lista 2 columnas | `OnDblClick` | Doble clic equivalente a pulsar el botón de abrir | `:164-169`, `.form.txt:648` |
| `ComandoAbrirFormulario` | botón | `OnClick` | Traduce la entidad a nombre de formulario, lo abre y se cierra | `:40-76`, `.form.txt:681-692` |
| `ComandoAyuda` | botón | `OnClick` | Abre el PDF de ayuda del formulario | `:78-100`, `.form.txt:543-553` |
| `cmdSalir` | botón | `OnClick` | Cierra el formulario | `:7-11`, `.form.txt:777-779` |
| `lblTitulo1` | etiqueta | — | Único rótulo; `Form_Open` copia su texto al título de la ventana | `:112`, `.form.txt:531-532` |

## Flujos

### Apertura (`Form_Open`, `:102-134`)

1. Dimensiona con `Ajustar Me` (`:110`; `FU.bas:735-758`) y copia el rótulo al título (`:112`).
2. Habilita el botón de abrir (`:114`).
3. `EstablecerLista` vacía la lista y la repuebla desde `m_ObjEntorno.Entidades` (`:115`; `:13-38`). Cada elemento es `clave;etiqueta` (`:30`), donde la clave es el valor de `EnumEntidades`.
4. Si `EstablecerLista` devuelve error, `Form_Open` avisa y cierra el formulario (`:116-118`, `:125-133`).

### Abrir una entidad (`ComandoAbrirFormulario_Click`, `:40-76`)

1. Deshabilita el botón (`:51`) y aborta con «No se ha seleccionado ninguna opción» si la columna 0 está vacía (`:52-55`).
2. Traduce la clave con `m_ObjEntorno.EntidadesFormularios` (`:56`) y aborta con «No se ha seleccionado ninguna entidad reconocida» si el resultado es vacío (`:57-60`).
3. Abre el formulario con `DoCmd.OpenForm` (`:61`), invoca `ListaEntidades_Click` (`:62`) y se cierra (`:63`).

### Alternativos

- `ListaEntidades_DblClick` (`:164-169`) solo actúa si el botón está habilitado; su manejador va con `On Error Resume Next`, sin notify.
- `ListaEntidades_Click` (`:136-162`) habilita el botón únicamente si hay fila seleccionada (`:145-148`).

### Errores

`ComandoAbrirFormulario_Click`, `ComandoAyuda_Click`, `Form_Open` y `ListaEntidades_Click` comparten el mismo patrón: `Err.Raise 1000` para el error interno ya redactado, `CorreoAlAdministrador` solo cuando `Err.Number <> 1000`, `DoCmd.Hourglass False` y `MsgBox` crítico (`:68-75`, `:92-99`, `:125-133`, `:154-161`).

## Reglas, validaciones y cálculos

- **Catálogo de entidades:** son cuatro, fijas en código, no en tabla: responsables, plantillas HTML, motivos HPS y justificaciones HPS (`Entorno.cls:180-208`). El diccionario se cachea en `m_ObjEntidades` la primera vez que se lee (`Entorno.cls:180-190`).
- **Traducción a formulario:** `EntidadesFormularios` mapea cada clave a un nombre de formulario (`Entorno.cls:210-238`): responsables, plantillas HTML, motivos HPS y, en este caso, un formulario cuyo nombre no sigue el patrón `Form…` sino `frm…`.
- **Nombre de formulario:** el destino se obtiene de un diccionario literal y se pasa a `DoCmd.OpenForm` (`:56`, `:61`); no se construye ninguna ruta ni se concatena entrada del usuario.
- **Sin control de rol:** este formulario no comprueba permisos. La comprobación vive en `Form_Form0BDOpciones`, que es quien lo abre (`Form_Form0BDOpciones.cls:137-138`).
- **Sin escritura:** no hay `INSERT`, `UPDATE` ni `DELETE`; los cuatro formularios de destino son a su vez gestores de su propia entidad.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Catálogo de entidades | diccionario en memoria | caché de `m_ObjEntorno` | `Entorno.cls:180-208` |
| Entidad → formulario | diccionario en memoria | caché de `m_ObjEntorno` | `Entorno.cls:210-238` |
| Ayuda | PDF homónimo en el directorio de documentación | E/S de fichero | `:84`; `FU.bas:475-500` |

Sin escrituras.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Ya se ve: el catálogo de entidades está codificado en el literal de `Entorno.cls` en lugar de en tabla, así que dar de alta un módulo nuevo exige recompilar (`:21-23`); la navegación es abrir/cerrar formularios con `DoCmd`, patrón que no tiene equivalente directo; y el título de la ventana se toma de una etiqueta en tiempo de ejecución (`:112`) en lugar de declararse en el formulario.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-7.** `RowSource` de diseño de `ListaEntidades` (`.form.txt:645`) contiene una lista de entidades que **no coincide** con `m_ObjEntorno.Entidades` (`:21-23`): los valores de diseño pertenecen a otro catálogo. `EstablecerLista` la vacía siempre (`:22`), así que solo sería visible si la lectura del catálogo fallara antes de llegar a `lst.RowSource = ""`; a verificar el comportamiento en ese caso.
- **H-HPS-S-8.** `ComandoAbrirFormulario_Click` deshabilita el botón al entrar (`:51`) y solo lo rehabilita `ListaEntidades_Click` (`:145-148`): en la ruta de error (`:53-60`, `:68-75`) el botón se queda deshabilitado hasta que el usuario vuelva a pulsar la lista. Además invoca `ListaEntidades_Click` justo antes de cerrarse (`:62-63`), una llamada sobre un formulario que va a desaparecer y que no aporta nada, ya que habilita y deshabilita el mismo botón dentro del mismo clic.

**Defectos de interfaz heredados (a verificar):** el título de diseño es `EXPEDIENTES Versión 2024-001 Administrador` (`.form.txt:39`) —otro formulario y otra versión— aunque `Form_Open` lo sobrescribe con el texto de la etiqueta (`:112`), de modo que el defecto solo es visible en diseño. `ComandoAyuda` tiene el rótulo `Salir` (`.form.txt:552`), idéntico al de `cmdSalir` (`.form.txt:778`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: catálogo de entidades en tabla de configuración con orden y permisos por entrada, en lugar de dos diccionarios literales; una única pantalla de administración con secciones en lugar de un menú que encadena ventanas; y el alta de módulos como dato, no como recompilación.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir el formulario se listan todas las entidades gestionables con su etiqueta.
2. Con ninguna fila seleccionada, el botón de abrir está deshabilitado.
3. Al seleccionar una fila, el botón se habilita.
4. Pulsar el botón abre el formulario de la entidad y cierra este.
5. El doble clic sobre una fila hace lo mismo que el botón.
6. Ante un error no previsto, el usuario recibe el aviso y el formulario se cierra.

## Pendientes

- Si `EstablecerLista` puede fallar antes de vaciar la lista (H-HPS-S-7).
- Si existe algún otro punto de entrada a `FormEntidades` además de `Form_Form0BDOpciones`.
- Qué comprobación de rol hace `Form_Form0BDOpciones` antes de abrir este formulario: está fuera de su fichero y por tanto fuera de esta ficha.
- Los nombres concretos de los cuatro formularios de destino (`:56`) se resuelven en tiempo de ejecución; solo constan en `Entorno.cls:210-238`.