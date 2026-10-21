# Form_Form0BDOpcionesTecnicos

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_Form0BDOpcionesTecnicos.cls` (167 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Menú de inicio reducido para el perfil técnico: dos accesos a búsqueda de expedientes (simplificada y completa), ayuda y salida de la aplicación. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) y no lee ni escribe en la base de datos por sí mismo. Las citas son `Form_Form0BDOpcionesTecnicos.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_frmSplash` | Tras `EVE`, abre este menú si `EsTecnico = Sí`; si no, `Form0BDOpciones` | `Form_frmSplash.cls:38, 44-50` |
| `Form_Form0BDOpciones` | `ComandoBuscadorTecnico_Click` cierra el menú técnico si estaba abierto y lo vuelve a abrir | `Form_Form0BDOpciones.cls:83-94` |

`EsTecnico` vale «Sí» cuando el usuario conectado no es administrador ni de calidad (`Variables Globales.bas:382-392`); ser administrador lo anula (`Variables Globales.bas:395-397`). El botón de `Form0BDOpciones` que abre este menú solo es visible para administradores (`Form_Form0BDOpciones.cls:253-259`; rótulo «Abrir como técnico», `Form_Form0BDOpciones.form.txt:3126-3129`).

Abre `FormExpedientesGestionTecnica` (`:95`) y `FormExpedientesGestion` en solo lectura (`:69`); abre la ayuda en PDF (`AbrirAyuda`, `:41`) y puede cerrar la base de datos (`:21`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComandoBusquedaSimple` (rótulo «Comando36», imagen) | botón | `OnClick` → `ComandoBusquedaSimple_Click` | Cierra `FormExpedientesGestionTecnica` si está abierto y lo abre sin `OpenArgs` | `:84-108`, `.form.txt:918-923` |
| `lblBusquedaSimple` («Búsqueda Simplificada») | etiqueta | `OnClick` | Delegado en `ComandoBusquedaSimple_Click` si el botón está habilitado | `:158-162`, `.form.txt:888-892` |
| `ComandoBusquedaCompleta` (rótulo «Comando36», imagen) | botón | `OnClick` → `ComandoBusquedaCompleta_Click` | Cierra `FormExpedientesGestion` si está abierto y lo abre con `OpenArgs:="SoloLectura"` | `:59-82`, `.form.txt:1509-1514` |
| `lblBusquedaCompleta` («Búsqueda Completa») | etiqueta | `OnClick` | Delegado en `ComandoBusquedaCompleta_Click` si el botón está habilitado | `:163-167`, `.form.txt:1478-1482` |
| `ComandoAyuda` (rótulo «Salir», información sobre herramientas «Ayuda») | botón | `OnClick` → `ComandoAyuda_Click` | `AbrirAyuda`; si devuelve error lo muestra | `:35-57`, `.form.txt:799-802` |
| `cmdSalir` («Salir de la aplicación») | botón | `OnClick` → `CmdSalir_Click` | Cierra este menú o la base de datos (ver Flujos) | `:8-32`, `.form.txt:2927-2930` |
| `lblVersion` | etiqueta | — | «Versión: » + `m_ObjEntorno.VersionAplicacion` | `:137`, `.form.txt:1448-1449` |
| `lblTitulo1` | etiqueta | — | Rótulo fijo de cabecera | `.form.txt:548-549` |
| `lblEstado` | etiqueta oculta | — | Sin uso en el `.cls` | `.form.txt:3008-3022` |

Ningún código del formulario cambia `Enabled` de los botones de búsqueda; en el `.form.txt` no se declara `Enabled` para ellos, así que la comprobación de las etiquetas (`:159, 164`) siempre se cumple con el diseño actual.

## Flujos

### Apertura (`Form_Open`, `:110-156`)

1. Reloj de arena activado y `Ajustar Me`, que ajusta el tamaño de la ventana a las secciones (`:116-119`; `FUNCIONES UTILES.bas:2703-2720`).
2. Si `m_ObjEntorno` no existe, ejecuta `EVE` para cargar el entorno; un error en `EVE` se eleva como error interno 1000 (`:121-126`).
3. Título de la ventana = `m_TituloFormulario` (`:135`), que compone versión, nombre del usuario y, según el caso, «(Sólo lectura)» y «EN PRUEBAS» (`Variables Globales.bas:448-462`). Sustituye el título de diseño (`.form.txt:41`).
4. Rellena `lblVersion` (`:137`) y quita el reloj de arena (`:141`).

### Búsquedas (`:59-108`)

1. Si el formulario destino ya está abierto (`FormularioAbierto`, `FUNCIONES UTILES.bas:2821-2834`), lo cierra sin guardar (`:66-68`, `:92-94`).
2. Lo abre: la búsqueda completa con `OpenArgs:="SoloLectura"` (`:69`), que `FormExpedientesGestion` traduce a `m_AbiertoParaEditar = False` (`Form_FormExpedientesGestion.cls:2910-2914`); la simplificada sin argumentos (`:95`).

### Salir (`CmdSalir_Click`, `:8-32`)

1. Si `Form0BDOpciones` está abierto (caso del administrador que entró con «Abrir como técnico»), solo cierra este menú (`:13-16`).
2. En otro caso pide confirmación «¿Desea realmente salir de la aplicación» (`:17`); con «Sí» ejecuta `Application.CloseCurrentDatabase` (`:18-21`).

### Errores

`Form_Open`, `ComandoAyuda_Click` y los dos botones de búsqueda envían `CorreoAlAdministrador` salvo para el error interno 1000 y muestran `MsgBox` crítico (`:49-56, 74-81, 100-107, 146-154`). `CmdSalir_Click` muestra el mensaje pero no envía correo (`:24-30`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| El menú técnico se muestra en el arranque solo con `EsTecnico = Sí` | `Form_frmSplash` | `Form_frmSplash.cls:44-50` |
| Técnico = ni administrador ni calidad | `EVE` | `Variables Globales.bas:382-397` |
| La búsqueda completa se abre siempre en solo lectura | `ComandoBusquedaCompleta_Click` | `:69` |
| Salir cierra la base de datos solo si no hay menú general abierto y el usuario confirma | `CmdSalir_Click` | `:13-21` |
| La ayuda abre un PDF único del directorio de ayuda del entorno; si no existe, «No se ha generado la ayuda para este formulario aún» | `AbrirAyuda` | `FUNCIONES UTILES.bas:3094-3101` |

## Lecturas y escrituras

Sin lecturas ni escrituras propias. La única carga indirecta es `EVE` cuando `m_ObjEntorno` no existe (`:121-122`; `Variables Globales.bas:282`), cuyo contenido se documenta en la ficha de `Form_frmSplash`. `CorreoAlAdministrador` (`FUNCIONES UTILES.bas:2943`) no se analiza aquí.

## Coexistencia

Sin SQL ni acceso a datos propio: no requiere adaptación para PostgreSQL. Lo que debe preservarse es el reparto por perfil (técnico frente a administrador o calidad) y que la búsqueda completa sea de solo lectura para este perfil. Detalle general en `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Un error en `Form_Open` no asigna `Cancel` (`:146-155`): el menú quedaría abierto sin entorno cargado y sin versión.
- `ComandoAyuda` rotula «Salir» en el `.form.txt` (`.form.txt:800`) y los dos botones de búsqueda «Comando36» (`.form.txt:919, 1510`); a verificar si la imagen tapa el texto.
- `OpenArgs:="SoloLectura"` (`:69`) se compara con «Sololectura» (`Form_FormExpedientesGestion.cls:2910`); funciona solo porque `Option Compare Database` no distingue mayúsculas (`Form_FormExpedientesGestion.cls:2`).
- `AbrirAyuda` calcula la ruta del PDF del formulario activo y la sobrescribe en la línea siguiente con un PDF fijo (`FUNCIONES UTILES.bas:3094-3095`): la ayuda no es por formulario.
- `m_Inicio` se asigna y no se usa (`:112, 126, 128`).
- `CmdSalir_Click` no notifica errores al administrador, a diferencia del resto de handlers (`:24-30`).
- `lblEstado` («Obteniendo los Exp Filtrados») está en el diseño sin código que lo use (`.form.txt:3021-3022`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: un único menú de inicio con opciones filtradas por rol en lugar de dos formularios; ayuda contextual por pantalla.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario que no es administrador ni de calidad ve, al entrar, solo búsqueda simplificada, búsqueda completa, ayuda y salir.
2. La búsqueda completa abierta desde este menú no permite editar.
3. Un administrador puede previsualizar el menú técnico; al salir de él vuelve al menú general sin cerrar la aplicación.
4. Salir desde el menú técnico pide confirmación antes de cerrar la sesión.
5. La pantalla muestra la versión de la aplicación.

## Pendientes

- Qué restringe `FormExpedientesGestionTecnica` frente a `FormExpedientesGestion` (su ficha propia).
- Valor de `Enabled` en tiempo de ejecución de los botones de búsqueda, por si otro módulo lo cambia.
- Contenido de `CorreoAlAdministrador` y de la carga completa de `EVE`.
- Si la imagen de los botones oculta los rótulos «Comando36» y «Salir» en pantalla.
