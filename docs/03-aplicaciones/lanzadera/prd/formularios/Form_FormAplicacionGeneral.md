# Form_FormAplicacionGeneral

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormAplicacionGeneral.cls` (129 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Contenedor de la gestión de aplicaciones: barra de menú con dos pestañas («Búsqueda» y «Datos Generales»), rótulo con la aplicación activa y subformulario `FormDetalle` que aloja la pestaña. No tiene origen de datos ni escribe en la base de datos. Las citas son `Form_FormAplicacionGeneral.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormMenuPrincipalAdmin` | `SUBMENU2_Click` (rótulo «Aplicaciones») anula `m_ObjUsuarioActivo`, carga `FormAplicacionGeneral` en su `FormDetalle`, marca `MENU1` como activo y llama a `Form_FormAplicacionGeneral.MENU1_Click` | `Form_FormMenuPrincipalAdmin.cls:241-260`, `Form_FormMenuPrincipalAdmin.form.txt:1080-1083` |

Aloja `FormAplicacionesGestion` (`:62-68`) y `FormAplicacionDatosGenerales` (`:73-77`) en `FormDetalle`. Escucha `AplicacionSeleccionada` y `QuitarAplicacionSeleccionada` de la pestaña de búsqueda mediante `WithEvents` (`:3, 54-60`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `MENU1` + `lblMenu1` («Búsqueda») | rectángulo + etiqueta | `OnClick` → `MENU1_Click` (público) | Carga `FormAplicacionesGestion`, anula `m_ObjAplicacionActiva` y se suscribe a sus eventos | `:62-72`, `.form.txt:262-305` |
| `MENU2` + `lblMenu2` («Datos Generales») | rectángulo + etiqueta | `OnClick` → `MENU2_Click` | Carga `FormAplicacionDatosGenerales` | `:73-81`, `.form.txt:327-374` |
| `MENU1`, `MENU2`, `lblMenu1`, `lblMenu2` | — | `OnMouseMove` | Resalta la opción bajo el cursor (color e indicador) | `:84-128` |
| `EncabezadoDelFormulario` | sección | `OnMouseMove` | Apaga ambos indicadores y colores | `:7-22`, `.form.txt:227-228` |
| `IND1`, `IND2` | indicadores | — | Marca visual de la pestaña resaltada | `:30-32` |
| `MENU3` | rectángulo | sin handler | Sin comportamiento identificado | `.form.txt:245` |
| `lblAplicacionActiva` | etiqueta | — | «Aplicación Activa: » más el nombre de la aplicación seleccionada | `:41, 55, 59` |
| `FormDetalle` | subformulario | — | Aloja la pestaña; vacío al abrir | `:39`, `.form.txt:497` |

## Flujos

### Apertura (`Form_Load`, `:24-52`)

1. Enciende `IND1`, apaga `IND2` y colorea `MENU1` como activo y `MENU2` como inactivo con los colores de `m_ObjEntorno` (`:30-35`).
2. Vacía `FormDetalle` (`:39`), anula `m_ObjUsuarioActivo` (`:40`) y deja el rótulo en «Aplicación Activa: » (`:41`).
3. El menú de administración llama después a `MENU1_Click`, que carga la búsqueda (`Form_FormMenuPrincipalAdmin.cls:256`).

### Navegación

- `MENU1_Click` solo cambia `SourceObject` si es distinto; en cualquier caso anula `m_ObjAplicacionActiva` y reasigna `m_Form` (`:63-67`).
- `MENU2_Click` solo cambia `SourceObject`; la pestaña de datos lee la aplicación activa al abrirse (`:74-76`; `Form_FormAplicacionDatosGenerales.cls:90-94`).
- Las etiquetas delegan en su rectángulo (`:70-72, 79-81, 101-105, 125-128`).

### Eventos de la búsqueda

- `AplicacionSeleccionada` añade el nombre al rótulo (`:58-60`).
- `QuitarAplicacionSeleccionada` deja el rótulo sin nombre (`:54-56`).

### Errores

Solo `Form_Load` tiene manejador: envía `CorreoAlAdministrador` salvo el error interno 1000 y muestra `MsgBox` crítico; el formulario queda abierto (`:43-51`). El resto de handlers no captura errores; dependen de que `m_ObjEntorno` exista (`:15-20, 93-98`).

## Reglas, validaciones y cálculos

- **Permisos:** sin comprobación propia. Solo se llega desde `FormMenuPrincipalAdmin`, que se abre cuando `EsAdministradorCalculado = Sí` (`Variables Globales.bas:724-725`).
- **Colores del menú:** `ColorMenuActivado` es un RGB fijo de `Entorno` (`Entorno.cls:722-728`).
- **Aplicación activa:** volver a «Búsqueda» siempre la anula (`:66`).

## Lecturas y escrituras

Sin lecturas ni escrituras propias en base de datos. Modifica las globales `m_ObjUsuarioActivo` y `m_ObjAplicacionActiva` (`:40, 66`; `Variables Globales.bas:100-101`).

## Coexistencia

Sin impacto directo en datos. Pendiente de `../coexistencia.md` solo como pieza de navegación.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `MENU1_Click` anula `m_ObjAplicacionActiva` (`:66`) pero no limpia `lblAplicacionActiva`: si la pestaña ya estaba cargada (`:63`), el rótulo y el campo `AplicacionSeleccionada` de la búsqueda siguen mostrando una aplicación que ya no está activa.
- `Form_Load` anula `m_ObjUsuarioActivo` (`:40`), global de la gestión de usuarios; repite la línea de `Form_FormUsuarioGeneral.cls:44` y el menú ya lo hace (`Form_FormMenuPrincipalAdmin.cls:243`).
- El menú llama a `MENU1_Click` sobre `Form_FormAplicacionGeneral`, la instancia por defecto de la clase (`Form_FormMenuPrincipalAdmin.cls:256`), no sobre `FormDetalle.Form`; con el formulario alojado como subformulario, no está comprobado que sea la misma instancia.
- `MENU3` existe sin handler (`.form.txt:245`).
- Los handlers `MouseMove` no tienen manejador de error (`:84-128`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: pestañas como rutas hijas (`/aplicaciones/buscar`, `/aplicaciones/:id`); mostrar la aplicación activa desde el estado de la ruta.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir «Aplicaciones» desde el menú de administración se muestra la búsqueda y el rótulo «Aplicación Activa: » sin nombre.
2. Seleccionar una aplicación en la búsqueda muestra su nombre en el rótulo.
3. Quitar la aplicación activa deja el rótulo sin nombre.
4. «Datos Generales» muestra la aplicación seleccionada.
5. Volver a «Búsqueda» deja sin aplicación activa.

## Pendientes

- Si `Form_FormAplicacionGeneral.MENU1_Click` actúa sobre la instancia alojada (`Form_FormMenuPrincipalAdmin.cls:256`).
- Uso previsto de `MENU3` (`.form.txt:245`).
