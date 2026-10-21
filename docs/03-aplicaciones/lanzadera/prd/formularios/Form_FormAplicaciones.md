# Form_FormAplicaciones

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormAplicaciones.cls` (156 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Contenedor del catálogo de aplicaciones: tres pestañas dibujadas con rectángulos (`MENU1`-`MENU3`), sus rótulos (`lblMenu1`-`lblMenu3`) e indicadores (`IND1`-`IND3`), y un subformulario `FormDetalle` que aloja uno de los tres listados de botones. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) y no lee ni escribe datos. Las citas son `Form_FormAplicaciones.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormMenuPrincipalAdmin` | `FormDetalle.SourceObject = "FormAplicaciones"` en su `MENU1_Click` | `Form_FormMenuPrincipalAdmin.cls:129-130` |
| `Form_FormMenuPrincipalCalidad` | igual | `Form_FormMenuPrincipalCalidad.cls:167-168` |
| `Form_FormMenuPrincipalUsuario` | igual | `Form_FormMenuPrincipalUsuario.cls:154-155` |

El menú principal que se abre depende del rol del usuario tras el login (`Variables Globales.bas:724-736`). Al cargarse, los tres menús llaman a su `MENU1_Click` y después fuerzan en el subformulario el listado [FormAplicacionesFueraOficina](Form_FormAplicacionesFueraOficina.md) con la pestaña 1 resaltada (`Form_FormMenuPrincipalAdmin.cls:52-70`, `Form_FormMenuPrincipalCalidad.cls:44-64`, `Form_FormMenuPrincipalUsuario.cls:43-59`).

Abre, dentro de `FormDetalle`: [FormAplicacionesFueraOficina](Form_FormAplicacionesFueraOficina.md), [FormAplicacionesOficina](Form_FormAplicacionesOficina.md) y [FormAplicacionesPrueba](Form_FormAplicacionesPrueba.md).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `MENU1` / `lblMenu1` (rótulo de «fuera de la oficina») | rectángulo / etiqueta | `OnClick` → `MENU1_Click` (pública) | Carga [FormAplicacionesFueraOficina](Form_FormAplicacionesFueraOficina.md) en `FormDetalle` si no estaba | `:39-50`, `.form.txt:247-248, 288-291` |
| `MENU2` / `lblMenu2` («Solo Oficina») | rectángulo / etiqueta | `OnClick` → `MENU2_Click` (pública) | Carga `FormAplicacionesOficina` | `:51-62`, `.form.txt:314-315, 357-360` |
| `MENU3` / `lblMenu3` («En Pruebas») | rectángulo / etiqueta | `OnClick` → `MENU3_Click` (pública) | Carga `FormAplicacionesPrueba` | `:63-74`, `.form.txt:481-482, 523-526` |
| `MENU1`-`MENU3`, `lblMenu1`-`lblMenu3`, `IND1` | rectángulos / etiquetas | `OnMouseMove` | Resaltan la pestaña bajo el ratón: muestran su `IND` y le dan `ColorMenuActivado`; apagan las otras dos | `:75-156` |
| `EncabezadoDelFormulario` | sección | `OnMouseMove` | Oculta los tres indicadores y devuelve los tres colores a `ColorMenuDesactivado` | `:6-25`, `.form.txt:229-230` |
| `FormDetalle` | subformulario | — | Aloja el listado activo; en diseño apunta a `FormAplicacionesPrueba` | `.form.txt:603-604` |
| `Imagen1`-`Imagen3` | imágenes | sin handler | Iconos decorativos de las pestañas | `.form.txt:378, 426, 544` |

## Flujos

### Apertura (`Form_Load`, `:27-38`)

1. Oculta `IND1`-`IND3` y pinta las tres pestañas con `m_ObjEntorno.ColorMenuDesactivado` (`:28-33`).
2. Vacía `FormDetalle.SourceObject` (`:35`): el listado de diseño no llega a mostrarse; lo fija después el menú principal.

### Cambio de pestaña

1. `MENU<n>_Click` compara `SourceObject` con el listado de la pestaña y solo lo reasigna si cambia (`:42-44, 54-56, 66-68`); cambiar `SourceObject` recarga el listado y dispara su `Form_Load`.
2. Los menús principales Admin y Calidad, si el usuario tiene `UltimaAplicacionAbierta`, invocan `Form_FormAplicaciones.MENU1_Click` o `MENU2_Click` según `EtiquetaAplicacionesAPulsar` («1» si la última aplicación no es de oficina, «2» si lo es) (`Form_FormMenuPrincipalAdmin.cls:132-164`; `classes/Usuario.cls:650-670`). En el menú de usuario esa rama está comentada (`Form_FormMenuPrincipalUsuario.cls:157-190`).

### Errores

Los tres `MENU<n>_Click` usan `On Error Resume Next` (`:40, 52, 64`): un fallo al cargar el listado no se notifica desde este formulario.

## Reglas, validaciones y cálculos

- Colores de pestaña: `ColorMenuActivado` / `ColorMenuDesactivado` del `Entorno` (`classes/Entorno.cls:722`).
- `AllowEdits` del contenedor: `False` en diseño (`.form.txt:11`); `MENU1_Click` lo deja en `True` tras cambiar (`:41-45`) y `MENU2_Click`/`MENU3_Click` lo dejan en `False` (`:53-57`, `:65-69`).
- No hay control de permisos en el contenedor: las tres pestañas se muestran a todos los roles; los permisos se aplican botón a botón en cada listado.

## Lecturas y escrituras

Ninguna propia.

## Coexistencia

Sin acceso a datos: no requiere adaptación para PostgreSQL. Ver `../coexistencia.md` cuando exista.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Al cargar, los menús principales llaman a `MENU1_Click` (que puede elegir la pestaña de oficina por la última aplicación) y justo después fuerzan [FormAplicacionesFueraOficina](Form_FormAplicacionesFueraOficina.md) y la pestaña 1 (`Form_FormMenuPrincipalAdmin.cls:52-70`): la preferencia por última aplicación quedaría anulada en la apertura.
- `AllowEdits` se trata de forma opuesta en `MENU1_Click` (`:41, 45`) y en `MENU2_Click`/`MENU3_Click` (`:53, 57`, `:65, 69`).
- `On Error Resume Next` en los tres cambios de pestaña oculta cualquier error (`:40, 52, 64`).
- `IND1` declara `OnMouseMove` y tiene handler (`.form.txt:268`, `:101-104`); `IND2` e `IND3` no, así que el resaltado se pierde al pasar sobre ellos.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: pestañas como rutas (`/aplicaciones/fuera-oficina`, `/aplicaciones/oficina`, `/aplicaciones/pruebas`) con la última pestaña recordada de forma explícita; resaltado por CSS en lugar de eventos de ratón.

## Criterios de aceptación

Propuestos, no aprobados:

1. La vista de aplicaciones muestra tres pestañas y abre por defecto la de fuera de la oficina.
2. Pulsar una pestaña muestra su listado sin recargar si ya estaba activo.
3. La pestaña activa se distingue visualmente de las otras dos.

## Pendientes

- Si `Form_FormAplicaciones.MENU1_Click`, llamado desde el menú principal (`Form_FormMenuPrincipalAdmin.cls:149, 163`), actúa sobre la instancia alojada en `FormDetalle` o sobre una instancia por defecto distinta del módulo de clase.
- Si la pestaña «En Pruebas» debe ser visible para usuarios no administradores (no hay código que la oculte).
- Efecto real de `AllowEdits` en un formulario sin `RecordSource` y en sus subformularios.
