# Form_FormNCAuditoriaControlEficaciaAlta

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoriaControlEficaciaAlta.cls` (239 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana auxiliar de alta del plan de control de eficacia de una NC de auditoría durante el alta o edición de los datos generales: recoge la fecha prevista y la descripción del control y delega la grabación en la ficha general (`Form_FormNCAuditoriaGeneral`). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) que opera sobre el objeto global `m_ObjNCAuditoriaActiva` (`Variables Globales.bas:127`). Las citas son `Form_FormNCAuditoriaControlEficaciaAlta.cls:línea` salvo indicación; `Form_FormNCAuditoriaGeneral.cls` abrevia `Form_FormNCAuditoriaGeneral.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormNCAuditoriaGeneral`, botón `ComandoControlEficaciaDatos` | Vuelca los datos generales al objeto, exige `DatosGeneralesOK`, cierra una instancia previa de este formulario y lo abre con `DoCmd.OpenForm` | `Form_FormNCAuditoriaGeneral.cls:50-64` |

Se cierra a sí mismo con `DoCmd.Close acForm, Me.Name, acSaveNo` tras grabar (`:95`) o desde «Salir» (`:36-39`). No devuelve valores: deja los datos en el objeto activo y dispara la grabación de la ficha general.

## Controles y eventos

`.form.txt` declara 23 bloques `Begin`: el formulario (título «ALTA DE ACCIÓN REALIZADA», `.form.txt:40`), 10 bloques de valores por defecto (`:132-194`), 3 secciones (`EncabezadoDelFormulario` `:204`, `Detalle` `:242`, `PieDelFormulario` `:473`) y los 9 controles de la tabla. El título del bloque de cabecera es «Alta control de eficacia» (`:222`) y `EstablecerDatos` lo copia al `Caption` (`:192`).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | Título «Alta control de eficacia» | `.form.txt:221-222` |
| `RequiereControlEficacia` | cuadro de texto bloqueado, `Tag=DATO` | — | Se fija a «Sí» al cargar | `:196`, `.form.txt:431-450` |
| `FechaPrevistaControlEficacia` | cuadro de texto, `Tag=DATO` | `AfterUpdate` | Fecha prevista; recalcula bordes | `:118-122`, `.form.txt:261-265` |
| `ControlEficacia` | cuadro de texto, `Tag=DATO` | `AfterUpdate` | Descripción del control | `:111-114`, `.form.txt:340-343` |
| `ComandoGrabar` | botón «Grabar» | `OnClick` | Valida y delega la grabación | `:41-109`, `.form.txt:393-395` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra sin guardar | `:36-39`, `.form.txt:488-490` |
| `Etiqueta254`, `255`, `256` | 3 etiquetas | — | Rótulos de los tres campos | `.form.txt:292-293, 313-314, 370-371` |

## Flujos

### Apertura (`Form_Load`, `:156-182`)

`EstablecerDatos` exige objeto activo («El formulario se ha abierto con parámetros insuficentes», `:188-191`), copia el título al `Caption`, habilita «Grabar», fija «Sí» en `RequiereControlEficacia` y pinta bordes (`:192-197`). **No precarga** fecha ni descripción desde el objeto: el alta empieza siempre en blanco.

### Grabar (`:41-109`)

1. Técnico ⇒ «No tiene autorización para esa acción» (`:52-55`).
2. Exige fecha prevista y descripción (`:56-63`) y las copia al objeto activo (`:64-67`).
3. `HaHabidoCambios` compara ambos campos con `Form_FormNCAuditoria.m_NCAlInicio`; sin cambios ⇒ «No hay cambios que guardar» (`:68-75`, `:207-235`).
4. Llama a `Form_FormNCAuditoriaGeneral.RellenarDatosObjeto` (`:76-79`) y exige `DatosGeneralesOK`; si no, «Antes de rellenar este apartado ha de tener todos los datos obligatorios principales rellenos» (`:80-83`).
5. Si `FormNCAuditoria` está abierto y su `FrmDetalle` muestra `FormNCAuditoriaGeneral`, ejecuta `Form_FormNCAuditoriaGeneral.ComandoGrabar_Click` (`:84-92`).
6. Cierra el formulario (`:95`).

### Errores

Error 1000 ⇒ aviso con `m_Error`; otro ⇒ mensaje con número y descripción más `CorreoAlAdministrador` (`:100-108`).

## Reglas, validaciones y cálculos

- **Rol:** los técnicos no graban (`:52-55`).
- **Obligatorios:** fecha prevista y descripción (`:56-63`); `DatosGeneralesOK` exige, entre otros, `Tipo`, `RequiereAccionCorrectiva` y `Descripcion` (`NCAuditoria.cls:756-775`; resto del cuerpo no revisado).
- **Indicador visual:** borde de campo relleno o no relleno por cada uno de los tres campos (`:124-154`; constantes en `Variables Globales.bas:83-86`).

## Lecturas y escrituras

El formulario no accede a datos por sí mismo. La persistencia ocurre en `Form_FormNCAuditoriaGeneral.cls:ComandoGrabar_Click` (`Form_FormNCAuditoriaGeneral.cls:78-198`) mediante `RegistrarDatosUnicos` y `RegistrarControlEficacia` (`classes/NCaUDITORIAOperaciones.cls:136-146, 262-267`); se describe en la ficha de la general. Sin transacciones en este formulario.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Sin SQL propio; las fechas viajan como texto (`:65`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Si `FormNCAuditoria` no está abierto, o su detalle no es la ficha general, el formulario se cierra sin grabar nada y sin aviso (`:84-95`).
- `Err.Raise 100` tras `RellenarDatosObjeto` (`:76-79`) no es el 1000 de los avisos controlados: el error iría al administrador por correo.
- El `Caption` del formulario es «ALTA DE ACCIÓN REALIZADA» (`.form.txt:40`), que no corresponde a un control de eficacia; `EstablecerDatos` lo sustituye en ejecución (`:192`).
- El método público `RellenarDatosObjeto` (`:9-34`) no lo llama este formulario; si la ficha general lo usa, su efecto está en `Form_FormNCAuditoriaGeneral.cls` (pendiente).
- `On Error Resume Next` en `EstableceColorBordes` y manejadores de `AfterUpdate` (`:112, 119, 125`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: integrar estos dos campos como sección de la ficha general, sin ventana auxiliar.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un técnico no puede grabar (`:52-55`).
2. No se graba sin fecha prevista ni descripción (`:56-63`).
3. No se graba si los datos generales obligatorios no están completos (`:80-83`).
4. Sin cambios respecto al estado inicial, informa de que no hay nada que guardar (`:72-75`).
5. Tras grabar desde la ficha general, la ventana se cierra (`:95`).

## Pendientes

- Cuerpo completo de `NCAuditoria.DatosGeneralesOK`.
- La persistencia del plan ocurre en `RegistrarDatosUnicos`, que llama a `RegistrarControlEficacia` si `RequiereControlEficacia = "Sí"` (`classes/NCaUDITORIAOperaciones.cls:262-267`); los dos campos llegan al objeto por asignación directa (`:64-67`), no por `RellenarDatosObjeto` de la general (`Form_FormNCAuditoriaGeneral.cls:515-540`, solo recorre sus propios controles).
