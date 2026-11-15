# Form_FormNCProyectoTipologiaGestion

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoTipologiaGestion.cls` (361 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Mantenimiento del catálogo de tipologías de NC de proyecto: muestra la lista de tipologías (tabla `TbTiposNCProyectos`) y permite dar de alta, editar y eliminar, delegando el alta y la edición en `FormTipologiaNCProyecto`. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`; la lista es de valores). Las citas son `Form_FormNCProyectoTipologiaGestion.cls:línea` salvo indicación; `constructor.bas` está en `modules/` y `TipologiaNCProyectos.cls` en `classes/`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_Form0BDOpcionesParteProyectos | Botón `ComandoConfiguracion`: rechaza a los técnicos («No tiene autorización para esa acción»), cierra la instancia abierta sin guardar y abre el formulario | `Form_Form0BDOpcionesParteProyectos.cls:87-102` |

Abre `FormTipologiaNCProyecto` (`:35, 76`) y queda a la escucha de su `WithEvents` `m_Form` (`:6, 37, 78`). Se cierra con «Salir» (`:15-18`, `acSaveNo`) o por error en `Form_Open` (`:183`). Declara tres eventos públicos (`AltaTipologia`, `EdicionTipologia`, `TipologiaEliminada`, `:10-12`); en el clon ningún otro módulo declara una variable `Form_FormNCProyectoTipologiaGestion`, así que nadie los consume.

## Controles y eventos

`.form.txt` declara 21 bloques `Begin`: el formulario (`.form.txt:4`), 11 de valores por defecto (`:134-205`), 3 secciones (`EncabezadoDelFormulario` `:210`, `Detalle` `:247`, `PieDelFormulario` `:442`) y los 6 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo1` | etiqueta | — | Texto «Tipologías de NC de proyecto» | `.form.txt:220-230` |
| `Lista` | lista (valores, 2 columnas: id oculto y tipología) | `OnClick`, `OnDblClick` | El clic carga la tipología y habilita «Editar» y «Eliminar»; el doble clic edita si «Editar» está habilitado | `:188-253, 284-291`, `.form.txt:304-325` |
| `ComandoAlta` («Alta») | botón | `OnClick` | Abre `FormTipologiaNCProyecto` vacío | `:20-52`, `.form.txt:344-355` |
| `ComandoEditar` («Editar», deshabilitado de diseño) | botón | `OnClick` | Abre `FormTipologiaNCProyecto` con la tipología seleccionada | `:54-93`, `.form.txt:391-404` |
| `ComandoEliminar` («Eliminar», deshabilitado de diseño) | botón | `OnClick` | Elimina la tipología seleccionada previa confirmación | `:95-149`, `.form.txt:255-268` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra el formulario | `:15-18`, `.form.txt:452-463` |

## Flujos

### Apertura (`Form_Open`, `:152-184`)

1. `Ajustar Me` (`:158`).
2. Habilita «Alta» salvo para técnicos (`EsTecnico`) (`:159-163`).
3. `RellenarLista` (`:165`).

### Lista (`RellenarLista`, `:293-333`)

Vacía el `RowSource` y carga `constructor.getTipologiasNCProyecto` (`:306-307`), ordenado por `Tipologia` (`constructor.bas:1408-1410`). Añade una fila por tipología: `IDTipo;Tipologia`, con `;` sustituido por `:` (`:320`). Si hay resultado, selecciona la fila de índice 1 y dispara `Lista_Click` (`:326-327`).

### Selección (`Lista_Click`, `:188-253`)

Deshabilita «Editar» y «Eliminar» (`:197-198`), fija «Alta» según el rol (`:199-203`). Sin id, libera la selección (`:206-211`). Carga con `getTipologiaNCProyecto` si cambia el id (`:213-225`), habilita «Editar» y habilita «Eliminar» solo si no es técnico (`:232-237`).

### Alta y edición

«Alta» exige no ser técnico (`:27-30`), pone `m_ObjTipologiaProyectoActiva` a `Nothing`, cierra la instancia abierta de `FormTipologiaNCProyecto` y la abre (`:31-37`). «Editar» reutiliza la tipología seleccionada o la carga; sin selección avisa «Seleccione un elemento de la lista» (`:60-69`), fija `m_ObjTipologiaProyectoActiva` y abre el formulario (`:71-78`). Al guardar, `FormTipologiaNCProyecto` lanza `AltaTipo` o `EdicionTipo` (`Form_FormTipologiaNCProyecto.cls:144-146, 195-197`); este formulario recarga la lista y reemite `AltaTipologia` o `EdicionTipologia` (`:256-268, 337-348`).

### Eliminar (`ComandoEliminar_Click`, `:95-149`)

Exige no ser técnico (`:102-105`), obtiene la tipología (`:106-114`) y pide confirmación «¿Desea realmente eliminar el tipo seleccionado?» con «No» por defecto (`:119`). `Eliminar` falla si el código ha formado parte de alguna NC («El código ha formado parte de n NC's», `TipologiaNCProyectos.cls:108-115, 212-218`); si no, ejecuta el borrado (`TipologiaNCProyectos.cls:223-226`). Después emite `TipologiaEliminada` y recarga la lista (`:131-133`).

### Errores

Error 1000 ⇒ advertencia con `m_Error`; otro ⇒ `CorreoAlAdministrador` y mensaje crítico (p. ej. `:43-50, 141-147`).

## Reglas, validaciones y cálculos

- Los técnicos no pueden dar de alta ni eliminar; pueden editar (`:27-30, 102-105, 159-163, 199-203, 233-237`).
- Una tipología usada por alguna NC de proyecto no se elimina (`TipologiaNCProyectos.cls:108-115`).
- Confirmación previa a la eliminación, con «No» por defecto (`:119`).
- `m_CodigoAscendente` y `m_TipologiaAscendente` están declaradas y no se usan (`:8-9`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listar tipologías | `TbTiposNCProyectos` | SQL dinámico sin parámetros (`SELECT * … ORDER BY Tipologia`) | `constructor.bas:1408-1412` |
| Cargar una tipología | `TbTiposNCProyectos` | SQL dinámico concatenado con `p_IDTipo` | `constructor.bas:1358-1360` |
| Comprobar uso antes de borrar | NC de proyecto por tipo | `constructor.getNCsPorTipoNCProyecto`, fuera de esta ficha | `TipologiaNCProyectos.cls:108` |
| Eliminar | `TbTiposNCProyectos` | SQL dinámico `DELETE` concatenado con `IDTipo`, ejecutado con `getdb().Execute`; sin transacción | `TipologiaNCProyectos.cls:223-226` |

## Coexistencia

`DELETE * FROM` es sintaxis Jet/ACE: PostgreSQL exige `DELETE FROM`. `SELECT *` con `ORDER BY Tipologia` depende de la intercalación del motor (mayúsculas y acentos). La concatenación de `IDTipo` debe revisarse (tipo del id, comillas); si `IDTipo` pasa a identidad de PostgreSQL, la carga posterior a un alta depende de cómo se recupere el id (en `FormTipologiaNCProyecto`, fuera de esta ficha). El borrado no usa transacción, de modo que la comprobación de uso y el `DELETE` no son atómicos (`TipologiaNCProyectos.cls:212, 226`).

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual**

- Catálogo único sin filtros ni paginación, ordenado por nombre (`constructor.bas:1410`).
- La edición está permitida a técnicos; el alta y la baja no (`:27-30, 102-105, 233-237`).

**Posibles defectos (a verificar)**

- `ComandoAlta_Click` no declara `On Error GoTo errores` (`:20-34`): el `Err.Raise 1000` para técnicos (`:29`) no llega al manejador de `:43` y produciría un error en tiempo de ejecución sin tratar, aunque el botón ya se deshabilita para técnicos (`:160, 200`).
- `RellenarLista` y `Lista_Click` seleccionan `lst.Selected(1)` (`:326`) sin cabecera en la lista; con una sola tipología el índice 1 no existe, y la primera nunca se selecciona por defecto.
- El manejador de `Form_Open` cierra con `DoCmd.Close acForm, Me` (`:183`), pasando el objeto y no el nombre.
- Los tres eventos públicos no tienen consumidor en el clon (`:10-12`).
- La comprobación de uso y el borrado no son atómicos (`TipologiaNCProyectos.cls:212-226`).

**Propuestas web pendientes de decisión**

- Sustituir la lista y los tres botones por una tabla editable con confirmación de borrado y mensaje claro de dependencias.
- Permitir desactivar una tipología en lugar de borrarla, para conservar el histórico.
- Aplicar los permisos de alta y baja en servidor, no solo deshabilitando botones.

## Criterios de aceptación

1. Desde el menú de proyectos, un usuario no técnico abre el formulario y ve las tipologías ordenadas por nombre (`Form_Form0BDOpcionesParteProyectos.cls:87-102`; `constructor.bas:1410`).
2. Un técnico ve «Alta» y «Eliminar» deshabilitados y puede editar (`:159-163, 233-237`).
3. «Editar» (o doble clic) abre `FormTipologiaNCProyecto` con la tipología elegida; sin selección avisa (`:60-78, 284-291`).
4. Tras un alta o edición, la lista se recarga (`:256-268, 337-348`).
5. «Eliminar» pide confirmación y no borra una tipología usada por alguna NC, mostrando el motivo (`:119`; `TipologiaNCProyectos.cls:108-115`).
6. Cada error muestra el mensaje y avisa al administrador salvo error 1000 (`:43-50`).

## Pendientes

- Formulario `FormTipologiaNCProyecto` (alta y edición): ficha propia, fuera de este lote.
- `constructor.getNCsPorTipoNCProyecto`: origen de datos del recuento de uso, no resuelto aquí (`TipologiaNCProyectos.cls:108`).
- Dónde se rellena `m_ObjTipologiaProyectoActiva` tras un alta (`:268, 348`): referencia global no resuelta en esta ficha.
