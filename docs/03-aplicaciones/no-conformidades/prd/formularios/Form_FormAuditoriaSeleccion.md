# Form_FormAuditoriaSeleccion

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormAuditoriaSeleccion.cls` (116 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Selector de auditoría previo a la gestión de NC de auditoría: ofrece una lista desplegable con todas las auditorías, fija la elegida como auditoría activa y abre `FormNCAuditoria`. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); la lista se rellena como «Lista de valores» desde el entorno en memoria. Las citas son `Form_FormAuditoriaSeleccion.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas` y `VG.bas`, `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_Form0BDOpcionesAuditorias | Cierra un `FormAuditoriaSeleccion` previo y lo abre sin argumentos | `Form_Form0BDOpcionesAuditorias.cls:163-167` |

Abre `FormNCAuditoria` (`:50`) tras cerrar uno ya abierto (`:46-48`) y se cierra a sí mismo (`:56`). Efecto global: asigna `m_ObjAuditoriaActiva` (`:35`; `VG.bas:126`) y pone `m_ObjNCAuditoriaActiva` a `Nothing` (`:49`).

## Controles y eventos

`.form.txt` declara 18 bloques `Begin`: el formulario (`.form.txt:4`), 9 bloques de valores por defecto (`.form.txt:127-186`), 3 secciones (`EncabezadoDelFormulario` `.form.txt:191`, `Detalle` `.form.txt:220`, `PieDelFormulario` `.form.txt:291`) y los 5 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` («AUDITORÍAS») | etiqueta | — | Título; se copia a la barra de título | `.form.txt:209-210`, `:90` |
| `Auditoria` (etiqueta `Etiqueta86`, «Auditoría») | cuadro combinado, lista de valores, 2 columnas (`0;1134`), `LimitToList` | — | Lista de auditorías: columna 0 = identificador (oculta), columna 1 = nombre | `.form.txt:229-262, 278-279` |
| `ComandoSeleccionar` («Seleccionar») | botón | `OnClick` | Valida, fija la auditoría activa y abre `FormNCAuditoria` | `:31-77`, `.form.txt:310-312` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | `:8-11`, `.form.txt:357-359` |
| Formulario | — | `OnOpen` | Ajusta tamaño, título y rellena el combo | `:83-114`, `.form.txt:37` |

## Flujos

### Apertura (`Form_Open`, `:83-114`)

1. Reloj de arena, `Ajustar Me` (`:92`; `FG.bas:615`) y `Caption` = título (`:94`).
2. `EstablecerCombos` llama a `EstablecerComboAuditoriasNombres` (`:19`), que vacía el origen de filas y añade `identificador;nombre` por cada auditoría del diccionario `m_ObjEntorno.ColAuditorias` (`FG.bas:3149-3162`). El diccionario se carga una vez con `constructor.getAuditorias`, `SELECT * FROM TbAuditorias ORDER BY IDAuditoria DESC` (`Entorno.cls:2201-2212`; `constructor.bas:4861-4862`), de modo que las auditorías más recientes salen primero.
3. `AvanceCerrar` y fin de espera (`:100-103`).

### Selección (`ComandoSeleccionar_Click`, `:31-77`)

1. Sin nombre en la columna 1 del combo ⇒ «Se ha de seleccionar una auditoría» (`:39-42`).
2. Carga la auditoría con `constructor.getAuditoria(p_ID:=Column(0))` (`:43`).
3. Si no devuelve objeto ⇒ mismo aviso (`:47-50`); si sí, cierra `FormNCAuditoria` abierto, borra la NC activa y abre `FormNCAuditoria` (`:52-58`).
4. Cierra este formulario (`:63`).

### Errores

Error 1000 ⇒ `MsgBox` de advertencia con `m_Error`; otro error ⇒ mensaje con número y detalle, `CorreoAlAdministrador` y `MsgBox` crítico (`:66-77, 106-114`). En `Form_Open` cualquier error cierra el formulario (`:113`).

## Reglas, validaciones y cálculos

- Debe elegirse una auditoría de la lista (`:39-42`); el combo limita a los valores de la lista (`.form.txt:229`).
- Sin comprobación de rol: el acceso lo gobierna el menú de auditorías (`Form_Form0BDOpcionesAuditorias`).
- Sin cálculos: el orden de la lista lo fija la consulta del entorno.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Lista de auditorías (una vez por sesión de entorno) | `TbAuditorias` | SQL dinámico concatenado, DAO `OpenRecordset` | `constructor.bas:4861-4863` |
| Auditoría elegida | `TbAuditorias` por `IDAuditoria` | SQL dinámico concatenado (el identificador se concatena sin comillas), DAO `OpenRecordset` | `constructor.bas:4815-4819` |
| Registro de errores no controlados | `TbCorreosEnviados` vía `CorreoAlAdministrador` | DAO `AddNew/Update` | `FG.bas:1291-1305` |

Sin escrituras propias ni transacciones.

## Coexistencia

- Ambas consultas son de selección sobre una sola tabla, sin funciones Jet; portables a PostgreSQL. El identificador se concatena como número (`constructor.bas:4817`): usar parámetros enlazados y revisar el tipo si el identificador deja de ser numérico.
- Los datos viajan por una cadena `identificador;nombre` en una lista de valores (`FG.bas:3160`); un nombre que contenga `;` rompería la columna (a verificar).
- El diccionario `ColAuditorias` se cachea en el entorno (`Entorno.cls:2206-2209`); una auditoría creada después no aparece hasta recargar el entorno.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La lista cacheada puede estar desfasada respecto a la tabla si otro usuario da de alta una auditoría (`Entorno.cls:2206-2209`).
- Si `FormNCAuditoria` falla al abrir, este formulario ya ha cambiado `m_ObjAuditoriaActiva` (`:35, 50`).
- El título de ventana del `.form.txt` es «Registro de NPedido» (`.form.txt:36`), residuo de otro formulario; lo sobrescribe `Form_Open` (`:94`).
- La clase de error 1000 en `Form_Open` cierra el formulario tras un aviso de «Advertencia» (`:111-113`), sin posibilidad de reintento.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: sustituir esta pantalla por un selector integrado en la cabecera de NC de auditoría.

## Criterios de aceptación

Propuestas, no aprobadas:

1. El selector muestra todas las auditorías, de mayor a menor identificador (`constructor.bas:4862`).
2. «Seleccionar» sin elegir muestra «Se ha de seleccionar una auditoría» y no cambia de pantalla (`:39-42`).
3. Con una auditoría elegida, queda como auditoría activa y se abre `FormNCAuditoria` sin NC activa (`:43, 49-50`).
4. «Salir» cierra el selector sin cambiar la auditoría activa (`:10`).

## Pendientes

- Contenido y reglas de `FormNCAuditoria`: ver su ficha.
- Quién puede abrir el menú de auditorías (Calidad o administración): ver ficha de `Form_Form0BDOpcionesAuditorias`.
