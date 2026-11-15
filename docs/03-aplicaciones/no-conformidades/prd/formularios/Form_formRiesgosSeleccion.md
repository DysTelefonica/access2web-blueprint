# Form_formRiesgosSeleccion

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_formRiesgosSeleccion.cls` (205 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Diálogo modal que permite elegir, de una lista multiselección, los riesgos del expediente que se asocian a una NC de proyecto. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); la lista se rellena por código con los riesgos candidatos que devuelve `RiesgoServicio` y la selección se guarda en base de datos (NC existente) o solo en memoria (NC en alta). Las citas son `Form_formRiesgosSeleccion.cls:línea` salvo indicación; `RiesgoRepositorio.bas` abrevia `modules/RiesgoRepositorio.bas`, `RiesgoServicio.cls`, `classes/RiesgoServicio.cls` y `NCProyecto.cls`, `classes/NCProyecto.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_FormNCProyectoGeneral, botón `btnSeleccionarRiesgos` | `DoCmd.OpenForm "formRiesgosSeleccion"` en modo diálogo (`acDialog`) con `OpenArgs = "<IDNoConformidad>\|<IDExpediente>"` (0 si la NC es nueva); exige NC activa y expediente seleccionado | `Form_FormNCProyectoGeneral.cls:1125-1157` |

Al volver (síncrono por `acDialog`), el llamante refresca el texto de riesgos con `ActualizarTextoRiesgosUI` (`Form_FormNCProyectoGeneral.cls:1157, 1201`). El formulario se cierra con `DoCmd.Close` sin guardar nada más (`:12, 191`).

## Controles y eventos

`.form.txt` declara 21 bloques `Begin`: el formulario (`.form.txt:4`), 11 de valores por defecto (`.form.txt:411-483`), 3 secciones (`EncabezadoDelFormulario` `.form.txt:495`, `Detalle` `.form.txt:534`, `PieDelFormulario` `.form.txt:620`) y los 6 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | Su `Caption` (diseño: «GESTIÓN DE LAS AUDITORÍAS») se sustituye al cargar por «RIESGOS ASOCIADOS», también en la barra de título | `:39-40`, `.form.txt:505-516` |
| `lstRiesgos` | lista de valores, multiselección extendida, 3 columnas (la primera, el identificador, con ancho 0) | — | Muestra identificador, código y descripción de cada riesgo candidato | `:78-87`, `.form.txt:542-559` |
| `btnQuitarTodos` («Quitar Todos») | botón | `OnClick` | Desmarca todos los elementos de la lista | `:197-206`, `.form.txt:575-588` |
| `btnAceptar` («Aceptar») | botón | `OnClick` | Guarda la selección y cierra | `:125-196`, `.form.txt:630-642` |
| `btnCancelar` («Cancelar») | botón | `OnClick` | Cierra sin guardar | `:11-13`, `.form.txt:668-681` |
| `txtSeleccion` | cuadro de texto oculto | — | Sin handler ni referencias en el `.cls` | `.form.txt:710-725` |

## Flujos

### Apertura (`Form_Open`, `:14-32`; `Form_Load`, `:33-76`)

1. Sin `OpenArgs` muestra «Faltan parámetros de apertura.» y cancela la apertura (`:17-21`). Separa los argumentos por `|` y guarda `m_IdNC` y `m_IdExpediente` si son numéricos (`:24-30`).
2. En `Form_Load`: `Ajustar Me` y título (`:38-40`). Con `m_IdNC > 0` carga la NC con `constructor.getNCProyecto`; con 0 usa el objeto global `m_ObjNCProyectoActiva` y le asigna el expediente (`:45-53`). Sin NC, error «No se ha podido inicializar el contexto…» (`:56-58`).
3. Obtiene los candidatos (`:62`; `RiesgoServicio.cls:14-26`): la NC debe tener expediente («La No Conformidad no tiene un expediente asignado.», `RiesgoServicio.cls:20-23`). Rellena la lista con `AddItem "id;código;descripción"` (`:85`) y preselecciona los riesgos que ya tiene la NC (`:70, 100-116`).

### Aceptar (`btnAceptar_Click`, `:125-196`)

1. Recoge los elementos marcados y, por cada uno presente en los candidatos, lo añade a una colección y a un diccionario (`:138-149`).
2. **NC existente** (`m_IdNC > 0`): `ActualizarRiesgosNC` (`:156`), `SincronizarCampoTextoNC` (`:161`), sustituye `m_ObjNCProyectoActiva.Riesgos` (`:165`) e invalida la caché de la NC (`:167`). Si falla el primero, muestra «Error al aceptar la selección de riesgos» (`:157, 195`).
3. **NC en alta**: solo memoria; asigna el diccionario a `m_ObjNC.Riesgos` y compone `CodRiesgosAsociados` con los códigos separados por `|` (`:179-187`). Los persiste después `NCProyectoOperaciones.RegistrarDatosUnicos`, según el comentario (`:177-178`); esa clase no se ha revisado aquí.
4. Cierra el formulario (`:191`).

### Errores

Todos los handlers muestran un `MsgBox` crítico con `Err.Description`; `Form_Load` no cierra el formulario tras el error (`:74-75`).

## Reglas, validaciones y cálculos

- **Candidatos:** riesgos de la última edición (`MAX(IDEdicion)`) del proyecto del expediente y sin fecha de retirada (`RiesgoRepositorio.bas:57-61`).
- **Sin mínimo ni máximo:** se admite una selección vacía; en edición borra todas las asociaciones (`RiesgoRepositorio.bas:96`).
- **Fecha:** asociaciones con `FechaRegistro` y `FechaDecison` = hoy, `ParaNC` = «Sí» fijo (`RiesgoRepositorio.bas:97-105`).
- **Rol:** no hay comprobación de rol en el formulario.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Riesgos candidatos | `TbRiesgos`, `TbProyectosEdiciones`, `TbProyectos` | SQL dinámico concatenado (expediente) | `RiesgoRepositorio.bas:57-63` |
| Riesgos asociados a la NC | `TbRiesgos`, `TbRiesgosNC` | SQL dinámico concatenado (NC) | `RiesgoRepositorio.bas:19-23` |
| Borrar asociaciones previas | `TbRiesgosNC` | SQL dinámico concatenado, `dbFailOnError` | `RiesgoRepositorio.bas:96` |
| Insertar asociaciones | `TbRiesgosNC` (`ID` = `Max + 1` por fila) | SQL dinámico concatenado, `dbFailOnError` | `RiesgoRepositorio.bas:100-105`; `constructor.bas:1165-1178` |
| Actualizar texto denormalizado | `TbNoConformidades.CodigoRiesgo` | SQL dinámico concatenado | `RiesgoServicio.cls:72` |
| Invalidar caché | tabla de caché de NC | consulta con parámetro | `CacheNCProyecto.bas:367-392` |
| Log «Vínculo de Riesgos» | log de NC de proyecto | DAO (`RegistrarLogProyecto`) | `RiesgoServicio.cls:47` |

Sin transacciones: borrado e inserciones son sentencias independientes.

## Coexistencia

Pendiente de `../coexistencia.md`. Puntos a adaptar: fecha pasada como literal `#mm/dd/yyyy#` (`RiesgoRepositorio.bas:97`), identificador `Max + 1` por fila (`RiesgoRepositorio.bas:100`), `DELETE` + `INSERT` sin transacción (`RiesgoRepositorio.bas:96-105`), subconsulta `MAX` en SQL concatenada (`RiesgoRepositorio.bas:57-61`) y nombres con acento en las columnas de la unión (`FechaDecison` está escrito así en el código, `RiesgoRepositorio.bas:104`).

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual**

- El título de diseño del `.form.txt` no es el que ve el usuario (`:39-40`).
- En edición los cambios se persisten al pulsar «Aceptar», sin esperar a «Registrar» en el formulario padre (`:152-167`).

**Posibles defectos (a verificar)**

- `Form_Open` accede a `vArgs(1)` sin comprobar que existan dos partes (`:24-30`): un `OpenArgs` sin `|` provocaría un error de índice sin controlar.
- Una descripción con `;` se partiría en más columnas, ya que `AddItem` usa `;` como separador (`:85`).
- `SincronizarCampoTextoNC` no devuelve error al formulario (`:161`, `RiesgoServicio.cls:59-73`): un fallo de la lectura o del `UPDATE` no se notifica al usuario.
- En edición se actualiza `m_ObjNCProyectoActiva` (`:165, 169`), no `m_ObjNC` (`:46`); si ambos fueran objetos distintos, el cargado en `Form_Load` quedaría desfasado.
- `If m_ObjNCProyectoActiva.Riesgos Is Nothing` (`:169`) solo es cierto si el diccionario está vacío y la propiedad lo reinicializa al leerla (`NCProyecto.cls:130-145`), por lo que `CodRiesgosAsociados` se vacía pero la rama podría no ejecutarse nunca.
- `lstRiesgos.ItemData` se usa como identificador (`:110, 139`) y depende de que la primera columna sea la enlazada; no se ha verificado `BoundColumn`.
- `GuardarAsociaciones` borra primero y no es atómico (`RiesgoRepositorio.bas:96-105`): un fallo a mitad deja la NC con asociaciones parciales.

**Propuestas web pendientes de decisión**

- Sustituir la lista por un selector multiselección con filtro y búsqueda.
- Persistir borrado e inserciones en una transacción y devolver el error de sincronización al usuario.

## Criterios de aceptación

1. Sin `OpenArgs` el formulario no se abre y muestra «Faltan parámetros de apertura.» (`:17-21`).
2. La lista contiene solo riesgos de la última edición del expediente y sin retirar, con código y descripción visibles y el identificador oculto (`RiesgoRepositorio.bas:57-61`; `.form.txt:549, 559`).
3. En una NC con riesgos asociados, esos elementos aparecen marcados al abrir (`:100-116`).
4. «Quitar Todos» desmarca todos y no cierra ni guarda (`:197-202`).
5. «Cancelar» cierra sin escribir en base de datos ni en memoria (`:11-13`).
6. «Aceptar» en NC existente sustituye las asociaciones de la NC por la selección, actualiza `CodigoRiesgo` e invalida la caché (`:152-167`).
7. «Aceptar» en NC en alta no escribe en base de datos y deja `CodRiesgosAsociados` con los códigos separados por `|` (`:179-187`).
8. Una NC sin expediente no permite listar candidatos (`RiesgoServicio.cls:20-23`; `Form_FormNCProyectoGeneral.cls:1133-1136`).

## Pendientes

- `NCProyectoOperaciones.RegistrarDatosUnicos` (`:177-178`): referencia citada en comentario, no revisada en esta ficha.
- `BoundColumn` de `lstRiesgos`: no figura entre las propiedades revisadas del `.form.txt`.
- Esquema físico de `TbRiesgos`, `TbRiesgosNC` y `TbProyectosEdiciones`: no hay definición en `src`; a resolver con `../erd-origen.md`.
