# Form_FormAnexosNoAlcanzables

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormAnexosNoAlcanzables.cls` (219 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Herramienta de mantenimiento: busca los anexos de `TbAnexos` cuyo archivo no es accesible en la ruta configurada y permite resubirlos. Es un formulario **no enlazado**; la lista se rellena con `AddItem` desde `Constructor.getAnexosNoAlcanzablesBusqueda`. `FG.bas` = `modules/Funciones Generales.bas`. Las citas son `Form_FormAnexosNoAlcanzables.cls:línea` salvo indicación. Capacidad de todos los controles: CAP-11 Anexos.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpcionesOtras` | etiqueta `lblAnexosNoAlcanzables`: `DoCmd.OpenForm "FormAnexosNoAlcanzables"` con `On Error Resume Next` | `Form_Form0BDOpcionesOtras.cls:19-22` |

No recibe `OpenArgs` ni escucha eventos. `cmdSalir` solo cierra (`:161-164`).

Controles: 13 (25 `Begin` en el `.form.txt`: 1 `Form`, 8 estilos por defecto, 1 `FormHeader`, 1 `Section`, 1 `FormFooter` y 13 controles).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComboIDEventoBusqueda` | combo, 4 columnas | — | Filtro por evento: `IDEvento`, `BUI`, `SUBSISTEMA`, `Equipo`; `Column(0)` alimenta la búsqueda | `:25`; `.form.txt:451-480` |
| `ComboBUIBusqueda` | combo | — | Filtro por BUI (`TbBUI`) | `:26`; `.form.txt:510-514` |
| `ComandoFiltrar` («Filtrar») | botón | `OnClick` → `ComandoFiltrar_Click` | Ejecuta `Filtrar` | `:64-82`; `.form.txt:947-949` |
| `ListaFiltrados` | lista, 5 columnas, lista de valores | `OnClick` → `ListaFiltrados_Click` | Resultado; al pulsar carga el `Anexo` y habilita Resubir | `:166-218`; `.form.txt:375-394` |
| `ComandoResubir` («Resubir», deshabilitado al abrir) | botón | `OnClick` → `ComandoResubir_Click` | Sustituye el archivo del anexo seleccionado | `:84-130`; `.form.txt:422-424` |
| `cmdSalir` («Salir») | botón (pie) | `OnClick` → `cmdSalir_Click` | Cierra con `acSaveNo` | `:161-164`; `.form.txt:996-998` |
| `ComandoLimpiarIDEvento`, `ComandoLimpiarBUI` | 2 botones con icono | sin `OnClick` | Sin acción asociada visible | `.form.txt:591-593, 769-770` |
| `lblEstado` | etiqueta | — | Se oculta al abrir | `:144`; `.form.txt:1039-1040` |
| `Etiqueta142`, `Etiqueta267`, `lblTitulo1`, `lblTitulo2` | 4 etiquetas | — | Rótulos «IDEVENTO», «BUI»; título «ANEXOS NO ALCANZABLES» (`lblTitulo1` a `Me.Caption`) | `:145`; `.form.txt:540, 567, 328, 350` |

## Flujos

1. **Apertura** (`Form_Open`, `:135-160`): `AjustarTamaño`, oculta `lblEstado`, copia el título a `Me.Caption` (el `.form.txt` trae un título «GESTIÓN DE LA FACTURACIÓN», `.form.txt:60`). La lista arranca vacía.
2. **Filtrar** (`:64-82`, `Filtrar` `:10-60`): vacía la lista poniendo `RowSource = "IDAnexo;BUI;IDEvento;IDMaterial;Archivo"` (`:23`), llama a `getAnexosNoAlcanzablesBusqueda(evento, BUI)` (`:24-26`) y, por cada anexo, calcula el BUI desde la entidad (evento, actividad→evento o material→actividad→evento, `:36-46`) y añade `IDAnexo;BUI;IDEvento;IDMaterial;NombreArchivo` (`:47`). Sin evento ni BUI no devuelve nada (`Constructor.bas:262-264`).
3. **Selección** (`:166-218`): `Column(0)` vacío ⇒ limpia y sale (`:178-185`); si no, carga el `Anexo` con `getAnexo` si cambió (`:186-198`) y habilita Resubir (`:205`).
4. **Resubir** (`:84-130`): carga el anexo si falta (`:93-100`); si el archivo ya existe, «El archivo ya es alcanzable» (`:104-107`); `Seleccionar(True, …)` y salida en silencio si no es archivo (`:108-111`); `Anexo.Resubir` (`:115`; `Anexo.cls:818-855`). No refresca la lista.
5. **Error:** `MsgBox` crítico con `m_Error` (`:124-129`); en la apertura, además cierra el formulario (`:159`).

## Reglas, validaciones y cálculos

- **Definición de «no alcanzable»:** `FileExists(URLAnexo)` es falso (`Constructor.bas:354, 386`). `URLAnexo` = directorio de gastos o de documentación + `NombreArchivo` (`Anexo.cls:151-160`).
- **Prioridad del filtro:** si hay evento, se ignora el BUI (`Constructor.bas:266-272`); por BUI recorre los eventos con ese BUI y fusiona sin duplicar por `IDAnexo` (`Constructor.bas:276-309`).
- **Por evento:** anexos con ese `IDEvento` (`Constructor.bas:340-366`) y, además, todos los anexos con material (`Constructor.bas:373-398`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Combo de eventos | `TbEventos`, `TbEquipos` | `RowSource` SQL guardada (`LEFT JOIN`, orden por fecha de alta descendente) | — | `.form.txt:474-476` |
| Combo de BUI | `TbBUI` | `RowSource` SQL guardada | — | `.form.txt:512` |
| Eventos por BUI | `TbEventos` | SQL dinámico concatenado | no | `Constructor.bas:276-279` |
| Anexos por evento | `TbAnexos` | SQL dinámico concatenado | no | `Constructor.bas:340-343` |
| Anexos con material | `TbAnexos`, `TbMaterial` | SQL fija con `INNER JOIN` | no | `Constructor.bas:373-375` |
| Anexo seleccionado | `TbAnexos` | SQL dinámico concatenado | no | `Constructor.bas:219-221` |
| Resubida | `TbAnexos` y archivo | `UPDATE` concatenado si cambia el nombre + copia de archivo | no | `Anexo.cls:840-847` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: la consulta de anexos con material se recorre completa y filtra en VBA con `FileExists` por fila (`Constructor.bas:354, 386`); con un almacén de archivos distinto de la ruta de red actual cambiaría la definición de «alcanzable». SQL concatenado (`Constructor.bas:220, 278, 342`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoResubir_Click` llama a `fso.FileExists(m_AnexoSeleccionado)` con el objeto `Anexo`, no con su ruta (`:104`); `Anexo` no declara miembro predeterminado, así que podría fallar en ejecución o dar siempre falso, anulando el aviso «ya es alcanzable».
- `Anexo.Resubir` se invoca sin comprobar que `m_AnexoSeleccionado` no sea `Nothing` (`:93-115`).
- Una búsqueda por evento añade **todos** los anexos con material sin filtrar por evento (`Constructor.bas:373-398`), así que el resultado puede incluir anexos de otros eventos.
- Los anexos de subcontratación, gasto o seguimiento de material no tienen rama para calcular el BUI (`:36-46`) y salen con la columna vacía.
- La primera fila de la lista es la cabecera `IDAnexo;BUI;…` (`:23`) y `Column(0)` valdría `IDAnexo` (`:178`), que `getAnexo` concatenaría en SQL (`Constructor.bas:219-221`).
- `RowSource` de diseño (`ID;Título;Fecha;NombreFacturaAnexa`) y título de ventana «GESTIÓN DE LA FACTURACIÓN» son restos de otro formulario (`.form.txt:60, 392`).
- `ComandoFiltrar_Click` cita en su mensaje `ComboIDEventoBusqueda_Change` (`:78`) y `ListaFiltrados_Click` otro nombre mal escrito (`:215`).
- Los botones de limpiar no tienen handler (`.form.txt:591-593, 769-770`): los combos no se pueden vaciar con ellos.
- Tras resubir con éxito no se refresca la lista: el anexo sigue apareciendo hasta volver a filtrar (`:115-123`).
- `Seleccionar` devuelve `#ERR|…` y se ignora en silencio (`:108-111`; `FG.bas:1143-1147`); cancelar el diálogo pisa `URLArchivoUltimo` con vacío (`FG.bas:1139-1141`).
- Coste: la consulta por evento recorre toda `TbAnexos`+`TbMaterial` y comprueba cada archivo en red (`Constructor.bas:373-398`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: informe de alcanzabilidad calculado en servidor y resubida con refresco de la fila.

## Criterios de aceptación

Propuestas, no aprobadas:

1. Al abrir, la lista está vacía y Resubir deshabilitado.
2. Filtrar por un evento lista sus anexos cuyo archivo no es accesible, con BUI, evento, material y nombre de archivo.
3. Filtrar por BUI lista los anexos no alcanzables de todos los eventos de ese BUI, sin duplicados.
4. Filtrar sin evento ni BUI no lista nada.
5. Seleccionar una fila habilita Resubir; seleccionar un valor vacío lo deshabilita.
6. Resubir sobre un archivo ya alcanzable avisa «El archivo ya es alcanzable» y no escribe.
7. Resubir con archivo válido copia el archivo a su destino y actualiza `NombreArchivo` si cambia la extensión.
8. Los botones de limpiar vacían su filtro, o se retiran (pendiente de decisión: hoy no hacen nada).

## Pendientes

- Qué anexos se esperan al buscar por evento (si el `INNER JOIN` sin filtro de evento es intencionado), `Constructor.bas:373-374`.
- `Anexo.Resubir` cuando `m_NombreFinal` cambia: no hay actualización del objeto cargado (`Anexo.cls:840-847`).
- Efecto de `Evento.BUI`, `Actividad.Evento` y `Material.Actividad` al calcular `m_BUI` (`:37, 40, 43`), fuera de este lote.
- Contenido de `DoCmd.Restore` en `AjustarTamaño` (`FG.bas:261`).
