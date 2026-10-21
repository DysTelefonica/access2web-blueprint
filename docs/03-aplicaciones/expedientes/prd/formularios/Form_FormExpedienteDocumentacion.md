# Form_FormExpedienteDocumentacion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpedienteDocumentacion.cls` (329 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Anexos» de la ficha del expediente: lista los documentos anexados, permite añadir, eliminar y abrir uno. Formulario no enlazado (sin `RecordSource`); trabaja sobre el DTO global `m_ObjExpedienteDTOActivo`. Las citas son `Form_FormExpedienteDocumentacion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormExpediente` | botón de navegación `NavAnexos` con destino `FormExpedienteDocumentacion` dentro de `FrmDetalle`; al entrar vuelca el DTO de la pestaña anterior | `Form_FormExpediente.form.txt:940, 958`; `Form_FormExpediente.cls:391-413` |

No tiene botón de salida ni lanza eventos. Lee el expediente de `m_ObjExpedienteDTOActivo` y el permiso de `Form_FormExpediente.m_AbiertoParaEditar` y `EsAdministrador` (`:287, 293-294`). `Form_Load` sustituye el rótulo de diseño por `m_TituloFormulario` (`:175`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaDocumentos` | ListBox, lista de valores, 2 columnas | `OnClick` → `ListaDocumentos_Click` | Muestra `ImagenAnexo` solo si hay fila seleccionada | `.cls:220-248`, `.form.txt:589-612` |
| `ListaDocumentos` | — | `OnDblClick` → `ListaDocumentos_DblClick` | Si `ImagenAnexo` es visible, abre el documento (`ImagenAnexo_Click`); `On Error Resume Next` | `.cls:322-330` |
| `ComandoAltaAnexo` (sin rótulo útil; tooltip «Alta»; `Tag = "EJECUTIVO"`) | botón | `OnClick` → `ComandoAltaAnexo_Click` | Selecciona un fichero y lo anexa | `.cls:10-77`, `.form.txt:735-740` |
| `ComandoEliminar` (tooltip «Eliminar elemento seleccionado»; `Tag = "EJECUTIVO"`) | botón | `OnClick` → `ComandoEliminar_Click` | Elimina el anexo seleccionado tras confirmar | `.cls:80-160`, `.form.txt:640-646` |
| `ImagenAnexo` (oculta por defecto) | imagen | `OnClick` → `ImagenAnexo_Click` | Abre el documento seleccionado (`AbrirEnLocal`) | `.cls:195-218`, `.form.txt:817-831` |

`ListaDocumentos`: columna 0 (oculta, ancho 0) = ruta completa; columna 1 = nombre del fichero (`.form.txt:594, 609`; filas añadidas en `:42, 267`). El `RowSource` del diseño conserva valores residuales; `RellenarListas` lo vacía al cargar (`:263`).

## Flujos

### Apertura (`Form_Load`, `.cls:166-193`)

1. Fija `AllowEdits = True` (`:176`) y llama a `EstablecerDatos` (`:177`).
2. `EstablecerDatos` exige `m_ObjExpedienteDTOActivo.Expediente`; si falta, «El formulario se ha abierto con parámetros insuficientes» (`:287-290`).
3. Los controles con `Tag = "EJECUTIVO"` (alta y eliminar) se habilitan solo si `m_AbiertoParaEditar` **y** `EsAdministrador = Sí` (`:291-303`).
4. `RellenarListas` recorre `ColAnexos` y añade «`URLDocumento`;nombre» por cada `ExpedienteAnexo`; selecciona la primera fila (`:261-271`).

### Alta de anexo (`ComandoAltaAnexo_Click`, `.cls:10-77`)

1. `Seleccionar` abre un selector de fichero (`:19`; `FUNCIONES UTILES.bas:2188-2228`). Si se cancela o el fichero no existe, sale sin aviso (`:23-29`).
2. **Expediente sin `IDExpediente` (alta):** guarda en `ColAnexos` la clave = nombre del fichero y el valor = ruta local, si la clave no existe ya; añade la fila a la lista (`:34-45`). La persistencia queda para `ExpedienteOperaciones.Registrar` (`ExpedienteOperaciones.cls:459-466`).
3. **Expediente existente:** `ExpedienteOperaciones.RegistrarAnexo` escribe en el acto (`:47-54`); si devuelve `"1"`, recarga `ColAnexos` desde `Expediente.Anexos` y rehace la lista (`:55-59`).
4. Selecciona `ListaDocumentos.Selected(ListCount)` (`:63`).

### Eliminación (`ComandoEliminar_Click`, `.cls:80-160`)

1. Sin fila seleccionada, sale sin aviso (`:93-99`). No administrador: «Operacion no autorizada» (`:100-103`).
2. Confirmación «¿Desea realmente eliminar el anexo seleccionado?», con «No» por defecto (`:107-110`).
3. **Alta:** intenta quitar de `ColAnexos` la clave `Column(0)` (ruta) y quita la fila de la lista (`:114-132`).
4. **Existente:** `ExpedienteOperaciones.EliminarAnexo` borra registro y fichero; recarga la lista (`:134-144`).

### Apertura de documento (`ImagenAnexo_Click`, `.cls:195-218`)

`AbrirEnLocal` (`FUNCIONES UTILES.bas:3033-3075`): si la ruta no existe, «No es accesible la ruta del archivo que se pretende abrir»; si es UNC, copia el fichero al directorio local del entorno (falla con «Tiene el archivo abierto» si la copia local está abierta) y lo abre con el programa asociado (`:3048-3066`).

### Errores

Todos los handlers muestran `MsgBox` crítico con `m_Error`; los errores no controlados (distintos de 1000) además envían `CorreoAlAdministrador` (`:71-76, 154-159, 187-192, 212-217, 242-247`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Alta y eliminar habilitados solo con edición y rol administrador | `EstablecerDatos` | `.cls:291-303` |
| Eliminar comprueba de nuevo `EsAdministrador`; alta no lo comprueba en código | handlers | `.cls:100-103`; `:10-77` |
| En alta no se admite dos veces el mismo nombre de fichero (comparación de texto) | `ComandoAltaAnexo_Click` | `.cls:38, 40` |
| En existente, un nombre ya registrado no se vuelve a insertar ni copiar, sin aviso | `RegistrarAnexo` | `ExpedienteOperaciones.cls:1557-1571` |
| `IDDocumento` = máximo + 1 (`DameID`) | `ExpedienteAnexo.IDDocumentoCalculado` | `ExpedienteAnexo.cls:95-106`; `FUNCIONES UTILES.bas:2851-2864` |
| Directorio del anexo = directorio de documentación del entorno + `IDExpediente` a 5 dígitos | `getURLDirectorioExpediente` | `ExpedienteAnexo.cls:41`; `FUNCIONES UTILES.bas:2310` |
| Al guardar un alta, cada clave de `ColAnexos` no numérica debe existir como fichero | `MotivoNoOK` | `ExpedienteOperaciones.cls:1044-1055` |

## Lecturas y escrituras

| Operación | Tabla o recurso | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Carga de anexos | `TbExpedientesAnexos` por `IDExpediente`; diccionario por `IDDocumento` | SQL dinámico DAO | no | `constructor.bas:4943-4968` |
| Alta (existente) | `SELECT` por expediente y nombre; `AddNew` en `TbExpedientesAnexos` (`IDDocumento`, `IdExpediente`, `NombreDocumento`); crea carpeta y copia el fichero | SQL dinámico DAO + sistema de ficheros | no (sin `p_db`, usa `getdb()`) | `ExpedienteOperaciones.cls:1538, 1557-1570` |
| Alta (expediente nuevo) | mismo `RegistrarAnexo` por cada clave de `ColAnexos`, con la conexión de `Registrar` | SQL dinámico DAO | sí, la de `Registrar` | `ExpedienteOperaciones.cls:459-466` |
| Eliminación (existente) | `DELETE` en `TbExpedientesAnexos` por expediente y nombre; borra el fichero si no está abierto (errores ignorados); `RegistrarUltimoCambio` | SQL dinámico DAO + ficheros | no | `ExpedienteOperaciones.cls:2658-2675` |
| Apertura | copia local del fichero y ejecución | sistema de ficheros | — | `FUNCIONES UTILES.bas:3052-3066` |

## Coexistencia

Pendiente de `coexistencia.md`. Puntos visibles: SQL concatenado con `NombreDocumento` entre comillas simples; `DameID` (máximo + 1) en lugar de autonumérico; `DELETE *` es sintaxis Jet; los ficheros viven en un recurso compartido fuera de la base y su alta y baja no forman parte de ninguna transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- En alta, `ColAnexos` usa como clave el nombre del fichero y como valor la ruta (`:41`), pero `Registrar` pasa la clave a `RegistrarAnexo` (`ExpedienteOperaciones.cls:463`) y `MotivoNoOK` comprueba `FileExists` sobre la clave (`:1049`): con un nombre sin ruta, el guardado fallaría con «El anexo no es accesible».
- En alta, `ComandoEliminar` busca en `ColAnexos` la ruta (`Column(0)`, `:93, 117`) pero la clave es el nombre: la fila desaparece de la lista y el anexo sigue en el DTO.
- En alta, si se vuelve a la pestaña, `RellenarListas` hace `Set m_Anexo = m_Col(m_ID)` sobre valores de texto (`:266`): error de tipos esperado al recargar.
- `Selected(ListCount)` (`:63`) apunta una posición más allá de la última fila.
- `RegistrarAnexo` no tiene manejador propio; si `URLDirectorio` falla, eleva 1000 sin rellenar `p_Error` (`ExpedienteOperaciones.cls:1552`) y el formulario muestra un `MsgBox` vacío. El comentario de rollback (`:1568`) no aplica fuera de `Registrar`.
- Alta en existente no llama a `RegistrarUltimoCambio`; la eliminación sí (`ExpedienteOperaciones.cls:2675`).
- `EliminarAnexo` usa `Execute` sin `dbFailOnError` (`:2663`) y borra el fichero del recurso compartido de forma irreversible.
- `Seleccionar` llama a `.Show` antes de fijar título, filtro y carpeta inicial (`FUNCIONES UTILES.bas:2208-2214`): el diálogo no los aplicaría.
- Nombres con comilla simple rompen el SQL concatenado (`ExpedienteOperaciones.cls:1557, 2662`).
- Tooltip de `ImagenAnexo` copiado de otra pantalla («Evidencia del acuerdo de revisión de los Riesgos…», `.form.txt:829`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: almacenamiento de documentos con identificador propio en vez de nombre de fichero; alta y baja de fichero y registro coherentes ante fallo; descarga en lugar de copia local.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin permiso de edición o sin rol administrador, alta y eliminación no están disponibles (regla 1).
2. Anexar un fichero ya anexado al mismo expediente no crea un segundo registro (reglas 3-4).
3. Eliminar pide confirmación con «No» por defecto y, al aceptar, el anexo desaparece de la lista y de `TbExpedientesAnexos`.
4. En un expediente nuevo, los anexos añadidos se persisten al guardar y uno eliminado antes de guardar no se persiste.
5. Abrir un documento inexistente muestra «No es accesible la ruta del archivo que se pretende abrir».

## Pendientes

- `FicheroAbierto` (`FUNCIONES UTILES.bas:2873`) y `Ejecutar` no se han analizado.
- Rótulo efectivo de los botones (`Caption = "Comando247"` con imagen) en ejecución.
- Si `DameID` con `getdb()` ve las filas no confirmadas de la transacción de `Registrar` cuando un alta lleva varios anexos.
- Origen de `URLDirectorioDocumentacion` y `URLDirectorioLocal` en `m_ObjEntorno`.
