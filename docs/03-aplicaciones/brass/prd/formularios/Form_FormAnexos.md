# Form_FormAnexos

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormAnexos.cls` (606 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Gestión de los documentos anexos de una entidad (evento, actividad, material, seguimiento de material, subcontratación o gasto): lista, adjuntar, abrir, resubir, eliminar y abrir el parte del evento. Es un formulario **no enlazado**; la entidad llega en la global `m_ObjEntidadParaAnexoActiva` (`Variables Globales.bas:252`) y la lista se rellena con `AddItem` desde `ColAnexos` de esa entidad. Persiste mediante la clase `Anexo`. `FG.bas` = `modules/Funciones Generales.bas`. Las citas son `Form_FormAnexos.cls:línea` salvo indicación. Capacidad de todos los controles: CAP-11 Anexos.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEventoGestion` (evento y actividad) | cierra `FormAnexos` si está abierto, asigna la entidad a la global y lo abre con el nombre del formulario llamante en `OpenArgs` | `Form_FormEventoGestion.cls:931-939, 982-988` |
| `Form_FormActividadGestion`, `Form_FormMaterialesGestion`, `Form_FormMaterialRepuestoEdicion`, `Form_FormMaterialReparacionEdicion`, `Form_FormMaterialSeguimiento`, `Form_FormSubContratacionGestion`, `Form_FormGastoGestion`, `Form_FormPlanificacionesBusqueda` | mismo patrón | `Form_FormActividadGestion.cls:220-226`; `Form_FormMaterialesGestion.cls:107-112`; `Form_FormMaterialRepuestoEdicion.cls:142-148`; `Form_FormMaterialReparacionEdicion.cls:115-121`; `Form_FormMaterialSeguimiento.cls:62-67`; `Form_FormSubContratacionGestion.cls:145-151`; `Form_FormGastoGestion.cls:77-82`; `Form_FormPlanificacionesBusqueda.cls:834-840` |

Sin entidad activa: «No se ha abierto el formulario con los datos necesarios» y se cierra (`:324-327, 338-344`). `cmdSalir` cierra también el formulario llamante (`m_FormularioQueAbre`) si sigue abierto (`:97-100`). Tras adjuntar o eliminar, relanza `ListaEventos_Click`/`ListaActividades_Click`/`ListaMateriales_Click` de `FormEventoGestion` y `ListaFiltrados_Click` de `FormActividadGestion` si están abiertos (`:65-77, 193-205`).

Controles: 28 (38 `Begin` en el `.form.txt`: 1 `Form`, 8 estilos por defecto, 1 `Section` y 28 controles).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaDocumentos` | lista, 4 columnas (`0;6237;0;0`), lista de valores | `OnClick` → `ListaDocumentos_Click` | Muestra el título; al pulsar carga el `Anexo` y habilita Anexo/Eliminar, y Resubir si el archivo no es accesible | `:471-521`; `.form.txt:342-360` |
| `TituloDoc` | texto no enlazado | — | Título del nuevo anexo; obligatorio al adjuntar | `:32-35, 42, 78`; `.form.txt:435` |
| `ComandoAnexar` («Adjuntar») | botón | `OnClick` → `ComandoAnexar_Click` | Registra un archivo elegido por diálogo | `:20-89`; `.form.txt:958-960` |
| `ComandoAnexo` («Anexo», deshabilitado al abrir) | botón | `OnClick` → `ComandoAnexo_Click` | Abre el archivo del anexo seleccionado | `:105-147`; `.form.txt:819-821` |
| `ComandoResubir` («Resubir») | botón | `OnClick` → `ComandoResubir_Click` | Sustituye el archivo de un anexo existente | `:264-308`; `.form.txt:1131-1133` |
| `ComandoEliminarDocumento` («Eliminar», deshabilitado al abrir) | botón | `OnClick` → `ComandoEliminarDocumento_Click` | Elimina registro y archivo tras confirmar | `:149-222`; `.form.txt:464-466` |
| `ComandoParte` («Parte», deshabilitado al abrir) | botón | `OnClick` → `ComandoParte_Click` | Abre el parte del evento | `:223-262`; `.form.txt:859-861` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra este y el llamante | `:90-104`; `.form.txt:505-507` |
| `IDEVENTO`, `IDActividad`, `IDMaterial`, `IDMaterialSeguimiento`, `IDSubContratacion`, `IDGasto` | 6 textos no enlazados, bloqueados | — | Muestran el identificador de la entidad activa (solo el que corresponde) | `:355-366, 370-402`; `.form.txt:627, 687, 747, 930, 1042, 1102` |
| `Etiqueta142`, `Etiqueta4`, `Etiqueta6`, `Etiqueta8`, `Etiqueta12`, `Etiqueta15`, `Etiqueta17`, `lblTitulo1`, `lblTitulo2` | 9 etiquetas | — | Rótulos («TÍTULO», «EVENTO», «ACTIVIDAD», «MATERIAL», «SubCon», «MATERIAL SEG», «Gasto», «ANEXOS»); `lblTitulo1` a `Me.Caption` | `:323`; `.form.txt:548, 598, 658, 718, 901, 1013, 1073, 777, 798` |
| `CuadroSuperior`, `Cuadro2`, `Marco103`, `Marco104`, `Marco3` | 2 rectángulos y 3 marcos | — | Decoración | `.form.txt:305, 328, 388, 411, 573` |

## Flujos

### Apertura (`Form_Open`, `:310-345`)

Ajusta tamaño, guarda `OpenArgs` y exige la entidad activa (`:319-327`). `EstablecerDatos` (`:346-421`) deshabilita Anexo, Parte y Eliminar, bloquea los seis cuadros de identificador, rellena el que corresponde según el tipo de la entidad (`TypeOf`, `:369-406`) y habilita `ComandoParte` si el evento tiene `Parte.URLParte` (`:373-382`); tipo desconocido ⇒ «Tipo de Anexo no reconocido» (`:403-405`). Termina con `EstablecerLista` (`:409`).

### Lista (`EstablecerLista`, `:555-604`)

Con `Sí` relee la entidad con `Constructor.get*` (`ActualizarEntidad`, `:523-553`). Pone `RowSource = "ID;IDAnexo;Título;Archivo"` (`:574`), recorre `ColAnexos` y añade `IDAnexo;Titulo;NombreArchivo` (`:583-587`). Con un único anexo (`ListCount = 2`, cabecera incluida) lo selecciona y ejecuta `ListaDocumentos_Click` (`:591-594`).

### Adjuntar (`ComandoAnexar_Click`, `:20-89`)

1. Título obligatorio: «El título es obligatorio» (`:32-35`).
2. `AsignarTipoAnexo` crea un `Anexo` nuevo y lo vincula a la entidad (`:36, 423-470`); asigna el título y deshabilita Eliminar, Anexo y Parte (`:42-45`).
3. `Seleccionar(True, …)` abre el diálogo de archivo (`:50`; `FG.bas:1117`); si el resultado no es un archivo existente, sale en silencio (`:51-53`).
4. `m_ObjAnexo.Registrar` (`Anexo.cls:690-756`), `EstablecerLista Sí`, refresco de pantallas abiertas, vacía `TituloDoc` (`:57-78`).

### Otras acciones

- **Abrir** (`:105-147`): comprueba que el archivo existe («Parece que el archivo no se pudo subir en su momento. Tiene que resubirlo», `:127-130`) y llama a `AbrirEnLocal` (`:131`; `FG.bas:218-260`), que copia el archivo al directorio local de trabajo y lo ejecuta.
- **Resubir** (`:264-308`): selecciona archivo (`:283`) y llama a `Anexo.Resubir` (`:290`; `Anexo.cls:818-855`).
- **Eliminar** (`:149-222`): confirma con «No» por defecto (`:175-180`), llama a `Anexo.Eliminar` (`:184`; `Anexo.cls:912-938`), recarga la lista y deshabilita los tres botones (`:188-210`).
- **Parte** (`:223-262`): exige que la entidad sea `Evento` («La entidad del anexo no es un evento», `:232-235`) y que tenga parte («El evento no tiene parte registrado», `:241-244`).
- **Errores:** `MsgBox` crítico con `m_Error`; el formulario sigue abierto (`:83-88` y equivalentes).

## Reglas, validaciones y cálculos

- **Título obligatorio** en el formulario (`:32-35`) y en la clase (`Anexo.cls:667-670`).
- **Título único por entidad:** `TituloRepetido` consulta `TbAnexos` por título e identificador de entidad (`Anexo.cls:856-910`); repetido ⇒ «El título ya ha sido usado para esa entidad» (`Anexo.cls:675-677`).
- **Identificador:** `IDAnexoCalculado` = máximo de `TbAnexos.IDAnexo` + 1 mediante `DameID` (`Anexo.cls:88`; `FG.bas:371-428`).
- **Nombre del archivo:** `<ID de entidad>_<ordinal de 2 cifras>.<extensión>`, con `/` sustituido por `-` (`Anexo.cls:427-430`); el directorio destino es el de gastos para gastos y el de documentación para el resto (`Anexo.cls:422-426`).
- **Archivo en uso:** no se registra si el origen o el destino parecen abiertos (`Anexo.cls:627-630, 661-665`).
- **Resubir** conserva la base del nombre y adopta la extensión del nuevo archivo; si cambia el nombre, actualiza `NombreArchivo` (`Anexo.cls:832-844`).
- **Eliminar:** impide borrar si el archivo está abierto (`Anexo.cls:921-926`); borra primero el registro y después el archivo (`Anexo.cls:927-932`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Cargar anexo seleccionado | `TbAnexos` | SQL dinámico concatenado (`WHERE IDAnexo=`) | no | `Constructor.bas:219-221` |
| Relectura de la entidad y sus anexos | según entidad | métodos `Constructor.get*` | no | `:530-541` |
| Alta de anexo | `TbAnexos` | DAO `AddNew`/`Update` sobre la tabla; después copia de archivo | no | `Anexo.cls:713-745` |
| Comprobación de título repetido | `TbAnexos` | SQL dinámico concatenado, con reintento si falla | no | `Anexo.cls:880-892` |
| Cambio de nombre en resubida | `TbAnexos` | `Execute` de `UPDATE` concatenado | no | `Anexo.cls:841-843` |
| Baja | `TbAnexos` | `Execute` de `DELETE` concatenado; luego borrado del archivo | no | `Anexo.cls:927-932` |
| Archivos | sistema de ficheros de red configurado | `FileSystemObject` (copiar, borrar, existe) | no | `Anexo.cls:745, 931`; `FG.bas:243-246` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `IDAnexo` calculado con `Max + 1` fuera de transacción (`FG.bas:371-428`) debe pasar a ser autonumérico o secuencia en PostgreSQL; SQL concatenado con `DELETE *` (`Anexo.cls:927`, dialecto Jet); el almacenamiento de archivos en rutas de red configuradas queda fuera de la base y debe seguir accesible en coexistencia; el vínculo entre registro y archivo no es atómico.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Anexo.Registrar` inserta la fila antes de copiar el archivo (`Anexo.cls:713-745`): si la copia falla, queda un registro sin archivo. El `If p_Error` posterior a la copia no puede detectar nada (`Anexo.cls:746-748`).
- `Anexo.Eliminar` borra el registro y luego el archivo (`Anexo.cls:927-932`): un fallo en el segundo paso deja un archivo huérfano.
- `ComandoAnexo_Click` usa `m_ObjAnexoSeleccionado` sin comprobar que sea `Nothing` si la lista no tiene selección (`:114-123`): error de objeto no definido, mostrado con el mensaje genérico (`:141-146`).
- `ComandoResubir_Click` no comprueba `m_ObjAnexoSeleccionado` antes de `Resubir` (`:272-290`).
- La primera fila de la lista es la cabecera `ID;IDAnexo;Título;Archivo` (`:574`) y la columna 0 de las filas es `IDAnexo` (`:586`), pero los nombres de cabecera tienen cuatro columnas y los datos tres (`:574` frente a `:586`); pulsar la fila de cabecera daría `Column(0) = "ID"`, que `getAnexo` concatenaría en SQL (`Constructor.bas:219-221`).
- `Seleccionar` devuelve `#ERR|…` si falla (`FG.bas:1143-1147`); `ComandoAnexar_Click` solo comprueba si el texto es un archivo existente (`:50-53`), por lo que el error se ignora en silencio. Al cancelar el diálogo, `URLArchivoUltimo` se pisa con vacío (`FG.bas:1139-1141`).
- Tras eliminar, el refresco usa `m_ObjAnexo.TipoAnexo` (`:194-199`) y no el anexo eliminado; correcto solo porque ambos comparten entidad.
- `ListaDocumentos_Click` no restablece `ComandoParte`; el botón depende solo de la apertura (`:380`, `:471-521`).
- Títulos con apóstrofo rompen la comprobación de título repetido (`Anexo.cls:880-882`).
- `Dim m_ID` local y `Private m_ID` de módulo con el mismo nombre (`:8`, `:473`), y variables de módulo sin uso (`m_URLAnexo`, `m_TipoAnexo`, `lst`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: registrar la fila solo tras copiar el archivo (o con estado pendiente), identificador generado por la base, y ocultar el nombre físico de la ruta al usuario.

## Criterios de aceptación

Propuestas, no aprobadas:

1. Sin entidad activa, el formulario muestra «No se ha abierto el formulario con los datos necesarios» y se cierra.
2. Solo el cuadro de identificador de la entidad activa aparece relleno y todos están bloqueados (`IDEVENTO`…`IDGasto`).
3. Adjuntar con `TituloDoc` vacío muestra «El título es obligatorio» y no abre el diálogo.
4. Adjuntar con un título ya usado por la misma entidad muestra «El título ya ha sido usado para esa entidad» y no escribe.
5. Adjuntar válido crea el registro y el archivo con nombre `<ID>_<nn>.<ext>`, refresca la lista y las pantallas abiertas y vacía `TituloDoc`.
6. Seleccionar un anexo habilita Anexo y Eliminar; Resubir solo si el archivo no es accesible.
7. Eliminar pide confirmación («No» por defecto) y, si se acepta, borra registro y archivo.
8. Parte solo está habilitado para un evento con parte registrado y abre ese documento.
9. Con un único anexo, se selecciona automáticamente al abrir.

## Pendientes

- Tabla y columnas reales de `TbAnexos` y relación con `TbMaterial` (ERD de origen).
- Código de `Anexo.GetURLAnexoCalculado`, `DameNumeroAnexosPorClase` y `getID` solo leído parcialmente (`Anexo.cls:384-442, 554`).
- Valor de las rutas de directorio de anexos en el entorno (`Anexo.cls:152-156`): configuración, no se copia.
- Si `Form_FormEventoGestion.ListaEventos_Click` y similares son públicos y su efecto exacto (`:66-71`).
