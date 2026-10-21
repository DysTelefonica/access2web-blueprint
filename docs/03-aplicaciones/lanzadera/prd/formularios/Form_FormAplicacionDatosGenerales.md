# Form_FormAplicacionDatosGenerales

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormAplicacionDatosGenerales.cls` (269 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Datos Generales» de la gestión de aplicaciones: ficha de configuración de una aplicación del catálogo (nombres, ejecutable, backend, iconos, carpetas, indicadores de pruebas y oficina, contraseña de la aplicación). Es un formulario **no enlazado** (sin `RecordSource`): lee la global `m_ObjAplicacionActiva` y guarda con `Aplicacion.Registrar` en `TbAplicaciones`. Las citas son `Form_FormAplicacionDatosGenerales.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormAplicacionGeneral` | `MENU2_Click` fija `FormDetalle.SourceObject = "FormAplicacionDatosGenerales"` | `Form_FormAplicacionGeneral.cls:73-77` |

No abre otros formularios ni lanza eventos. Recibe la aplicación por la global `m_ObjAplicacionActiva` (`Variables Globales.bas:101`), que fija la pestaña de búsqueda (`Form_FormAplicacionesGestion.cls:108`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComandoRegistrar` | botón, deshabilitado en diseño | `OnClick` | Valida cambios y guarda la aplicación | `:175-237`, `.form.txt:942-956` |
| `IDAplicacion` | texto oculto y deshabilitado | — | ID de la aplicación cargada o recién creada | `:149-150, 222`, `.form.txt:1420-1438` |
| `NombreAplicacion` («Nombre»), `NombreCorto` («Corto»), `NombreEjecutable` («Ejecutable») | textos | — | Identificación de la aplicación | `:96-103, 195-197`, `.form.txt:925, 1032, 1089` |
| `Pass` («Pass») | texto sin máscara | — | Contraseña de la aplicación; se muestra y se guarda tal cual | `:104-106, 198`, `.form.txt:1201-1214` |
| `EjecucionEnOficina` («En OFicina»), `EnPruebas`, `ConIconoEnLanzadera` | combos `Sí;No` limitados a la lista | — | Indicadores de la aplicación | `:107-116, 199-201`, `.form.txt:1135-1154, 1258-1277, 1353-1372` |
| `NombreArchivoDatos` («A. Datos») | texto | — | Fichero del backend | `:117-119, 202`, `.form.txt:1402` |
| `NombreIcono`, `NombreIconoLanzadera`, `NombreIconoParaArbol` | textos | — | Iconos | `:120-129, 203-205`, `.form.txt:1542, 1599, 1998` |
| `NombreCarpeta`, `NombreCarpetaTemporal`, `NombreCarpetaDocumentacion`, `NombreDirectorioIconos`, `NombreDirectorioAyuda`, `NombreDirectorioRecursos` | textos | — | Carpetas de la aplicación | `:130-148, 206-211`, `.form.txt:1656-1941` |

Los controles llevan `Tag` «PROTEGIDO», «NOPROTEGIDO» o «DATO» (`.form.txt:927-2000`); ningún código de `src` lee esos valores (búsqueda de `PROTEGIDO` sin resultados en `.cls`/`.bas`).

## Flujos

### Apertura (`Form_Open`, `:240-269`)

1. `EstablecerDatos` (`:251`). Sin aplicación activa, crea una `Aplicacion` vacía en la global y **sale sin tocar permisos** (`:90-93`).
2. Con aplicación activa, copia cada propiedad no vacía a su control (`:94-148`). Con `IDAplicacion`, guarda una copia recién leída de la base en `m_ObjAplicacionAlInicio` (`:149-155`).
3. Administrador (`EsAdministradorCalculado = Sí`): habilita `ComandoRegistrar` y la edición; resto: los deshabilita (`:158-164`).
4. Error: correo al administrador salvo el 1000, `MsgBox` crítico y cierre del formulario (`:260-268`).

### Guardado (`ComandoRegistrar_Click`, `:175-237`)

1. Si hay copia inicial y `HaHabidoCambios` es falso, aborta con «Ningún cambio que registrar» (`:182-190`).
2. Copia los 17 controles a la global con `Nz` (`:194-211`) y llama a `Registrar` (`:212`).
3. Relee la aplicación guardada, la deja como copia inicial y muestra su ID (`:218-222`).
4. Error: acumulado en `m_Error`, correo al administrador salvo el 1000 y `MsgBox` crítico (`:229-236`).

## Reglas, validaciones y cálculos

`Aplicacion.Registrar` valida en este orden y corta en la primera regla incumplida (`Aplicacion.cls:1794-1865`):

| Regla | Mensaje | Evidencia |
|---|---|---|
| `NombreAplicacion` obligatorio y único (excluye la propia aplicación al editar) | «Ha de indicar el NombreAplicacion» · «Ya existe en la base de datos un nombre idéntico» | `Aplicacion.cls:1794-1805, 1757-1761` |
| `NombreCorto` obligatorio y único | «Ha de indicar el NombreCorto» · «… un nombre corto idéntico» | `Aplicacion.cls:1806-1817` |
| Obligatorios: ejecutable, archivo de datos, carpeta temporal, carpeta, icono, icono de lanzadera, icono de árbol, directorio de iconos y de recursos | «Se ha de indicar …» | `Aplicacion.cls:1818-1845, 1850-1857` |
| `EjecucionEnOficina`, `EnPruebas`, `ConIconoEnLanzadera` ∈ {Sí, No} | «Se ha de indicar la …» | `Aplicacion.cls:1846-1849, 1858-1865` |

`Pass`, `NombreCarpetaDocumentacion` y `NombreDirectorioAyuda` son opcionales. Al editar, un campo vacío se guarda como `Null` (`Aplicacion.cls:1938-2022`); en un alta, los vacíos no se asignan (`Aplicacion.cls:1883-1933`).

`HaHabidoCambios` compara 15 campos distintos (16 comparaciones) con la copia inicial (`:11-76`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura inicial y relectura | `TbAplicaciones` por `IDAplicacion` | SQL dinámico concatenado | no | `:151, 218`; `Constructor.bas:482-484` |
| Unicidad de nombre y nombre corto | `TbAplicaciones` | SQL dinámico concatenado | no | `Aplicacion.cls:1757-1765` |
| Alta: ID nuevo | `Max(IDAplicacion) + 1` | SQL dinámico | no | `Aplicacion.cls:1878`, `Aplicacion.cls:1153`, `Funciones Genreales.bas:473-486` |
| Alta o edición | `TbAplicaciones` | DAO `AddNew`/`Edit` + `Update` | no | `Aplicacion.cls:1867-2024` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: ID calculado en VBA con `Max + 1`, sin bloqueo frente a altas concurrentes (`Funciones Genreales.bas:473-486`); validación de unicidad en VBA, sin constancia de índice único; textos concatenados sin escapar comillas (`Aplicacion.cls:1758`); valores «Sí»/«No» como texto de dos caracteres.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **Alta probablemente inalcanzable o destructiva.** Sin aplicación activa, `EstablecerDatos` sale antes de habilitar la edición (`:90-93`), así que el botón queda deshabilitado (`.form.txt:943`) y el formulario no editable (`.form.txt:11`). Si se llegara a guardar sin `IDAplicacion`, `Registrar` abre la tabla entera y, si no está vacía, entra por `.Edit` y **sobrescribe la primera fila** en lugar de añadir (`Aplicacion.cls:1867-1876, 1936-1937`).
- En un alta, el ID calculado se asigna a la propiedad pero no al campo `IDAplicacion` del registro (`Aplicacion.cls:1878`); si la columna no es autonumérica, la fila queda sin ID y la relectura de `:218-222` falla.
- `HaHabidoCambios` compara `NombreIconoLanzadera` dos veces (`:36, 48`) y no compara `ConIconoEnLanzadera` ni `NombreIconoParaArbol`, que sí se guardan (`:201, 204`): cambiar solo esos dos campos da «Ningún cambio que registrar».
- `ComandoRegistrar_Click` vuelca los controles en la global antes de validar (`:194-212`): si `Registrar` falla, la aplicación activa en memoria queda con valores no guardados.
- La contraseña de la aplicación se muestra en claro (sin máscara, `.form.txt:1201-1214`) y se guarda tal cual en `TbAplicaciones.Pass` (`Aplicacion.cls:1895-1896, 1958-1959`).
- Las etiquetas `Tag` no las usa ningún código (`.form.txt:927-2000`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: alta explícita con ruta propia; identificador generado por la base; restricciones únicas en base de datos para nombre y nombre corto; secreto de la aplicación cifrado y nunca devuelto al cliente.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario no administrador ve los datos sin poder editar ni registrar.
2. Guardar sin cambios devuelve «Ningún cambio que registrar» y no escribe.
3. Guardar con nombre o nombre corto ya usados por otra aplicación devuelve el mensaje de duplicado.
4. Cada obligatorio vacío devuelve su mensaje y no escribe.
5. Editar y vaciar un campo opcional lo deja nulo.
6. Cambiar solo «Icono En Lanzadera» o «Nombre Icono Arbol» cuenta como cambio (corrige el defecto, pendiente de decisión).

## Pendientes

- Si `TbAplicaciones.IDAplicacion` es autonumérico (el ERD no lo indica, `../../ERD/schema.sql:11`).
- Si existe alguna vía real de alta en el legacy, dado el defecto descrito.
- Uso previsto de los `Tag` «PROTEGIDO» y «DATO».
- Campos de `TbAplicaciones` que este formulario no gestiona (`NombreFuncionPublicacion`, `TituloAplicacion`, `URLDIrectorioIconoAplicacion`, `Comando`; `../../ERD/schema.sql:18-32`).
