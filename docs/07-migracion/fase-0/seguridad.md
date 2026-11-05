# Fase 0: Seguridad (`Seguridad`)

Issue #1574. Épica `docs/07-migracion/epica-coexistencia-postgresql.md` (fase 0).
Base de análisis: repositorio `00_SEGURIDAD` (GitHub `DysTelefonica/seguridad`), rama `main`, SHA `e22d85417c1fe11449be78c3da4666a5ed7c7fdb`.
Fuentes: ERD del front y del backend con `scripts/dysflow-erd/extract-erd.ps1` (solo lectura, sobre copia) y trazado del código con CodeGraph y con `read`/`grep` sobre `src/`. CodeGraph sí estaba operativo en este entorno (`codegraph explore` devuelve el fuente actual del disco), así que no hay fallback que declarar.

Las rutas del front son relativas a ese repositorio.

## Base de análisis

| Hecho | Evidencia |
|---|---|
| La rama `release/2025-001.rc1` existe, en SHA `d02684e6e789f6d978765dbbbb5daa7dddb140dc` | `git branch -a` en `00_SEGURIDAD` devuelve `main`, `release/2025-001.rc1` y sus equivalentes en `origin/` |
| La RC1 **no aporta nada** al código de la aplicación: su `src/` y su binario son idénticos a los de `main` | `git diff --stat origin/main origin/release/2025-001.rc1 -- Seguridad.accdb src/` no devuelve ninguna línea. Los 134 ficheros que difieren entre ambas ramas son todos de la propagación de skills (`.agents/skills/`, `.github/ISSUE_TEMPLATE/`, `AGENTS.md`, `.team-skills.yaml`) |
| La RC1 nunca se fusionó en `main`, pero `main` sí se fusionó en la RC1 | `git merge-base --is-ancestor origin/release/2025-001.rc1 origin/main` falla; PR #5 (`baseRefName` `release/2025-001.rc1`, `headRefName` `main`, fusionado el 2026-10-23, «chore: alinear release/2025-001.rc1 con main») es el único PR que toca la RC1, y va en sentido contrario |
| **No hay ninguna evidencia de UAT**, ni pendiente ni superada | El repositorio no tiene `docs/`; `gh issue list --repo DysTelefonica/seguridad --state all` no devuelve ninguna issue; `gh api repos/DysTelefonica/seguridad/tags` no devuelve ninguna etiqueta; `README.md:17` dice «Pendiente de analizar: no hay todavía código exportado, backend identificado ni medición de tamaño» |
| El README está desfasado respecto al repositorio | `README.md:13` dice que el código exportado no se versiona, pero `fd19e49 feat: export VBA source from Seguridad.accdb 2025-001` lo exportó y `src/` está versionado |
| La versión que declara el catálogo es `2025-001` | `README.md:9` («Binario de la aplicación, versión `2025-001`») y `README.md:13` («Copiado de la carpeta de versiones de la aplicación (`Seguridad\Versiones\2025-001`)») |
| Seguridad es una aplicación de «Solo Oficina» | `TbAplicaciones` de `Lanzadera_Datos.accdb` fila `IDAplicacion` 16: `NombreEjecutable` `Seguridad.accdb`, `NombreArchivoDatos` `Seguridad_datos.accdb`, `EjecucionEnOficina` `Sí`, `EnPruebas` `No`. El formulario `Form_FormAplicacionesOficina` de la Lanzadera es el que la ofrece (`src/forms/Form_FormAplicacionesOficina.form.txt:796`, botón `ComandoSeguridad`, `Tag ="16"`) |
| La etiqueta «Ver. Servidor» no se puede verificar desde el repositorio | La Lanzadera la pinta en tiempo de ejecución desde el `.ini` del servidor (`src/classes/Aplicacion.cls:1133`, `LeerIni(Me.URLIniServidor, "VersionAplicacion", …)`), y el `.ini` no está versionado (`README.md:13`). El único valor en el repositorio es la etiqueta de diseño `Ver. Servidor: 0.03.003` (`src/forms/Form_FormAplicacionesOficina.form.txt:1441`), que `setVersiones` sobrescribe al arrancar (`src/modules/Funciones Genreales.bas:2326`). El front lee la misma clave de su propio `.ini` (`src/classes/Entorno.cls:137`, `:142`) |

Consecuencia: como la RC1 y `main` comparten código y binario, el inventario de abajo vale para las dos, y se fija `main` como base de análisis. No hay una UAT pendiente que justifique estudiar la RC1, ni una UAT superada que la descarte: no hay evidencia de UAT en ninguna parte.

## Estructura del front

| Hecho | Valor |
|---|---|
| Tablas del front | 6, todas propias, 0 vinculadas |
| Tablas propias del front | `Tb0AuxOrdenarColeccion`, `Tb0FiltroRegistros`, `Tb0FiltroRegistrosConRelaciones`, `Tb0RegistrosParaRelacionar`, `Tb0RelacionesDeUnRegistro`, `TbAuxURLNoAlcanzables` |
| Vínculos del front | Ninguno: el ERD del front no declara ninguna tabla vinculada y el volcado directo de `TableDefs` del binario devuelve 19 definiciones, de las que 13 son `MSys*` y las 6 restantes son propias |
| Formularios | 54 (ERD y `src/forms/`: 54 `.form.txt` y 54 `.cls`) |
| Consultas guardadas del front | 4: `Consulta1` (tipo 64 en el binario, con SQL `INSERT INTO Tb0FiltroRegistros …`), `Consulta2` (tipo 0, selección), `Consulta3` (tipo 32, con SQL `DELETE … FROM Tb0TipoRegistrosRelacionados`) y `Consulta4` (tipo 0, selección). En `Consulta1` y `Consulta3` el tipo declarado y el SQL guardado no coinciden. El SQL está en `src/queries/Consulta1.sql` … `Consulta4.sql` |
| Módulos y clases | 5 módulos (`Constructor.bas`, `Funciones Generales.bas`, `JsonConverter.bas`, `Módulo1.bas`, `Variables de Entorno.bas`) y 9 clases (`Adjunto`, `Correo`, `Entorno`, `Log`, `Registro`, `RegistroBusqueda`, `TipoRegistro`, `Usuario`, `UsuarioAplicacionPermisos`) |
| Macros o `AutoExec` | Ninguno: `AutoExec`, `AutoOpen`, `AutoClose` y `Sub Auto` no aparecen en ningún fichero de `src/` |
| Controles con `RowSourceType` de tabla o consulta | 0: los 116 `RowSourceType` del front son `"Value List"` |
| Controles con `ControlSource` con valor | 33, todos en los cinco formularios de filtro y todos con un valor de columna |

La propiedad `RowSourceTypeInt =1` que acompaña a cada combo y lista del export es un artefacto de la exportación, no un indicador de origen en tabla: en el export de Brass ese mismo atributo aparece exactamente 79 veces, tantas como `RowSourceType ="Value List"`, y ninguna vez en los combos que sí llevan SQL.

### Una regla del propietario que la medición refuta

La regla «salvo Brass, los formularios son desconectados» **no se cumple en Seguridad**: cinco formularios están unidos a un recordset, con `RecordSource` a nivel de formulario.

| Formulario | `RecordSource` | Tabla |
|---|---|---|
| `Form_FormFiltro` | `src/forms/Form_FormFiltro.form.txt:51` | `Tb0FiltroRegistros` |
| `Form_FormFiltroAvanzado` | `src/forms/Form_FormFiltroAvanzado.form.txt:46` | `Tb0FiltroRegistros` |
| `Form_FormFiltroParaRelacion` | `src/forms/Form_FormFiltroParaRelacion.form.txt:47` | `Tb0RegistrosParaRelacionar` |
| `Form_FormFiltroRegistrosConRelaciones` | `src/forms/Form_FormFiltroRegistrosConRelaciones.form.txt:55` | `Tb0FiltroRegistrosConRelaciones` |
| `Form_FormFiltroRegistrosRelacionados` | `src/forms/Form_FormFiltroRegistrosRelacionados.form.txt:55` | `Tb0RelacionesDeUnRegistro` |

La mitigación es que las cinco tablas son **propias del front**, no del backend: son tablas de trabajo de filtro y de relación. La forma es la de HPS (formulario unido a la caché local) y no la de Brass (formulario unido al backend), pero el efecto para la migración es el mismo: cuando el front desaparezca, las cinco tablas tienen que mudarse al backend o desaparecer con él.

## Recordsets de tipo tabla

El indicio se confirma con evidencia: hay **18 aperturas de recordset cuyo argumento es el nombre de una tabla y no una sentencia `SELECT`**. En DAO, `OpenRecordset` con un nombre de tabla abre un recordset de tipo tabla (`dbOpenTable`), que no existe en PostgreSQL y que además es de solo lectura frente a cambios concurrentes.

| Sitio | Tabla | Base de datos |
|---|---|---|
| `src/modules/Constructor.bas:111` (asignación en `:109`) | `TbUsuariosAplicaciones` | Lanzadera |
| `src/modules/Constructor.bas:214` (`:213`) | `Tb0TipoRegistros` | Backend propio |
| `src/modules/Constructor.bas:264` (`:263`) | `Tb0HerramientaDocAyuda` | Backend propio |
| `src/modules/Constructor.bas:568` (`:559`) | `Tb0TipoRegistros` | Backend propio |
| `src/modules/Constructor.bas:832` (`:831`) | `Tb0Adjuntos` | Backend propio |
| `src/modules/Constructor.bas:941` (`:940`) | `Tb0Adjuntos` | Backend propio |
| `src/classes/Registro.cls:683` (`:682`) | `TipoRegistro.NombreTabla` (dinámica) | Backend propio |
| `src/classes/Registro.cls:1344` (`:1343`) | `m_NombreTabla` (dinámica) | Backend propio |
| `src/classes/Registro.cls:1537` (`:1536`) | `Tb0RegistrosRelacionados` | Backend propio |
| `src/modules/Módulo1.bas:1046` (`:1045`) | `TipoRegistroActivo.NombreTabla` (dinámica) | Backend propio |
| `src/modules/Funciones Generales.bas:137` (`:136`) | `TbCorreosEnviados` | Correos |
| `src/classes/Correo.cls:150` (`:149`) | `TbCorreosEnviados` | Correos |
| `src/classes/Log.cls:77` (`:76`) | `TbLog` | Front (error, ver más abajo) |
| `src/modules/Funciones Generales.bas:921` (`:920`) | `Tb0FiltroRegistros` | Front |
| `src/modules/Funciones Generales.bas:968` (`:967`) | `Tb0RegistrosParaRelacionar` | Front |
| `src/modules/Funciones Generales.bas:1021` (`:1020`) | `Tb0RelacionesDeUnRegistro` | Front |
| `src/modules/Funciones Generales.bas:1103` (`:1102`) | `Tb0FiltroRegistrosConRelaciones` | Front |

Tres de ellos son el patrón más difícil: el nombre de la tabla no está en el código, sino en el catálogo `Tb0TipoRegistros.NombreTabla`, que el front lee del backend y concatena en el SQL (`src/classes/Registro.cls:330`, `:682`, `:1343`, `:1746`; `src/modules/Constructor.bas:381`, `:1069`, `:1196`, `:1307`; `src/modules/Módulo1.bas:1045`). El esquema de la aplicación no está escrito en el código: está escrito en datos.

Este mismo patrón ya estaba registrado para Seguridad en la ficha de aperturas directas de la Lanzadera (`docs/07-migracion/fase-0/aperturas-directas-lanzadera.md:134`), que cita `Constructor.bas:111` como ejemplo de recordset de tipo tabla que no sobrevive a una conexión ODBC.

## Aperturas directas

| Apertura | Código | Observación |
|---|---|---|
| Backend propio | `src/modules/Variables de Entorno.bas:484` (`getdb()`) | Ruta `\\datoste\aplicaciones_dys\Aplicaciones PpD\Seguridad\Seguridad_datos.accdb`, construida en `:475-478`. La credencial va literal en el código: `MS Access;PWD=***` |
| Lanzadera | `src/modules/Variables de Entorno.bas:422` (`getdbLanzadera()`) | Ruta `…\0Lanzadera\Lanzadera_Datos.accdb` (`:419-420`). Credencial literal `PWD=***`. Consumida en `Constructor.bas:72`, `:111`, `:167`, en `Usuario.cls:282` y en `UsuarioAplicacionPermisos.cls:263` |
| Correos | `src/modules/Variables de Entorno.bas:457` (`getdbCorreo()`) | Ruta `…\00Recursos\Correos_datos.accdb` (`:454-455`). Credencial literal `PWD=***`. Consumida en `Funciones Generales.bas:137` y `Correo.cls:150` |
| `DoCmd.OpenTable` o `TransferDatabase` | Ninguno | No hay ninguna apertura de tabla por interfaz |
| `CurrentDb()` | `src/classes/Log.cls:77`, `src/modules/Funciones Generales.bas:915`, `:962`, `:1012`, `:1090`, `:1092`, `:921`, `:968`, `:1021`, `:1103` | Todas apuntan a tablas del propio front, menos una: `Log.cls:77` |

Los tres `OpenDatabase` llevan la contraseña en el literal de la cadena de conexión (`Variables de Entorno.bas:422`, `:457`, `:484`). La credencial es la misma que abre los demás binarios del ecosistema; aquí va enmascarada.

`Usuario.cls:282` y `UsuarioAplicacionPermisos.cls:263` leen metadatos de otra base con `getdbLanzadera().TableDefs("TbUsuariosAplicaciones")` y `TableDefs("TbUsuariosAplicacionesPermisos")`. El acceso por `TableDefs` tampoco tiene equivalente en PostgreSQL.

### Un defecto medido en `Log.cls`

`src/classes/Log.cls:76-77` abre `TbLog` contra `CurrentDb()`, es decir, contra el front. `TbLog` **no está en el front**: el ERD del front declara 6 tablas y ninguna es `TbLog`; la tabla vive en el backend y tiene 2260 registros. Además, el identificador del log sí se calcula contra el backend (`src/classes/Log.cls:35` → `Constructor.getID("TbLog", "IDLog", …)` → `src/modules/Constructor.bas:764`, que usa `getdb()`).

El defecto no se manifiesta porque `RegistrarLog` (`src/modules/Funciones Generales.bas:746`) no tiene ningún llamador en `src/`: la ruta de alta del log es código muerto.

## Tablas de otros backends que consume el front

| Tabla | Backend dueño | Evidencia |
|---|---|---|
| `TbUsuariosAplicaciones` | Lanzadera | `src/modules/Constructor.bas:45-46` (SQL), `:109-111` (recordset de tipo tabla), `src/classes/Usuario.cls:282` (`TableDefs`) |
| `TbUsuariosAplicacionesPermisos` | Lanzadera | `src/modules/Constructor.bas:164-166` (SQL), `src/classes/UsuarioAplicacionPermisos.cls:263` (`TableDefs`) |
| `TbCorreosEnviados` | Correos | `src/modules/Funciones Generales.bas:136-137`, `src/classes/Correo.cls:149-150` |

Las tres están bien ubicadas: son la autoridad de identidad, de permisos y de correo. Ninguna de las dos aplicaciones a las que pertenecen está en el alcance de la migración por ahora, así que el frente que hay que resolver en la fase 2 es doble: la tabla y el modo de acceso.

## Backend `Seguridad_datos`

| Hecho | Valor |
|---|---|
| Tablas | 36 propias, 0 vinculadas |
| Relaciones | 4: 3 del usuario (`Tb0AdjuntosTb0AdjuntosIndexaciones`, `Tb0TipoRegistrosTb0RegistrosPendientes`, `Tb0TipoRegistrosTb0TipoRegistrosDescripcion`) y 1 de sistema (`MSysNavPaneGroupCategoriesMSysNavPaneGroups`) |
| Consultas guardadas | 1: `Consulta1`, con SQL `SELECT TbENSServiciosTelefonica.AsuntoENS … WHERE (((TbENSServiciosTelefonica.AsuntoENS) Like '*' Or (TbENSServiciosTelefonica.AsuntoENS) Is Null));` |
| Formularios | 0 |
| Tablas sin clave primaria | 2: `Tb0Vinculo` (1 campo `ID`, 0 registros) y `TbNombramientoJS` (17 registros) |
| Campos multivalor, de datos adjuntos u objeto OLE | Ninguno: ningún campo de las 36 tablas tiene `IsComplex` y el único tipo largo (`dbLongText`, valor 12) se ha verificado con contenido real, no por el tipo. `Tb0AdjuntosIndexaciones.TextoAdjunto` alcanza 9474 caracteres, así que es texto largo, no OLE. Los adjuntos no se guardan en la base: `Tb0Adjuntos` guarda solo el nombre del fichero (`src/classes/Adjunto.cls:204-213`) y el binario vive en el sistema de ficheros |
| Tablas con 0 registros | 9: `Tb0TipoRegistroUso`, `Tb0Vinculo`, `TbAuditorias`, `TbBuzonTEOficio`, `TbCarta`, `TbCertificados`, `TbInformes`, `TbInspeccionesAuditoriasSeguridad`, `TbRecibosMP` |

### El catálogo de tipos de registro

El backend tiene una tabla de catálogo, `Tb0TipoRegistros` (21 registros), que dice para cada tipo de registro cuál es su tabla física, cuál es su clave y qué campo es el asunto, las notas y la fecha. Se lee en `src/modules/Constructor.bas:213-214` y `:559-568`.

Las 21 tablas que nombra, con su clave:

| ID | Tipo | Tabla | Clave |
|---|---|---|---|
| 1 | Solicitud de integración | `TbMatClasificadoSIntegracion` | `IDSMatClasificadoSIntegracion` |
| 2 | Recibo de entrega y recepción (MP) | `TbMatClasificadoReciboEntrega` | `IDMatClasificadoReciboEntrega` |
| 3 | Recibo de transporte de material clasificado | `TbMatClasificadoReciboTransporte` | `IDMatClasificadoReciboTransporte` |
| 4 | Solicitud documentación clasificada | `TbSolDocClas` | `IDSolDocClas` |
| 5 | BUZÓN TE_Seguridad_Industrial | `TbBuzonTE` | `IDBuzonTE` |
| 6 | Comunicación nombramiento Jefe de Seguridad | `TbNombramientoJS` | `IDNombramientoJS` |
| 7 | HSEM / HSES | `TbHSEMHSES` | `IDHSEMHSES` |
| 8 | Seguridad de edificios | `TbSeguridadEdificio` | `IDSeguridadEdificio` |
| 9 | Informes de seguimiento | `TbInformes` | `IDInforme` |
| 10 | Acreditación del Sistema | `TbAcreditacionSistema` | `IDAcreditacionSistema` |
| 11 | Ficha del contratista | `TbFichaContratista` | `IDFichaContratista` |
| 12 | Inventario de documentación clasificada | `TbInventarioDocumentacionClasificada` | `IDInventarioDocumentacionClasificada` |
| 13 | Correspondencia | `TbCorrespondencia` | `IDCorrespondencia` |
| 15 | Formación | `TbFormacion` | `IDFormacion` |
| 16 | Cripto/Cifra | `TbCripto` | `IDCripto` |
| 17 | Inspecciones / Auditorías de seguridad | `TbInspeccionesAuditoriasSeguridad` | `IDInspeccionesAuditoriasSeguridad` |
| 18 | Certificados | `TbCertificados` | `IDCertificado` |
| 19 | Poderes seguridad | `TbPoderesSeguridad` | `IDPoderesSeguridad` |
| 20 | Planes y Procedimientos de Seguridad | `TbPlanesSeguridad` | `IDPlanSeguridad` |
| 21 | ENS Servicios de Telefónica | `TbENSServiciosTelefonica` | `IDENS` |
| 22 | Normativas | `TbNormativas` | `IDNormativa` |

El identificador 14 no existe en el catálogo: hay un hueco en la numeración entre `Correspondencia` (13) y `Formación` (15). Ninguna tabla lo reclama.

Las 21 tablas comparten seis columnas (`FechaRegistro`, `UsuarioRegistra`, `JuridicaTdE`, `JuridicaTME`, `JuridicaTSOL`, `JuridicaTIS`) y añaden entre 4 y 13 columnas propias de su tipo. El código no las conoce: las alcanza por el nombre que guarda el catálogo.

## Patrones de riesgo del VBA

| Patrón | Medida | Dónde | Qué hay que cambiar |
|---|---|---|---|
| Generadores `Max + 1` | 2 | `src/modules/Constructor.bas:762-774` (`getID`: `SELECT Max(<ID>) AS Maximo FROM <tabla>` y `CStr(m_IDMax + 1)`) y `src/modules/Funciones Generales.bas:178-191` (`DameID`: `SELECT Max(<tabla>.<campo>) AS MaxID FROM <tabla>` y `CStr(lngIDMax + 1)`) | Identidad o autonumérico. `getID` se llama desde `src/classes/Registro.cls:62` para las 21 tablas de tipo de registro, y `DameID` desde `src/modules/Funciones Generales.bas:134` para `TbCorreosEnviados` |
| `Like '*'` | 7 en el front | `src/modules/Constructor.bas:345`, `:350`, `:355`, `:360`, `:365`, `:370` (filtros de asunto, notas y las cuatro marcas de jurisdicción) y `src/modules/Módulo1.bas:77` | `ILIKE '%'` o, mejor, quitar el filtro: `Like '*'` no filtra nada |
| Literales de fecha con `#` | 2 | `src/modules/Constructor.bas:341` (`Between #<fecha># And #<fecha>#`, con las fechas formateadas en `:334` y `:339`) y `src/queries/Consulta1.sql:2` (`#1/29/2019#`) | `DATE '…'` o `TIMESTAMP '…'` |
| `Date()` o `Now()` dentro de SQL | 0 | — | No hay ninguno. `Now()` solo aparece en asignación de campos de VBA: `src/classes/Adjunto.cls:186`, `src/classes/Correo.cls:167`, `src/classes/Registro.cls:1715`, `src/modules/Funciones Generales.bas:146`, `:789` |
| `Seek` | 0 | — | No hay ninguno |
| `dbOpenTable` explícito | 0 | — | El tipo de tabla no se pide: se hereda del hecho de pasar un nombre de tabla a `OpenRecordset`. Los 18 sites están en el apartado «Recordsets de tipo tabla» |
| Recordsets sin `dbSeeChanges` | 43 de 43 | Ninguna de las 43 llamadas a `OpenRecordset` del `src/` pasa un segundo argumento, así que ninguna fija opciones de apertura | Decidir el nivel de aislamiento y reescribir el patrón: en PostgreSQL el cursor no es un recordset de tipo tabla |
| `Option Compare Database` | 67 de 68 | Todos los módulos, clases y code-behind de formulario salvo `src/modules/JsonConverter.bas` | Fija la semántica de comparación de cadenas al estilo de Access. En PostgreSQL hay que decidir la collation y reescribir las comparaciones |
| `DELETE * FROM` | 13 | `src/classes/Adjunto.cls:253`, `src/classes/Registro.cls:860`, `:1577`, `:1603`, `src/forms/Form_Form0RegistrosBusqueda.cls:972`, `src/forms/Form_Form0RegistrosParaRelacion.cls:375`, `src/modules/Funciones Generales.bas:914`, `:961`, `:1011`, `:1089`, `:1091`, `src/modules/Módulo1.bas:1031`, `:1104` | `DELETE FROM <tabla>` |
| `Is Null` | 7 | `src/modules/Constructor.bas:345`, `:350`, `:355`, `:360`, `:365`, `:370` y `src/modules/Módulo1.bas:77` | `IS NULL` |
| SQL con valores concatenados | 4 familias | `src/modules/Constructor.bas:341` (fechas del formulario), `:346` (texto del filtro de asunto), `:517` (`NombreTabla`), `src/classes/Registro.cls:331` y `:1748` (identificador del registro), `src/modules/Constructor.bas:167` (correo del usuario) | Parametrizar. El texto del filtro de pantalla llega concatenado dentro del literal de SQL |

Una peculiaridad más: `src/classes/Registro.cls:1344` y `:1351` abren el recordset dos veces en la rama de alta (la primera en el `If`, la segunda después del `End If`), y la segunda descarta la primera. Es una fuga de recordset, no un problema de PostgreSQL, pero conviene resolverla antes del corte.

## Veredicto de propiedad

**Se quedan en el backend**: las 26 tablas que el VBA del front necesita.

- Las **21 tablas de tipo de registro** que nombra `Tb0TipoRegistros.NombreTabla` (medido: 21 filas) y a las que el código llega por SQL construido con ese nombre (`src/classes/Registro.cls:330`, `:682`, `:1343`, `:1746`; `src/modules/Constructor.bas:381`, `:1069`, `:1196`, `:1307`; `src/modules/Módulo1.bas:1045`).
- Las **5 tablas de infraestructura** citadas por nombre: `Tb0TipoRegistros` (`src/modules/Constructor.bas:213-214`, `:559-568`), `Tb0RegistrosRelacionados` (`src/classes/Registro.cls:1536-1537`, `:1577`, `:1603`; `src/modules/Constructor.bas:1308`), `Tb0Adjuntos` (`src/classes/Adjunto.cls:204-205`, `:253`; `src/modules/Constructor.bas:831`, `:940`) y `Tb0HerramientaDocAyuda` (`src/modules/Constructor.bas:263-264`).

**No las usa el front**: 10 tablas del backend, con cero referencias en `src/`: `Tb0AdjuntosIndexaciones`, `Tb0Conexiones`, `Tb0RegistrosPendientes`, `Tb0TipoRegistrosDescripcion`, `Tb0TipoRegistroUso`, `Tb0Vinculo`, `TbAuditorias`, `TbBuzonTEOficio`, `TbCarta` y `TbRecibosMP`. Seis de ellas están vacías. Dos están atadas al catálogo por una relación declarada (`Tb0RegistrosPendientes` y `Tb0TipoRegistrosDescripcion` dependen de `Tb0TipoRegistros`), y una por una relación con `Tb0Adjuntos` (`Tb0AdjuntosIndexaciones`, que guarda el texto indexado de un adjunto y que tiene 1 registro). Requieren comprobar otros consumidores antes de proponer nada (tarea T6).

**Las usa el front pero no existen**: `Tb0TipoRegistrosRelacionados`, que aparece en las dos consultas guardadas `Consulta3` (`src/queries/Consulta3.sql:1-3`) y `Consulta4` (`src/queries/Consulta4.sql:1-3`) del front, y no está ni en el front (6 tablas) ni en el backend (36 tablas). Es una referencia muerta: esas consultas fallarían en runtime.

**Están mal ubicadas**: las 4 tablas de trabajo del front que alimentan los cinco formularios unidos a recordset (`Tb0FiltroRegistros`, `Tb0FiltroRegistrosConRelaciones`, `Tb0RegistrosParaRelacionar`, `Tb0RelacionesDeUnRegistro`). Incumplen la regla de que las tablas del aplicativo viven en su backend, y sin ellas los formularios de filtro y de relación no tienen origen de datos. Son candidatas a mudarse al backend.

**No las usa nadie, según el código**: `Tb0AuxOrdenarColeccion` (3 registros) y `TbAuxURLNoAlcanzables` (100 registros), las dos del front, sin ninguna referencia en `src/`.

**Las usa el código, pero desde la base equivocada**: `TbLog`. El código la abre contra `CurrentDb()`, es decir, contra el front (`src/classes/Log.cls:76-77`), y la tabla vive en el backend. La ruta está muerta porque `RegistrarLog` no tiene llamadores.

## Propuestas para después del descubrimiento global

Ninguna se aplica: la puerta de descubrimiento global sigue cerrada (#1576).

| Id | Propuesta | Grupo (fase 1) | Comprobación previa | Reversión |
|---|---|---|---|---|
| G1 | Declarar clave primaria en `TbNombramientoJS` (17 registros, sin clave) y en `Tb0Vinculo` (0 registros, sin clave). | (a) Imprescindible | Contar duplicados de `IDNombramientoJS` antes de fijar la clave; `Tb0Vinculo` solo tiene una columna `ID` y ningún índice: decidir antes si se retira (G2) en lugar de declararle clave. | Quitar la clave primaria. |
| G2 | Poner en cuarentena las 10 tablas del backend que el front no referencia, renombrándolas a `zzBAJA_<tabla>`. Antes hay que quitar las tres relaciones que las atan: `Tb0AdjuntosTb0AdjuntosIndexaciones`, `Tb0TipoRegistrosTb0RegistrosPendientes` y `Tb0TipoRegistrosTb0TipoRegistrosDescripcion`. | (a) Imprescindible | Confirmar en el grafo global que ningún otro front ni la web las leen (tarea T6) y que no hay consumidores fuera de Access. Para `Tb0AdjuntosIndexaciones` hay que decidir qué pasa con el texto indexado de los adjuntos, porque guarda contenido que el front no lee pero que alguien puede haber previsto para la web. | Quitar el prefijo `zzBAJA_` y volver a declarar las relaciones. |
| G3 | Resolver las dos referencias muertas: `Tb0TipoRegistrosRelacionados`, que usan las consultas guardadas `Consulta3` y `Consulta4` del front y no existe en ninguna base; y `TbLog`, que `src/classes/Log.cls:76-77` abre contra el front cuando vive en el backend. | (a) Imprescindible | Reproducir en runtime el alta y el borrado de un registro, el alta de una relación entre dos registros y la ruta de log (`RegistrarLog`, `src/modules/Funciones Generales.bas:746`). Si la ruta de log sigue sin llamadores, la decisión es borrar el código muerto en lugar de crear la tabla. | Restaurar la tabla en la base que le corresponda o restituir la referencia. |
| G4 | Sustituir los dos generadores `Max + 1` (`src/modules/Constructor.bas:762-774` y `src/modules/Funciones Generales.bas:178-191`) por identidad o autonumérico. | (a) Imprescindible | Enumerar las tablas cuyo identificador se calcula así: las 21 de tipo de registro (`src/classes/Registro.cls:62`), `Tb0TipoRegistros`, `Tb0Adjuntos`, `Tb0RegistrosRelacionados`, `TbLog` y `TbCorreosEnviados` (esta última en otro backend). Verificar que el identificador se empieza a calcular también en el alta, no solo en la edición. | Volver a la consulta `SELECT Max(…)`. |
| G5 | Declarar la integridad referencial que el ERD no recoge: `Tb0RegistrosRelacionados.IDTipoRegistro` se une en SQL con `Tb0TipoRegistros.IDTipoRegistro` y las 21 tablas de tipo de registro se unen con su catálogo por la columna de texto `NombreID`, que no es clave foránea. | (a) Imprescindible | Contar los huérfanos en los dos sentidos antes de declarar las claves foráneas, y comprobar que el tipo de `Tb0TipoRegistros.IDTipoRegistro` coincide con el de la columna referenciada. | Deshacer las relaciones y volver a la unión en SQL. |
| G6 | Mudar al backend las 4 tablas de trabajo del front que alimentan los cinco formularios unidos a recordset: `Tb0FiltroRegistros`, `Tb0FiltroRegistrosConRelaciones`, `Tb0RegistrosParaRelacionar` y `Tb0RelacionesDeUnRegistro`. Sin ellas, los formularios de filtro y de relación se quedan sin origen de datos en cuanto el front desaparezca. | (a) Imprescindible | Comprobar que el backend puede alojar el mismo esquema y decidir si son tablas de sesión por usuario o compartidas: `Tb0FiltroRegistros` se borra y se repuebla por sesión (`src/modules/Funciones Generales.bas:914-921`), lo que sugiere que son de sesión, no de aplicación. | Devolver las tablas al front. |
| G7 | Retirar la credencial literal de las tres cadenas de conexión (`src/modules/Variables de Entorno.bas:422`, `:457`, `:484`) y leerla de la configuración. | (a) Imprescindible | Es la misma credencial que abre todos los binarios del ecosistema. Comprobar que ninguna otra copia del literal queda en el front, y que el cambio no rompe la apertura del backend, de la Lanzadera ni de Correos. | Restaurar las tres líneas originales. |
| G8 | Reescribir el SQL que no compila en PostgreSQL: 13 `DELETE * FROM`, 7 `Is Null`, 7 `Like '*'`, 2 literales de fecha con `#` y el `Between #…#` de `src/modules/Constructor.bas:341`. | (a) Imprescindible | Ejecutar en runtime los cuatro caminos que borran las tablas de filtro (`src/modules/Funciones Generales.bas:914`, `:961`, `:1011`, `:1089`) y la búsqueda con los cuatro filtros de jurisdicción. Comprobar que `Like '*'` se traduce por un filtro vacío y no por un `ILIKE '%'`, que en tablas grandes no es lo mismo. | Volver al SQL original de Access. |
| G9 | Parametrizar el SQL que concatena valores: el texto del filtro de asunto y de notas (`src/modules/Constructor.bas:343`, `:348`), las fechas del filtro (`:341`), el nombre de tabla del catálogo (`:517`), el identificador del registro (`src/classes/Registro.cls:331`, `:1748`) y el correo del usuario (`src/modules/Constructor.bas:164-167`). | (a) Imprescindible | El texto del filtro llega desde un control del formulario: es entrada de usuario. Comprobar en runtime que los filtros siguen devolviendo lo mismo con comillas y signos de porcentaje en el texto. | Volver a la concatenación. |
| G10 | Retirar o reubicar `Tb0AuxOrdenarColeccion` (3 registros) y `TbAuxURLNoAlcanzables` (100 registros): ninguna de las dos tiene una referencia en `src/`. | (b) Recomendable si es barato | Comprobar en el binario que no hay ningún formulario, consulta ni macro que las use, y que la lista de anexos inalcanzables se reconstruye en memoria (`src/modules/Constructor.bas:928-960` recorre `Tb0Adjuntos` y comprueba el fichero en disco, sin escribir en ninguna tabla). | Restaurar las tablas desde la copia de seguridad. |
| G11 | Decidir el destino del patrón «una tabla por tipo de registro»: 21 tablas con la misma forma básica, elegidas por un catálogo en datos y accedidas por SQL concatenado. En PostgreSQL esto es una tabla única con columna discriminadora, o 21 tablas con SQL dinámico; la decisión cambia el número de tablas y el SQL de todo el front. | (c) Aplazar a la refactorización final | Levantar el catálogo completo de columnas de las 21 tablas y medir cuántas comparten además de las seis columnas comunes, para acotar la opción de tabla única. El coste está medido en el catálogo de SQL por tabla de `docs/07-migracion/fase-0/controles-conectados.md`, que hay que extender con Seguridad. | La decisión no cambia el legacy: es un cambio de diseño para la fase 1. |
| G12 | Fijar la semántica de comparación de cadenas: `Option Compare Database` está en 67 de los 68 módulos, clases y code-behind de formulario (falta en `src/modules/JsonConverter.bas`). | (c) Aplazar a la refactorización final | Enumerar las comparaciones de cadenas que dependen de la regla de Access (acentos, mayúsculas y ordenación) y decidir la collation de PostgreSQL antes de la reescritura. | Restablecer el valor de la opción en el proyecto de Access. |

Lo que **no** se toca ahora: las 21 tablas de tipo de registro, las 5 tablas de infraestructura del backend, las 4 tablas de trabajo del front (más allá de la propuesta G6), los vínculos de otros backends que el front no usa (no tiene ninguno) y la ruta `\\datoste\…` de producción, salvo la propuesta G7 sobre la credencial.

## Anexo: SQL literal de los cinco formularios unidos a recordset

Ningún control tiene SQL como origen: los 116 `RowSourceType` del front son `"Value List"`, y las cinco asignaciones en tiempo de ejecución (`src/forms/Form_Form0RegistroConSusRelaciones.cls:20`, `src/forms/Form_Form0RegistrosParaRelacion.cls:633`, `src/forms/Form_Form0TipoRegistroEleccion.cls:266`, `src/modules/Módulo1.bas:73`, `:784`) tampoco contienen SQL: tres rellenan el control con listas de valores literales y dos lo vacían (`cmb.RowSource = ""`). El SQL de la aplicación se construye en código, en variables, y es el que hay que reescribir.

| SQL literal | Uso | Tabla |
|---|---|---|
| `SELECT Tb0FiltroRegistros.IDRegistro, Tb0FiltroRegistros.IDTipoRegistro, Tb0FiltroRegistros.AliasTipoRegistro AS Tipo, Tb0FiltroRegistros.Asunto, Tb0FiltroRegistros.Fecha FROM Tb0FiltroRegistros; ` | `Form_FormFiltro` (`src/forms/Form_FormFiltro.form.txt:51`) y `Form_FormFiltroAvanzado` (`src/forms/Form_FormFiltroAvanzado.form.txt:46`) | `Tb0FiltroRegistros` |
| `SELECT Tb0RegistrosParaRelacionar.IDRegistro, Tb0RegistrosParaRelacionar.IDTipoRegistro, Tb0RegistrosParaRelacionar.AliasTipoRegistro AS Tipo, Tb0RegistrosParaRelacionar.Asunto, Tb0RegistrosParaRelacionar.Fecha FROM Tb0RegistrosParaRelacionar; ` | `Form_FormFiltroParaRelacion` (`src/forms/Form_FormFiltroParaRelacion.form.txt:47`) | `Tb0RegistrosParaRelacionar` |
| `SELECT Tb0FiltroRegistrosConRelaciones.IDRegistro, Tb0FiltroRegistrosConRelaciones.IDTipoRegistro, Tb0FiltroRegistrosConRelaciones.AliasTipoRegistro AS Tipo, Tb0FiltroRegistrosConRelaciones.Asunto, Tb0FiltroRegistrosConRelaciones.Fecha, Tb0FiltroRegistrosConRelaciones.JuridicaTdE AS TdE, Tb0FiltroRegistrosConRelaciones…` | `Form_FormFiltroRegistrosConRelaciones` (`src/forms/Form_FormFiltroRegistrosConRelaciones.form.txt:55`) | `Tb0FiltroRegistrosConRelaciones` |
| `SELECT Tb0RelacionesDeUnRegistro.IDRegistro, Tb0RelacionesDeUnRegistro.IDTipoRegistro, Tb0RelacionesDeUnRegistro.AliasTipoRegistro AS Tipo, Tb0RelacionesDeUnRegistro.Asunto, Tb0RelacionesDeUnRegistro.Fecha, Tb0RelacionesDeUnRegistro.JuridicaTdE AS TdE, Tb0RelacionesDeUnRegistro.JuridicaTME AS TME, Tb0RelacionesDeUnRegistro…` | `Form_FormFiltroRegistrosRelacionados` (`src/forms/Form_FormFiltroRegistrosRelacionados.form.txt:55`) | `Tb0RelacionesDeUnRegistro` |

Las dos últimas sentencias están truncadas en el export: la propiedad `RecordSource` se corta en el límite del export y el resto del SQL no está en `src/`. Es el mismo caso que las divergencias de export señaladas en `docs/07-migracion/fase-0/controles-conectados.md`: el SQL completo hay que sacarlo de una sesión de Access.

## Lo que no se pudo verificar

1. **El estado de UAT.** No hay ninguna evidencia: ni en el repositorio (no hay `docs/`), ni en el tracker (la lista de issues está vacía), ni en las etiquetas de git (no hay ninguna). La ficha se apoya en la ausencia de evidencia, no en un documento de UAT. La etiqueta «Ver. Servidor» que muestra la Lanzadera no se puede verificar desde aquí porque el `.ini` del servidor no está versionado.
2. **El contenido completo de dos `RecordSource`.** El export los trunca; falta el resto del SQL de `Form_FormFiltroRegistrosConRelaciones` y de `Form_FormFiltroRegistrosRelacionados`.
3. **Si el SQL exportado coincide con el binario.** No se ha ejecutado `verify_code` de dysflow sobre los 54 formularios de Seguridad. La comparación con el binario se ha hecho solo para las propiedades que el ERD y el volcado directo de `TableDefs` dan, no para el SQL de cada control.
4. **Los consumidores de las 10 tablas no usadas.** Exige el grafo de vínculos de todos los backends (tarea T6) y saber si la web o un proceso programado las leen.
5. **La tabla `Tb0TipoRegistrosRelacionados`.** Se ha comprobado que no existe en ninguna de las dos bases, pero no cuándo dejó de existir ni si el código que la usaba sigue vivo en el binario.
6. **El comportamiento en runtime.** No se ha ejecutado la aplicación: la ficha se apoya en la lectura del código exportado y en el volcado del esquema. En particular, no se ha comprobado si `Log.Alta` falla de verdad en runtime ni si la doble apertura del recordset de `Registro.cls:1344` y `:1351` produce síntomas observables.
