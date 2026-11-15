# Fase 0: Registros (`Registro_Ent_Salida`)

Épica `docs/07-migracion/epica-coexistencia-postgresql.md` (fase 0, #1545). Ficha gemela: [`brass.md`](brass.md) (#1574).
Base de análisis: repositorio `C:\00repos\codigo\00_REGISTROS`, rama `main`, SHA `44f36d3774a4870e7f70d8c8aa71710965f61411` (tag `2025-001`).
Backends medidos: `C:\00repos\datos\Registro_Ent_Salida_Datos.accdb` (backend propio) y `C:\00repos\datos\Registro_Datos.accdb` (backend legacy). Son copias de desarrollo, no producción (`\\datoste\…`): los recuentos son orientativos y hay que repetirlos sobre producción antes de cualquier acción que borre datos.
Fuentes: ERD del front (`front-registro/erd.json`), volcado de esquema del backend (`docs/ERD-auxiliar/Registro_Ent_Salida_Datos.sql`), lectura de solo lectura con dysflow sobre las copias locales, matriz de vínculos `TbTablasAVincular` de `Lanzadera_Datos.accdb` y trazado del código con CodeGraph.

CodeGraph se usó por CLI sobre el índice propio del repositorio (`00_REGISTROS/.codegraph/`, `codegraph status` → 34 ficheros, 2029 nodos). El servidor MCP de CodeGraph solo expone el espacio de trabajo de este repositorio y no el de `00_REGISTROS`, así que el rastreo se completó con `read`/`grep` sobre `src/` y con las consultas de solo lectura de dysflow.

## Base de análisis

| Hecho | Evidencia |
|---|---|
| La rama base es `main`, SHA `44f36d3774a4870e7f70d8c8aa71710965f61411` | `git branch -a` en `00_REGISTROS` devuelve solo `main` y `origin/main`; `git log` → `44f36d3 feat: export VBA source from Registro_Ent_Salida.accdb 2025-001` |
| No hay rama RC1 ni ciclo de release en curso | No existe ninguna rama `rc1` ni etiqueta posterior a `2025-001`; el README solo declara la versión (`README.md:3`: «**Versión**: 2025-001») |
| No hay evidencia de UAT | El repositorio no tiene `docs/`, ni parte de resultados, ni firma; tampoco hay rama de trabajo con UAT pendiente. La base real de este inventario es `main` del SHA indicado |
| `src/` es una copia derivada del binario, no la fuente de verdad | `README.md:5` («Es una copia derivada del binario; la fuente de verdad sigue siendo el `.accdb`»); `src/export-manifest.json:3` (`generatedAt` 2026-10-24) |
| El ERD de backend entregado es de **otro** archivo | `backend-registro/erd.json` describe 5 tablas (`DClases`, `Entradas`, `Nombres`, `Salidas`, `TbConfiguracion`): es `Registro_Datos.accdb`, no el backend al que apunta el front |
| El catálogo de la Lanzadera distingue dos aplicaciones de registro | `TbAplicaciones` de `Lanzadera_Datos.accdb`: `IDAplicacion = 10` «Registro» → `Registro_Ent_Salida.accdb` / `Registro_Ent_Salida_Datos.accdb`, e `IDAplicacion = 101` «RegistroAntiguo» → `Registro.accdb` / `Registro_Datos.accdb` |

Consecuencia: el inventario de abajo describe `main` de `00_REGISTROS` y el backend `Registro_Ent_Salida_Datos.accdb`, no `main` de otro repositorio ni una RC1. El «registro antiguo» es **otra aplicación** del catálogo, con su propio front, no un módulo de la aplicación estudiada.

### `Registro_Datos.accdb` y `Registro_Ent_Salida_Datos.accdb` no son el mismo backend

No son el mismo archivo con dos nombres: son **dos backends distintos**, y el segundo además **vincula** una tabla del primero.

| Comprobación | Resultado |
|---|---|
| `SELECT COUNT(*) FROM Entradas` sobre `Registro_Ent_Salida_Datos.accdb` | Error: «El motor de base de datos de Microsoft Access no puede encontrar la tabla o consulta de entrada 'Entradas'» |
| `SELECT COUNT(*) FROM Salidas` sobre `Registro_Ent_Salida_Datos.accdb` | 13630 filas: `Salidas` sí existe, pero como **vínculo** a `Registro_Datos.accdb` (`docs/ERD-auxiliar/Registro_Ent_Salida_Datos.sql:4`: `-- LINKED/SKIPPED: [Salidas] (given file does not exist: C:\00repos\datos\Registro_Datos.accdb)`) |
| `Registro_Datos.accdb` | 5 tablas propias, 0 vínculos, 1 relación (la de sistema `MSysNavPaneGroupCategoriesMSysNavPaneGroups`), 1 consulta guardada (`Consulta1`), ninguna sin clave primaria |
| Archivos | Distinto nombre, distinto tamaño y distinta fecha: `Registro_Datos.accdb` (5.193.728 bytes, 11-03-2024) y `Registro_Ent_Salida_Datos.accdb` (11.239.424 bytes, 22-05-2025) |

Consecuencia: el ERD de backend entregado mide el backend de la aplicación «RegistroAntiguo», no el que usa el front. El esquema del backend propio se ha tomado de `docs/ERD-auxiliar/Registro_Ent_Salida_Datos.sql` y de las consultas de solo lectura sobre la copia local.

### El catálogo de vínculos de la Lanzadera

`TbTablasAVincular` relaciona aplicaciones y tablas. La columna `IDBBDD` es **fiable para saber qué aplicación vincula qué tabla** y **no lo es para saber quién es el dueño físico**: la matriz no cubre las aplicaciones modernas (Expedientes, Riesgos, HPS, Solicitudes HPS, Condor y Contratos), de modo que atribuye al backend que enlaza una tabla la tabla de otro backend. Casos medidos: `TbExpedientes` figura bajo `IDBBDD = 2` (AGEDYS) pero vive en `Expedientes_datos.accdb`; `TbCorreosEnviados` figura bajo `IDBBDD = 1`, `2` y `8` pero vive en `Correos_datos.accdb`; `TbProyectos` figura bajo `IDBBDD = 2` y `IDBBDD = 5` y hay una tabla local con ese nombre en cada uno.

Para el backend de Registros (`IDBBDD = 10`, nueve tablas) la matriz dice:

| Tabla | Aplicaciones que la vinculan según la matriz |
|---|---|
| `TbEntradas_RES` | 3 (AGEDYS TEC), 4 (AGEDO), 10 (Registro) |
| `TbSalidas_RES` | 2 (AGEDYS ECO), 3 (AGEDYS TEC), 4 (AGEDO), 10 (Registro) |
| `TbLog` | 10 (Registro) |
| `TbEntradaDestinatarios_RES` | 10 (Registro) |
| `TbSalidaDestinatarios_RES` | 10 (Registro) |
| `TbEntradaPersonasContacto_RES` | 10 (Registro) |
| `TbCorreosEnviados_RES` | 10 (Registro) |
| `TbSalidas_RES_Indexacion` | 2 (AGEDYS ECO), 3 (AGEDYS TEC) |
| `Salidas` | 2 (AGEDYS ECO) |

Tres diferencias con el binario:

1. La matriz dice que Registro vincula **`Entradas` y `Salidas`** desde `IDBBDD = 101` (`Registro_Datos.accdb`, el backend de «RegistroAntiguo»). El front inventariado **no** tiene esos dos vínculos. Es la explicación de por qué la búsqueda de registros antiguos no funciona (ver «Referencias que no resuelven»).
2. La matriz dice que Registro vincula `TbCorreosEnviados_RES` desde su propio backend; el front vincula `TbCorreosEnviados`, sin sufijo, desde `Correos_datos.accdb`.
3. El front vincula `TbSalidas_RES_Indexacion`, que la matriz no le atribuye; la matriz la atribuye a AGEDYS ECO y AGEDYS TEC.

La matriz no registra ningún vínculo a `Registro_Datos.accdb` desde la aplicación «RegistroAntiguo» hacia otro backend, y `IDBBDD = 14` (una sola tabla, `Salidas`) no corresponde a ninguna aplicación registrada: hay un tercer archivo con una tabla `Salidas` que no se ha identificado.

## Estructura del front

| Hecho | Valor |
|---|---|
| Tablas del front | 13: 0 propias y 13 vinculadas |
| Consultas guardadas del front | 3: `Consulta1`, `Consulta2`, `Consulta3` (`src/queries/Consulta1.sql`, `:2`, `:3`) |
| Formularios | 11 |
| Macros y `AutoExec` | Ninguno: cero coincidencias de `autoexec`, `autoopen` y `macro` en `src/` |
| Formularios con `RecordSource` | 0: la cadena `recordsource` no aparece en ninguna línea de `src/` |
| Controles con `ControlSource` | 5, todos con el valor `IDActividad` y sin `RecordSource` de formulario: inertes (`Form_FormRegistroAntiguoDetalle.form.txt:918`, `Form_FormRegistroEntradaAlta.form.txt:949`, `Form_FormRegistroEntradaEdicion.form.txt:942`, `Form_FormRegistroSalidaAlta.form.txt:942`, `Form_FormRegistroSalidaEdicion.form.txt:942`) |
| Controles con `RowSourceType = Table/Query` | 37 |
| Controles con `RowSourceType = Value List` | 17 |

Los formularios son **desconectados**: los controles se rellenan en tiempo de ejecución, como expone el propio código (`Form_FormRegistrosAntiguosBusqueda.cls:289` fija `Me.ComboNregistroBusqueda.RowSource`; `:422` fija `lst.RowSource`; `Form_FormRegistrosGestion.cls:1091` y `:1221` hacen lo mismo con `ListaFiltrados`). La ausencia de `RecordSource` no es ausencia de uso: el origen de datos está en los `RowSource` guardados en el diseño y en las asignaciones en runtime.

### Destino de los 13 vínculos

Cinco rutas distintas para cuatro backends. La misma tabla del mismo backend se declara con grafías diferentes de servidor y de carpeta:

| Tablas | `DATABASE=` de la cadena de conexión |
|---|---|
| `TbEntradaDestinatarios_RES`, `TbEntradaPersonasContacto_RES`, `TbEntradas_RES`, `TbLog`, `TbSalidaDestinatarios_RES`, `TbSalidas_RES` | `\\DATOSTE\aplicaciones_dys\Aplicaciones PpD\Registros\Registro_Ent_Salida_Datos.accdb` |
| `TbSalidas_RES_Indexacion` | `\\DATOSTE\Aplicaciones_dys\Aplicaciones PpD\REGISTROS\Registro_Ent_Salida_Datos.accdb` |
| `TbExpedientes`, `TbProyectos` | `\\DATOSTE\aplicaciones_dys\Aplicaciones PpD\Proyectos\AGEDYS_DATOS.accdb` |
| `TbUsuariosAplicaciones` | `\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\Lanzadera_Datos.accdb` |
| `TbConexionesRegistro`, `TbUsuariosAplicacionesPermisos` | `\\datoste\APLICACIONES_DYS\Aplicaciones PpD\0Lanzadera\Lanzadera_Datos.accdb` |
| `TbCorreosEnviados` | `\\datoste\aplicaciones_dys\Aplicaciones PpD\00Recursos\Correos_datos.accdb` |

Variaciones a documentar: `\\DATOSTE` frente a `\\datoste`; `APLICACIONES_DYS` frente a `aplicaciones_dys` frente a `Aplicaciones_dys`; y el backend propio declarado con dos grafías de carpeta, `…\Registros\` en seis tablas y `…\REGISTROS\` en `TbSalidas_RES_Indexacion`. Todas las cadenas llevan la credencial; en este documento se escriben como `PWD=***`.

Los 13 vínculos apuntan a producción: no hay ruta local en el front.

### Consultas guardadas del front

Tres, y las tres consultan tablas vinculadas desde el front, que es justo lo que la fase 2 quiere eliminar:

| Consulta | Qué hace | Observación |
|---|---|---|
| `Consulta1` | Une `TbSalidas_RES` y `TbEntradas_RES` por `UbicacionFisica` | Correcta |
| `Consulta2` | Listado de `TbSalidas_RES` con `((Year([FSalida]))=2023)` literal | El año está escrito en la consulta: hay que parametrizarla |
| `Consulta3` | Une `TbUsuariosAplicaciones` con `TbConexionesRegistro` y filtra por el SSID fijo `S-1-5-21-2675229729-2419695801-670233423-1001` | Solo lectura; el SSID está escrito en la consulta |

## Aperturas directas

| Apertura | Código | Observación |
|---|---|---|
| Backend propio | `src/modules/Variables de Entorno.bas:352-366` (`getdb()`) | Ruta remota `…\REGISTROS\Registro_Ent_Salida_Datos.accdb` (`:355`, con `m_URLRutaAplicacionRemota` definida en `:215`); ruta local `…\000datoslocal\…` (`:353`). La credencial va literal en la cadena (`PWD=***`) |
| Lanzadera | `src/modules/Variables de Entorno.bas:379-391` (`getdbLanzadera()`) | 8 llamadas en 7 procedimientos: `constructor.bas:74, 114, 182`, `Usuario.cls:178`, `UsuarioAplicacionPermisos.cls:263`, `Funciones Generales.bas:2767, 2772, 2834`. Escribe `TbAplicacionesAperturas`, que **no** está entre los vínculos del front |
| Correos | `src/modules/Variables de Entorno.bas:412-428` (`getdbCorreo()`) | 2 llamadas: `Funciones Generales.bas:2969, 2972` |
| `CurrentDb()` | 34 menciones en 10 ficheros, 2 de ellas en bloques comentados (`Funciones Generales.bas:2339, 2398`) | Es el mecanismo mayoritario: todas las lecturas y escrituras del front sobre sus vínculos pasan por aquí |
| `DoCmd.OpenDatabase`, `TransferDatabase`, `DoCmd.OpenTable`, `TransferText` | Ninguno | Los 3 `OpenDatabase` del repositorio son los de `Variables de Entorno.bas` |

La rama de datos locales está muerta: `Application.TempVars("DatosEnLocal")` se fija a `"No"` en `src/modules/Variables de Entorno.bas:210` y la línea que lo fijaba a `"Sí"` está comentada (`:209`), de modo que `getdb()` siempre resuelve la ruta remota de producción. `Application.TempVars("EnDesarrollo")` está igual, a `"No"` (`:212`).

`getdb()` se usa en 3 puntos: `constructor.bas:242`, `Funciones Generales.bas:512` (`DameID1`) y `JsonConverter.bas:1257`.

## Tablas de otros backends usadas por código

| Tabla | Backend real | Cómo la usa el front | Quién más la vincula | Dueño |
|---|---|---|---|---|
| `TbUsuariosAplicaciones` | `Lanzadera_Datos` | Vínculo (71 referencias) y apertura directa | Es identidad de la Lanzadera | Lanzadera |
| `TbUsuariosAplicacionesPermisos` | `Lanzadera_Datos` | Vínculo (6 referencias) y apertura directa (`constructor.bas:114`) | Es identidad de la Lanzadera | Lanzadera |
| `TbConexionesRegistro` | `Lanzadera_Datos` | Vínculo (5 referencias, solo `Consulta3`) | — | Lanzadera |
| `TbAplicacionesAperturas` | `Lanzadera_Datos` | Solo por apertura directa (3 referencias): `Funciones Generales.bas:2767, 2772, 2834` | — | Lanzadera |
| `TbProyectos` | `AGEDYS_DATOS` (tabla local de AGEDYS) | Vínculo, en 4 `RowSource` de combos (12 referencias) | La matriz la sitúa también en el backend de Riesgos (IDBBDD 5) | AGEDYS, compartida con Registros |
| `TbExpedientes` | `Expedientes_datos` | Vínculo, en 5 `RowSource` de combos (15 referencias) | AGEDYS ECO, AGEDYS TEC y AGEDO la vinculan desde `AGEDYS_DATOS` | Expedientes: **mal ubicada** para Registros |
| `TbCorreosEnviados` | `Correos_datos` | Vínculo, escritura y lectura (5 referencias) | AGEDYS ECO, AGEDYS TEC, No Conformidades y Sirdee | Correos: **dueño por decidir** |
| `tbUsuarios` | `Lanzadera_Datos` | Referencia en un `RowSource` (`Form_FormRegistroEntradaAlta.form.txt:1879`) | — | Lanzadera, pero **no está vinculada** en el front: la referencia no resuelve |

Sobre R-1 (cada front consulta solo su backend): el front de Registros **no se cumple**. Consulta cuatro backends: el propio, `AGEDYS_DATOS`, `Lanzadera_Datos` y `Correos_datos`, y abre dos de ellos directamente. La excepción de las tablas de identidad es la misma que se acepta en Brass y está descrita en la [fase 0 de la Lanzadera](lanzadera-datos.md#verificación-de-r-1-y-r-2); la lectura de `TbExpedientes` y `TbProyectos` desde AGEDYS no lo es: son tablas de negocio de otra aplicación.

## Backend `Registro_Ent_Salida_Datos.accdb`

| Hecho | Valor |
|---|---|
| Tablas | 11 propias, 1 vinculada (`Salidas` → `Registro_Datos.accdb`) |
| Relaciones | 0 de usuario: la única relación es `MSysNavPaneGroupCategoriesMSysNavPaneGroups`, artefacto del panel de navegación |
| Consultas guardadas | 1: `Consulta1` (resuelve por nombre; `Consulta2` no existe en este backend) |
| Tablas sin clave primaria | No verificable con la interfaz de solo lectura usada: el volcado de que se dispone no incluye índices (ver «Lo que no se pudo verificar») |
| Tipos problemáticos | Ninguno: no hay campos multivalor, de datos adjuntos ni objeto OLE. Los tipos presentes son texto, entero largo, fecha/hora y memo (`docs/ERD-auxiliar/Registro_Ent_Salida_Datos.sql` y `get_schema` de `TbEntradas_RES` y `TbSalidas_RES`) |
| Duplicados en las claves | Ninguno: 635/635 `IDEntrada` distintos, 4053/4053 `IDSalida` distintos, 3341/3341 `IDLog` distintos |

Filas por tabla (copia local):

| Tabla | Filas | Vinculada en el front | La vinculan además (matriz) |
|---|---|---|---|
| `TbEntradas_RES` | 635 | Sí | AGEDYS TEC, AGEDO |
| `TbSalidas_RES` | 4053 | Sí | AGEDYS ECO, AGEDYS TEC, AGEDO |
| `TbLog` | 3341 | Sí | — |
| `TbSalidas_RES_Indexacion` | 1063 | Sí | AGEDYS ECO, AGEDYS TEC |
| `TbEntradaDestinatarios_RES` | 37 | Sí | — |
| `TbConexiones` | 12 | No | — |
| `TbSalidaDestinatarios_RES` | 5 | Sí | — |
| `TbEntradas_RES_Indexacion` | 2 | No | — |
| `TbEntradaPersonasContacto_RES` | 0 | Sí | — |
| `TbEntradas_indexacion` | 0 | No | — |
| `TbSalidas_Indexacion` | 0 | No | — |
| `Salidas` (vínculo) | 13630 | No | AGEDYS ECO |

Huérfanos medidos:

| Relación | Huérfanos |
|---|---|
| `TbEntradaDestinatarios_RES.IDEntrada` → `TbEntradas_RES.IDEntrada` | 0 |
| `TbSalidaDestinatarios_RES.IDSalida` → `TbSalidas_RES.IDSalida` | 0 |
| `TbSalidas_RES_Indexacion.IDSalida` → `TbSalidas_RES.IDSalida` | 12 |
| `TbEntradas_RES_Indexacion.IDEntrada` → `TbEntradas_RES.IDEntrada` | 0 |
| `TbLog.IDEntrada` → `TbEntradas_RES.IDEntrada` | 613 (de 2339 con valor; las otras 1002 filas tienen `IDEntrada` nulo) |
| `TbLog.IDSalida` → `TbSalidas_RES.IDSalida` | 3 |

## Backend legacy `Registro_Datos.accdb`

Es el backend de la aplicación «RegistroAntiguo» (`IDAplicacion = 101`, front `Registro.accdb`), no del front estudiado.

| Hecho | Valor |
|---|---|
| Tablas | 5 propias, 0 vinculadas: `DClases` (9), `Entradas` (14905), `Nombres` (57), `Salidas` (13630), `TbConfiguracion` (1) |
| Relaciones | 1, y es la de sistema del panel de navegación |
| Consultas guardadas | 1: `Consulta1` |
| Tablas sin clave primaria | Ninguna: `DClases` por `Cod`, `Entradas` por `IDEntrada`, `Nombres` por `Cod`, `Salidas` por `IDSalida`, `TbConfiguracion` por `DirectorioPorDefectoAnexosEntrada` |
| Quién lo usa | El backend propio de Registros mantiene un **vínculo** a su tabla `Salidas`; la matriz de la Lanzadera dice además que el front de Registros vincula `Entradas` y `Salidas` desde aquí, y que AGEDYS ECO y AGEDYS TEC vinculan `Salidas` |

## Referencias que no resuelven

El front no vincula `Entradas` ni `Salidas`, pero el código las usa. Son referencias rotas: fallan en runtime con «no puede encontrar la tabla o consulta de entrada». La matriz de la Lanzadera dice que esos dos vínculos deberían existir (ver «El catálogo de vínculos de la Lanzadera»), de modo que el diagnóstico es **un front al que le faltan dos vínculos**, no código muerto.

| Referencia | Código | Qué le pasa |
|---|---|---|
| `Entradas` | `Form_FormRegistrosAntiguosBusqueda.cls:288, 330-331, 340, 366, 373, 379, 533, 539`; `Form_FormRegistroAntiguoDetalle.cls:90`; `Funciones Generales.bas:2511` | La tabla vive en `Registro_Datos.accdb` y el front no la vincula. Todo el formulario de búsqueda de registros antiguos depende de ella |
| `Salidas` | `Form_FormRegistrosAntiguosBusqueda.cls:288, 334-335, 342, 368, 375, 381, 535, 541`; `Form_FormRegistroAntiguoDetalle.cls:95`; `Form_FormRegistrosAntiguosBusqueda.form.txt:960`; `Funciones Generales.bas:2452, 2455, 2468, 2514` | Igual que `Entradas`. Además `Salidas.Anio` (`:2455`) no existe en el esquema actual de `Salidas` |
| `TbSalidas` | `Form_FormRegistroSalidaEdicion.form.txt:1470` | Tabla que no existe en ningún backend del ecosistema: el combo de destinatarios de la salida en edición no carga su lista |
| `tbUsuarios` | `Form_FormRegistroEntradaAlta.form.txt:1879` | Existe en `Lanzadera_Datos` (con ese nombre en minúsculas) pero el front no la vincula: el combo de interesados no carga su lista |
| `TbHerramientaDocAyuda` | `Funciones Generales.bas:317` (dentro de `URLCompletaArchivoAyuda`, `:294`) | La tabla existe en `AGEDYS_DATOS` (10 filas) pero no está vinculada. La función no tiene llamadores: `VerAyudaFormulario` (`:175`) no se invoca desde ningún sitio |
| `TbExpedientes.Expediente` y `TbExpedientes.TITULOEXP` | `Form_FormRegistroEntradaAlta.form.txt:2332`, `Form_FormRegistroEntradaEdicion.form.txt:2035`, `Form_FormRegistroSalidaAlta.form.txt:2325`, `Form_FormRegistroSalidaEdicion.form.txt:1728`, `Form_FormRegistrosGestion.form.txt:3995` | El esquema vigente de `TbExpedientes` (en `Expedientes_datos.accdb`) tiene `IDExpediente`, `Nemotecnico`, `Titulo` y `CodProyecto`; no tiene `Expediente` ni `TÍTULOEXP` |
| `Entradas.Año` y `Salidas.Año` | `Form_FormRegistrosAntiguosBusqueda.form.txt:2414` | Columnas inexistentes en el esquema vigente |
| `TbSalidas_Indexacion` | `Form_FormRegistrosAntiguosBusqueda.form.txt:961-964` | La tabla existe en el backend propio (0 filas) pero no está vinculada en el front, así que el `RowSource` guardado del formulario no resuelve |

Dos advertencias sobre el diagnóstico de estas referencias: no son un error de captura, son la consecuencia de un cambio de esquema (el backend nuevo renombró `Entradas`/`Salidas` a `TbEntradas_RES`/`TbSalidas_RES`) que no se acompañó en el front. Y el `RowSource` de `TbExpedientes` apunta a `AGEDYS_DATOS`, donde `TbExpedientes` es a su vez un **vínculo** a `Expedientes_datos.accdb`: Access no recorre cadenas de vínculos, de modo que ese combo no puede resolver ni aunque existieran las columnas.

## Veredicto de propiedad

**Propias de Registros y compartidas** (vive en `Registro_Ent_Salida_Datos.accdb`; la usa el front y la vinculan además otras aplicaciones): `TbEntradas_RES` (184 referencias en el front; la vinculan AGEDYS TEC y AGEDO) y `TbSalidas_RES` (194; la vinculan AGEDYS ECO, AGEDYS TEC y AGEDO). Son las dos tablas que escriben y leen el front, y las que hacen del backend un núcleo compartido.

**Propias de Registros** (la usa el front; la matriz no registra ninguna otra aplicación que la vincule): `TbLog` (13), `TbEntradaDestinatarios_RES` (4) y `TbSalidaDestinatarios_RES` (6).

**No la usa el front, pero no es huérfana**: `TbSalidas_RES_Indexacion` (1063 filas). La matriz la vincula desde AGEDYS ECO y AGEDYS TEC; el front la vincula pero no la usa en ninguna línea de código. Lo que sobra es el vínculo del front, no la tabla.

**No la usa el front ni ninguna otra aplicación**: `TbEntradaPersonasContacto_RES` (0 filas; la matriz la atribuye a Registro), `TbEntradas_indexacion` (0 filas), `TbEntradas_RES_Indexacion` (2 filas), `TbSalidas_Indexacion` (0 filas) y `TbConexiones` (12 filas). Ninguna aparece en el SQL ni en los `RowSource` del front, y ninguna figura en la matriz. Son candidatas a retirada, sujetas al grafo de vínculos de todos los backends (tarea T6).

**Mal ubicadas**: `TbExpedientes`, que el front de Registros usa en 5 combos pero cuya dueña física es Expedientes (está en `Expedientes_datos.accdb`; `AGEDYS_DATOS` solo la vincula, y la matriz la atribuye a AGEDYS porque no cubre las aplicaciones modernas). `TbProyectos`, que es de AGEDYS y la usa Registros en 4 combos: en ese caso es «propia y compartida», no mal ubicada.

**Dueño por decidir**: `TbCorreosEnviados`, de `Correos_datos`, que la usan Registros, AGEDYS ECO, AGEDYS TEC, No Conformidades y Sirdee. `Salidas` y `Entradas`, del backend legacy, que la matriz vincula a Registros y AGEDYS ECO.

**Fuera del alcance de esta ficha**: `Registro_Datos.accdb` y las 5 tablas de la aplicación «RegistroAntiguo», que el front estudiado ya no alcanza pero a la que el catálogo y el backend propio siguen apuntando.

### Coste medido

Referencias en `src/` por tabla (solo este front; sumar los demás fronts es trabajo del grafo global). Los ficheros `.form.txt` incluyen una copia del código de su `.cls`, así que el recuento de `src/` duplica parte de las referencias: el coste de cambio real de un formulario es la mitad de lo que indica la tabla.

| Tabla | Referencias |
|---|---|
| `TbSalidas_RES` | 194 |
| `TbEntradas_RES` | 184 |
| `TbUsuariosAplicaciones` | 71 |
| `Salidas` (no resuelve) | 57 |
| `Entradas` (no resuelve) | 39 |
| `TbExpedientes` | 15 |
| `TbLog` | 13 |
| `TbProyectos` | 12 |
| `TbUsuariosAplicacionesPermisos` | 6 |
| `TbSalidaDestinatarios_RES` | 6 |
| `tbUsuarios` (no resuelve) | 6 |
| `TbConexionesRegistro` | 5 |
| `TbCorreosEnviados` | 5 |
| `TbSalidas` (no resuelve) | 4 |
| `TbEntradaDestinatarios_RES` | 4 |
| `TbSalidas_Indexacion` (no resuelve) | 3 |
| `TbAplicacionesAperturas` | 3 |
| `TbHerramientaDocAyuda` (no resuelve) | 2 |
| `TbSalidas_RES_Indexacion` | 0 |
| `TbEntradaPersonasContacto_RES` | 0 |
| `TbEntradas_RES_Indexacion` | 0 |
| `TbEntradas_indexacion` | 0 |
| `TbConexiones` | 0 |

## Propuestas para después del descubrimiento global

Ninguna se aplica: la puerta de descubrimiento global sigue cerrada y la fase 1 no empieza hasta terminar el inventario de todas las herramientas. Grupos según la [clasificación de la épica](../epica-coexistencia-postgresql.md#clasificación-de-tablas): (a) imprescindible antes de migrar, (b) recomendable si es barato, (c) aplazar a la refactorización final.

| Id | Propuesta | Grupo | Coste | Comprobación previa | Reversión |
|---|---|---|---|---|---|
| G1 | Poner en cuarentena las 5 tablas del backend propio que no usa ninguna aplicación: renombrar a `zzBAJA_<tabla>`. | (a) | 0 referencias; 12 filas en `TbConexiones` y 2 en `TbEntradas_RES_Indexacion` | Confirmar con el grafo de vínculos de todos los backends que ninguna aplicación las vincula (tarea T6). La matriz de la Lanzadera no las registra. Repetir los recuentos sobre producción | Quitar el prefijo |
| G2 | Sustituir los generadores `Max + 1` por autonumérico o identidad. Hay cinco: `DameID` (`Funciones Generales.bas:437`), `DameID1` (`:495`), `constructor.getID` (`constructor.bas:222`), `DameNRegistro` (`:374`) y `DameCRNuevo` (`:1957`) | (a) | `DameID` se llama con `TbEntradas_RES` (`:578`), `TbSalidas_RES` (`:762`), `TbLog` (`:1907`) y `TbCorreosEnviados` (`Correo.bas:58`); `DameID1` con `TbAplicacionesAperturas` (`:2767`); `constructor.getID` con `TbCorreosEnviados` (`:2969`) | Comprobar en la copia que no hay huecos ni duplicados en los identificadores actuales y definir el punto de corte de la numeración | Dejar el generador en su sitio y volver al contador anterior |
| G3 | Declarar clave primaria en las tablas del backend propio que no la tengan, una vez se pueda leer el catálogo de índices (hoy no verificable, ver «Lo que no se pudo verificar») | (a) | 11 tablas | Leer los índices de las 11 tablas con una extracción que los incluya y comprobar duplicados antes de fijar la clave | Quitar la clave |
| G4 | Reponer los dos vínculos que la matriz declara y el front no tiene: `Entradas` y `Salidas` desde `Registro_Datos.accdb`, y corregir las columnas `Año` que el código usa y el esquema no tiene | (a) | 25 referencias en código (`.bas`/`.cls`) y 2 `RowSource` guardados | Comprobar en producción si el front desplegado los tiene y este binario los perdió. Reproducir en runtime la búsqueda de registros antiguos y el detalle de registro antiguo | Quitar los vínculos y volver al SQL original |
| G5 | Resolver las referencias que no tienen tabla: `TbSalidas` (combo de destinatarios en la salida en edición) y `tbUsuarios` (combo de interesados en el alta de entrada) | (a) | 2 `RowSource` guardados | `tbUsuarios` existe en `Lanzadera_Datos`: la alternativa es vincularla. `TbSalidas` no existe en ningún backend: hay que decidir si se corrige a `TbSalidas_RES` o si el combo sobra | Restituir el `RowSource` original |
| G6 | Decidir el destino de `Registro_Datos.accdb` y de sus 5 tablas: `Entradas` (14905), `Salidas` (13630), `Nombres`, `DClases` y `TbConfiguracion`, que son la aplicación «RegistroAntiguo» del catálogo | (a) | 0 referencias vivas desde el front; 3 vínculos entrantes (el backend propio y `AGEDYS_DATOS` hacia `Salidas`) | Comprobar si «RegistroAntiguo» sigue en uso y si `AGEDYS` usa de verdad `Salidas`; comprobar consumidores fuera de Access (Excel, informes, scripts). Si se conserva, declararlo como archivo histórico de solo lectura | Nada que revertir: es una decisión, no un cambio |
| G7 | Quitar el vínculo del front a `TbSalidas_RES_Indexacion`: la tabla la usan AGEDYS ECO y AGEDYS TEC, no Registros | (b) | 0 referencias en el código; 1063 filas en la tabla | Confirmar en producción que ningún flujo del front la usa (búsqueda de texto sobre `src/` y sobre los `RowSource` guardados) | Reponer el vínculo |
| G8 | Aislar el desarrollo de la producción: los vínculos del front apuntan a `\\datoste\…` y `getdb()` resuelve siempre la ruta remota porque `Application.TempVars("DatosEnLocal")` está fijado a `"No"` (`Variables de Entorno.bas:210`) | (b) | 13 vínculos y 3 abridores | Confirmar con el propietario el entorno de cada copia y los permisos de escritura en el share | Restituir la configuración original |
| G9 | Normalizar las rutas: cinco grafías distintas para cuatro backends, con `\\DATOSTE` frente a `\\datoste`, tres formas de escribir el nombre de la carpeta de aplicaciones y dos del backend propio | (b) | 13 vínculos | Comprobar en la copia que la unificación no rompe la resolución de ningún vínculo; hacer el cambio en la copia, no en producción | Restituir la cadena original |
| G10 | Retirar la credencial literal de las tres cadenas de apertura (`Variables de Entorno.bas:363, 391, 428`) y leerla de la configuración | (b) | 3 literales | Es la misma credencial que abre todos los binarios del ecosistema. Comprobar que no queda ninguna otra copia del literal en el front | Restaurar la línea original |
| G11 | Mover las 3 consultas guardadas del front al backend y parametrizar `Consulta2` (año literal) y `Consulta3` (SSID literal) | (c) | 3 consultas, 2 literales | Comprobar que ninguna otra aplicación usa esas consultas por su nombre y que la parametrización no rompe las llamadas | Devolver las consultas al front |
| G12 | Corregir el `RowSource` de `TbExpedientes` (columnas `Expediente` y `TÍTULOEXP`, que no existen) y de `TbProyectos` para el esquema vigente (`IDExpediente`, `Nemotecnico`, `Titulo`, `CodProyecto`), y apuntar al backend dueño, `Expedientes_datos`, en vez de a `AGEDYS_DATOS` | (b) | 5 `RowSource` | Reproducir en runtime la carga de los combos de expediente y de proyecto en alta y edición de entrada y salida | Restituir el `RowSource` original |
| G13 | Retirar la dirección de correo nominal fija escrita en el código (`Correo.bas:44`) | (b) | 1 literal | Comprobar a quién corresponde la copia oculta y cómo se configura hoy | Sustituir por la dirección configurada |
| G14 | Renombrar `Tb*_RES` a una nomenclatura estable | (c) | 378 referencias (`TbEntradas_RES` 184 + `TbSalidas_RES` 194) más las 3 consultas | Dejarlo para la refactorización final, tras el apagado | — |

Lo que **no** se toca ahora: los 13 vínculos del front, las 5 tablas que sí usa el front, la tabla `Salidas` del backend legacy mientras `AGEDYS` la use, y las 3 consultas guardadas del front.

## Consultas de reproducción

Sobre `C:\00repos\datos\Registro_Ent_Salida_Datos.accdb`, en modo de solo lectura:

```sql
SELECT COUNT(*) AS n FROM Entradas;         -- error: no existe
SELECT COUNT(*) AS n FROM Salidas;          -- 13630, es un vínculo a Registro_Datos.accdb
SELECT COUNT(*) FROM (SELECT DISTINCT IDEntrada FROM TbEntradas_RES);  -- 635
SELECT COUNT(*) FROM TbLog WHERE IDEntrada <> 0
  AND IDEntrada NOT IN (SELECT IDEntrada FROM TbEntradas_RES);         -- 613 huérfanos
```

Sobre `C:\00repos\datos\Lanzadera_Datos.accdb`:

```sql
SELECT IDAplicacion, NombreAplicacion, NombreEjecutable, NombreArchivoDatos
  FROM TbAplicaciones WHERE IDAplicacion IN (10, 101);

SELECT DISTINCT NombreTabla FROM TbTablasAVincular WHERE IDBBDD=10 ORDER BY NombreTabla;

SELECT IDAplicacion, NombreTabla FROM TbTablasAVincular
  WHERE IDBBDD=10 OR IDBBDD=101 GROUP BY IDAplicacion, NombreTabla;
```

## Lo que no se pudo verificar

1. **Claves primarias e índices del backend propio.** El ERD disponible (`backend-registro/erd.json`) es de `Registro_Datos.accdb`; el volcado de `Registro_Ent_Salida_Datos.accdb` (`docs/ERD-auxiliar/Registro_Ent_Salida_Datos.sql`) no incluye índices y la interfaz de solo lectura usada no los expone. La propuesta G3 queda condicionada a esa lectura.
2. **Producción.** Los vínculos apuntan a `\\datoste\…`, no alcanzable desde este entorno: el ERD del front registra un error de vínculo por cada una de las 13 tablas. Todo el esquema y los recuentos del backend propio salen de la copia local `C:\00repos\datos\Registro_Ent_Salida_Datos.accdb` (22-05-2025), que puede no coincidir con producción. Los recuentos de `Expedientes_datos.accdb` y `AGEDYS_DATOS.accdb` son de copias locales del 04-11-2026 y del 23-10-2025.
3. **UAT.** No hay rama RC1, ni `docs/`, ni parte de resultados ni firma en el repositorio de Registros. La base de análisis es `main` del SHA indicado y no hay reconciliación con una RC1 que hacer.
4. **Reglas de validación y macros de datos** de las tablas del backend propio: no se han inspeccionado.
5. **Comportamiento en runtime.** Las referencias rotas se deducen del esquema, de la lista de vínculos y de la matriz, no de un fallo observado: los vínculos no se pudieron abrir.
6. **La aplicación «RegistroAntiguo»** (`Registro.accdb`) no se ha inventariado: no hay repositorio con su código en este entorno. Su front sigue registrado en `TbAplicaciones` y su backend sigue en el share.
7. **El backend `IDBBDD = 14`**, que aporta una sola tabla `Salidas` y no corresponde a ninguna aplicación registrada: no se ha identificado el archivo.
8. **Consumidores fuera de Access** (Excel, informes, scripts programados) de las 5 tablas que no usa ninguna aplicación, y de `Registro_Datos.accdb` entero.
9. **Coste por columna.** Se ha medido el coste por tabla; el desglose por columna y la suma de los demás fronts quedan para el grafo global.
10. **Fidelidad de la matriz de vínculos.** `TbTablasAVincular` no coincide con el front en `TbCorreosEnviados_RES` ni en `TbSalidas_RES_Indexacion`, y atribuye a otros dueños tablas que viven en otro backend. No se ha averiguado si la diferencia es de la matriz, del front o de ambos.
