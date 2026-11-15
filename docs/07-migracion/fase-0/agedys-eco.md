# Fase 0: AGEDYS ECO (`Gestion Proyectos Eco 30`)

Issue #1574. Épica `docs/07-migracion/epica-coexistencia-postgresql.md` (fase 0).
Base de análisis: repositorio `00_AGEDYS_ECO` (GitHub `DysTelefonica/AGEDYS_ECO`), rama `main`, SHA `a263ead7c2e6bab3aae85fdb07ce2193d5f91dba`.
Fuentes: ERD del front, de `AGEDYS_DATOS.accdb` y de `Registro_Ent_Salida_Datos.accdb` con `scripts/dysflow-erd/extract-erd.ps1` (solo lectura, sobre copia) y trazado del código con CodeGraph y con `read`/`grep` sobre `src/`. CodeGraph sí estaba operativo en este entorno (`codegraph status` sobre `00_AGEDYS_ECO` devuelve 163 ficheros y 8726 nodos; `codegraph explore` devuelve el fuente actual del disco), así que no hay fallback que declarar.

Las rutas del front son relativas a `C:\00repos\codigo\00_AGEDYS_ECO`. Las de los backends, a `C:\00repos\datos`.

Por decisión del propietario (issue #1623, decisión 13 de la épica), AGEDYS Técnico y AGEDYS ECO se sustituyen por una sola aplicación web, **APGEDYS**, que muestra vista de técnico o de gestor económico según el usuario que entre. Esta ficha inventaría la mitad económica. La mitad técnica es la ficha gemela de AGEDYS Técnico.

## Base de análisis

| Hecho | Evidencia |
|---|---|
| La rama base es `main` | `git branch -a` en `00_AGEDYS_ECO` devuelve `main` y `chore/dysflow-init` en local, y en remoto `main`, `chore/agents-md-inicial`, `chore/dysflow-init`, `chore/issue-forms`, `chore/team-skills-manifest`, `skill-fleet/AGEDYS_ECO` y `skill-fleet/AGEDYS_ECO-vendored` |
| **No hay ninguna rama RC1** | Las siete ramas remotas son `main`, cuatro `chore/*` y dos `skill-fleet/*`. Ninguna es de versión, de release ni de UAT |
| El SHA base es `a263ead7c2e6bab3aae85fdb07ce2193d5f91dba` | `git rev-parse main`; el commit es «Merge pull request #5 from DysTelefonica/skill-fleet/AGEDYS_ECO», de 2026-10-23 |
| El árbol de trabajo no tiene cambios versionados | `git status --short` solo devuelve `?? .codegraph/` y `?? "Gestion Proyectos Eco 30.accdb"`, ambos sin versionar |
| **No hay ninguna evidencia de UAT** | El repositorio no tiene `docs/`: solo `README.md` (2 líneas), `CLAUDE.md`, `AGENTS.md`, `.agents/`, `.atl/`, `.dysflow/` y `src/`. `AGENTS.md:628` y `:670` mencionan UAT solo como disparador de skills del catálogo («Tests de contrato, UAT o helpers de formularios», «capability-docs, feature-acceptance-uat»), no como estado de esta aplicación |
| La versión que muestra la Lanzadera no se puede verificar desde el repositorio | La Lanzadera la pinta en tiempo de ejecución desde el `.ini` del servidor (`00_LANZADERA/src/classes/Aplicacion.cls:1133`, `LeerIni(Me.URLIniServidor, "VersionAplicacion", "", Me.Error)`; `00_LANZADERA/src/modules/Funciones Genreales.bas:2326`, `m_Titulo = "Ver. Servidor: " & m_objAplicacion.VersionServidor`). El share `\\datoste\…` no es alcanzable desde este entorno. El front lee la misma clave de su propio `.ini` (`src/classes/Entorno.cls:2199`, `LeerIni("VersionAplicacion", "#ERR#")`), y **ese `.ini` no está en el repositorio**: el hermano AGEDYS Técnico sí lo versiona (`00_AGEDYS/GestionTecnicaProyectos.ini`, `VersionAplicacion=2026-001`) |
| Lo que sí se puede verificar del catálogo | `Lanzadera_Datos.accdb` → `TbAplicaciones`, fila `IDAplicacion` 2: `NombreAplicacion` «AGEDYS ECO», `NombreEjecutable` «Gestion Proyectos Eco 30.accdb», `NombreArchivoDatos` «AGEDYS_DATOS.accdb», `NombreCarpeta` «Proyectos», `EnPruebas` «No». La fila 3 es «AGEDYS TEC» con `NombreArchivoDatos` «AGEDYS_DATOS.accdb» |
| La identidad de la aplicación en el código | `src/modules/Variables Globales.bas:688`, `IDAplicacion = "2"`, que coincide con el `IDAplicacion` del catálogo |

Consecuencia: no hay reconciliación con una RC1 que hacer. El inventario de abajo describe `main` en el SHA indicado, que es también el contenido del binario del repositorio.

## Estructura del front

| Hecho | Valor |
|---|---|
| Binario | `Gestion Proyectos Eco 30.accdb`, 37.617.664 bytes, `sha256` `721d112df8486528…`, fecha 2026-10-23 |
| Tablas del front | 10, **todas propias**, 0 vinculadas |
| Tablas propias del front | `TbAuxExpedientesVivos`, `TbAuxExpPrado`, `TbAuxFacturas`, `TbAuxFacturasCliente`, `TbAuxFacturasDetalleNoDYS`, `TbAuxGTVContratos`, `TbAuxParaProvisiones`, `TbAuxPedidos`, `TbAuxPedidosGustavo`, `TbAuxProyectos` |
| Vínculos del front | **Ninguno.** El ERD del front no declara ninguna tabla vinculada. El contraste con el hermano AGEDYS Técnico es el dato relevante: su front tiene 127 vínculos, 101 de ellos a `\\datoste\aplicaciones_dys\Aplicaciones PpD\Proyectos\AGEDYS_DATOS.accdb` |
| Relaciones del front | 1, y es de sistema (`MSysNavPaneGroupCategoriesMSysNavPaneGroups`) |
| Consultas guardadas del front | 1: `Consulta1`, con SQL `SELECT TbAuxFacturas.* FROM TbAuxFacturas WHERE (((TbAuxFacturas.FECHAFACTURA) Between #12/1/2024# And #1/31/2025#));` (`src/queries/Consulta1.sql`) |
| Formularios | 40 (ERD y `src/forms/`: 40 `.form.txt` y 40 `.cls`) |
| Módulos y clases | 8 módulos (`constructor.bas`, `FUNCIONES UTILES.bas`, `Instalador.bas`, `JsonConverter.bas`, `Módulo1.bas`, `TESTS.bas`, `Variables Globales.bas`, `mdlBarraComando.bas`) y 65 clases |
| Macros o `AutoExec` | Ninguno: el export de dysflow no produce carpeta de macros, y `AutoExec`, `AutoOpen`, `AutoClose` y `Sub Auto` no aparecen en ningún fichero de `src/` |
| Controles con `RowSourceType` de tabla o consulta | **0**: los 39 `RowSourceType` del front son `"Value List"` |
| Controles con `RecordSource` | **0** en todo `src/` |
| Controles con `ControlSource` | **0** en todo `src/` |

La propiedad `RowSourceTypeInt =1` que acompaña a cada combo y lista del export es un artefacto de la exportación, no un indicador de origen en tabla: aparece exactamente 39 veces, tantas como `RowSourceType ="Value List"`, y ninguna vez en un control que lleve SQL.

### La regla del propietario sobre los formularios desconectados, medida

La regla «salvo Brass, los formularios son desconectados» **se cumple en AGEDYS ECO, y por el motivo más fuerte posible**: no hay ni un solo control unido a datos. Ni `RecordSource` a nivel de formulario, ni `ControlSource` con valor, ni un combo o lista con origen en tabla o consulta.

Los 39 controles con `RowSource` llevan listas literales, y ninguno de esos literales es SQL: son cabeceras de columnas que el código rellena en tiempo de ejecución. El anexo recoge los 39.

### Recordsets de tipo tabla

El indicio se confirma: hay **24 aperturas de recordset cuyo argumento es el nombre de una tabla y no una sentencia `SELECT`**. En DAO, `OpenRecordset` con un nombre de tabla abre un recordset de tipo tabla (`dbOpenTable`), que no existe en PostgreSQL y que además es de solo lectura frente a cambios concurrentes.

| Sitio | Tabla | Base de datos |
|---|---|---|
| `src/classes/AdjudicacionCartaOperaciones.cls:253` (apertura en `:254`) | `TbAdjudicacionesOfertasCartas` | AGEDYS_DATOS |
| `src/classes/AdjudicacionEnvioOperaciones.cls:60` (apertura en `:61`) | `TbAdjudicacionesEnvioAlSuministrador` | AGEDYS_DATOS |
| `src/classes/AnexoOperaciones.cls:117` (apertura en `:118`) | `TbDpDDocAnexada` | AGEDYS_DATOS |
| `src/classes/CartasDatosFijosOperaciones.cls:65` (apertura en `:66`) | `TbDpDCartaSolicitudOfertaDatosFijos` | AGEDYS_DATOS |
| `src/classes/CestaAnexoOrdinariaOperaciones.cls:132` (apertura en `:133`) | `TbCestasAnexosOrdinaria` | AGEDYS_DATOS |
| `src/classes/CorreoOperaciones.cls:577` (apertura en `:578`) | `TbCorreosEnviados` | Correos |
| `src/classes/Entorno.cls:830` (apertura en `:831`) | `TbDpDCartaSolicitudOfertaDatosFijos` | AGEDYS_DATOS |
| `src/classes/ExpedienteCodigoComprasOperaciones.cls:65` (apertura en `:72`) | `TbExpedientesCodigoCompras` | AGEDYS_DATOS |
| `src/classes/FacturaOperaciones.cls:988` (apertura en `:994`) | `TbFacturasDetalle` | AGEDYS_DATOS |
| `src/classes/GTVContratoOperaciones.cls:95` (apertura en `:96`) | `TbGTVContratos` | AGEDYS_DATOS |
| `src/classes/ProvisionOperaciones.cls:40` (apertura en `:41`) | `TbCalendarioProvisiones` | AGEDYS_DATOS |
| `src/classes/ProyectoOperaciones.cls:535` y `:549` (apertura en `:562`) | `TbNPedido` | AGEDYS_DATOS |
| `src/classes/SalidaOperaciones.cls:106` (apertura en `:107`) | `TbSalidas_RES` | AGEDYS_DATOS → Registro |
| `src/classes/suministradorOperaciones.cls:182` (apertura en `:188`) | `TbSuministradoresSAP` | AGEDYS_DATOS |
| `src/modules/FUNCIONES UTILES.bas:407` (apertura en `:408`) | `TbCorreosEnviados` | Correos |
| `src/modules/TESTS.bas:95` (apertura en `:96`) | `TbAuxPedidosGustavo` | Front |
| `src/classes/ConsultaOperaciones.cls:182`, `:241`, `:319`, `:387`, `:434`, `:483`, `:555` | las 7 tablas `TbAux*` de trabajo | Front |

En ocho de ellos el nombre de la tabla es el valor de una rama de un `If` cuya otra rama es un `SELECT *`: son las ramas de alta, donde el código abre la tabla para insertar. Las otras 16 abren la tabla sin más.

Los siete últimos son el patrón de la función «Consultas»: la consulta guardada se ejecuta contra `AGEDYS_DATOS`, su resultado se copia a una tabla `TbAux*` del front y de ahí se exporta a Excel. El nombre de la tabla de destino no está en el código, lo dice la columna `TbAux` del catálogo `TbConsultasDetalle` (`src/classes/ConsultaOperaciones.cls:85-110`).

## Aperturas directas

Todas las aperturas están en `src/modules/Variables Globales.bas`. Ninguna usa vínculos: el front abre sus backends por código.

| Apertura | Código | Destino | Llamadas |
|---|---|---|---|
| Backend propio y compartido | `src/modules/Variables Globales.bas:467` (`getdb()`) | `m_URLRutaAplicacionLocal` o `m_URLRutaAplicacionRemota` + `AGEDYS_DATOS.accdb` (`:475`, `:477`). En remoto, `\\datoste\aplicaciones_dys\Aplicaciones PpD\proyectos\AGEDYS_DATOS.accdb` | 255 |
| Lanzadera | `src/modules/Variables Globales.bas:439` (`getdbLanzadera()`) | `…\0Lanzadera\Lanzadera_Datos.accdb` (`:449`) | 12 |
| Correos | `src/modules/Variables Globales.bas:519` (`getdbCorreo()`) | `…\00Recursos\Correos_datos.accdb` (`:530`) | 5 |
| Registro de entradas y salidas | `src/modules/Variables Globales.bas:552` (`getdbSalida()`) | `…\REGISTROS\Registro_Ent_Salida_Datos.accdb` (`:561`) | 2 |
| AGEDO | `src/modules/Variables Globales.bas:620` (`getdbAgedo()`) | `…\AGEDO\AGEDO20_Datos.accdb` (`:629`) | 2 |
| Apertura de correo | `src/modules/Variables Globales.bas:495` (`getdbCorreoApertura()`) | `…\00Recursos\Correo_Apertura.accdb` (`:505`) | **0** |
| Registro antiguo | `src/modules/Variables Globales.bas:587` (`getdbSalidaAntigua()`) | `…\REGISTROS\Registro_Datos.accdb` (`:597`) | **0** |

Las siete llevan la credencial literal en la cadena de conexión (`"MS Access;PWD=" & … & ""`, en `:458`, `:485`, `:510`, `:543`, `:577`, `:611`, `:643`). Es la misma credencial que abre todos los binarios del ecosistema; aquí va enmascarada como `PWD=***`.

`DoCmd.OpenTable` y `TransferDatabase`: ninguno. `DoCmd.OpenDatabase`: ninguno.

`getdbCorreoApertura()` y `getdbSalidaAntigua()` no tienen ningún llamador en `src/`: son código muerto.

### Un detalle de resolución que conviene verificar

Cuatro de los siete abridores eligen su ruta base según `Application.TempVars("DatosEnLocal")`, y lo hacen en el sentido habitual: `…\Lanzadera\Lanzadera_Datos.accdb`, `…\proyectos\AGEDYS_DATOS.accdb` y `…\00Recursos\Correos_datos.accdb` salen de `m_URLRutaAplicacionesRemotas` cuando el dato no está en local (`:449`, `:477`, `:530`), y de `m_URLRutaAplicacionesLocal` cuando sí (`:447`, `:475`, `:528`).

Los tres restantes hacen lo contrario: `getdbSalida()` (`:560-563`), `getdbSalidaAntigua()` (`:596-599`) y `getdbAgedo()` (`:628-631`) toman `m_URLRutaAplicacionesLocal` en la rama «No» y `m_URLRutaAplicacionesRemotas` en la rama «Sí». Como `m_URLRutaAplicacionesLocal` solo se rellena cuando el dato está en local (`src/modules/Variables Globales.bas:693`; en el resto de los casos se deja a cadena vacía, `:698`), la ruta que construyen esas tres funciones en el modo normal de producción es relativa (`REGISTROS\Registro_Ent_Salida_Datos.accdb`), y depende de que el front esté publicado en una carpeta que contenga esa subcarpeta. `src/classes/Entorno.cls:1877` construye `URLDIRECTORIOANEXOSSALIDASREGISTRO` como `URLDirectorioAplicaciones & "REGISTROS\ANEXOS\SALIDAS\"`, lo que sugiere que esa subcarpeta existe junto al ejecutable. No se ha podido comprobar en runtime si la ruta resuelve; queda como propuesta con comprobación previa, no como defecto confirmado.

## Tablas de otros backends que consume el front

El front no usa vínculos, así que todo lo que no es del front llega por `getdb()` o por un abridor directo. El camino de una tabla es siempre el mismo: el front abre `AGEDYS_DATOS.accdb`, y ahí la tabla es o propia o un vínculo hacia su backend dueño.

| Tabla | Backend dueño | Cómo llega el front | Evidencia |
|---|---|---|---|
| `TbSalidas_RES` | Registro (`Registro_Ent_Salida_Datos.accdb`) | Doble camino: por `getdb()` (que resuelve el vínculo de AGEDYS_DATOS) y directo con `getdbSalida()` | `src/classes/Salida.cls:88`, `:187-190`; `src/classes/SalidaOperaciones.cls:106-107` (alta) y `:166-169` (borrado, por `getdbSalida()`); `src/modules/Instalador.bas:433-435` (por `getdb()`) y `:670-674` (por `getdbSalida()`) |
| `TbEntradas_RES` | Registro (`Registro_Ent_Salida_Datos.accdb`) | **Ninguna: el front no la referencia** | Cero apariciones de `TbEntradas_RES` en `src/`. Véase el apartado siguiente |
| `TbExpedientes`, `TbExpedientesResponsables` | Expedientes (`Expedientes_datos.accdb`) | Por `getdb()`, que resuelve el vínculo de AGEDYS_DATOS | `src/modules/constructor.bas:1458`, `:1462` (TbExpedientes, 19 referencias) y `:1340`, `:1398` (TbExpedientesResponsables, 8) |
| `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` | Lanzadera (`Lanzadera_Datos.accdb`) | Doble camino: por `getdb()` y directo con `getdbLanzadera()`, incluido el acceso por metadatos | `src/classes/Usuario.cls:282` y `src/classes/UsuarioAplicacionPermisos.cls:263` (`getdbLanzadera().TableDefs(...)`); `src/modules/constructor.bas:793`, `:832`, `:906`, `:1509`, `:1544`, `:1599`, `:1651`, `:1706` |
| `TbAplicacionesAperturas` | Lanzadera (`Lanzadera_Datos.accdb`) | Directo con `getdbLanzadera()` | `src/modules/FUNCIONES UTILES.bas:966` y `:970-971` |
| `TbDocumentos`, `TbDocumentosID` | AGEDO (`AGEDO20_Datos.accdb`) | Directo con `getdbAgedo()` y por `getdb()` | `src/classes/DocumentoOperaciones.cls:42`, `:114`, `:123-125`, `:194-196`; `src/modules/constructor.bas:5588` |
| `TbCorreosEnviados` | Correos (`Correos_datos.accdb`) | Doble camino: por `getdb()` y directo con `getdbCorreo()` | `src/classes/CorreoOperaciones.cls:28` (borrado por `getdb()`) y `:578`; `src/modules/FUNCIONES UTILES.bas:405`, `:408`; `src/modules/Instalador.bas:536` |

AGEDO está fuera del alcance de la migración (decisión 6 de la épica): sus dos tablas se inventarían como dependencia, no como objetivo.

Dos de estas lecturas merecen una línea propia: `getdbLanzadera().TableDefs("TbUsuariosAplicaciones")` y `getdbLanzadera().TableDefs("TbUsuariosAplicacionesPermisos")` leen **metadatos** de otra base de datos. El acceso por `TableDefs` no tiene equivalente en PostgreSQL.

## La pregunta central: de quién son `TbSalidas_RES` y `TbEntradas_RES`

El enunciado de partida era que el front de AGEDYS ECO trabaja contra tablas de otros backends, entre ellas `TbEntradas_RES` y `TbSalidas_RES`, y que hacía falta saber de qué backend salen y quién es su dueño real. La medición lo resuelve, y corrige una de las dos premisas.

### Dónde viven las dos tablas

| Tabla | `Registro_Ent_Salida_Datos.accdb` | `AGEDYS_DATOS.accdb` | Quién las usa |
|---|---|---|---|
| `TbSalidas_RES` | **Propia**, 4053 filas, clave `PrimaryKey` | **Vinculada** a `…\REGISTROS\Registro_Ent_Salida_Datos.accdb` | Registro (7 vínculos en la matriz), AGEDYS TEC (4), AGEDO (2), **AGEDYS ECO (3)** |
| `TbEntradas_RES` | **Propia**, 635 filas, clave `PrimaryKey` | **Vinculada** a la misma ruta | Registro (7), AGEDYS TEC (4), AGEDO (2). **AGEDYS ECO no la vincula** |

El dueño real es **Registro**, la aplicación `IDAplicacion` 10 del catálogo de la Lanzadera (`NombreEjecutable` «Registro_Ent_Salida.accdb», `NombreArchivoDatos` «Registro_Ent_Salida_Datos.accdb»). AGEDYS ECO es un **escritor secundario legítimo**: registra en el registro de salidas las adjudicaciones de oferta y los envíos al suministrador (`src/classes/SalidaOperaciones.cls`, `src/classes/AdjudicacionCartaOperaciones.cls`, `src/classes/AdjudicacionEnvioOperaciones.cls`). En eso, la tabla está bien ubicada: el escritor no es el dueño, y el dueño la usa más que nadie.

### La premisa que no se cumple

`TbEntradas_RES` **no la usa el front de AGEDYS ECO**: cero referencias en `src/`. Y la matriz de vínculos de la Lanzadera lo confirma: para `IDAplicacion` 2 (AGEDYS ECO), el backend de Registro (`IDBBDD` 10) aporta tres tablas, `Salidas`, `TbSalidas_RES` y `TbSalidas_RES_Indexacion`, y `TbEntradas_RES` no está entre ellas. La matriz de la [ficha de Registros](registros.md) coincide: `TbEntradas_RES` la vinculan AGEDYS TEC, AGEDO y Registro, no AGEDYS ECO.

Lo que sí hereda el front de AGEDYS ECO es el vínculo de `AGEDYS_DATOS` a `TbEntradas_RES`, que no usa. Eso hace que la tabla llegue al front sin que su código la necesite: es dependencia heredada, no dependencia propia.

### El recorrido de `TbSalidas_RES`, que es el que importa

El front llega a la misma tabla física por dos caminos distintos, y cada uno abre una base distinta:

```text
AGEDYS ECO
  ├─ getdb() ──────────────► AGEDYS_DATOS.accdb ──vínculo──► Registro_Ent_Salida_Datos.accdb ──► TbSalidas_RES
  │   (alta: SalidaOperaciones.cls:106-107;
  │    lectura de numeración: Salida.cls:88, :187-190)
  └─ getdbSalida() ────────► Registro_Ent_Salida_Datos.accdb ─────────────────────────────► TbSalidas_RES
      (borrado: SalidaOperaciones.cls:166-169;
       lectura en el instalador: Instalador.bas:670-674)
```

En producción los dos caminos terminan en el mismo fichero, porque el vínculo de `AGEDYS_DATOS` apunta a `\\datoste\aplicaciones_dys\Aplicaciones PpD\REGISTROS\Registro_Ent_Salida_Datos.accdb`, que es el mismo fichero que abre `getdbSalida()`. En el entorno de desarrollo, el vínculo de la copia local apunta a `C:\00repos\datos\Registro_Ent_Salida_Datos.accdb`, y `getdbSalida()` no resuelve a la misma ruta: con `DatosEnLocal` a «No» construye una ruta relativa (véase el apartado anterior).

El riesgo no es hoy la duplicación del dato, es la **asimetría**: la alta va por un camino y el borrado por otro. Si el vínculo de `AGEDYS_DATOS` se rompiera, o se apuntara a otra copia, el front escribiría en una tabla y borraría en otra, sin ningún aviso. En la web, el mismo error se traduciría en filas que no se pueden borrar.

### `TbSalidas_RES_Indexacion`

La matriz de la Lanzadera declara que AGEDYS ECO vincula `TbSalidas_RES_Indexacion` (backend de Registro), pero **el front no la referencia en ninguna línea**: cero apariciones de `Indexacion` en `src/`. La tabla tiene 1063 filas y su dueño es Registro. Lo que sobra es la declaración de vínculo, no la tabla. Es el mismo caso que ya está registrado en la [ficha de Registros](registros.md) para AGEDYS TEC.

## Backends alcanzados

### `AGEDYS_DATOS.accdb` (backend de AGEDYS, AGEDYS ECO y AGEDYS TEC)

| Hecho | Valor |
|---|---|
| Tablas | 137: 102 propias y 35 vinculadas |
| Destino de las 35 vínculos | `Expedientes_datos.accdb` (23), `Lanzadera_Datos.accdb` (4), `AGEDO20_Datos.accdb` (3), `Registro_Ent_Salida_Datos.accdb` (2: `TbEntradas_RES` y `TbSalidas_RES`), `Correos_datos.accdb` (1), `Registro_Datos.accdb` (1) y una autovinculación a sí mismo (1) |
| Relaciones | 4: 2 del usuario (`TbFacturasDetalleTbVisadoFacturas_Nueva`, `TbVisadoFacturas_NuevaTbSuministradoresPuntuacionConDetalle`) y 2 de sistema |
| Consultas guardadas | 5: `Consulta1` (tipo 48, con SQL) y `Consulta2`, `Consulta3`, `Consulta4`, `Consulta5` (tipo 0, de selección) |
| Tablas que el front de AGEDYS ECO necesita | 43: 35 propias de AGEDYS_DATOS y 8 que llegan por vínculo (`TbSalidas_RES`, `TbExpedientes`, `TbExpedientesResponsables`, `TbDocumentos`, `TbDocumentosID`, `TbCorreosEnviados`, `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`) |

Las 35 propias que el front necesita, con el número de referencias de SQL en `src/`:

`TbProyectos` (54), `TbNPedido` (46), `TbFacturasDetalle` (29), `TbAdjudicacionesOfertasCartas` (23), `TbGTVContratos` (17), `TbDpDDocAnexada` (15), `TbGTV` (15), `TbVisadoFacturas_Nueva` (14), `TbSuministradoresSAP` (14), `TbFacturasPrincipal` (13), `TbAdjudicacionesEnvioAlSuministrador` (11), `TbFacturasCliente` (10), `TbFacturasDetalleNoDYS` (10), `TbDpDInformeCondicionamiento` (7), `TbExpedientesCodigoCompras` (7), `TbCalendarioProvisiones` (5), `TbGTVContratosExpedientes` (5), `TbCalendarioProvisionesDetalle` (4), `TbOrdinariasTipoAnexos` (4), `TbVisadosGenerales` (4), `TbCestasAnexosOrdinaria` (3), `TbDPDComiteOfertas` (3), `TbSolicitudesOfertasPrevias` (3), `TbSuministradoresPuntuacionConDetalle` (3), `TbConsultasDetalle` (2), `TbConsultasPrincipal` (2), `TbSolicitudesEnvioAlSuministrador` (2), `TbAuxExpedientesVivos` (1), `TbDpDAsuntoDefensa` (1), `TbDpDCartaSolicitudOfertaDatos` (1), `TbDpDCartaSolicitudOfertaDatosFijos` (1), `TbDpDCartaSolicitudOfertaPlantillasSegunReqBasicos` (1), `TbDPDOfertaSuministrador` (1), `TbExpedientesAnualidades` (1), `TbProductos` (1).

### La matriz de vínculos de la Lanzadera para AGEDYS ECO está desfasada

`Lanzadera_Datos.accdb` → `TbTablasAVincular` declara para `IDAplicacion` 2 (AGEDYS ECO) 78 vínculos: 71 de `AGEDYS_DATOS` (`IDBBDD` 2), 3 de Registro (`IDBBDD` 10), 3 de la Lanzadera (`IDBBDD` 12) y 1 de `IDBBDD` 14, un identificador de backend que la tabla de la épica no documenta y que aporta `Salidas`.

De los 71 vínculos declarados a `AGEDYS_DATOS`, **6 tablas ya no existen** en ese backend: `TbDPDSO`, `TbDPDSODatosCarta`, `TbDPDSOEnvios`, `TbDpDCartaSolicitudOfertaRequerida`, `TbReferencias` y `tblCentros`. Y al revés: 72 tablas que **sí existen** en `AGEDYS_DATOS` no están declaradas para AGEDYS ECO, entre ellas `TbSalidas_RES`, `TbEntradas_RES`, `TbDocumentos`, `TbExpedientes`, `TbFacturasDetalleNoDYS`, `TbGTVContratosExpedientes`, `TbConsultasPrincipal`, `TbCalendarioProvisionesDetalle` y `TbOrdinariasTipoAnexos`, entre otras.

La consecuencia práctica es que **la matriz no sirve para reconstruir el grafo de este front**: el front real no tiene vínculos, y los que la matriz declara no coinciden con lo que el código abre. El grafo de AGEDYS ECO hay que sacarlo del código, como se ha hecho aquí, no de la matriz.

### `Registro_Ent_Salida_Datos.accdb`

| Hecho | Valor |
|---|---|
| Tablas | 12: 11 propias y 1 vinculada |
| Tablas propias | `TbEntradas_RES` (635), `TbSalidas_RES` (4053), `TbSalidas_RES_Indexacion` (1063), `TbEntradas_RES_Indexacion` (2), `TbEntradaDestinatarios_RES` (37), `TbSalidaDestinatarios_RES` (5), `TbEntradaPersonasContacto_RES` (0), `TbEntradas_indexacion` (0), `TbSalidas_Indexacion` (0), `TbConexiones` (12), `TbLog` (3341) |
| Vínculo | `Salidas` → `Registro_Datos.accdb`, el backend antiguo |
| Relaciones | 1, de sistema |

### `Lanzadera_Datos.accdb`

Identidad y permisos: `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` y `TbAplicacionesAperturas`. Están bien ubicadas: son la autoridad de identidad y de permisos del ecosistema.

## Patrones de riesgo del VBA

Medidos sobre `.bas` y `.cls` de `src/`, sin contar los `.form.txt`, que reexportan el mismo code-behind.

| Patrón | Medida | Dónde | Qué hay que cambiar |
|---|---|---|---|
| Generadores `Max + 1` | 2 | `src/modules/FUNCIONES UTILES.bas:534` (`DameID`: `SELECT Max(<tabla>.<campo>) AS Maximo FROM <tabla>` y `CStr(lngOrdinalMaximo + 1)`, `:519-553`) y `src/modules/constructor.bas:2410` (`getID`: `SELECT Max(IDEnvio) AS MáxDeIDEnvio`) | Identidad o autonumérico. `DameID` se llama con `TbSalidas_RES` (`src/classes/Salida.cls:88`), `TbCorreosEnviados` (`src/classes/Correo.cls:47`, `src/modules/FUNCIONES UTILES.bas:405`), `TbAplicacionesAperturas` (`src/modules/FUNCIONES UTILES.bas:966`), `TbDocumentosID` (`src/classes/DocumentoOperaciones.cls:114`), `TbCestasAnexosOrdinaria` (`src/classes/CestaAnexoOrdinaria.cls:35`), `TbProyectos` (`src/classes/Proyecto.cls:304`) y `TbSuministradoresSAP` (`src/classes/suministradorOperaciones.cls:191`) |
| Generadores que apuntan a la base equivocada | 2 | `src/classes/CestaAnexoOrdinaria.cls:35` (`DameID("TbCestasAnexosOrdinaria", IDCestaAnexo, , m_Error)`) y `src/classes/Proyecto.cls:304` (`DameID("TbProyectos", "ID_PROYECTOS", , m_Error)`) | Los dos omiten el argumento de base de datos, así que `DameID` cae a `CurrentDb()` (`src/modules/FUNCIONES UTILES.bas:531-532`), que en este front es el front. Ninguna de las dos tablas está en el front. El primero es alcanzable: `src/classes/CestaAnexoOrdinariaOperaciones.cls:120` lo invoca y `:132-133` abre la tabla con `getdb()`. El segundo no tiene llamadores |
| Recordsets sin opciones de apertura | 215 de 215 | Ninguna de las 215 llamadas a `OpenRecordset` del `src/` pasa un segundo argumento | Decidir el nivel de aislamiento y reescribir el patrón: en PostgreSQL el cursor no es un recordset de tipo tabla |
| Recordsets de tipo tabla | 24 | Véase el apartado anterior | `OpenRecordset` con nombre de tabla en vez de `SELECT` |
| `DELETE * FROM` | 5 | `src/classes/AnexoOperaciones.cls:32`, `src/classes/CestaAnexoOrdinariaOperaciones.cls:51`, `src/classes/CorreoOperaciones.cls:28`, `src/classes/DocumentoOperaciones.cls:42`, `src/classes/ProvisionOperaciones.cls:65` | `DELETE FROM <tabla>` |
| `Is Null` | 90 | Distribuidos por `src/classes/`, `src/modules/constructor.bas` e `src/modules/Instalador.bas` | `IS NULL` |
| `Like '*'` | 20 | `src/forms/Form_FormPedidosGestion.cls:704`, `:711`, `:716`; `src/forms/Form_FormSuministradoresInformeValoracion.cls:80`; `src/modules/constructor.bas:93`, `:98`, `:1103`, `:1108`, … | `ILIKE '%'` o, mejor, quitar el filtro: `Like '*'` no filtra nada |
| Literales de fecha con `#` | 4 | `src/classes/FacturaOperaciones.cls:59`, `:218`; `src/classes/ProyectoOperaciones.cls:376`; `src/modules/Instalador.bas:222` | `DATE '…'` o `TIMESTAMP '…'`. El formato es `mm/dd/yyyy` (`Format(…, "mm/dd/yyyy")`), con el mes y el día intercambiados respecto al formato de Access |
| Fechas de la consulta guardada | 1 | `src/classes/ConsultaOperaciones.cls:943`, `:950` (sustituye `[Introduzca la Primera Fecha]` por `#…#`) | `DATE '…'` |
| `Option Compare Database` | 112 de 113 | Todos los módulos, clases y code-behind de formulario salvo `src/modules/JsonConverter.bas` | Fija la semántica de comparación de cadenas al estilo de Access. En PostgreSQL hay que decidir la collation y reescribir las comparaciones |
| `Seek` | 0 | — | — |
| SQL con valores concatenados | 2 familias | `src/classes/ConsultaOperaciones.cls:943`, `:950` (fechas del diálogo) y `src/modules/constructor.bas:2410` y `:2417` (el límite de `IDEnvio` en `getID`) | Parametrizar |
| `.AddNew` sobre recordset | 36 | `src/classes/AdjudicacionCartaOperaciones.cls:256`, `src/classes/SalidaOperaciones.cls:108`, `src/classes/ConsultaOperaciones.cls:191`, `:251`, `:332`, … | En PostgreSQL no hay `AddNew` sobre cursor: es un `INSERT` |
| `Date()` o `Now()` dentro de SQL | 0 | — | No hay ninguno. `Now()` solo aparece en asignación de campos de VBA (`src/classes/DocumentoOperaciones.cls:203`, `src/classes/FacturaClienteOperaciones.cls:409`, `src/classes/FacturaOperaciones.cls:1009`) |
| `DoCmd.OpenTable`, `TransferDatabase` | 0 | — | — |

## Veredicto de propiedad

Las clases son las de la [clasificación de la épica](../epica-coexistencia-postgresql.md#clasificación-de-tablas).

**Se quedan en el backend (`AGEDYS_DATOS`)**: las 43 tablas que el VBA del front necesita (las 35 propias más las 8 que llegan por vínculo). No hay ninguna mal ubicada entre ellas: el backend de AGEDYS lo usan también AGEDYS TEC y AGEDYS, y su dueño no es AGEDYS ECO.

**Se quedan en su backend, que es de otro**: `TbSalidas_RES` (Registro), `TbExpedientes` y `TbExpedientesResponsables` (Expedientes), `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` y `TbAplicacionesAperturas` (Lanzadera), `TbDocumentos` y `TbDocumentosID` (AGEDO), `TbCorreosEnviados` (Correos). Todas están bien ubicadas en su dueño: son la autoridad de identidad, de permisos, de correo y de registro. Lo que no está resuelto es el modo de acceso, que es el objeto de la fase 2.

**Están mal ubicadas en el front**: las 10 tablas del front. Siete son de trabajo de la función «Consultas» (`TbAuxProyectos`, `TbAuxPedidos`, `TbAuxFacturas`, `TbAuxFacturasCliente`, `TbAuxFacturasDetalleNoDYS`, `TbAuxGTVContratos`, `TbAuxExpedientesVivos`), y las tres restantes no las usa el código. Incumplen la regla de que las tablas del aplicativo viven en su backend, y desaparecen con el front: en APGEDYS tienen que mudarse al backend o desaparecer con él.

**No las usa el front, y son huérfanas o casi**:

- `TbAuxExpPrado` (0 filas, 18 campos, sin clave): cero referencias en el `src/` de AGEDYS ECO. La tabla con ese nombre **sí** la usa el front de AGEDYS Técnico (`00_AGEDYS/src/forms/Form_FormConsultas.cls:426`, `:428`, `:469`; `00_AGEDYS/src/modules/Funciones Generales.bas:16`, `:19`, `:37`). Es una copia muerta de una tabla que pertenece a otro front.
- `TbAuxParaProvisiones` (213 filas, 14 campos, sin clave): cero referencias en `src/`. La función de provisiones usa `TbCalendarioProvisiones` y `TbCalendarioProvisionesDetalle`, que están en AGEDYS_DATOS (`src/classes/ProvisionOperaciones.cls:40`, `:65`, `:308`, `:314`; `src/modules/constructor.bas:208`).
- `TbAuxPedidosGustavo` (70 filas, 1 campo): solo la usa `src/modules/TESTS.bas:95-96`, una lista manual de números de pedido que el módulo de pruebas cierra en lote. No es una tabla de negocio.

**Referencia muerta**: `TbDPDDocAnexada` (`src/classes/AnexoOperaciones.cls:32`, un `DELETE`). La tabla que existe en AGEDYS_DATOS se llama `TbDpDDocAnexada`, con las letras en otro orden. Es una referencia que en runtime falla.

**Código muerto**: `getdbCorreoApertura()` y `getdbSalidaAntigua()` (0 llamadas), `Proyecto.ID_PROYECTOSCalculado` (`src/classes/Proyecto.cls:298-316`, sin llamadores) y la consulta guardada `Consulta1` del front (`src/queries/Consulta1.sql`), que lee una tabla `TbAux*` de trabajo con las fechas literales `12/1/2024` y `1/31/2025`.

**Colisión de nombres**: `TbAuxExpedientesVivos` existe en el front (28 filas) y en AGEDYS_DATOS. El código la usa siempre contra el front (`src/classes/ConsultaOperaciones.cls:550`, `:556`), pero el nombre no dice de qué base se trata.

**Una tabla que el código de un front usa y el otro no**: `TbConsultasDetalle` y `TbConsultasPrincipal` están en AGEDYS_DATOS y las lee el front de AGEDYS ECO (`src/modules/constructor.bas:6699`, `:6740`, `:6783`, `:6836`). El catálogo es **compartido** entre las dos mitades de AGEDYS, y el front de AGEDYS ECO solo implementa una parte: `src/classes/ConsultaOperaciones.cls:53` rechaza la consulta «Consulta Prado completa» con el mensaje «No codificada para este aplicativo. Pruebe con AGEDYS técnico». La tabla `TbConsultasDetalle` tiene 23 consultas guardadas, con su `SQLConsulta`, su `SQLSimple` y el nombre de la tabla `TbAux` donde volcar el resultado.

## Propuestas para después del descubrimiento global

Ninguna se aplica: la puerta de descubrimiento global sigue cerrada (#1576). Los identificadores llevan prefijo `E` para no colisionar con los de Expedientes. Grupos según la [clasificación de la épica](../epica-coexistencia-postgresql.md#clasificación-de-tablas): (a) imprescindible antes de migrar, (b) recomendable si es barato, (c) aplazar a la refactorización final.

| Id | Propuesta | Grupo | Coste | Comprobación previa | Reversión |
|---|---|---|---|---|---|
| E1 | Unificar el acceso a `TbSalidas_RES`: el alta va por `getdb()` (que resuelve el vínculo de AGEDYS_DATOS) y el borrado por `getdbSalida()` (que abre Registro directamente), sobre la misma tabla física. Decidir un único camino y aplicarlo a los dos. | (a) | 5 referencias de SQL sobre `TbSalidas_RES`; 4 sitios (`src/classes/Salida.cls:88`, `:187-190`, `src/classes/SalidaOperaciones.cls:106-107`, `:166-169`) más 2 del instalador | Comprobar en producción que el vínculo de `AGEDYS_DATOS` y la ruta de `getdbSalida()` siguen resolviendo al mismo fichero, y reproducir el ciclo completo de una salida (alta, consulta por `NSalida`, borrado) antes de tocar nada | Restituir la línea original de cada sitio |
| E2 | Retirar la credencial literal de las siete cadenas de apertura (`src/modules/Variables Globales.bas:458`, `:485`, `:510`, `:543`, `:577`, `:611`, `:643`) y leerla de la configuración. | (a) | 7 literales | Es la misma credencial que abre todos los binarios del ecosistema. Comprobar que no queda ninguna otra copia del literal en el front y que el cambio no rompe ninguna de las siete aperturas | Restaurar las siete líneas originales |
| E3 | Corregir la referencia muerta `TbDPDDocAnexada` (`src/classes/AnexoOperaciones.cls:32`): la tabla se llama `TbDpDDocAnexada`. | (a) | 1 sitio | Reproducir en runtime el borrado de un anexo de expediente y confirmar que hoy falla | Restituir el nombre original |
| E4 | Corregir los dos generadores de identificador que apuntan al front: `src/classes/CestaAnexoOrdinaria.cls:35` y `src/classes/Proyecto.cls:304` pasan el argumento de base de datos vacío, y `DameID` cae a `CurrentDb()` (`src/modules/FUNCIONES UTILES.bas:531-532`). | (a) | 2 sitios; el primero es alcanzable desde `src/classes/CestaAnexoOrdinariaOperaciones.cls:120` | Reproducir en runtime el alta de un anexo de cesta ordinaria y comprobar que la lectura del contador falla hoy. Para el segundo, comprobar que la propiedad no tiene llamadores antes de decidir entre corregirla y borrarla | Restituir el argumento omitido |
| E5 | Declarar la clave primaria de `TbAuxGTVContratos`, la única tabla propia del front que el código usa y que no la tiene (136 filas, 11 campos). | (a) | 1 tabla | Contar duplicados antes de fijar la clave. El orden de la clave debe ser el de la consulta guardada `IDConsulta` 17, 19 y 20, que la puebla | Quitar la clave |
| E6 | Declarar en PostgreSQL las 43 tablas que el front necesita, con el dueño que se deduce aquí: 35 en el esquema de AGEDYS, 2 en el de Expedientes, 3 en el de la Lanzadera, 2 en el de AGEDO y 1 en el de Registro. | (a) | 43 tablas | Contrastar la lista con el grafo de vínculos de todos los backends (tarea T6) y con el resultado de la ficha de AGEDYS Técnico, para que las 30 tablas compartidas por las dos mitades no se dupliquen | Es una decisión de diseño, no un cambio en el legacy |
| E7 | Sustituir los dos generadores `Max + 1` (`src/modules/FUNCIONES UTILES.bas:534` y `src/modules/constructor.bas:2410`) por identidad o autonumérico. | (a) | 2 generadores, 7 tablas que los usan | Comprobar en la copia que no hay huecos ni duplicados en los identificadores actuales y definir el punto de corte de la numeración | Volver a la consulta `SELECT Max(…)` |
| E8 | Reescribir el SQL que no compila en PostgreSQL: 5 `DELETE * FROM`, 90 `Is Null`, 20 `Like '*'`, 4 literales de fecha con `#` y los 2 de `src/classes/ConsultaOperaciones.cls:943`, `:950`. | (a) | 121 sitios | Ejecutar en runtime los caminos afectados: borrado de anexo, de cesta, de correo y de documento, los filtros de la gestión de pedidos y la consulta con fechas del diálogo de «Consultas». Comprobar que `Like '*'` se traduce por un filtro vacío y no por un `ILIKE '%'` | Volver al SQL original de Access |
| E9 | Parametrizar el SQL que concatena valores: las fechas del diálogo (`src/classes/ConsultaOperaciones.cls:943`, `:950`) y el límite de `IDEnvio` (`src/modules/constructor.bas:2410`, `:2417`). | (a) | 4 sitios | Las fechas del diálogo son entrada de usuario. Comprobar que el resultado no cambia con comillas y signos de porcentaje en el texto | Volver a la concatenación |
| E10 | Migrar los 24 recordsets de tipo tabla a sentencias `SELECT` explícitas. | (a) | 24 sitios | Enumerar los 24 con su tabla y su base, y reproducir cada alta y cada borrado antes de reescribir | Volver a la forma con nombre de tabla |
| E11 | Mudar al backend las 7 tablas de trabajo de la función «Consultas» (`TbAuxProyectos`, `TbAuxPedidos`, `TbAuxFacturas`, `TbAuxFacturasCliente`, `TbAuxFacturasDetalleNoDYS`, `TbAuxGTVContratos`, `TbAuxExpedientesVivos`), o declararlas como temporales del front. | (a) | 7 tablas, 2146 filas en total | Son tablas de volcado y exportación: se borran y se repueblan en cada ejecución (`src/classes/ConsultaOperaciones.cls:177`, `:236`, `:314`, `:382`, `:429`, `:478`, `:550`). Comprobar que el backend puede alojar el mismo esquema y decidir si son de sesión o de aplicación | Devolver las tablas al front |
| E12 | Poner en cuarentena las 3 tablas del front sin uso (`TbAuxExpPrado`, `TbAuxParaProvisiones`, `TbAuxPedidosGustavo`), renombrándolas a `zzBAJA_<tabla>`. | (b) | 3 tablas, 283 filas | Confirmar que ninguna otra aplicación las usa y que no hay consumidores fuera de Access. `TbAuxExpPrado` la usa el front de AGEDYS Técnico, pero **su propia copia**, no la de ECO: la comprobación debe hacerse sobre cada front por separado | Quitar el prefijo `zzBAJA_` y restaurar las filas |
| E13 | Decidir el destino de la consulta guardada `Consulta1` del front (`src/queries/Consulta1.sql`), que lee `TbAuxFacturas` con las fechas literales `12/1/2024` y `1/31/2025`. | (b) | 1 consulta | Comprobar que no la invoca ningún formulario ni ningún módulo por su nombre, y decidir si se traslada al backend con la consulta de provisiones o se retira | Devolver la consulta al front |
| E14 | Retirar el código muerto: `getdbCorreoApertura()` (`src/modules/Variables Globales.bas:495`), `getdbSalidaAntigua()` (`:587`), `Proyecto.ID_PROYECTOSCalculado` (`src/classes/Proyecto.cls:298-316`) y `TESTS.bas` (`.CerrarPedidos` y sus dos ayudantes), que solo se usan desde un módulo de pruebas. | (b) | 2 abridores, 1 propiedad, 1 módulo | Confirmar que no hay ninguna otra ruta que los invoque (por ejemplo, una macro de la Lanzadera o una llamada por reflection) y que el front publicado no lleva un `.ini` que los habilite | Restituir el código original |
| E15 | Verificar la resolución de las rutas de `getdbSalida()`, `getdbSalidaAntigua()` y `getdbAgedo()`, que eligen la base de ruta contraria a la de sus cuatro hermanos (`src/modules/Variables Globales.bas:560-563`, `:596-599`, `:628-631`). | (b) | 3 abridores | Comprobar en el entorno de publicación real si la ruta relativa que construyen resuelve al fichero correcto. Si no resuelve, decidir si la corrección es una línea de código o un cambio de la estructura de publicación | Restituir la línea original |
| E16 | Normalizar la matriz de vínculos de la Lanzadera para AGEDYS ECO: 6 de los 71 vínculos declarados a `AGEDYS_DATOS` apuntan a tablas que ya no existen, y 72 tablas que sí existen no están declaradas. | (b) | 6 vínculos colgados | Confirmar que la matriz es la que usa la Lanzadera al publicar, y decidir si se corrige o si la publicación deja de usarla para este front. Requiere el grafo global (tarea T6) | Restituir la matriz original |
| E17 | Fijar la semántica de comparación de cadenas: `Option Compare Database` está en 112 de los 113 módulos, clases y code-behind de formulario (falta en `src/modules/JsonConverter.bas`). | (c) | 113 ficheros | Enumerar las comparaciones de cadenas que dependen de la regla de Access (acentos, mayúsculas y ordenación) y decidir la collation de PostgreSQL antes de la reescritura | Restablecer el valor de la opción en el proyecto de Access |
| E18 | Decidir el destino de `Option Compare Database` y de los 215 recordsets sin opciones de apertura como una sola política de acceso a datos, porque en APGEDYS las dos mitades van a compartir tabla. | (c) | 215 llamadas | Levantar el catálogo de SQL por tabla de `docs/07-migracion/fase-0/controles-conectados.md` y extenderlo con AGEDYS ECO, para acotar el trabajo antes de la reescritura | La decisión no cambia el legacy |

Lo que **no** se toca ahora: las 43 tablas que el front necesita, los vínculos que `AGEDYS_DATOS` declara hacia otros backends, las 3 tablas de la Lanzadera, las 2 de AGEDO y las 7 de staging del front más allá de la propuesta E11, y la ruta `\\datoste\…` de producción, salvo la propuesta E2 sobre la credencial.

## Anexo: controles con origen de datos

Ninguno. Los 39 `RowSourceType` del front son `"Value List"` y las 39 asignaciones en tiempo de ejecución que rellenan esos controles tampoco contienen SQL: todas son listas de cabeceras de columnas, y tres están vacías (`cmb.RowSource = ""`).

| Literal | Formulario y línea | Contenido |
|---|---|---|
| `IDAnexo;TÍTULO` | `src/forms/Form_FormAnexarDocumentacion.cls:225` | Cabeceras |
| `ID;FInicial;FFinal` | `src/forms/Form_FormCalendarioProvisiones.cls:59` | Cabeceras |
| `IDTipoAnexo;ANEXO;EXISTE` | `src/forms/Form_FormCestaOrdinariaAnexos.cls:55` | Cabeceras |
| `IDConsulta;Consulta;Descripción` | `src/forms/Form_FormConsultas.cls:24`, `:77`, `:110` | Cabeceras |
| `IDCategoria;Categoría;Descripción` | `src/forms/Form_FormConsultas.cls:27` | Cabeceras |
| `ID;EXPEDIENTE;CÓDIGO` | `src/forms/Form_FormExpedienteCodigosGestion.cls:107` | Cabeceras |
| `NFACTURA;F.EMISIÓN;BASE IMP.;T.IMP;EXENTO;IMPORTE;CODEXP` | `src/forms/Form_FormFacturaAlClienteGestion.cls:230` | Cabeceras |
| `IDFactura;DOCUMENTO;IMPORTE;DPD;DESCRIPCIÓN;SUMINISTRADOR;TIP` | `src/forms/Form_FormFacturasAProveedorGestion.cls:131`; `src/forms/Form_FormTareasSobreFacturasPendientes.cls:36`, `:524` | Cabeceras |
| `ID;NPEDIDO;NFactura;F.Grab;Importe;Suministrador` | `src/forms/Form_FormFacturasNODYSGestion.cls:407` | Cabeceras |
| `ID;Expediente;Título` | `src/forms/Form_FormGTV.cls:530`, `:605` | Cabeceras |
| `IDGTV;CESTA;DESCRIPCIONGTV;FECHAGRABACION;FECHACESTA` | `src/forms/Form_FormGTVGestion.cls:545`; `src/forms/Form_FormTareasSobreCestasDeComprasGenerales.cls:31` | Cabeceras |
| `IDGTVCONTRATO;CONTRATO;POS;IMPORTE;SUMINISTRADOR` | `src/forms/Form_FormGTVGestion.cls:549`, `:671` | Cabeceras |
| `CONTRATO;POS;DPD;DESCRIPCIÓN` | `src/forms/Form_FormGTVGestion.cls:554`, `:675`, `:851` | Cabeceras |
| `NCARTA;NPEDIDO;TEXTO;RECT` | `src/forms/Form_FormPedidoAdjudicacion.cls:41` | Cabeceras |
| `IDEnvio;NSALIDA;FECHASALIDA;IDCORREO;EMAIL` | `src/forms/Form_FormPedidoAdjudicacion.cls:42`, `:467`, `:488`, `:507` | Cabeceras |
| `DPD;Descripción;Sumin;Cesta;NPedido` | `src/forms/Form_FormPedidosGestion.cls:55` | Cabeceras |
| `IDFactura;NDOCUMENTO;NFACTURA;NPEDIDO;IMPORTE` | `src/forms/Form_FormProyectoDetalleFacturas.cls:90` | Cabeceras |
| `NAC.;CIF;Suministrador;email` | `src/forms/Form_FormSuministradores.cls:60`; `src/forms/Form_FormSuministradoresGestion.cls:64` | Cabeceras |
| `DPD;TIPO;IMPORTE;Descripción;Sumin` | `src/forms/Form_FormTareasSobreProyectosPendientes.cls:71` | Cabeceras |
| `""` | `src/forms/Form_FormFacturaAlCliente.cls:465`; `src/forms/Form_FormFacturaAlClienteGestion.cls:428`; `src/forms/Form_FormTareasSobreFacturasPendientes.cls:114` | Control vaciado |

El SQL de la aplicación no está en los controles: se construye en código, en variables, y es el que hay que reescribir. Los dos formularios que filtran por expediente, proyecto y suministrador (`Form_FormPedidosGestion` y `Form_FormSuministradoresInformeValoracion`) no lo hacen con `RowSource`, sino con `WHERE` concatenado en `m_Where…` (`src/forms/Form_FormPedidosGestion.cls:704`, `:711`, `:716`; `src/forms/Form_FormSuministradoresInformeValoracion.cls:80`).

## Qué queda en la parte económica de APGEDYS

Por la decisión 13 de la épica (issue #1623), AGEDYS Técnico y AGEDYS ECO se sustituyen por una sola aplicación web, APGEDYS, con vista de técnico o de gestor económico según el usuario que entre. Esta ficha inventaría la parte económica; lo que sigue es el reparto que se deduce de la medición.

### Lo que es económico y solo económico

| Elemento | Dónde vive hoy | Evidencia |
|---|---|---|
| Facturas de proveedor y de cliente | `TbFacturasDetalle`, `TbFacturasPrincipal`, `TbFacturasCliente`, `TbFacturasDetalleNoDYS`, `TbVisadoFacturas_Nueva` en `AGEDYS_DATOS` | 29, 13, 10, 10 y 14 referencias de SQL en el `src/` de ECO |
| Visado económico y aceptación | `src/classes/FacturaVisadoOperaciones.cls`, `src/classes/FacturaOperaciones.cls` | `VisarPorEconomia`, `AceptacionEconomica` (`src/classes/FacturaOperaciones.cls:59`) |
| Pedidos y su adjudicación | `TbNPedido`, `TbAdjudicacionesOfertasCartas`, `TbAdjudicacionesEnvioAlSuministrador`, `TbSolicitudesOfertasPrevias`, `TbSolicitudesEnvioAlSuministrador` | 46, 23, 11, 3 y 2 referencias |
| Compras generales | `TbGTV`, `TbGTVContratos`, `TbGTVContratosExpedientes`, `TbDPDComiteOfertas`, `TbDPDOfertaSuministrador` | 15, 17, 5, 3 y 1 referencias |
| Valoración de suministradores | `TbSuministradoresSAP`, `TbSuministradoresPuntuacionConDetalle` | 14 y 3 referencias |
| Provisiones | `TbCalendarioProvisiones`, `TbCalendarioProvisionesDetalle` | 5 y 4 referencias |

**Cinco tablas propias de `AGEDYS_DATOS` las usa solo la parte económica**: `TbCalendarioProvisiones`, `TbCalendarioProvisionesDetalle`, `TbConsultasPrincipal`, `TbFacturasDetalleNoDYS` y `TbAuxExpedientesVivos` (esta última, con el agravante de que el nombre existe también en el front y en el backend). Esta ficha es la que inventaría la vista de gestor económico; la de AGEDYS Técnico, la vista de técnico.

### Lo que es técnico

Diez tablas propias de `AGEDYS_DATOS` las referencia el front de AGEDYS Técnico y no el de AGEDYS ECO: `TbBloqueos`, `TbDpDPrepCartaOfAcepDetalle`, `TbGTV_Temp`, `TbLog`, `TbRequisitosSTG`, `TbSuministradoresNotasReferencia`, `TbTemp_PedidosEstado`, `TbTipoDocumento`, `TbTmp_DPDs` y `TbTmp_Facturas`. Ninguna de las dos mitades las posee: son del front de AGEDYS Técnico y habría que inventariarlas en su ficha.

### Lo que las dos mitades comparten

**30 tablas propias de `AGEDYS_DATOS`** las referencian los dos fronts: `TbAdjudicacionesEnvioAlSuministrador`, `TbAdjudicacionesOfertasCartas`, `TbCestasAnexosOrdinaria`, `TbConsultasDetalle`, `TbDPDCartaSolicitudOfertaDatos`, `TbDPDComiteOfertas`, `TbDPDOfertaSuministrador`, `TbDpDAsuntoDefensa`, `TbDpDCartaSolicitudOfertaDatosFijos`, `TbDpDCartaSolicitudOfertaPlantillasSegunReqBasicos`, `TbDpDDocAnexada`, `TbDpDInformeCondicionamiento`, `TbExpedientesAnualidades`, `TbExpedientesCodigoCompras`, `TbFacturasCliente`, `TbFacturasDetalle`, `TbFacturasPrincipal`, `TbGTV`, `TbGTVContratos`, `TbGTVContratosExpedientes`, `TbNPedido`, `TbOrdinariasTipoAnexos`, `TbProductos`, `TbProyectos`, `TbSolicitudesEnvioAlSuministrador`, `TbSolicitudesOfertasPrevias`, `TbSuministradoresPuntuacionConDetalle`, `TbSuministradoresSAP`, `TbVisadoFacturas_Nueva` y `TbVisadosGenerales`.

Hay que añadir tres dependencias más que ninguna de las dos mitades posee, pero que las dos tocan:

- `TbExpedientes` y `TbExpedientesResponsables`, propiedad de Expedientes.
- `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` y `TbAplicacionesAperturas`, propiedad de la Lanzadera.

**Y un catálogo compartido, que es el hallazgo de mayor alcance para APGEDYS**: `TbConsultasPrincipal` y `TbConsultasDetalle` están en `AGEDYS_DATOS` y las lee el front de AGEDYS ECO (`src/modules/constructor.bas:6699`, `:6740`, `:6783`, `:6836`). El catálogo de 23 consultas guardadas es **de las dos mitades**: cada front implementa las suyas y las demás las rechaza por su nombre. En APGEDYS, con una sola aplicación y una sola vista que cambia según el rol, ese catálogo deja de necesitar la partición: o se ejecuta entero, o se documenta qué consulta es de qué vista.

### Qué tablas quedarían en un backend compartido

Con lo medido, el reparto que se deduce es:

| Backend | Tablas | De quién son |
|---|---|---|
| `AGEDYS_DATOS.accdb` → esquema AGEDYS en PostgreSQL | 43 que la parte económica necesita, de las cuales 30 son compartidas con la parte técnica | Compartidas |
| `Registro_Ent_Salida_Datos.accdb` → esquema Registro | `TbSalidas_RES` (y `TbEntradas_RES`, que el front no usa) | De Registro, con AGEDYS ECO como escritor secundario |
| `Expedientes_datos.accdb` → esquema Expedientes | `TbExpedientes`, `TbExpedientesResponsables` | De Expedientes |
| `Lanzadera_Datos.accdb` → esquema Lanzadera | `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`, `TbAplicacionesAperturas` | De la Lanzadera, que es la autoridad de identidad y permisos |
| `Correos_datos.accdb` | `TbCorreosEnviados` | De Correos |
| `AGEDO20_Datos.accdb` | `TbDocumentos`, `TbDocumentosID` | De AGEDO, fuera del alcance de la migración: dependencia, no objetivo |

El backend propio de la parte económica **no existe**: no hay ningún `AGEDYS_ECO_Datos.accdb` en `C:\00repos\datos`, ni ninguna tabla que la Lanzadera atribuya a una aplicación «AGEDYS ECO» en la columna `NombreArchivoDatos` (la fila `IDAplicacion` 2 declara «AGEDYS_DATOS.accdb», el mismo fichero que declara la fila 3, AGEDYS TEC). Las 10 tablas propias del front de AGEDYS ECO son de staging y no de negocio, y las 43 tablas de negocio que la parte económica necesita ya viven en un backend compartido con la parte técnica. **No hay una tabla de la parte económica que necesite mudarse a un backend propio**: lo que hay que decidir es si el backend de AGEDYS se mantiene como uno solo para las dos mitades de APGEDYS o si se parte, y la medición dice que partirlo no aporta nada y rompe 30 tablas compartidas.

## Consultas de reproducción

Sobre `C:\00repos\datos\Lanzadera_Datos.accdb`, en modo de solo lectura:

```sql
SELECT IDAplicacion, NombreAplicacion, NombreEjecutable, NombreArchivoDatos, EnPruebas
  FROM TbAplicaciones WHERE IDAplicacion IN (2, 3, 10, 12);

SELECT IDBBDD, COUNT(*) AS Tablas
  FROM (SELECT DISTINCT IDBBDD, NombreTabla FROM TbTablasAVincular WHERE IDAplicacion=2)
  GROUP BY IDBBDD;

SELECT IDBBDD, NombreTabla FROM TbTablasAVincular
  WHERE IDAplicacion=2 AND IDBBDD=10;

SELECT IDConsulta, Compleja, TbAux, NombreConsulta FROM TbConsultasDetalle ORDER BY IDConsulta;
```

Sobre `C:\00repos\datos\AGEDYS_DATOS.accdb`:

```sql
SELECT COUNT(*) AS n FROM TbProyectos;
SELECT COUNT(*) AS n FROM TbSalidas_RES;      -- error: la tabla es un vínculo, no existe localmente
```

Sobre `C:\00repos\datos\Registro_Ent_Salida_Datos.accdb`:

```sql
SELECT COUNT(*) AS n FROM TbSalidas_RES;      -- 4053
SELECT COUNT(*) AS n FROM TbEntradas_RES;     -- 635
SELECT COUNT(*) AS n FROM TbSalidas_RES_Indexacion;  -- 1063
```

Extracción de ERD, sobre copia temporal y en solo lectura:

```powershell
$env:ACCDB_PASSWORD = <se pasa como variable de entorno, no se escribe en ningún fichero>
.\scripts\dysflow-erd\extract-erd.ps1 -Path "C:\00repos\codigo\00_AGEDYS_ECO\Gestion Proyectos Eco 30.accdb" -OutDir "<salida temporal>"
```

## Lo que no se pudo verificar

1. **El estado de UAT.** No hay ninguna evidencia: ni en el repositorio (no hay `docs/`), ni en las ramas (no hay RC1), ni en las etiquetas de git. La ficha se apoya en la ausencia de evidencia, no en un documento de UAT.
2. **La versión «Ver. Servidor 2026-002».** La Lanzadera la compone en tiempo de ejecución desde el `.ini` del servidor, que no está versionado y cuyo share no es alcanzable desde aquí. Lo verificado es el catálogo: `TbAplicaciones` fila 2, con `EnPruebas` «No». El `.ini` del front tampoco está en el repositorio, así que la versión que el front muestra al usuario (`src/forms/Form_Form0BDOpciones.cls:390`) tampoco es comprobable.
3. **Si el SQL exportado coincide con el binario.** No se ha ejecutado `verify_code` de dysflow sobre los 40 formularios de AGEDYS ECO. La comparación con el binario se ha hecho para lo que el ERD y el volcado directo de `TableDefs` dan: el número de tablas, sus vínculos, sus claves y su número de filas.
4. **La resolución real de las rutas de `getdbSalida()`, `getdbSalidaAntigua()` y `getdbAgedo()`.** No se ha ejecutado la aplicación. Que la rama «No» construya una ruta relativa es una lectura del código; que esa ruta resuelva o no depende de la estructura de publicación, que no se ha inspeccionado.
5. **El comportamiento en runtime de los defectos medidos.** `TbDPDDocAnexada`, los dos `DameID` sin base de datos y la doble vía a `TbSalidas_RES` son conclusiones de lectura de código y de esquema, no fallos observados.
6. **Los consumidores de las 3 tablas huérfanas del front y de las 67 tablas de `AGEDYS_DATOS` que no referencia ninguna de las dos mitades de AGEDYS.** Exige el grafo de vínculos de todos los backends (tarea T6) y saber si la web o un proceso programado las leen.
7. **La tabla `TbSalidas_RES_Indexacion` en el front de AGEDYS ECO.** La matriz de la Lanzadera declara el vínculo y el código no la referencia, pero el front publicado en el servidor podría traer vínculos creados por la Lanzadera que la copia del repositorio no tiene. El ERD del front de este repositorio no declara ningún vínculo, lo que hace pensar que la publicación los crea.
8. **El catálogo `TbConsultasDetalle` en producción.** Las 23 consultas y su reparto entre las dos mitades se han medido sobre la copia local de `AGEDYS_DATOS.accdb` (23-10-2025 según la marca de tiempo del fichero), que puede no coincidir con producción.
