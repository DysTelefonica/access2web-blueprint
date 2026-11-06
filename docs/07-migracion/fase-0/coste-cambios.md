# Coste de los cambios por tabla y por columna

Entregable «coste de cada cambio estructural» y parte del entregable «por tabla» de la fase 0 de la [épica](../epica-coexistencia-postgresql.md#fase-0-inventario). Para cada tabla que usa alguna aplicación del alcance, y para cada una de sus columnas, cuenta cuántas líneas de código la nombran, sumando todos los fronts que la usan. Es el criterio de coste de la fase 1: renombrar, partir o cambiar el tipo de una tabla o columna cuesta, como mínimo, tocar esas líneas.

Cubre las 299 tablas de los backends de la [copia de referencia](copia-referencia.md) que tienen consumidores según la [clasificación de las tablas](clasificacion-tablas.md), con las mismas ramas base y los mismos commits. Las huérfanas no se cuentan: van a cuarentena (`CLS-H1`), no se cambian.

## Método

- **Qué tabla es cada nombre.** No se busca el nombre físico de la tabla, sino los nombres con los que cada aplicación la usa, tal como los resolvió la clasificación: su vínculo, un alias de vínculo (Expedientes llama `TbProyectos_Riesgos` a la `TbProyectos` de Gestión de Riesgos) o la tabla propia. Así no se cuentan dos veces las líneas de una tabla que existe con el mismo nombre en dos backends, como `TbProyectos` en `AGEDYS_DATOS` y en `Gestion_Riesgos_Datos`.
- **Qué se lee.** El VBA de módulos, clases, formularios e informes sin comentarios ni módulos de test, y las propiedades de origen de datos de los formularios (`RecordSource`, `RowSource` y `ControlSource`).
- **Tabla.** Líneas que nombran la tabla en las aplicaciones que la usan, y ficheros en los que aparece.
- **Columna.** Líneas que nombran la columna como campo (`!Columna`, `.Columna`, `[Columna]`) o dentro del texto de un SQL, contadas solo en los ficheros que nombran su tabla.
- **Tipos y reglas.** Las columnas, sus tipos, reglas de validación y valores por defecto se leen de la copia de referencia con DAO en solo lectura, sobre copias temporales (`inventory-data.ps1 -Mode Columns`).
- **Se reproduce** con `collect-usage.ps1`, `inventory-data.ps1 -Mode Columns`, `classify-tables.py`, que escribe también `resolved.tsv` con la resolución de nombres, y `table-cost.py`, todos en `scripts/dysflow-erd/`. El último escribe `table-cost.tsv`, `column-cost.tsv` y `column-risks.tsv`, con todas las tablas y columnas.

## Coste por backend

| Backend | Tablas | Columnas | Líneas de código |
|---|---:|---:|---:|
| `Gestion_Brass_Gestion_Datos` | 50 | 398 | 3 319 |
| `AGEDYS_DATOS` | 45 | 578 | 3 026 |
| `GestionContratos_datos` | 19 | 256 | 2 583 |
| `Gestion_Riesgos_Datos` | 34 | 297 | 1 310 |
| `Expedientes_datos` | 38 | 394 | 1 280 |
| `NoConformidades_Datos` | 21 | 203 | 865 |
| `Lanzadera_Datos` | 14 | 127 | 674 |
| `HPST` | 15 | 220 | 346 |
| `Solicitudes_HPS_datos` | 9 | 107 | 338 |
| `Registro_Ent_Salida_Datos` | 6 | 56 | 291 |
| `condor_datos` | 14 | 249 | 191 |
| `Correos_datos` | 1 | 17 | 109 |
| `Registro_Datos` | 2 | 29 | 75 |
| `AGEDO20_Datos` | 2 | 37 | 47 |
| `Seguridad_datos` | 26 | 314 | 29 |
| `Control_Cambios_datos` | 3 | 19 | 13 |
| **Total** | **299** | **3 301** | **14 496** |

Las 25 tablas con consumidores que no nombra ninguna línea de VBA se usan por otra vía: 21 de `Seguridad_datos`, que Seguridad elige desde sus datos (`Tb0TipoRegistros.NombreTabla`) y trata con un código genérico; `TbResponsablesExpedientes` (`AGEDYS_DATOS`), `ListaPrevia` y `TbExpedientes_antes` (`Expedientes_datos`), que solo leen consultas guardadas; y `TbUsuariosAplicacionesTareas` (`Lanzadera_Datos`), que vincula Tareas, cuyo código no se analiza. Su coste está en esas consultas y en ese código genérico, no en estas cifras.

## Las veinte tablas más caras

| Backend | Tabla | La usan | Líneas | Ficheros |
|---|---|---|---:|---:|
| `AGEDYS_DATOS` | `TbProyectos` | AGEDYS ECO, AGEDYS Técnico, Expedientes, Registros, Tareas | 792 | 73 |
| `GestionContratos_datos` | `AnexoFases` | Contratos | 508 | 12 |
| `Expedientes_datos` | `TbExpedientes` | AGEDYS ECO, AGEDYS Técnico, Condor, Expedientes, Gestión de Riesgos, HPS, No Conformidades, Solicitudes HPS, Tareas | 497 | 60 |
| `Gestion_Brass_Gestion_Datos` | `TbEventos` | Brass | 460 | 66 |
| `Gestion_Brass_Gestion_Datos` | `TbEquipos` | Brass | 370 | 45 |
| `Lanzadera_Datos` | `TbUsuariosAplicaciones` | AGEDYS ECO, AGEDYS Técnico, Brass, Condor, Contratos, Expedientes, Gestión de Riesgos, HPS, Lanzadera, No Conformidades, Registros, Seguridad, Solicitudes HPS, Tareas | 360 | 48 |
| `Gestion_Riesgos_Datos` | `TbRiesgos` | Gestión de Riesgos, No Conformidades | 350 | 26 |
| `Gestion_Brass_Gestion_Datos` | `TbActividades` | Brass | 331 | 42 |
| `Gestion_Riesgos_Datos` | `TbProyectosEdiciones` | Expedientes, Gestión de Riesgos, No Conformidades | 318 | 23 |
| `Gestion_Brass_Gestion_Datos` | `TbMaterial` | Brass | 316 | 26 |
| `GestionContratos_datos` | `Contratos` | AGEDYS ECO, Contratos | 308 | 22 |
| `AGEDYS_DATOS` | `TbNPedido` | AGEDYS ECO, AGEDYS Técnico, Tareas | 278 | 36 |
| `AGEDYS_DATOS` | `TbFacturasDetalle` | AGEDYS ECO, AGEDYS Técnico, Tareas | 253 | 25 |
| `AGEDYS_DATOS` | `TbSuministradoresSAP` | AGEDYS ECO, AGEDYS Técnico, Tareas | 237 | 37 |
| `NoConformidades_Datos` | `TbNoConformidades` | Condor, Expedientes, Gestión de Riesgos, No Conformidades | 223 | 20 |
| `Gestion_Riesgos_Datos` | `TbProyectos` | AGEDYS Técnico, Expedientes, Gestión de Riesgos, No Conformidades | 192 | 19 |
| `GestionContratos_datos` | `AnexoSA` | Contratos | 188 | 8 |
| `GestionContratos_datos` | `AnexoSCS` | Contratos | 184 | 8 |
| `GestionContratos_datos` | `AnexoSFS` | Contratos | 184 | 8 |
| `GestionContratos_datos` | `AnexoSFT` | Contratos | 184 | 8 |

Las tablas más caras son las del núcleo de cada aplicación, y casi todas son de una sola aplicación o de su familia (AGEDYS y Expedientes). Las compartidas por muchas aplicaciones, como `TbUsuariosAplicaciones` (catorce consumidores), cuestan menos líneas por aplicación pero obligan a coordinar las versiones de todas.

## Las quince columnas más caras

| Backend | Tabla | Columna | Líneas |
|---|---|---|---:|
| `Expedientes_datos` | `TbExpedientes` | `IDExpediente` | 1 000 |
| `Expedientes_datos` | `TbExpedientesResponsables` | `IdExpediente` | 732 |
| `Expedientes_datos` | `TbExpedientesSuministradores` | `IDExpediente` | 596 |
| `AGEDYS_DATOS` | `TbProyectos` | `CODPROYECTOS` | 581 |
| `Expedientes_datos` | `TbExpedientesRACS` | `IDExpediente` | 538 |
| `Gestion_Brass_Gestion_Datos` | `TbEventos` | `IDEvento` | 536 |
| `Expedientes_datos` | `TbExpedientesAnualidades` | `IDExpediente` | 502 |
| `Expedientes_datos` | `TbExpedientesPECAL` | `IDExpediente` | 502 |
| `AGEDYS_DATOS` | `TbProyectos` | `IDExpediente` | 462 |
| `Gestion_Brass_Gestion_Datos` | `TbActividades` | `IDEvento` | 439 |
| `Expedientes_datos` | `TbExpedientesConEntidades` | `IDExpediente` | 416 |
| `Expedientes_datos` | `TbExpedientesComerciales` | `IDExpediente` | 414 |
| `Gestion_Riesgos_Datos` | `TbRiesgos` | `IDEdicion` | 403 |
| `Gestion_Brass_Gestion_Datos` | `TbEventos` | `BUI` | 380 |
| `Expedientes_datos` | `TbExpedientesCPVs` | `IDExpediente` | 377 |

Las columnas más nombradas son claves de relación (`IDExpediente`, `IDEvento`, `CODPROYECTOS`). Cambiar el nombre o el tipo de una de ellas es el cambio más caro posible y no está en ninguna propuesta; las propuestas tocan sobre todo tablas enteras (cuarentena) y claves primarias que faltan.

De las 3 301 columnas, **706 no aparecen en ninguna línea** de los ficheros que nombran su tabla. Es la lista de candidatas a dejar de llevarse a PostgreSQL, pero no una prueba: una columna se puede leer con `SELECT *`, en una consulta guardada o desde un formulario conectado. Se revisan tabla a tabla en la fase 1, con la lista de `column-cost.tsv`.

## Tipos, reglas y valores por defecto

Sobre las 3 301 columnas de las tablas que migran:

- **Ningún tipo problemático.** No hay columnas multivalor, de datos adjuntos ni de objeto OLE. Los tipos son texto corto (1 701), entero largo (510), fecha (524), texto largo (324), sí/no (124), doble (84), entero (32), simple (1) y byte (1).
- **Ninguna regla de validación** en el nivel de columna.
- **173 columnas con valor por defecto.** Solo dos usan funciones de Access que PostgreSQL no conoce: `TbCacheNCProyecto.FechaCache` (`Date()`) y `TbLogCache.FechaOperacion` (`Now()`), las dos en `NoConformidades_Datos`. Se traducen a `CURRENT_DATE` y `now()` al crear el esquema.

## Límites

- **Es lectura de código, no ejecución.** Una línea en un procedimiento que nadie llama cuenta igual.
- **El recuento de columnas es una cota superior.** Un nombre de columna común (`IDExpediente`) cuenta en todos los ficheros que nombran la tabla, aunque la línea se refiera a otra tabla con la misma columna.
- **No cuenta las consultas guardadas** ni el código de Tareas, que sigue en Access y no se analiza (decisión 14). El SQL de las consultas está en el [anexo de controles conectados](controles-conectados.md).
- **No mide las macros de datos** de las tablas: DAO no las expone. Se comprueban en el ensayo de la fase 3, al exportar el esquema.
