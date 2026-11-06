# Patrones de riesgo del VBA

Entregable «por front» de la fase 0 de la [épica](../epica-coexistencia-postgresql.md#fase-0-inventario): el código de cada aplicación que no sobreviviría tal cual a que sus datos pasen a PostgreSQL por ODBC. Cubre las trece aplicaciones de la [clasificación de las tablas](clasificacion-tablas.md), con las mismas ramas base y los mismos commits. Contratos no migra, pero sigue viva y se mide igual.

Seis fichas ya medían estos patrones a mano (HPS, Condor, Gestión de Riesgos, Seguridad, Solicitudes HPS y AGEDYS ECO). Este documento los mide en todas con un solo método, que se puede volver a ejecutar en cualquier momento. Cuando una ficha y este recuento no coinciden, manda este recuento.

## Método

- **Qué se lee.** El VBA de los módulos y clases, y el código de los formularios e informes (la sección `CodeBehindForm`). Se quitan los comentarios y los módulos de test, y las líneas partidas con ` _` se unen en una sola.
- **Dónde se busca.** La mayoría de los patrones está en el SQL, así que se buscan dentro de los literales de texto de cada línea.
- **Se reproduce** con `scripts/dysflow-erd/collect-usage.ps1` y `scripts/dysflow-erd/risk-patterns.py`. El segundo escribe `patterns.tsv`, con fichero, línea y código de cada coincidencia, y `summary.tsv`, con el recuento.

| Patrón | Qué se cuenta | Por qué es un riesgo |
|---|---|---|
| Generador `Max + 1` | Procedimientos que leen un máximo (`Max(` en el SQL o `DMax`) y le suman 1; no cuentan los contadores `x = x + 1`. Se cuenta el procedimiento, una vez. | Con varios usuarios a la vez dos registros reciben el mismo identificador: ya pasó con `IDCorreo` (decisión 22). En PostgreSQL el identificador lo asigna la base. |
| Recordset que edita sin `dbSeeChanges` | `OpenRecordset` sin `dbSeeChanges` cuyo recordset se edita después (`.AddNew` o `.Edit`, también dentro de un `With`) en el mismo procedimiento. | Con SQL Server, Access exige `dbSeeChanges` para editar tablas con columna de identidad (error 3622). Si con psqlODBC pasa lo mismo se comprueba en el ensayo de la fase 3. |
| `OpenRecordset` sin `dbSeeChanges` | Todas las aperturas sin esa opción, editen o no. | Cota superior del anterior: las de solo lectura no fallan. |
| Literal de fecha con `#` | `#` como delimitador de fecha dentro del SQL. No cuentan las máscaras de número (`"#,##0.00"`). | La sintaxis `#mm/dd/aaaa#` es de Access; PostgreSQL usa `'aaaa-mm-dd'`. Las consultas que pasan tal cual al servidor fallan. |
| `Date()`, `Now()` o `Time()` en el SQL | Las funciones de fecha de Access escritas dentro del texto del SQL. | El servidor no las conoce (`CURRENT_DATE`, `now()`), y la hora pasa a ser la del servidor, no la del puesto. |
| `LIKE` con comodines de Access | `LIKE` con `*` o `?` dentro del SQL. | PostgreSQL usa `%` y `_`: el filtro no devuelve nada o devuelve otra cosa. Hay muchos `Like '*'`, que significan «cualquier valor» y se pueden quitar. |
| `Seek` o `dbOpenTable` | Su uso en el código. | No existen sobre tablas vinculadas por ODBC. |
| `Option Compare Database` | Módulos con esa cabecera, sobre el total de módulos. | Compara texto sin distinguir mayúsculas; PostgreSQL sí las distingue por defecto. |

## Recuento por aplicación

| Aplicación | `Max + 1` | Editan sin `dbSeeChanges` | `OpenRecordset` sin `dbSeeChanges` | Fecha con `#` | Fechas de Access en SQL | `LIKE` con `*` | `Seek` | `Option Compare Database` |
|---|---:|---:|---:|---:|---:|---:|---:|---|
| Lanzadera | 1 | 15 | 59 | 7 | 2 | 4 | 0 | 76 de 78 |
| Expedientes | 1 | 39 | 244 | 8 | 7 | 16 | 0 | 181 de 183 |
| Condor | 11 | 50 | 92 | 1 | 2 | 1 | 0 | 183 de 189 |
| Seguridad | 2 | 10 | 43 | 2 | 0 | 12 | 0 | 121 de 122 |
| HPS | 7 | 32 | 168 | 6 | 8 | 7 | 0 | 114 de 115 |
| Solicitudes HPS | 1 | 12 | 90 | 13 | 2 | 5 | 0 | 103 de 105 |
| Gestión de Riesgos | 3 | 64 | 336 | 30 | 14 | 27 | 0 | 261 de 276 |
| No Conformidades | 3 | 52 | 248 | 23 | 18 | 4 | 0 | 178 de 180 |
| Registros | 2 | 14 | 48 | 1 | 1 | 65 | 0 | 30 de 31 |
| Brass | 2 | 108 | 480 | 49 | 27 | 122 | 0 | 213 de 214 |
| AGEDYS Técnico | 2 | 79 | 438 | 13 | 5 | 68 | 0 | 149 de 150 |
| AGEDYS ECO | 1 | 47 | 219 | 8 | 1 | 29 | 0 | 151 de 152 |
| Contratos | 1 | 40 | 255 | 1 | 1 | 8 | 0 | 84 de 85 |
| **Total** | **37** | **562** | **2 720** | **162** | **88** | **368** | **0** | |

Lectura:

- **`Seek` y `dbOpenTable` no aparecen en ninguna aplicación.** Ese riesgo de la épica queda descartado.
- **Casi todos los módulos llevan `Option Compare Database`.** La diferencia de mayúsculas no se arregla módulo a módulo: se decide una vez para todos, con la intercalación de la base o comparando en minúsculas en las consultas.
- **Brass concentra el SQL de Access**: 49 fechas con `#`, 27 funciones de fecha y 122 `LIKE` con `*`. Es la penúltima en el [orden de paso a la web](orden-cortes-web.md), así que tiene margen.
- **Los recordsets que editan son 562, no 2 720.** Ese es el trabajo real si psqlODBC exige `dbSeeChanges`.

## Generadores `Max + 1`

Los 37 procedimientos. Cada uno se sustituye por un identificador que asigna la base: autonumérico en Access y, después, identidad en PostgreSQL.

- **Lanzadera** (1): `DameID` (`modules/Funciones Genreales.bas:455`).
- **Expedientes** (1): `DameID` (`modules/FUNCIONES UTILES.bas:2863`).
- **Condor** (11): `getSiguienteIDAdjunto` (`modules/AdjuntoRepositorio.bas:37`), `getSiguienteIDCorreo` (`modules/CorreoRepositorio.bas:11`), `getSiguienteIDDatosCDCA` (`modules/DatosCDCARepositorio.bas:12`), `getSiguienteIDDatosCDCASUB` (`modules/DatosCDCASUBRepositorio.bas:5`), `getSiguienteIDDatosPC` (`modules/DatosPCRepositorio.bas:74`), `getSiguienteIDDatosPCSUB` (`modules/DatosPCSUBRepositorio.bas:75`), `LogToTable` (`modules/ErrorLogger.bas:21`), `DameID` (`modules/FUNCIONES UTILES.bas:516`), `getSiguienteIDLogCambios` (`modules/LogCambioRepositorio.bas:33`), `getSiguienteIDLogErrores` (`modules/LogErrorRepositorio.bas:43`), `getSiguienteOrdinalParaClave` (`modules/SolicitudRepositorio.bas:96`).
- **Seguridad** (2): `getID` (`modules/Constructor.bas:751`), `DameID` (`modules/Funciones Generales.bas:160`).
- **HPS** (7): `Phv_ResolveDestinationId` (`classes/UsuarioLifecycleTransactionCoordinator.cls:293`), `getID` (`classes/UsuarioSICA.cls:321`), `getID` (`modules/Constructor.bas:2504`), `DameIDUsuario` (`modules/Constructor.bas:2568`), `DameIDAnexo` (`modules/Constructor.bas:2631`), `DameIDObservaciones` (`modules/Constructor.bas:2715`), `DameID` (`modules/Funciones Generales.bas:2776`).
- **Solicitudes HPS** (1): `DameID` (`modules/FUNCIONES UTILES.bas:760`).
- **Gestión de Riesgos** (3): `getCodigo` (`classes/NC.cls:483`), `ControlCambios_Cache_SiguienteBuildId` (`modules/ControlCambios.bas:1443`), `DameID` (`modules/Funciones Generales.bas:2362`).
- **No Conformidades** (3): `DameID` (`modules/Funciones Generales.bas:2105`), `getNAccionACCalculado` (`modules/Funciones Generales.bas:3933`), `getID` (`modules/constructor.bas:1174`).
- **Registros** (2): `DameID1` (`modules/Funciones Generales.bas:495`), `getID` (`modules/constructor.bas:222`).
- **Brass** (2): `getIDParte` (`classes/Parte.cls:481`), `DameID1` (`modules/Funciones Generales.bas:332`).
- **AGEDYS Técnico** (2): `DameID1` (`modules/Funciones Generales.bas:1261`), `DameID` (`modules/Funciones Generales.bas:1506`).
- **AGEDYS ECO** (1): `DameID` (`modules/FUNCIONES UTILES.bas:519`).
- **Contratos** (1): `DameID` (`modules/Funciones Generales.bas:213`).

## Límites

- **Es lectura de código, no ejecución.** Un patrón dentro de un procedimiento que nadie llama cuenta igual, y el SQL que se compone en variables a lo largo de varias líneas solo se ve si el patrón cae en un mismo literal.
- **Las consultas guardadas no están en este recuento.** Su SQL está en el [anexo de controles conectados](controles-conectados.md) y en `queries.tsv` de la recogida; los patrones de ese SQL se revisan con él.
- **Falsos positivos posibles:** un `#` dentro de un texto que no es SQL, o un `Max(` y un `+ 1` del mismo procedimiento que no tienen relación. La lista de generadores se ha revisado a mano; los demás recuentos son cotas.
