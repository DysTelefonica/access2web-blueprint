# Fase 0: inventario de `Lanzadera_Datos`

> **Issue:** #1560 · **Épica:** [coexistencia progresiva](../epica-coexistencia-postgresql.md), fase 0 y decisión 12 · **Fecha:** 2026-11-04

## Propósito

Inventario de solo lectura del backend de la Lanzadera, la primera aplicación en coexistir. Termina en una lista de **acciones aplicables ya en el legacy** (fase 1), cada una con su comprobación previa y su reversión.

## Fuente y método

- **Backend inventariado:** `C:\00repos\datos\Lanzadera_Datos.accdb` (copia de desarrollo del 2026-11-04). No es producción (`\\datoste\...`): los recuentos de filas son orientativos y hay que repetirlos sobre producción antes de aplicar una acción que borre datos.
- **Otra copia:** `00_LANZADERA\Lanzadera_Datos.accdb` (2026-09-17) tiene el mismo esquema y menos filas (por ejemplo, `TbAplicaciones` 19 frente a 22). Es la que lee dysflow por defecto.
- **Estructura:** [`scripts/dysflow-erd/extract-erd.ps1`](../../../scripts/dysflow-erd/extract-erd.ps1) sobre una copia temporal abierta con DAO en solo lectura. Extrae campos, índices, relaciones, reglas de validación, recuento de filas y vínculos; un vínculo con origen inaccesible queda registrado en `linkError` sin detener la extracción.
- **Consumidores:** el mismo script sobre los 16 backends de `C:\00repos\datos` dentro del alcance; se cuentan los vínculos cuyo origen es `Lanzadera_Datos.accdb`.
- **Referencias en código:** recuento por nombre exacto de tabla en `src` (`.bas`, `.cls`, `.form.txt`, `.report.txt`, `.sql`) de los 13 fronts con repositorio. Es un recuento por nombre: `tbUsuarios` coincide con una tabla propia de HPS y sus referencias fuera de la Lanzadera no se cuentan.
- **Consultas guardadas:** el front no tiene ninguna (`export_queries`); el backend tiene `Consulta1` a `Consulta4`, de prueba, que solo leen tablas de identidad.
- **Uso vivo o muerto:** CodeGraph-VBA (índice del 2026-11-04) y búsqueda de llamadores en `src`, incluidas las secciones de código de los `.form.txt`; `find_references` de dysflow con `scope: binary` para el binario. CodeGraph modela las tablas solo de forma parcial (nodos sintéticos en algunas líneas de SQL), así que no sirve para contar referencias; sí para el grafo de llamadas. Quedan sin revisar las macros, la cinta de opciones y las expresiones de propiedades de formularios, que no se exportan a `src`.
- **Patrones de VBA del front:** ya inventariados en la [coexistencia de Lanzadera](../../03-aplicaciones/lanzadera/prd/coexistencia.md), apartados 4 y 5; no se repiten aquí.

## Resumen

| Hecho | Valor |
|---|---|
| Tablas propias / vinculadas | 35 / 1 (`TbCorreosEnviados`, de `Correos_datos.accdb`) |
| Tablas sin clave primaria | 1: `TbTablasAVincular` |
| Campos multivalor, adjuntos, OLE o hipervínculo | 0 |
| Reglas de validación | 0 |
| Relaciones con integridad referencial | 5, todas con actualización y borrado en cascada |
| Campos Sí/No guardados como texto | 28; valores reales solo `Sí`, `No` y nulo |
| Tablas que vinculan otros backends | 5 |
| Candidatas a huérfanas | 22 (20 sin referencias y 2 con referencias solo en código muerto) |

El backend es apto para PostgreSQL sin cambios de tipo: las tablas tienen clave primaria salvo una y no hay tipos sin equivalente.

## Acciones para aplicar ya en el legacy

Antes de cualquier acción: copia de seguridad del backend de producción y ventana sin usuarios conectados. Las consultas de comprobación se ejecutan en producción, no en la copia de desarrollo.

| Id | Acción | Grupo (fase 1) | Comprobación previa | Reversión |
|---|---|---|---|---|
| A1 | Corregir la ayuda: el código consulta `TbHerramientaDocAyuda` (`Funciones Genreales.bas:538`) y la tabla se llama `Tb0HerramientaDocAyuda`. Renombre la tabla a `TbHerramientaDocAyuda`, o retire la llamada si la ayuda ya no se usa. | Imprescindible | Abra la ayuda desde la Lanzadera: hoy debe fallar. | Volver a renombrar. |
| A2 | Poner en cuarentena las 6 tablas vacías y sin referencias: `TbAplicacionesVideos`, `TbCuestionarioPreguntas`, `TbCuestionarios`, `TbCuestionaroRespuestas`, `TbUsuarioAplicacionesSolicitud`, `TbVideosCuestionario`. Renombrar a `zzBAJA_<tabla>`. | Imprescindible | `SELECT COUNT(*)` de cada una en producción da 0. | Quitar el prefijo. |
| A3 | Poner en cuarentena `TbTablasAVincular` (411 filas, sin clave primaria, 0 referencias en ningún `src`). Si resulta que la usa un front sin repositorio, en lugar de cuarentena añada la clave primaria (`IDBBDD`, `IDAplicacion`, `NombreTabla`). | Imprescindible | Confirme que ningún front fuera de los repositorios relinka desde esta tabla. Para la clave: `SELECT IDBBDD, IDAplicacion, NombreTabla FROM TbTablasAVincular GROUP BY IDBBDD, IDAplicacion, NombreTabla HAVING COUNT(*) > 1` debe devolver 0 filas. | Quitar el prefijo o la clave. |
| A4 | Poner en cuarentena las 15 candidatas restantes con datos: las 13 sin referencias de la ficha y `TbVideosVisionados` y `TbUsuariosCorreosEnvio`, cuyas referencias están todas en código muerto (la clase `Visionado` nunca se instancia; `Constructor.getCorreo` no tiene llamadores). Empezar por las de uso más improbable: `TbPermisos` (69 de sus 205 filas apuntan a aplicaciones inexistentes), `TbAplicacionesEdiciones`, `TbAplicacionesEdicionesCambios`, `TbDetalleVersiones`, `TbParametros` y `TbAplicacionesParametros`. | Imprescindible | El [protocolo de huérfanas](../epica-coexistencia-postgresql.md#protocolo-de-huérfanas): confirmar que no las leen Excel, informes ni tareas programadas. `tbUsuarios` contiene datos personales (DNI): no se borra sin decisión expresa. | Quitar el prefijo. |
| A5 | Borrar los registros huérfanos de `TbConexionUltimaAppAbierta.IDUltimaAplicacionAbierta` → `TbAplicaciones` (9 huérfanos) y declarar la relación con integridad referencial. (`TbVideosVisionados`, que figuraba aquí, pasa a A4.) | Imprescindible | Las consultas de huérfanos de [reproducción](#reproducción) devuelven los mismos recuentos. | Restaurar las filas de la copia y quitar la relación. |
| A6 | Declarar con integridad referencial las relaciones que ya cumplen los datos (0 huérfanos): `TbAplicacionesAperturas.IDAplicacion` → `TbAplicaciones`, `TbVideos.IDAplicacion` → `TbAplicaciones`, `TbVideosCategorias.IDVideo` → `TbVideos`, `TbVideosCategorias.IDCategoria` → `TbCategorias`. Sin borrado en cascada. | Recomendable si es barato | Consultas de huérfanos con 0 filas. Compruebe que la Lanzadera no borra aplicaciones ni vídeos con datos dependientes: con la relación, ese borrado fallará. | Quitar la relación. |
| A7 | Borrar las consultas de prueba `Consulta1` a `Consulta4` del backend. | Recomendable si es barato | Ningún front las referencia (0 apariciones en `src`). | Recrearlas desde este inventario. |

Lo que **no** se toca ahora:

- **Los 28 campos Sí/No de texto** se quedan como texto: sus valores ya son limpios y cambiarlos exige tocar 183 líneas de VBA solo en la Lanzadera. En PostgreSQL bastará una restricción `CHECK` sobre `Sí` y `No`.
- **`TbUsuariosAplicaciones.Id` (`Integer` de 16 bits, máximo 32 767; hoy 190 filas)** no se cambia en Access: el ensanchamiento a `integer` se hace directamente en el DDL de PostgreSQL, sin tocar el legacy ni revincular sus 7 consumidores.
- **La clave natural `CorreoUsuario`** (texto) de las tablas de identidad se mantiene: la referencian 7 backends y más de 500 líneas de código.
- **`TbUsuariosAplicacionesTareas`** está mal ubicada (la usa solo `Tareas_datos1`; ver R-2): se registra su dueño real y no se mueve.

## Consumidores de `Lanzadera_Datos`

| Tabla | Backends que la vinculan |
|---|---|
| `TbUsuariosAplicaciones` | AGEDYS_DATOS, Expedientes_datos, Gestion_Riesgos_Datos, HPST, NoConformidades_Datos, Tareas_datos1, condor_datos |
| `TbUsuariosAplicacionesPermisos` | Los mismos 7 |
| `TbAplicaciones` | AGEDYS_DATOS, Gestion_Riesgos_Datos |
| `TbConexionesRegistro` | AGEDYS_DATOS |
| `TbUsuariosAplicacionesTareas` | Tareas_datos1 |

Las tablas de identidad (`TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos`) suman 559 referencias en 13 fronts: cualquier cambio de su estructura es el más caro del ecosistema y queda fuera de la coexistencia.

## Verificación de R-1 y R-2

- **R-1 (cada front consulta solo su backend): se cumple en el front de la Lanzadera y se refuta para su backend.** El front de la Lanzadera abre solo `Lanzadera_Datos` (`getdb()`) y tiene vinculado lo que necesita de otros (`TbCorreosEnviados`). Pero **11 fronts abren `Lanzadera_Datos.accdb` directamente**, con la ruta escrita en el código y sin pasar por un vínculo de su backend: AGEDYS (`Variables Globales.bas:447`), AGEDYS ECO (`Variables Globales.bas:447`), Brass (`Variables Globales.bas:528`), Contratos (`Variables Entorno.bas:257`), Registros (`Variables de Entorno.bas:380`), Gestión de Riesgos (`Variables Globales.bas:465`), HPS (`VariablesEntorno.bas:553`), Solicitudes HPS (`Variables Globales.bas:438`), Seguridad (`Variables de Entorno.bas:411`), Expedientes (`getdbLanzadera()`) y Condor (`FUNCIONES UTILES.bas:252`). No Conformidades accede por vínculos en `main`. Todos leen las tablas de identidad por esa conexión; solo AGEDYS, AGEDYS ECO, Brass, Contratos y Registros tienen código que escribe `TbAplicacionesAperturas`, y solo en Contratos y Registros es alcanzable desde eventos. Ningún front escribe `TbConexionesRegistro` en su `src`. Las rutas de escritura propias de la Lanzadera (`RegistroApertura`, `RegistroCierre`, `RegistroConexion`) no tienen llamadores. Detalle por front: [aperturas directas de `Lanzadera_Datos`](aperturas-directas-lanzadera.md) (#1570).
- **Corrección del 2026-11-04 (#1570).** La versión anterior de este apartado daba Condor como probable y atribuía a los fronts con apertura directa la escritura de `TbAplicacionesAperturas` y `TbConexionesRegistro`. El inventario de #1570 confirma Condor y limita las escrituras observadas a `TbAplicacionesAperturas` en cinco fronts.
- **R-2 (cada tabla vive en el backend de quien la usa): una excepción.** `TbUsuariosAplicacionesTareas` vive en `Lanzadera_Datos`, pero solo la usa `Tareas_datos1`. Su dueño real es Tareas, que no figura en el catálogo de la Lanzadera.

## Hallazgos del ecosistema

Fuera del alcance de este backend, pero condicionan las fases siguientes:

- **`Expedientes_datos` es el núcleo más vinculado:** lo vinculan AGEDYS_DATOS (23 tablas), condor_datos (6), Gestion_Riesgos_Datos, HPST, NoConformidades_Datos (4 cada uno), Solicitudes_HPS_datos y Tareas_datos1 (2 cada uno). `TbTablasAVincular` no lo reflejaba.
- **Dependencia circular AGEDYS ↔ Expedientes:** `Expedientes_datos` vincula 15 tablas de `AGEDYS_DATOS` y este vincula 23 de `Expedientes_datos`. Ambos backends tendrán que trasladarse en el mismo corte (H-2).
- **Un vínculo de desarrollo apunta a producción:** `TbExpedientes1` de la copia de `Expedientes_datos` enlaza con `\\datoste\...\AGEDYS_DATOS.accdb`. Revíselo: desde desarrollo se puede leer o escribir en producción.
- **`AGEDO20_Datos`, fuera de alcance, sigue vinculado** desde AGEDYS_DATOS (3 tablas) y Gestion_Riesgos_Datos (2). Hay que decidir si esas tablas viajan con sus consumidores.
- **`Tareas_datos1` no está en el catálogo** de la Lanzadera, pero vincula tablas de la Lanzadera, de AGEDYS y de Expedientes.

## Ficha por tabla

Clasificación según la [épica](../epica-coexistencia-postgresql.md#clasificación-de-tablas). «Candidata a huérfana» significa 0 referencias en código y ningún vínculo desde otro backend; no autoriza el borrado por sí sola.

| Tabla | Clave primaria | Filas | Refs. Lanzadera | Refs. otros fronts | Vinculada por | Clasificación |
|---|---|---|---|---|---|---|
| `Tb0HerramientaDocAyuda` | NombreFormulario, NombreArchivoAyuda | 4 | 0 | 1 | — | propia, con nombre roto (A1) |
| `TbAplicaciones` | IDAplicacion | 22 | 18 | 0 | AGEDYS_DATOS, Gestion_Riesgos_Datos | propia y compartida |
| `TbAplicacionesAperturas` | IDApertura | 1697 | 19 | 14 | — | propia y compartida (la escriben otros fronts por apertura directa) |
| `TbAplicacionesEdiciones` | IDAplicacion, IDVersion | 1262 | 0 | 0 | — | candidata a huérfana |
| `TbAplicacionesEdicionesCambios` | IDCambio | 1364 | 0 | 0 | — | candidata a huérfana |
| `TbAplicacionesEstados` | PerfilAplicacion | 6 | 0 | 0 | — | candidata a huérfana |
| `TbAplicacionesParametros` | IDAplicacion, IDParametro | 60 | 0 | 0 | — | candidata a huérfana |
| `TbAplicacionesPerfiles` | IDAplicacion, Perfil | 34 | 0 | 0 | — | candidata a huérfana |
| `TbAplicacionesVideos` | IDAplicacionVideo | 0 | 0 | 0 | — | candidata a huérfana |
| `TbCategorias` | IDCategoria | 2 | 20 | 0 | — | propia |
| `TbConexiones` | Usuario | 24 | 0 | 0 | — | candidata a huérfana |
| `TbConexionesRegistro` | IDConexion | 10360 | 24 | 5 | AGEDYS_DATOS | propia y compartida |
| `TbConexionUltimaAppAbierta` | IDConexion | 1469 | 4 | 0 | — | propia |
| `TbConfiguracion` | Clave | 1 | 1 | 3 | — | propia |
| `TbCuestionarioPreguntas` | IDPregunta | 0 | 0 | 0 | — | candidata a huérfana |
| `TbCuestionarios` | IDCuestionario | 0 | 0 | 0 | — | candidata a huérfana |
| `TbCuestionaroRespuestas` | IDRespuesta | 0 | 0 | 0 | — | candidata a huérfana |
| `TbDetalleVersiones` | IDAplicacion, IDVersion, IDDetalle | 28 | 0 | 0 | — | candidata a huérfana |
| `TbParametros` | IDParametro | 53 | 0 | 0 | — | candidata a huérfana |
| `TbPermisos` | IDAplicacion, Usuario | 205 | 0 | 0 | — | candidata a huérfana |
| `TbTablasAVincular` | **ninguna** | 411 | 0 | 0 | — | candidata a huérfana |
| `TbUbicaciones` | NombreUbicacion, Sirdee | 8 | 0 | 0 | — | candidata a huérfana |
| `TbUsuarioAplicacionesSolicitud` | CorreoUsuario | 0 | 0 | 0 | — | candidata a huérfana |
| `TbUsuarioConfiguracion` | UsuarioDeRed | 14 | 0 | 0 | — | candidata a huérfana |
| `tbUsuarios` | Id | 156 | 0 | 0 | — | candidata a huérfana |
| `TbUsuariosAplicaciones` | CorreoUsuario | 190 | 39 | 390 | AGEDYS_DATOS, Expedientes_datos, Gestion_Riesgos_Datos, HPST, NoConformidades_Datos, Tareas_datos1, condor_datos | propia y compartida |
| `TbUsuariosAplicacionesPermisos` | CorreoUsuario, IDAplicacion | 655 | 16 | 114 | AGEDYS_DATOS, Expedientes_datos, Gestion_Riesgos_Datos, HPST, NoConformidades_Datos, Tareas_datos1, condor_datos | propia y compartida |
| `TbUsuariosAplicacionesTareas` | CorreoUsuario | 7 | 0 | 0 | Tareas_datos1 | mal ubicada |
| `TbUsuariosCorreosEnvio` | IDCorreo | 1042 | 3 | 0 | — | candidata a huérfana (solo código muerto) |
| `TbUsuariosHistoricoContrasenias` | Usuario, PassAntigua | 483 | 20 | 0 | — | propia |
| `TbUsuariosTareasDiarias` | UsuarioRed | 8 | 0 | 0 | — | candidata a huérfana |
| `TbVideos` | IDVideo | 16 | 31 | 0 | — | propia |
| `TbVideosCategorias` | IDCategoriaVideo | 15 | 27 | 0 | — | propia |
| `TbVideosCuestionario` | IDVideoCuestionario | 0 | 0 | 0 | — | candidata a huérfana |
| `TbVideosVisionados` | IDVisionado | 138 | 4 | 0 | — | candidata a huérfana (solo código muerto) |

## Reproducción

```bash
# Estructura de un backend (sobre copia temporal, solo lectura)
ACCDB_PASSWORD="$ACCESS_VBA_PASSWORD" powershell.exe -NoProfile -ExecutionPolicy Bypass \
  -File scripts/dysflow-erd/extract-erd.ps1 -Path 'C:/00repos/datos/Lanzadera_Datos.accdb' -OutDir "$TEMP/fase0"
```

Huérfanos de una relación no declarada (dysflow `query_execute`, `mode: "read"`, `target: "backend"`):

```sql
SELECT COUNT(*) FROM TbVideosVisionados AS c
LEFT JOIN TbVideos AS p ON c.IDVideo = p.IDVideo
WHERE p.IDVideo IS NULL AND c.IDVideo IS NOT NULL;
```

El mismo patrón sirve para el resto de pares de la acción A5 y A6, y para `TbPermisos`, `TbAplicacionesEdiciones`, `TbAplicacionesPerfiles` y `TbDetalleVersiones` frente a `TbAplicaciones.IDAplicacion`.
