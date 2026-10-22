# Matriz campo a campo — Gestión de Riesgos

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Issue:** #1079 · **Fase:** 3 · **PRD:** [README](README.md)

Cada tabla del backend de Gestión de Riesgos, con sus columnas y la evidencia de su uso en `src`. **Ninguna fila está aprobada**: la topología PostgreSQL global sigue abierta y el ERD de origen está pendiente del propietario (D-GR-4 del [README](README.md)).

## Fuente y límites

Se han cruzado dos fuentes, porque ninguna por sí sola sirve:

| Fuente | Tablas | Qué aporta | Qué le falta |
|---|---|---|---|
| [`ERD/schema.sql`](../ERD/schema.sql) (estudio previo, `origin/main`) | 68 `CREATE TABLE` (la cabecera anuncia 84) | Columna, tipo y número de filas | Claves primarias, índices, relaciones, nulabilidad |
| `ERD/Estructura_Datos.md` (propio del legacy, 2026-02-13) | 81 | Columna, tipo, longitud y **procedencia de las 15 tablas vinculadas** | Lo mismo, y es anterior al commit fijado en casi cinco meses |

La unión da **85 tablas**: 64 en ambas, 17 solo en `Estructura_Datos.md` y 4 solo en `schema.sql`. Las columnas se toman de `schema.sql` cuando existe y de `Estructura_Datos.md` cuando no; el tipo se expresa con la notación de `schema.sql` (`TEXT`, `LONG`, `MEMO`...) y, si solo hay `Estructura_Datos.md`, con el código numérico de tipo de Access (10 = texto, 4 = entero largo, 8 = fecha/hora, 12 = memo).

**Lo que sigue pendiente** (D-GR-4): claves primarias, índices, relaciones y nulabilidad. El backend `Gestion_Riesgos_Datos.accdb` no está en el clon y el workflow de extracción está en el PR #1144, abierto. Hasta entonces **cada fila es una propuesta**, no una confirmación, y las relaciones que se citen en las fichas son **inferidas por código**.

## Reglas de mapeo de tipos (propuesta)

| Tipo de origen | Tipo objetivo propuesto | Decisión abierta |
|---|---|---|
| `LONG` | `INTEGER` | Identidad: Access usa `Max + 1` (`DameID`, `src/modules/Funciones Generales.bas:2191, 2223`); la numeración de la web no está decidida |
| `SHORT_DATE_TIME` | `TIMESTAMP` | Fecha o fecha y hora, y zona horaria |
| `TEXT(n)` | `VARCHAR(n)` | `Sí`/`No` en texto: conservar hasta decidir `BOOLEAN` |
| `MEMO` | `TEXT` | — |
| `DOUBLE` | `DOUBLE PRECISION` | — |

## Clase de dato

Heurística **por nombre de columna**, sin mirar contenido: `credencial` (contraseñas, hashes, tokens), `personal` (nombres, correos, responsables, teléfonos) o `—`. Sirve para priorizar la revisión de privacidad y de cifrado; no sustituye a una clasificación formal.

## Advertencias de lectura

Antes de usar el recuento por evidencia, tres cosas:

1. **`TbID*` (15 tablas de secuencia).** Ninguna aparece citada en `src`: los identificadores los calcula `DameID`, que hace `Max + 1` sobre el nombre de tabla que le pasan (`Funciones Generales.bas:2191, 2223`). Son tablas de secuencia heredadas, probablemente sin uso tras ese cambio. «Sin evidencia» aquí **no** significa «sin uso»: hay que verificar contra el binario.
2. **`~TMPCLP576781` y `~TMPCLP603661`.** Tablas temporales que Access crea al ejecutar consultas de unión. No son estructura que haya que migrar.
3. **Tablas de caché (6).** `TbArbolCacheMeta`, `TbCacheArbolRiesgosMeta`, `TbCacheArbolRiesgosNodo`, `TbCacheControlCambiosMeta`, `TbCacheControlCambiosRow` y `TbCachePublicabilidadEdicion` materializan en disco el árbol de riesgos y el control de cambios. La migrabilidad de estas tablas está directamente ligada al problema de rendimiento del árbol (D88 del estudio previo): en la web el árbol debería calcularse con una CTE recursiva, no persistirse. **Son candidatas a no migrarse**; es una decisión abierta, no un hecho.

## Tablas de otros backends

Quince tablas están **vinculadas**: no las aporta esta aplicación. El detalle por tabla está en [vinculadas-a-otros-backends](matriz-campos/vinculadas-a-otros-backends.md); el análisis, en `coexistencia.md` (tanda posterior de esta cadena).

Por origen: `Lanzadera_Datos.accdb` (3 — identidad y permisos, también usadas por otras aplicaciones), `Expedientes_datos.accdb` (4), `AGEDYS_DATOS.accdb` (4), `AGEDO20_Datos.accdb` (2), `NoConformidades_Datos.accdb` (1) y `Gestion_Riesgos_Datos.accdb` (1, `TbProyectos1`, que es D-GR-5: un vínculo a sí mismo no puede ser).

## Resumen por tabla

| Tabla | Columnas | Filas (snapshot) | Ficheros de `src` | Vinculada | Detalle |
|---|---|---|---|---|---|
| **Proyecto, edición y riesgo** | | | | | |
| ProyectoRespCalidad | 2 | 16 | 1 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| **Evidencias, anexos y suministradores** | | | | | |
| TbAnexos | 9 | 484 | 2 |  | [detalle](matriz-campos/evidencias-y-suministradores.md) |
| TbAnexosAntigua | 9 | 2 | 0 |  | [detalle](matriz-campos/evidencias-y-suministradores.md) |
| **Identidad y permisos (vinculadas a Lanzadera)** | | | | | |
| TbAplicaciones | 0 | — | 0 | vinculada | [detalle](matriz-campos/identidad-y-permisos-vinculadas.md) |
| **Cachés y tablas auxiliares** | | | | | |
| TbArbolCacheMeta | 5 | 0 | 0 |  | [detalle](matriz-campos/cache-y-secundarias.md) |
| **Proyecto, edición y riesgo** | | | | | |
| TbAux | 1 | 13 | 0 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| TbAuxProyectosRiesgos | 5 | 2 | 1 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| **Cachés y tablas auxiliares** | | | | | |
| TbAuxSumExp | 2 | 6 | 0 |  | [detalle](matriz-campos/cache-y-secundarias.md) |
| **Proyecto, edición y riesgo** | | | | | |
| TbBibliotecaRiesgos | 9 | 65 | 2 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| **Cachés y tablas auxiliares** | | | | | |
| TbCacheArbolRiesgosMeta | 3 | 31 | 0 |  | [detalle](matriz-campos/cache-y-secundarias.md) |
| TbCacheArbolRiesgosNodo | 16 | 2089 | 0 |  | [detalle](matriz-campos/cache-y-secundarias.md) |
| TbCacheControlCambiosMeta | 5 | 50 | 1 |  | [detalle](matriz-campos/cache-y-secundarias.md) |
| TbCacheControlCambiosRow | 7 | 321 | 1 |  | [detalle](matriz-campos/cache-y-secundarias.md) |
| TbCachePublicabilidadEdicion | 10 | 0 | 1 |  | [detalle](matriz-campos/cache-y-secundarias.md) |
| **Publicación, control de cambios y tareas** | | | | | |
| TbCambiosExplicacion | 2 | 5 | 1 |  | [detalle](matriz-campos/publicacion-y-control-cambios.md) |
| **Catálogos y valores posibles** | | | | | |
| TbConfiguracionVisionRiesgos | 4 | 35 | 1 |  | [detalle](matriz-campos/catalogos-y-valores.md) |
| **Vinculadas a otros backends** | | | | | |
| TbDPD | 0 | — | 0 | vinculada | [detalle](matriz-campos/vinculadas-a-otros-backends.md) |
| TbDocumentos | 0 | — | 2 | vinculada | [detalle](matriz-campos/vinculadas-a-otros-backends.md) |
| TbDocumentosID | 0 | — | 0 | vinculada | [detalle](matriz-campos/vinculadas-a-otros-backends.md) |
| TbExpedientes | 0 | — | 1 | vinculada | [detalle](matriz-campos/vinculadas-a-otros-backends.md) |
| TbExpedientes1 | 0 | — | 1 | vinculada | [detalle](matriz-campos/vinculadas-a-otros-backends.md) |
| TbExpedientesResponsables | 0 | — | 2 | vinculada | [detalle](matriz-campos/vinculadas-a-otros-backends.md) |
| TbExpedientesSuministradores | 0 | — | 1 | vinculada | [detalle](matriz-campos/vinculadas-a-otros-backends.md) |
| **Publicación, control de cambios y tareas** | | | | | |
| TbExplicacionCarencias | 3 | 6 | 2 |  | [detalle](matriz-campos/publicacion-y-control-cambios.md) |
| **Catálogos y valores posibles** | | | | | |
| TbFelix | 2 | 47 | 0 |  | [detalle](matriz-campos/catalogos-y-valores.md) |
| TbHerramientaDocAyuda | 2 | 26 | 1 |  | [detalle](matriz-campos/catalogos-y-valores.md) |
| **Secuencias de identificadores y temporales** | | | | | |
| TbIDAccionPlanContingencia | 5 | 251 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| TbIDAccionPlanMitigacion | 5 | 1074 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| TbIDAnexos | 5 | 107 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| TbIDBibliotecaRiesgos | 5 | 51 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| TbIDCorreosEnviados | 5 | 264 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| TbIDDocumentos | 5 | 27 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| TbIDEdiciones | 5 | 209 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| TbIDLog | 5 | 1077 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| TbIDPlanContingencia | 5 | 154 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| TbIDPlanMitigacion | 5 | 640 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| TbIDProyectos | 5 | 71 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| TbIDRiesgos | 5 | 889 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| TbIDRiesgosAIntegrar | 5 | 58 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| TbIDTareas | 5 | 53 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| **Publicación, control de cambios y tareas** | | | | | |
| TbLog | 12 | 5817 | 0 |  | [detalle](matriz-campos/publicacion-y-control-cambios.md) |
| TbLogPublicaciones | 11 | 329 | 2 |  | [detalle](matriz-campos/publicacion-y-control-cambios.md) |
| **Vinculadas a otros backends** | | | | | |
| TbNoConformidades | 0 | — | 4 | vinculada | [detalle](matriz-campos/vinculadas-a-otros-backends.md) |
| **Proyecto, edición y riesgo** | | | | | |
| TbOrigenesRiesgosDetalles | 2 | 5 | 0 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| TbProyectoEdicionesCorreoRevision | 5 | 3 | 2 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| TbProyectos | 25 | 81 | 12 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| **Vinculadas a otros backends** | | | | | |
| TbProyectos1 | 0 | — | 0 | vinculada | [detalle](matriz-campos/vinculadas-a-otros-backends.md) |
| **Proyecto, edición y riesgo** | | | | | |
| TbProyectosEdiciones | 22 | 325 | 9 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| **Evidencias, anexos y suministradores** | | | | | |
| TbProyectosEdicionesSuministradores | 4 | 20 | 2 |  | [detalle](matriz-campos/evidencias-y-suministradores.md) |
| TbProyectosEdicionesSuministradores1 | 4 | — | 0 |  | [detalle](matriz-campos/evidencias-y-suministradores.md) |
| **Proyecto, edición y riesgo** | | | | | |
| TbProyectosResponsablesCalidad | 2 | 39 | 0 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| **Evidencias, anexos y suministradores** | | | | | |
| TbProyectosSuministradores | 4 | 0 | 2 |  | [detalle](matriz-campos/evidencias-y-suministradores.md) |
| TbProyectosSuministradores1 | 4 | — | 0 |  | [detalle](matriz-campos/evidencias-y-suministradores.md) |
| **Proyecto, edición y riesgo** | | | | | |
| TbProyectosTipo | 3 | 0 | 0 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| **Vinculadas a otros backends** | | | | | |
| TbResponsablesExpedientes | 0 | — | 0 | vinculada | [detalle](matriz-campos/vinculadas-a-otros-backends.md) |
| **Proyecto, edición y riesgo** | | | | | |
| TbRiesgos | 47 | 2924 | 11 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| TbRiesgosAIntegrar | 20 | 175 | 3 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| TbRiesgosAreasImpacto | 4 | 15 | 1 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| TbRiesgosMaterializaciones | 11 | 32 | 6 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| TbRiesgosNC | 6 | 32 | 2 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| **Planes de mitigación y de contingencia** | | | | | |
| TbRiesgosPlanContingenciaDetalle | 11 | 1817 | 4 |  | [detalle](matriz-campos/planes.md) |
| TbRiesgosPlanContingenciaDetalleReversa | 6 | 27 | 3 |  | [detalle](matriz-campos/planes.md) |
| TbRiesgosPlanContingenciaPpal | 8 | 1028 | 4 |  | [detalle](matriz-campos/planes.md) |
| TbRiesgosPlanMitigacionDetalle | 11 | 4221 | 4 |  | [detalle](matriz-campos/planes.md) |
| TbRiesgosPlanMitigacionDetalleReversa | 6 | 41 | 3 |  | [detalle](matriz-campos/planes.md) |
| TbRiesgosPlanMitigacionPpal | 8 | 2843 | 5 |  | [detalle](matriz-campos/planes.md) |
| **Proyecto, edición y riesgo** | | | | | |
| TbRiesgosPorTipoProyecto | 2 | 0 | 0 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| TbRiesgosValoracion | 3 | 25 | 1 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| **Vinculadas a otros backends** | | | | | |
| TbSuministradores | 0 | — | 1 | vinculada | [detalle](matriz-campos/vinculadas-a-otros-backends.md) |
| TbSuministradoresSAP | 0 | — | 0 | vinculada | [detalle](matriz-campos/vinculadas-a-otros-backends.md) |
| **Catálogos y valores posibles** | | | | | |
| TbTablas | 9 | 41 | 0 |  | [detalle](matriz-campos/catalogos-y-valores.md) |
| **Publicación, control de cambios y tareas** | | | | | |
| TbTareas | 11 | 36 | 0 |  | [detalle](matriz-campos/publicacion-y-control-cambios.md) |
| TbTareasExplicaciones | 3 | 10 | 1 |  | [detalle](matriz-campos/publicacion-y-control-cambios.md) |
| **Proyecto, edición y riesgo** | | | | | |
| TbUltimoProyecto | 4 | 29 | 3 |  | [detalle](matriz-campos/proyecto-edicion-riesgo.md) |
| **Identidad y permisos (vinculadas a Lanzadera)** | | | | | |
| TbUsuariosAplicaciones | 0 | — | 2 | vinculada | [detalle](matriz-campos/identidad-y-permisos-vinculadas.md) |
| TbUsuariosAplicacionesPermisos | 0 | — | 2 | vinculada | [detalle](matriz-campos/identidad-y-permisos-vinculadas.md) |
| **Catálogos y valores posibles** | | | | | |
| TbValoresPosiblesContingencia | 1 | 4 | 0 |  | [detalle](matriz-campos/catalogos-y-valores.md) |
| TbValoresPosiblesEstadoRiesgo | 1 | 7 | 0 |  | [detalle](matriz-campos/catalogos-y-valores.md) |
| TbValoresPosiblesMitigacion | 2 | 4 | 0 |  | [detalle](matriz-campos/catalogos-y-valores.md) |
| TbValoresPosiblesPlazoCalidadCosteVulnerabilidad | 1 | 5 | 0 |  | [detalle](matriz-campos/catalogos-y-valores.md) |
| TbValoresPosiblesValoracion | 1 | 4 | 0 |  | [detalle](matriz-campos/catalogos-y-valores.md) |
| **Publicación, control de cambios y tareas** | | | | | |
| tbCambios | 8 | 6725 | 2 |  | [detalle](matriz-campos/publicacion-y-control-cambios.md) |
| tbCambiosParaPublicacion | 8 | 68 | 0 |  | [detalle](matriz-campos/publicacion-y-control-cambios.md) |
| **Secuencias de identificadores y temporales** | | | | | |
| `~TMPCLP576781` | 4 | 0 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |
| `~TMPCLP603661` | 4 | 0 | 0 |  | [detalle](matriz-campos/secuencias-y-temporales.md) |

Totales: **85 tablas**, 41 con evidencia en `src` de producción y 44 sin ella. El detalle columna a columna está en las nueve páginas de familia.

## Tandas de la matriz

Esta tanda cubre las 85 tablas con su columna, tipo, clase de dato y evidencia de uso. Quedan pendientes: confirmar claves, índices, relaciones y nulabilidad con el ERD de origen (D-GR-4), decidir la correspondencia columna a columna con el esquema web —hoy no existe módulo `app/src/modules/gestion_riesgos`— y resolver si las tablas de caché se migran.
