# Fase 0: backends compartidos con consumidores dentro del alcance

Issue #1638. Épica `docs/07-migracion/epica-coexistencia-postgresql.md` (fase 0).
Alcance: `Correos_datos` y `Registro_Ent_Salida_Datos`. Sus aplicaciones (Correos y el grupo de Registros) no entran en la migración, pero sus tablas las consumen aplicaciones que sí entran, y con la D92 la web va a leer esas mismas tablas en PostgreSQL. Se inventarían como **dependencias**, no como objetivos de migración.
Método: ERD con `scripts/dysflow-erd/extract-erd.ps1`, que abre una copia temporal en solo lectura con DAO. Recuentos de la copia de desarrollo, no de producción.

## `Correos_datos`

| Hecho | Valor |
|---|---|
| Tablas | 1: `TbCorreosEnviados` |
| Filas | 35 394 |
| Clave primaria | **Ninguna** |
| Índices | `IDCorreo`, `IDEdicion` (ninguno único) |
| Relaciones de aplicación | 0 (la única relación del ERD es del panel de navegación de Access) |
| Consultas guardadas | 1: `Consulta1` |

**Consumidores dentro del alcance:** seis backends vinculan `TbCorreosEnviados` (Expedientes, Condor, No Conformidades, Gestión de Riesgos, HPS, Lanzadera) y el front de Registro escribe en ella. Según el grafo (`grafo-dependencias.md`), **seis backends la consumen**. El front de Expedientes la escribe con un `INSERT`; No Conformidades, Gestión de Riesgos, HPS y Solicitudes HPS también escriben sobre ella.

Es el punto de tensión más grande del ecosistema: una tabla sin clave, con 35 394 filas, escrita desde varias aplicaciones y leída desde seis backends. Cuatro fichas piden clave primaria sobre ella, desde cuatro aplicaciones distintas (Conflicto X-01 del catálogo).

## `Registro_Ent_Salida_Datos`

Es el backend real del front de Registros, distinto de `Registro_Datos.accdb`.

| Tabla | Filas | Clave primaria | Índices |
|---|---|---|---|
| `Salidas` | vínculo a `Registro_Datos.accdb` | `PrimaryKey` | `IDEntrada`, `sEGUNDA` |
| `TbEntradas_RES` | 635 | `PrimaryKey` | `IDEntrada`, `NEntrada` |
| `TbSalidas_RES` | 4 053 | `PrimaryKey` | `IDSalida` |
| `TbSalidas_RES_Indexacion` | 1 063 | `PrimaryKey` | `IDEntrada` |
| `TbLog` | 3 341 | `PrimaryKey` | `IDAR`, `IDEdicion`, `IDLog`, `IDProyecto` |
| `TbConexiones` | 12 | `PrimaryKey` | — |
| `TbEntradaDestinatarios_RES` | 37 | `PrimaryKey` | `IDEntrada` |
| `TbSalidaDestinatarios_RES` | 5 | `PrimaryKey` | `IDEntrada` |
| `TbEntradas_RES_Indexacion` | 2 | `PrimaryKey` | `IDEntrada` |
| `TbEntradaPersonasContacto_RES` | 0 | `PrimaryKey` | `IDEntrada` |
| `TbEntradas_indexacion` | 0 | `PrimaryKey` | `IDEntrada` |
| `TbSalidas_Indexacion` | 0 | `PrimaryKey` | `IDEntrada` |

Todas tienen clave primaria y todas tienen índice sobre `IDEntrada`. Relaciones de aplicación: 0. Consultas guardadas: 1.

**Dos rutas para la misma tabla.** El front de Registros crea con `getdb()` (`SalidaOperaciones.cls:106-107`) y borra con `getdbSalida()` (`:166-169`), y la numeración se lee en `Salida.cls:88, 187-190`. Dos caminos distintos hacia `TbSalidas_RES` en el mismo front: es un riesgo de coherencia al migrar, no una limpieza.

**Dueño y consumidores.** `TbSalidas_RES` y `TbEntradas_RES` son de Registro, pero AGEDYS ECO las usa como origen y las referencia desde su propio front. La ficha de AGEDYS ECO las declara como tabla de Registro con vínculo en `AGEDYS_DATOS`, y anota que `TbEntradas_RES` no la usa ninguno de los dos fronts de AGEDYS.

**`TbLog` tiene 613 filas huérfanas** en `IDEntrada`, ya registradas en la ficha de Registros.

## Propuestas

Ninguna se aplica: la puerta de descubrimiento global sigue cerrada (#1576).

| Id | Propuesta | Grupo (fase 1) | Comprobación previa | Reversión |
|---|---|---|---|---|
| C1 | Fijar clave primaria en `TbCorreosEnviados`: hoy tiene 35 394 filas y dos índices no únicos, `IDCorreo` e `IDEdicion`. Antes hay que decidir quién es su dueño. | Imprescindible | Contar filas con `IDCorreo` duplicado y con `IDEdicion` duplicado en una copia reciente de producción. Verificar qué aplicaciones escriben y con qué frecuencia. | Quitar la clave. |
| C2 | Decidir si `TbCorreosEnviados` es un servicio propio de la plataforma o sigue siendo un backend más del legacy. Hoy la escriben al menos cinco aplicaciones y la leen seis backends. | Imprescindible | Es una decisión de producto y de arquitectura, no de limpieza. Si pasa a servicio, hay que decidir su esquema y su API antes de migrar nada que dependa de ella. | Sin cambios mientras no se decida. |
| C3 | Unificar el acceso a `TbSalidas_RES` en el front de Registros: hoy se crea por `getdb()` y se borra por `getdbSalida()`. | Imprescindible | Reproducir en runtime un alta y un borrado de salida y comprobar que las dos rutas apuntan al mismo conjunto. | Volver a las dos rutas originales. |
| C4 | Declarar en `Registro_Ent_Salida_Datos` las relaciones que los datos ya cumplen: `TbSalidas_RES.IDEntrada` y `TbEntradas_RES.IDEntrada` hacia `Salidas` e `TbEntradas_RES` respectivamente. | Recomendable si es barato | Consultas de huérfanos con 0 filas sobre la copia de producción. Sin borrado en cascada. | Quitar la relación. |
| C5 | Revisar las 613 filas huérfanas de `TbLog.IDEntrada` y las 5 tablas de indexación vacías. | Recomendable si es barato | Confirmar que ninguna consulta guardada las usa. `Consulta1` es la única y su SQL hay que leerlo. | Restaurar las filas desde la copia. |
| C6 | Declarar explícitamente que `TbSalidas_RES` y `TbEntradas_RES` son de Registro y que AGEDYS ECO es consumidor, no propietario. | Imprescindible | Es la regla de propiedad por uso del front propietario, aplicada a un backend de aplicación fuera del alcance. | Sin cambios mientras no se decida. |
| C7 | Revisar la matriz de vínculos de la lanzadera para este par de backends: la ficha de AGEDYS ECO detectó 6 tablas declaradas que ya no existen y 72 existentes sin declarar. | Recomendable si es barato | Volcar `TbTablasAVincular` y contrastar con el ERD. | Restaurar la matriz anterior. |

## Conflictos con lo ya catalogado

- **X-01** (catálogo único): cuatro propuestas de clave primaria sobre `TbCorreosEnviados` desde cuatro aplicaciones. C1 las sustituye por una sola, pero la decisión de fondo sigue siendo del propietario (C2).
- **H2** (ficha de HPS) proponía retirar el vínculo que expone `TbSolicitudesFechas`; la ficha de Solicitudes HPS demuestra que el front escribe en tablas del backend de HPS. Afecta a este par solo de forma indirecta, pero confirma que las propuestas sobre vínculos necesitan el grafo completo.

## Lo que no se pudo verificar

1. Los vínculos en producción (`\\datoste\…`) no son alcanzables desde este entorno: los campos se leyeron del ERD, no del origen.
2. Los recuentos son de la copia de desarrollo; `TbSalidas_RES` con 4 053 filas y `TbEntradas_RES` con 635 deben verificarse en una copia reciente de producción antes de aprobar cualquier acción.
3. El SQL de `Consulta1`, el único guardado en cada backend, no se ha leído.
4. Si `Registro_Ent_Salida_Datos` y `Registro_Datos` comparten tablas o solo se parecen: `Salidas` es un vínculo al segundo, y el primero tiene sus propias tablas `*_RES`.

## Backends de aplicaciones fuera del alcance pendientes de ficha

`SICA_datos` (4 tablas, todas sin clave primaria), `Control_Cambios_datos` (3 tablas), `AGEDO20_Datos` (66 tablas propias, 12 relaciones, 8 sin clave primaria) y `Tareas_datos1` (4 tablas propias y 13 vínculos; es el dueño de `TbUsuariosAplicacionesTareas`, ya marcada como mal ubicada). Se inventarían en un documento aparte.
