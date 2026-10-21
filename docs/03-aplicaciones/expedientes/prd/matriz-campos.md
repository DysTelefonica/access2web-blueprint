# Matriz campo a campo — Expedientes

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Issue:** #1073 · **Fase:** 3 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Cada columna de origen de las tablas locales del backend de Expedientes, con la evidencia de su uso en `src` de RC1 y un destino PostgreSQL **propuesto**. Ninguna fila está aprobada: el [ERD global](../../../07-migracion/roadmap-descubrimiento-prd.md#erd-objetivo-postgresql-global) y la topología (un esquema por aplicación o varias bases) siguen abiertos.

## Fuente y límites

- **Columnas de origen:** [`ERD/schema.sql`](../ERD/schema.sql), extracción Jackcess del backend `Expedientes_datos.accdb` del snapshot anterior (73 tablas, 49 locales y 24 vinculadas). **No está revalidada contra el backend de RC1**: el ERD de origen queda pendiente porque `dysflow` no está disponible en el entorno de trabajo (el binario no existe en este host) y el backend no está en el clon de solo lectura de RC1 (solo el frontend).
- **Qué falta en esa fuente:** el `schema.sql` solo lista columna y tipo. No hay claves primarias, índices, relaciones, nulabilidad ni valores por defecto, salvo lo que el propio `src` declara en DDL por DAO (tablas E2E, ver más abajo). Sin ellos, la reconciliación de claves y los criterios de nulos quedan **pendientes del ERD de origen**.
- **Evidencia en `src`:** se busca cada columna en los ficheros de producción de RC1 que citan la tabla (excluidos los módulos `Test_*`). Es lectura estática, no prueba de que el campo se lea o escriba en ejecución.

## Reglas de mapeo de tipos (propuesta)

| Tipo de origen | Tipo objetivo propuesto | Decisión abierta |
|---|---|---|
| `LONG` | `BIGINT` | Identidad: conservar el ID legacy como clave; generación de IDs nuevos sin definir (hoy `Max + 1` en VBA) |
| `INT` | `SMALLINT` | — |
| `TEXT(n)` | `VARCHAR(n)` | Valores `Sí`/`No` en `TEXT(2)`: conservar el texto legacy hasta decidir si pasan a `BOOLEAN` |
| `MEMO` | `TEXT` | — |
| `DOUBLE` | `DOUBLE PRECISION` | Importes: `NUMERIC` frente a `DOUBLE PRECISION` (redondeo) |
| `SHORT_DATE_TIME` | `TIMESTAMP` | `DATE` o `TIMESTAMP`; zona horaria |
| `BOOLEAN` | `BOOLEAN` | — |

**Convención del destino:** esquema `expedientes`, mismo nombre de tabla y de columna que el origen, como ya hace la migración [`0002_expedientes_schema.py`](../../../../app/migrations/versions/0002_expedientes_schema.py) para los catálogos. Las columnas que esa migración ya crea se marcan «existente (migración 0002)»; el resto, «propuesta».

## Leyenda de la columna «Evidencia en `src`»

- `F` — la columna aparece como `Fields("columna")` en un fichero que cita la tabla (primer hallazgo).
- `S` — el nombre de la columna aparece en un fichero que cita la tabla, sin `Fields(...)`; es coincidencia de nombre y no prueba lectura ni escritura (poco fiable para nombres comunes).
- `—` — sin cita en los ficheros de producción de RC1.

## Resumen por tabla

| Tabla | Columnas | Filas (snapshot) | Ficheros de `src` | Disposición | Detalle |
|---|---|---|---|---|---|
| `Copia de TbExpedientes` | 54 | 0 | 0 | histórica o auxiliar; comparte 53 de 54 columnas con `TbExpedientes`; pendiente de decisión de negocio | sin detalle por columna |
| `Copia de TbExpedientesConEntidades` | 21 | 365 | 0 | histórica o auxiliar; comparte 3 de 21 columnas con `TbExpedientes`; pendiente de decisión de negocio | sin detalle por columna |
| `ListaPrevia` | 52 | 99 | 0 | histórica o auxiliar; comparte 10 de 52 columnas con `TbExpedientes`; pendiente de decisión de negocio | sin detalle por columna |
| `TbAusExpPostAGEDO` | 1 | 203 | 0 | auditoría, configuración o soporte | tanda posterior |
| `TbAuxEstadosMartina` | 3 | 143 | 0 | auditoría, configuración o soporte | tanda posterior |
| `TbAuxNemotecnico` | 1 | 0 | 0 | auditoría, configuración o soporte | tanda posterior |
| `TbCambios` | 10 | 0 | 1 | auditoría, configuración o soporte | tanda posterior |
| `TbComerciales` | 3 | 26 | 4 | catálogo | [catalogos](matriz-campos/catalogos.md#tbcomerciales) |
| `TbConfMostrarEstado` | 3 | 23 | 3 | auditoría, configuración o soporte | tanda posterior |
| `TbCPV` | 3 | 73 | 4 | catálogo | [catalogos](matriz-campos/catalogos.md#tbcpv) |
| `TbDatosEconomicosExpedientes` | 7 | 91 | 1 | auditoría, configuración o soporte | tanda posterior |
| `TbE2EExportBatch` | 10 | 0 | 0 | auditoría, configuración o soporte | tanda posterior |
| `TbE2EExportBatchDetalle` | 8 | 0 | 0 | auditoría, configuración o soporte | tanda posterior |
| `TbE2EExportSeleccionTemp` | 5 | 5 | 0 | auditoría, configuración o soporte | tanda posterior |
| `TbE2EJsonDestinationUserConfig` | 5 | 4 | 0 | auditoría, configuración o soporte | tanda posterior |
| `TbEjercitos` | 3 | 5 | 4 | catálogo | [catalogos](matriz-campos/catalogos.md#tbejercitos) |
| `TbEstados` | 3 | 9 | 1 | catálogo | [catalogos](matriz-campos/catalogos.md#tbestados) |
| `TbExpAgedys` | 49 | 116 | 1 | histórica o auxiliar; comparte 3 de 49 columnas con `TbExpedientes`; pendiente de decisión de negocio | sin detalle por columna |
| `TbExpAGEDYS1` | 49 | 116 | 0 | histórica o auxiliar; comparte 3 de 49 columnas con `TbExpedientes`; pendiente de decisión de negocio | sin detalle por columna |
| `TbExpedientes` | 61 | 453 | 20 | cabecera y read-model | [cabecera](matriz-campos/cabecera.md#tbexpedientes) |
| `TbExpedientes_antes` | 49 | 457 | 0 | histórica o auxiliar; comparte 3 de 49 columnas con `TbExpedientes`; pendiente de decisión de negocio | sin detalle por columna |
| `TbExpedientesAnexos` | 3 | 710 | 3 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedientesanexos) |
| `TbExpedientesAnualidades` | 11 | 174 | 5 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedientesanualidades) |
| `TbExpedientesCadenaContratacion` | 9 | 0 | 0 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedientescadenacontratacion) |
| `TbExpedientesCodigoCompras` | 3 | 0 | 4 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedientescodigocompras) |
| `TbExpedientesComerciales` | 3 | 333 | 5 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedientescomerciales) |
| `TbExpedientesConEntidades` | 21 | 451 | 7 | cabecera y read-model | [cabecera](matriz-campos/cabecera.md#tbexpedientesconentidades) |
| `TbExpedientesCPVs` | 3 | 429 | 6 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedientescpvs) |
| `TbExpedientesE2E` | 7 | 0 | 0 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedientese2e) |
| `TbExpedientesHitos` | 6 | 46 | 4 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedienteshitos) |
| `TbExpedientesJefaturas` | 3 | 0 | 2 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedientesjefaturas) |
| `TbExpedientesJuridicas` | 6 | 417 | 3 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedientesjuridicas) |
| `TbExpedientesLugaresEjecucion` | 3 | 194 | 5 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedienteslugaresejecucion) |
| `TbExpedientesModificados` | 6 | 37 | 5 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedientesmodificados) |
| `TbExpedientesPECAL` | 3 | 366 | 5 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedientespecal) |
| `TbExpedientesRACS` | 3 | 193 | 4 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedientesracs) |
| `TbExpedientesResponsables` | 6 | 730 | 8 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedientesresponsables) |
| `TbExpedientesSuministradores` | 7 | 713 | 9 | relación del expediente | [relaciones](matriz-campos/relaciones.md#tbexpedientessuministradores) |
| `TbGradosClasificacion` | 3 | 5 | 5 | catálogo | [catalogos](matriz-campos/catalogos.md#tbgradosclasificacion) |
| `TbJefaturas` | 3 | 3 | 0 | catálogo | [catalogos](matriz-campos/catalogos.md#tbjefaturas) |
| `TbJuridicas` | 4 | 17 | 2 | catálogo | [catalogos](matriz-campos/catalogos.md#tbjuridicas) |
| `TbLugaresEjecucion` | 3 | 34 | 4 | catálogo | [catalogos](matriz-campos/catalogos.md#tblugaresejecucion) |
| `TbOficinasPrograma` | 3 | 37 | 5 | catálogo | [catalogos](matriz-campos/catalogos.md#tboficinasprograma) |
| `TbOrganosContratacion` | 3 | 46 | 4 | catálogo | [catalogos](matriz-campos/catalogos.md#tborganoscontratacion) |
| `TbPECAL` | 3 | 6 | 5 | catálogo | [catalogos](matriz-campos/catalogos.md#tbpecal) |
| `TbRACS` | 4 | 33 | 4 | catálogo | [catalogos](matriz-campos/catalogos.md#tbracs) |
| `TbResponsablesPorRol` | 3 | 11 | 2 | catálogo | [catalogos](matriz-campos/catalogos.md#tbresponsablesporrol) |
| `TbSuministradores` | 10 | 72 | 7 | catálogo | [catalogos](matriz-campos/catalogos.md#tbsuministradores) |
| `TbUltimoCambio` | 4 | 6 | 3 | auditoría, configuración o soporte | tanda posterior |

## Tandas de la matriz

Esta tanda añade los catálogos (14 tablas). Siguen las tablas transversales, las que solo `src` define (`TbE2ECola`, `TbE2EEnlaces`) y las vinculadas a otros backends. Las tablas históricas no tienen detalle por columna: el [ledger de la matriz de migración](../migration-matrix.md) las conserva hasta que negocio decida su disposición.
