# Matriz campo a campo — transversales

> Parte de la [matriz campo a campo](../matriz-campos.md). Reglas de mapeo, leyenda y límites de la fuente en ese documento. Commit fijado `ec6a8c3`.

## TbCambios

10 columnas · 0 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbCambios.IDCambio` | LONG(4) | S `constructor.bas:5246` | `expedientes.TbCambios.IDCambio` | BIGINT | propuesta |
| `TbCambios.NombreTabla` | TEXT(255) | — | `expedientes.TbCambios.NombreTabla` | VARCHAR(255) | propuesta |
| `TbCambios.NombreCampoID` | TEXT(255) | — | `expedientes.TbCambios.NombreCampoID` | VARCHAR(255) | propuesta |
| `TbCambios.ValorCampoID` | LONG(4) | — | `expedientes.TbCambios.ValorCampoID` | BIGINT | propuesta |
| `TbCambios.NombreCampo` | TEXT(255) | — | `expedientes.TbCambios.NombreCampo` | VARCHAR(255) | propuesta |
| `TbCambios.ValorInicial` | TEXT(255) | — | `expedientes.TbCambios.ValorInicial` | VARCHAR(255) | propuesta |
| `TbCambios.ValorFinal` | TEXT(255) | — | `expedientes.TbCambios.ValorFinal` | VARCHAR(255) | propuesta |
| `TbCambios.FechaCambio` | SHORT_DATE_TIME(8) | S `constructor.bas:589` | `expedientes.TbCambios.FechaCambio` | TIMESTAMP | propuesta |
| `TbCambios.IDUsuarioCambio` | LONG(4) | S `constructor.bas:588` | `expedientes.TbCambios.IDUsuarioCambio` | BIGINT | propuesta |
| `TbCambios.Accion` | TEXT(255) | — | `expedientes.TbCambios.Accion` | VARCHAR(255) | propuesta |

## TbUltimoCambio

4 columnas · 6 filas en la extracción del snapshot · citada en 3 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbUltimoCambio.ID` | LONG(4) | F `UltimoCambioOperaciones.cls:40` | `expedientes.TbUltimoCambio.ID` | BIGINT | propuesta |
| `TbUltimoCambio.IDExpediente` | LONG(4) | F `UltimoCambioOperaciones.cls:41` | `expedientes.TbUltimoCambio.IDExpediente` | BIGINT | propuesta |
| `TbUltimoCambio.FechaCambio` | SHORT_DATE_TIME(8) | F `UltimoCambioOperaciones.cls:42` | `expedientes.TbUltimoCambio.FechaCambio` | TIMESTAMP | propuesta |
| `TbUltimoCambio.IDUsuarioCambio` | LONG(4) | F `UltimoCambioOperaciones.cls:43` | `expedientes.TbUltimoCambio.IDUsuarioCambio` | BIGINT | propuesta |

## TbConfMostrarEstado

3 columnas · 23 filas en la extracción del snapshot · citada en 3 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbConfMostrarEstado.ID` | LONG(4) | F `FUNCIONES UTILES.bas:4131` | `expedientes.TbConfMostrarEstado.ID` | BIGINT | propuesta |
| `TbConfMostrarEstado.UsuarioRed` | TEXT(255) | F `FUNCIONES UTILES.bas:4132` | `expedientes.TbConfMostrarEstado.UsuarioRed` | VARCHAR(255) | propuesta |
| `TbConfMostrarEstado.MostrarEstado` | TEXT(2) | F `FUNCIONES UTILES.bas:4133` | `expedientes.TbConfMostrarEstado.MostrarEstado` | VARCHAR(2) · Sí/No como texto | propuesta |

## TbDatosEconomicosExpedientes

7 columnas · 91 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbDatosEconomicosExpedientes.IDExpediente` | LONG(4) | F `ExpedienteJsonExporter.bas:817` | `expedientes.TbDatosEconomicosExpedientes.IDExpediente` | BIGINT | propuesta |
| `TbDatosEconomicosExpedientes.MontoTotal` | DOUBLE(8) | — | `expedientes.TbDatosEconomicosExpedientes.MontoTotal` | DOUBLE PRECISION | propuesta |
| `TbDatosEconomicosExpedientes.CosteExternoPrevisto` | DOUBLE(8) | — | `expedientes.TbDatosEconomicosExpedientes.CosteExternoPrevisto` | DOUBLE PRECISION (¿NUMERIC?) | propuesta |
| `TbDatosEconomicosExpedientes.CosteInternoPrevisto` | DOUBLE(8) | — | `expedientes.TbDatosEconomicosExpedientes.CosteInternoPrevisto` | DOUBLE PRECISION (¿NUMERIC?) | propuesta |
| `TbDatosEconomicosExpedientes.GastoMaximoPrevisto` | DOUBLE(8) | — | `expedientes.TbDatosEconomicosExpedientes.GastoMaximoPrevisto` | DOUBLE PRECISION | propuesta |
| `TbDatosEconomicosExpedientes.InversionPrevista` | DOUBLE(8) | — | `expedientes.TbDatosEconomicosExpedientes.InversionPrevista` | DOUBLE PRECISION | propuesta |
| `TbDatosEconomicosExpedientes.Observaciones` | TEXT(255) | F `ExpedienteJsonExporter.bas:834` | `expedientes.TbDatosEconomicosExpedientes.Observaciones` | VARCHAR(255) | propuesta |

## TbAuxNemotecnico

1 columnas · 0 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbAuxNemotecnico.Nemotecnico` | TEXT(255) | — | `expedientes.TbAuxNemotecnico.Nemotecnico` | VARCHAR(255) | propuesta |

## TbAusExpPostAGEDO

1 columnas · 203 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbAusExpPostAGEDO.ID` | LONG(4) | — | `expedientes.TbAusExpPostAGEDO.ID` | BIGINT | propuesta |

## TbAuxEstadosMartina

3 columnas · 143 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbAuxEstadosMartina.ID` | LONG(4) | — | `expedientes.TbAuxEstadosMartina.ID` | BIGINT | propuesta |
| `TbAuxEstadosMartina.EstadoMartina` | TEXT(255) | — | `expedientes.TbAuxEstadosMartina.EstadoMartina` | VARCHAR(255) | propuesta |
| `TbAuxEstadosMartina.Igual` | TEXT(2) | — | `expedientes.TbAuxEstadosMartina.Igual` | VARCHAR(2) · Sí/No como texto | propuesta |

## TbE2EExportBatch

10 columnas · 0 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbE2EExportBatch.IDBatch` | LONG(4) | — | `expedientes.TbE2EExportBatch.IDBatch` | BIGINT | propuesta |
| `TbE2EExportBatch.SessionId` | TEXT(100) | — | `expedientes.TbE2EExportBatch.SessionId` | VARCHAR(100) | propuesta |
| `TbE2EExportBatch.UsuarioConectado` | TEXT(255) | — | `expedientes.TbE2EExportBatch.UsuarioConectado` | VARCHAR(255) | propuesta |
| `TbE2EExportBatch.Estado` | TEXT(50) | — | `expedientes.TbE2EExportBatch.Estado` | VARCHAR(50) | propuesta |
| `TbE2EExportBatch.CreatedAt` | SHORT_DATE_TIME(8) | — | `expedientes.TbE2EExportBatch.CreatedAt` | TIMESTAMP | propuesta |
| `TbE2EExportBatch.StartedAt` | SHORT_DATE_TIME(8) | — | `expedientes.TbE2EExportBatch.StartedAt` | TIMESTAMP | propuesta |
| `TbE2EExportBatch.CompletedAt` | SHORT_DATE_TIME(8) | — | `expedientes.TbE2EExportBatch.CompletedAt` | TIMESTAMP | propuesta |
| `TbE2EExportBatch.TotalSeleccionados` | LONG(4) | — | `expedientes.TbE2EExportBatch.TotalSeleccionados` | BIGINT | propuesta |
| `TbE2EExportBatch.TotalExportados` | LONG(4) | — | `expedientes.TbE2EExportBatch.TotalExportados` | BIGINT | propuesta |
| `TbE2EExportBatch.ErrorMessage` | MEMO | — | `expedientes.TbE2EExportBatch.ErrorMessage` | TEXT | propuesta |

## TbE2EExportBatchDetalle

8 columnas · 0 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbE2EExportBatchDetalle.IDBatchDetalle` | LONG(4) | — | `expedientes.TbE2EExportBatchDetalle.IDBatchDetalle` | BIGINT | propuesta |
| `TbE2EExportBatchDetalle.IDBatch` | LONG(4) | — | `expedientes.TbE2EExportBatchDetalle.IDBatch` | BIGINT | propuesta |
| `TbE2EExportBatchDetalle.IDExpediente` | LONG(4) | — | `expedientes.TbE2EExportBatchDetalle.IDExpediente` | BIGINT | propuesta |
| `TbE2EExportBatchDetalle.OrdinalSeleccion` | LONG(4) | — | `expedientes.TbE2EExportBatchDetalle.OrdinalSeleccion` | BIGINT | propuesta |
| `TbE2EExportBatchDetalle.HashExportado` | TEXT(64) | — | `expedientes.TbE2EExportBatchDetalle.HashExportado` | VARCHAR(64) | propuesta |
| `TbE2EExportBatchDetalle.Estado` | TEXT(50) | — | `expedientes.TbE2EExportBatchDetalle.Estado` | VARCHAR(50) | propuesta |
| `TbE2EExportBatchDetalle.CreatedAt` | SHORT_DATE_TIME(8) | — | `expedientes.TbE2EExportBatchDetalle.CreatedAt` | TIMESTAMP | propuesta |
| `TbE2EExportBatchDetalle.ExportedAt` | SHORT_DATE_TIME(8) | — | `expedientes.TbE2EExportBatchDetalle.ExportedAt` | TIMESTAMP | propuesta |

## TbE2EExportSeleccionTemp

5 columnas · 5 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbE2EExportSeleccionTemp.IDTemp` | LONG(4) | — | `expedientes.TbE2EExportSeleccionTemp.IDTemp` | BIGINT | propuesta |
| `TbE2EExportSeleccionTemp.UsuarioConectado` | TEXT(255) | — | `expedientes.TbE2EExportSeleccionTemp.UsuarioConectado` | VARCHAR(255) | propuesta |
| `TbE2EExportSeleccionTemp.SessionId` | TEXT(100) | — | `expedientes.TbE2EExportSeleccionTemp.SessionId` | VARCHAR(100) | propuesta |
| `TbE2EExportSeleccionTemp.IDExpediente` | LONG(4) | — | `expedientes.TbE2EExportSeleccionTemp.IDExpediente` | BIGINT | propuesta |
| `TbE2EExportSeleccionTemp.CreatedAt` | SHORT_DATE_TIME(8) | — | `expedientes.TbE2EExportSeleccionTemp.CreatedAt` | TIMESTAMP | propuesta |

## TbE2EJsonDestinationUserConfig

5 columnas · 4 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbE2EJsonDestinationUserConfig.IDConfig` | LONG(4) | — | `expedientes.TbE2EJsonDestinationUserConfig.IDConfig` | BIGINT | propuesta |
| `TbE2EJsonDestinationUserConfig.UsuarioRed` | TEXT(255) | — | `expedientes.TbE2EJsonDestinationUserConfig.UsuarioRed` | VARCHAR(255) | propuesta |
| `TbE2EJsonDestinationUserConfig.RutaDestino` | MEMO | — | `expedientes.TbE2EJsonDestinationUserConfig.RutaDestino` | TEXT | propuesta |
| `TbE2EJsonDestinationUserConfig.CreatedAt` | SHORT_DATE_TIME(8) | — | `expedientes.TbE2EJsonDestinationUserConfig.CreatedAt` | TIMESTAMP | propuesta |
| `TbE2EJsonDestinationUserConfig.UpdatedAt` | SHORT_DATE_TIME(8) | — | `expedientes.TbE2EJsonDestinationUserConfig.UpdatedAt` | TIMESTAMP | propuesta |
