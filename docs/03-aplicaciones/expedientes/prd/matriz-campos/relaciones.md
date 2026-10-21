# Matriz campo a campo — relaciones

> Parte de la [matriz campo a campo](../matriz-campos.md). Reglas de mapeo, leyenda y límites de la fuente en ese documento. Commit fijado `ec6a8c3`.

## TbExpedientesAnexos

3 columnas · 710 filas en la extracción del snapshot · citada en 3 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesAnexos.IDDocumento` | LONG(4) | F `ExpedienteOperaciones.cls:1562` | `expedientes.TbExpedientesAnexos.IDDocumento` | BIGINT | propuesta |
| `TbExpedientesAnexos.IDExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:92` | `expedientes.TbExpedientesAnexos.IDExpediente` | BIGINT | propuesta |
| `TbExpedientesAnexos.NombreDocumento` | TEXT(255) | F `ExpedienteOperaciones.cls:1564` | `expedientes.TbExpedientesAnexos.NombreDocumento` | VARCHAR(255) | propuesta |

## TbExpedientesAnualidades

11 columnas · 174 filas en la extracción del snapshot · citada en 5 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesAnualidades.IDAnualidad` | LONG(4) | F `ExpedienteOperaciones.cls:1266` | `expedientes.TbExpedientesAnualidades.IDAnualidad` | BIGINT | propuesta |
| `TbExpedientesAnualidades.IDExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:92` | `expedientes.TbExpedientesAnualidades.IDExpediente` | BIGINT | propuesta |
| `TbExpedientesAnualidades.Año` | INT(2) | F `ExpedienteOperaciones.cls:1268` | `expedientes.TbExpedientesAnualidades.Año` | SMALLINT | propuesta |
| `TbExpedientesAnualidades.BIIVA` | DOUBLE(8) | F `ExpedienteOperaciones.cls:1269` | `expedientes.TbExpedientesAnualidades.BIIVA` | DOUBLE PRECISION | propuesta |
| `TbExpedientesAnualidades.BIIPSI` | DOUBLE(8) | F `ExpedienteOperaciones.cls:1270` | `expedientes.TbExpedientesAnualidades.BIIPSI` | DOUBLE PRECISION | propuesta |
| `TbExpedientesAnualidades.BIIGIC` | DOUBLE(8) | F `ExpedienteOperaciones.cls:1271` | `expedientes.TbExpedientesAnualidades.BIIGIC` | DOUBLE PRECISION | propuesta |
| `TbExpedientesAnualidades.BIEXENTA` | DOUBLE(8) | F `ExpedienteOperaciones.cls:1272` | `expedientes.TbExpedientesAnualidades.BIEXENTA` | DOUBLE PRECISION | propuesta |
| `TbExpedientesAnualidades.IVA` | DOUBLE(8) | F `ExpedienteOperaciones.cls:1273` | `expedientes.TbExpedientesAnualidades.IVA` | DOUBLE PRECISION | propuesta |
| `TbExpedientesAnualidades.IPSI` | DOUBLE(8) | F `ExpedienteOperaciones.cls:1274` | `expedientes.TbExpedientesAnualidades.IPSI` | DOUBLE PRECISION | propuesta |
| `TbExpedientesAnualidades.IGIC` | DOUBLE(8) | F `ExpedienteOperaciones.cls:1275` | `expedientes.TbExpedientesAnualidades.IGIC` | DOUBLE PRECISION | propuesta |
| `TbExpedientesAnualidades.PeriodoFacturacion` | TEXT(255) | F `ExpedienteOperaciones.cls:1276` | `expedientes.TbExpedientesAnualidades.PeriodoFacturacion` | VARCHAR(255) | propuesta |

## TbExpedientesCadenaContratacion

9 columnas · 0 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesCadenaContratacion.ID` | LONG(4) | — | `expedientes.TbExpedientesCadenaContratacion.ID` | BIGINT | propuesta |
| `TbExpedientesCadenaContratacion.IDPadre` | LONG(4) | — | `expedientes.TbExpedientesCadenaContratacion.IDPadre` | BIGINT | propuesta |
| `TbExpedientesCadenaContratacion.IDExpediente` | LONG(4) | — | `expedientes.TbExpedientesCadenaContratacion.IDExpediente` | BIGINT | propuesta |
| `TbExpedientesCadenaContratacion.IDSuministrador` | LONG(4) | — | `expedientes.TbExpedientesCadenaContratacion.IDSuministrador` | BIGINT | propuesta |
| `TbExpedientesCadenaContratacion.AplicaCalidad` | TEXT(2) | — | `expedientes.TbExpedientesCadenaContratacion.AplicaCalidad` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientesCadenaContratacion.AplicaRiesgos` | TEXT(2) | — | `expedientes.TbExpedientesCadenaContratacion.AplicaRiesgos` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientesCadenaContratacion.AplicaContratosClasificados` | TEXT(2) | — | `expedientes.TbExpedientesCadenaContratacion.AplicaContratosClasificados` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientesCadenaContratacion.AplicaHPS` | TEXT(255) | — | `expedientes.TbExpedientesCadenaContratacion.AplicaHPS` | VARCHAR(255) | propuesta |
| `TbExpedientesCadenaContratacion.Observaciones` | MEMO | — | `expedientes.TbExpedientesCadenaContratacion.Observaciones` | TEXT | propuesta |

## TbExpedientesCodigoCompras

3 columnas · 0 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesCodigoCompras.ID` | LONG(4) | F `ExpedienteOperaciones.cls:1341` | `expedientes.TbExpedientesCodigoCompras.ID` | BIGINT | propuesta |
| `TbExpedientesCodigoCompras.IDExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:92` | `expedientes.TbExpedientesCodigoCompras.IDExpediente` | BIGINT | propuesta |
| `TbExpedientesCodigoCompras.CodCompras` | TEXT(255) | F `ExpedienteOperaciones.cls:1343` | `expedientes.TbExpedientesCodigoCompras.CodCompras` | VARCHAR(255) | propuesta |

## TbExpedientesComerciales

3 columnas · 333 filas en la extracción del snapshot · citada en 5 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesComerciales.IDComercialExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:1307` | `expedientes.TbExpedientesComerciales.IDComercialExpediente` | BIGINT | propuesta |
| `TbExpedientesComerciales.IDComercial` | LONG(4) | F `ExpedienteOperaciones.cls:1309` | `expedientes.TbExpedientesComerciales.IDComercial` | BIGINT | propuesta |
| `TbExpedientesComerciales.IDExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:92` | `expedientes.TbExpedientesComerciales.IDExpediente` | BIGINT | propuesta |

## TbExpedientesCPVs

3 columnas · 429 filas en la extracción del snapshot · citada en 6 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesCPVs.IDCPVExpediente` | LONG(4) | F `ExpedienteCPVOperaciones.cls:42` | `expedientes.TbExpedientesCPVs.IDCPVExpediente` | BIGINT | propuesta |
| `TbExpedientesCPVs.IDCPV` | LONG(4) | F `ExpedienteCPVOperaciones.cls:40` | `expedientes.TbExpedientesCPVs.IDCPV` | BIGINT | propuesta |
| `TbExpedientesCPVs.IDExpediente` | LONG(4) | F `ExpedienteCPVOperaciones.cls:41` | `expedientes.TbExpedientesCPVs.IDExpediente` | BIGINT | propuesta |

## TbExpedientesE2E

7 columnas · 0 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesE2E.IDExpediente` | LONG(4) | — | `expedientes.TbExpedientesE2E.IDExpediente` | BIGINT | propuesta |
| `TbExpedientesE2E.HashPayload` | TEXT(255) | — | `expedientes.TbExpedientesE2E.HashPayload` | VARCHAR(255) | propuesta |
| `TbExpedientesE2E.Estado` | TEXT(20) | — | `expedientes.TbExpedientesE2E.Estado` | VARCHAR(20) | propuesta |
| `TbExpedientesE2E.FechaCreacion` | SHORT_DATE_TIME(8) | — | `expedientes.TbExpedientesE2E.FechaCreacion` | TIMESTAMP | propuesta |
| `TbExpedientesE2E.FechaModificacion` | SHORT_DATE_TIME(8) | — | `expedientes.TbExpedientesE2E.FechaModificacion` | TIMESTAMP | propuesta |
| `TbExpedientesE2E.UsuarioCreacion` | TEXT(100) | — | `expedientes.TbExpedientesE2E.UsuarioCreacion` | VARCHAR(100) | propuesta |
| `TbExpedientesE2E.UsuarioModificacion` | TEXT(100) | — | `expedientes.TbExpedientesE2E.UsuarioModificacion` | VARCHAR(100) | propuesta |

## TbExpedientesHitos

6 columnas · 46 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesHitos.IDHitoExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:1804` | `expedientes.TbExpedientesHitos.IDHitoExpediente` | BIGINT | propuesta |
| `TbExpedientesHitos.IDExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:92` | `expedientes.TbExpedientesHitos.IDExpediente` | BIGINT | propuesta |
| `TbExpedientesHitos.Descripcion` | TEXT(255) | F `ExpedienteOperaciones.cls:1806` | `expedientes.TbExpedientesHitos.Descripcion` | VARCHAR(255) | propuesta |
| `TbExpedientesHitos.FechaHito` | SHORT_DATE_TIME(8) | F `ExpedienteOperaciones.cls:1807` | `expedientes.TbExpedientesHitos.FechaHito` | TIMESTAMP | propuesta |
| `TbExpedientesHitos.FechaGarantiaHito` | SHORT_DATE_TIME(8) | F `ExpedienteOperaciones.cls:1810` | `expedientes.TbExpedientesHitos.FechaGarantiaHito` | TIMESTAMP | propuesta |
| `TbExpedientesHitos.Importe` | DOUBLE(8) | F `ExpedienteOperaciones.cls:1813` | `expedientes.TbExpedientesHitos.Importe` | DOUBLE PRECISION (¿NUMERIC?) | propuesta |

## TbExpedientesJefaturas

3 columnas · 0 filas en la extracción del snapshot · citada en 2 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesJefaturas.IDJefaturaExpediente` | LONG(4) | — | `expedientes.TbExpedientesJefaturas.IDJefaturaExpediente` | BIGINT | propuesta |
| `TbExpedientesJefaturas.IDJefatura` | LONG(4) | S `constructor.bas:3451` | `expedientes.TbExpedientesJefaturas.IDJefatura` | BIGINT | propuesta |
| `TbExpedientesJefaturas.IDExpediente` | LONG(4) | F `ExpedienteJsonExporter.bas:817` | `expedientes.TbExpedientesJefaturas.IDExpediente` | BIGINT | propuesta |

## TbExpedientesJuridicas

6 columnas · 417 filas en la extracción del snapshot · citada en 3 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesJuridicas.IDExpedienteJuridica` | LONG(4) | — | `expedientes.TbExpedientesJuridicas.IDExpedienteJuridica` | BIGINT | propuesta |
| `TbExpedientesJuridicas.IDExpediente` | LONG(4) | F `ExpedienteJsonExporter.bas:817` | `expedientes.TbExpedientesJuridicas.IDExpediente` | BIGINT | propuesta |
| `TbExpedientesJuridicas.IDJuridica` | LONG(4) | S `Instalador.bas:1237` | `expedientes.TbExpedientesJuridicas.IDJuridica` | BIGINT | propuesta |
| `TbExpedientesJuridicas.IDSuministrador` | LONG(4) | F `ExpedienteJsonExporter.bas:765` | `expedientes.TbExpedientesJuridicas.IDSuministrador` | BIGINT | propuesta |
| `TbExpedientesJuridicas.ContratistaPrincipal` | TEXT(2) | F `Instalador.bas:1255` | `expedientes.TbExpedientesJuridicas.ContratistaPrincipal` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientesJuridicas.SubContratista` | TEXT(2) | F `Instalador.bas:1256` | `expedientes.TbExpedientesJuridicas.SubContratista` | VARCHAR(2) · Sí/No como texto | propuesta |

## TbExpedientesLugaresEjecucion

3 columnas · 194 filas en la extracción del snapshot · citada en 5 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesLugaresEjecucion.IDExpedienteLugarEjecucion` | LONG(4) | F `ExpedienteOperaciones.cls:1407` | `expedientes.TbExpedientesLugaresEjecucion.IDExpedienteLugarEjecucion` | BIGINT | propuesta |
| `TbExpedientesLugaresEjecucion.IDExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:92` | `expedientes.TbExpedientesLugaresEjecucion.IDExpediente` | BIGINT | propuesta |
| `TbExpedientesLugaresEjecucion.IDLugarEjecucion` | LONG(4) | F `ExpedienteOperaciones.cls:1409` | `expedientes.TbExpedientesLugaresEjecucion.IDLugarEjecucion` | BIGINT | propuesta |

## TbExpedientesModificados

6 columnas · 37 filas en la extracción del snapshot · citada en 5 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesModificados.IDExpedienteModificado` | LONG(4) | F `ExpedienteOperaciones.cls:1907` | `expedientes.TbExpedientesModificados.IDExpedienteModificado` | BIGINT | propuesta |
| `TbExpedientesModificados.IDExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:92` | `expedientes.TbExpedientesModificados.IDExpediente` | BIGINT | propuesta |
| `TbExpedientesModificados.NModificado` | TEXT(255) | F `ExpedienteModificadoOperaciones.cls:69` | `expedientes.TbExpedientesModificados.NModificado` | VARCHAR(255) | propuesta |
| `TbExpedientesModificados.FechaFirmaModificado` | SHORT_DATE_TIME(8) | F `ExpedienteModificadoOperaciones.cls:71` | `expedientes.TbExpedientesModificados.FechaFirmaModificado` | TIMESTAMP | propuesta |
| `TbExpedientesModificados.FechaFinModificado` | SHORT_DATE_TIME(8) | F `ExpedienteModificadoOperaciones.cls:78` | `expedientes.TbExpedientesModificados.FechaFinModificado` | TIMESTAMP | propuesta |
| `TbExpedientesModificados.Descripcion` | MEMO | F `ExpedienteModificadoOperaciones.cls:68` | `expedientes.TbExpedientesModificados.Descripcion` | TEXT | propuesta |

## TbExpedientesPECAL

3 columnas · 366 filas en la extracción del snapshot · citada en 5 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesPECAL.IDPECALExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:1443` | `expedientes.TbExpedientesPECAL.IDPECALExpediente` | BIGINT | propuesta |
| `TbExpedientesPECAL.IDExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:92` | `expedientes.TbExpedientesPECAL.IDExpediente` | BIGINT | propuesta |
| `TbExpedientesPECAL.IDPECAL` | LONG(4) | F `ExpedienteOperaciones.cls:1445` | `expedientes.TbExpedientesPECAL.IDPECAL` | BIGINT | propuesta |

## TbExpedientesRACS

3 columnas · 193 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesRACS.IDRacExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:1479` | `expedientes.TbExpedientesRACS.IDRacExpediente` | BIGINT | propuesta |
| `TbExpedientesRACS.IDExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:92` | `expedientes.TbExpedientesRACS.IDExpediente` | BIGINT | propuesta |
| `TbExpedientesRACS.IDRAC` | LONG(4) | F `ExpedienteOperaciones.cls:1481` | `expedientes.TbExpedientesRACS.IDRAC` | BIGINT | propuesta |

## TbExpedientesResponsables

6 columnas · 730 filas en la extracción del snapshot · citada en 8 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesResponsables.IDExpedienteResponsable` | LONG(4) | F `Form_FormExpedienteEntidades.cls:434` | `expedientes.TbExpedientesResponsables.IDExpedienteResponsable` | BIGINT | propuesta |
| `TbExpedientesResponsables.IdExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:92` | `expedientes.TbExpedientesResponsables.IdExpediente` | BIGINT | propuesta |
| `TbExpedientesResponsables.IdUsuario` | LONG(4) | F `Form_FormExpedienteEntidades.cls:435` | `expedientes.TbExpedientesResponsables.IdUsuario` | BIGINT | propuesta |
| `TbExpedientesResponsables.CorreoSiempre` | TEXT(2) | F `Form_FormExpedienteEntidades.cls:437` | `expedientes.TbExpedientesResponsables.CorreoSiempre` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientesResponsables.EsJefeProyecto` | TEXT(2) | F `Form_FormExpedienteEntidades.cls:436` | `expedientes.TbExpedientesResponsables.EsJefeProyecto` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientesResponsables.esPreventa` | TEXT(2) | F `Form_FormExpedienteEntidades.cls:438` | `expedientes.TbExpedientesResponsables.esPreventa` | VARCHAR(2) · Sí/No como texto | propuesta |

## TbExpedientesSuministradores

7 columnas · 713 filas en la extracción del snapshot · citada en 9 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesSuministradores.IDExpedienteSuministrador` | LONG(4) | F `Instalador.bas:1252` | `expedientes.TbExpedientesSuministradores.IDExpedienteSuministrador` | BIGINT | propuesta |
| `TbExpedientesSuministradores.IDExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:92` | `expedientes.TbExpedientesSuministradores.IDExpediente` | BIGINT | propuesta |
| `TbExpedientesSuministradores.IDSuministrador` | LONG(4) | F `ExpedienteJsonExporter.bas:765` | `expedientes.TbExpedientesSuministradores.IDSuministrador` | BIGINT | propuesta |
| `TbExpedientesSuministradores.Descripcon` | MEMO | S `ExpedienteSuministrador.cls:17` | `expedientes.TbExpedientesSuministradores.Descripcon` | TEXT | propuesta |
| `TbExpedientesSuministradores.ContratistaPrincipal` | TEXT(2) | F `ExpedienteOperaciones.cls:387` | `expedientes.TbExpedientesSuministradores.ContratistaPrincipal` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientesSuministradores.SubContratista` | TEXT(2) | F `Instalador.bas:1256` | `expedientes.TbExpedientesSuministradores.SubContratista` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientesSuministradores.IDPadre` | LONG(4) | F `ExpedienteJsonExporter.bas:766` | `expedientes.TbExpedientesSuministradores.IDPadre` | BIGINT | propuesta |
