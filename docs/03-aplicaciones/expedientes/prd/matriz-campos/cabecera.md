# Matriz campo a campo — cabecera del expediente

> Parte de la [matriz campo a campo](../matriz-campos.md). Reglas de mapeo, leyenda y límites de la fuente en ese documento. Commit fijado `ec6a8c3`.

## TbExpedientes

61 columnas · 453 filas en la extracción del snapshot · citada en 20 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientes.IDExpediente` | LONG(4) | F `ExpedienteOperaciones.cls:92` | `expedientes.TbExpedientes.IDExpediente` | BIGINT | propuesta |
| `TbExpedientes.IDExpedientePadre` | LONG(4) | F `ExpedienteOperaciones.cls:100` | `expedientes.TbExpedientes.IDExpedientePadre` | BIGINT | propuesta |
| `TbExpedientes.Nemotecnico` | TEXT(255) | F `Expediente.cls:864` | `expedientes.TbExpedientes.Nemotecnico` | VARCHAR(255) | propuesta |
| `TbExpedientes.Titulo` | MEMO | F `ExpedienteOperaciones.cls:112` | `expedientes.TbExpedientes.Titulo` | TEXT | propuesta |
| `TbExpedientes.ImporteLicitacion` | DOUBLE(8) | F `ExpedienteOperaciones.cls:125` | `expedientes.TbExpedientes.ImporteLicitacion` | DOUBLE PRECISION (¿NUMERIC?) | propuesta |
| `TbExpedientes.ImporteContratacion` | DOUBLE(8) | F `ExpedienteOperaciones.cls:131` | `expedientes.TbExpedientes.ImporteContratacion` | DOUBLE PRECISION (¿NUMERIC?) | propuesta |
| `TbExpedientes.CodProyecto` | TEXT(255) | F `ExpedienteOperaciones.cls:138` | `expedientes.TbExpedientes.CodProyecto` | VARCHAR(255) | propuesta |
| `TbExpedientes.CodExp` | TEXT(255) | F `ExpedienteOperaciones.cls:144` | `expedientes.TbExpedientes.CodExp` | VARCHAR(255) | propuesta |
| `TbExpedientes.CodExpLargo` | TEXT(255) | F `ExpedienteOperaciones.cls:150` | `expedientes.TbExpedientes.CodExpLargo` | VARCHAR(255) | propuesta |
| `TbExpedientes.CodS4H` | TEXT(255) | F `ExpedienteOperaciones.cls:156` | `expedientes.TbExpedientes.CodS4H` | VARCHAR(255) | propuesta |
| `TbExpedientes.FechaInicioContrato` | SHORT_DATE_TIME(8) | F `ExpedienteOperaciones.cls:325` | `expedientes.TbExpedientes.FechaInicioContrato` | TIMESTAMP | propuesta |
| `TbExpedientes.FechaFinContrato` | SHORT_DATE_TIME(8) | F `ExpedienteOperaciones.cls:331` | `expedientes.TbExpedientes.FechaFinContrato` | TIMESTAMP | propuesta |
| `TbExpedientes.FechaFinGarantia` | SHORT_DATE_TIME(8) | F `ExpedienteOperaciones.cls:361` | `expedientes.TbExpedientes.FechaFinGarantia` | TIMESTAMP | propuesta |
| `TbExpedientes.EsAM` | TEXT(2) | F `ExpedienteOperaciones.cls:162` | `expedientes.TbExpedientes.EsAM` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientes.EsLote` | TEXT(2) | F `ExpedienteOperaciones.cls:168` | `expedientes.TbExpedientes.EsLote` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientes.EsBasado` | TEXT(255) | F `ExpedienteOperaciones.cls:174` | `expedientes.TbExpedientes.EsBasado` | VARCHAR(255) | propuesta |
| `TbExpedientes.EsExpediente` | TEXT(2) | F `ExpedienteOperaciones.cls:180` | `expedientes.TbExpedientes.EsExpediente` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientes.Ordinal` | TEXT(255) | F `ExpedienteOperaciones.cls:186` | `expedientes.TbExpedientes.Ordinal` | VARCHAR(255) | propuesta |
| `TbExpedientes.IdGradoClasificacion` | LONG(4) | F `ExpedienteOperaciones.cls:192` | `expedientes.TbExpedientes.IdGradoClasificacion` | BIGINT | propuesta |
| `TbExpedientes.IDOrganoContratacion` | LONG(4) | F `ExpedienteOperaciones.cls:198` | `expedientes.TbExpedientes.IDOrganoContratacion` | BIGINT | propuesta |
| `TbExpedientes.IDOficinaPrograma` | LONG(4) | F `ExpedienteOperaciones.cls:204` | `expedientes.TbExpedientes.IDOficinaPrograma` | BIGINT | propuesta |
| `TbExpedientes.IDEjercito` | LONG(4) | F `ExpedienteOperaciones.cls:210` | `expedientes.TbExpedientes.IDEjercito` | BIGINT | propuesta |
| `TbExpedientes.AccesoSharepoint` | MEMO | F `ExpedienteOperaciones.cls:216` | `expedientes.TbExpedientes.AccesoSharepoint` | TEXT | propuesta |
| `TbExpedientes.Observaciones` | MEMO | F `ExpedienteOperaciones.cls:222` | `expedientes.TbExpedientes.Observaciones` | TEXT | propuesta |
| `TbExpedientes.FechaCreacion` | SHORT_DATE_TIME(8) | F `UltimoCambioOperaciones.cls:67` | `expedientes.TbExpedientes.FechaCreacion` | TIMESTAMP | propuesta |
| `TbExpedientes.IDUsuarioCreacion` | TEXT(255) | F `UltimoCambioOperaciones.cls:68` | `expedientes.TbExpedientes.IDUsuarioCreacion` | VARCHAR(255) | propuesta |
| `TbExpedientes.FechaUltimoCambio` | SHORT_DATE_TIME(8) | F `UltimoCambioOperaciones.cls:72` | `expedientes.TbExpedientes.FechaUltimoCambio` | TIMESTAMP | propuesta |
| `TbExpedientes.IDUsuarioUltimoCambio` | TEXT(255) | F `UltimoCambioOperaciones.cls:73` | `expedientes.TbExpedientes.IDUsuarioUltimoCambio` | VARCHAR(255) | propuesta |
| `TbExpedientes.IDEstado` | LONG(4) | S `ExpedienteJsonExporter.bas:968` | `expedientes.TbExpedientes.IDEstado` | BIGINT | propuesta |
| `TbExpedientes.Ambito` | TEXT(10) | F `ExpedienteOperaciones.cls:228` | `expedientes.TbExpedientes.Ambito` | VARCHAR(10) | propuesta |
| `TbExpedientes.NPedido` | TEXT(255) | F `ExpedienteOperaciones.cls:234` | `expedientes.TbExpedientes.NPedido` | VARCHAR(255) | propuesta |
| `TbExpedientes.IDResponsableCalidad` | LONG(4) | F `ExpedienteOperaciones.cls:240` | `expedientes.TbExpedientes.IDResponsableCalidad` | BIGINT | propuesta |
| `TbExpedientes.Adjudicado` | TEXT(2) | F `ExpedienteJsonExporter.bas:837` | `expedientes.TbExpedientes.Adjudicado` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientes.EnPeriodoDeAdjudicacion` | TEXT(2) | F `ExpedienteJsonExporter.bas:838` | `expedientes.TbExpedientes.EnPeriodoDeAdjudicacion` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientes.Tipo` | TEXT(255) | F `ExpedienteOperaciones.cls:252` | `expedientes.TbExpedientes.Tipo` | VARCHAR(255) | propuesta |
| `TbExpedientes.TipoInforme` | TEXT(255) | F `ExpedienteOperaciones.cls:276` | `expedientes.TbExpedientes.TipoInforme` | VARCHAR(255) | propuesta |
| `TbExpedientes.AGEDYSAplica` | TEXT(2) | F `ExpedienteOperaciones.cls:258` | `expedientes.TbExpedientes.AGEDYSAplica` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientes.AGEDYSGenerico` | TEXT(2) | F `ExpedienteOperaciones.cls:264` | `expedientes.TbExpedientes.AGEDYSGenerico` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientes.HPSAplica` | TEXT(2) | F `ExpedienteOperaciones.cls:270` | `expedientes.TbExpedientes.HPSAplica` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientes.CadenaPecal` | TEXT(255) | S `Expediente.cls:50` | `expedientes.TbExpedientes.CadenaPecal` | VARCHAR(255) | propuesta |
| `TbExpedientes.Pecal` | TEXT(2) | F `ExpedienteOperaciones.cls:1616` | `expedientes.TbExpedientes.Pecal` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientes.POSTAGEDO` | TEXT(2) | F `ExpedienteOperaciones.cls:282` | `expedientes.TbExpedientes.POSTAGEDO` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientes.FECHAPREOFERTA` | SHORT_DATE_TIME(8) | F `ExpedienteOperaciones.cls:301` | `expedientes.TbExpedientes.FECHAPREOFERTA` | TIMESTAMP | propuesta |
| `TbExpedientes.APLICAESTADO` | TEXT(2) | F `ExpedienteOperaciones.cls:288` | `expedientes.TbExpedientes.APLICAESTADO` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientes.FECHAINICIOLICITACION` | SHORT_DATE_TIME(8) | F `ExpedienteOperaciones.cls:295` | `expedientes.TbExpedientes.FECHAINICIOLICITACION` | TIMESTAMP | propuesta |
| `TbExpedientes.FECHAOFERTA` | SHORT_DATE_TIME(8) | F `ExpedienteOperaciones.cls:307` | `expedientes.TbExpedientes.FECHAOFERTA` | TIMESTAMP | propuesta |
| `TbExpedientes.FECHAADJUDICACION` | SHORT_DATE_TIME(8) | F `ExpedienteOperaciones.cls:343` | `expedientes.TbExpedientes.FECHAADJUDICACION` | TIMESTAMP | propuesta |
| `TbExpedientes.FECHAFIRMACONTRATO` | SHORT_DATE_TIME(8) | F `ExpedienteOperaciones.cls:349` | `expedientes.TbExpedientes.FECHAFIRMACONTRATO` | TIMESTAMP | propuesta |
| `TbExpedientes.GARANTIAMESES` | TEXT(255) | F `ExpedienteOperaciones.cls:337` | `expedientes.TbExpedientes.GARANTIAMESES` | VARCHAR(255) | propuesta |
| `TbExpedientes.FECHACERTIFICACION` | SHORT_DATE_TIME(8) | F `ExpedienteOperaciones.cls:355` | `expedientes.TbExpedientes.FECHACERTIFICACION` | TIMESTAMP | propuesta |
| `TbExpedientes.FECHAPERDIDA` | SHORT_DATE_TIME(8) | F `ExpedienteOperaciones.cls:319` | `expedientes.TbExpedientes.FECHAPERDIDA` | TIMESTAMP | propuesta |
| `TbExpedientes.FECHADESESTIMADA` | SHORT_DATE_TIME(8) | F `ExpedienteOperaciones.cls:313` | `expedientes.TbExpedientes.FECHADESESTIMADA` | TIMESTAMP | propuesta |
| `TbExpedientes.ESTADO` | TEXT(255) | F `ExpedienteOperaciones.cls:368` | `expedientes.TbExpedientes.ESTADO` | VARCHAR(255) | propuesta |
| `TbExpedientes.CodigoActividad` | TEXT(255) | F `ExpedienteOperaciones.cls:374` | `expedientes.TbExpedientes.CodigoActividad` | VARCHAR(255) | propuesta |
| `TbExpedientes.AplicaTareaS4H` | TEXT(2) | F `ExpedienteOperaciones.cls:380` | `expedientes.TbExpedientes.AplicaTareaS4H` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientes.ContratistaPrincipal` | TEXT(2) | F `ExpedienteOperaciones.cls:387` | `expedientes.TbExpedientes.ContratistaPrincipal` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientes.IDResponsableSeguridad` | LONG(4) | F `ExpedienteOperaciones.cls:246` | `expedientes.TbExpedientes.IDResponsableSeguridad` | BIGINT | propuesta |
| `TbExpedientes.ObjetoContrato` | MEMO | F `ExpedienteOperaciones.cls:118` | `expedientes.TbExpedientes.ObjetoContrato` | TEXT | propuesta |
| `TbExpedientes.OrdinalE2E` | LONG(4) | F `ExpedienteJsonExporter.bas:818` | `expedientes.TbExpedientes.OrdinalE2E` | BIGINT | propuesta |
| `TbExpedientes.HashActual` | TEXT(64) | S `E2EEnlacesUIHelper.bas:173` | `expedientes.TbExpedientes.HashActual` | VARCHAR(64) | propuesta |
| `TbExpedientes.HashUltimaExportacion` | TEXT(64) | S `E2EEnlacesUIHelper.bas:174` | `expedientes.TbExpedientes.HashUltimaExportacion` | VARCHAR(64) | propuesta |

## TbExpedientesConEntidades

21 columnas · 451 filas en la extracción del snapshot · citada en 7 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbExpedientesConEntidades.IDExpediente` | LONG(4) | F `ExpedienteEntidadOperaciones.cls:60` | `expedientes.TbExpedientesConEntidades.IDExpediente` | BIGINT | propuesta |
| `TbExpedientesConEntidades.Clasificacion` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:74` | `expedientes.TbExpedientesConEntidades.Clasificacion` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.OrganoContratacion` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:75` | `expedientes.TbExpedientesConEntidades.OrganoContratacion` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.OficinaPrograma` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:76` | `expedientes.TbExpedientesConEntidades.OficinaPrograma` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.Ejercito` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:77` | `expedientes.TbExpedientesConEntidades.Ejercito` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.Estado` | TEXT(255) | F `ExpedienteOperaciones.cls:368` | `expedientes.TbExpedientesConEntidades.Estado` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.ResponsableCalidad` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:78` | `expedientes.TbExpedientesConEntidades.ResponsableCalidad` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.ResponsableSeguridad` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:79` | `expedientes.TbExpedientesConEntidades.ResponsableSeguridad` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.CadenaPecal` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:85` | `expedientes.TbExpedientesConEntidades.CadenaPecal` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.Pecal` | TEXT(2) | F `ExpedienteEntidadOperaciones.cls:86` | `expedientes.TbExpedientesConEntidades.Pecal` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbExpedientesConEntidades.CadenaContratistas` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:92` | `expedientes.TbExpedientesConEntidades.CadenaContratistas` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.CadenaSubContratistas` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:97` | `expedientes.TbExpedientesConEntidades.CadenaSubContratistas` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.CadenaSuministradores` | TEXT(255) | S `FUNCIONES UTILES.bas:115` | `expedientes.TbExpedientesConEntidades.CadenaSuministradores` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.CadenaComerciales` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:102` | `expedientes.TbExpedientesConEntidades.CadenaComerciales` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.CadenaJPs` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:108` | `expedientes.TbExpedientesConEntidades.CadenaJPs` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.CadenaRACs` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:113` | `expedientes.TbExpedientesConEntidades.CadenaRACs` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.CadenaCorreoRACs` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:114` | `expedientes.TbExpedientesConEntidades.CadenaCorreoRACs` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.CadenaHitos` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:119` | `expedientes.TbExpedientesConEntidades.CadenaHitos` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.TipoParaLista` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:80` | `expedientes.TbExpedientesConEntidades.TipoParaLista` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.CadenaLugares` | TEXT(255) | F `ExpedienteEntidadOperaciones.cls:126` | `expedientes.TbExpedientesConEntidades.CadenaLugares` | VARCHAR(255) | propuesta |
| `TbExpedientesConEntidades.CadenaJuridicas` | TEXT(255) | — | `expedientes.TbExpedientesConEntidades.CadenaJuridicas` | VARCHAR(255) | propuesta |
