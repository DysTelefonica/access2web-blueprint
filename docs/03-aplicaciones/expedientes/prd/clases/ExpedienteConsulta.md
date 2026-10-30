# ExpedienteConsulta

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/classes/ExpedienteConsulta.cls` (1320 líneas) · **PRD:** [README](../README.md)

Vista plana de solo lectura de un expediente para listados y exportaciones: 43 propiedades `Get`; al menos `FechaFinGarantia` devuelve su valor en caché o, si está vacío, lo toma del `Expediente` asociado (`:455-470`). Las citas son `ExpedienteConsulta.cls:línea`.

**Propiedades:** `IDExpediente`, `Nemotecnico`, `CodExpLargo`, `CodExp`, `Tipo`, `IDExpedientePadre`, `Titulo`, `ObjetoContrato`, `Contratistas`, `Ejercito`, `ESTADO`, `JefeProyecto`, `Calidad`, `Seguridad`, `Adjudicado`, `EnPeriodoDeAdjudicacion`, `ImporteLicitacion`, `ImporteContratacion`, `CodProyecto`, `CodS4H`, `Clasificacion`, `OrganoContratacion`, `OficinaPrograma`, `AccesoSharepoint`, `Observaciones`, `Ambito`, `NPedido`, `Ordinal`, `PECAL`, `CadenaPecal`, `POSTAGEDO`, `AGEDYSAplica`, `AGEDYSGenerico`, `HPSAplica`, las tres fechas de contrato y garantía, los indicadores `EsAM`, `EsLote`, `EsBasado` y `EsExpediente`, `ExpedienteEntidad` y `ColCampos`.

## Uso

**Ningún fichero de `src` instancia la clase ni cita su nombre fuera de ella.** Aparece como fuente en `docs/DISCOVERY/README.md` del repositorio de origen (sección de lectura, listados y filtros). Los listados vivos de [FormExpedientesGestion](../formularios/Form_FormExpedientesGestion.md) trabajan con `ExpedienteCompleto`.

## Comportamiento que la migración debe conservar

Ninguno verificable en `src`: el conjunto de columnas que expone puede servir de referencia del modelo de lectura de listados, pero no hay constancia de que se use.

## Pendientes

- Confirmar que no se usa desde el binario (informes o consultas guardadas), fuera de la lectura estática de `src`.
