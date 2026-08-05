# Expedientes — modelo físico y diccionario de datos

## Autoridad y fecha

Fuente física: `C:\00repos\datos\Expedientes_datos.accdb`, resuelta explícitamente el 2026-08-05 mediante Dysflow `list_tables`, `get_schema`, `get_relationships` y `count_rows`. Se detectaron **49 tablas de usuario**. El frontend `staging\Expedientes.accdb` no sustituye al backend autoritativo.

## Diccionario completo

El diccionario fuente completo, con **cada tabla y cada campo**, está cosechado en [Estructura_Datos.md](../../../../documentacion/OPENSPEC/00_EXPEDIENTES/docs/ERD/Estructura_Datos.md). Ese documento se conserva como fuente primaria del inventario de columnas; esta página añade la normalización, procedencia, relaciones y perfilado exigidos para la migración. No se han copiado valores de filas.

### Tablas de negocio principales

| Grupo | Tablas |
|---|---|
| Agregado | `TbExpedientes`, `TbExpedientesConEntidades`, `TbDatosEconomicosExpedientes` |
| Relaciones | `TbExpedientesAnualidades`, `Comerciales`, `CPVs`, `RACS`, `PECAL`, `LugaresEjecucion`, `Responsables`, `Suministradores`, `Jefaturas`, `Juridicas`, `Hitos`, `Modificados`, `Anexos`, `CodigoCompras`, `CadenaContratacion` |
| Catálogos | `TbComerciales`, `TbCPV`, `TbEjercitos`, `TbEstados`, `TbGradosClasificacion`, `TbJefaturas`, `TbJuridicas`, `TbLugaresEjecucion`, `TbOficinasPrograma`, `TbOrganosContratacion`, `TbPECAL`, `TbRACS`, `TbResponsablesPorRol`, `TbSuministradores` |
| Auditoría/preferencias | `TbCambios`, `TbUltimoCambio`, `TbConfMostrarEstado` |
| Históricos/auxiliares | `Copia de TbExpedientes`, `Copia de TbExpedientesConEntidades`, `ListaPrevia`, `TbExpedientes_antes`, `TbAusExpPostAGEDO`, `TbAuxEstadosMartina`, `TbAuxNemotecnico`, `TbExpAgedys`, `TbExpAGEDYS1` |
| E2E/operación | `TbE2EExportBatch`, `TbE2EExportBatchDetalle`, `TbE2EExportSeleccionTemp`, `TbE2EJsonDestinationUserConfig`, `TbExpedientesE2E` |

## Relaciones físicas confirmadas

`TbExpedientes.IDExpediente` relaciona con las tablas de hijos y joins; `TbSuministradores.IDSuministrador` relaciona con `TbExpedientesSuministradores`. `IDExpedientePadre` es además una relación jerárquica autorreferente inferida por código y datos, aunque no aparece como FK física explícita en todas las relaciones Access.

## Semántica Access que debe conservarse

- Tipos 1/3/4/7/8/10/12 observados: Boolean, Integer/Long, Currency, DateTime, Text y Memo según columna; confirmar mapeo final por campo.
- `Sí/No` se almacena frecuentemente como texto de longitud 2, no como Boolean; no convertirlo silenciosamente.
- `Null` y cadena vacía se distinguen en formularios, DTO y `Registrar`.
- Autonuméricos/IDs, jerarquías, Memo, URLs SharePoint, ficheros referenciados, cachés denormalizadas y tablas temporales necesitan tratamiento explícito.
- No se observan campos Access de tipo Attachment/OLE ni multivalor en el esquema autoritativo cosechado; los anexos son filas con `IDDocumento`/`NombreDocumento` y deben comprobarse contra el almacenamiento externo.

## Perfil agregado (privacidad segura)

| Tabla | Filas |
|---|---:|
| `TbExpedientes` | 453 |
| `TbExpedientesConEntidades` | 451 |
| `TbExpedientesAnexos` | 712 |
| `TbExpedientesSuministradores` | 713 |
| `TbExpedientesResponsables` | 730 |
| `TbExpedientesComerciales` | 333 |
| `TbExpedientesCPVs` | 429 |
| `TbExpedientesPECAL` | 366 |
| `TbExpedientesJuridicas` | 417 |
| `TbExpedientesLugaresEjecucion` | 194 |
| `TbExpedientesAnualidades` | 174 |
| `TbExpedientesHitos` | 46 |
| `TbExpedientesModificados` | 37 |
| `TbSuministradores` | 72 |

El resto de recuentos está registrado en la evidencia de sesión; las dos tablas `Copia de...` no devolvieron recuento mediante el wrapper y quedan como comprobación pendiente.

Comprobaciones: `TbExpedientes` tiene 453 IDs no nulos, 49 `CodExp` vacíos y 60 `Nemotecnico` vacíos; no se detectaron duplicados de `CodExp`, huérfanos de suministrador ni huérfanos de padre; no se detectaron intervalos contrato fin anteriores a inicio. Fechas serializadas se mantienen como rangos agregados, sin valores de filas.
