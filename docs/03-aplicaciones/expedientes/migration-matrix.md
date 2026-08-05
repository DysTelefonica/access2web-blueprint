# Expedientes — matriz de migración (scaffold de preservación)

Esta matriz no decide el esquema PostgreSQL. Su regla es conservadora: todo campo/registro se conserva hasta que negocio apruebe una disposición.

| Fuente | Significado/candidato de dominio | Transformación | Estado | Validación/rechazo | Decisión abierta |
|---|---|---|---|---|---|
| `TbExpedientes.*` | expediente y ciclo de vida | separar identidad, clasificación, fechas y estados; conservar valores originales | mapeado preliminar | recuento 1:1, IDs y códigos reconciliados | catálogo de estados final |
| `IDExpedientePadre` | jerarquía AM/lote/basado | FK autorreferente preservando ID legacy | mapeado preliminar | cero huérfanos; detectar ciclos | política de ciclos |
| `CodExp`, `CodExpLargo`, `Nemotecnico` | identificadores funcionales | conservar columnas originales y normalizadas | necesita decisión | unicidad y vacíos; no inventar sustituto | claves canónicas |
| flags `EsAM/EsLote/EsBasado/EsExpediente` y `Tipo` | clasificación/tipo | tabla de valores + valor legacy | necesita decisión | coherencia tipo/jerarquía | fuente de verdad entre flags y Tipo |
| `TbExpedientes*` joins | relaciones N:N y jerarquías | entidades de relación con IDs legacy | mapeado preliminar | cardinalidad y FK; preservar duplicados | restricciones nuevas |
| `TbExpedientesConEntidades` | read-model denormalizado | conservar como snapshot/reconstruible, nunca descartar | necesita decisión | comparar cadenas contra joins | retención histórica |
| `TbExpedientesAnexos` | documentación | metadatos + referencia de almacenamiento | external reference | comprobar existencia de fichero/URL sin copiar datos | repositorio destino |
| `TbCambios`, `TbUltimoCambio` | auditoría | evento de cambio y último estado | mapeado preliminar | conteos y timestamps | retención/legal |
| tablas E2E/temp | soporte técnico de exportación | migrar solo si se requiere trazabilidad histórica; conservar dump si no | technical-only candidate | documentar antes de excluir | decisión explícita, no obsolescencia |
| tablas copia/históricas/auxiliares | histórico o soporte | conservar en zona legacy hasta clasificación | needs business decision | comparar recuentos y uso en código | archivo vs dominio |
| integraciones AGEDYS/HPS/riesgos/NC/correos | referencias externas | claves externas y snapshot contractual | external reference | reconciliar por identificador, sin asumir propiedad | ownership y sincronización |

## Ledger

- **Mapeado preliminar:** cabecera, relaciones, catálogos, auditoría.
- **Necesita decisión:** flags frente a tipo, read-model, históricos, E2E, anexos y contratos externos.
- **Candidato técnico:** temporales y helpers de exportación; no retirado.
- **Referencia externa:** documentos, identidad, AGEDYS, HPS, riesgos, NC y correo.
- **No hay campos/filas declarados obsoletos.**

Cada fila futura debe expandirse a `source table.field` para los 49 esquemas cosechados, con transformación, regla de reconciliación, rechazo y disposición aprobada. La fuente completa de campos es el ERD enlazado en `data-model.md`.
