# Reconciliation Runbook — Expedientes

Runbook de reconciliación para la migración de Expedientes. Define los conteos, muestras, hash y rollback para verificar la integridad de la migración.

**Design**: `openspec/changes/expedientes-web-migration/design.md` (D-EXP-9)
**Spec**: `uat-cutover-legacy-retirement.md` (EXP-CAP-057..063)

## Conteos de reconciliación

| Métrica | Legacy | Nuevo | Criterio |
|---|---|---|---|
| Total expedientes | `SELECT COUNT(*) FROM TbExpedientes` | `SELECT COUNT(*) FROM expedientes` | Iguales |
| Total anexos | `SELECT COUNT(*) FROM TbExpedientesAnexos` | `SELECT COUNT(*) FROM expediente_anexos` | Iguales |
| Total anualidades | `SELECT COUNT(*) FROM TbExpedientesAnualidades` | `SELECT COUNT(*) FROM expediente_anualidades` | Iguales |
| Total hitos | `SELECT COUNT(*) FROM TbExpedientesHitos` | `SELECT COUNT(*) FROM expediente_hitos` | Iguales |
| Total modificados | `SELECT COUNT(*) FROM TbExpedientesModificados` | `SELECT COUNT(*) FROM expediente_modificados` | Iguales |
| Total E2E batches | `SELECT COUNT(*) FROM TbE2EExportBatch` | `SELECT COUNT(*) FROM e2e_batches` | Iguales |

**Criterio de reconciliación**: Los conteos deben ser iguales entre legacy y nuevo. Si hay diferencia, investigar antes de continuar.

## Muestras de reconciliación

| Muestra | Tamaño | Criterio |
|---|---|---|
| Expedientes aleatorios | 100 | Todos los campos coinciden |
| Anexos aleatorios | 50 | Todos los campos coinciden |
| Anualidades aleatorias | 50 | Todos los campos coinciden |
| Hitos aleatorios | 50 | Todos los campos coinciden |
| E2E batches aleatorios | 20 | Todos los campos coinciden |

**Criterio de reconciliación**: Las muestras deben coincidir campo-a-campo entre legacy y nuevo. Si hay diferencia, investigar antes de continuar.

## Hash de reconciliación

| Hash | Algoritmo | Aplica a | Criterio |
|---|---|---|---|
| Hash de expediente | FNV-1a | `expedientes` | Hash igual entre legacy y nuevo |
| Hash de E2E batch | FNV-1a | `e2e_batches` | Hash igual entre legacy y nuevo |

**Criterio de reconciliación**: Los hash deben ser iguales entre legacy y nuevo. Si hay diferencia, investigar antes de continuar.

**Golden de continuidad**: El hash FNV-1a requiere golden de continuidad antes de aprobarse (D-EXP-7). El golden se genera a partir de un expediente de referencia y se verifica en cada migración.

## Rollback

**Escenario**: La migración falla o se detecta un problema crítico.

**Procedimiento**:
1. **Detener el tráfico al nuevo**: Cambiar el routing para dirigir el tráfico al legacy.
2. **Verificar el legacy**: Confirmar que el legacy está operativo y los datos están intactos.
3. **Rollback del schema**: Ejecutar `DROP SCHEMA expedientes CASCADE` para eliminar el schema nuevo.
4. **Verificar el rollback**: Confirmar que el legacy está operativo y el schema nuevo fue eliminado.
5. **Documentar el rollback**: Registrar el rollback en el log de migración.

**Criterio de rollback**: El rollback devuelve el tráfico al legacy sin reverse-write automático (D-EXP-9). El legacy permanece intacto durante todo el proceso.

**Reversibilidad**: La migración es reversible antes del cutover. Después del cutover, el rollback requiere restaurar desde backup.

## Verificación

- [ ] Los conteos de reconciliación son iguales entre legacy y nuevo
- [ ] Las muestras de reconciliación coinciden campo-a-campo
- [ ] Los hash de reconciliación son iguales entre legacy y nuevo
- [ ] El golden de continuidad fue generado y verificado
- [ ] El procedimiento de rollback fue documentado y probado
