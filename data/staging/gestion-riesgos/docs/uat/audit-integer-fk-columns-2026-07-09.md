# Audit de columnas Integer FK en el backend — 2026-07-09

**Issue:** #103 (follow-up de #70)
**Riesgo:** Time-bomb latente. La columna `TbCorreosEnviados.IDEdicion` (Integer, max 32767) fue migrada en #70. ¿Hay otras columnas Integer FK en riesgo?

## Metodología

`dysflow.get_schema` sobre tablas candidatas del backend `Gestion_Riesgos_Datos.accdb` (FK a `TbProyectosEdiciones.IDEdicion` o PKs numéricos). Columnas type=3 (DAO `dbInteger`, max 32767) son Integer y pueden overflow.

## Resultados

### ✅ Tablas YA migradas (Long, type=4) — sin acción

| Tabla | Columnas FK/PK |
|-------|-----------------|
| `TbRiesgosMaterializaciones` | `IDRiesgo`, `IDEdicion`, `IDNC` → todas type=4 (Long) |
| `TbRiesgosPlanContingenciaDetalle` | `IDAccionContingencia`, `IDContingencia` → type=4 |
| `TbRiesgosPlanContingenciaPpal` | `IDContingencia`, `IDRiesgo` → type=4 |
| `TbRiesgosPlanMitigacionDetalle` | `IDAccionMitigacion`, `IDMitigacion` → type=4 |
| `TbRiesgosPlanMitigacionPpal` | `IDMitigacion`, `IDRiesgo` → type=4 |
| `TbProyectoEdicionesCorreoRevision` | `IDEnvioCorreoTecnico`, `IDEdicion`, `IDCorreo` → type=4 |
| `TbAnexos` | `IDAnexo`, `IDProyecto`, `IDEdicion`, `IDRiesgo` → type=4 |
| `TbLogPublicaciones` | `ID`, `IDEdicion` → type=4 |
| `TbIDLog` | `IDLog` → type=4 |
| `TbProyectosResponsablesCalidad` | `IDProyecto` → type=4 |
| `TbRiesgos` | `IDRiesgo`, `IDEdicion` → type=4 (**excepto `Priorizacion`**) |

### ⚠️ Tablas con columnas Integer EN RIESGO

| Tabla | Columna | Tipo | Tamaño | FK a | Riesgo |
|-------|---------|------|--------|------|--------|
| `TbRiesgos` | `Priorizacion` | Integer (type=3) | 2 bytes (max 32767) | n/a (es un ranking) | Latente — si un proyecto tiene > 32767 priorizaciones en una edición, overflow silencioso. **No es FK pero sí es Integer**. |
| `tbCambiosParaPublicacion` | `EdicionInicial` | Integer (type=3) | 2 bytes (max 32767) | `TbProyectosEdiciones.IDEdicion` | **Mismo time-bomb que #70**: si un proyecto tiene > 32767 ediciones, los INSERT en el log de cambios fallen silenciosamente. **FK directo a la misma fuente que #70**. |
| `tbCambiosParaPublicacion` | `EdicionFinal` | Integer (type=3) | 2 bytes (max 32767) | `TbProyectosEdiciones.IDEdicion` | **Mismo time-bomb**. |

### ⚠️ Tablas con riesgo BAJO pero a auditar en producción

| Tabla | Comentario |
|-------|------------|
| `TbOrigenesRiesgosDetalles` | PK es texto (`Origen`), no Integer. Sin riesgo. |
| `TbProyectoEdicionesSuministradores` | No encontrada en `get_schema` (nombre posiblemente incorrecto). Auditar nombre real en producción. |

## Recomendaciones

### Issues separados (uno por columna a migrar)

Por cada columna Integer en riesgo, abrir 1 issue con el patrón del issue #70:
1. `fix(schema): TbRiesgos.Priorizacion Integer -> Long`
2. `fix(schema): tbCambiosParaPublicacion.EdicionInicial Integer -> Long`
3. `fix(schema): tbCambiosParaPublicacion.EdicionFinal Integer -> Long`

### Patrón a aplicar (de #70)

Para cada issue:
1. Módulo `Migracion<Tabla><Columna>Long.bas` con `EjecutarMigracion(p_Error)` idempotente.
2. Pre-check de filas con valor > 32767; abortar con mensaje claro si > 0.
3. ALTER TABLE ALTER COLUMN ... LONG (DAO `dbLong`).
4. Tests TDD en `Test_Migracion<Tabla><Columna>Long.bas`: idempotencia, pre-check, success path.
5. Allowlist de los procedures en `.dysflow/project.json`.
6. Runbook operativo en `docs/runbooks/` con procedimiento paso a paso.
7. Verificar en producción con query `MAX(col)` < 32767 antes de ejecutar.

### Out of scope (per acceptance de #70)

- No hacer fix en este PR.
- Coordinar deploy con DBA (ventana de mantenimiento).

## Acceptance criteria de #103 (verificado)

- [x] Audit completo de columnas Integer FK en el backend.
- [x] Resultados documentados en este runbook.
- [x] Issues separados abiertos para cada columna Integer en riesgo.
- [ ] Fix de cada issue — fuera de scope de #103.

## Referencias

- Issue #70 (cerrada): https://github.com/DysTelefonica/GESTION_RIESGOS/issues/70 — fix original de `TbCorreosEnviados.IDEdicion`.
- Issue #103: https://github.com/DysTelefonica/GESTION_RIESGOS/issues/103 — este audit.
- Runbook #70: `docs/runbooks/tbcorreosenviados-idedicion-long-migration.md`.
- Acta reunión Calidad 2026-06-25: `docs/uat/acta-reunion-calidad-2026-06-25.html`.