# Runbook: tbCambiosParaPublicacion.EdicionInicial + EdicionFinal Integer → Long

**Issue:** [#107](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/107) — `fix(schema): tbCambiosParaPublicacion.EdicionInicial + EdicionFinal Integer -> Long`

**Autor:** consumer `gestion_riesgos` (sesion 2026-07-14)

**Patrón:** sigue el contrato de `#70` (TbCorreosEnviados.IDEdicion) pero con **dos columnas** en una sola migración.

**Status schema (staging 2026-07-14):**
- `tbCambiosParaPublicacion.EdicionInicial`: `dbInteger` (type=3, size=2, max 32,767), `required=false` — FK a `TbProyectosEdiciones.IDEdicion` (Long)
- `tbCambiosParaPublicacion.EdicionFinal`: `dbInteger` (type=3, size=2, max 32,767), `required=true` — FK a `TbProyectosEdiciones.IDEdicion` (Long)
- Filas: 68
- Valores: `EdicionInicial ∈ [2, 5]`, `EdicionFinal ∈ [1, 6]`, **ninguna fuera de rango Integer**

---

## Por qué esta migración

`TbProyectosEdiciones.IDEdicion` es `dbLong` (max 2^31-1). Cualquier edicion nueva con ID > 32,767 en un proyecto rompera el INSERT en `tbCambiosParaPublicacion` con error DAO 3035 (data type conversion overflow). Ventana actual muy amplia (max=6 en staging), pero el time-bomb existe por patron estructural.

El mismo patron estructural se aplicó ya en #70 (TbCorreosEnviados) y #103 (TbRiesgos.Priorizacion). Auditorias #103 y su follow-up #107 listan las columnas restantes del proyecto que aún son `Integer` en FKs.

---

## Alcance del cambio

| Recurso | Cambia | Como |
|---|---|---|
| `src/modules/MigracionTbCambiosParaPublicacionEdicionLong.bas` | sí | nuevo módulo idempotente |
| `src/modules/Test_TbCambiosParaPublicacion_Edicion.bas` | sí | 3 átomos TDD + 1 RunAll |
| `.dysflow/project.json` | sí | 4 nuevas procedures en allowlist |
| `tests/tests.vba.cambios-publicacion-edicionlong.json` | sí | test manifest |
| `tbCambiosParaPublicacion.EdicionInicial` (type) | sí | Integer → Long (via helper) |
| `tbCambiosParaPublicacion.EdicionFinal` (type) | sí | Integer → Long (via helper) |
| Datos | no | ALTER preserva todas las filas |
| Codigo de producto (modulos/classes/forms) | no | sin lecturas criticas con riesgo |
| Backend prod (cambios fisicos en columnas) | NO en este PR | operacional via runbook |

---

## Pre-checklist (operador DBA, ANTES del ALTER)

1. **Backups**:
   ```powershell
   # Copia local del backend
   $today = Get-Date -Format "yyyyMMdd-HHmmss"
   Copy-Item "Gestion_Riesgos_Datos.accdb" "Gestion_Riesgos_Datos.bak-$today-pre-edicionlong-cambios-publicacion"
   # Confirmar SHA
   Get-FileHash "Gestion_Riesgos_Datos.accdb" -Algorithm SHA256
   ```
2. **Audit de datos (replica el query del módulo para validar pre-condition):**
   ```sql
   SELECT COUNT(*) AS Total,
          MAX(EdicionInicial) AS MaxIni, MIN(EdicionInicial) AS MinIni,
          MAX(EdicionFinal) AS MaxFin, MIN(EdicionFinal) AS MinFin,
          SUM(IIF(EdicionInicial > 32767 OR EdicionFinal > 32767, 1, 0)) AS FueraDeRango
   FROM tbCambiosParaPublicacion;
   ```
   **Regla**: `FueraDeRango` debe ser `0`. Si es >0, NO correr el ALTER — respaldar las filas problemáticas y remediar manualmente.
3. **Lock del backend**: detener todos los clientes Access conectados o avisar coordinacion para ventana de mantenimiento.
4. **Verificar que el modulo TDD esté corriendo** en el entorno target. En staging: usuario compila + corre `tests.vba.cambios-publicacion-edicionlong.json`. Si los 3 átomos pasan, modulo funciona. Si fallan, NO continuar con operacional.

---

## Ejecución (dev / staging)

Una vez que el código está mergeado a staging y el operador ya compiló + tests verdes:

### Desde Access / VBE (interactivo)
```vba
? MigracionTbCambiosParaPublicacionEdicionLong.EjecutarMigracion("")
' Retorna EnumSiNo.Si (=0) en exito
' Retorna EnumSiNo.No (=1) si hay pre-check fail
' Propaga Err.Raise 1000 si la migracion falla
```

### Idempotencia
```vba
' 2da llamada en la misma sesion: no-op (ya es LONG,LONG)
? MigracionTbCambiosParaPublicacionEdicionLong.EjecutarMigracion("")
' Tambien EnumSiNo.Si
```

### Verificacion post-ALTER
```sql
-- Debe devolver 4 en ambas (dbLong = 4)
SELECT Type FROM MSysObjects
WHERE Name IN ('tbCambiosParaPublicacion');
-- Luego abrir la tabla en DAO y:
--   tdf!EdicionInicial.Type = 4  (dbLong)
--   tdf!EdicionFinal.Type = 4    (dbLong)
```

O via modulo dysflow (read-only):
```js
await tools.dysflow.get_schema({tableName: "tbCambiosParaPublicacion", backendPath: "Gestion_Riesgos_Datos.accdb"})
// Devuelve type=4 en ambos EdicionInicial y EdicionFinal
```

---

## Post-checklist

1. **Tests atómicos en verde** (`tests.vba.cambios-publicacion-edicionlong.json`):
   - `PersisteValorSobreIntegerMax`: inserto 50.000 y 50.001 (sobre Integer max), ambas deben round-trip OK
   - `MigrationEsIdempotente`: dos llamadas, ambas Si, ambas columnas terminan `dbLong`
   - `RegistrosExistentesPreservados`: conteo de filas antes == despues (ALTER nunca pierde filas)
2. **Regresion bloque 0** (sanidad): `tests.vba.smoke.json` (sin cambios en comportamiento de otras tablas).
3. **Verificacion manual**: abrir `tbCambiosParaPublicacion` en Access, confirmar que `EdicionInicial` y `EdicionFinal` aceptan valores > 32.767 sin overflow.
4. **Commit PR de cierre** con SHA del header + test reference. Mensaje:
   ```
   fix(schema): migra tbCambiosParaPublicacion.EdicionInicial + EdicionFinal Integer -> Long (issue #107)

   - ALTER via helper idempotente `MigracionTbCambiosParaPublicacionEdicionLong`
   - 3 tests atomicos + 1 RunAll (Test_TbCambiosParaPublicacion_Edicion)
   - Runbook: docs/uat/runbook-tb-cambios-publicacion-edicionlong-migration.md
   - Audit: 68 filas, max EdicionInicial=5, max EdicionFinal=6, 0 fuera de rango

   Refs: #107, audit #103
   ```

---

## Rollback

Si el ALTER daña algo (improbable — DAO `ALTER COLUMN LONG` sobre Integer no toca filas, solo metadata):

1. **Restaurar backend desde backup**:
   ```powershell
   Copy-Item "Gestion_Riesgos_Datos.bak-$today-pre-edicionlong-cambios-publicacion" "Gestion_Riesgos_Datos.accdb" -Force
   ```
2. **Verificar que no hay connectores activos**: list_access_operations antes y despues.
3. **Re-importar el codigo con `git revert <commit-sha>`** en gestion_riesgos.

Si solo el modulo VBA da problemas (no el schema) — el usuario puede borrar `MigracionTbCambiosParaPublicacionEdicionLong.bas` y `Test_TbCambiosParaPublicacion_Edicion.bas` desde Access sin tocar la tabla (las columnas quedan como Long, que es lo deseable).

---

## Coordinacion operacional

- **Ventana**: pedir a Natalia (Calidad) coordinacion con usuarios finales (~5 min de indisponibilidad del backend mientras corre el ALTER + verificacion).
- **Quien ejecuta**: el DBA o el responsable del backend. **NO el IA** (el consumer IA no tiene acceso a prod).
- **Cuando**: despues de Fase A (per planning staging-acceptance-contract #115), este es trabajo de mantencion no urgente (el time-bomb no presiona hasta que un proyecto tenga > 32.767 ediciones, lo cual es improbable a corto plazo).

---

## Referencias

- **Issue**: [#107](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/107)
- **Auditoria previa**: #103 (auditoría completa, suggestion #107 + #106)
- **Patrón**: #70 (TbCorreosEnviados) — `MigracionCorreosEnviadosIDEdicionLong.bas` + `Test_TbCorreosEnviados_IDEdicion.bas`
- **Sister trabajo**: #106 (TbRiesgos.Priorizacion)
- **Migracion helper**: `src/modules/MigracionTbCambiosParaPublicacionEdicionLong.bas` (278 lineas)
- **TDD atoms**: `src/modules/Test_TbCambiosParaPublicacion_Edicion.bas` (3 tests + RunAll)
- **Test manifest**: `tests/tests.vba.cambios-publicacion-edicionlong.json`
