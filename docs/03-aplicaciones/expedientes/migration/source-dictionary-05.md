# Source Dictionary Tranche 05 — Tablas 41–49 (re-entries)

Las posiciones ordinales 41–49 en `data-model.md` (§Tabla de mapeo ordinal → nombre de tabla) apuntan a nombres de tabla **sin el prefijo `Tb`** que existen en `schema.sql` como alias o referencias a las tablas ya documentadas en las tranches 01–04. No son tablas independientes: son nombres alternativos que el ordinal mapping registra para completar la cobertura de posiciones 01–49.

**Fuente**: `docs/03-aplicaciones/expedientes/data-model.md` (ordinal mapping) + `schema.sql` para verificación.
**Esta tranche no introduce nueva información** — documenta los re-entries para cerrar el mapeo ordinal completo.

## Resumen de tablas

| # | Tabla legacy | Alias de ordinal | Detallada en | Estado |
|---|---|---|---|---|
| 41 | `Comerciales` | ordinal 02 | Tranche 01 | RE-ENTRY → ir a ordinal 02 |
| 42 | `CPVs` | ordinal 03 | Tranche 01 | RE-ENTRY → ir a ordinal 03 |
| 43 | `RACS` | ordinal 11 | Tranche 02 | RE-ENTRY → ir a ordinal 11 |
| 44 | `PECAL` | ordinal 10 | Tranche 01 | RE-ENTRY → ir a ordinal 10 |
| 45 | `LugaresEjecucion` | ordinal 07 | Tranche 01 | RE-ENTRY → ir a ordinal 07 |
| 46 | `Responsables` | ordinal 12 | Tranche 02 | RE-ENTRY → ir a ordinal 12 |
| 47 | `Suministradores` | ordinal 13 | Tranche 02 | RE-ENTRY → ir a ordinal 13 |
| 48 | `Jefaturas` | ordinal 05 | Tranche 01 | RE-ENTRY → ir a ordinal 05 |
| 49 | `Juridicas` | ordinal 06 | Tranche 01 | RE-ENTRY → ir a ordinal 06 |

## Detalle por tabla

### 41. Comerciales → RE-ENTRY (ordinal 02)

**Tabla real**: `TbComerciales` — catálogo de comerciales.
**Ordinal canónico**: 02.
**Detallada en**: [Source Dictionary Tranche 01, entrada 02](source-dictionary-01.md).
**No es una tabla independiente** — `schema.sql` contiene `CREATE TABLE [TbComerciales]` y la ordinal mapping registra `Comerciales` como nombre alternativo para el ordinal 41.

### 42. CPVs → RE-ENTRY (ordinal 03)

**Tabla real**: `TbCPV` — catálogo de códigos CPV.
**Ordinal canónico**: 03.
**Detallada en**: [Source Dictionary Tranche 01, entrada 03](source-dictionary-01.md).
**No es una tabla independiente** — `schema.sql` contiene `CREATE TABLE [TbCPV]` y la ordinal mapping registra `CPVs` como nombre alternativo.

### 43. RACS → RE-ENTRY (ordinal 11)

**Tabla real**: `TbRACS` — catálogo de responsables de actuación contractual.
**Ordinal canónico**: 11.
**Detallada en**: [Source Dictionary Tranche 02, entrada 11](source-dictionary-02.md).
**No es una tabla independiente** — `schema.sql` contiene `CREATE TABLE [TbRACS]` y la ordinal mapping registra `RACS` como nombre alternativo.

### 44. PECAL → RE-ENTRY (ordinal 10)

**Tabla real**: `TbPECAL` — catálogo de códigos PECAL.
**Ordinal canónico**: 10.
**Detallada en**: [Source Dictionary Tranche 01, entrada 10](source-dictionary-01.md).
**No es una tabla independiente** — `schema.sql` contiene `CREATE TABLE [TbPECAL]` y la ordinal mapping registra `PECAL` como nombre alternativo.

### 45. LugaresEjecucion → RE-ENTRY (ordinal 07)

**Tabla real**: `TbLugaresEjecucion` — catálogo de lugares de ejecución.
**Ordinal canónico**: 07.
**Detallada en**: [Source Dictionary Tranche 01, entrada 07](source-dictionary-01.md).
**No es una tabla independiente** — `schema.sql` contiene `CREATE TABLE [TbLugaresEjecucion]` y la ordinal mapping registra `LugaresEjecucion` como nombre alternativo.

### 46. Responsables → RE-ENTRY (ordinal 12)

**Tabla real**: `TbResponsablesPorRol` — catálogo de responsables por rol.
**Ordinal canónico**: 12.
**Detallada en**: [Source Dictionary Tranche 02, entrada 12](source-dictionary-02.md).
**No es una tabla independiente** — `schema.sql` contiene `CREATE TABLE [TbResponsablesPorRol]` y la ordinal mapping registra `Responsables` como nombre alternativo.

### 47. Suministradores → RE-ENTRY (ordinal 13)

**Tabla real**: `TbSuministradores` — catálogo de suministradores.
**Ordinal canónico**: 13.
**Detallada en**: [Source Dictionary Tranche 02, entrada 13](source-dictionary-02.md).
**No es una tabla independiente** — `schema.sql` contiene `CREATE TABLE [TbSuministradores]` y la ordinal mapping registra `Suministradores` como nombre alternativo.

### 48. Jefaturas → RE-ENTRY (ordinal 05)

**Tabla real**: `TbJefaturas` — catálogo de jefaturas.
**Ordinal canónico**: 05.
**Detallada en**: [Source Dictionary Tranche 01, entrada 05](source-dictionary-01.md).
**No es una tabla independiente** — `schema.sql` contiene `CREATE TABLE [TbJefaturas]` y la ordinal mapping registra `Jefaturas` como nombre alternativo.

### 49. Juridicas → RE-ENTRY (ordinal 06)

**Tabla real**: `TbJuridicas` — catálogo de entidades jurídicas.
**Ordinal canónico**: 06.
**Detallada en**: [Source Dictionary Tranche 01, entrada 06](source-dictionary-01.md).
**No es una tabla independiente** — `schema.sql` contiene `CREATE TABLE [TbJuridicas]` y la ordinal mapping registra `Juridicas` como nombre alternativo.

## Explicación: por qué hay re-entries

El ordinal mapping de `data-model.md` fue cosechado mediante `dysflow list_tables` sobre `Expedientes_datos.accdb`. Algunas tablas Access aparecen en el resultado con y sin el prefijo `Tb` (por ejemplo, `TbComerciales` y `Comerciales`). El ordinal mapping las registra como entradas separadas para completar el rango 01–49, pero ambas apuntan al mismo objeto en `schema.sql`.

El nombre con prefijo `Tb` es la forma canónica y es la que se usa en las tranches 01–04. Las re-entries simplemente cierran el mapeo ordinal para que cualquier referencia a posición 41–49 tenga una entrada en esta documentación.

## Navegación

Tranche anterior: [Source Dictionary Tranche 04 — Tablas 31–40](source-dictionary-04.md).
Tranche siguiente: — (última tranche).
[← Back to Expedientes migration](README.md).
