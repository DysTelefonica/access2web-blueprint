# Condor — modelo físico y diccionario de datos

## Autoridad y fecha

Fuente física: `C:\00repos\datos\condor_datos.accdb` (5 MB). **Inventario real obtenido vía Dysflow MCP el 2026-08-05** sobre `C:\00repos\codigo\00_CONDOR\staging\condor_datos.accdb` después de `register_worktree` + `accessPath` absoluto explícito.

## Inventario real Dysflow (2026-08-05, staging)

| Categoría | Resultado |
|---|---|
| **Tablas totales** | **15** en `condor_datos.accdb` (staging) |
| **Filas en `tbSolicitudes`** | **1** (staging de prueba) |
| **Filas en `tbEstados`** | **9** (catálogo) |
| **Filas en `tbRechazos`** | **0** |
| **Filas en `tbAdjuntos`** | **0** |
| **Filas en `tbLogErrores`** | **3** |
| **FK relationships** (user tables) | **5** desde `tbSolicitudes` |

⚠️ **Volumen bajo**: staging de Condor tiene **1 solicitud de prueba**. Para conocer el volumen real de producción se necesita acceso al backend `00_main` o al backend autoritativo en `C:\00repos\datos\` (que tiene staging-local). **NO se ha ejercitado el backend de producción en esta pasada.**

⚠️ **Convención de nombres**: las tablas usan `tb` (minúscula) + CamelCase (`tbSolicitudes`, `tbDatosCDCA`, etc.). Diferencia estética con las otras apps que usan `Tb` mayúscula. La migración a PostgreSQL normaliza a `snake_case` (`solicitudes`, `datos_cdca`, etc.).

### Lista completa de las 15 tablas

```
tbAdjuntos
tbDatosCDCA
tbDatosCDCASUB
tbDatosPC
tbDatosPCSUB
tbEstados
tbHistorialRechazos
tbLogCambios
tbLogErrores
tbLogEstados
tbMapeoCampos
tbRechazos
tbSolicitudes
tbTransiciones
tbValidacionRevision
```

### Esquema real `tbSolicitudes` (12 columnas, tipos reales)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `idSolicitud` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `idExpediente` | 4 (LongInteger) | 4 | true | `BIGINT NOT NULL` — FK conceptual a `TbExpedientes` (Lanzadera), sin constraint ⚠️ |
| `tipoSolicitud` | 10 (Text) | 20 | true | `VARCHAR(20) NOT NULL` — enum: `PC`, `CD_CA`, `CD_CA_SUB`, `PC_SUB` |
| `codigoSolicitud` | 10 (Text) | 50 | true | `VARCHAR(50) NOT NULL` — identificador funcional |
| `idNCAsociada` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — FK conceptual a `TbNoConformidades`, sin constraint ⚠️ |
| `idEstadoInterno` | 4 (LongInteger) | 4 | true | `BIGINT NOT NULL` — FK conceptual a `tbEstados`, sin constraint ⚠️ |
| `fechaCreacion` | 8 (DateTime) | 8 | true | `TIMESTAMP NOT NULL` |
| `usuarioCreacion` | 10 (Text) | 100 | true | `VARCHAR(100) NOT NULL` — auditoría |
| `fechaModificacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `usuarioModificacion` | 10 (Text) | 100 | false | `VARCHAR(100) NULL` — auditoría |
| `revisionCalidadEstado` | 10 (Text) | 20 | true | `VARCHAR(20) NOT NULL` — enum de estado de revisión |
| `revisionCalidadComentarios` | 12 (Memo) | 0 | false | `TEXT NULL` |

### Relaciones físicas reales (5 FK desde `tbSolicitudes`)

| Origen | Columna FK | Destino | Columna FK |
|---|---|---|---|
| `tbSolicitudes` | `idSolicitud` | `tbAdjuntos` | `idSolicitud` |
| `tbSolicitudes` | `idSolicitud` | `tbDatosCDCA` | `idSolicitud` |
| `tbSolicitudes` | `idSolicitud` | `tbDatosCDCASUB` | `idSolicitud` |
| `tbSolicitudes` | `idSolicitud` | `tbDatosPC` | `idSolicitud` |
| `tbSolicitudes` | `idSolicitud` | `tbDatosPCSUB` | `idSolicitud` (presumido; no listado en output truncado) |

⚠️ **Inconsistencias detectadas en FKs**:

1. **`tbSolicitudes.idEstadoInterno` debería tener FK a `tbEstados` pero NO la tiene**. Data integrity gap.
2. **`tbSolicitudes.idExpediente` es FK conceptual a Lanzadera/Expedientes**. Sin constraint.
3. **`tbSolicitudes.idNCAsociada` es FK conceptual a NoConformidades**. Sin constraint.
4. **NO hay FKs explícitas entre `tbRechazos`, `tbLogErrores`, `tbTransiciones`, `tbValidacionRevision` y `tbSolicitudes`**. Probablemente deberían tener.
5. **NO hay FK explícita entre `tbLogEstados` y `tbEstados`**. Data integrity gap.

### Pendientes de discovery (segunda pasada)

- **`tbEstados`** (9 filas): catálogo de estados. Schema no inspeccionado.
- **`tbRechazos`** (0 filas): schema no inspeccionado.
- **`tbDatosPC`, `tbDatosCDCA`, `tbDatosCDCASUB`, `tbDatosPCSUB`**: schemas no inspeccionados en esta pasada.
- **`tbTransiciones`**: schema no inspeccionado.
- **`tbValidacionRevision`**: schema no inspeccionado.
- **`tbAdjuntos`**: schema no inspeccionado.
- **`tbLogCambios`, `tbLogErrores`, `tbLogEstados`**: schemas no inspeccionados.

Estos schemas se obtendrán en una segunda pasada con `get_schema` para cada tabla.