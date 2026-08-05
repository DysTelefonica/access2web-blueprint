# HPS — modelo físico y diccionario de datos

## Autoridad y fecha

Fuente física: `C:\00repos\datos\HPST.accdb` (7 MB). **Inventario real obtenido vía Dysflow MCP el 2026-08-05** después de `register_worktree` + `accessPath` absoluto explícito. Las observaciones sobre `.dysflow/project.json` legacy se aplican (ver [NoConformidades § D89 invalidado](../no-conformidades/security-rules.md#d89--diagnóstico-del-fallo-de-list_objects-de-dysflow--invalidado) para el patrón general).

## Inventario real Dysflow (2026-08-05)

| Categoría | Resultado |
|---|---|
| **Tablas totales** | **22** en `HPST.accdb` |
| **Filas en `TbUsuarios`** | **345** (usuarios activos) |
| **Filas en `TbUsuariosHistoricos`** | **242** (usuarios históricos) |
| **Filas en `TbHPS`** | **1280** (relaciones usuario-curso) |
| **Filas en `TbObservaciones`** | **330** (observaciones) |
| **FK relationships** (user tables) | **6** |
| **Columnas en `TbUsuarios`** | **27** |

### Lista completa de las 22 tablas

```
Copia de TbExpedienteLugares (legacy copy)
Copia de TbExpedientes (legacy copy)
Copia de TbUsuarios (legacy copy)
Copia de TbUsuariosEntidades (legacy copy)
Errores de pegado (sentinel de errores en operaciones de pegado masivo)
TbAnexosUsuariosHistoricos
TbAnexosUsuariosHPS
TbAnexosUsuariosSICA
TbAuxCursos
TbConsultas
TbHPS
TbHPSEquivalencia
TbHPSGrado
TbJuridicasContratacion
TbMotivoHPS
TbObservaciones
TbObservacionesHistoricas
TbUsuarioAnexos
TbUsuarios
TbUsuariosEntidades
TbUsuariosHistoricos
TbUsuariosSICA
```

⚠️ **Duplicación detectada**: `TbHPS` y `TbUsuariosHistoricos` aparecen tanto en el frontend como en el backend (Dysflow emite `ACCESS_TABLE_AMBIGUOUS` al consultarlas sin `target`). Esto es un patrón legacy de sincronización o cache local. Disposición: ver D92 en [migration-matrix.md](migration-matrix.md#d92--disposiciones-específicas-de-hps).

⚠️ **4 tablas "Copia de..."**: patrón legacy de copia antes de cambios masivos. Probablemente obsoletas. Requieren disposición explícita (mantener en legacy / archivar / descartar). Ver D92.

⚠️ **Sentinel `Errores de pegado`**: tabla para capturar errores en operaciones de pegado masivo. Migración debe decidir si preservar (con política de retención) o transformar en un log estructurado canónico (D27).

### Esquema real `TbUsuarios` (27 columnas, tipos reales)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `ID` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `DNI` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **dato personal (D92)** |
| `Nombre` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **dato personal** |
| `Apellido_1` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **dato personal** |
| `Apellido_2` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **dato personal** |
| `Telefono` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **dato personal** |
| `Correo_e` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **dato personal** |
| `IDExpediente` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — sin FK física ⚠️ |
| `F_Nacimiento` | 8 (DateTime) | 8 | false | `DATE NULL` — **dato personal** |
| `LugarNacimiento` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **dato personal** |
| `Motivo_HPS` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `F_Curso` | 8 (DateTime) | 8 | false | `DATE NULL` |
| `CursoEnVigor` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ boolean como texto ('Sí'/'No') |
| `Requiere_Curso` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ boolean como texto |
| `FechaPrimeraConvocatoria` | 8 (DateTime) | 8 | false | `DATE NULL` |
| `FechaSegundaConvocatoria` | 8 (DateTime) | 8 | false | `DATE NULL` |
| `FechaCorreoNoCurso` | 8 (DateTime) | 8 | false | `DATE NULL` |
| `F_Baja` | 8 (DateTime) | 8 | false | `DATE NULL` |
| `LugarPrestacionServicio` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `IDEmpresaUsuario` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — sin FK física ⚠️ |
| `IDEmpresaHPS` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — sin FK física ⚠️ |
| `IDJuridicaContrato` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — sin FK física ⚠️ |
| `FAvisoConcesion` | 8 (DateTime) | 8 | false | `DATE NULL` |
| `RequiereComunicacionConcesion` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ boolean como texto |
| `FechaHPSConcesionMinima` | 8 (DateTime) | 8 | false | `DATE NULL` |
| `IDSolicitud` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — sin FK física ⚠️ |
| `CadenaContratistas` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |

### Relaciones físicas reales (6 FK entre user tables)

| Origen | Columna FK | Destino | Columna FK |
|---|---|---|---|
| `TbUsuariosHistoricos` | `ID` | `TbObservacionesHistoricas` | `ID` ⚠️ |
| `TbUsuariosSICA` | `ID` | `TbAnexosUsuariosSICA` | `IDUsuarioSICA` |
| `TbUsuarios` | `ID` | `TbAnexosUsuariosHPS` | `IDUsuario` |
| `TbUsuarios` | `ID` | `TbHPS` | `IDUsuario` |
| `TbUsuarios` | `ID` | `TbObservaciones` | `ID` ⚠️ |
| `TbUsuarios` | `ID` | `TbUsuarioAnexos` | `IDUsuario` |

⚠️ **Inconsistencias detectadas en FKs**:

1. `TbUsuariosHistoricos.ID → TbObservacionesHistoricas.ID` (PK-to-PK genérico, debería ser por `IDUsuarioHistorico`).
2. `TbUsuarios.ID → TbObservaciones.ID` (PK-to-PK genérico, debería ser por `IDUsuario`).
3. **NO hay FK física** entre `TbUsuarios` y `TbExpedientes` pese a tener `IDExpediente`. La relación es conceptual/por código.
4. **NO hay FK física** entre `TbUsuarios` y `TbEmpresas` pese a `IDEmpresaUsuario` e `IDEmpresaHPS`. La relación es por ID.
5. **NO hay FK física** entre `TbUsuarios` y `TbJuridicasContratacion` pese a `IDJuridicaContrato`. La relación es por ID.
6. **NO hay FK física** entre `TbUsuarios` y `TbSolicitudes` pese a `IDSolicitud`. La relación es por ID.

Migración debe decidir si formalizar las FK conceptuales (requiere migrar datos consistentes) o mantener como referencia conceptual (D86 / D87 sobre la forma hexagonal).