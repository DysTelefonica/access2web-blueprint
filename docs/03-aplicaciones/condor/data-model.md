# Condor — modelo físico y diccionario de datos

## Autoridad y fecha

Fuente física: **`C:\00repos\datos\condor_datos.accdb`** (5 MB) — backend autoritativo. **Inventario real obtenido vía Dysflow MCP el 2026-08-05** con `accessPath` + `backendPath` absolutos explícitos. Confirmado que el `staging/condor_datos.accdb` local tiene la misma estructura y volumen que el autoritativo (staging lee del autoritativo vía `TbConfiguracionBackends`).

## Inventario real Dysflow (2026-08-05, backend autoritativo)

| Categoría | Resultado |
|---|---|
| **Tablas totales** | **15** en `condor_datos.accdb` |
| **Filas en `tbSolicitudes`** | **1** (backend autoritativo casi vacío; staging tiene el mismo volumen) |
| **Filas en `tbEstados`** | **9** (catálogo) |
| **Filas en `tbRechazos`** | **0** |
| **Filas en `tbAdjuntos`** | **0** |
| **Filas en `tbLogCambios`** | **5** |
| **Filas en `tbLogErrores`** | **3** |
| **Filas en `tbLogEstados`** | **0** |
| **Filas en `tbMapeoCampos`** | **183** ⚠️ config de mapeo legacy → moderno |
| **Filas en `tbDatosCDCA`** | **0** |
| **Filas en `tbDatosCDCASUB`** | **0** |
| **Filas en `tbDatosPC`** | **0** |
| **Filas en `tbDatosPCSUB`** | **0** |
| **Filas en `tbHistorialRechazos`** | **0** |
| **Filas en `tbValidacionRevision`** | (no contado en esta pasada) |
| **FK relationships** (user tables) | **5** desde `tbSolicitudes` |

⚠️ **Hallazgo del volumen**: el backend autoritativo también está casi vacío (1 solicitud, 9 estados, 5 logs). **NO es un sistema en producción con miles de filas**. Posiblemente Condor está en despliegue temprano o se está migrando de otro backend.

⚠️ **Hallazgo de `tbMapeoCampos`**: **183 filas** de config de mapeo entre columnas legacy y modernas. Es **config que se preserva como datos**, no código. La nueva plataforma debe migrar estas 183 filas como `INSERT INTO tb_mapeo_campos VALUES (...)` o equivalente.

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

## Schemas reales (Dysflow, segunda pasada)

### `tbSolicitudes` (12 columnas, 1 fila)

Cabecera de Solicitud. FKs a `tbEstados` (conceptual), Lanzadera `TbExpedientes` (cross-app), NoConformidades `TbNoConformidades` (cross-app).

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `idSolicitud` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `idExpediente` | 4 (LongInteger) | 4 | true | `BIGINT NOT NULL` — FK conceptual a `TbExpedientes` (Lanzadera) |
| `tipoSolicitud` | 10 (Text) | 20 | true | `VARCHAR(20) NOT NULL` — enum: `PC`, `CD_CA`, `CD_CA_SUB`, `PC_SUB` |
| `codigoSolicitud` | 10 (Text) | 50 | true | `VARCHAR(50) NOT NULL` — identificador funcional |
| `idNCAsociada` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — FK conceptual a NoConformidades |
| `idEstadoInterno` | 4 (LongInteger) | 4 | true | `BIGINT NOT NULL` — FK conceptual a `tbEstados` ⚠️ sin constraint |
| `fechaCreacion` | 8 (DateTime) | 8 | true | `TIMESTAMP NOT NULL` |
| `usuarioCreacion` | 10 (Text) | 100 | true | `VARCHAR(100) NOT NULL` — auditoría |
| `fechaModificacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `usuarioModificacion` | 10 (Text) | 100 | false | `VARCHAR(100) NULL` — auditoría |
| `revisionCalidadEstado` | 10 (Text) | 20 | true | `VARCHAR(20) NOT NULL` — enum de estado de revisión |
| `revisionCalidadComentarios` | 12 (Memo) | 0 | false | `TEXT NULL` |

### `tbEstados` (6 columnas, 9 filas)

Catálogo de estados del workflow. Define el ciclo de vida de una Solicitud.

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `idEstado` | 4 (LongInteger) | 4 | false | `BIGSERIAL` PK (aunque `required:false` en schema, es lógica PK) |
| `nombreEstado` | 10 (Text) | 50 | true | `VARCHAR(50) NOT NULL` UNIQUE — nombre del estado |
| `descripcion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `esEstadoInicial` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `esEstadoFinal` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `orden` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — orden en el workflow |

### `tbTransiciones` (4 columnas, ⚠️ workflow declarativo)

**Tabla crítica**: define las transiciones permitidas entre estados y el **rol requerido** para ejecutar cada transición. Es un **workflow declarativo** — la nueva plataforma debe preservar este patrón.

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `idTransicion` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `idEstadoOrigen` | 4 (LongInteger) | 4 | true | `BIGINT NOT NULL` — FK conceptual a `tbEstados` |
| `idEstadoDestino` | 4 (LongInteger) | 4 | true | `BIGINT NOT NULL` — FK conceptual a `tbEstados` |
| `rolRequerido` | 10 (Text) | 50 | true | `VARCHAR(50) NOT NULL` — enum: Administrador, Calidad, Técnico |

**Disposición**: preservar como `transiciones` en PostgreSQL con FKs explícitas a `estados(origen)` y `estados(destino)`. La columna `rolRequerido` se traduce a una verificación de capabilities (D45-D46) en la nueva plataforma.

### `tbRechazos` (10 columnas, 0 filas)

Cabecera de Rechazo. Una Solicitud puede tener varios Rechazos (subsanación iterativa).

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `idRechazo` | 4 (LongInteger) | 4 | false | `BIGSERIAL` PK |
| `idSolicitud` | 4 (LongInteger) | 4 | true | `BIGINT NOT NULL` — FK conceptual a `tbSolicitudes` ⚠️ |
| `fechaRechazo` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `motivoPrincipal` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `areaAfectada` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `comentarios` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `usuarioRechazo` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `esActivo` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT TRUE` |
| `notasSubsanacion` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `CambiosTecnico` | 12 (Memo) | 0 | false | `TEXT NULL` |

### `tbHistorialRechazos` (presumido similar a `tbRechazos`; schema no inspeccionado)

⚠️ Schema no inspeccionado en esta pasada. Pendiente para una segunda iteración.

### `tbAdjuntos` (8 columnas, 0 filas)

Adjuntos de una Solicitud. Vinculados a `etapaWF` (etapa del workflow) y `TipoAccion`.

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `idAdjunto` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `idSolicitud` | 4 (LongInteger) | 4 | true | `BIGINT NOT NULL` — FK conceptual a `tbSolicitudes` |
| `etapaWF` | 10 (Text) | 50 | false | `VARCHAR(50) NULL` — etapa del workflow donde se subió |
| `nombreArchivo` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `fechaSubida` | 8 (DateTime) | 8 | true | `TIMESTAMP NOT NULL` |
| `usuarioSubida` | 10 (Text) | 100 | true | `VARCHAR(100) NOT NULL` — auditoría |
| `descripcion` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `TipoAccion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |

**Disposición**: en la nueva plataforma los adjuntos se migran al object storage (D16); la tabla local queda como metadatos con referencia al storage.

### `tbDatosPC` (41 columnas, 0 filas) — **Solicitud tipo PC**

**Tabla MUY ancha** (41 columnas) — datos completos de una Propuesta de Cambio (PC). Formulario regulatorio de procurement defense.

| Columna | Tipo DAO | Notas |
|---|---|---|
| `idDatosPC` | 4 | PK |
| `idSolicitud` | 4 | FK a `tbSolicitudes` |
| `refContratoInspeccionOficial` | 10 (Text 100, required) | Referencia del contrato |
| `refSuministrador` | 10 (Text 100, required) | Referencia del suministrador |
| `denominacionContrato` | 12 (Memo) | |
| `suministradorNombreDir` | 12 (Memo) | |
| `objetoContrato` | 12 (Memo) | |
| `descripcionMaterialAfectado` | 12 (Memo) | |
| `numPlanoEspecificacion` | 12 (Memo) | |
| `descripcionPropuestaCambio` | 12 (Memo) | |
| `descripcionPropuestaCambioCont` | 12 (Memo) | |
| `motivoCorregirDeficiencias` | 1 (YesNo) | |
| `motivoMejorarCapacidad` | 1 (YesNo) | |
| `motivoAumentarNacionalizacion` | 1 (YesNo) | |
| `motivoMejorarSeguridad` | 1 (YesNo) | |
| `motivoMejorarFiabilidad` | 1 (YesNo) | |
| `motivoMejorarCosteEficacia` | 1 (YesNo) | |
| `motivoOtros` | 1 (YesNo) | |
| `motivoOtrosDetalle` | 12 (Memo) | |
| `incidenciaCoste` | 10 (Text 50) | |
| `incidenciaPlazo` | 10 (Text 50) | |
| `incidenciaSeguridad` | 1 (YesNo) | |
| `incidenciaFiabilidad` | 1 (YesNo) | |
| `incidenciaMantenibilidad` | 1 (YesNo) | |
| `incidenciaIntercambiabilidad` | 1 (YesNo) | |
| `incidenciaVidaUtilAlmacen` | 1 (YesNo) | |
| `incidenciaFuncionamientoFuncion` | 1 (YesNo) | |
| `impactoClasificacion` | 10 (Text 255) | |
| `CambioAfectaAMaterial` | 10 (Text 255) | |
| `firmaOficinaTecnicaNombre` | 10 (Text 100) | |
| `firmaRepSuministradorNombre` | 10 (Text 100) | |
| `racCodigo` | 10 (Text 50) | Código RAC |
| `observacionesRAC` | 12 (Memo) | |
| `racNombre` | 10 (Text 255) | |
| `racDecision` | 10 (Text 50) | Decisión del RAC |
| `racRechazoMotivos` | 12 (Memo) | |
| `obsAprobacionAutoridadDiseno` | 12 (Memo) | |
| `NombreAutoridadDiseno` | 10 (Text 100) | |
| `decisionFinal` | 10 (Text 50) | |
| `obsDecisionFinal` | 12 (Memo) | |
| `NombreFirmanteFinal` | 10 (Text 100) | |

### `tbDatosCDCA` (39 columnas, 0 filas) — **Solicitud tipo CD_CA**

**Tabla MUY ancha** (39 columnas) — datos completos de una Comunicación de Desviación / Corrección de Anomalía (CD_CA).

| Columna | Tipo DAO | Notas |
|---|---|---|
| `idDatosCDCA` | 4 | PK |
| `idSolicitud` | 4 | FK a `tbSolicitudes` |
| `numContrato` | 10 (Text 100) | |
| `refSuministrador` | 10 (Text 100) | |
| `SuministradorNombreDir` | 12 (Memo) | |
| `refDesviacionesPrevias` | 10 (Text 100) | |
| `requiereModificacionContrato` | 1 (YesNo) | |
| `identificacionMaterial` | 12 (Memo) | |
| `numPlanoEspecificacion` | 12 (Memo) | |
| `cantidadPeriodo` | 10 (Text 50) | |
| `numSerieLote` | 10 (Text 100) | |
| `causaNC` | 12 (Memo) | |
| `descripcionImpactoNC` | 12 (Memo) | |
| `descripcionImpactoNCCont` | 12 (Memo) | |
| `afectaPrestaciones` | 1 (YesNo) | |
| `afectaSeguridad` | 1 (YesNo) | |
| `afectaFiabilidad` | 1 (YesNo) | |
| `afectaVidaUtil` | 1 (YesNo) | |
| `afectaMedioambiente` | 1 (YesNo) | |
| `afectaIntercambiabilidad` | 1 (YesNo) | |
| `afectaMantenibilidad` | 1 (YesNo) | |
| `afectaApariencia` | 1 (YesNo) | |
| `afectaOtros` | 1 (YesNo) | |
| `impactoCoste` | 10 (Text 50) | |
| `clasificacionNC` | 10 (Text 50) | |
| `esSuministradorAD` | 1 (YesNo) | Es suministrador la autoridad de diseño |
| `identificacionAutoridadDiseno` | 10 (Text 100) | |
| `efectoFechaEntrega` | 12 (Memo) | |
| `firmaAprobacionRespIngenieriaNombre` | 10 (Text 255) | |
| `firmaAprobacionRespProduccionNombre` | 10 (Text 255) | |
| `firmaAprobacionRespCalidadNombre` | 10 (Text 255) | |
| `firmaAprobacionRespDisenioNombre` | 10 (Text 255) | |
| `firmaAprobacionRepresentanteSumNombre` | 10 (Text 255) | |
| `racCodigo` | 10 (Text 50) | |
| `observacionesRAC` | 12 (Memo) | |
| `racNombre` | 10 (Text 255) | |
| `racDecision` | 10 (Text 50) | |
| `decisionFinal` | 10 (Text 50) | |
| `observacionesFinales` | 12 (Memo) | |
| `NombreFirmanteFinal` | 10 (Text 100) | |

### `tbDatosCDCASUB` y `tbDatosPCSUB` (schemas no inspeccionados en esta pasada)

⚠️ Pendiente para una segunda iteración. Análogos a `tbDatosCDCA` y `tbDatosPC` pero para los sub-tipos `CD_CA_SUB` y `PC_SUB`.

### `tbValidacionRevision` (11 columnas) — **hash para validación**

Tabla de validación de revisión con **hash de datos**.

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `Id` | 4 (LongInteger) | 4 | false | `BIGSERIAL` PK |
| `idSolicitud` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — FK conceptual a `tbSolicitudes` ⚠️ |
| `ordinal` | 3 (Integer) | 2 | false | `SMALLINT NULL` |
| `idAdjunto` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — FK conceptual a `tbAdjuntos` ⚠️ |
| `FechaEnvio` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaRecepcion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `Resultado` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Comentarios` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `Usuario` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `HashDatos` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **hash para validación** |

⚠️ **`HashDatos`** se preserva tal cual — la nueva plataforma usa este hash para verificar integridad.

### `tbLogCambios` (11 columnas, 5 filas) — **log de cambios general con suplantación**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `idLogCambio` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `fechaHora` | 8 (DateTime) | 8 | true | `TIMESTAMP NOT NULL` |
| `usuario` | 10 (Text) | 100 | true | `VARCHAR(100) NOT NULL` |
| `suplantadoPor` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **quién suplantó al usuario** (registro de impersonación) |
| `tabla` | 10 (Text) | 50 | true | `VARCHAR(50) NOT NULL` |
| `registro` | 4 (LongInteger) | 4 | true | `BIGINT NOT NULL` |
| `campo` | 10 (Text) | 50 | false | `VARCHAR(50) NULL` |
| `valorAnterior` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `valorNuevo` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `tipoOperacion` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` — INSERT, UPDATE, DELETE |

⚠️ **`suplantadoPor`** confirma que **Condor tiene un sistema de impersonación** que registra quién suplantó a quién en cada cambio. Coherente con `rolUsuario` / `rolUsuarioReal` (preparación) y `g_blnImpersonando` (estado). Migrar a logs estructurados canónicos (D27).

### `tbLogErrores` (schema no inspeccionado en detalle; 3 filas)

⚠️ Pendiente para una segunda iteración.

### `tbLogEstados` (6 columnas, 0 filas) — **log de transiciones de estado**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `idLogEstado` | 4 (LongInteger) | 4 | false | `BIGSERIAL` PK |
| `idSolicitud` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — FK conceptual a `tbSolicitudes` ⚠️ |
| `idEstadoAnterior` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — FK conceptual a `tbEstados` ⚠️ |
| `idEstadoNuevo` | 4 (LongInteger) | 4 | false | `BIGINT NULL` — FK conceptual a `tbEstados` ⚠️ |
| `fechaTransicion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `usuarioTransicion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |

⚠️ **NO hay FKs físicas declaradas** desde `idSolicitud`, `idEstadoAnterior`, `idEstadoNuevo`. Data integrity gap.

### `tbMapeoCampos` (schema no inspeccionado; 183 filas)

⚠️ **Pendiente para segunda iteración**. Volumen alto (183 filas) sugiere que es **config de mapeo entre columnas legacy y modernas** — debe preservarse como datos (no como código).

### Pendientes de discovery (segunda pasada)

- **`tbDatosCDCASUB`**, **`tbDatosPCSUB`**: schemas no inspeccionados.
- **`tbLogErrores`**: schema no inspeccionado en detalle.
- **`tbMapeoCampos`**: schema no inspeccionado en detalle (183 filas).
- **`tbHistorialRechazos`**: schema no inspeccionado.

Estos schemas se obtendrán en una iteración posterior.

### Relaciones físicas reales (5 FK desde `tbSolicitudes`, parciales)

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
4. **`tbRechazos`, `tbLogEstados`, `tbValidacionRevision`, `tbAdjuntos`, `tbHistorialRechazos` no tienen FKs explícitas a `tbSolicitudes`** pese a tener `idSolicitud`. Data integrity gap.
5. **`tbValidacionRevision.idAdjunto` no tiene FK explícita a `tbAdjuntos`**. Data integrity gap.
6. **`tbLogEstados.idEstadoAnterior` / `idEstadoNuevo` no tienen FKs explícitas a `tbEstados`**. Data integrity gap.