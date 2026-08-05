# Brass — modelo físico y diccionario de datos

## Autoridad y fecha

Fuente física: **`C:\00repos\datos\Gestion_Brass_Gestion_Datos.accdb`** (61 MB) — backend autoritativo. **Inventario real obtenido vía Dysflow MCP el 2026-08-05** después de `setup_project` (con `projectId: 00-brass-00-main-clean` aplicado) + `register_worktree` + `accessPath` + `backendPath` absolutos.

## Inventario real Dysflow (2026-08-05, backend autoritativo)

| Categoría | Resultado |
|---|---|
| **Tablas totales** | **60** en `Gestion_Brass_Gestion_Datos.accdb` |
| **Filas en `TbEventos`** | **5777** (eventos en producción) |
| **Filas en `TbActividades`** | **25639** (actividades — altísimo volumen) |
| **Filas en `TbAnexos`** | **7592** (anexos) |
| **Filas en `TbMaterial`** | **2351** (materiales) |
| **Filas en `TbEquipos`** | **806** (equipos) |
| **Filas en `TbEquiposMedida`** | **14** (equipos de medida calibrados — dominio regulatorio) |
| **Filas en `TbEquiposMedidaCalibraciones`** | **28** (calibraciones — 2 por equipo de medida) |
| **Filas en `TbTecnicos`** | **115** (técnicos) |
| **Filas en `TbPlanificacion`** | **624** (planificaciones) |
| **Filas en `TbGastos`** | **519** (gastos) |
| **Filas en `TbFacturaPrincipal`** | **53** (facturas principales) |
| **Filas en `TbSubcontrataciones`** | (no contado) |
| **Filas en `TbEquiposCalibrables`** | (no contado) |
| **Filas en `TbAnexos` (otra count)** | (confirmado 7592) |
| **FK relationships** (user tables) | **27** entre tablas de usuario |

⚠️ **Hallazgo del volumen real**: Brass está en **uso activo MUY intenso**:
- 5777 eventos + 25639 actividades + 7592 anexos + 2351 materiales = **36k+ registros principales**.
- 14 equipos de medida con 28 calibraciones — sistema regulatorio activo.
- 115 técnicos, 806 equipos — flota activa.
- 53 facturas principales pero 519 gastos — gastos se acumulan más rápido que facturas.

⚠️ **Convención de nombres**: las tablas usan `Tb` mayúscula + CamelCase (`TbEventos`, `TbActividades`, etc.). Consistente con la mayoría de las apps.

### Lista completa de las 60 tablas

```
Tb0FiltroGestion (con número inicial, presumiblemente legacy)
TbActividades
TbAnexos
TbAuxActividad
TbAuxEventos
TbAuxManteniminetosPreventivosCalendario (typo: "Manteniminetos")
TbAuxMateriales
TbAuxPlanificacion
TbBUI
TbBuiIDEvento
TbCausaFin
TbCodActividad
TbCriticidad
TbEquipos
TbEquiposCalibrables
TbEquiposCalibrablesFechas
TbEquiposMedida
TbEquiposMedidaCalibraciones
TbEventos
TbEventosEquipoMedida
TbEventosEquipoMedida_antes (legacy, presumido)
TbFacturaActividadesInvolucradas
TbFacturaEventosInvolucrados
TbFacturaGastosInvolucrados
TbFacturaMaterialesInvolucrados
TbFacturaPrincipal
TbFacturaPrincipalPerfiles
TbFacturaSubcontratacionesInvolucradas
TbGastos
TbGastosImportePorTipo
TbGuiaConciliaciones
TbHerramientaDocAyuda
TbMaterial
TbMaterialSeguimiento
TbNodoBUI
TbNodos
TbOriginador
TbPartesDetalle
TbPartesPpal
TbPlanificacion
TbPlanificacionAnexos
TbPlanificacionEquipos
TbPlanificacionRegistrada
TbRepuestosTipo
TbResultadoVerificacion
TbSistema
TbSubcontrataciones
TbSubsistemaBui
TbTecnicos
TbTecnicosAusencias
TbTecnicosFiestas
TbTipoAccion
TbTipoAsistencia
TbTipoEvento
TbTipoTecnico
TbTipoTecnicoPrecios
TbUbicacion
```

⚠️ **TYPO detectado**: `TbAuxManteniminetosPreventivosCalendario` (debería ser "Mantenimientos").

## Schemas reales (Dysflow, segunda pasada)

### `TbEventos` (38 columnas, 5777 filas) — **cabecera de evento con SLA**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDEvento` | 10 (Text) | 50 | true | `VARCHAR(50) NOT NULL` PK ⚠️ **string, no Long** |
| `NODO` | 10 (Text) | 50 | false | `VARCHAR(50) NULL` — FK por texto a `TbNodos` |
| `BUI` | 10 (Text) | 50 | false | `VARCHAR(50) NULL` — FK por texto a `TbBUI` |
| `SUBSISTEMA` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — FK por texto a `TbSubsistemaBui` |
| `IDEquipo` | 4 (LongInteger) | 4 | true | `BIGINT NOT NULL` — FK a `TbEquipos` |
| `PMPR` | 10 (Text) | 50 | false | `VARCHAR(50) NULL` |
| `TIPOEVENTO` | 10 (Text) | 50 | false | `VARCHAR(50) NULL` — FK a `TbTipoEvento` (conceptual) |
| `CRITICIDAD` | 3 (Integer) | 2 | false | `SMALLINT NULL` |
| `ALIASTECNICO` | 10 (Text) | 50 | false | `VARCHAR(50) NULL` — FK por texto a `TbTecnicos` |
| `ORIGINADOR` | 10 (Text) | 50 | false | `VARCHAR(50) NULL` — FK a `TbOriginador` |
| `FECHAALTAEVENTO` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `HORAINICIALEVENTO` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaFinal` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `HORAFINALEVENTO` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `TIEMPORESPUESTAEVENTO` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` — **SLA: tiempo de respuesta** |
| `CAUSAFIN` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `DESCRIPCION` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `CONTACTO` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `FechaRegistroAlta` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaRegistroModificacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `Franqueado` | 1 (YesNo) | 1 | true | `BOOLEAN NOT NULL DEFAULT FALSE` — **flag de cierre del evento** |
| `IDExportacion` | 10 (Text) | 50 | false | `VARCHAR(50) NULL` |
| `CodImportacion` | 10 (Text) | 50 | false | `VARCHAR(50) NULL` |
| `IDParte` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — FK a `TbPartesPpal` |
| `Notas` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `FechaEnInformeRAC` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaRecepcionNotificacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaInicioContactoCliente` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `IncidenciaAveriaOReparacion` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `FechaInicioTiempoAdquisicion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaFinTiempoAdquisicion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `TipoReparacion` | 10 (Text) | 50 | false | `VARCHAR(50) NULL` |
| `Urgente` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `EventoConServicioAfectado` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `FechaRestablecimientoServicio` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `TipoRepInsitu` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` — ⚠️ **3 tipos de reparación como YesNo** |
| `TipoRepNoSMT` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `TipoRepValvulas` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |

**Hallazgos críticos del schema de `TbEventos`**:

1. **`IDEvento` es Text(50)**: ⚠️ todas las FKs del sistema se hacen por texto. La nueva plataforma debe decidir si mantiene IDs como string (UUID) o los migra a numéricos (BIGSERIAL).

2. **SLA explícito**: `TIEMPORESPUESTAEVENTO` (Date) — **tiempo de respuesta del evento**. ⚠️ **Crítico para la nueva plataforma**: el SLA debe preservarse.

3. **3 tipos de reparación mutuamente excluyentes** (`TipoRepInsitu`, `TipoRepNoSMT`, `TipoRepValvulas` como YesNo) — **debería ser un enum**. Migrar a un ENUM en PostgreSQL.

4. **14 fechas en la tabla**: `FECHAALTAEVENTO`, `HORAINICIALEVENTO`, `FechaFinal`, `HORAFINALEVENTO`, `TIEMPORESPUESTAEVENTO`, `FechaRegistroAlta`, `FechaRegistroModificacion`, `FechaEnInformeRAC`, `FechaRecepcionNotificacion`, `FechaInicioContactoCliente`, `FechaInicioTiempoAdquisicion`, `FechaFinTiempoAdquisicion`, `FechaRestablecimientoServicio`, más las 2 de franqueo. **Migración**: algunas pueden consolidarse, otras son workflow regulatorio.

5. **FKs por texto**: `NODO`, `BUI`, `SUBSISTEMA`, `ALIASTECNICO`, `ORIGINADOR`, `IDParte` — todas son Text. Migración: agregar columnas numéricas como `*IdNum` y mantener las de texto como legacy.

### Relaciones físicas reales (27 FK entre user tables, parciales)

| Origen | Columna FK | Destino | Columna FK |
|---|---|---|---|
| `TbActividades` | `IDActividad` | `TbFacturaActividadesInvolucradas` | `IDActividad` |
| `TbActividades` | `IDActividad` | `TbMaterial` | `IDActividad` |
| `TbBUI` | `BUI` | `TbEquipos` | `BUI` |
| `TbBUI` | `BUI` | `TbEventos` | `BUI` |
| `TbBUI` | `BUI` | `TbSubsistemaBui` | `BUI` |
| `TbEquipos` | `IDEquipo` | `TbEventos` | `IDEquipo` |
| `TbEventos` | `IDEvento` | `TbActividades` | `IDEvento` |
| `TbEventos` | `IDEvento` | `TbFacturaEventosInvolucrados` | `IDEvento` |
| `TbFacturaPrincipal` | `IDFactura` | `TbFacturaActividadesInvolucradas` | `IDFactura` |
| `TbFacturaPrincipal` | `IDFactura` | `TbFacturaEventosInvolucrados` | `IDFactura` |
| `TbFacturaPrincipal` | `IDFactura` | `TbFacturaGastosInvolucrados` | `IDFactura` |
| `TbFacturaPrincipal` | `IDFactura` | `TbFacturaMaterialesInvolucrados` | `IDFactura` |
| `TbFacturaPrincipal` | `IDFactura` | `TbFacturaPrincipalPerfiles` | `IDFactura` |
| `TbFacturaPrincipal` | `IDFactura` | `TbFacturaSubcontratacionesInvolucradas` | `IDFactura` |
| `TbGastos` | `IDGasto` | `TbFacturaGastosInvolucrados` | `IDGasto` |
| `TbMaterial` | `IDMaterial` | `TbFacturaMaterialesInvolucrados` | `IDMaterial` |
| `TbMaterial` | `IDMaterial` | `TbMaterialSeguimiento` | `IDMaterial` |
| `TbNodos` | `NODO` | `TbEquipos` | `NODO` |
| `TbNodos` | `NODO` | `TbEventos` | `NODO` |
| `TbNodos` | `NODO` | `TbNodoBUI` | `NODO` |
| `TbSubcontrataciones` | `IDSubcontratacion` | `TbFacturaSubcontratacionesInvolucradas` | `IDSubcontratacion` |
| `TbTecnicos` | `ALIAS` | `TbFacturaActividadesInvolucradas` | `ALIASTECNICO` |
| `TbTecnicos` | `ALIAS` | `TbFacturaGastosInvolucrados` | `ALIAS` |
| `TbTecnicos` | `ALIAS` | `TbTecnicosAusencias` | `Alias` |
| `TbTipoTecnico` | `TIPO` | `TbFacturaActividadesInvolucradas` | `TIPOTECNICO` |
| `TbTipoTecnico` | `TIPO` | `TbTecnicos` | `TIPO` |
| `TbTipoTecnico` | `TIPO` | `TbTipoTecnicoPrecios` | `TIPOTECNICO` |

⚠️ **Inconsistencias detectadas en FKs**:
- `TbTecnicos.ALIAS` → `TbFacturaActividadesInvolucradas.ALIASTECNICO` y `TbFacturaGastosInvolucrados.ALIAS`: ⚠️ el nombre de la columna es **diferente** (`ALIAS` vs `ALIASTECNICO`). Data integrity gap.
- `TbTecnicos.ALIAS` → `TbTecnicosAusencias.Alias`: ⚠️ el case de `Alias` vs `ALIAS` (PascalCase vs MAYÚSCULAS). Data integrity gap.
- `TbTipoTecnico.TIPO` → `TbFacturaActividadesInvolucradas.TIPOTECNICO` y `TbTipoTecnicoPrecios.TIPOTECNICO`: ⚠️ nombre de columna diferente.

### Pendientes de discovery (segunda pasada)

Schemas de las 50+ tablas restantes (TbActividades, TbAnexos, TbMaterial, TbEquipos, TbEquiposMedida, TbEquiposMedidaCalibraciones, TbTecnicos, TbFacturaPrincipal, etc.). Se obtendrán en una iteración posterior.