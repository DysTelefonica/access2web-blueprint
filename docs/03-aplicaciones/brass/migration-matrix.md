# Brass — matriz de migración (scaffold de preservación)

Esta matriz no decide el esquema PostgreSQL. Su regla es conservadora: todo campo/registro se conserva hasta que negocio apruebe una disposición. Disposiciones específicas de Brass en § D104-D107.

| Fuente | Significado/candidato de dominio | Transformación | Estado | Validación/rechazo | Decisión abierta |
|---|---|---|---|---|---|
| `TbEventos` (5777 filas) | evento regulatorio con SLA | separar identidad, fechas, SLA, franqueo, facturación | mapeado preliminar | unicidad por `IDEvento` (Text); **ID string ⚠️** | **D104** (contraseña); **D105** (ID string) |
| `TbActividades` (25639 filas) | actividad de un evento | FK a `TbEventos` + datos específicos | mapeado preliminar | cardinalidad 1:N con `TbEventos` | cardinalidad |
| `TbFacturaPrincipal` (53 filas) | cabecera de factura | separar identidad, fecha, importes, perfil | mapeado preliminar | unicidad; cardinalidad con 5 tablas de detalle | cardinalidad |
| `TbFacturaActividadesInvolucradas` | actividades en una factura | FK a `TbFacturaPrincipal` + `TbActividades` | mapeado preliminar | cardinalidad N:M | cardinalidad |
| `TbFacturaEventosInvolucrados` | eventos en una factura | FK a `TbFacturaPrincipal` + `TbEventos` | mapeado preliminar | cardinalidad N:M | cardinalidad |
| `TbFacturaGastosInvolucrados` | gastos en una factura | FK a `TbFacturaPrincipal` + `TbGastos` | mapeado preliminar | cardinalidad N:M | cardinalidad |
| `TbFacturaMaterialesInvolucrados` | materiales en una factura | FK a `TbFacturaPrincipal` + `TbMaterial` | mapeado preliminar | cardinalidad N:M | cardinalidad |
| `TbFacturaSubcontratacionesInvolucradas` | subcontrataciones en una factura | FK a `TbFacturaPrincipal` + `TbSubcontrataciones` | mapeado preliminar | cardinalidad N:M | cardinalidad |
| `TbFacturaPrincipalPerfiles` | perfiles de factura | FK a `TbFacturaPrincipal` | mapeado preliminar | cardinalidad 1:N con `TbFacturaPrincipal` | cardinalidad |
| `TbGuiaConciliaciones` | guías de conciliación | entidad de auditoría | mapeado preliminar | unicidad; relación con facturas (presumida) | retencion |
| `TbEquipos` (806 filas) | equipos (no calibrados) | entidad de catálogo | mapeado preliminar | unicidad; jerarquía BUI/Nodo | cardinalidad |
| `TbEquiposCalibrables` (volumen TBD) | equipos que requieren calibración | entidad de catálogo | mapeado preliminar | unicidad; FK a `TbEquipos` (presumida) | cardinalidad |
| `TbEquiposCalibrablesFechas` | fechas de calibración por equipo | FK a `TbEquiposCalibrables` | mapeado preliminar | cardinalidad 1:N | cardinalidad |
| `TbEquiposMedida` (14 filas) | equipos de medida (calibrados) | entidad de catálogo con calibración | mapeado preliminar | unicidad; **dominio regulatorio** | versionado |
| `TbEquiposMedidaCalibraciones` (28 filas) | calibraciones de equipos de medida | FK a `TbEquiposMedida` + `TbEventosEquipoMedida` | mapeado preliminar | **dominio regulatorio CRÍTICO** | retencion + versionado |
| `TbEventosEquipoMedida` (volumen TBD) | equipos de medida asignados a un evento | FK a `TbEventos` + `TbEquiposMedida` + `TbEquiposMedidaCalibraciones` | mapeado preliminar | cardinalidad N:M | cardinalidad |
| `TbEventosEquipoMedida_antes` | legacy (presumido) | descartar o archivar | mapeado preliminar | comparar recuentos y uso en código | archivo vs dominio |
| `TbMaterial` (2351 filas) | materiales | entidad de catálogo | mapeado preliminar | unicidad; FK a `TbActividades` | cardinalidad |
| `TbMaterialSeguimiento` | seguimiento de materiales | FK a `TbMaterial` | mapeado preliminar | cardinalidad 1:N | cardinalidad |
| `TbAnexos` (7592 filas) | anexos de eventos/actividades | FK a `TbEventos` + path externo | external reference | comprobar existencia de fichero/URL sin copiar datos | repositorio destino (D16) |
| `TbTecnicos` (115 filas) | técnicos | entidad de catálogo | mapeado preliminar | unicidad; **FK por `ALIAS` Text ⚠️** | cardinalidad |
| `TbTecnicosAusencias` | ausencias de técnicos | FK a `TbTecnicos` (⚠️ `Alias` vs `ALIAS`) | mapeado preliminar | cardinalidad 1:N | cardinalidad |
| `TbTecnicosFiestas` | fiestas de técnicos | FK a `TbTecnicos` (presumida) | mapeado preliminar | cardinalidad 1:N | cardinalidad |
| `TbTipoTecnico` | tipos de técnico | entidad de catálogo | mapeado preliminar | unicidad; **FK por `TIPO` Text** | versionado |
| `TbTipoTecnicoPrecios` | precios por tipo de técnico | FK a `TbTipoTecnico` | mapeado preliminar | cardinalidad 1:N | cardinalidad |
| `TbTecnicosHoras` (presumido) | horas trabajadas por técnico | FK a `TbTecnicos` (presumida) | mapeado preliminar | cardinalidad 1:N | cardinalidad |
| `TbGastos` (519 filas) | gastos | entidad de gastos | mapeado preliminar | unicidad; FK a `TbEventos` (presumida) | cardinalidad |
| `TbGastosImportePorTipo` | importes de gastos por tipo | FK a `TbGastos` | mapeado preliminar | cardinalidad 1:N | cardinalidad |
| `TbSubcontrataciones` | subcontrataciones | entidad de subcontratación | mapeado preliminar | unicidad; FK a `TbEventos` (presumida) | cardinalidad |
| `TbPlanificacion` (624 filas) | planificación | entidad de planificación | mapeado preliminar | unicidad; FK a `TbEventos` (presumida) | cardinalidad |
| `TbPlanificacionAnexos` | anexos de planificación | FK a `TbPlanificacion` | mapeado preliminar | cardinalidad 1:N | cardinalidad |
| `TbPlanificacionEquipos` | equipos de planificación | FK a `TbPlanificacion` + `TbEquipos` | mapeado preliminar | cardinalidad N:M | cardinalidad |
| `TbPlanificacionRegistrada` | planificación registrada (presumido) | FK a `TbPlanificacion` (presumida) | mapeado preliminar | cardinalidad 1:1 | cardinalidad |
| `TbPartesPpal` | cabecera de parte de trabajo | entidad de parte | mapeado preliminar | unicidad; FK a `TbEventos` (conceptual) | cardinalidad |
| `TbPartesDetalle` | detalle de parte de trabajo | FK a `TbPartesPpal` | mapeado preliminar | cardinalidad 1:N | cardinalidad |
| `TbBUI` | Business Unit Identifier | entidad de catálogo | mapeado preliminar | unicidad; jerarquía BUI | cardinalidad |
| `TbNodos` | nodos de la organización | entidad de catálogo | mapeado preliminar | unicidad; FK por `NODO` Text ⚠️ | cardinalidad |
| `TbNodoBUI` | relación nodo-BUI | FK a `TbNodos` + `TbBUI` | mapeado preliminar | cardinalidad N:M | cardinalidad |
| `TbSubsistemaBui` | subsistemas de BUI | FK a `TbBUI` | mapeado preliminar | cardinalidad 1:N | cardinalidad |
| `TbBuiIDEvento` | relación BUI-evento | FK a `TbBUI` + `TbEventos` (presumida) | mapeado preliminar | cardinalidad N:M | cardinalidad |
| `TbOrigenador` | originador (catálogo) | entidad de catálogo | mapeado preliminar | unicidad; FK a `TbEventos.ORIGINADOR` | versionado |
| `TbCausaFin` | causas de fin de evento | entidad de catálogo | mapeado preliminar | unicidad; FK a `TbEventos.CAUSAFIN` | versionado |
| `TbTipoEvento` | tipos de evento | entidad de catálogo | mapeado preliminar | unicidad | versionado |
| `TbTipoAccion` | tipos de acción | entidad de catálogo | mapeado preliminar | unicidad | versionado |
| `TbTipoAsistencia` | tipos de asistencia | entidad de catálogo | mapeado preliminar | unicidad | versionado |
| `TbCriticidad` | niveles de criticidad | entidad de catálogo | mapeado preliminar | unicidad | versionado |
| `TbCodActividad` | códigos de actividad | entidad de catálogo | mapeado preliminar | unicidad | versionado |
| `TbRepuestosTipo` | tipos de repuestos | entidad de catálogo | mapeado preliminar | unicidad | versionado |
| `TbUbicacion` | ubicaciones | entidad de catálogo | mapeado preliminar | unicidad | versionado |
| `TbSistema` | sistemas | entidad de catálogo | mapeado preliminar | unicidad; jerarquía BUI | cardinalidad |
| `TbResultadoVerificacion` | resultados de verificación | entidad de catálogo | mapeado preliminar | unicidad; FK a `TbEventos` (presumida) | versionado |
| `TbConfiguracion` (presumido) | configuración del sistema | config key-value | mapeado preliminar | unicidad | migrar a config del módulo |
| `TbHerramientaDocAyuda` | doc de ayuda | contenido estático | mapeado preliminar | unicidad | migrar a docs |
| `Tb0FiltroGestion` | filtro de gestión (presumido) | config legacy | mapeado preliminar | unicidad; **con número inicial `0` ⚠️ legacy** | descartar o archivar |
| `TbAuxActividad` | auxiliar de actividad | descartar o archivar | mapeado preliminar | comparar recuentos | archivo vs dominio |
| `TbAuxEventos` | auxiliar de eventos | descartar o archivar | mapeado preliminar | comparar recuentos | archivo vs dominio |
| `TbAuxManteniminetosPreventivosCalendario` (typo) | auxiliar de mantenimientos preventivos | descartar o archivar | mapeado preliminar | comparar recuentos | archivo vs dominio |
| `TbAuxMateriales` | auxiliar de materiales | descartar o archivar | mapeado preliminar | comparar recuentos | archivo vs dominio |
| `TbAuxPlanificacion` | auxiliar de planificación | descartar o archivar | mapeado preliminar | comparar recuentos | archivo vs dominio |
| catálogos compartidos con Lanzadera (vía `getdbLanzadera`) | identidad / permisos / entorno | identidad vía adaptador unificado (D9-D10) | mapeado preliminar | coherencia | acoplamiento directo a romper (D86/D87) |

## Ledger

- **Mapeado preliminar:** cabecera de eventos, actividades, facturación completa, equipos con calibración, técnicos, planificación, partes, BUI.
- **Necesita decisión:** D104 (contraseña), D105 (ID string), D106 (FKs por texto), D107 (booleanos como YesNo).
- **External reference:** anexos, identidad, Lanzadera.
- **No hay campos/filas declarados obsoletos** (salvo los marcados en D104-D107).

## D104 · Contraseña hardcodeada `"dpddpd"` como REAL en `Variables Globales.bas:560`

**Estado**: PROPUESTO. **CRÍTICO** ⚠️⚠️⚠️. Detalle completo en [Seguridad § D104](security-rules.md#d104--contraseña-hardcodeada-dpddpd-como-real-en-variables-globalesbas560).

Resumen:

- Brass NO usa fallback como Condor (D93). La contraseña `"dpddpd"` se pasa **directamente** a `OpenDatabase` en `Variables Globales.bas:560`:
  ```vba
  Set db = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
  ```
- Es la **contraseña real del backend** (no un fallback).
- Si el repo `00_BRASS` se versiona, esta contraseña queda en git.
- La rotación de la contraseña NO surte efecto sin cambio de código.

## D105 · `IDEvento` es `Text(50)`, no Long

**Estado**: PROPUESTO.

Todas las FKs del sistema Brass se hacen por **texto** (no por ID numérico). Esto es una **decisión de diseño legacy**. La nueva plataforma debe decidir si:

- **Mantiene IDs como string** (UUID, más portable cross-database, más compatible con la integración).
- **Migra a numéricos** (BIGSERIAL, más eficiente para joins, más consistente con otras apps).

Decisión pendiente con el equipo de backend. **Mientras tanto, preservar tal cual**.

## D106 · FKs por texto (coherente con D105)

**Estado**: PROPUESTO.

`NODO`, `BUI`, `SUBSISTEMA`, `ALIASTECNICO`, `ORIGINADOR`, `IDParte` son `Text` y se usan como FKs. Esto es una **decisión de diseño legacy** que se debe **documentar y preservar** en la nueva plataforma (decisión coherente con D105).

⚠️ **Inconsistencias detectadas en nombres de columnas de FK**:
- `TbTecnicos.ALIAS` → `TbFacturaActividadesInvolucradas.ALIASTECNICO` (⚠️ nombre diferente).
- `TbTecnicos.ALIAS` → `TbTecnicosAusencias.Alias` (⚠️ case diferente).
- `TbTipoTecnico.TIPO` → `TbFacturaActividadesInvolucradas.TIPOTECNICO` (⚠️ nombre diferente).

Migración: en PostgreSQL, **unificar nombres de columnas** (`ALIAS` en todos los casos) y agregar columnas numéricas como `*IdNum` para joins eficientes.

## D107 · Booleanos como YesNo (mejor que Text 2)

**Estado**: PROPUESTO.

A diferencia de otras apps (que usan `Text 2` para 'Sí'/'No'), Brass usa **`YesNo` real** (type 1) en la mayoría de las columnas booleanas (`Franqueado`, `IncidenciaAveriaOReparacion`, `Urgente`, `EventoConServicioAfectado`, `TipoRepInsitu`, `TipoRepNoSMT`, `TipoRepValvulas`).

⚠️ **Inconsistencia cross-cutting D102**: otras apps usan `Text 2` para booleanos. Brass está mejor en este aspecto. **Recomendación**: en la nueva plataforma, estandarizar a `BOOLEAN` (Brass ya lo hace correctamente).