# Audit — Form_FormExpedienteHitos (Phase 3.3 / PR-9)

> **Change**: `forms-thin-coverage` (Phase 3.3 — PR-9, batch of 6 forms)
> **Form**: `Form_FormExpedienteHitos` (pestaña de hitos del expediente)
> **Date**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight)

## 1. STATUS

| Fase | Estado |
|---|---|
| Audit | **COMPLETE** |
| Helper extraction | **DONE** (6 helpers) |
| Binario sync | DONE |
| Commit | **DONE** |

## 2. FORM METADATA

| Atributo | Valor |
|---|---|
| Líneas .cls | 257 |
| Handlers totales | 3 (Form_Load + ComandoRegistrar + ComandoEliminar) |
| Public methods | 3 (`RellenarListaHitos`, `RellenarListas`, `EstablecerDatos`) — **MANTENER firmas** |
| Public events | 0 |
| WithEvents | 0 |
| MsgBox call sites | 3 |
| DoCmd.OpenForm en .cls | 0 |
| Forms abiertas | 0 |

## 3. CONTROLES REFERENCIADOS (Me.X)

| Control | Tipo | Usado en |
|---|---|---|
| `FechaHito` | TextBox | ComandoRegistrar, Form_Load |
| `DESCRIPCION` | TextBox | ComandoRegistrar |
| `FechaGarantiaHito` | TextBox | ComandoRegistrar |
| `Importe` | TextBox | ComandoRegistrar |
| `ListaHitos` | ListBox | ComandoRegistrar, ComandoEliminar, RellenarListaHitos, RellenarListas |
| `ComandoRegistrar` | CommandButton (parent) | Form_Load |
| `AllowEdits` | Form property | Form_Load |

## 4. ENTIDAD Y PROPIEDADES REFERENCIADAS

| Entidad | Propiedades usadas |
|---|---|
| `ExpedienteOperaciones` | `Expediente`, `RegistrarHito(p_Hito, p_Error)`, `EliminarHito(p_FechaHito, p_Error)` |
| `ExpedienteDTO` (global) | `Expediente`, `ColHitos` |
| `Expediente` | `Hitos` |
| `ExpedienteHito` | `FechaHito`, `DESCRIPCION`, `FechaGarantiaHito`, `Importe` |

## 5. INTER-FORM CALLSITES

| Target | Mecanismo |
|---|---|
| `MostrarPopupProgreso` / `CerrarPopupProgreso` | Module call (UI guard) |
| `CorreoAlAdministrador` | Module call (error notification) |
| `EsAdministrador`, `AbiertoParaEditar` | Module call (auth/state) |

## 6. ANTI-PATTERN DETECTADO

⚠️ **`pregunta = MsgBox(...)`** (3 sites) — bug pre-existente, ya no compila bajo Option Explicit.
⚠️ **`On Error GoTo errores + Err.Raise 1000`** (3 handlers) — patrón de control flow.
⚠️ **`VBA.DoEvents / DoCmd.Hourglass True/False`** en handlers — UI guard se queda en form pero helper no debe repetirlo.

## 7. RECOMENDACIÓN DE PATRÓN

**Pattern**: SIMPLE — alta/baja hitos. Helper recibe Dictionary wrapper con `{Expediente, ColHitos}`.

### 7.1 Helpers implementados (6)

| Helper | Signature | Returns |
|---|---|---|
| `ExpedienteHitos_AltaHito` | `(p_DTO, p_FechaHito, p_Descripcion, p_FechaGarantiaHito, p_Importe, p_Error)` | `{registrado, colRefreshed}` |
| `ExpedienteHitos_EliminarHito` | `(p_DTO, p_FechaHito, p_EsAdmin, p_Error)` | `{eliminado, colRefreshed}` |
| `ExpedienteHitos_RellenarListaHitos` | `(p_DTO, p_Error)` | `{rowSource, count}` |
| `ExpedienteHitos_RellenarListas` | `(p_DTO, p_Error)` | `{rowSource, count}` |
| `ExpedienteHitos_EstablecerDatos` | `(p_DTO, p_EsAdmin, p_AbiertoParaEditar, p_Error)` | `{ejecutivosEnabled, expedienteOK}` |
| `ExpedienteHitos_Form_Load` | `(p_DTO, p_EsAdmin, p_AbiertoParaEditar, p_Error)` | `{expedienteOK, ejecutivoEnabled, hideComandoRegistrar}` |

### 7.2 UI orchestration que se queda en el form
- `ComandoRegistrar_Click` — UI guard + AltaHito + RellenarListaHitos
- `ComandoEliminar_Click` — UI guard + EliminarHito + RellenarListaHitos
- `Form_Load` — orquesta inicialización
- `EstablecerDatos` / `RellenarListas` / `RellenarListaHitos` — MANTENER firmas; cuerpo: helper + apply RowSource

## 8. NOTAS Y RIESGOS

1. **`m_ObjExpedienteDTOActivo.Expediente.Hitos`** — refresh después de alta/baja; helper lo hace best-effort.
2. **`pregunta = MsgBox(...)`** — bug pre-existente corregido en form thinned.

## 9. REWORK (2026-06-26, PR-9 / Phase 3.3)

### 9.1 Files added

| File | Lines | Purpose |
|---|---|---|
| `src/modules/modExpedienteHitosHelper.bas` | ~470 | Pure-data helpers (6 Public Functions + private utilities) |
| `src/modules/Test_ExpedienteHitosHelper.bas` | ~150 | TDD atoms (12 Public atoms + RunAll) |

### 9.2 Form .cls changes

| File | Before | After (est.) | Delta |
|---|---|---|---|
| `src/forms/Form_FormExpedienteHitos.cls` | 257 | ~150 | -42% |

**Public surface preserved** (per e2e-methodology §6):
- `EstablecerDatos(Optional ByRef p_Error As String) As String` — signature unchanged.
- `RellenarListas(Optional ByRef p_Error As String) As String` — signature unchanged.
- `RellenarListaHitos(Optional ByRef p_Error As String) As String` — signature unchanged.

**Removed**:
- `On Error GoTo errores + Err.Raise 1000` in handlers.
- `VBA.DoEvents / DoCmd.Hourglass True/False` blocks from helpers.
- `pregunta = MsgBox(...)` bug.