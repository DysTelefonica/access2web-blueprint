# Audit — Form_FormExpediente (Phase 3.2 / PR-8)

> **Change**: `forms-thin-coverage` (Phase 3.2 — PR-8, batch of 6 forms)
> **Form**: `Form_FormExpediente` (main expediente view/edit, tab manager)
> **Date**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight)

## 1. STATUS

| Fase | Estado |
|---|---|
| Audit | **COMPLETE** — helpers not yet extracted |
| Helper extraction | **PENDING** (4-5 helpers planned) |
| Binario sync | PENDING |

## 2. FORM METADATA

| Atributo | Valor |
|---|---|
| Líneas .cls | 402 |
| Handlers totales | 14 (Form_Load + 7 tab click + 4 botones + GuardarPestanaActivaSiEsGeneralOFechas + GuardarPestanaActivaAntesDeNavegar) |
| Public methods | 1 (`EstablecerDatos(Optional ByRef p_Error As String) As String`) — **MANTENER** (lo llaman otros forms vía RaiseEvent / directo) |
| Public events | 2 (`AltaExpediente`, `ExpedienteEditado`) |
| MsgBox call sites | 3 (GuardarPestanaActivaAntesDeNavegar, ComandoActualizarCompleto, ComandoAyuda, ComandoRegistrar, Form_Load — error handling) |
| DoCmd.OpenForm en .cls | 0 (navegación se hace via subforms ya cargados) |
| DoCmd.Close en .cls | 1 (cmdSalir, Form_Load errores) |
| Forms abiertas | 0 propias (delega a FormExpedienteGeneral/Fechas/etc. como subforms dentro de FrmDetalle) |
| WithEvents | 0 |

## 3. CONTROLES REFERENCIADOS (Me.X) — verificación contra .form.txt

| Control | Tipo | Usado en | Validado en .form.txt |
|---|---|---|---|
| `InsideHeight` | Form property | Form_Load | OK (property) |
| `InsideWidth` | Form property | Form_Load | OK (property) |
| `ComandoRegistrar` | CommandButton | Form_Load | OK (existe en .form.txt) |
| `AllowEdits` | Form property | Form_Load, cmdSalir | OK (property) |
| `Caption` | Form property | Form_Load, Form_Load | OK (property) |
| `lblTitulo` | Label | EstablecerDatos | OK (existe en .form.txt NameMap) |
| `lblUltimaModificacion` | Label | EstablecerDatos | OK |
| `ComandoActualizarCompleto` | CommandButton | EstablecerDatos | OK |
| `OpenArgs` | Form property | Form_Load | OK (property) |
| `FrmDetalle` | Subform | Form_Load (comentado) | OK (existe en .form.txt — NavigationControl NavigationSubform) |

## 4. ENTIDAD Y PROPIEDADES REFERENCIADAS

| Entidad | Propiedades usadas | Helper existente que las maneja |
|---|---|---|
| `Expediente` | `TIpo`, `Derivable`, `IDExpediente`, `TituloFormulario`, `UltimaModificacionTexto`, `TipoCalculado`, `Nemotecnico` | `Helper_ExpedienteEdicion` (save), `Helper_ExpedienteAlta` (registrar) |
| `ExpedienteOperaciones` | `RegistrarExpEntidades` | directo |
| `ExpedienteDTO` | `.Expediente` (acceso global `m_ObjExpedienteDTOActivo`) | n/a (es DTO) |

## 5. INTER-FORM CALLSITES

| Target | Mecanismo | Acción |
|---|---|---|
| `Helper_ExpedienteEdicion.GuardarPestanaActiva` | Module call | Save active tab sub-form |
| `Helper_ExpedienteEdicion.HayCambiosFormularioPestana` | Module call | Detect dirty state |
| `Helper_ExpedienteEdicion.DebeGuardarPestanaGeneralFechas` | Module call | Decide if save is required |
| `Helper_ExpedienteEdicion.getFormPestanaActiva` | Module call | Get active subform |
| `Helper_ExpedienteEdicion.NombrePestanaDesdeSourceObject` | Module call | Tab name resolution |
| `Helper_ExpedienteEdicion.EsPestanaGeneralOFechas` | Module call | Branch logic |
| `Helper_ExpedienteEdicion.ActualizarSnapshotFormularioPestana` | Module call | Post-save state |
| `Helper_ExpedienteEdicion.NavegarAPestanaSiSolicitada` | Module call | Tab navigation |
| `Helper_ExpedienteEdicion.EstablecerDTODeExpediente` | Module call | DTO init |
| `Helper_ExpedienteEdicion.SetFormGeneralFromExpediente` | Module call | Initial load |
| `constructor.getExpediente` | Module call | DAO fetch |
| `BackendResolver.LeeConfiguracionLocal` | Module call | Backend init |
| `AntiSpamIsOperationInProgress` / `AntiSpamEnterOperation` / `AntiSpamExitOperation` | Module call | Anti-double-click guard |

## 6. RECOMENDACIÓN DE PATRÓN

**Pattern**: SPECIAL — Tab manager con save-delegation. NO encaja en Gestion/Alta clásico.

### 6.1 Helpers planeados (5)

| Helper | Signature (target) | Responsabilidad |
|---|---|---|
| `Expediente_Form_Load` | `(p_OpenArgs As String, p_BackendActivo As String, p_EsAdmin As Boolean, p_HasExpediente As Boolean, p_Error As String) As String` | Inicializa caption, visibilidad de botones según OpenArgs / admin / expediente, dispara SetFormGeneral si OpenArgs con ID |
| `Expediente_EstablecerDatos` | `(p_Expediente As Object, p_Error As String) As String` | Pone lblTitulo, lblUltimaModificacion, visibilidad de ComandoActualizarCompleto según IDExpediente |
| `Expediente_Tab_Seleccionar_Guardar` | `(p_NombrePestana As String, p_AbiertoParaEditar As Boolean, p_AllowEdits As Boolean, p_Error As String) As String` | Decide si debe guardar la pestaña activa antes de cambiar (solo General/Fechas) |
| `Expediente_ComandoRegistrar_Click` | `(p_Expediente As Object, p_Error As String) As String` | Save completo via GuardarPestanaActivaSiEsGeneralOFechas con popup |
| `Expediente_ComandoActualizarCompleto_Click` | `(p_Expediente As Object, p_Error As String) As String` | Llama `m_ExpOp.RegistrarExpEntidades` para refrescar entidades desde DAO |

### 6.2 UI orchestration que se queda en el form (per rule #1)

- `cmdSalir_Click` — guarda si es editable + `DoCmd.Close acForm, Me.Name, acSaveNo`
- 7 `tabX_Click` — llaman `Expediente_Tab_Seleccionar_Guardar`
- `ComandoAyuda_Click` — `AbrirAyuda m_Error` (no necesita helper)
- `Form_Load` — orquesta: `BackendResolver.LeeConfiguracionLocal`, `Helper_ExpedienteEdicion.EstablecerDatos`, opcional `Helper_ExpedienteEdicion.EstablecerDTODeExpediente` + `SetFormGeneralFromExpediente` + `NavegarAPestanaSiSolicitada`

### 6.3 Public methods que MANTENER

- `EstablecerDatos(Optional ByRef p_Error As String) As String` — llamado desde otros forms (Form_FormExpedientesParaCambioTipo via sub-form, sub-forms General/Fechas/etc.). Cambia firma para que el helper interno sea puro-datos, pero el wrapper público mantiene la firma por compatibilidad.

## 7. NOTAS Y RIESGOS

1. **No `ByRef p_Form`** ya está respetado en este form — el refactor solo añade helpers puros-datos que sustituyen llamadas inline a `Helper_ExpedienteEdicion.*` y `m_ExpOp.Registrar*`.
2. **AntiSpam wrappers** se quedan en el form (son UI guard, no lógica de negocio).
3. **Eventos públicos** (`AltaExpediente`, `ExpedienteEditado`) son API consumida por otros forms — MANTENER.
4. **MsgBox en errores** — el helper retorna error via `p_Error` (Telefónica D&S convention), el form se encarga del MsgBox (UI orchestration).
5. **`getSnapshot`, `m_SnapshotInicial`** — el form_Alta tiene su propio snapshot; este form NO, solo verifica `AbiertoParaEditar` + `AllowEdits` en cmdSalir.

## 8. VERIFICATION CHECKLIST (post-extraction)

| Check | Target |
|---|---|
| Helper signatures: ZERO `ByRef p_Form` | OK esperado |
| Helpers: ZERO `DoCmd.OpenForm` / `Forms(...)` / `Application.Echo` | OK esperado |
| Test atoms: ZERO `DoCmd.OpenForm` / `Forms(...)` / `Screen.ActiveForm` | OK esperado |
| Atoms estimated | ~25-35 (5 helpers × 5-7 atoms promedio) |
| Pattern | NEW — "Tab manager / View-only" |
