# Audit — Form_FormExpedienteAlta (Phase 3.2 / PR-8)

> **Change**: `forms-thin-coverage` (Phase 3.2 — PR-8, batch of 6 forms)
> **Form**: `Form_FormExpedienteAlta` (alta de nuevo expediente)
> **Date**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight)

## 1. STATUS

| Fase | Estado |
|---|---|
| Audit | **COMPLETE** — helpers not yet extracted |
| Helper extraction | **PENDING** (6-8 helpers planned) |
| Binario sync | PENDING |

## 2. FORM METADATA

| Atributo | Valor |
|---|---|
| Líneas .cls | 675 |
| Handlers totales | 23 (Form_Load + Form_Unload + 12 botones/sub + 4 WithEvents callbacks + 4 sub helpers Privates) |
| Public methods | 2 (`EstablecerDatos(Optional ByRef p_Error As String) As String`, `EstablecerDatosConTipo` Private) — `EstablecerDatos` MANTENER |
| Public events | 1 (`Alta(m_ObjExpediente As Expediente)`) |
| MsgBox call sites | 14 (casi todos los handlers de botón + Form_Load + Form_Unload + WithEvents callbacks + Ordinal_AfterUpdate) |
| DoCmd.OpenForm en .cls | 4 (`FormGradosClasificacionGestion`, `FormOficinasProgramaGestion`, `FormEjercitosGestion`, `FormOrganoContratacionGestion`) |
| DoCmd.Close en .cls | 5 (cmdSalir, ComandoRegistrar, ComandoElegitClasificacion, ComandoElegirOficinaPrograma, ComandoElegirEjercito, ComandoElegirOC) |
| Forms abiertas | 4 gestión (OC, Clasif, Ejercito, OfProg) |
| WithEvents | 4 (`m_FormOC`, `m_FormClasificacion`, `m_FormEjercito`, `m_FormOficinaPrograma`) |

## 3. CONTROLES REFERENCIADOS (Me.X) — verificación contra .form.txt

| Control | Tipo | Usado en |
|---|---|---|
| `IdGradoClasificacion` | ComboBox | ComandoLimpiarClasificacion, m_FormClasificacion_Seleccionar |
| `IDOrganoContratacion` | ComboBox | ComandoLimpiarOC, m_FormOC_Seleccionar |
| `IDOficinaPrograma` | ComboBox | ComandoLimpiarOficinaPrograma, m_FormOficinaPrograma_Seleccionar |
| `ComandoVerSharepoint` | CommandButton | AccesoSharepoint_Exit |
| `AccesoSharepoint` | TextBox | AccesoSharepoint_Exit, ComandoVerSharepoint |
| `Ambito` | ComboBox | Ambito_AfterUpdate |
| `HPSAplica` | (calculado) | Ambito_AfterUpdate |
| `Nemotecnico` | TextBox | ComandoRellenarNemotecnico, EstablecerDatos |
| `lblTitulo` | Label | Form_Load, EstablecerDatos |
| `lblNumero` | Label | EstablecerDatos |
| `Ordinal` | TextBox | EstablecerDatos, Ordinal_AfterUpdate |
| `ComandoRellenarNemotecnico` | CommandButton | EstablecerDatos |
| `TIpo` | ComboBox | Form_Load, EstablecerDatos |
| `AllowEdits` | Form property | Form_Load |

## 4. ENTIDAD Y PROPIEDADES REFERENCIADAS

| Entidad | Propiedades usadas |
|---|---|
| `Expediente` | `TIpo`, `Nemotecnico`, `IDExpedientePadre`, `TIPOEnum`, `ExpedientePadre`, `Ejercito`, `EsAM`, `EsLote`, `EsExpediente`, `IDExpediente`, `Ordinal`, `OrdinalCalculado` |
| `ExpedienteOperaciones` | `RegistrarAlta` |
| `ExpedienteDTO` | `.Expediente` (vía global `m_ObjExpedienteDTOActivo`) |
| `ExpedienteModificado` (vía `ExpedienteOrdinalUsado`) | `CodExp`, `Nemotecnico`, `Titulo` |
| `GradoClasificacion` | `GradoClasificacion` (texto) |
| `OrganoContratacion` | `OrganoContratacion` (texto) |
| `OficinaPrograma` | `OficinaPrograma` (texto) |
| `Ejercito` | `Ejercito` (texto) |

## 5. INTER-FORM CALLSITES

| Target | Mecanismo |
|---|---|
| `FormularioAbierto("FormGradosClasificacionGestion")` | `DoCmd.Close` + `DoCmd.OpenForm` |
| `FormularioAbierto("FormOficinasProgramaGestion")` | `DoCmd.Close` + `DoCmd.OpenForm` |
| `FormularioAbierto("FormEjercitosGestion")` | `DoCmd.Close` + `DoCmd.OpenForm` |
| `FormularioAbierto("FormOrganoContratacionGestion")` | `DoCmd.Close` + `DoCmd.OpenForm` |
| `FormularioAbierto("FormExpedienteAltaTipo")` | notificación de Alta via `m_FormTipo.ReplicarAltaEvento` |
| `Forms("FormGradosClasificacionGestion")` / `"FormOficinasProgramaGestion"` / `"FormEjercitosGestion"` / `"FormOrganoContratacionGestion"` | set `WithEvents` refs |
| `Helper_ExpedienteAlta.EstablecerCombos` | Module call |
| `Helper_ExpedienteAlta.SetDatosDeExpedienteTipo` | Module call |
| `Helper_ExpedienteAlta.HeredarDatosDePadre` | Module call |
| `Helper_ExpedienteAlta.SetDTOFromGeneral` | Module call |
| `Helper_ExpedienteEdicion.Ajustar` | Module call |
| `Helper_ExpedienteEdicion.getSnapshot` | Module call |
| `Helper_Expedicion.CalcularOrdinalSiguiente` | Module call |
| `Helper_Expedicion.CalcularNemotecnicoCalculado` | Module call |
| `Helper_ExpedienteEliminacion.ExpedienteOrdinalUsado` | Module call |
| `AntiSpam*` | Module call (UI guard) |
| `constructor.getExpediente` | Module call |

## 6. RECOMENDACIÓN DE PATRÓN

**Pattern**: ALTA/EDICIÓN + inter-form WithEvents. Más complejo que el patrón estándar por los 4 WithEvents.

### 6.1 Helpers planeados (8)

| Helper | Signature | Responsabilidad |
|---|---|---|
| `ExpedienteAlta_Form_Load` | `(p_OpenArgs As String, p_DTO As Object, p_Error As String) As String` | Parsea `IDExpedientePadre\|TipoEnum`, llama `SetDatosDeExpedienteTipo` + `HeredarDatosDePadre`, fija label título, allowEdits, snapshot inicial |
| `ExpedienteAlta_EstablecerDatos` | `(p_DTO As Object, p_TipoEnum As Long, p_IDExpedientePadre As String, p_Error As String) As String` | Decide visibilidad de ordinal/lblNumero/ComandoRellenarNemotecnico según tipo; setea `Me.TIpo`, `Me.Nemotecnico`; calcula `OrdinalCalculado` |
| `ExpedienteAlta_EstablecerDatosConTipo` | `(p_DTO As Object, p_TipoTexto As String, p_Error As String) As String` | Traduce `TIpo` (texto) a `EsAM`/`EsLote`/`EsExpediente` (Sí/No) en el DTO |
| `ExpedienteAlta_ComandoRegistrar_Click` | `(p_DTO As Object, p_Error As String) As String` | `SetDTOFromGeneral` + `ExpedienteOperaciones.RegistrarAlta` + `RaiseEvent Alta` + retornar `notifyTipo:Boolean` (indica si notificar AltaTipo) |
| `ExpedienteAlta_Ambito_AfterUpdate` | `(p_AmbitoActual As String, p_HPSAplicaActual As String, p_Error As String) As String` | Si Ambito="HPS", retorna `setHPSAplica: "Sí"` |
| `ExpedienteAlta_Ordinal_AfterUpdate` | `(p_Ordinal As String, p_IDExpedientePadre As String, p_PromptResult As Long, p_Error As String) As String` | Llama `ExpedienteOrdinalUsado`, si hay conflicto → MsgBox via p_PromptResult. Retorna `conflicto:{CodExp/Nemotecnico/Titulo}` |
| `ExpedienteAlta_ComandoRellenarNemotecnico_Click` | `(p_DTO As Object, p_Error As String) As String` | `CalcularOrdinalSiguiente` + `CalcularNemotecnicoCalculado`, retorna `nemotecnico:String` |
| `ExpedienteAlta_NotificarAltaTipo` | `(p_Expediente As Object, p_Error As String) As String` | Wrapper puro-datos: indica si `FormularioAbierto("FormExpedienteAltaTipo")` retornaría `True` (en test, esto se inyecta via stub) + ejecuta `ReplicarAltaEvento` |

### 6.2 UI orchestration que se queda en el form

- 4 `ComandoElegir*_Click` — `DoCmd.Close acForm "FormX"` + `DoCmd.OpenForm "FormX", OpenArgs:="PARA SELECCIONAR"` + set `WithEvents` ref. (Regla #1: navegación UI vive en el form)
- 4 `m_Form*_Seleccionar` callbacks — asignan `Me.IdX = obj.X` (puro UI wiring)
- 3 `ComandoLimpiar*_Click` — `Me.IdX = Null` (puro UI)
- `ComandoVerSharepoint_Click` — `Ejecutar 1, "open", ...` (file open, UI)
- `cmdSalir_Click` — `DoCmd.Close`
- `Form_Unload` — pregunta si hay cambios sin guardar (vía snapshot)
- `AccesoSharepoint_Exit` — enable/disable `ComandoVerSharepoint` (UI wiring puro)

### 6.3 Public methods que MANTENER

- `EstablecerDatos(Optional ByRef p_Error As String) As String` — **MANTENER firma**; internamente llama al helper puro-datos `ExpedienteAlta_EstablecerDatos`

### 6.4 WithEvents handlers — qué hacer

Los 4 WithEvents callbacks (`m_FormOC_Seleccionar`, etc.) NO pueden eliminarse — son API del form. PERO su lógica interna (setear `Me.IdX`) se queda. No tienen suficiente lógica para extraer.

## 7. NOTAS Y RIESGOS

1. **Forma con MÁS handlers** del lote (23). El test surface es el más grande.
2. **Inter-form WithEvents** añade complejidad — los helpers NO pueden testear la apertura del form de selección (eso es UI), pero SÍ pueden testear la lógica de `Ordinal_AfterUpdate`, `ComandoRegistrar`, `ComandoRellenarNemotecnico`, etc.
3. **`FormularioAbierto` y `Forms(...)`** se usan en handlers de selección. Esos se quedan en el form (UI orchestration). NO van a helpers.
4. **Public event `Alta`** — disparado por `RaiseEvent Alta(m_ObjExpedienteDTOActivo.Expediente)`. API consumida por `FormExpedienteAltaTipo`. MANTENER.
5. **Snapshot en Form_Unload** — comparar `m_SnapshotInicial` vs `getSnapshot(Me, ...)`. La comparación es UI (captura estado del form), NO se mete a helper.
6. **`FormExpedienteAltaTipo` (otro form del lote)** referencia `Form_FormExpedienteAlta` via `Forms("FormExpedienteAlta")` en `ComandoAlta_Click`. Esto es acoplamiento entre forms del mismo lote — el refactor debe preservar este contrato.

## 8. VERIFICATION CHECKLIST (post-extraction)

| Check | Target |
|---|---|
| Helper signatures: ZERO `ByRef p_Form` | OK esperado |
| Helpers: ZERO `DoCmd.OpenForm` / `Forms(...)` | OK esperado (excepto `NotificarAltaTipo` que es wrapper con flag) |
| Test atoms: ZERO `DoCmd.OpenForm` / `Forms(...)` / `Screen.ActiveForm` | OK esperado |
| Atoms estimated | ~40-55 (8 helpers × 5-7 atoms promedio) |
| Pattern | SPECIAL — Alta + inter-form |
