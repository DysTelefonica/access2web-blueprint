# Audit — Form_FormExpedienteCambioTipo (Phase 3.2 / PR-8)

> **Change**: `forms-thin-coverage` (Phase 3.2 — PR-8, batch of 6 forms)
> **Form**: `Form_FormExpedienteCambioTipo` (cambio de tipo de expediente)
> **Date**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight)

## 1. STATUS

| Fase | Estado |
|---|---|
| Audit | **COMPLETE** |
| Helper extraction | **PENDING** (3-4 helpers planned) |
| Binario sync | PENDING |

## 2. FORM METADATA

| Atributo | Valor |
|---|---|
| Líneas .cls | 340 |
| Handlers totales | 8 (Form_Open + cmdSalir + ComandoAyuda + ComandoBuscar + ComandoRegistrar + Tipo_AfterUpdate + m_Form_Seleccionado + EstablecerDatos) |
| Public methods | 1 (`EstablecerDatos(Optional ByRef p_Error As String) As String`) — **MANTENER firma** |
| Public events | 1 (`CambioRealizado()`) |
| MsgBox call sites | 6 (ComandoAyuda, ComandoBuscar, ComandoRegistrar, Form_Open, m_Form_Seleccionado, Tipo_AfterUpdate) |
| DoCmd.OpenForm en .cls | 1 (`FormExpedientesParaCambioTipo`) |
| DoCmd.Close en .cls | 2 (cmdSalir, ComandoBuscar) |
| Forms abiertas | 1 (`FormExpedientesParaCambioTipo`) |
| WithEvents | 1 (`m_form` → Form_FormExpedientesParaCambioTipo) |

## 3. CONTROLES REFERENCIADOS (Me.X)

| Control | Tipo | Usado en |
|---|---|---|
| `TIpo` | ComboBox | EstablecerDatos, ComandoRegistrar, Tipo_AfterUpdate |
| `TIPOACTUAL` | (Label o TextBox) | EstablecerDatos |
| `IDExpedientePadre` | ComboBox | EstablecerDatos, ComandoRegistrar, Tipo_AfterUpdate, m_Form_Seleccionado |
| `Caption` | Form property | Form_Open |

## 4. ENTIDAD Y PROPIEDADES REFERENCIADAS

| Entidad | Propiedades usadas |
|---|---|
| `Expediente` (global `m_ObjExpedienteActivo`) | `TipoCalculado`, `TipoCalculadoTexto`, `IDExpedientePadre`, `EsAM`, `EsLote`, `EsExpediente`, `EsBasado` |
| `ExpedienteOperaciones` | `RegistrarCambioTipo` |
| `Expediente` (en `m_ExpPadre`) | `IDExpediente`, `EsAM`, `EsLote` |

## 5. INTER-FORM CALLSITES

| Target | Mecanismo |
|---|---|
| `FormularioAbierto("FormExpedientesParaCambioTipo")` | `DoCmd.Close` + `DoCmd.OpenForm` |
| `Forms("FormExpedientesParaCambioTipo")` | set `WithEvents m_form` |
| `Helper_ExpedienteEdicion.Ajustar` | Module call (UI layout) |
| `constructor.getExpediente` | Module call (DAO lookup) |
| `MostrarPopupProgreso` / `CerrarPopupProgreso` | Module call (UI guard) |
| `AbrirAyuda` | Module call |

## 6. RECOMENDACIÓN DE PATRÓN

**Pattern**: SPECIAL — 5-branch dispatch por tipo. NO encaja en Gestion/Alta clásico.

### 6.1 Helpers planeados (4)

| Helper | Signature | Responsabilidad |
|---|---|---|
| `ExpedienteCambioTipo_EstablecerDatos` | `(p_Expediente As Object, p_Error As String) As String` | Retorna JSON: `{tipoActual:String, opciones: ["Convertir a Lote", ...]}` según `TipoCalculado`. El form llena el combo (UI wiring). |
| `ExpedienteCambioTipo_Tipo_AfterUpdate` | `(p_TipoElegido As String, p_Error As String) As String` | Retorna JSON: `{idPadreLocked:Boolean, idPadreEnabled:Boolean}` según la rama del tipo elegido. El form aplica `Me.IDExpedientePadre.Locked/Enabled` (UI). |
| `ExpedienteCambioTipo_ComandoRegistrar` | `(p_Expediente As Object, p_TipoElegido As String, p_IDExpedientePadre As String, p_PromptResult As Long, p_Error As String) As String` | 5-branch dispatch: valida IDExpedientePadre, setea `EsAM/EsLote/EsExpediente/EsBasado/IDExpedientePadre` en el DTO, llama `ExpedienteOperaciones.RegistrarCambioTipo`. Retorna `{registrado:Boolean, validation:String}`. |
| `ExpedienteCambioTipo_ValidarPadre` | `(p_TipoElegido As String, p_IDExpedientePadre As String, p_Error As String) As String` | Helper puro: si requiere padre, valida que existe y del tipo correcto. Retorna `{valido:Boolean, error:String}`. |

### 6.2 UI orchestration que se queda en el form

- `cmdSalir_Click` — `DoCmd.Close`.
- `ComandoAyuda_Click` — `AbrirAyuda`.
- `ComandoBuscar_Click` — `DoCmd.OpenForm "FormExpedientesParaCambioTipo"`. (regla #1)
- `Form_Open` — `Ajustar Me` + `EstablecerDatos`.
- `m_Form_Seleccionado` — `Me.IDExpedientePadre = m_expediente.IDExpediente` (UI wiring).
- `EstablecerDatos` — MANTENER firma; internamente llama al helper.

### 6.3 Lógica de 5 ramas (extracto del .cls original)

```
"Convertir a Acuerdo Marco"      → IDExpedientePadre="", EsAM="Sí", EsLote="No", EsExpediente="No", EsBasado="No"
"Convertir a Lote"               → requiere IDExpedientePadre de tipo AM, EsAM="No", EsLote="Sí"
"Convertir a Basado"             → requiere IDExpedientePadre de tipo AM o Lote, EsBasado="Sí"
"Convertir a Expediente Independiente" → IDExpedientePadre="" (validación), EsExpediente="Sí"
"Convertir a Lote de Acuerdo Marco" → requiere IDExpedientePadre, EsLote="Sí"
```

Esta lógica ES business logic y DEBE ir al helper `ExpedienteCambioTipo_ComandoRegistrar`.

## 7. NOTAS Y RIESGOS

1. **5 ramas en ComandoRegistrar** — el helper debe cubrir TODAS con tests (5 happy paths + 4 sad paths para validaciones).
2. **Validación de padre** — `m_ExpPadre.EsAM <> "Sí"` etc. Es lógica de negocio. Va al helper `ExpedienteCambioTipo_ValidarPadre` para tests granulares.
3. **`Tipo_AfterUpdate`** — UI state puro (Locked/Enabled del combo padre). El helper retorna qué hacer; el form aplica.
4. **`m_form` (WithEvents)** — `m_form_Seleccionado(m_expediente As Expediente)` es pass-through. NO requiere helper.

## 8. VERIFICATION CHECKLIST (post-extraction)

| Check | Target |
|---|---|
| Helper signatures: ZERO `ByRef p_Form` | OK esperado |
| Helpers: ZERO `DoCmd.OpenForm` / `Forms(...)` | OK esperado |
| Test atoms: ZERO `DoCmd.OpenForm` / `Forms(...)` / `Screen.ActiveForm` | OK esperado |
| Atoms estimated | ~20-30 (4 helpers × 5-8 atoms promedio,especialmente el dispatch de 5 ramas) |
| Pattern | SPECIAL — 5-branch dispatch |
