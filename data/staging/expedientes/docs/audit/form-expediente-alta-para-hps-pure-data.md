# Audit — Form_FormExpedienteAltaParaHPS (Phase 3.2 / PR-8)

> **Change**: `forms-thin-coverage` (Phase 3.2 — PR-8, batch of 6 forms)
> **Form**: `Form_FormExpedienteAltaParaHPS` (alta específica HPS)
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
| Líneas .cls | 461 |
| Handlers totales | 15 (Form_Load + cmdSalir + 6 botones add/del + 3 AfterUpdate + 3 WithEvents Alta callbacks) |
| Public methods | 5 (`EstablecerDatos`, `EstablecerCombos`, `RellenarListaContratistas`, `RellenarListaSubContratistas`, `RellenarListaLugares`) — **TODAS MANTENER firma** (API pública) |
| Public events | 1 (`AltaExpediente(m_expediente As Expediente)`) |
| MsgBox call sites | 7 (Form_Load error, Comando* buttons, error paths) |
| DoCmd.OpenForm en .cls | 0 |
| DoCmd.Close en .cls | 1 (cmdSalir, Form_Load error) |
| Forms abiertas | 0 propias (delega a `Form_FormSuministrador` y `Form_FormLugarEjecucion` via WithEvents) |
| WithEvents | 4 (`m_FormClasificacion`, `m_ForContratista`, `m_ForSubContratista`, `m_FormLugar`) |

## 3. CONTROLES REFERENCIADOS (Me.X) — verificación contra .form.txt

| Control | Tipo | Usado en |
|---|---|---|
| `IdGradoClasificacion` | ComboBox | EstablecerCombos |
| `Contratistas` | ComboBox | EstablecerCombos, ComandoAltaContratista, Contratistas_AfterUpdate, m_ForContratista_Alta |
| `SubContratistas` | ComboBox | EstablecerCombos, ComandoAltaSubContratista, SubContratistas_AfterUpdate, m_ForSubContratista_Alta |
| `LugaresEjecucion` | ComboBox | EstablecerCombos, ComandoAltaLugarEjecucion, LugaresEjecucion_AfterUpdate |
| `ListaContratistas` | ListBox | RellenarListaContratistas, ComandoEliminarContratista |
| `ListaSubContratistas` | ListBox | RellenarListaSubContratistas, ComandoEliminarSubContratista |
| `ListaLugaresEjecucion` | ListBox | RellenarListaLugares, ComandoEliminarLugarEjecucion |
| `Titulo` | TextBox | SetDTOFromForm |
| `Nemotecnico` | TextBox | SetDTOFromForm |
| `CodExp` | TextBox | SetDTOFromForm |
| `Observaciones` | TextBox | SetDTOFromForm |
| `Caption` | Form property | Form_Load |
| `AllowEdits` | Form property | Form_Load |

## 4. ENTIDAD Y PROPIEDADES REFERENCIADAS

| Entidad | Propiedades usadas |
|---|---|
| `Expediente` | `Titulo`, `Nemotecnico`, `CodExp`, `IdGradoClasificacion`, `Observaciones`, `Ambito`, `HPSAplica`, `POSTAGEDO`, `TIpo` |
| `ExpedienteOperaciones` | `Registrar` |
| `ExpedienteDTO` | `ColArbolSuministradores`, `ColLugaresEjecucion`, `Expediente` (vía global `m_ObjExpedienteDTOActivo`) |
| `ExpedienteSuministrador` | `IDExpedienteSuministrador`, `IDSuministrador`, `IdPadre`, `Tag`, `ContratistaPrincipal`, `SubContratista`, `Descripcon` |
| `Suministrador` | `Nombre`, `IDSuministrador` |
| `LugarEjecucion` | `IDLugarEjecucion`, `LugarEjecucion` |
| `GradoClasificacion` | `IdGradoClasificacion`, `GradoClasificacion` |
| `Entorno` (global `m_ObjEntorno`) | `GradosClasificaciones`, `Suministradores`, `LugaresEjecucion` |

## 5. INTER-FORM CALLSITES

| Target | Mecanismo |
|---|---|
| `Form_FormGradoClasificacion` | `WithEvents m_FormClasificacion` — definido pero NO instanciado en este .cls (lo hace el form padre que abre) |
| `Form_FormSuministrador` (×2: contratista + sub) | `WithEvents m_ForContratista` / `m_ForSubContratista` — Alta callback |
| `Form_FormLugarEjecucion` | `WithEvents m_FormLugar` — Alta callback |
| `Helper_ExpedienteAlta.Ajustar` | Module call (UI layout) |
| `constructor.getSuministrador` | Module call (DAO lookup) |
| `constructor.getLugarEjecucion` | Module call (DAO lookup) |
| `MostrarPopupProgreso` / `CerrarPopupProgreso` | Module call (UI guard) |

## 6. RECOMENDACIÓN DE PATRÓN

**Pattern**: SPECIAL — Alta con DTO en memoria + add/del via botones. Cada botón es un mini-CRUD.

### 6.1 Helpers planeados (8)

| Helper | Signature | Responsabilidad |
|---|---|---|
| `ExpedienteHPS_Form_Load` | `(p_DTO As Object, p_Error As String) As String` | Inicializa `ColArbolSuministradores` si Nothing, llama `Ajustar Me`, set AllowEdits, llama `EstablecerDatos` |
| `ExpedienteHPS_EstablecerCombos_Construir` | `(p_DTO As Object, p_Error As String) As String` | Retorna JSON con `clasificacionesRowSource`, `contratistasRowSource`, `subcontratistasRowSource`, `lugaresRowSource` (cada uno como semicolon-separated). El form hace `cmb.AddItem` por su cuenta (UI wiring) |
| `ExpedienteHPS_RellenarLista_Construir` | `(p_DTO As Object, p_TipoLista As String, p_Error As String) As String` | Para `p_TipoLista` ∈ `{"CONTRATISTAS", "SUBCONTRATISTAS", "LUGARES"}`, retorna JSON con `rowSource` (semicolon-separated). El form hace `lst.AddItem` |
| `ExpedienteHPS_AltaSuministrador` | `(p_DTO As Object, p_IDSuministrador As String, p_Tipo As String, p_Error As String) As String` | Wrapper puro: crea `ExpedienteSuministrador` con `ContratistaPrincipal="Sí"` o `SubContratista="Sí"`, lo añade a `ColArbolSuministradores`. Retorna `added:Boolean, duplicate:Boolean` |
| `ExpedienteHPS_EliminarSuministrador` | `(p_DTO As Object, p_IDSuministrador As String, p_Tipo As String, p_Error As String) As String` | Wrapper puro: busca por `IDSuministrador` + tipo, elimina de `ColArbolSuministradores`. Retorna `removed:Boolean` |
| `ExpedienteHPS_AltaLugar` | `(p_DTO As Object, p_IDLugar As String, p_Error As String) As String` | Añade `LugarEjecucion` a `ColLugaresEjecucion` si no existe. Retorna `added:Boolean, duplicate:Boolean` |
| `ExpedienteHPS_EliminarLugar` | `(p_DTO As Object, p_IDLugar As String, p_Error As String) As String` | Elimina de `ColLugaresEjecucion`. Retorna `removed:Boolean` |
| `ExpedienteHPS_ComandoRegistrar_Click` | `(p_DTO As Object, p_Error As String) As String` | SetDTOFromForm + setea `Ambito="HPS"`, `HPSAplica="Sí"`, `POSTAGEDO="Sí"`, `TIpo="Expediente Sólo para HPS"` + `ExpedienteOperaciones.Registrar` + `RaiseEvent AltaExpediente` |

### 6.2 UI orchestration que se queda en el form

- `RellenarListaContratistas` / `RellenarListaSubContratistas` / `RellenarListaLugares` — iteran el JSON del helper y hacen `lst.AddItem`. MANTENER wrappers públicos (API), pero el contenido es UI wiring.
- `EstablecerDatos` / `EstablecerCombos` — similares, MANTENER firma pública; el helper interno retorna datos puros, el form aplica a controles.
- `cmdSalir_Click` — `DoCmd.Close`.
- 3 `*_AfterUpdate` — llaman al botón alta correspondiente (UI chaining).
- 3 `m_For*_Alta` callbacks — refrescan combos + ejecutan alta.
- 6 botones add/delete — orquestan: extraer ID del combo/list → llamar helper → re-renderizar lista.
- `Form_Load` errores — MsgBox.

### 6.3 Public methods que MANTENER

- `EstablecerDatos` — MANTENER firma
- `EstablecerCombos` — MANTENER firma
- `RellenarListaContratistas` / `RellenarListaSubContratistas` / `RellenarListaLugares` — MANTENER firma (públicas; el form padre o tests pueden llamarlas)

## 7. NOTAS Y RIESGOS

1. **Forma con DTO complejo** (`ColArbolSuministradores` jerárquico). Los helpers deben aceptar el DTO como `Object` (Scripting.Dictionary) — la "interfaz" está implícita en la estructura del Dictionary.
2. **Public Sub `SetDTOFromForm`** — Private; el helper `ExpedienteHPS_ComandoRegistrar_Click` necesita acceso equivalente. Inlining su lógica en el helper está bien.
3. **`m_FormClasificacion`** declarado pero no usado directamente — dejarlo, no es bloqueante.
4. **Anti-pattern check**: el form YA está limpio de `ByRef p_Form`. Solo tiene `Me.X` que son UI wiring legítimo. MsgBox es UI feedback — se queda.
5. **No `DoCmd.OpenForm`** en este form — los WithEvents apuntan a forms que YA están abiertos por el form padre. El handler de Alta callback es pass-through.

## 8. VERIFICATION CHECKLIST (post-extraction)

| Check | Target |
|---|---|
| Helper signatures: ZERO `ByRef p_Form` | OK esperado |
| Helpers: ZERO `DoCmd.OpenForm` / `Forms(...)` | OK esperado |
| Test atoms: ZERO `DoCmd.OpenForm` / `Forms(...)` / `Screen.ActiveForm` | OK esperado |
| Atoms estimated | ~40-55 (8 helpers × 5-7 atoms promedio) |
| Pattern | SPECIAL — Alta con DTO en memoria |
