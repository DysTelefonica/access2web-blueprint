# Audit — Form_FormExpedienteAltaTipo (Phase 3.2 / PR-8)

> **Change**: `forms-thin-coverage` (Phase 3.2 — PR-8, batch of 6 forms)
> **Form**: `Form_FormExpedienteAltaTipo` (selector de tipo de alta)
> **Date**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight)

## 1. STATUS

| Fase | Estado |
|---|---|
| Audit | **COMPLETE** |
| Helper extraction | **NONE — KEEP AS-IS** |
| Binario sync | N/A |
| Pattern | **Verified-static** — pure UI navigation hub |

## 2. FORM METADATA

| Atributo | Valor |
|---|---|
| Líneas .cls | 280 |
| Handlers totales | 8 (Form_Open + cmdSalir + ComandoAlta + ComandoAyuda + ListaTipoAlta_Click + ListaTipoAlta_DblClick + 2 WithEvents) |
| Public methods | 1 (`ReplicarAltaEvento(p_Exp As Expediente)`) — **MANTENER** (lo llama Form_FormExpedienteAlta) |
| Public events | 1 (`Alta(p_Exp As Expediente)`) |
| MsgBox call sites | 5 (ComandoAlta, ComandoAyuda, Form_Open, ListaTipoAlta_Click — todos error paths) |
| DoCmd.OpenForm en .cls | 2 (`FormExpedienteAltaParaHPS`, `FormExpedienteAlta`) |
| DoCmd.Close en .cls | 3 (cmdSalir, ComandoAlta × 2 — cierra este form y abre el siguiente) |
| Forms abiertas | 2 (`FormExpedienteAltaParaHPS` para HPS, `FormExpedienteAlta` para el resto) |
| WithEvents | 2 (`m_FormExpediente_ParaHPS`, `m_FormExpedienteAlta`) |

## 3. CONTROLES REFERENCIADOS (Me.X)

| Control | Tipo | Usado en |
|---|---|---|
| `ComandoAlta` | CommandButton | Form_Open, ListaTipoAlta_Click, ListaTipoAlta_DblClick |
| `Caption` | Form property | Form_Open |
| `AllowEdits` | Form property | Form_Open |
| `OpenArgs` | Form property | Form_Open |
| `ListaTipoAlta` | ListBox | EstablecerLista, ComandoAlta_Click, ListaTipoAlta_Click |

## 4. ENTIDAD Y PROPIEDADES REFERENCIADAS

| Entidad | Propiedades usadas |
|---|---|
| `Expediente` (global `m_ObjExpedienteActivo`) | `IDExpediente` |
| `ExpedienteDTO` (global `m_ObjExpedienteDTOActivo`) | `Expediente` |
| `Entorno` (global `m_ObjEntorno`) | `ColTiposExpedientesNoDerivados`, `ColTiposExpedientes` |

## 5. INTER-FORM CALLSITES

| Target | Mecanismo |
|---|---|
| `FormularioAbierto("FormExpedienteAltaParaHPS")` | `DoCmd.Close` + `DoCmd.OpenForm` |
| `FormularioAbierto("FormExpedienteAlta")` | `DoCmd.Close` + `DoCmd.OpenForm` |
| `Forms("FormExpedienteAlta")` | set `WithEvents` ref |
| `Helper_ExpedienteAlta.SetDatosDeExpedienteTipo` | Module call |
| `Helper_ExpedienteEdicion.Ajustar` | Module call (UI layout) |
| `AbrirAyuda` | Module call |

## 6. RECOMENDACIÓN: KEEP AS-IS (Verified-static)

### 6.1 Por qué NO extraer helpers

Este form es un **navigation hub puro**:

1. **`ComandoAlta_Click`** — único path con lógica condicional. Lee el tipo seleccionado del listbox (column 0) y dispatch:
   - `EXPHPS` → cierra `FormExpedienteAltaParaHPS` si está abierto, abre `FormExpedienteAltaParaHPS`
   - resto → cierra `FormExpedienteAlta` si está abierto, abre `FormExpedienteAlta` con `OpenArgs:="IDPadre|TipoEnum"`
   - Esto es **UI orchestration pura** (per e2e-methodology rule #1: navigation stays in form).

2. **`EstablecerLista`** — rellena `ListaTipoAlta` con `m_ObjEntorno.ColTiposExpedientesNoDerivados`. Es UI wiring (lst.AddItem). No hay cálculo de negocio.

3. **`ListaTipoAlta_Click` / `_DblClick`** — enable/disable `ComandoAlta`. Puro UI state.

4. **`Form_Open`** — `Ajustar Me`, parsea `OpenArgs`, `EstablecerLista`, `Ajustar visibilidad de ComandoAlta` según `EsAdministrador`. Puro UI.

5. **`m_FormExpedienteAlta_Alta`** — re-raise event `Alta(m_ObjExpediente)`. Passthrough.

6. **`ReplicarAltaEvento(p_Exp)`** — re-raise `Alta(p_Exp)`. Passthrough. API pública consumida por `Form_FormExpedienteAlta`.

### 6.2 Verificación de la regla de decisión (per task spec)

> "If pure UI navigation (no business logic): KEEP AS-IS, document as Verified-static"

✅ **Aplica**:
- Sin business logic (no validaciones, no cálculos de datos)
- Sin lectura/escritura de DTO que no sea re-raise event
- Sin DAO calls
- Toda la lógica es navegación (`DoCmd.OpenForm`/`DoCmd.Close`) + relleno de listbox + enable/disable de botón

### 6.3 Lo que NO se hace

- NO se crea `modExpedienteAltaTipoHelper.bas`
- NO se crea `Test_ExpedienteAltaTipoHelper.bas`
- NO se reduce el .cls — se queda tal cual
- NO se añade nada al PR desde este form

## 7. CONTRATO PRESERVADO

Este form es consumido por:
- `Form_FormExpedientesParaCambioTipo` (vía OpenArgs)
- `Form_FormExpedienteAlta` (vía `m_FormExpedienteAlta.ReplicarAltaEvento`)

El refactor de los otros 5 forms del lote debe **preservar** este contrato:
- `Form_FormExpedienteAlta.ComandoRegistrar_Click` llama a `m_FormTipo.ReplicarAltaEvento` — **PRESERVAR**
- El evento `Alta(m_ObjExpediente As Expediente)` público — **PRESERVAR**

## 8. VERIFICATION CHECKLIST

| Check | Target |
|---|---|
| Helpers extraídos | 0 (correcto) |
| Test atoms | 0 (correcto) |
| Anti-pattern check | N/A (no se modifica) |
| Source binary sync | N/A (no se modifica) |
| Pattern | **Verified-static — navigation hub** |
