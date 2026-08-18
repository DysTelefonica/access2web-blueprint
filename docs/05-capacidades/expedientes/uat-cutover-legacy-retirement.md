[← Back to capabilities-index](../capabilities-index.md) · [← Back to DOCS](../../../DOCS.md)

# Expedientes · uat-cutover-legacy-retirement (D54 / EXP-CAP-057..063)

Resumen a nivel repo del slice D54 del task plan. Fuente de verdad: [`openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md`](../../../../openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md).

## Capacidades

| ID | Título | Sustitución | Puerta de retirada |
|---|---|---|---|
| [EXP-CAP-057](../../../../openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md#requirement-exp-cap-057--retirada-win32-red) | Retirada Win32 red | Llamadas Win32 al filesystem/red → `DocumentStoragePort` + Postgres | Cobertura de tests ≥ 95 % + 30 días sin invocación legacy |
| [EXP-CAP-058](../../../../openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md#requirement-exp-cap-058--retirada-procesos-win32) | Retirada procesos Win32 | `Shell()` y `CreateProcess` → `CommandExecutorPort` async | Cobertura de tests ≥ 95 % + 30 días sin invocación legacy |
| [EXP-CAP-059](../../../../openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md#requirement-exp-cap-059--retirada-oleactivex) | Retirada OLE/ActiveX | Controles OLE embebidos (`MSComctlLib.TreeView`, `MSComctlLib.ImageList`) → componentes web (HTMX + Alpine.js) | Reemplazo verificado en UAT + 30 días sin uso legacy |
| [EXP-CAP-060](../../../../openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md#requirement-exp-cap-060--retirada-globals) | Retirada globals | `Variables Globales.bas` con flags temporales → `Configuration` tipada + `ContextVar` | Tests `Test_NoMutableGlobals` verdes + 30 días sin `SetVariableGlobales*` |
| [EXP-CAP-061](../../../../openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md#requirement-exp-cap-061--retirada-selector-backend) | Retirada selector backend | `TbConfiguracionBackends` con test/modo en frontend → `BackendSelectorPort` en backend | Cobertura de tests ≥ 95 % + 30 días sin acceso directo al `.accdb` |
| [EXP-CAP-062](../../../../openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md#requirement-exp-cap-062--retirada-popups) | Retirada popups | `MsgBox`, `DoCmd.OpenForm` con popup → notificaciones in-app (HTMX + toast) | Cobertura de tests ≥ 95 % + 30 días sin popup legacy |
| [EXP-CAP-063](../../../../openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md#requirement-exp-cap-063--retirada-menú-json-hub) | Retirada menú JSON-hub | Tabla `TbMenuLanzaderaJson` con prompts en Access → endpoints REST del backend | Cobertura de tests ≥ 95 % + 30 días sin escritura legacy |

Cada Requirement tiene escenarios de validación y de concurrencia o fallo. Ver el spec.

## Cómo se aplica a access2web-blueprint

- **D8** (hexagonal global): cada retirada es un swap de adapter (legacy → web) preservando el dominio intacto.
- **D82** (Expand and Contract en cada release): las retiradas son aditivas; rollback = `DROP SCHEMA <módulo> CASCADE` con el legacy intacto.
- **DA-1** (ROOT_PACKAGE = `app.src.modules`; PURE_LAYERS = {domain, ports, application}): ningún dominio importa `win32`, `ole32`, `MSComctlLib`; los adapters legacy se aíslan hasta la retirada final.
- **D77** (Docker desde día uno): el contenedor del Expedientes no monta el filesystem de Windows; las llamadas Win32 fallan en runtime por construcción, no por validación.
- **D155** (walkthrough aplicado a las 8 apps): la metodología v3/v4 (D146, D155, D168) deja evidencia de los `tool_warnings` conocidos (#1408 OPEN, #1412 OPEN) que aplican también a las retiradas.

## Core invariants

- **Retiradas graduales, no big-bang**: cada retirada tiene una bandera de feature-flag que controla el rollout; el legacy coexiste con la nueva implementación durante 30 días.
- **Cobertura de tests ≥ 95 % antes de retirar**: las suites `Test_*Helper_*` del legacy VBA se preservan como evidencia (D87) y los nuevos tests pytest cubren la sustitución.
- **Rollback siempre posible**: el legacy `.accdb` queda en `data/staging/<app>/` con el binario original, sin modificaciones; cualquier release fallida vuelve al estado anterior con `DROP SCHEMA`.
- **No se reintroducen legacy globals**: tras retirar `Variables Globales.bas`, los tests pinean `Test_NoMutableGlobals` que falla si cualquier módulo accede a globals mutables.

## Contributor checklist

- [ ] Si el PR sustituye un legacy adapter, mantiene la cobertura de tests del adaptador reemplazado (D87).
- [ ] Si el PR añade un feature-flag de retirada, documenta el criterio de «30 días sin invocación legacy».
- [ ] Si el PR elimina código legacy, el path `.accdb` original sigue accesible vía `data/staging/<app>/`.
- [ ] Si el PR introduce un componente web que reemplaza un ActiveX, el componente web tiene su test de smoke.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Lista de comprobación final

- [ ] Las 7 capabilities referencian anchors al spec fuente.
- [ ] Las decisiones D8/D82/DA-1/D77/D155 citadas están vigentes en `docs/architecture.md`.
- [ ] Tono castellano peninsular formal con usted.
- [ ] Cross-references desde DOCS.md y `docs/03-aplicaciones/expedientes/README.md` siguen resolviendo.

## Navigation

Previous: [integrations](integrations.md) | Next: [capabilities-index](../capabilities-index.md)
