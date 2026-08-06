# Form_FormExpedienteSuministradores — pure-data audit

## Scope
- Source form: `src/forms/Form_FormExpedienteSuministradores.cls`
- Helper: `src/modules/modExpedienteSuministradoresHelper.bas`
- Tests: `src/modules/Test_ExpedienteSuministradoresHelper.bas`

## Extracted pure-data behavior
- Context tag parsing (`RELID`, `IDS`).
- UTE/socio label composition.
- Button enablement state for selected root/child nodes.
- Drop-target classification for root vs child destination.

## UI/DAO that intentionally stays in the form
- TreeView/ListView COM interaction, drag/drop event arguments, command bars, and context menus.
- DAO/service calls through `ExpedienteSuministradorServicio` and repository functions.
- Opening supplier management/detail forms.

## Anti-pattern check
- Helper/tests do not receive `ByRef p_Form`.
- Tests do not call `DoCmd.OpenForm`, `Forms(...)`, or `Screen.ActiveForm`.
