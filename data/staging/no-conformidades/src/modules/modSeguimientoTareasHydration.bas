Attribute VB_Name = "modSeguimientoTareasHydration"
Option Compare Database
Option Explicit

' =============================================================================
' Test instrumentation for modFrmNCProyectoSeguimientoTareasListadoHelper (formerly NCProyectoSeguimientoTareasListadoHelper; renamed 2026-06-26 per slice 7 of form-thin-helper-refactor).
'
' Hosts the 9 counters/flags/strings that the helper reads inside its
' ApplySeguimientoTareasFormFilters / GetTareasNCProyectosFiltrados paths
' when m_TestingMode = True, plus the cache-error R11 seam toggle used by
' Task 4.x. The companion test module
' (Test_modFrmNCProyectoSeguimientoTareasListadoHelper.bas) toggles the flags,
' sets the cache-error text, and asserts the helper delegation counts.
'
' Why a standard module and not Public members of the helper class:
'   - Project convention (verified via Select-String VB_GlobalNameSpace=True):
'     ZERO class modules in src/ use VB_GlobalNameSpace=True or
'     VB_PredeclaredId=True. Public members of helper classes cannot be
'     reached unqualified from the test module.
'   - Test_NCProyectoSeguimientoTareasListadoHelper.bas reads/writes these
'     WITHOUT module qualification (e.g.
'     m_TestSeguimientoTareasHelperDelegationSeamEnabled = True); only
'     Public variables in a standard module are reachable that way.
'   - Pattern matches m_TestingMode / m_BackendSandboxURL in Variables
'     Globales.bas and the modSegTareasHydration bas created 2026-06-25:
'     shared test state in standard modules, not classes.
'
' History: variables introduced in commit aa1ef79 (feat(issue-55)) together
' with the helper but the declarations were never committed. The binary
' compiled at the time only because the source was incomplete and the
' declarations were added by hand to the binary at some point; subsequent
' binary syncs (slice 0.1, hotfix #111, hotfix #119) discarded the manual
' patch and the declarations vanished, leaving the source's Option Explicit
' + unqualified reads unable to resolve. Same root cause as the
' modSegTareasHydration bug (PR #120).
' =============================================================================

' --- Delegation seam: form-level hook gating whether TestHook_*
'     paths run inside the helper.
Public m_TestFormSeguimientoTareasDelegationHookEnabled As Boolean

' --- Delegation seam: helper-level hook toggling delegation tracking.
Public m_TestSeguimientoTareasHelperDelegationSeamEnabled As Boolean

' --- Delegation counters: tracked on each helper invocation.
Public m_TestSeguimientoTareasHelperDelegationCallCount As Long
Public m_TestSeguimientoTareasHelperLastResponsableCalidad As String
Public m_TestSeguimientoTareasHelperLastResponsable As String
Public m_TestSeguimientoTareasHelperLastEstado As String
Public m_TestSeguimientoTareasHelperLastIDExpediente As String

' --- Cache-error R11 seam: forces the cache to fail so the helper's
'     fallback path can be exercised deterministically.
Public m_TestTareasHelperCacheErrorSeamEnabled As Boolean
Public m_TestTareasHelperCacheErrorText As String
