Attribute VB_Name = "modSegTareasHydration"
Option Compare Database
Option Explicit

' =============================================================================
' Test instrumentation for SegTareasProyecto hydration counters.
'
' Hosts the 4 counters that SegTareasProyecto.cls increments inside its
' AR/AC/NC property getters when m_TestingMode = True. The companion test
' module (Test_modFrmNCProyectoSeguimientoTareasListadoHelper.bas) toggles the
' counters-enabled flag and asserts the counts.
'
' Why a standard module and not on SegTareasProyecto:
'   - Project convention: NO class module uses VB_GlobalNameSpace=True or
'     VB_PredeclaredId=True (verified via Select-String — zero matches in src/).
'     So the variables cannot be Public members of a global default instance.
'   - Test_NCProyectoSeguimientoTareasListadoHelper.bas reads/writes them
'     WITHOUT module qualification; only Public variables in a standard
'     module are reachable that way.
'   - Pattern matches m_TestingMode / m_BackendSandboxURL in Variables
'     Globales.bas: shared state in standard modules, not classes.
'
' History: variables introduced in commit aa1ef79 (feat(issue-55)) but the
' declarations were never committed. The binary compiled at the time only
' because the source was incomplete and the declarations were added by hand
' to the binary at some point; subsequent binary syncs (slice 0.1, hotfix #111)
' discarded the manual patch and the declarations vanished, leaving the
' source's Option Explicit + unqualified reads unable to resolve.
' =============================================================================

Public m_TestSegTareasProyectoHydrationCountersEnabled As Boolean
Public m_TestSegTareasProyectoARHydrationCount As Long
Public m_TestSegTareasProyectoACHydrationCount As Long
Public m_TestSegTareasProyectoNCHydrationCount As Long
