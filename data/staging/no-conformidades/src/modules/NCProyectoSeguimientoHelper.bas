Attribute VB_Name = "NCProyectoSeguimientoHelper"
Option Compare Database
Option Explicit

' ============================================
' NCProyectoSeguimientoHelper
' Helper testeable para Form_FormNCProyectoSeguimiento.
' Refactor de Issue #85: extrae la configuracion que el Form_Open debe aplicar
' (flags + TimerInterval) a un metodo puro, stateless y testable sin UI.
' ============================================

Public Function InicializarConfigParaFormOpen() As Object
    ' Devuelve un Dictionary con la configuracion que el form's Form_Open
    ' debe aplicar: flags + TimerInterval. Stateles: no tiene estado de modulo,
    ' no usa Me, no usa DoCmd, no toca DB.
    Set InicializarConfigParaFormOpen = CreateObject("Scripting.Dictionary")
    InicializarConfigParaFormOpen.Add "CargaInicialPendiente", True
    InicializarConfigParaFormOpen.Add "IndicadoresReiniciando", EnumSino.No
    InicializarConfigParaFormOpen.Add "TimerIntervalMs", 100
End Function

Public Function CargarIndicadoresSeguimientoProyecto( _
    ByVal p_Reiniciando As EnumSino, _
    Optional ByRef p_DuracionSegundos As Double, _
    Optional ByRef p_Error As String _
    ) As Boolean

    Dim startTime As Single

    On Error GoTo errores

    p_Error = ""
    p_DuracionSegundos = 0
    startTime = Timer

    PintarIndicadores p_Reiniciando:=p_Reiniciando, p_Modo:="PROYECTO", p_Error:=p_Error
    p_DuracionSegundos = ElapsedSeconds(startTime)
    If p_Error <> "" Then Err.Raise 1000

    CargarIndicadoresSeguimientoProyecto = True
    Exit Function

errores:
    If p_DuracionSegundos = 0 Then p_DuracionSegundos = ElapsedSeconds(startTime)
    If Err.Number <> 1000 Then
        p_Error = "El método CargarIndicadoresSeguimientoProyecto ha devuelto el error: " & Err.Description
    End If
End Function

Private Function ElapsedSeconds(ByVal p_StartTime As Single) As Double
    Dim endTime As Single

    endTime = Timer
    If endTime < p_StartTime Then
        ElapsedSeconds = (86400# - p_StartTime) + endTime
    Else
        ElapsedSeconds = endTime - p_StartTime
    End If
End Function
