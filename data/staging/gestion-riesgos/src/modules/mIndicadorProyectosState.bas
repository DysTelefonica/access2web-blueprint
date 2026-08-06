Attribute VB_Name = "mIndicadorProyectosState"
Option Compare Database
Option Explicit

' Persistencia en memoria (durante la sesión) de:
' - proyectos excluidos
' - proyectos seleccionados (los que quedan tras aceptar)
Public g_Indicador_Excluidos As Object         ' Scripting.Dictionary (late binding)
Public g_Indicador_Seleccionados As Object     ' Scripting.Dictionary (late binding)
Public g_Indicador_ProyectosFormAbierto As Boolean
Public g_Indicador_ProyectosConfirmados As Boolean

Public Sub IndicadorState_Init()
    On Error Resume Next

    If g_Indicador_Excluidos Is Nothing Then
        Set g_Indicador_Excluidos = CreateObject("Scripting.Dictionary")
        g_Indicador_Excluidos.CompareMode = 1 ' TextCompare
    End If

    If g_Indicador_Seleccionados Is Nothing Then
        Set g_Indicador_Seleccionados = CreateObject("Scripting.Dictionary")
        g_Indicador_Seleccionados.CompareMode = 1 ' TextCompare
    End If
End Sub

Public Sub IndicadorState_ResetExcluidos()
    IndicadorState_Init
    g_Indicador_Excluidos.RemoveAll
End Sub

Public Sub IndicadorState_ResetSeleccionados()
    IndicadorState_Init
    g_Indicador_Seleccionados.RemoveAll
End Sub

Public Sub IndicadorState_ClearAll()
    IndicadorState_Init
    g_Indicador_Excluidos.RemoveAll
    g_Indicador_Seleccionados.RemoveAll
    g_Indicador_ProyectosConfirmados = False
    g_Indicador_ProyectosFormAbierto = False
End Sub

Public Sub IndicadorState_SetExcluido(ByVal p_IDProyecto As Long)
    IndicadorState_Init
    If p_IDProyecto > 0 Then
        If Not g_Indicador_Excluidos.Exists(CStr(p_IDProyecto)) Then
            g_Indicador_Excluidos.Add CStr(p_IDProyecto), p_IDProyecto
        End If
    End If
End Sub

Public Function IndicadorState_EsExcluido(ByVal p_IDProyecto As Long) As Boolean
    IndicadorState_Init
    If p_IDProyecto > 0 Then
        IndicadorState_EsExcluido = g_Indicador_Excluidos.Exists(CStr(p_IDProyecto))
    Else
        IndicadorState_EsExcluido = False
    End If
End Function

Public Sub IndicadorState_SetSeleccionadosDesdeLista(ByVal lst As ListBox)
    Dim dict As Object
    Dim i As Long
    Dim idp As String

    IndicadorState_Init
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' TextCompare

    ' OJO: la lista tiene encabezado en fila 0
    If lst.ListCount > 1 Then
        For i = 1 To lst.ListCount - 1
            idp = Nz(lst.Column(0, i), "")
            If IsNumeric(idp) Then
                If Not dict.Exists(CStr(idp)) Then dict.Add CStr(idp), CLng(idp)
            End If
        Next
    End If

    Set g_Indicador_Seleccionados = dict
End Sub

Public Function IndicadorState_CountSeleccionados() As Long
    IndicadorState_Init
    If g_Indicador_Seleccionados Is Nothing Then
        IndicadorState_CountSeleccionados = 0
    Else
        IndicadorState_CountSeleccionados = g_Indicador_Seleccionados.Count
    End If
End Function

Public Function IndicadorState_TieneSeleccionados() As Boolean
    IndicadorState_TieneSeleccionados = (IndicadorState_CountSeleccionados() > 0)
End Function

Public Function IndicadorState_GetIdsCsv(ByRef p_Error As String) As String
    Dim k As Variant
    Dim s As String

    On Error GoTo errores
    p_Error = ""
    IndicadorState_Init

    If g_Indicador_Seleccionados Is Nothing Or g_Indicador_Seleccionados.Count = 0 Then
        p_Error = "No hay proyectos seleccionados."
        Exit Function
    End If

    For Each k In g_Indicador_Seleccionados.keys
        If Len(s) > 0 Then s = s & ","
        s = s & CStr(k)
    Next

    IndicadorState_GetIdsCsv = s
    Exit Function

errores:
    p_Error = "IndicadorState_GetIdsCsv: " & Err.Number & vbCrLf & Err.Description
End Function

Public Sub IndicadorState_MarcarFormAbierto(ByVal p_Abierto As Boolean)
    g_Indicador_ProyectosFormAbierto = p_Abierto
End Sub

Public Sub IndicadorState_MarcarConfirmado(ByVal p_Confirmado As Boolean)
    g_Indicador_ProyectosConfirmados = p_Confirmado
End Sub

Public Function IndicadorState_HaConfirmado() As Boolean
    IndicadorState_HaConfirmado = (g_Indicador_ProyectosConfirmados = True)
End Function

Public Sub IndicadorState_ResetAll()
    On Error Resume Next
    IndicadorState_Init

    If Not g_Indicador_Excluidos Is Nothing Then g_Indicador_Excluidos.RemoveAll
    If Not g_Indicador_Seleccionados Is Nothing Then g_Indicador_Seleccionados.RemoveAll

    g_Indicador_ProyectosFormAbierto = False
    g_Indicador_ProyectosConfirmados = False
End Sub

Public Function IndicadorState_HaConfirmadoProyectos() As Boolean
    IndicadorState_HaConfirmadoProyectos = (g_Indicador_ProyectosConfirmados = True)
End Function
