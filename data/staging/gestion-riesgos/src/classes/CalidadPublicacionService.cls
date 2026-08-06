VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalidadPublicacionService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function SaveAndRefreshTareasCalidad( _
    ByVal p_IDEdicion As String, _
    ByVal p_IDRiesgo As String, _
    ByVal p_Mutacion As String, _
    Optional ByRef p_Error As String _
) As CalidadPublicacionResultado
    On Error GoTo EH

    Dim resultado As CalidadPublicacionResultado
    Dim riesgoActual As riesgo
    Dim edicionActual As Edicion
    Dim scopeRefresco As String
    Dim tareasOriginal As TareasCalidad
    Dim habiaTareasOriginal As Boolean
    Dim restaurarTareasOriginal As Boolean

    p_Error = ""
    Set resultado = New CalidadPublicacionResultado
    resultado.TareasAntes = 0
    resultado.TareasDespues = 0
    resultado.RiesgoInvalidado = EnumSiNo.No
    resultado.EdicionRehidratada = EnumSiNo.No
    habiaTareasOriginal = Not m_ObjTareasCalidad Is Nothing
    If habiaTareasOriginal Then
        Set tareasOriginal = m_ObjTareasCalidad
    End If
    restaurarTareasOriginal = True

    If Trim$(p_IDEdicion) = "" Or Trim$(p_IDRiesgo) = "" Or _
        Trim$(p_IDEdicion) = "0" Or Trim$(p_IDRiesgo) = "0" Then
        p_Error = "CalidadPublicacionService.SaveAndRefreshTareasCalidad: identificadores inválidos"
        Set SaveAndRefreshTareasCalidad = resultado
        Exit Function
    End If

    Set riesgoActual = GetCachedRiesgo(p_IDRiesgo:=p_IDRiesgo, p_Error:=p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If riesgoActual Is Nothing Then
        p_Error = "No se pudo obtener el riesgo (IDRiesgo=" & p_IDRiesgo & ")"
        Err.Raise 1000
    End If

    resultado.TareasAntes = ContarTareasCalidadPendientes(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    scopeRefresco = ResolverScopeRefrescoCalidadRiesgo(p_Mutacion, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Set edicionActual = InvalidarYRehidratarRiesgoPorScope( _
                        p_IDEdicion:=p_IDEdicion, _
                        p_IDRiesgo:=p_IDRiesgo, _
                        p_Scope:=scopeRefresco, _
                        p_Error:=p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If edicionActual Is Nothing Then
        p_Error = "No se pudo rehidratar la edición (IDEdicion=" & p_IDEdicion & ")"
        Err.Raise 1000
    End If
    resultado.RiesgoInvalidado = EnumSiNo.Sí
    resultado.EdicionRehidratada = EnumSiNo.Sí

    If m_ObjTareasCalidad Is Nothing Then
        Set m_ObjTareasCalidad = New TareasCalidad
    End If
    m_ObjTareasCalidad.RefrescarPendientesCalidad p_Error
    If p_Error <> "" Then Err.Raise 1000

    resultado.TareasDespues = ContarTareasCalidadPendientes(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    restaurarTareasOriginal = False
    Set SaveAndRefreshTareasCalidad = resultado
    Exit Function

EH:
    If restaurarTareasOriginal Then
        On Error Resume Next
        If habiaTareasOriginal Then
            Set m_ObjTareasCalidad = tareasOriginal
        Else
            Set m_ObjTareasCalidad = Nothing
        End If
        On Error GoTo 0
    End If
    If p_Error = "" Then
        p_Error = "CalidadPublicacionService.SaveAndRefreshTareasCalidad: " & Err.Description
    End If
    Set SaveAndRefreshTareasCalidad = resultado
End Function

Private Function ContarTareasCalidadPendientes(Optional ByRef p_Error As String) As Long
    On Error GoTo EH

    Dim reseteandoTareas As EnumSiNo
    Dim regularizandoTareas As EnumSiNo
    Dim tareasTexto As String

    p_Error = ""
    If m_ObjTareasCalidad Is Nothing Then
        reseteandoTareas = EnumSiNo.No
        regularizandoTareas = EnumSiNo.No
        EstablecerTareasCalidad reseteandoTareas, regularizandoTareas, p_Error
        If p_Error <> "" Then Err.Raise 1000
    End If

    tareasTexto = m_ObjTareasCalidad.TareasTotalesCalidad
    If m_ObjTareasCalidad.Error <> "" Then
        p_Error = m_ObjTareasCalidad.Error
        Err.Raise 1000
    End If
    If Not IsNumeric(tareasTexto) Then
        p_Error = "El total de tareas de calidad no es numérico"
        Err.Raise 1000
    End If

    ContarTareasCalidadPendientes = CLng(tareasTexto)
    Exit Function

EH:
    If p_Error = "" Then
        p_Error = "CalidadPublicacionService.ContarTareasCalidadPendientes: " & Err.Description
    End If
End Function
