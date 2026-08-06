Attribute VB_Name = "Test"
Option Compare Database
Option Explicit

Public Function HTMLHoy( _
                            Optional ByRef p_Error As String) As String
    
    
    
    On Error GoTo errores
    MostrarAperturas EnumApertura.HoyTodas, , , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTMLHoy ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    Debug.Print p_Error
End Function


Public Function CadenaUsuariosUsandoApp( _
                                        p_NumeroApp As String, _
                                        Optional ByRef p_Error As String) As String
    
    
    Dim m_App As Aplicacion
    Dim m_Usuario As Variant
    Dim m_Col As Scripting.Dictionary
    Dim m_Cadena As String
    
    On Error GoTo errores
    Set m_App = Constructor.getAplicacion(p_IDAplicacion:=p_NumeroApp, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_App Is Nothing Then
        p_Error = "Aplicación no existente"
        Err.Raise 1000
    End If
    Set m_Col = m_App.ColUsuariosUsandola(p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        CadenaUsuariosUsandoApp = "Nadie está usando la aplicación"
        Err.Raise 1000
    End If
    For Each m_Usuario In m_Col
        If m_Cadena = "" Then
            m_Cadena = CStr(m_Usuario)
        Else
            m_Cadena = m_Cadena & vbNewLine & CStr(m_Usuario)
        End If
    Next
    CadenaUsuariosUsandoApp = m_Cadena
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CadenaUsuariosUsandoApp ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    Debug.Print p_Error
End Function


