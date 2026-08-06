Attribute VB_Name = "Pruebas"
Option Compare Database
Option Explicit

Public Function PruebaUsuariosCalidad() As String
    Dim m_Col As Scripting.Dictionary
    Dim m_Usuario As usuario
    Dim m_ID As Variant
    On Error GoTo Errores
   
    Set m_Col = m_ObjEntorno.ResponsablesCalidad
    
    If m_Col Is Nothing Then
        PruebaUsuariosCalidad = "NINGÚN DATO FILTRADO"
        Exit Function
    End If
    
    'Debug.Print "REGISTROS " & m_Col.Count
    For Each m_ID In m_Col
        Set m_Usuario = m_Col(m_ID)
       ' Debug.Print m_Usuario.nombre
        Set m_Usuario = Nothing
    Next
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Pruebas.PruebaUsuariosCalidad"
    Debug.Print errObj.FullDescription ' En pruebas, imprimimos el error a la consola.
End Function

Public Function PruebaUsuariosTecnicos() As String
    Dim m_Col As Scripting.Dictionary
    Dim m_Usuario As usuario
    Dim m_ID As Variant
    On Error GoTo Errores

    Set m_Col = m_ObjEntorno.ResponsablesTecnicos
    
    If m_Col Is Nothing Then
        PruebaUsuariosTecnicos = "NINGÚN DATO FILTRADO"
        Exit Function
    End If
    
    'Debug.Print "REGISTROS " & m_Col.Count
    For Each m_ID In m_Col
        Set m_Usuario = m_Col(m_ID)
        'Debug.Print m_Usuario.nombre
        Set m_Usuario = Nothing
    Next

    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Pruebas.PruebaUsuariosTecnicos"
    Debug.Print errObj.FullDescription
End Function
