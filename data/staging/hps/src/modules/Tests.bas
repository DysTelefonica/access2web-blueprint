Attribute VB_Name = "Tests"
Option Compare Database
Option Explicit
Public Function getAleatorio(p_NumeroInicial As Integer, p_NumeroFinal As Integer, Optional ByRef p_Error As String) As Integer

    getAleatorio = Abs(Int((p_NumeroInicial - p_NumeroFinal + 1) * Rnd + p_NumeroInicial))
End Function


Public Function AnexosInalcanzables(Optional ByRef p_Error As String) As String
    Dim ColAnexos As Scripting.Dictionary
    Dim m_Anexo As AnexoUsuarioHPS
    Dim m_ID As Variant
    Dim URLAnexo As String
    Dim ColAnexosInalcanzables As Scripting.Dictionary
    
    On Error GoTo errores
    
    'vamos a obtener la colección de anexos
       
    
    Set ColAnexos = Constructor.getAnexosUsuarioHPS(, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    For Each m_ID In ColAnexos
        
        Set m_Anexo = ColAnexos(m_ID)
        
        URLAnexo = m_Anexo.URLAnexo
        If Not fso.FileExists(URLAnexo) Then
            If ColAnexosInalcanzables Is Nothing Then
                Set ColAnexosInalcanzables = New Scripting.Dictionary
                ColAnexosInalcanzables.CompareMode = TextCompare
            End If
            Set ColAnexosInalcanzables(m_ID) = m_Anexo
        End If
        Set m_Anexo = Nothing
    Next
    For Each m_ID In ColAnexosInalcanzables.Keys
        Set m_Anexo = ColAnexosInalcanzables(m_ID)
        URLAnexo = m_Anexo.URLAnexo
        If Not fso.FileExists(URLAnexo) Then
            VBA.DoEvents
            Debug.Print m_ID, URLAnexo
            VBA.DoEvents
        End If
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexosInalcanzables ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
    
End Function
Public Function DNIValido(Optional ByRef p_Error As String) As String
    
    Dim mObjExp As Object
    
    
    On Error GoTo errores
    
    Set mObjExp = CreateObject("VBSCRIPT.RegExp")
    mObjExp.Pattern = "((([X-Z])|([LM])){1}([-]?)((\d){7})([-]?)([A-Z]{1}))|((\d{8})([-]?)([A-Z]))"
    Debug.Print mObjExp.Test("02248439M")
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ExpresionRegular ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function NumeroTelValido(NTel As String, Optional ByRef p_Error As String) As Boolean
    
    Dim mObjExp As Object
    Dim Patron As String
    
    On Error GoTo errores
    Patron = "^(\(?\+[\d]{1,3}\)?)\Sí([\d]{1,5})\Sí([\d][\s\.-|.]?){6,7}$"
    Set mObjExp = CreateObject("VBSCRIPT.RegExp")
    mObjExp.Pattern = Patron
    NumeroTelValido = mObjExp.Test(NTel)
    Set mObjExp = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método NumeroTelValido ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function


Public Function TestDatosLocal( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_UsuarioHPS As UsuarioHPS
    Const m_ID As String = "774"
    Dim m_DatosLocal As DatosLocal
    Dim m_Campo As Variant
    
    On Error GoTo errores
    Set m_UsuarioHPS = Constructor.getUsuarioHPS(m_ID)
    Set m_DatosLocal = getDatosLocalDeUsuario(p_UsuarioHPS:=m_UsuarioHPS, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    For Each m_Campo In m_DatosLocal.ColCampos
        Debug.Print m_Campo, m_DatosLocal.getPropiedad(m_Campo)
    Next
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TestDatosLocal ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

