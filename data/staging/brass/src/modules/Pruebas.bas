Attribute VB_Name = "Pruebas"
Option Compare Database
Option Explicit




'Set m_ColNombreCamposLetraColumna = New Scripting.Dictionary
'    With m_ColNombreCamposLetraColumna
'        .Add "B", m_ColParaEnumNombreCamposEnInforme("6") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("6"))
'        .Add "F", m_ColParaEnumNombreCamposEnInforme("2") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("2"))
'        .Add "B", m_ColParaEnumNombreCamposEnInforme("3") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("3"))
'        .Add "F", m_ColParaEnumNombreCamposEnInforme("4") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("4"))
'        .Add "B", m_ColParaEnumNombreCamposEnInforme("5") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("5"))
'        .Add "F", m_ColParaEnumNombreCamposEnInforme("9") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("9"))
'        .Add "B", m_ColParaEnumNombreCamposEnInforme("7") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("7"))
'        .Add "F", m_ColParaEnumNombreCamposEnInforme("10") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("10"))
'        .Add "B", m_ColParaEnumNombreCamposEnInforme("8") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("8"))
'        .Add "F", m_ColParaEnumNombreCamposEnInforme("28") & "|" & m_ColParaNombreCamposInformeTabla(m_ColParaEnumNombreCamposEnInforme("28"))
'
'    End With

Public Function PruebagetActividadesGetronic(p_IDEvento As String, Optional ByRef p_Error As String) As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_Tipo As Variant
    Dim m_Resultado As String
    On Error GoTo errores
    
    m_Resultado = Constructor.getActividadesGetronic(p_IDEvento, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Debug.Print m_Resultado
    Exit Function
errores:
    If Err.Number <> 0 Then
        p_Error = "EL método PruebagetActividadesGetronic ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function
