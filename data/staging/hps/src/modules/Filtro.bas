Attribute VB_Name = "Filtro"
Option Compare Database
Option Explicit

Public Function EstablecerFiltroUsuarios( _
                                        p_Formulario As Form, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim frm As Form
    Dim m_SQLSinWhere As String
    Dim m_Where As String
    Dim m_WhereEstado As String
    Dim m_WhereEspecialidad As String
    Dim strEstado As String
    Dim m_SQL As String
    Dim m_Parametro As String
    On Error GoTo errores
    
    p_Error = ""
    If Not FormularioAbierto("FormInicial00Principal") Then
        Exit Function
    End If
    Set frm = Forms("FormInicial00Principal").Controls("SubFormCentral").Form
    
    Set frm = frm.Controls("SubFormResultados").Form
    
    m_SQLSinWhere = "SELECT * " & _
                    "FROM TbDatosLocal "
    
    If m_ObjColParaFiltros Is Nothing Then
        p_Error = "No se ha rellenado la colección de parámetros"
        Err.Raise 1000
    End If
    If Not m_ObjColParaFiltros.Exists("Estado") Then
        p_Error = "No se ha rellenado la colección de parámetros"
        Err.Raise 1000
    End If
    
    m_WhereEstado = getWhereEstado(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_Formulario.Name = "FormInicial03ConsultasEstado" Then
        If m_WhereEstado <> "" Then
            m_WhereEstado = "WHERE " & m_WhereEstado
            m_SQL = m_SQLSinWhere & " " & m_WhereEstado & ";"
        Else
            m_SQL = m_SQLSinWhere
        End If
        frm.RecordSource = m_SQL
    
    ElseIf p_Formulario.Name = "FormInicial03ConsultasGrado" Then
        m_Where = getWhereGrado(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_Where = "" And m_WhereEstado = "" Then
            m_SQL = m_SQLSinWhere
        ElseIf m_Where <> "" And m_WhereEstado = "" Then
            m_SQL = m_SQLSinWhere & " " & "WHERE " & m_Where & ";"
        ElseIf m_Where = "" And m_WhereEstado <> "" Then
            m_SQL = m_SQLSinWhere & " " & "WHERE " & m_WhereEstado & ";"
        ElseIf m_Where <> "" And m_WhereEstado <> "" Then
            m_SQL = m_SQLSinWhere & " " & "WHERE (" & m_Where & ")" & " AND " & "(" & m_WhereEstado & ")" & ";"
        End If
        frm.RecordSource = m_SQL
   
    Else
        m_Where = getWhereParametro(p_Error)
'        If InStr(1, m_Where, "[Nombre]") <> 0 Then
'            m_Where = Replace(m_Where, "[Nombre]", "[TbDatosLocal].[Nombre]")
'        End If
'        If InStr(1, m_Where, "JuridicaContrato='") <> 0 Then
'            m_Where = Replace(m_Where, "JuridicaContrato=", "JuridicaContrato='")
'            If Right(m_Where, 1) = "'" Then
'                 m_Where = Left(m_Where, Len(m_Where) - 1)
'            End If
'            m_Where = "TbDatosLocal." & m_Where
'        ElseIf InStr(1, m_Where, "EmpresaUsuario='") <> 0 Then
'            m_Where = Replace(m_Where, "EmpresaUsuario='", "EmpresaUsuario='")
'            If Right(m_Where, 1) = "'" Then
'                 m_Where = Left(m_Where, Len(m_Where) - 1)
'            End If
'            m_Where = "TbDatosLocal." & m_Where
'        ElseIf InStr(1, m_Where, "EmpresaTramitadora='") <> 0 Then
'            m_Where = Replace(m_Where, "EmpresaTramitadora='", "EmpresaTramitadora='")
'            If Right(m_Where, 1) = "'" Then
'                 m_Where = Left(m_Where, Len(m_Where) - 1)
'            End If
'            m_Where = "TbDatosLocal." & m_Where
'        ElseIf InStr(1, m_Where, "IDExpediente=") <> 0 Then
'            m_Where = "TbDatosLocal." & m_Where
'        End If
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_Where = "" And m_WhereEstado = "" Then
            m_SQL = m_SQLSinWhere
        ElseIf m_Where <> "" And m_WhereEstado = "" Then
            m_SQL = m_SQLSinWhere & " " & "WHERE " & m_Where & ";"
        ElseIf m_Where = "" And m_WhereEstado <> "" Then
            m_SQL = m_SQLSinWhere & " " & "WHERE " & m_WhereEstado & ";"
        ElseIf m_Where <> "" And m_WhereEstado <> "" Then
            m_SQL = m_SQLSinWhere & " " & "WHERE (" & m_Where & ")" & " AND " & "(" & m_WhereEstado & ")" & ";"
        End If
        frm.RecordSource = m_SQL
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerFiltroUsuarios ha devuelto el error: " & Err.Description
    End If
    
End Function
Public Function getWhereParametro( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    On Error GoTo errores
    Dim m_Parametro As String
    Dim m_NombreCampo As Variant
    Dim m_NombreCampoResultado As String
    Dim m_ComparativoInicial As String
    Dim m_ComparativoFinal As String
    If m_ObjColParaFiltros Is Nothing Then
        Exit Function
    End If
    For Each m_NombreCampo In m_ObjColParaFiltros.Keys
        If CStr(m_NombreCampo) = "Estado" Then
            GoTo SiguienteCampo
        End If
        m_Parametro = m_ObjColParaFiltros(m_NombreCampo)
        If CStr(m_NombreCampo) = "Nombre" Then
            m_NombreCampoResultado = "[Nombre] & ' ' & [Apellido_1] & ' ' & [Apellido_2] "
            m_ComparativoInicial = "like'*"
            m_ComparativoFinal = "*'"
        ElseIf m_NombreCampo = "IDExpediente" Then
             m_NombreCampoResultado = m_NombreCampo
            m_ComparativoInicial = "="
            m_ComparativoFinal = ""
        Else
            m_NombreCampoResultado = m_NombreCampo
            m_ComparativoInicial = "='"
            m_ComparativoFinal = "'"
        End If
        If m_Parametro <> "" Then
            If getWhereParametro = "" Then
                getWhereParametro = m_NombreCampoResultado & m_ComparativoInicial & m_Parametro & m_ComparativoFinal
            Else
                getWhereParametro = getWhereParametro & " AND " & m_NombreCampoResultado & m_ComparativoInicial & m_Parametro & m_ComparativoFinal
            End If
        End If
SiguienteCampo:
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getWhereParametro ha devuelto el error: " & Err.Description
    End If
End Function
    
Public Function getWhereEstado( _
                                    Optional ByRef p_Error As String _
                                    ) As String

    
    Dim m_EstadoVisibleHPS As String
    On Error GoTo errores
    
    If m_ObjColParaFiltros Is Nothing Then
        Exit Function
    End If
    If Not m_ObjColParaFiltros.Exists("Estado") Then
        Exit Function
    End If
    m_EstadoVisibleHPS = m_ObjColParaFiltros("Estado")
    If m_EstadoVisibleHPS = "" Then
        Exit Function
    End If
    If Not m_ObjEntorno.ColVisiblesHPSEstados.Exists(m_EstadoVisibleHPS) Then
        p_Error = "No existe el estado indicado"
        Err.Raise 1000
    End If
    If m_EstadoVisibleHPS = m_ObjEntorno.ColEstadosVisiblesHPS(CStr(EnumEstadoVisibleHPS.Activo)) Then
        getWhereEstado = "HPS_NAC_Activo='Sí' OR HPS_OTAN_Activo='Sí' OR HPS_UE_Activo='Sí' OR HPS_ESA_Activo='Sí'"
    ElseIf m_EstadoVisibleHPS = m_ObjEntorno.ColEstadosVisiblesHPS(CStr(EnumEstadoVisibleHPS.ActivoAPuntoCaducar)) Then
        getWhereEstado = "HPS_NAC_ApuntoDeCaducar='Sí' OR HPS_OTAN_ApuntoDeCaducar='Sí' OR HPS_UE_ApuntoDeCaducar='Sí' OR HPS_ESA_ApuntoDeCaducar='Sí'"
    ElseIf m_EstadoVisibleHPS = m_ObjEntorno.ColEstadosVisiblesHPS(CStr(EnumEstadoVisibleHPS.Baja)) Then
        getWhereEstado = "HPS_NAC_Baja='Sí' OR HPS_OTAN_Baja='Sí' OR HPS_UE_Baja='Sí' OR HPS_ESA_Baja='Sí'"
    
    ElseIf m_EstadoVisibleHPS = m_ObjEntorno.ColEstadosVisiblesHPS(CStr(EnumEstadoVisibleHPS.Caducada)) Then
        getWhereEstado = "HPS_NAC_Caducado='Sí' OR HPS_OTAN_Caducado='Sí' OR HPS_UE_Caducado='Sí' OR HPS_ESA_Caducado='Sí'"
    
    ElseIf m_EstadoVisibleHPS = m_ObjEntorno.ColEstadosVisiblesHPS(CStr(EnumEstadoVisibleHPS.Irregular)) Then
        getWhereEstado = "HPS_NAC_Irregular='Sí' OR HPS_OTAN_Irregular='Sí' OR HPS_UE_Irregular='Sí' OR HPS_ESA_Irregular='Sí'"
    ElseIf m_EstadoVisibleHPS = m_ObjEntorno.ColEstadosVisiblesHPS(CStr(EnumEstadoVisibleHPS.NoPosee)) Then
        getWhereEstado = "HPS_NAC_SIN_DATOS='Sí' OR HPS_OTAN_SIN_DATOS='Sí' OR HPS_UE_SIN_DATOS='Sí' OR HPS_ESA_SIN_DATOS='Sí'"
    ElseIf m_EstadoVisibleHPS = m_ObjEntorno.ColEstadosVisiblesHPS(CStr(EnumEstadoVisibleHPS.PendienteRenovacion)) Then
        getWhereEstado = "HPS_NAC_PendienteRenovacion='Sí' OR HPS_OTAN_PendienteRenovacion='Sí' OR HPS_UE_PendienteRenovacion='Sí' OR HPS_ESA_PendienteRenovacion='Sí'"
    ElseIf m_EstadoVisibleHPS = m_ObjEntorno.ColEstadosVisiblesHPS(CStr(EnumEstadoVisibleHPS.SinOrdenDeRenovacion)) Then
        getWhereEstado = "HPS_NAC_Renovacion='Sí' OR HPS_OTAN_Renovacion='Sí' OR HPS_UE_Renovacion='Sí'OR HPS_ESA_Renovacion='Sí'"
    ElseIf m_EstadoVisibleHPS = m_ObjEntorno.ColEstadosVisiblesHPS(CStr(EnumEstadoVisibleHPS.Solicitada)) Then
        getWhereEstado = "HPS_NAC_Solicitado='Sí' OR HPS_OTAN_Solicitado='Sí' OR HPS_UE_Solicitado='Sí' OR HPS_ESA_Solicitado='Sí'"
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Ver getWhereEstado ha devuelto el error:" & vbNewLine & Err.Description
    End If
    
End Function

Public Function getWhereGrado( _
                                Optional ByRef p_Error As String _
                                ) As String

        
    On Error GoTo errores
    
    Dim m_LineaNAC As String
    Dim m_LineaOTAN As String
    Dim m_LineaUE As String
    Dim m_LineaESA As String
    Dim m_WhereEspecialidad As String
    Dim m_Comparador As String
    Dim m_ObjColLineas As Collection
    Dim m_linea As Variant
    Dim m_Resto As String
    Dim m_ValorNAC As String
    Dim m_ValorOTAN As String
    Dim m_ValorUE As String
    Dim m_ValorESA As String
    Dim m_ValorYO As String
    Dim m_ValorEspecialidad As String
    Dim m_WhereGrado As String
    
    Dim m_Col As Scripting.Dictionary
    
    On Error GoTo errores
    If m_ObjColParaFiltros Is Nothing Then
        p_Error = "Sin datos"
        Err.Raise 1000
    End If
    If Not m_ObjColParaFiltros.Exists("GradoNAC") Or _
        Not m_ObjColParaFiltros.Exists("GradoOTAN") Or _
        Not m_ObjColParaFiltros.Exists("GradoESA") Or _
        Not m_ObjColParaFiltros.Exists("GradoUE") Or _
        Not m_ObjColParaFiltros.Exists("Especialidad") Or _
        Not m_ObjColParaFiltros.Exists("YO") Then
        p_Error = "Sin datos suficientes"
        Err.Raise 1000
    End If
    m_ValorNAC = m_ObjColParaFiltros("GradoNAC")
    m_ValorOTAN = m_ObjColParaFiltros("GradoOTAN")
    m_ValorESA = m_ObjColParaFiltros("GradoESA")
    m_ValorUE = m_ObjColParaFiltros("GradoUE")
    m_ValorEspecialidad = m_ObjColParaFiltros("Especialidad")
    m_ValorYO = m_ObjColParaFiltros("YO")
    
    If m_ValorYO = "Y" Then
        m_Comparador = " AND "
    Else
        m_Comparador = " OR "
    End If
    If m_ValorNAC = "No Posee" Then
        m_LineaNAC = "HPS_NAC_ESTADO='No Posee'"
    Else
        If m_ValorNAC <> "" Then
            m_LineaNAC = "HPS_NAC_Grado='" & m_ValorNAC & "'"
        End If
    End If
    If m_ValorOTAN = "No Posee" Then
        m_LineaOTAN = "HPS_OTAN_ESTADO='No Posee'"
    Else
        If m_ValorOTAN <> "" Then
            m_LineaOTAN = "HPS_OTAN_Grado='" & m_ValorOTAN & "'"
        End If
    End If
    If m_ValorUE = "No Posee" Then
        m_LineaUE = "HPS_UE_ESTADO='No Posee'"
    Else
        If m_ValorUE <> "" Then
            m_LineaUE = "HPS_UE_Grado='" & m_ValorUE & "'"
        End If
    End If
    If m_ValorESA = "No Posee" Then
        m_LineaESA = "HPS_ESA_ESTADO='No Posee'"
    Else
        If m_ValorESA <> "" Then
            m_LineaESA = "HPS_ESA_Grado='" & m_ValorESA & "'"
        End If
    End If
    If m_ValorEspecialidad <> "" Then
        m_WhereEspecialidad = "HPS_NAC_Especialidad='" & m_ValorEspecialidad & "' " & _
                            "or" & " " & _
                            "HPS_OTAN_Especialidad='" & m_ValorEspecialidad & "' " & _
                            "or" & " " & _
                            "HPS_ESA_Especialidad='" & m_ValorEspecialidad & "' " & _
                            "or" & " " & _
                            "HPS_UE_Especialidad='" & m_ValorEspecialidad & "'"
    End If
    
    If Not (m_LineaNAC = "" And m_LineaOTAN = "" And m_ValorUE = "" And m_ValorESA = "" And m_WhereEspecialidad = "") Then
        Set m_ObjColLineas = New Collection
        If m_LineaNAC <> "" Then
            m_ObjColLineas.Add m_LineaNAC
        End If
        If m_LineaOTAN <> "" Then
            m_ObjColLineas.Add m_LineaOTAN
        End If
        If m_LineaUE <> "" Then
            m_ObjColLineas.Add m_LineaUE
        End If
        If m_LineaESA <> "" Then
            m_ObjColLineas.Add m_LineaESA
        End If
        If m_ObjColLineas.Count > 0 Then
            For Each m_linea In m_ObjColLineas
                If m_Resto = "" Then
                    m_Resto = m_linea & " "
                Else
                    m_Resto = m_Resto & m_Comparador & m_linea & " "
                End If
            Next
            If Right(m_Resto, Len(m_Comparador)) = m_Comparador Then
                m_WhereGrado = Left(m_Resto, Len(m_Resto) - Len(m_Comparador))
            Else
                m_WhereGrado = m_Resto
            End If
        End If
        If m_WhereGrado = "" Then
            m_WhereGrado = m_WhereEspecialidad
        Else
            If m_WhereEspecialidad <> "" Then
                m_WhereGrado = m_WhereGrado & " " & m_Comparador & " " & m_WhereEspecialidad
           
            End If
            
        End If
        getWhereGrado = m_WhereGrado
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Ver getWhereGrado ha devuelto el error:" & vbNewLine & Err.Description
    End If
    
End Function



