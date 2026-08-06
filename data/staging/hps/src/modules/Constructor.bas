Attribute VB_Name = "Constructor"
Option Compare Database
Option Explicit

Public Function getUsuarioConectadoPorMaquina( _
                                                Optional ByRef p_Error As String _
                                                ) As Usuario
    Dim objNetwork As Object
    On Error GoTo errores
    Set objNetwork = CreateObject("Wscript.Network")
    Set getUsuarioConectadoPorMaquina = Constructor.getUsuario(, objNetwork.UserName, , , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set objNetwork = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioConectadoPorMaquina ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getSuministrador( _
                                Optional p_ID As String, _
                                Optional p_Nombre As String, _
                                Optional ByRef p_Error As String _
                                ) As Suministrador

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    
    On Error GoTo errores
    If p_ID = "" And p_Nombre = "" Then
        Exit Function
    End If
    If p_ID <> "" Then
        m_SQL = "SELECT * FROM TbSuministradores " & _
                "WHERE IDSuministrador=" & p_ID & ";"
    Else
        m_SQL = "SELECT * FROM TbSuministradores " & _
                "WHERE Nombre Like'*" & p_Nombre & "*';"
    End If
    
    Set rcdDatos = getdbExpedientes().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getSuministrador = New Suministrador
        For Each m_Campo In getSuministrador.ColCampos
            getSuministrador.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSuministrador ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getcadenaContratistas( _
                                         p_IDExp As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Nombre As String
    
    On Error GoTo errores
    m_SQL = "SELECT TbSuministradores.Nombre, TbSuministradores.Nemotecnico " & _
            "FROM TbSuministradores INNER JOIN TbExpedientesSuministradores " & _
            "ON TbSuministradores.IDSuministrador = TbExpedientesSuministradores.IDSuministrador " & _
            "WHERE (((TbExpedientesSuministradores.IDExpediente)=" & p_IDExp & "));"
    
   
    
    Set rcdDatos = getdbExpedientes().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_Nombre = Nz(.Fields("Nemotecnico"), "")
            If m_Nombre = "" Then
                m_Nombre = .Fields("Nombre")
            End If
            
            If getcadenaContratistas = "" Then
                getcadenaContratistas = m_Nombre
            Else
                getcadenaContratistas = getcadenaContratistas & "|" & m_Nombre
            End If
            
            m_Nombre = ""
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getcadenaContratistas ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSuministradores( _
                                Optional p_SoloTramitadorasHPS As EnumSiNo = EnumSiNo.No, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_Suministrador As Suministrador
    
    On Error GoTo errores
    If p_SoloTramitadorasHPS = Empty Then
        p_SoloTramitadorasHPS = EnumSiNo.No
    End If
    
    If p_SoloTramitadorasHPS = EnumSiNo.No Then
        m_SQL = "SELECT * FROM TbSuministradores ORDER BY Nombre;"
    Else
        m_SQL = "SELECT * FROM TbSuministradores " & _
                "WHERE TramitadoraHPS='Sí' ORDER BY Nombre;"
    End If
    
    Set rcdDatos = getdbExpedientes().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Suministrador = New Suministrador
            For Each m_Campo In m_Suministrador.ColCampos
                m_Suministrador.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getSuministradores Is Nothing Then
                Set getSuministradores = New Scripting.Dictionary
                getSuministradores.CompareMode = TextCompare
            End If
            If Not getSuministradores.Exists(m_Suministrador.IDSuministrador) Then
                getSuministradores.Add m_Suministrador.IDSuministrador, m_Suministrador
            End If
            Set m_Suministrador = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSuministradores ha devuelto el error: " & Err.Description
    End If
End Function
Public Function gelCadenaContratistas( _
                                            Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_CadenaContratistas As String
    
    On Error GoTo errores
   
    
    m_SQL = "SELECT DISTINCT  CadenaContratistas " & _
            "FROM TbExpedientesConEntidades " & _
            "WHERE Not CadenaContratistas Is Null;"
    
    Set rcdDatos = getdbExpedientes().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
           
            m_CadenaContratistas = .Fields("CadenaContratistas")
            If gelCadenaContratistas Is Nothing Then
                Set gelCadenaContratistas = New Scripting.Dictionary
                gelCadenaContratistas.CompareMode = TextCompare
            End If
            If Not gelCadenaContratistas.Exists(m_CadenaContratistas) Then
                gelCadenaContratistas.Add m_CadenaContratistas, m_CadenaContratistas
            End If
            
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método gelCadenaContratistas ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getMotivos( _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_Motivo As String
    
    On Error GoTo errores
    m_SQL = "TbMotivoHPS"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_Motivo = .Fields("MotivoHPS")
            
            If getMotivos Is Nothing Then
                Set getMotivos = New Scripting.Dictionary
                getMotivos.CompareMode = TextCompare
            End If
            If Not getMotivos.Exists(m_Motivo) Then
                getMotivos.Add m_Motivo, m_Motivo
            End If
            
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getMotivos ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getAnexoUsuarioHPS( _
                                    p_IDAnexo As String, _
                                    Optional ByRef p_Error As String, _
                                    Optional ByVal p_Db As DAO.Database = Nothing _
                                    ) As AnexoUsuarioHPS
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDAnexo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbAnexosUsuariosHPS " & _
            "WHERE IDAnexo=" & p_IDAnexo & ";"
    If p_Db Is Nothing Then
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
    Else
        Set rcdDatos = p_Db.OpenRecordset(m_SQL)
    End If
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getAnexoUsuarioHPS = New AnexoUsuarioHPS
        For Each m_Campo In getAnexoUsuarioHPS.ColCampos
            getAnexoUsuarioHPS.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    ' Remember the backend the instance was loaded from so lazy lookups
    ' (Usuario, URLAnexo, etc.) reuse it instead of calling getdb(),
    ' which would resolve a different store in test contexts where
    ' TempVars("DatosEnLocal") / m_URLRutaAplicacionesLocal do not match
    ' the injected HPST.accdb. Nothing is a valid value: legacy callers
    ' that pass Nothing keep the existing getdb()-fallback behaviour.
    Set getAnexoUsuarioHPS.p_Db = p_Db
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexoUsuarioHPS ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getAnexosUsuarioHPS( _
                                    Optional p_ID As String, _
                                    Optional ByRef p_AnexosHistoricos As EnumSiNo = EnumSiNo.No, _
                                    Optional ByRef p_Error As String, _
                                    Optional ByVal p_Db As DAO.Database = Nothing _
                                    ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Anexo As AnexoUsuarioHPS
    
    On Error GoTo errores
    
    If p_ID = "" Then
        If p_AnexosHistoricos = Empty Then
            m_SQL = "TbAnexosUsuariosHPS"
        ElseIf p_AnexosHistoricos = EnumSiNo.Sí Then
            m_SQL = "SELECT * " & _
                    "FROM TbAnexosUsuariosHPS " & _
                    "WHERE EsHistorico=Sí  ORDER BY NombreAnexo;"
        ElseIf p_AnexosHistoricos = EnumSiNo.No Then
            m_SQL = "SELECT * " & _
                    "FROM TbAnexosUsuariosHPS " & _
                    "WHERE EsHistorico=No ORDER BY NombreAnexo;"
        End If
    Else
        If p_AnexosHistoricos = Empty Then
            m_SQL = "SELECT * " & _
                    "FROM TbAnexosUsuariosHPS " & _
                    "WHERE IDUsuario=" & p_ID & "  ORDER BY NombreAnexo;"
        ElseIf p_AnexosHistoricos = EnumSiNo.Sí Then
            m_SQL = "SELECT * " & _
                    "FROM TbAnexosUsuariosHPS " & _
                    "WHERE IDUsuario=" & p_ID & " AND EsHistorico='Sí'  ORDER BY NombreAnexo;"
        ElseIf p_AnexosHistoricos = EnumSiNo.No Then
            m_SQL = "SELECT * " & _
                    "FROM TbAnexosUsuariosHPS " & _
                    "WHERE IDUsuario=" & p_ID & " AND EsHistorico='No'  ORDER BY NombreAnexo;"
        End If
        
            
    End If
    If p_Db Is Nothing Then
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
    Else
        Set rcdDatos = p_Db.OpenRecordset(m_SQL)
    End If
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Anexo = New AnexoUsuarioHPS
            For Each m_Campo In m_Anexo.ColCampos
                m_Anexo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAnexosUsuarioHPS Is Nothing Then
                Set getAnexosUsuarioHPS = New Scripting.Dictionary
                getAnexosUsuarioHPS.CompareMode = TextCompare
            End If
            If Not getAnexosUsuarioHPS.Exists(CStr(m_Anexo.IDAnexo)) Then
                getAnexosUsuarioHPS.Add CStr(m_Anexo.IDAnexo), m_Anexo
            End If
            Set m_Anexo = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexosUsuarioHPS ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getAnexoUsuarioSICA( _
                                    Optional p_IDAnexo As String, _
                                    Optional p_IDUsuarioSICA As String, _
                                    Optional ByRef p_Error As String _
                                    ) As AnexoUsuarioSICA
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDAnexo = "" And p_IDUsuarioSICA = "" Then
        Exit Function
    End If
    If p_IDAnexo <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbAnexosUsuariosSICA " & _
                "WHERE IDAnexo=" & p_IDAnexo & ";"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbAnexosUsuariosSICA " & _
                "WHERE IDUsuarioSICA='" & p_IDUsuarioSICA & "';"
    End If
    
        
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getAnexoUsuarioSICA = New AnexoUsuarioSICA
        For Each m_Campo In getAnexoUsuarioSICA.ColCampos
            getAnexoUsuarioSICA.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexoUsuarioSICA ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getAnexoUsuarioHistorico( _
                                            p_IDAnexo As String, _
                                            Optional ByRef p_Error As String _
                                            ) As AnexoUsuarioHistorico
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDAnexo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbAnexosUsuariosHistoricos " & _
            "WHERE IDAnexo=" & p_IDAnexo & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getAnexoUsuarioHistorico = New AnexoUsuarioHistorico
        For Each m_Campo In getAnexoUsuarioHistorico.ColCampos
            getAnexoUsuarioHistorico.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexoUsuarioHistorico ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getAnexosUsuarioHistorico( _
                                            p_IDUsuarioHistorico As String, _
                                            Optional ByRef p_Error As String, _
                                            Optional ByVal p_Db As DAO.Database = Nothing _
                                            ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Anexo As AnexoUsuarioHistorico
    
    On Error GoTo errores
    
    If p_IDUsuarioHistorico = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbAnexosUsuariosHistoricos " & _
            "WHERE IDUsuario=" & p_IDUsuarioHistorico & ";"
    If p_Db Is Nothing Then
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
    Else
        Set rcdDatos = p_Db.OpenRecordset(m_SQL)
    End If
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Do While Not .EOF
            Set m_Anexo = New AnexoUsuarioHistorico
            For Each m_Campo In m_Anexo.ColCampos
                m_Anexo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAnexosUsuarioHistorico Is Nothing Then
                Set getAnexosUsuarioHistorico = New Scripting.Dictionary
                getAnexosUsuarioHistorico.CompareMode = TextCompare
            End If
            If Not getAnexosUsuarioHistorico.Exists(m_Anexo.IDAnexo) Then
                getAnexosUsuarioHistorico.Add m_Anexo.IDAnexo, m_Anexo
            End If
            Set m_Anexo = Nothing
            .MoveNext
        Loop
        
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAnexosUsuarioHistorico ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getExpediente( _
                                p_IDExpediente As String, _
                                Optional ByRef p_Error As String _
                                ) As Expediente
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDExpediente = "" Then
        Exit Function
    End If
    m_SQL = "SELECT  * " & _
            "FROM TbExpedientes " & _
            "WHERE IDExpediente=" & p_IDExpediente & ";"
   
        
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getExpediente = New Expediente
        For Each m_Campo In getExpediente.ColCampos
            getExpediente.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getExpediente ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getHPS( _
                        p_TipoHPS As String, _
                        p_ID As String, _
                        Optional ByRef p_Error As String _
                        ) As HPS
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If p_TipoHPS = "" Or p_ID = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbHPS " & _
            "WHERE IDUsuario=" & p_ID & " AND TipoHPS='" & p_TipoHPS & "';"
        
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getHPS = New HPS
        For Each m_Campo In getHPS.ColCampos
            getHPS.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHPS ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getObservacion( _
                                p_IDObservacion As String, _
                                Optional ByRef p_Error As String _
                                ) As Observacion
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    
    On Error GoTo errores
    If p_IDObservacion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbObservaciones " & _
                "WHERE IDObservacion=" & p_IDObservacion & ";"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getObservacion = New Observacion
        For Each m_Campo In getObservacion.ColCampos
            getObservacion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getObservacion ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getObservacionHistorica( _
                                        p_IDObservacion As String, _
                                        Optional ByRef p_Error As String _
                                        ) As ObservacionHistorica
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    
    On Error GoTo errores
    If p_IDObservacion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbObservacionesHistoricas " & _
                "WHERE IDObservacion=" & p_IDObservacion & ";"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getObservacionHistorica = New ObservacionHistorica
        For Each m_Campo In getObservacionHistorica.ColCampos
            getObservacionHistorica.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getObservacionHistorica ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuarioHPS( _
                                Optional p_IDUsuario As String, _
                                Optional p_DNI As String, _
                                Optional ByRef p_Error As String, _
                                Optional ByVal p_Db As DAO.Database = Nothing _
                                ) As UsuarioHPS
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    
    On Error GoTo errores
    If p_IDUsuario = "" And p_DNI = "" Then
        Exit Function
    End If
    If p_IDUsuario <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbUsuarios " & _
                "WHERE ID=" & p_IDUsuario & ";"
    ElseIf p_DNI <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbUsuarios " & _
                "WHERE DNI='" & p_DNI & "';"
    End If
    
    If p_Db Is Nothing Then
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
    Else
        Set rcdDatos = p_Db.OpenRecordset(m_SQL)
    End If
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        Set getUsuarioHPS = New UsuarioHPS
        For Each m_Campo In getUsuarioHPS.ColCampos
            getUsuarioHPS.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioHPS ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuariosHistoricos( _
                                            Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_usuario As UsuarioHistorico
    
    On Error GoTo errores
    m_SQL = "TbUsuariosHistoricos"
    
   Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_usuario = New UsuarioHistorico
            For Each m_Campo In m_usuario.ColCampos
               ' Debug.Print m_Campo
               ' If CStr(m_Campo) = "SuministradorUsuario" Then Stop
                m_usuario.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getUsuariosHistoricos Is Nothing Then
                Set getUsuariosHistoricos = New Scripting.Dictionary
                getUsuariosHistoricos.CompareMode = TextCompare
            End If
            If Not getUsuariosHistoricos.Exists(CStr(m_usuario.ID)) Then
                getUsuariosHistoricos.Add CStr(m_usuario.ID), m_usuario
            End If
            Set m_usuario = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosHistoricos ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuariosLocalDeOrigen( _
                                            Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_DatoLocal As DatosLocal
    
    On Error GoTo errores
    m_SQL = "SELECT ID, DNI, TbUsuarios.Nombre, Apellido_1, Apellido_2, Telefono, TbSuministradores_1.Nombre AS EmpresaUsuario, " & _
            "TbSuministradores.Nombre AS EmpresaTramitadora, CadenaContratistas AS JuridicaContrato, Correo_e, F_Nacimiento,LugarNacimiento, " & _
            "Motivo_HPS, TbExpedientes.IDExpediente, F_Curso, CursoEnVigor, F_Baja, FAvisoConcesion, " & _
            "FechaHPSConcesionMinima, RequiereComunicacionConcesion, TbExpedientes.CodExp, Requiere_Curso, " & _
            "FechaPrimeraConvocatoria, FechaCorreoNoCurso, FechaSegundaConvocatoria " & _
            "FROM ((TbUsuarios INNER JOIN TbSuministradores " & _
            "ON TbUsuarios.IDEmpresaUsuario = TbSuministradores.IDSuministrador) " & _
            "INNER JOIN TbSuministradores AS TbSuministradores_1 " & _
            "ON TbUsuarios.IDEmpresaHPS = TbSuministradores_1.IDSuministrador) " & _
            "INNER JOIN TbExpedientes ON TbUsuarios.IDExpediente = TbExpedientes.IDExpediente;"
    
   Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_DatoLocal = New DatosLocal
            For Each m_Campo In m_DatoLocal.ColCampos
                m_DatoLocal.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getUsuariosLocalDeOrigen Is Nothing Then
                Set getUsuariosLocalDeOrigen = New Scripting.Dictionary
                getUsuariosLocalDeOrigen.CompareMode = TextCompare
            End If
            If Not getUsuariosLocalDeOrigen.Exists(CStr(m_DatoLocal.ID)) Then
                getUsuariosLocalDeOrigen.Add CStr(m_DatoLocal.ID), m_DatoLocal
            End If
            Set m_DatoLocal = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosLocalDeOrigen ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuariosHPSEnEmpresa( _
                                        p_IDEmpresa As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_usuario As UsuarioHPS
    
    On Error GoTo errores
    If p_IDEmpresa = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbUsuarios " & _
            "WHERE IDEmpresaHPS=" & p_IDEmpresa & _
            " or IDEmpresaUsuario=" & p_IDEmpresa & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_usuario = New UsuarioHPS
            For Each m_Campo In m_usuario.ColCampos
                m_usuario.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getUsuariosHPSEnEmpresa Is Nothing Then
                Set getUsuariosHPSEnEmpresa = New Scripting.Dictionary
                getUsuariosHPSEnEmpresa.CompareMode = TextCompare
            End If
            If Not getUsuariosHPSEnEmpresa.Exists(m_usuario.ID) Then
                getUsuariosHPSEnEmpresa.Add m_usuario.ID, m_usuario
            End If
            Set m_usuario = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "SELECT TbUsuarios.* " & _
            "FROM TbExpedientes INNER JOIN TbUsuarios ON TbExpedientes.IDexpediente = TbUsuarios.IDexpediente " & _
            "WHERE (((TbExpedientes.IDJuridicaContrato)=" & p_IDEmpresa & "));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_usuario = New UsuarioHPS
            For Each m_Campo In m_usuario.ColCampos
                m_usuario.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next

            If getUsuariosHPSEnEmpresa Is Nothing Then
                Set getUsuariosHPSEnEmpresa = New Scripting.Dictionary
                getUsuariosHPSEnEmpresa.CompareMode = TextCompare
            End If
            If Not getUsuariosHPSEnEmpresa.Exists(m_usuario.ID) Then
                getUsuariosHPSEnEmpresa.Add m_usuario.ID, m_usuario
            End If
            Set m_usuario = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosHPSEnEmpresa ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuariosHistoricosEnEmpresa( _
                                                p_IDEmpresa As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_UsuarioHistorico As UsuarioHistorico
    
    On Error GoTo errores
    If p_IDEmpresa = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbUsuariosHistoricos " & _
            "WHERE IDEmpresaHPS=" & p_IDEmpresa & _
            " or IDEmpresaUsuario=" & p_IDEmpresa & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_UsuarioHistorico = New UsuarioHistorico
            For Each m_Campo In m_UsuarioHistorico.ColCampos
                m_UsuarioHistorico.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getUsuariosHistoricosEnEmpresa Is Nothing Then
                Set getUsuariosHistoricosEnEmpresa = New Scripting.Dictionary
                getUsuariosHistoricosEnEmpresa.CompareMode = TextCompare
            End If
            If Not getUsuariosHistoricosEnEmpresa.Exists(m_UsuarioHistorico.ID) Then
                getUsuariosHistoricosEnEmpresa.Add m_UsuarioHistorico.ID, m_UsuarioHistorico
            End If
            Set m_UsuarioHistorico = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "SELECT TbUsuariosHistoricos.* " & _
            "FROM TbExpedientes INNER JOIN TbUsuariosHistoricos ON TbExpedientes.IDexpediente = TbUsuariosHistoricos.IDexpediente " & _
            "WHERE (((TbExpedientes.IDJuridicaContrato)=" & p_IDEmpresa & "));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_UsuarioHistorico = New UsuarioHistorico
            For Each m_Campo In m_UsuarioHistorico.ColCampos
                m_UsuarioHistorico.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next

            If getUsuariosHistoricosEnEmpresa Is Nothing Then
                Set getUsuariosHistoricosEnEmpresa = New Scripting.Dictionary
                getUsuariosHistoricosEnEmpresa.CompareMode = TextCompare
            End If
            If Not getUsuariosHistoricosEnEmpresa.Exists(m_UsuarioHistorico.ID) Then
                getUsuariosHistoricosEnEmpresa.Add m_UsuarioHistorico.ID, m_UsuarioHistorico
            End If
            Set m_UsuarioHistorico = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosHistoricosEnEmpresa ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuariosHPSEnExpediente( _
                                            p_IDExpediente As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_UsuarioHPS As UsuarioHPS
    
    On Error GoTo errores
    If p_IDExpediente = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbUsuarios " & _
            "WHERE IDExpediente=" & p_IDExpediente & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_UsuarioHPS = New UsuarioHPS
            For Each m_Campo In m_UsuarioHPS.ColCampos
                m_UsuarioHPS.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getUsuariosHPSEnExpediente Is Nothing Then
                Set getUsuariosHPSEnExpediente = New Scripting.Dictionary
                getUsuariosHPSEnExpediente.CompareMode = TextCompare
            End If
            If Not getUsuariosHPSEnExpediente.Exists(m_UsuarioHPS.ID) Then
                getUsuariosHPSEnExpediente.Add m_UsuarioHPS.ID, m_UsuarioHPS
            End If
            Set m_UsuarioHPS = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosHPSEnExpediente ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getUsuariosHistoricosHPSEnEmpresa( _
                                                p_IDEmpresa As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_UsuarioHPS As UsuarioHistorico
    
    On Error GoTo errores
    If p_IDEmpresa = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbUsuariosHistoricos " & _
            "WHERE IDEmpresaUsuario=" & p_IDEmpresa & " " & _
            "OR IDEmpresaHPS=" & p_IDEmpresa & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_UsuarioHPS = New UsuarioHistorico
            For Each m_Campo In m_UsuarioHPS.ColCampos
                m_UsuarioHPS.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getUsuariosHistoricosHPSEnEmpresa Is Nothing Then
                Set getUsuariosHistoricosHPSEnEmpresa = New Scripting.Dictionary
                getUsuariosHistoricosHPSEnEmpresa.CompareMode = TextCompare
            End If
            If Not getUsuariosHistoricosHPSEnEmpresa.Exists(m_UsuarioHPS.ID) Then
                getUsuariosHistoricosHPSEnEmpresa.Add m_UsuarioHPS.ID, m_UsuarioHPS
            End If
            Set m_UsuarioHPS = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosHistoricosHPSEnEmpresa ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuarioHistorico( _
                                        Optional p_IDUsuario As String, _
                                        Optional p_DNI As String, _
                                        Optional ByRef p_Error As String, _
                                        Optional ByVal p_Db As DAO.Database = Nothing _
                                        ) As UsuarioHistorico
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    
    On Error GoTo errores
    If p_IDUsuario = "" And p_DNI = "" Then
        Exit Function
    End If
    If p_IDUsuario <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbUsuariosHistoricos " & _
                "WHERE ID=" & p_IDUsuario & ";"
    ElseIf p_DNI <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbUsuariosHistoricos " & _
                "WHERE DNI='" & p_DNI & "';"
    End If
    
    If p_Db Is Nothing Then
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
    Else
        Set rcdDatos = p_Db.OpenRecordset(m_SQL)
    End If
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        Set getUsuarioHistorico = New UsuarioHistorico
        For Each m_Campo In getUsuarioHistorico.ColCamposSimplificado
            'If CStr(m_Campo) = "IDJuridicaContrato" Then Stop
            getUsuarioHistorico.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioHistorico ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuarioSICA( _
                                Optional p_IDUsuarioSICA As String, _
                                Optional p_IDUsuario As String, _
                                Optional p_IDUsuarioHistorico As String, _
                                Optional ByRef p_Error As String _
                                ) As UsuarioSICA
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    
    On Error GoTo errores
    If p_IDUsuarioSICA = "" And p_IDUsuario = "" And p_IDUsuarioHistorico = "" Then
        Exit Function
    End If
    If p_IDUsuarioSICA <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbUsuariosSICA " & _
                "WHERE ID='" & p_IDUsuarioSICA & "';"
    ElseIf p_IDUsuario <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbUsuariosSICA " & _
                "WHERE IDHPS=" & p_IDUsuario & ";"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbUsuariosSICA " & _
                "WHERE IDHPSHistorico=" & p_IDUsuarioHistorico & ";"
    End If
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        Set getUsuarioSICA = New UsuarioSICA
        For Each m_Campo In getUsuarioSICA.ColCampos
            getUsuarioSICA.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioSICA ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getUsuariosSICAEnEmpresa( _
                                            p_IDEmpresa As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_UsuarioSICA As UsuarioSICA
    
    On Error GoTo errores
    If p_IDEmpresa = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbUsuariosSICA " & _
            "WHERE IDEmpresaTramitadora=" & p_IDEmpresa & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_UsuarioSICA = New UsuarioSICA
            For Each m_Campo In m_UsuarioSICA.ColCampos
                m_UsuarioSICA.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getUsuariosSICAEnEmpresa Is Nothing Then
                Set getUsuariosSICAEnEmpresa = New Scripting.Dictionary
                getUsuariosSICAEnEmpresa.CompareMode = TextCompare
            End If
            If Not getUsuariosSICAEnEmpresa.Exists(m_UsuarioSICA.ID) Then
                getUsuariosSICAEnEmpresa.Add m_UsuarioSICA.ID, m_UsuarioSICA
            End If
            Set m_UsuarioSICA = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosSICAEnEmpresa ha devuelto el error: " & Err.Description
    End If
End Function



Public Function getDatosLocal( _
                                p_IDUsuario As String, _
                                Optional ByRef p_Error As String _
                                ) As DatosLocal
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDUsuario = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
        "FROM TbDatosLocal " & _
        "WHERE ID=" & p_IDUsuario & ";"
        
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getDatosLocal = New DatosLocal
        For Each m_Campo In getDatosLocal.ColCampos
            getDatosLocal.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDatosLocal ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getUsuario( _
                            Optional p_ID As String, _
                            Optional p_UsuarioRed As String, _
                            Optional p_Nombre As String, _
                            Optional p_Correo As String, _
                            Optional ByRef p_Error As String _
                            ) As Usuario

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_NombreCampoID As String
    Dim m_EsNumeroID As Boolean
    Dim m_Where As String
    Dim m_valorID As String
    Dim m_SQLInicial As String
    
    On Error GoTo errores
    If p_ID = "" And p_UsuarioRed = "" And p_Nombre = "" And p_Correo = "" Then
        Exit Function
    End If
    m_SQLInicial = "SELECT TbUsuariosAplicaciones.* " & _
                    "FROM TbUsuariosAplicaciones "
    
    If p_ID <> "" Then
        m_NombreCampoID = "ID"
        m_valorID = p_ID
        m_EsNumeroID = True
    ElseIf p_UsuarioRed <> "" Then
        m_NombreCampoID = "UsuarioRed"
        m_valorID = p_UsuarioRed
        m_EsNumeroID = False
    ElseIf p_Nombre <> "" Then
        m_NombreCampoID = "Nombre"
        m_valorID = p_Nombre
        m_EsNumeroID = False
    ElseIf p_Correo <> "" Then
        m_NombreCampoID = "CorreoUsuario"
        m_valorID = p_Correo
        m_EsNumeroID = False
    End If
    
    If m_EsNumeroID Then
        m_Where = m_NombreCampoID & "=" & m_valorID & ";"
    Else
        m_Where = m_NombreCampoID & "='" & m_valorID & "';"
    End If
    m_SQL = m_SQLInicial & "WHERE " & m_Where
    Set rcdDatos = getdbLanzadera().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        Set getUsuario = New Usuario
        For Each m_Campo In getUsuario.ColCampos
            getUsuario.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuario ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuarios( _
                                Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_ObjUsuario As Usuario
    
    
    On Error GoTo errores
    
    m_SQL = "TbUsuariosAplicaciones"
    
    Set rcdDatos = getdbLanzadera().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjUsuario = New Usuario
            For Each m_Campo In m_ObjUsuario.ColCampos
                m_ObjUsuario.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getUsuarios Is Nothing Then
                Set getUsuarios = New Scripting.Dictionary
                getUsuarios.CompareMode = TextCompare
            End If
            If Not getUsuarios.Exists(CStr(m_ObjUsuario.UsuarioRed)) Then
                getUsuarios.Add CStr(m_ObjUsuario.UsuarioRed), m_ObjUsuario
            End If
            Set m_ObjUsuario = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarios ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuariosParaIndicadores( _
                                            p_TipoIndicador As EnumTipoIndicador, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_Obj As Object
    Dim m_valorID As String
    
    
    On Error GoTo errores
    If p_TipoIndicador = 0 Then
        Exit Function
    End If
    If p_TipoIndicador = EnumTipoIndicador.PendientesCursoHPS Then
        m_SQL = "SELECT TbDatosLocal.* " & _
                "FROM TbDatosLocal " & _
                "WHERE (((TbDatosLocal.Requiere_Curso)='Sí'));"
    ElseIf p_TipoIndicador = EnumTipoIndicador.PtesPrimeraConvocatoriaCurso Then
        m_SQL = "SELECT TbDatosLocal.* " & _
                "FROM TbDatosLocal " & _
                "WHERE (((TbDatosLocal.Requiere_PrimeraConvocatoriaCurso)='Sí'));"
    ElseIf p_TipoIndicador = EnumTipoIndicador.PtesSegundaConvocatoriaCurso Then
        m_SQL = "SELECT TbDatosLocal.* " & _
                "FROM TbDatosLocal " & _
                "WHERE (((TbDatosLocal.Requiere_SegundaConvocatoriaCurso)='Sí'));"
    ElseIf p_TipoIndicador = EnumTipoIndicador.PtesEnvioCorreoJefeSeguridadCurso Then
        m_SQL = "SELECT TbDatosLocal.* " & _
                "FROM TbDatosLocal " & _
                "WHERE (((TbDatosLocal.Requiere_CorreoJefeSeguridadCurso)='Sí'));"
        
        
    
    ElseIf p_TipoIndicador = EnumTipoIndicador.HPSAPuntoDeCaducar Then
        m_SQL = "SELECT TbDatosLocal.* " & _
                "FROM TbDatosLocal " & _
                "WHERE HPS_NAC_ApuntoDeCaducar='Sí' OR " & _
                "HPS_OTAN_ApuntoDeCaducar='Sí' or " & _
                "HPS_ESA_ApuntoDeCaducar='Sí' OR " & _
                "HPS_UE_ApuntoDeCaducar='Sí';"
    
    ElseIf p_TipoIndicador = EnumTipoIndicador.HPSACaducadas Then
        m_SQL = "SELECT TbDatosLocal.* " & _
                "FROM TbDatosLocal " & _
                "WHERE HPS_NAC_Caducado='Sí' OR " & _
                "HPS_OTAN_Caducado='Sí' or " & _
                "HPS_ESA_Caducado='Sí' OR " & _
                "HPS_UE_Caducado='Sí';"
    ElseIf p_TipoIndicador = EnumTipoIndicador.HPSASolicitandose Then
        m_SQL = "SELECT TbDatosLocal.* " & _
                "FROM TbDatosLocal " & _
                "WHERE Not HPS_UE_F_Solicitud Is Null OR " & _
                "Not HPS_OTAN_F_Solicitud Is Null OR " & _
                "NOt HPS_ESA_F_Solicitud Is Null OR " & _
                "Not HPS_UE_F_Solicitud Is Null;"
    End If
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Obj = New DatosLocal
            
            For Each m_Campo In m_Obj.ColCampos
                m_Obj.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            m_valorID = m_Obj.getPropiedad("ID", p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If getUsuariosParaIndicadores Is Nothing Then
                Set getUsuariosParaIndicadores = New Scripting.Dictionary
                getUsuariosParaIndicadores.CompareMode = TextCompare
            End If
            If Not getUsuariosParaIndicadores.Exists(m_valorID) Then
                getUsuariosParaIndicadores.Add m_valorID, m_Obj
            End If
            Set m_Obj = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosParaIndicadores ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getAplicacionesPermisos( _
                                            p_CorreoUsuario As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_ObjUsuarioAplicacionPermisos As UsuarioAplicacionPermisos
            
    On Error GoTo errores
    If p_CorreoUsuario = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbUsuariosAplicacionesPermisos.* " & _
            "FROM TbUsuariosAplicacionesPermisos " & _
            "WHERE CorreoUsuario='" & p_CorreoUsuario & "';"
    Set rcdDatos = getdbLanzadera().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjUsuarioAplicacionPermisos = New UsuarioAplicacionPermisos
            For Each m_Campo In m_ObjUsuarioAplicacionPermisos.ColCampos
                m_ObjUsuarioAplicacionPermisos.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getAplicacionesPermisos Is Nothing Then
                Set getAplicacionesPermisos = New Scripting.Dictionary
                getAplicacionesPermisos.CompareMode = TextCompare
            End If
            If Not getAplicacionesPermisos.Exists(CStr(m_ObjUsuarioAplicacionPermisos.IDAplicacion)) Then
                getAplicacionesPermisos.Add m_ObjUsuarioAplicacionPermisos.IDAplicacion, m_ObjUsuarioAplicacionPermisos
            End If
            Set m_ObjUsuarioAplicacionPermisos = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getAplicacionesPermisos ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getObjAleatorio( _
                                    p_TipoObjeto As EnumTipoObjeto, _
                                    Optional ByRef p_Error As String _
                                    ) As Object
    
    Dim m_SQL As String
    Dim m_NombreTabla As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    
    Dim m_FilaInicial As Long
    Dim m_FilaFinal As Long
    Dim m_FilaAleatoria As Long
    Dim m_Fila As Long
    
    On Error GoTo errores
     p_Error = ""
     
    If p_TipoObjeto = EnumTipoObjeto.UsuarioHPS Then
        m_NombreTabla = "TbUsuarios"
    ElseIf p_TipoObjeto = EnumTipoObjeto.UsuarioSICA Then
        m_NombreTabla = "TbUsuariosSICA"
    Else
        p_Error = "Sólo admite USuarios y UsuariosSICA"
        Err.Raise 1000
    End If
    
    m_SQL = "SELECT " & m_NombreTabla & ".* " & _
            "FROM " & m_NombreTabla & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveLast
        .MoveFirst
        m_FilaFinal = .RecordCount
        m_FilaInicial = 1
        m_FilaAleatoria = Int((m_FilaFinal - m_FilaInicial + 1) * Rnd + m_FilaInicial)
        m_Fila = 1
        Do While Not .EOF
            If .AbsolutePosition + 1 = m_FilaAleatoria Then
                If p_TipoObjeto = EnumTipoObjeto.UsuarioHPS Then
                    Set getObjAleatorio = New UsuarioHPS
                ElseIf p_TipoObjeto = EnumTipoObjeto.UsuarioSICA Then
                    Set getObjAleatorio = New UsuarioSICA
                End If
                
                For Each m_Campo In getObjAleatorio.ColCampos
                    getObjAleatorio.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                rcdDatos.Close
                Set rcdDatos = Nothing
                Exit Function
            End If
            
            .MoveNext
            
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getObjAleatorio ha devuelto un error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function


Public Function getColResultadoBusqueda( _
                                        p_Palabra As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
   
    Dim rcdDatos As DAO.Recordset
    Dim fld As DAO.Field
    Dim m_SQL As String
    Dim Apartado As String
    Dim m_NombreCampo As String
    Dim m_Valor As String
    Dim m_Resultado As String
    On Error Resume Next
    
    If p_Palabra = "" Then
        Exit Function
    End If
    
    Apartado = "Usuarios"
    m_SQL = "TbDatosLocal"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                For Each fld In .Fields
                    m_NombreCampo = fld.Name
                    m_Valor = Nz(fld.value, "")
                    If m_Valor <> "" Then
                        If InStr(1, m_Valor, p_Palabra) <> 0 Then
                            If getColResultadoBusqueda Is Nothing Then
                                Set getColResultadoBusqueda = New Scripting.Dictionary
                                getColResultadoBusqueda.CompareMode = TextCompare
                            End If
                            m_Resultado = Nz(.Fields("ID").value, "") & "|" & Apartado & "|" & m_NombreCampo & "|" & m_Valor
                            If Not getColResultadoBusqueda.Exists(m_Resultado) Then
                                getColResultadoBusqueda.Add m_Resultado, m_Resultado
                            End If
                            
                        End If
                    End If
                    
                Next
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Apartado = "Usuarios SICA"
    m_SQL = "TbUsuariosSICALocal"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
    
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                For Each fld In .Fields
                    m_NombreCampo = fld.Name
                    m_Valor = Nz(fld.value, "")

                    If m_Valor <> "" Then
                        If InStr(1, m_Valor, p_Palabra) <> 0 Then
                            If getColResultadoBusqueda Is Nothing Then
                                Set getColResultadoBusqueda = New Scripting.Dictionary
                                getColResultadoBusqueda.CompareMode = TextCompare
                            End If
                            m_Resultado = Nz(.Fields("ID").value, "") & "|" & Apartado & "|" & m_NombreCampo & "|" & m_Valor
                            If Not getColResultadoBusqueda.Exists(m_Resultado) Then
                                getColResultadoBusqueda.Add m_Resultado, m_Resultado
                            End If
                        End If
                    End If
                    'Debug.Print rcdDatos.Fields("ID").Value, m_NombreCampo, m_Valor
                Next
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Apartado = "Usuarios Históricos"
    m_SQL = "TbUsuariosHistoricosLocal"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                For Each fld In .Fields
                    m_NombreCampo = fld.Name
                    m_Valor = Nz(fld.value, "")

                    If m_Valor <> "" Then
                        If InStr(1, m_Valor, p_Palabra) <> 0 Then
                            If getColResultadoBusqueda Is Nothing Then
                                Set getColResultadoBusqueda = New Scripting.Dictionary
                                getColResultadoBusqueda.CompareMode = TextCompare
                            End If
                            m_Resultado = Nz(.Fields("ID").value, "") & "|" & Apartado & "|" & m_NombreCampo & "|" & m_Valor
                            If Not getColResultadoBusqueda.Exists(m_Resultado) Then
                                getColResultadoBusqueda.Add m_Resultado, m_Resultado
                            End If
                        End If
                    End If
                    'Debug.Print rcdDatos.Fields("ID").Value, m_NombreCampo, m_Valor
                Next
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing

    
    Apartado = "Observaciones"
    m_SQL = "TbObservacionesLocal"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                For Each fld In .Fields
                    m_NombreCampo = fld.Name
                    m_Valor = Nz(fld.value, "")
                    If m_Valor <> "" Then
                        If InStr(1, m_Valor, p_Palabra) <> 0 Then
                            If getColResultadoBusqueda Is Nothing Then
                                Set getColResultadoBusqueda = New Scripting.Dictionary
                                getColResultadoBusqueda.CompareMode = TextCompare
                            End If
                            'If .Fields("IDObservacion") = "504" Then Stop
                            m_Resultado = Nz(.Fields("IDObservacion").value, "") & "|" & Apartado & "|" & m_NombreCampo & "|" & m_Valor
                            If Not getColResultadoBusqueda.Exists(m_Resultado) Then
                                getColResultadoBusqueda.Add m_Resultado, m_Resultado
                            End If
                        End If
                    End If
                    
                Next
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Apartado = "Observaciones Históricas"
    m_SQL = "TbObservacionesHistoricasLocal"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                For Each fld In .Fields
                    m_NombreCampo = fld.Name
                    m_Valor = Nz(fld.value, "")
                    If m_Valor <> "" Then
                        If InStr(1, m_Valor, p_Palabra) <> 0 Then
                            If getColResultadoBusqueda Is Nothing Then
                                Set getColResultadoBusqueda = New Scripting.Dictionary
                                getColResultadoBusqueda.CompareMode = TextCompare
                            End If
                            m_Resultado = Nz(.Fields("IDObservacion").value, "") & "|" & Apartado & "|" & m_NombreCampo & "|" & m_Valor
                            If Not getColResultadoBusqueda.Exists(m_Resultado) Then
                                getColResultadoBusqueda.Add m_Resultado, m_Resultado
                            End If
                        End If
                    End If
                    
                Next
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Apartado = "Anexos"
    m_SQL = "TbUsuarioAnexos"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                For Each fld In .Fields
                    m_NombreCampo = fld.Name
                    m_Valor = Nz(fld.value, "")
                    If m_Valor <> "" Then
                        If InStr(1, m_Valor, p_Palabra) <> 0 Then
                            If getColResultadoBusqueda Is Nothing Then
                                Set getColResultadoBusqueda = New Scripting.Dictionary
                                getColResultadoBusqueda.CompareMode = TextCompare
                            End If
                            m_Resultado = Nz(.Fields("IDAnexo").value, "") & "|" & Apartado & "|" & m_NombreCampo & "|" & m_Valor
                            If Not getColResultadoBusqueda.Exists(m_Resultado) Then
                                getColResultadoBusqueda.Add m_Resultado, m_Resultado
                            End If
                        End If
                    End If
                    
                Next
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Apartado = "Anexos SICA"
    m_SQL = "TbAnexosUsuariosSICA"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                For Each fld In .Fields
                    m_NombreCampo = fld.Name
                    m_Valor = Nz(fld.value, "")
                    If m_Valor <> "" Then
                        If InStr(1, m_Valor, p_Palabra) <> 0 Then
                            If getColResultadoBusqueda Is Nothing Then
                                Set getColResultadoBusqueda = New Scripting.Dictionary
                                getColResultadoBusqueda.CompareMode = TextCompare
                            End If
                            m_Resultado = Nz(rcdDatos.Fields("IDAnexo").value, "") & "|" & Apartado & "|" & m_NombreCampo & "|" & m_Valor
                            If Not getColResultadoBusqueda.Exists(m_Resultado) Then
                                getColResultadoBusqueda.Add m_Resultado, m_Resultado
                            End If
                            
                        End If
                    End If
                    
                Next
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Apartado = "Observaciones"
    m_SQL = "TbObservaciones"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                For Each fld In rcdDatos.Fields
                    m_NombreCampo = fld.Name
                    m_Valor = Nz(fld.value, "")
                    If m_Valor <> "" Then
                        If InStr(1, m_Valor, p_Palabra) <> 0 Then
                            If getColResultadoBusqueda Is Nothing Then
                                Set getColResultadoBusqueda = New Scripting.Dictionary
                                getColResultadoBusqueda.CompareMode = TextCompare
                            End If
                            m_Resultado = Nz(rcdDatos.Fields("IDObservacion").value, "") & "|" & Apartado & "|" & m_NombreCampo & "|" & m_Valor
                            If Not getColResultadoBusqueda.Exists(m_Resultado) Then
                                getColResultadoBusqueda.Add m_Resultado, m_Resultado
                            End If
                        End If
                    End If
                    
                Next
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Apartado = "Observaciones Históricas"
    m_SQL = "TbObservacionesHistoricas"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                For Each fld In rcdDatos.Fields
                    m_NombreCampo = fld.Name
                    m_Valor = Nz(fld.value, "")
                    If m_Valor <> "" Then
                        If InStr(1, m_Valor, p_Palabra) <> 0 Then
                            If getColResultadoBusqueda Is Nothing Then
                                Set getColResultadoBusqueda = New Scripting.Dictionary
                                getColResultadoBusqueda.CompareMode = TextCompare
                            End If
                            m_Resultado = Nz(rcdDatos.Fields("IDObservacion").value, "") & "|" & Apartado & "|" & m_NombreCampo & "|" & m_Valor
                            If Not getColResultadoBusqueda.Exists(m_Resultado) Then
                                getColResultadoBusqueda.Add m_Resultado, m_Resultado
                            End If
                        End If
                    End If
                    
                Next
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColResultadoBusqueda ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function getListaObservacionesParaExcel( _
                                                p_ID As String, _
                                                Optional ByRef p_Error As String _
                                            ) As String
                                
    Dim rcdDatos As DAO.Recordset
    Dim m_Fecha As String
    Dim m_Observacion As String
    Dim m_Cadena As String
    On Error GoTo errores
    
    p_Error = ""
    If p_ID = "" Then
        p_Error = "Ha de indicar el ID del usuario"
        Err.Raise 1000
    End If
    
    
    m_SQL = "SELECT TbObservacionesLocal.Fecha, TbObservacionesLocal.Observacion " & _
            "FROM TbObservacionesLocal " & _
            "WHERE (((TbObservacionesLocal.ID)=" & p_ID & ") AND ((TbObservacionesLocal.Tipo)='General')) " & _
            "ORDER BY TbObservacionesLocal.Fecha DESC;"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                m_Fecha = Nz(.Fields("Fecha"), "")
                m_Observacion = Nz(.Fields("Observacion"), "")
                If m_Cadena = "" Then
                    m_Cadena = m_Fecha & " | " & m_Observacion
                Else
                    m_Cadena = m_Cadena & vbCrLf & m_Fecha & " | " & m_Observacion
                End If
                .MoveNext
            Loop
        End If
    End With
    getListaObservacionesParaExcel = m_Cadena
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaObservacionesParaExcel ha devuelto el error: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    
    
End Function
Public Function getListaDatosLocalPorWhere( _
                                            Optional p_NombreCampo As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_ObjUsuario As UsuarioHPS
    Dim m_ID As String
    Dim m_objColCampos As Collection
    Dim m_NombreCampo As Variant
    Dim m_Valor As String
    Dim m_Where As String
    On Error GoTo errores
    
    If p_NombreCampo = "" Then
        m_SQL = "SELECT TbDatosLocal.* " & _
            "FROM TbDatosLocal;"
    Else
        m_Where = "WHERE HPS_NAC_" & p_NombreCampo & "='Sí' or " & _
                "HPS_OTAN_" & p_NombreCampo & " ='Sí' or " & _
                "HPS_ESA_" & p_NombreCampo & " ='Sí' or " & _
                "HPS_UE_" & p_NombreCampo & " ='Sí';"
        m_SQL = "SELECT TbDatosLocal.* " & _
            "FROM TbDatosLocal " & _
            m_Where
    End If
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Set m_ObjUsuario = Nothing
            Do While Not .EOF
                
                Set m_ObjUsuario = New UsuarioHPS
                For Each m_Campo In m_ObjUsuario.ColCampos
                    m_ObjUsuario.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next

                If getListaDatosLocalPorWhere Is Nothing Then
                    Set getListaDatosLocalPorWhere = New Scripting.Dictionary
                    getListaDatosLocalPorWhere.CompareMode = TextCompare
                End If
                If Not getListaDatosLocalPorWhere.Exists(m_ObjUsuario.ID) Then
                    getListaDatosLocalPorWhere.Add m_ObjUsuario.ID, m_ObjUsuario
                End If
                Set m_ObjUsuario = Nothing
                
                .MoveNext
            Loop
            
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaDatosLocalPorWhere ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function




Public Function getListaUsuarioAPuntoDeCaducar( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary

    
    On Error GoTo errores
    Set getListaUsuarioAPuntoDeCaducar = getListaDatosLocalPorWhere("ApuntoDeCaducar", p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaUsuarioAPuntoDeCaducar ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getObservaciones( _
                                    Optional p_ID As String, _
                                    Optional p_Tipo As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_Observacion As Observacion
    
    On Error GoTo errores
    
    If p_ID = "" Then
        Exit Function
    End If
    If p_Tipo = Empty Then
        m_SQL = "SELECT * " & _
                "FROM TbObservaciones " & _
                "WHERE ID=" & p_ID & " ORDER BY TbObservaciones.Fecha DESC;"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbObservaciones " & _
                "WHERE ID=" & p_ID & _
                " AND Tipo='" & p_Tipo & "'ORDER BY TbObservaciones.Fecha DESC;"
    End If
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Observacion = New Observacion
            For Each m_Campo In m_Observacion.ColCampos
                m_Observacion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next

            
            If getObservaciones Is Nothing Then
                Set getObservaciones = New Scripting.Dictionary
                getObservaciones.CompareMode = TextCompare
            End If
            If Not getObservaciones.Exists(m_Observacion.IDObservacion) Then
                getObservaciones.Add m_Observacion.IDObservacion, m_Observacion
            End If
            Set m_Observacion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getObservaciones ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuariosHPS( _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_UsuarioHPS As UsuarioHPS
    
    On Error GoTo errores
    m_SQL = "TbUsuarios"
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_UsuarioHPS = New UsuarioHPS
            For Each m_Campo In m_UsuarioHPS.ColCampos
                m_UsuarioHPS.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next

            
            If getUsuariosHPS Is Nothing Then
                Set getUsuariosHPS = New Scripting.Dictionary
                getUsuariosHPS.CompareMode = TextCompare
            End If
            If Not getUsuariosHPS.Exists(m_UsuarioHPS.ID) Then
                getUsuariosHPS.Add m_UsuarioHPS.ID, m_UsuarioHPS
            End If
            Set m_UsuarioHPS = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosHPS ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuariosHPSNoEntidades( _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_UsuarioHPS As UsuarioHPS
    
    On Error GoTo errores
    m_SQL = "SELECT TbUsuarios.* " & _
            "FROM TbUsuarios LEFT JOIN TbUsuariosEntidades ON TbUsuarios.ID = TbUsuariosEntidades.ID " & _
            "WHERE (((TbUsuariosEntidades.ID) Is Null));"
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_UsuarioHPS = New UsuarioHPS
            For Each m_Campo In m_UsuarioHPS.ColCampos
                m_UsuarioHPS.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next

            
            If getUsuariosHPSNoEntidades Is Nothing Then
                Set getUsuariosHPSNoEntidades = New Scripting.Dictionary
                getUsuariosHPSNoEntidades.CompareMode = TextCompare
            End If
            If Not getUsuariosHPSNoEntidades.Exists(m_UsuarioHPS.ID) Then
                getUsuariosHPSNoEntidades.Add m_UsuarioHPS.ID, m_UsuarioHPS
            End If
            Set m_UsuarioHPS = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosHPSNoEntidades ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getUsuariosEnEntidadNoExistentes( _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ID As String
    
    On Error GoTo errores
    m_SQL = "SELECT TbUsuariosEntidades.ID " & _
            "FROM TbUsuariosEntidades LEFT JOIN TbUsuarios ON TbUsuariosEntidades.ID = TbUsuarios.ID " & _
            "WHERE (((TbUsuarios.ID) Is Null));"
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_ID = Nz(.Fields("ID"), "")
            If Not m_ID = "" Then
                If getUsuariosEnEntidadNoExistentes Is Nothing Then
                    Set getUsuariosEnEntidadNoExistentes = New Scripting.Dictionary
                    getUsuariosEnEntidadNoExistentes.CompareMode = TextCompare
                End If
                If Not getUsuariosEnEntidadNoExistentes.Exists(m_ID) Then
                    getUsuariosEnEntidadNoExistentes.Add m_ID, m_ID
                End If
            End If

            
            
            
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariosEnEntidadNoExistentes ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getObservacionesHistoricas( _
                                            p_ID As String, _
                                            Optional p_Tipo As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_ObservacionHistorica As ObservacionHistorica
    
    On Error GoTo errores
    
    If p_ID = "" Then
        Exit Function
    End If
    If p_Tipo = Empty Then
        m_SQL = "SELECT * " & _
                "FROM TbObservacionesHistoricas " & _
                "WHERE ID=" & p_ID & " ORDER BY Fecha DESC;"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbObservacionesHistoricas " & _
                "WHERE ID=" & p_ID & _
                " AND Tipo='" & p_Tipo & "' ORDER BY Fecha DESC;"
    End If
    
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObservacionHistorica = New ObservacionHistorica
            For Each m_Campo In m_ObservacionHistorica.ColCampos
                m_ObservacionHistorica.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next

            
            If getObservacionesHistoricas Is Nothing Then
                Set getObservacionesHistoricas = New Scripting.Dictionary
                getObservacionesHistoricas.CompareMode = TextCompare
            End If
            If Not getObservacionesHistoricas.Exists(m_ObservacionHistorica.IDObservacion) Then
                getObservacionesHistoricas.Add m_ObservacionHistorica.IDObservacion, m_ObservacionHistorica
            End If
            Set m_ObservacionHistorica = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getObservacionesHistoricas ha devuelto el error: " & Err.Description
    End If
End Function






Public Function getListaColeccionesGradosHPS( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    
    Dim m_SQL As String
    Dim m_TipoHPS As String
    Dim m_Grado As String
    Dim m_ColGradosNAC As Scripting.Dictionary
    Dim m_ColGradosOTAN As Scripting.Dictionary
    Dim m_ColGradosUE As Scripting.Dictionary
    Dim m_ColGradosESA As Scripting.Dictionary
    
    On Error GoTo errores
    p_Error = ""
    
    
    m_SQL = "SELECT * " & _
            "FROM TbHPSGrado;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                m_TipoHPS = Nz(.Fields("TipoHPS"), "")
                m_Grado = Nz(.Fields("Grado"), "")
                If m_TipoHPS = "Nacional" Then
                    If m_ColGradosNAC Is Nothing Then
                        Set m_ColGradosNAC = New Scripting.Dictionary
                        m_ColGradosNAC.CompareMode = TextCompare
                    End If
                    m_ColGradosNAC.Add m_Grado, m_Grado
                ElseIf m_TipoHPS = "OTAN" Then
                    If m_ColGradosOTAN Is Nothing Then
                        Set m_ColGradosOTAN = New Scripting.Dictionary
                        m_ColGradosOTAN.CompareMode = TextCompare
                    End If
                    If Not m_ColGradosOTAN.Exists(m_Grado) Then
                        m_ColGradosOTAN.Add m_Grado, m_Grado
                    End If
                    
                ElseIf m_TipoHPS = "UE" Then
                    If m_ColGradosUE Is Nothing Then
                        Set m_ColGradosUE = New Scripting.Dictionary
                        m_ColGradosUE.CompareMode = TextCompare
                    End If
                    If Not m_ColGradosUE.Exists(m_Grado) Then
                        m_ColGradosUE.Add m_Grado, m_Grado
                    End If
                ElseIf m_TipoHPS = "ESA" Then
                    If m_ColGradosESA Is Nothing Then
                        Set m_ColGradosESA = New Scripting.Dictionary
                        m_ColGradosESA.CompareMode = TextCompare
                    End If

                   If Not m_ColGradosESA.Exists(m_Grado) Then
                        m_ColGradosESA.Add m_Grado, m_Grado
                    End If
                Else
                    p_Error = "Tipo HPS desconocido"
                    Err.Raise 1000
                End If

                .MoveNext
            Loop

        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Not m_ColGradosNAC Is Nothing Then
        If getListaColeccionesGradosHPS Is Nothing Then
            Set getListaColeccionesGradosHPS = New Scripting.Dictionary
            getListaColeccionesGradosHPS.CompareMode = TextCompare
        End If
        If Not getListaColeccionesGradosHPS.Exists("Nacional") Then
            getListaColeccionesGradosHPS.Add "Nacional", m_ColGradosNAC
        End If
        
    End If
    If Not m_ColGradosOTAN Is Nothing Then
        If getListaColeccionesGradosHPS Is Nothing Then
            Set getListaColeccionesGradosHPS = New Scripting.Dictionary
            getListaColeccionesGradosHPS.CompareMode = TextCompare
        End If
        If Not getListaColeccionesGradosHPS.Exists("OTAN") Then
            getListaColeccionesGradosHPS.Add "OTAN", m_ColGradosOTAN
        End If
        
    End If
    If Not m_ColGradosUE Is Nothing Then
        If getListaColeccionesGradosHPS Is Nothing Then
            Set getListaColeccionesGradosHPS = New Scripting.Dictionary
            getListaColeccionesGradosHPS.CompareMode = TextCompare
        End If
        If Not getListaColeccionesGradosHPS.Exists("UE") Then
            getListaColeccionesGradosHPS.Add "UE", m_ColGradosUE
        End If
        
    End If
    If Not m_ColGradosESA Is Nothing Then
        If getListaColeccionesGradosHPS Is Nothing Then
            Set getListaColeccionesGradosHPS = New Scripting.Dictionary
            getListaColeccionesGradosHPS.CompareMode = TextCompare
        End If
        If Not getListaColeccionesGradosHPS.Exists("ESA") Then
            getListaColeccionesGradosHPS.Add "ESA", m_ColGradosESA
        End If
        
    End If
    

    Exit Function
errores:

    If Err.Number <> 1000 Then
        p_Error = "El método getListaColeccionesGradosHPS ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getID( _
                        p_NombreTabla As String, _
                        p_NombreCampoID As String, _
                        Optional ByRef p_Error As String _
                        ) As String

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ID As Long
    On Error GoTo errores
    
    
    If p_NombreTabla = "" Then
        p_Error = "Se ha de indicar p_NombreTabla"
        Err.Raise 1000
    End If
    If p_NombreCampoID = "" Then
        p_Error = "Se ha de indicar p_NombreCampoID"
        Err.Raise 1000
    End If
    If p_NombreTabla = "TbUsuarios" Or p_NombreTabla = "TbUsuariosHistoricos" Then
        getID = DameIDUsuario(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreTabla = "TbObservaciones" Or p_NombreTabla = "TbObservacionesHistoricas" Then
        getID = DameIDObservaciones(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreTabla = "TbUsuarioAnexos" Or p_NombreTabla = "TbAnexosHistoricos" Or p_NombreTabla = "TbUsuarioHistoricosAnexos" Then
        getID = DameIDAnexo(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    m_SQL = "SELECT Max(" & p_NombreTabla & "." & p_NombreCampoID & ") AS MáxID " & _
            "FROM " & p_NombreTabla & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("MáxID"), "")) Then
                m_ID = .Fields("MáxID")
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    getID = CStr(m_ID + 1)
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getID ha devuelto el error: " & Err.Description
    End If
End Function
Private Function DameIDUsuario( _
                                Optional ByRef p_Error As String _
                                ) As String

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_NombreTabla As String
    Dim m_NombreCampoID As String
    Dim m_ID As Long
    Dim m_IDMaxUsuario As Long
    Dim m_IDMaxUsuarioHistorico As Long
    On Error GoTo errores
    
    m_NombreTabla = "TbUsuarios"
    m_NombreCampoID = "ID"
    m_ID = 0
    
    m_SQL = "SELECT Max(" & m_NombreTabla & "." & m_NombreCampoID & ") AS MáxID " & _
            "FROM " & m_NombreTabla & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("MáxID"), "")) Then
                m_IDMaxUsuario = .Fields("MáxID")
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    m_NombreTabla = "TbUsuariosHistoricos"
    m_NombreCampoID = "ID"
   
    
    m_SQL = "SELECT Max(" & m_NombreTabla & "." & m_NombreCampoID & ") AS MáxID " & _
            "FROM " & m_NombreTabla & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("MáxID"), "")) Then
                m_IDMaxUsuarioHistorico = .Fields("MáxID")
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    If m_IDMaxUsuarioHistorico > m_IDMaxUsuario Then
        m_ID = m_IDMaxUsuarioHistorico
    Else
        m_ID = m_IDMaxUsuario
    End If
    
    DameIDUsuario = CStr(m_ID + 1)
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameIDUsuario ha devuelto el error: " & Err.Description
    End If
End Function

Private Function DameIDAnexo( _
                                Optional ByRef p_Error As String _
                                ) As String

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_NombreTabla As String
    Dim m_NombreCampoID As String
    Dim m_ID As Long
    Dim m_IDMaxAnexoEHistorico As Long
    Dim m_IDMaxAnexo As Long
    Dim m_IDMaxAnexoHistorico As Long
    Dim m_IDMaxAnexoUsuarioHistorico As Long
    On Error GoTo errores
    
    m_NombreTabla = "TbUsuarioAnexos"
    m_NombreCampoID = "IDAnexo"
    m_ID = 0
    
    m_SQL = "SELECT Max(" & m_NombreTabla & "." & m_NombreCampoID & ") AS MáxID " & _
            "FROM " & m_NombreTabla & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("MáxID"), "")) Then
                m_IDMaxAnexo = .Fields("MáxID")
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_NombreTabla = "TbAnexosHistoricos"
    m_NombreCampoID = "IDAnexo"
   
    
    m_SQL = "SELECT Max(" & m_NombreTabla & "." & m_NombreCampoID & ") AS MáxID " & _
            "FROM " & m_NombreTabla & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("MáxID"), "")) Then
                m_IDMaxAnexoHistorico = .Fields("MáxID")
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If m_IDMaxAnexoHistorico > m_IDMaxAnexo Then
        m_IDMaxAnexoEHistorico = m_IDMaxAnexoHistorico
    Else
        m_IDMaxAnexoEHistorico = m_IDMaxAnexo
    End If
    
    m_NombreTabla = "TbUsuarioHistoricosAnexos"
    m_NombreCampoID = "IDAnexo"
   
    
    m_SQL = "SELECT Max(" & m_NombreTabla & "." & m_NombreCampoID & ") AS MáxID " & _
            "FROM " & m_NombreTabla & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("MáxID"), "")) Then
                m_IDMaxAnexoUsuarioHistorico = .Fields("MáxID")
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    If m_IDMaxAnexoUsuarioHistorico > m_IDMaxAnexoEHistorico Then
        m_ID = m_IDMaxAnexoUsuarioHistorico
    Else
        m_ID = m_IDMaxAnexoEHistorico
    End If
    
    DameIDAnexo = CStr(m_ID + 1)
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameIDAnexo ha devuelto el error: " & Err.Description
    End If
End Function
Private Function DameIDObservaciones( _
                                    Optional ByRef p_Error As String _
                                    ) As String

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_NombreTabla As String
    Dim m_NombreCampoID As String
    Dim m_ID As Long
    Dim m_IDMaxUsuario As Long
    Dim m_IDMaxUsuarioHistorico As Long
    On Error GoTo errores
    
    m_NombreTabla = "TbObservaciones"
    m_NombreCampoID = "IDObservacion"
    m_ID = 0
    
    m_SQL = "SELECT Max(" & m_NombreTabla & "." & m_NombreCampoID & ") AS MáxID " & _
            "FROM " & m_NombreTabla & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("MáxID"), "")) Then
                m_IDMaxUsuario = .Fields("MáxID")
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    m_NombreTabla = "TbObservacionesHistoricas"
    m_NombreCampoID = "IDObservacion"
    
    
    m_SQL = "SELECT Max(" & m_NombreTabla & "." & m_NombreCampoID & ") AS MáxID " & _
            "FROM " & m_NombreTabla & ";"
   Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("MáxID"), "")) Then
                m_IDMaxUsuarioHistorico = .Fields("MáxID")
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    If m_IDMaxUsuarioHistorico > m_IDMaxUsuario Then
        m_ID = m_IDMaxUsuarioHistorico
    Else
        m_ID = m_IDMaxUsuario
    End If
    
    DameIDObservaciones = CStr(m_ID + 1)
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameIDObservaciones ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getExpedientes( _
                                Optional p_Nombre As String, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Expediente As Expediente
    
    On Error GoTo errores
    
    
    If p_Nombre = "" Then
        m_SQL = "TbExpedientes"
    Else
         m_SQL = "SELECT  * " & _
            "FROM TbExpedientes " & _
            "WHERE Nemotecnico Like '*" & p_Nombre & "*' " & _
            "OR CodExp Like '*" & p_Nombre & "*' " & _
            "OR Titulo Like '*" & p_Nombre & "*';"
            
        
       
    End If
    
    
    Set rcdDatos = getdbExpedientes().OpenRecordset(m_SQL)
    With rcdDatos
         If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
               Set m_Expediente = New Expediente
               For Each m_Campo In m_Expediente.ColCampos
                  m_Expediente.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                   If p_Error <> "" Then
                       Err.Raise 1000
                   End If
               Next
               If getExpedientes Is Nothing Then
                   Set getExpedientes = New Scripting.Dictionary
                   getExpedientes.CompareMode = TextCompare
               End If
               If Not getExpedientes.Exists(m_Expediente.IDExpediente) Then
                   getExpedientes.Add m_Expediente.IDExpediente, m_Expediente
               End If
               Set m_Expediente = Nothing
               .MoveNext
            Loop
         End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getExpedientes ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getExpedientesUsados( _
                                        Optional ByRef p_Error As String _
                                        ) As Collection

    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_Expediente As Expediente

    
    On Error GoTo errores
    m_SQL = "SELECT  distinct TbExpedientes.* " & _
            "FROM TbExpedientes INNER JOIN TbUsuarios ON TbExpedientes.IDExpediente = TbUsuarios.IDExpediente " & _
            "ORDER BY TbExpedientes.CodExp"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
                Set m_Expediente = New Expediente
                For Each m_Campo In m_Expediente.ColCampos
                    Debug.Print m_Campo
                    If CStr(m_Campo) = "POSTAGEDO" Then Stop
                    m_Expediente.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                
                If getExpedientesUsados Is Nothing Then
                    Set getExpedientesUsados = New Scripting.Dictionary
                    getExpedientesUsados.CompareMode = TextCompare
                End If
                If Not getExpedientesUsados.Exists(m_Expediente.IDExpediente) Then
                    getExpedientesUsados.Add m_Expediente.IDExpediente, m_Expediente
                End If
                Set m_Expediente = Nothing
                .MoveNext
            Loop
        End If
       
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "SELECT  distinct TbExpedientes.* " & _
            "FROM TbExpedientes INNER JOIN TbUsuariosHistoricos ON TbExpedientes.IDExpediente = TbUsuariosHistoricos.IDExpediente " & _
            "ORDER BY TbExpedientes.CodExp"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
                Set m_Expediente = New Expediente
                For Each m_Campo In m_Expediente.ColCampos
                    Debug.Print m_Campo
                    m_Expediente.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                
                If getExpedientesUsados Is Nothing Then
                    Set getExpedientesUsados = New Scripting.Dictionary
                    getExpedientesUsados.CompareMode = TextCompare
                End If
                If Not getExpedientesUsados.Exists(m_Expediente.IDExpediente) Then
                    getExpedientesUsados.Add m_Expediente.IDExpediente, m_Expediente
                End If
                Set m_Expediente = Nothing
                .MoveNext
            Loop
        End If
       
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getExpedientesUsados ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getListaParaCombo( _
                                    p_NombreCampo As String, _
                                    Optional p_NombreTabla As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Collection

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Dato As String
    
    On Error GoTo errores
    If p_NombreCampo = "" Then
        p_Error = "Se ha de indicar el nombre del Campo"
        Err.Raise 1000
    End If
    If p_NombreTabla = "" Then
        p_NombreTabla = "TbUsuarios"
    End If
    m_SQL = "SELECT DISTINCT " & p_NombreTabla & "." & p_NombreCampo & " " & _
            "FROM " & p_NombreTabla & " " & _
            "WHERE ((Not (" & p_NombreTabla & "." & p_NombreCampo & ") Is Null));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                m_Dato = Nz(.Fields(p_NombreCampo), "")
                If m_Dato <> "" Then
                    If getListaParaCombo Is Nothing Then
                        Set getListaParaCombo = New Collection
                    End If
                    getListaParaCombo.Add m_Dato
                End If
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaParaCombo ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getListaUsuariosHPSParaCombo( _
                                            Optional ByRef p_Error As String _
                                            ) As Collection

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    Dim m_Nombre As String
    Dim m_Apellido1 As String
    Dim m_Apellido2 As String
    
    On Error GoTo errores
    
    m_SQL = "SELECT TbUsuarios.ID,TbUsuarios.Nombre,TbUsuarios.Apellido_1,TbUsuarios.Apellido_2 " & _
            "FROM TbUsuarios ORDER BY Nombre;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                m_Nombre = Nz(.Fields("Nombre"), "")
                m_Apellido1 = Nz(.Fields("Apellido_1"), "")
                m_Apellido2 = Nz(.Fields("Apellido_2"), "")
                If getListaUsuariosHPSParaCombo Is Nothing Then
                    Set getListaUsuariosHPSParaCombo = New Collection
                End If
                getListaUsuariosHPSParaCombo.Add m_Nombre & " " & m_Apellido1 & " " & m_Apellido2 & "(" & Nz(.Fields("ID"), "") & ")"
                
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaUsuariosHPSParaCombo ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getListaUsuariosHistoricosParaCombo( _
                                            Optional ByRef p_Error As String _
                                            ) As Collection

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    Dim m_Nombre As String
    Dim m_Apellido1 As String
    Dim m_Apellido2 As String
    
    On Error GoTo errores
    
    m_SQL = "SELECT TbUsuariosHistoricos.ID,TbUsuariosHistoricos.Nombre,TbUsuariosHistoricos.Apellido_1,TbUsuariosHistoricos.Apellido_2 " & _
            "FROM TbUsuariosHistoricos ORDER BY Nombre;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                m_Nombre = Nz(.Fields("Nombre"), "")
                m_Apellido1 = Nz(.Fields("Apellido_1"), "")
                m_Apellido2 = Nz(.Fields("Apellido_2"), "")
                If getListaUsuariosHistoricosParaCombo Is Nothing Then
                    Set getListaUsuariosHistoricosParaCombo = New Collection
                End If
                getListaUsuariosHistoricosParaCombo.Add m_Nombre & " " & m_Apellido1 & " " & m_Apellido2 & "(" & Nz(.Fields("ID"), "") & ")"
                
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaUsuariosHistoricosParaCombo ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getConsultas( _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Consulta As Consulta
    
    On Error GoTo errores
    
    
    m_SQL = "TbConsultas"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Consulta = New Consulta
            For Each m_Campo In m_Consulta.ColCampos
                m_Consulta.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next

            If getConsultas Is Nothing Then
                Set getConsultas = New Scripting.Dictionary
                getConsultas.CompareMode = TextCompare
            End If
            If Not getConsultas.Exists(CStr(m_Consulta.IDConsulta)) Then
                getConsultas.Add CStr(m_Consulta.IDConsulta), m_Consulta
            End If
            Set m_Consulta = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getConsultas ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getSolicitudesHPSParaExpediente( _
                                                    p_IDExpediente As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    
    Dim m_JSON As Object
    Dim m_Col As Collection
    Dim i As Long
    Dim m_Valor As String
    
    Dim m_Cadena As String
    
    On Error GoTo errores
    If p_IDExpediente = "" Then
        Exit Function
    End If
    Set m_JSON = getJSonDeTabla("TbSolicitudes", "IDExpediente", p_IDExpediente, getdbSolicitudHPS(), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_JSON Is Nothing Then
        Exit Function
    End If
    Set m_Col = m_JSON
    For i = 1 To m_Col.Count
        m_Valor = m_JSON(i)("Nombre") & " " & m_JSON(i)("Apellido1") & " " & m_JSON(i)("Apellido2") & " " & m_JSON(i)("DNI")
        m_Valor = TextoParsedoParaTxt(m_Valor)
        
        
        If m_Cadena = "" Then
            m_Cadena = m_Valor
        Else
            m_Cadena = m_Cadena & vbNewLine & m_Valor
        End If
    Next
    getSolicitudesHPSParaExpediente = m_Cadena
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesHPSParaExpediente ha devuelto" & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function getSolicitudesHPSParaSuministrador( _
                                                    p_IDSuministrador As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    
    Dim m_JSON As Object
    Dim m_Col As Collection
    Dim i As Long
    Dim m_Valor As String
    Dim m_SQL As String
    Dim m_Cadena As String
    
    On Error GoTo errores
    If p_IDSuministrador = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbSolicitudes " & _
            "WHERE IDEmpresaUsuario=" & p_IDSuministrador & " " & _
            "OR IDEmpresaTramitadora=" & p_IDSuministrador & ";"
    Set m_JSON = getJSonDeSQL(m_SQL, getdbSolicitudHPS(), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_JSON Is Nothing Then
        Exit Function
    End If
    Set m_Col = m_JSON
    For i = 1 To m_Col.Count
        m_Valor = m_JSON(i)("Nombre") & " " & m_JSON(i)("Apellido1") & " " & m_JSON(i)("Apellido2") & " " & m_JSON(i)("DNI")
        m_Valor = TextoParsedoParaTxt(m_Valor)
        
        
        If m_Cadena = "" Then
            m_Cadena = m_Valor
        Else
            m_Cadena = m_Cadena & vbNewLine & m_Valor
        End If
    Next
    getSolicitudesHPSParaSuministrador = m_Cadena
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudesHPSParaSuministrador ha devuelto" & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function getExpedientesBusqueda( _
                                        Optional p_PalabraClave As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Expediente As Expediente
    
    On Error GoTo errores
    m_SQL = "SELECT * " & _
            "FROM TbExpedientes; "
       
    Set rcdDatos = getdbExpedientes().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_Expediente = New Expediente
                For Each m_Campo In m_Expediente.ColCampos
                    m_Expediente.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If p_PalabraClave <> "" Then
                    If p_PalabraClave <> m_Expediente.IDExpediente And _
                        InStr(1, m_Expediente.Nemotecnico, p_PalabraClave) = 0 And _
                        InStr(1, m_Expediente.Titulo, p_PalabraClave) = 0 Then
                        GoTo siguiente
                    End If
                End If

                If getExpedientesBusqueda Is Nothing Then
                    Set getExpedientesBusqueda = New Scripting.Dictionary
                    getExpedientesBusqueda.CompareMode = TextCompare
                End If
                If Not getExpedientesBusqueda.Exists(CStr(m_Expediente.IDExpediente)) Then
                    getExpedientesBusqueda.Add CStr(m_Expediente.IDExpediente), m_Expediente
                End If
siguiente:
                .MoveNext
            Loop
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getExpedientesBusqueda ha devuelto el error: " & Err.Description
    End If
End Function


