Attribute VB_Name = "SuministradoresHelper"
Option Compare Database
Option Explicit

Public Function getSuministradorEnEdicion( _
                                                Optional p_Id As String, _
                                                Optional p_IDEdicion As String, _
                                                Optional p_IDSuministrador As String, _
                                                Optional ByRef p_Error As String _
                                                ) As SuministradorParaEvidencias
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    p_Error = ""
    
    If p_Id = "" And (p_IDEdicion = "" Or p_IDSuministrador = "") Then
        Exit Function
    End If
    If p_Id <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbProyectosEdicionesSuministradores " & _
                "WHERE ID=" & p_Id & ";"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbProyectosEdicionesSuministradores " & _
                "WHERE IDEdicion=" & p_IDEdicion & " " & _
                "AND IDSuministrador=" & p_IDSuministrador & ";"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            Set getSuministradorEnEdicion = New SuministradorParaEvidencias
            For Each m_Campo In getSuministradorEnEdicion.ColCampos
                getSuministradorEnEdicion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSuministradorEnEdicion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSuministradoresEnEdicion( _
                                                p_IDEdicion As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Suministrador As Suministrador
  
    On Error GoTo errores
    p_Error = ""
    
    If Nz(p_IDEdicion, "") = "" Then
        p_Error = "El parámetro p_IDEdicion es obligatorio."
        Exit Function
    End If
    
    ' Suministradores registrados en TbProyectosEdicionesSuministradores para la edición.
    m_SQL = "SELECT S.* " & _
            "FROM TbSuministradores AS S " & _
            "INNER JOIN TbProyectosEdicionesSuministradores AS ES ON S.IDSuministrador = ES.IDSuministrador " & _
            "WHERE ES.IDEdicion=" & p_IDEdicion & ";"
    
    ' Inicializar diccionario SIEMPRE (antes del EOF check)
    Set getSuministradoresEnEdicion = New Scripting.Dictionary
    getSuministradoresEnEdicion.CompareMode = TextCompare

    Set rcdDatos = getdb().OpenRecordset(m_SQL, dbOpenSnapshot)
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
            If Not getSuministradoresEnEdicion.Exists(m_Suministrador.IDSuministrador) Then
                getSuministradoresEnEdicion.Add m_Suministrador.IDSuministrador, m_Suministrador
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
        p_Error = "EL método getSuministradoresEnEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getSuministradoresEvidenciasEnEdicion( _
                                                        p_IDEdicion As String, _
                                                        Optional ByRef p_Error As String _
                                                        ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_EdicionSuministrador As SuministradorParaEvidencias
    
    On Error GoTo errores
    p_Error = ""
    
    If Nz(p_IDEdicion, "") = "" Then
        p_Error = "El parámetro p_IDEdicion es obligatorio."
        Exit Function
    End If
    
    m_SQL = "SELECT ES.* " & _
            "FROM TbProyectosEdicionesSuministradores AS ES " & _
            "WHERE ES.IDEdicion=" & p_IDEdicion & ";"
    
    ' Inicializar diccionario SIEMPRE (antes del EOF check)
    Set getSuministradoresEvidenciasEnEdicion = New Scripting.Dictionary
    getSuministradoresEvidenciasEnEdicion.CompareMode = TextCompare

    Set rcdDatos = getdb().OpenRecordset(m_SQL, dbOpenSnapshot)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_EdicionSuministrador = New SuministradorParaEvidencias
            For Each m_Campo In m_EdicionSuministrador.ColCampos
                m_EdicionSuministrador.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If Not getSuministradoresEvidenciasEnEdicion.Exists(m_EdicionSuministrador.ID) Then
                getSuministradoresEvidenciasEnEdicion.Add m_EdicionSuministrador.ID, m_EdicionSuministrador
            End If
            Set m_EdicionSuministrador = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método getSuministradoresEvidenciasEnEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

'-------------------------------------------
' Nombre: getSubcontratistasPorExpediente
' Propósito: Obtener subcontratistas directos (primer nivel) del expediente
'            cuyos padres raíz tienen ConsorcioPropio="Sí" en TbSuministradores.
' Retorno: Scripting.Dictionary (clave: IDSuministrador; valor: objeto Suministrador)
'------------------------------------------
Public Function getSubcontratistasPorExpediente( _
                                                p_IDExpediente As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Suministrador As Suministrador

    On Error GoTo errores
    p_Error = ""

    If Nz(p_IDExpediente, "") = "" Then
        p_Error = "El parámetro p_IDExpediente es obligatorio."
        Exit Function
    End If

    m_SQL = getSQLSubcontratistasPorExpediente(p_IDExpediente, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    ' Inicializar diccionario SIEMPRE (antes del EOF check)
    Set getSubcontratistasPorExpediente = New Scripting.Dictionary
    getSubcontratistasPorExpediente.CompareMode = TextCompare

    Set rcdDatos = getdb().OpenRecordset(m_SQL, dbOpenSnapshot)
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
            If Not getSubcontratistasPorExpediente.Exists(m_Suministrador.IDSuministrador) Then
                getSubcontratistasPorExpediente.Add m_Suministrador.IDSuministrador, m_Suministrador
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
        p_Error = "EL método getSubcontratistasPorExpediente ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getSQLSubcontratistasPorExpediente( _
                                                    p_IDExpediente As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String

    Dim m_SQL As String

    On Error GoTo errores
    p_Error = ""

    If Nz(p_IDExpediente, "") = "" Then
        p_Error = "El parámetro p_IDExpediente es obligatorio."
        Exit Function
    End If

    ' Hijos directos de empresas propias raíz del expediente.
    ' Si un hijo también es ConsorcioPropio='Sí', sus hijos NO se incluyen.
    m_SQL = ""
    m_SQL = m_SQL & "SELECT S.* "
    m_SQL = m_SQL & "FROM ((((TbExpedientesSuministradores AS ES "
    m_SQL = m_SQL & "INNER JOIN TbExpedientesSuministradores AS ESC "
    m_SQL = m_SQL & "ON ES.IDPadre = ESC.IDExpedienteSuministrador) "
    m_SQL = m_SQL & "INNER JOIN TbSuministradores AS SC "
    m_SQL = m_SQL & "ON ESC.IDSuministrador = SC.IDSuministrador) "
    m_SQL = m_SQL & "LEFT JOIN TbExpedientesSuministradores AS ESP "
    m_SQL = m_SQL & "ON ESC.IDPadre = ESP.IDExpedienteSuministrador) "
    m_SQL = m_SQL & "LEFT JOIN TbSuministradores AS SP "
    m_SQL = m_SQL & "ON ESP.IDSuministrador = SP.IDSuministrador) "
    m_SQL = m_SQL & "INNER JOIN TbSuministradores AS S "
    m_SQL = m_SQL & "ON ES.IDSuministrador = S.IDSuministrador "
    m_SQL = m_SQL & "WHERE ES.IDExpediente = " & p_IDExpediente & " "
    m_SQL = m_SQL & "AND ESC.IDExpediente = " & p_IDExpediente & " "
    m_SQL = m_SQL & "AND SC.ConsorcioPropio = 'Sí' "
    m_SQL = m_SQL & "AND (ESC.IDPadre IS NULL OR SP.ConsorcioPropio IS NULL OR SP.ConsorcioPropio <> 'Sí') "
    m_SQL = m_SQL & "ORDER BY S.IDSuministrador;"

    getSQLSubcontratistasPorExpediente = m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método getSQLSubcontratistasPorExpediente ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function SincronizarSuministradoresEnEdicion( _
                                                    Optional ByVal p_IDEdicion As String, _
                                                    Optional p_Edicion As Edicion, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    
    Dim m_ColFaltan As Scripting.Dictionary
    Dim m_ColSobran As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Edicion As Edicion
    
    On Error GoTo errores
    p_Error = ""
    
    If Not p_Edicion Is Nothing Then
        Set m_Edicion = p_Edicion
    Else
        Set m_Edicion = Constructor.getEdicion(p_IDEdicion:=p_IDEdicion, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If m_Edicion Is Nothing Then
        p_Error = "No se ha podido determinar la edición"
        Err.Raise 1000
    End If
    
    If Not m_Edicion.Proyecto Is Nothing Then
        VaciarColecciones m_Edicion, m_Edicion.Proyecto, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    Set m_ColFaltan = m_Edicion.ColSuministradoresFaltan
    p_Error = m_Edicion.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ColSobran = m_Edicion.ColSuministradoresSobran
    p_Error = m_Edicion.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If Not m_ColSobran Is Nothing Then
        For Each m_ID In m_ColSobran
            EliminarSuministradorEvidenciaEnEdicion CStr(m_ID), m_Edicion.IDEdicion, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End If
    
    If Not m_ColFaltan Is Nothing Then
        For Each m_ID In m_ColFaltan
            AltaSuministradorEvidenciaEnEdicion CStr(m_ID), m_Edicion.IDEdicion, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End If
    
    If Not m_Edicion.Proyecto Is Nothing Then
        VaciarColecciones m_Edicion, m_Edicion.Proyecto, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        m_Edicion.ColSuministradoresEvidenciasVaciar
        m_Edicion.ColSuministradoresFaltanVaciar
        m_Edicion.ColSuministradoresSobranVaciar
        m_Edicion.ColSuministradoresVaciar
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SincronizarSuministradoresEnEdicion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function AltaSuministradorEvidenciaEnEdicion( _
                                                    p_IDSuministrador As String, _
                                                    ByVal p_IDEdicion As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As SuministradorParaEvidencias
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_SumEv As SuministradorParaEvidencias
    
    On Error GoTo errores
    p_Error = ""
    
    If p_IDSuministrador = "" Or p_IDEdicion = "" Then
        p_Error = "Falta IDSuministrador o p_IDEdicion"
        Err.Raise 1000
    End If
    
    m_SQL = "SELECT ES.* " & _
            "FROM TbProyectosEdicionesSuministradores AS ES " & _
            "WHERE ES.IDEdicion=" & p_IDEdicion & " " & _
            "AND ES.IDSuministrador=" & p_IDSuministrador & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Set m_SumEv = New SuministradorParaEvidencias
            m_SumEv.ID = m_SumEv.IDCalculado
            m_SumEv.IDEdicion = p_IDEdicion
            m_SumEv.IDSuministrador = p_IDSuministrador
            .AddNew
                .Fields("ID") = m_SumEv.ID
                .Fields("IDEdicion") = m_SumEv.IDEdicion
                .Fields("IDSuministrador") = m_SumEv.IDSuministrador
            .Update
        Else
            Set m_SumEv = getSuministradorEnEdicion(p_IDEdicion:=p_IDEdicion, p_IDSuministrador:=p_IDSuministrador, p_Error:=p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End With
    
    rcdDatos.Close
    Set rcdDatos = Nothing
    Set AltaSuministradorEvidenciaEnEdicion = m_SumEv
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AltaSuministradorEvidenciaEnEdicion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EliminarSuministradorEvidenciaEnEdicion( _
                                                    p_IDSuministrador As String, _
                                                    ByVal p_IDEdicion As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    
    Dim m_SQL As String
    
    On Error GoTo errores
    p_Error = ""
    
    If p_IDSuministrador = "" Or p_IDEdicion = "" Then
        p_Error = "Falta IDSuministrador o p_IDEdicion"
        Err.Raise 1000
    End If
    
    m_SQL = "DELETE FROM TbProyectosEdicionesSuministradores " & _
            "WHERE IDEdicion=" & p_IDEdicion & " " & _
            "AND IDSuministrador=" & p_IDSuministrador & ";"
    getdb().Execute m_SQL
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EliminarSuministradorEvidenciaEnEdicion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EvidenciasSuministradoresCompletadas( _
                                                        Optional ByVal p_IDEdicion As String, _
                                                        Optional actualizando As EnumSiNo = EnumSiNo.No, _
                                                        Optional ByRef p_Error As String _
                                                        ) As EnumSiNo
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    p_Error = ""
    
    If p_IDEdicion = "" Then
        p_Error = "No se ha podido determinar la edición"
        Err.Raise 1000
    End If
    If actualizando = EnumSiNo.Sí Then
        SincronizarSuministradoresEnEdicion p_IDEdicion, , p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    m_SQL = "SELECT ES.* " & _
            "FROM TbProyectosEdicionesSuministradores AS ES " & _
            "WHERE ES.IDEdicion=" & p_IDEdicion & " " & _
            "AND ES.IDAnexo Is Null;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL, dbOpenSnapshot)
    With rcdDatos
        If .EOF Then
            EvidenciasSuministradoresCompletadas = EnumSiNo.Sí
        Else
            EvidenciasSuministradoresCompletadas = EnumSiNo.No
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método EvidenciasSuministradoresCompletadas ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function DetalleSuministradoresSinEvidenciaEnEdicion( _
                                                            Optional ByVal p_IDEdicion As String, _
                                                            Optional actualizando As EnumSiNo = EnumSiNo.No, _
                                                            Optional ByRef p_Error As String _
                                                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Detalle As String
    Dim m_Nombre As String
    
    On Error GoTo errores
    p_Error = ""
    
    If p_IDEdicion = "" Then
        p_Error = "No se ha podido determinar la edición"
        Err.Raise 1000
    End If
    If actualizando = EnumSiNo.Sí Then
        SincronizarSuministradoresEnEdicion p_IDEdicion, , p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    m_SQL = "SELECT S.IDSuministrador, S.Nombre " & _
            "FROM TbSuministradores AS S " & _
            "INNER JOIN TbProyectosEdicionesSuministradores AS ES " & _
            "ON S.IDSuministrador = ES.IDSuministrador " & _
            "WHERE ES.IDEdicion=" & p_IDEdicion & " " & _
            "AND ES.IDAnexo Is Null " & _
            "ORDER BY S.Nombre;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL, dbOpenSnapshot)
    With rcdDatos
        Do While Not .EOF
            m_Nombre = Nz(.Fields("Nombre").value, "")
            If m_Nombre = "" Then
                m_Nombre = "IDSuministrador " & Nz(.Fields("IDSuministrador").value, "")
            End If
            If m_Detalle <> "" Then
                m_Detalle = m_Detalle & "; "
            End If
            m_Detalle = m_Detalle & m_Nombre
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    If m_Detalle <> "" Then
        DetalleSuministradoresSinEvidenciaEnEdicion = "Falta evidencia de acuerdo con: " & m_Detalle
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método DetalleSuministradoresSinEvidenciaEnEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function VaciarColecciones( _
                                    ByVal p_Edicion As Edicion, _
                                    ByVal p_Proyecto As Proyecto, _
                                    Optional ByRef p_Error As String _
                                    ) As EnumSiNo

    On Error GoTo errores
    p_Error = ""
    
    p_Proyecto.ColSuministradoresVaciar
    p_Proyecto.Expediente.SuministradoresVaciar
    p_Edicion.ColSuministradoresEvidenciasVaciar
    p_Edicion.ColSuministradoresFaltanVaciar
    p_Edicion.ColSuministradoresSobranVaciar
    p_Edicion.ColSuministradoresVaciar
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método VaciarColecciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

'-------------------------------------------
' Nombre: getSuministradoresFaltantesEnEdicion
' Propósito: Obtener suministradores que deberían estar en una edición según
'            los subcontratistas actuales del expediente, pero no están en
'            TbProyectosEdicionesSuministradores.
'------------------------------------------
Public Function getSuministradoresFaltantesEnEdicion( _
                                                     p_IDExpediente As String, _
                                                     p_IDEdicion As String, _
                                                     Optional ByRef p_Error As String _
                                                     ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Suministrador As Suministrador

    On Error GoTo errores
    p_Error = ""

    If Nz(p_IDExpediente, "") = "" Then
        p_Error = "El parámetro p_IDExpediente es obligatorio."
        Exit Function
    End If

    If Nz(p_IDEdicion, "") = "" Then
        p_Error = "El parámetro p_IDEdicion es obligatorio."
        Exit Function
    End If

    m_SQL = ""
    m_SQL = m_SQL & "SELECT E.IDEdicion, E.IDProyecto, E.IDExpediente, S.* "
    m_SQL = m_SQL & "FROM ((SELECT DISTINCT PE.IDEdicion, P.IDProyecto, P.IDExpediente, ES.IDSuministrador "
    m_SQL = m_SQL & "FROM (((((TbProyectosEdiciones AS PE "
    m_SQL = m_SQL & "INNER JOIN TbProyectos AS P ON PE.IDProyecto = P.IDProyecto) "
    m_SQL = m_SQL & "INNER JOIN TbExpedientesSuministradores AS ES ON P.IDExpediente = ES.IDExpediente) "
    m_SQL = m_SQL & "INNER JOIN TbExpedientesSuministradores AS ESC ON ES.IDPadre = ESC.IDExpedienteSuministrador) "
    m_SQL = m_SQL & "INNER JOIN TbSuministradores AS SC ON ESC.IDSuministrador = SC.IDSuministrador) "
    m_SQL = m_SQL & "LEFT JOIN TbExpedientesSuministradores AS ESP ON ESC.IDPadre = ESP.IDExpedienteSuministrador) "
    m_SQL = m_SQL & "LEFT JOIN TbSuministradores AS SP ON ESP.IDSuministrador = SP.IDSuministrador "
    m_SQL = m_SQL & "WHERE PE.IDEdicion = " & p_IDEdicion & " "
    m_SQL = m_SQL & "AND P.IDExpediente = " & p_IDExpediente & " "
    m_SQL = m_SQL & "AND ESC.IDExpediente = " & p_IDExpediente & " "
    m_SQL = m_SQL & "AND SC.ConsorcioPropio = 'Sí' "
    m_SQL = m_SQL & "AND (ESC.IDPadre IS NULL OR SP.ConsorcioPropio IS NULL OR SP.ConsorcioPropio <> 'Sí')) AS E "
    m_SQL = m_SQL & "INNER JOIN TbSuministradores AS S ON E.IDSuministrador = S.IDSuministrador) "
    m_SQL = m_SQL & "LEFT JOIN (SELECT DISTINCT IDSuministrador "
    m_SQL = m_SQL & "FROM TbProyectosEdicionesSuministradores "
    m_SQL = m_SQL & "WHERE IDEdicion = " & p_IDEdicion & ") AS PESX "
    m_SQL = m_SQL & "ON E.IDSuministrador = PESX.IDSuministrador "
    m_SQL = m_SQL & "WHERE PESX.IDSuministrador IS NULL "
    m_SQL = m_SQL & "ORDER BY S.IDSuministrador;"

    ' Inicializar diccionario SIEMPRE (antes del EOF check)
    ' Así si no hay filas, retornamos {} en vez de Nothing
    Set getSuministradoresFaltantesEnEdicion = New Scripting.Dictionary
    getSuministradoresFaltantesEnEdicion.CompareMode = TextCompare

    Dim db As DAO.Database
    Set db = getdb(p_Error)
    If db Is Nothing Then
        Exit Function
    End If

    Set rcdDatos = db.OpenRecordset(m_SQL, dbOpenSnapshot)
    If rcdDatos Is Nothing Then
        Exit Function
    End If
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
            If Not getSuministradoresFaltantesEnEdicion.Exists(m_Suministrador.IDSuministrador) Then
                getSuministradoresFaltantesEnEdicion.Add m_Suministrador.IDSuministrador, m_Suministrador
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
        p_Error = "EL método getSuministradoresFaltantesEnEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

'-------------------------------------------
' Nombre: getSuministradoresSobrantesEnEdicion
' Propósito: Obtener suministradores que están en TbProyectosEdicionesSuministradores
'            para una edición, pero ya no pertenecen a los subcontratistas actuales
'            del expediente.
'------------------------------------------
Public Function getSuministradoresSobrantesEnEdicion( _
                                                    p_IDExpediente As String, _
                                                    p_IDEdicion As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Suministrador As Suministrador

    On Error GoTo errores
    p_Error = ""

    If Nz(p_IDExpediente, "") = "" Then
        p_Error = "El parámetro p_IDExpediente es obligatorio."
        Exit Function
    End If

    If Nz(p_IDEdicion, "") = "" Then
        p_Error = "El parámetro p_IDEdicion es obligatorio."
        Exit Function
    End If

    m_SQL = ""
    m_SQL = m_SQL & "SELECT PESX.IDEdicion, PE.IDProyecto, P.IDExpediente, S.* "
    m_SQL = m_SQL & "FROM ((((SELECT DISTINCT IDEdicion, IDSuministrador "
    m_SQL = m_SQL & "FROM TbProyectosEdicionesSuministradores "
    m_SQL = m_SQL & "WHERE IDEdicion = " & p_IDEdicion & ") AS PESX "
    m_SQL = m_SQL & "INNER JOIN TbProyectosEdiciones AS PE ON PESX.IDEdicion = PE.IDEdicion) "
    m_SQL = m_SQL & "INNER JOIN TbProyectos AS P ON PE.IDProyecto = P.IDProyecto) "
    m_SQL = m_SQL & "INNER JOIN TbSuministradores AS S ON PESX.IDSuministrador = S.IDSuministrador) "
    m_SQL = m_SQL & "LEFT JOIN (SELECT DISTINCT ES.IDSuministrador "
    m_SQL = m_SQL & "FROM (((TbExpedientesSuministradores AS ES "
    m_SQL = m_SQL & "INNER JOIN TbExpedientesSuministradores AS ESC ON ES.IDPadre = ESC.IDExpedienteSuministrador) "
    m_SQL = m_SQL & "INNER JOIN TbSuministradores AS SC ON ESC.IDSuministrador = SC.IDSuministrador) "
    m_SQL = m_SQL & "LEFT JOIN TbExpedientesSuministradores AS ESP ON ESC.IDPadre = ESP.IDExpedienteSuministrador) "
    m_SQL = m_SQL & "LEFT JOIN TbSuministradores AS SP ON ESP.IDSuministrador = SP.IDSuministrador "
    m_SQL = m_SQL & "WHERE ES.IDExpediente = " & p_IDExpediente & " "
    m_SQL = m_SQL & "AND ESC.IDExpediente = " & p_IDExpediente & " "
    m_SQL = m_SQL & "AND SC.ConsorcioPropio = 'Sí' "
    m_SQL = m_SQL & "AND (ESC.IDPadre IS NULL OR SP.ConsorcioPropio IS NULL OR SP.ConsorcioPropio <> 'Sí')) AS E "
    m_SQL = m_SQL & "ON PESX.IDSuministrador = E.IDSuministrador "
    m_SQL = m_SQL & "WHERE P.IDExpediente = " & p_IDExpediente & " "
    m_SQL = m_SQL & "AND E.IDSuministrador IS NULL "
    m_SQL = m_SQL & "ORDER BY S.IDSuministrador;"

    ' Inicializar diccionario SIEMPRE (antes del EOF check)
    Set getSuministradoresSobrantesEnEdicion = New Scripting.Dictionary
    getSuministradoresSobrantesEnEdicion.CompareMode = TextCompare

    Set rcdDatos = getdb().OpenRecordset(m_SQL, dbOpenSnapshot)
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
            If Not getSuministradoresSobrantesEnEdicion.Exists(m_Suministrador.IDSuministrador) Then
                getSuministradoresSobrantesEnEdicion.Add m_Suministrador.IDSuministrador, m_Suministrador
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
        p_Error = "EL método getSuministradoresSobrantesEnEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


