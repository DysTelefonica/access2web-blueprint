Attribute VB_Name = "PublicabilidadEdicion"
Option Compare Database
Option Explicit

Private Const CACHE_PUBLICABILIDAD_ALGORITMO_VERSION As Long = 2
Private Const CACHE_PUBLICABILIDAD_TIPO_EDICION As String = "EDICION"
Private Const CACHE_PUBLICABILIDAD_TIPO_RIESGO As String = "RIESGO"
Private Const CACHE_PUBLICABILIDAD_TABLA As String = "TbCachePublicabilidadEdicion"

Private Function CachePublicabilidad_ExisteTabla(p_db As DAO.Database, p_NombreTabla As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error GoTo errores

    If p_db Is Nothing Or Len(p_NombreTabla) = 0 Then
        CachePublicabilidad_ExisteTabla = False
        Exit Function
    End If

    For Each tdf In p_db.TableDefs
        If StrComp(tdf.Name, p_NombreTabla, vbTextCompare) = 0 Then
            CachePublicabilidad_ExisteTabla = True
            Exit Function
        End If
    Next

    CachePublicabilidad_ExisteTabla = False
    Set tdf = Nothing
    Exit Function
errores:
    CachePublicabilidad_ExisteTabla = False
    Err.Clear
    Set tdf = Nothing
End Function

Private Function CachePublicabilidad_ExisteCampo( _
                                                ByVal p_tdf As DAO.TableDef, _
                                                ByVal p_NombreCampo As String _
                                                ) As Boolean
    Dim fld As DAO.Field
    On Error GoTo errores

    Set fld = p_tdf.Fields(p_NombreCampo)
    CachePublicabilidad_ExisteCampo = True
    Set fld = Nothing
    Exit Function
errores:
    CachePublicabilidad_ExisteCampo = False
    Err.Clear
    Set fld = Nothing
End Function

Private Function CachePublicabilidad_ValidarSchema( _
                                                ByVal p_db As DAO.Database, _
                                                Optional ByRef p_Error As String _
                                                ) As EnumSiNo
    Dim tdf As DAO.TableDef
    Dim m_Campo As Variant
    Dim m_CamposRequeridos As Variant
    On Error GoTo errores

    p_Error = ""
    Set tdf = p_db.TableDefs(CACHE_PUBLICABILIDAD_TABLA)
    m_CamposRequeridos = Array("IDEdicion", "Tipo", "IDRiesgo", "Publicable", "Veredicto", "CodigoRiesgo", "Descripcion", "ChecksJson", "AlgoritmoVersion", "UpdatedAt")

    For Each m_Campo In m_CamposRequeridos
        If Not CachePublicabilidad_ExisteCampo(tdf, CStr(m_Campo)) Then
            p_Error = "La tabla " & CACHE_PUBLICABILIDAD_TABLA & " no tiene el campo requerido: " & CStr(m_Campo)
            Err.Raise 1000
        End If
    Next

    CachePublicabilidad_ValidarSchema = EnumSiNo.Sí
    Set tdf = Nothing
    Exit Function
errores:
    CachePublicabilidad_ValidarSchema = EnumSiNo.No
    If Err.Number <> 1000 Then
        p_Error = "El método CachePublicabilidad_ValidarSchema ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Set tdf = Nothing
End Function

Private Function CachePublicabilidad_SqlText(p_Value As Variant) As String
    If IsNull(p_Value) Then
        CachePublicabilidad_SqlText = "NULL"
        Exit Function
    End If
    CachePublicabilidad_SqlText = "'" & Replace(CStr(p_Value), "'", "''") & "'"
End Function

Private Function CachePublicabilidad_SqlLong(p_Value As Variant) As String
    If IsNull(p_Value) Or p_Value = "" Then
        CachePublicabilidad_SqlLong = "NULL"
        Exit Function
    End If
    If Not IsNumeric(p_Value) Then
        CachePublicabilidad_SqlLong = "NULL"
        Exit Function
    End If
    CachePublicabilidad_SqlLong = CStr(CLng(p_Value))
End Function

Private Function CachePublicabilidad_SqlYesNo(p_Value As Variant) As String
    CachePublicabilidad_SqlYesNo = IIf(CBool(p_Value), "True", "False")
End Function

Public Function CachePublicabilidad_AsegurarSchema(Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim m_UsarCache As String
    Dim m_ErrorLocal As String
    Dim m_TablaCreada As Boolean
    On Error GoTo errores

    p_Error = ""
    m_UsarCache = "No"
    On Error Resume Next
    m_UsarCache = CStr(Application.TempVars("Publicabilidad_Usar_Cache"))
    On Error GoTo errores
    If UCase$(Trim$(m_UsarCache)) = "NO" Then
        CachePublicabilidad_AsegurarSchema = "OK"
        Exit Function
    End If

    Set db = getdb(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    db.TableDefs.Refresh

    m_TablaCreada = False
    If Not CachePublicabilidad_ExisteTabla(db, CACHE_PUBLICABILIDAD_TABLA) Then
        On Error Resume Next
        db.Execute "CREATE TABLE " & CACHE_PUBLICABILIDAD_TABLA & " (" & _
                   "IDEdicion LONG NOT NULL, " & _
                   "Tipo TEXT(10) NOT NULL, " & _
                   "IDRiesgo LONG NOT NULL, " & _
                   "Publicable YESNO NOT NULL, " & _
                   "Veredicto LONG, " & _
                   "CodigoRiesgo TEXT(50), " & _
                   "Descripcion MEMO, " & _
                   "ChecksJson MEMO, " & _
                   "AlgoritmoVersion LONG NOT NULL, " & _
                   "UpdatedAt DATETIME NOT NULL" & _
                   ");"
        If Err.Number = 0 Then
            m_TablaCreada = True
        ElseIf InStr(1, Err.Description, "existe", vbTextCompare) > 0 Then
            Err.Clear
        Else
            p_Error = "El método CachePublicabilidad_AsegurarSchema ha devuelto el error: " & vbNewLine & Err.Description
            Err.Clear
            On Error GoTo errores
            Err.Raise 1000
        End If
        On Error GoTo errores

        If m_TablaCreada Then
            db.Execute "CREATE UNIQUE INDEX UX_" & CACHE_PUBLICABILIDAD_TABLA & " ON " & CACHE_PUBLICABILIDAD_TABLA & " (IDEdicion, Tipo, IDRiesgo);"
            db.Execute "CREATE INDEX IX_" & CACHE_PUBLICABILIDAD_TABLA & "_Consulta ON " & CACHE_PUBLICABILIDAD_TABLA & " (IDEdicion, Tipo, Publicable);"
            db.Execute "CREATE INDEX IX_" & CACHE_PUBLICABILIDAD_TABLA & "_Riesgo ON " & CACHE_PUBLICABILIDAD_TABLA & " (IDRiesgo);"
            db.TableDefs.Refresh
        End If
    End If

    db.TableDefs.Refresh
    m_ErrorLocal = ""
    If CachePublicabilidad_ValidarSchema(db, m_ErrorLocal) = EnumSiNo.No Then
        p_Error = m_ErrorLocal
        Err.Raise 1000
    End If

    CachePublicabilidad_AsegurarSchema = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CachePublicabilidad_AsegurarSchema ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CachePublicabilidad_InvalidarEdicion( _
                                                    ByVal p_IDEdicion As String, _
                                                    Optional p_db As DAO.Database = Nothing, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    Dim db As DAO.Database
    Dim m_SQL As String
    On Error GoTo errores

    p_Error = ""
    If p_IDEdicion = "" Or Not IsNumeric(p_IDEdicion) Then
        CachePublicabilidad_InvalidarEdicion = "OK"
        Exit Function
    End If

    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Set db = p_db
    End If

    If CachePublicabilidad_ExisteTabla(db, CACHE_PUBLICABILIDAD_TABLA) Then
        m_SQL = "DELETE FROM " & CACHE_PUBLICABILIDAD_TABLA & " WHERE IDEdicion=" & CStr(CLng(p_IDEdicion)) & ";"
        db.Execute m_SQL
    End If

    CachePublicabilidad_InvalidarEdicion = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CachePublicabilidad_InvalidarEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CachePublicabilidad_InvalidarRiesgo( _
                                                ByVal p_IDRiesgo As String, _
                                                Optional p_db As DAO.Database = Nothing, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim m_IDEdicion As String
    Dim m_ErrorLocal As String
    Dim m_SQL As String
    On Error GoTo errores

    p_Error = ""
    If p_IDRiesgo = "" Or Not IsNumeric(p_IDRiesgo) Then
        CachePublicabilidad_InvalidarRiesgo = "OK"
        Exit Function
    End If

    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Set db = p_db
    End If

    m_SQL = "SELECT IDEdicion FROM TbRiesgos WHERE IDRiesgo=" & CStr(CLng(p_IDRiesgo)) & ";"
    Set rcd = db.OpenRecordset(m_SQL)
    If Not rcd.EOF Then
        m_IDEdicion = CStr(Nz(rcd.Fields("IDEdicion"), ""))
    End If
    rcd.Close
    Set rcd = Nothing

    If m_IDEdicion <> "" Then
        m_ErrorLocal = ""
        Call CachePublicabilidad_InvalidarEdicion(m_IDEdicion, db, m_ErrorLocal)
        If m_ErrorLocal <> "" Then
            p_Error = m_ErrorLocal
            Err.Raise 1000
        End If
    End If

    CachePublicabilidad_InvalidarRiesgo = "OK"
    Exit Function
errores:
    On Error Resume Next
    If Not rcd Is Nothing Then
        rcd.Close
    End If
    On Error GoTo 0
    If Err.Number <> 1000 Then
        p_Error = "El método CachePublicabilidad_InvalidarRiesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Function CachePublicabilidad_ChecksJsonDesdeChecks( _
                                                        ByVal p_Checks As Scripting.Dictionary, _
                                                        Optional ByRef p_Error As String _
                                                        ) As String
    Dim m_ChecksArr As Collection
    Dim m_Key As Variant
    Dim m_Chk As Scripting.Dictionary
    On Error GoTo errores
    p_Error = ""

    Set m_ChecksArr = New Collection
    If Not p_Checks Is Nothing Then
        For Each m_Key In p_Checks
            Set m_Chk = p_Checks(m_Key)
            m_ChecksArr.Add ConstruirCheckJSON(m_Chk)
        Next
    End If

    CachePublicabilidad_ChecksJsonDesdeChecks = ConvertToJson(m_ChecksArr, 0)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CachePublicabilidad_ChecksJsonDesdeChecks ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Function CachePublicabilidad_InsertFila( _
                                                ByVal p_rcd As DAO.Recordset, _
                                                ByVal p_IDEdicion As Long, _
                                                ByVal p_Tipo As String, _
                                                ByVal p_IDRiesgo As Long, _
                                                ByVal p_Publicable As Boolean, _
                                                ByVal p_Veredicto As Variant, _
                                                ByVal p_CodigoRiesgo As Variant, _
                                                ByVal p_Descripcion As Variant, _
                                                ByVal p_ChecksJson As Variant, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim m_UsarCache As String
    On Error GoTo errores
    p_Error = ""

    m_UsarCache = "No"
    On Error Resume Next
    m_UsarCache = CStr(Application.TempVars("Publicabilidad_Usar_Cache"))
    On Error GoTo errores
    If UCase$(Trim$(m_UsarCache)) = "NO" Then
        CachePublicabilidad_InsertFila = "OK"
        Exit Function
    End If

    With p_rcd
        .AddNew
            .Fields("IDEdicion") = p_IDEdicion
            .Fields("Tipo") = p_Tipo
            .Fields("IDRiesgo") = p_IDRiesgo
            .Fields("Publicable") = p_Publicable
            If IsNull(p_Veredicto) Or IsEmpty(p_Veredicto) Then
                .Fields("Veredicto") = Null
            Else
                .Fields("Veredicto") = p_Veredicto
            End If
            If IsNull(p_CodigoRiesgo) Or IsEmpty(p_CodigoRiesgo) Then
                .Fields("CodigoRiesgo") = Null
            Else
                .Fields("CodigoRiesgo") = p_CodigoRiesgo
            End If
            If IsNull(p_Descripcion) Or IsEmpty(p_Descripcion) Then
                .Fields("Descripcion") = Null
            Else
                .Fields("Descripcion") = p_Descripcion
            End If
            If IsNull(p_ChecksJson) Or IsEmpty(p_ChecksJson) Then
                .Fields("ChecksJson") = Null
            Else
                .Fields("ChecksJson") = p_ChecksJson
            End If
            .Fields("AlgoritmoVersion") = CACHE_PUBLICABILIDAD_ALGORITMO_VERSION
            .Fields("UpdatedAt") = Now()
        .Update
    End With

    CachePublicabilidad_InsertFila = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CachePublicabilidad_InsertFila ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CachePublicabilidad_RecalcularEdicionYResetear( _
                                                            ByVal p_Edicion As Edicion, _
                                                            Optional p_db As DAO.Database = Nothing, _
                                                            Optional ByRef p_Error As String _
                                                            ) As String
    Dim m_ErrorLocal As String
    Dim m_UsarCache As String
    On Error GoTo errores
    p_Error = ""

    m_UsarCache = "No"
    On Error Resume Next
    m_UsarCache = CStr(Application.TempVars("Publicabilidad_Usar_Cache"))
    On Error GoTo errores
    If UCase$(Trim$(m_UsarCache)) = "NO" Then
        p_Edicion.PublicableResetear
        CachePublicabilidad_RecalcularEdicionYResetear = "OK"
        Exit Function
    End If

    m_ErrorLocal = ""
    If CachePublicabilidad_RecalcularEdicion(p_Edicion, p_db, m_ErrorLocal) = EnumSiNo.No Then
        If m_ErrorLocal <> "" Then
            p_Error = m_ErrorLocal
            Err.Raise 1000
        End If
    End If

    p_Edicion.PublicableResetear
    CachePublicabilidad_RecalcularEdicionYResetear = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CachePublicabilidad_RecalcularEdicionYResetear ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CachePublicabilidad_RecalcularEdicion( _
                                                    ByVal p_Edicion As Edicion, _
                                                    Optional p_db As DAO.Database = Nothing, _
                                                    Optional ByRef p_Error As String _
                                                    ) As EnumSiNo
    Dim db As DAO.Database
    Dim wksLocal As DAO.Workspace
    Dim rcdCache As DAO.Recordset
    Dim m_TransIniciada As Boolean
    Dim m_ColRiesgos As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Riesgo As riesgo
    Dim m_Datos As tPublicabilidadRiesgoDatos
    Dim m_Checks As Scripting.Dictionary
    Dim m_Veredicto As EnumPublicabilidadVeredicto
    Dim m_ErrorLocal As String
    Dim m_EdicionChecks As Scripting.Dictionary
    Dim m_EdicionPublicable As Boolean
    Dim m_EdicionPublicableFinal As Boolean
    Dim m_EdicionChecksJson As String
    Dim m_RiesgoChecksJson As String
    Dim m_SQL As String
    Dim m_EsTransaccionLocal As Boolean
    Dim m_UsarCache As String

    On Error GoTo errores
    p_Error = ""

    If p_Edicion Is Nothing Then
        p_Error = "Se ha de indicar la edición"
        Err.Raise 1000
    End If
    If p_Edicion.IDEdicion = "" Or Not IsNumeric(p_Edicion.IDEdicion) Then
        p_Error = "No se ha podido determinar el IDEdición"
        Err.Raise 1000
    End If

    m_UsarCache = "No"
    On Error Resume Next
    m_UsarCache = CStr(Application.TempVars("Publicabilidad_Usar_Cache"))
    On Error GoTo errores
    If UCase$(Trim$(m_UsarCache)) = "NO" Then
        CachePublicabilidad_RecalcularEdicion = EnumSiNo.No
        Exit Function
    End If

    m_ErrorLocal = ""
    Call CachePublicabilidad_AsegurarSchema(m_ErrorLocal)
    If m_ErrorLocal <> "" Then
        p_Error = m_ErrorLocal
        Err.Raise 1000
    End If

    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set wksLocal = DBEngine.Workspaces(0)
        wksLocal.BeginTrans
        m_TransIniciada = True
        m_EsTransaccionLocal = True
    Else
        Set db = p_db
        m_EsTransaccionLocal = False
        m_TransIniciada = False
    End If

    m_SQL = "DELETE FROM " & CACHE_PUBLICABILIDAD_TABLA & " WHERE IDEdicion=" & CStr(CLng(p_Edicion.IDEdicion)) & ";"
    db.Execute m_SQL

    Set rcdCache = db.OpenRecordset(CACHE_PUBLICABILIDAD_TABLA)

    m_ErrorLocal = ""
    If EvaluarPublicabilidadEdicion(p_Edicion, m_EdicionChecks, m_EdicionPublicable, db, m_ErrorLocal) = EnumSiNo.No Then
        p_Error = m_ErrorLocal
        Err.Raise 1000
    End If
    m_EdicionPublicableFinal = m_EdicionPublicable

    m_ErrorLocal = ""
    m_EdicionChecksJson = CachePublicabilidad_ChecksJsonDesdeChecks(m_EdicionChecks, m_ErrorLocal)
    If m_ErrorLocal <> "" Then
        p_Error = m_ErrorLocal
        Err.Raise 1000
    End If

    m_ErrorLocal = ""
    Set m_ColRiesgos = getRiesgosPorEdicionPub(p_Edicion.IDEdicion, db, EnumSiNo.Sí, m_ErrorLocal)
    If m_ErrorLocal <> "" Then
        p_Error = m_ErrorLocal
        Err.Raise 1000
    End If

    If Not m_ColRiesgos Is Nothing Then
        For Each m_ID In m_ColRiesgos
            Set m_Riesgo = m_ColRiesgos(m_ID)

            m_ErrorLocal = ""
            If ConstruirDatosPublicabilidadRiesgo(m_Riesgo, m_Datos, db, m_ErrorLocal) = EnumSiNo.No Then
                p_Error = "Error al evaluar publicabilidad (" & m_Riesgo.CodigoRiesgo & "): " & m_ErrorLocal
                Err.Raise 1000
            End If

            m_ErrorLocal = ""
            Call EvaluarPublicabilidadRiesgo(m_Datos, m_Checks, m_Veredicto, m_ErrorLocal)
            If m_ErrorLocal <> "" Then
                p_Error = "Error al evaluar publicabilidad (" & m_Riesgo.CodigoRiesgo & "): " & m_ErrorLocal
                Err.Raise 1000
            End If

            m_ErrorLocal = ""
            m_RiesgoChecksJson = CachePublicabilidad_ChecksJsonDesdeChecks(m_Checks, m_ErrorLocal)
            If m_ErrorLocal <> "" Then
                p_Error = m_ErrorLocal
                Err.Raise 1000
            End If

            If m_Veredicto = EnumPublicabilidadVeredicto.NoPublicable Then
                m_EdicionPublicableFinal = False
            End If

            m_ErrorLocal = ""
            Call CachePublicabilidad_InsertFila( _
                                                rcdCache, _
                                                CLng(p_Edicion.IDEdicion), _
                                                CACHE_PUBLICABILIDAD_TIPO_RIESGO, _
                                                CLng(m_Riesgo.IDRiesgo), _
                                                (m_Veredicto <> EnumPublicabilidadVeredicto.NoPublicable), _
                                                CLng(m_Veredicto), _
                                                m_Riesgo.CodigoRiesgo, _
                                                m_Riesgo.DescripcionParaLista, _
                                                m_RiesgoChecksJson, _
                                                m_ErrorLocal _
                                                )
            If m_ErrorLocal <> "" Then
                p_Error = m_ErrorLocal
                Err.Raise 1000
            End If

            Set m_Checks = Nothing
            Set m_Riesgo = Nothing
        Next
    End If

    m_ErrorLocal = ""
    Call CachePublicabilidad_InsertFila(rcdCache, CLng(p_Edicion.IDEdicion), CACHE_PUBLICABILIDAD_TIPO_EDICION, 0, m_EdicionPublicableFinal, Null, Null, Null, m_EdicionChecksJson, m_ErrorLocal)
    If m_ErrorLocal <> "" Then
        p_Error = m_ErrorLocal
        Err.Raise 1000
    End If

    rcdCache.Close
    Set rcdCache = Nothing

    If m_EsTransaccionLocal Then
        wksLocal.CommitTrans
    End If
    CachePublicabilidad_RecalcularEdicion = EnumSiNo.Sí
    Exit Function

errores:
    On Error Resume Next
    If Not rcdCache Is Nothing Then
        rcdCache.Close
    End If
    If m_EsTransaccionLocal And Not wksLocal Is Nothing Then
        If m_TransIniciada Then
            wksLocal.Rollback
        End If
    End If
    On Error GoTo 0

    CachePublicabilidad_RecalcularEdicion = EnumSiNo.No
    If Err.Number <> 1000 Then
        p_Error = "El método CachePublicabilidad_RecalcularEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CachePublicabilidad_LeerEdicionPublicable( _
                                                        ByVal p_IDEdicion As String, _
                                                        ByRef p_Publicable As Boolean, _
                                                        Optional ByRef p_Error As String _
                                                        ) As EnumSiNo
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim m_SQL As String
    Dim m_ErrorLocal As String
    Dim m_UsarCache As String

    On Error GoTo errores
    p_Error = ""

    m_UsarCache = "No"
    On Error Resume Next
    m_UsarCache = CStr(Application.TempVars("Publicabilidad_Usar_Cache"))
    On Error GoTo errores
    If UCase$(Trim$(m_UsarCache)) = "NO" Then
        CachePublicabilidad_LeerEdicionPublicable = EnumSiNo.No
        Exit Function
    End If

    If p_IDEdicion = "" Or Not IsNumeric(p_IDEdicion) Then
        p_Error = "No se ha podido determinar el IDEdición"
        Err.Raise 1000
    End If
    m_ErrorLocal = ""
    Call CachePublicabilidad_AsegurarSchema(m_ErrorLocal)
    If m_ErrorLocal <> "" Then
        p_Error = m_ErrorLocal
        Err.Raise 1000
    End If

    Set db = getdb(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    m_SQL = "SELECT Publicable " & _
            "FROM " & CACHE_PUBLICABILIDAD_TABLA & " " & _
            "WHERE IDEdicion=" & CStr(CLng(p_IDEdicion)) & " " & _
            "AND Tipo=" & CachePublicabilidad_SqlText(CACHE_PUBLICABILIDAD_TIPO_EDICION) & " " & _
            "AND IDRiesgo=0 " & _
            "AND AlgoritmoVersion=" & CStr(CACHE_PUBLICABILIDAD_ALGORITMO_VERSION) & ";"

    Set rcd = db.OpenRecordset(m_SQL)
    If rcd.EOF Then
        rcd.Close
        Set rcd = Nothing
        CachePublicabilidad_LeerEdicionPublicable = EnumSiNo.No
        Exit Function
    End If

    p_Publicable = CBool(Nz(rcd.Fields("Publicable"), False))
    rcd.Close
    Set rcd = Nothing

    CachePublicabilidad_LeerEdicionPublicable = EnumSiNo.Sí
    Exit Function
errores:
    CachePublicabilidad_LeerEdicionPublicable = EnumSiNo.No
    If Err.Number <> 1000 Then
        p_Error = "El método CachePublicabilidad_LeerEdicionPublicable ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Function CachePublicabilidad_CargarInformeDesdeCache( _
                                                            ByVal p_IDEdicion As String, _
                                                            ByRef p_RiesgosArr As Collection, _
                                                            ByRef p_Total As Long, _
                                                            ByRef p_Pub As Long, _
                                                            ByRef p_NoPub As Long, _
                                                            ByRef p_NoAplica As Long, _
                                                            ByRef p_EdicionPublicable As Boolean, _
                                                            ByRef p_EdicionChecksArr As Collection, _
                                                            Optional ByRef p_Error As String _
                                                            ) As EnumSiNo
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim rcdEd As DAO.Recordset
    Dim m_SQL As String
    Dim m_EdSQL As String
    Dim m_ChecksJson As String
    Dim m_ChecksObj As Object
    Dim m_Veredicto As EnumPublicabilidadVeredicto
    Dim m_R As Scripting.Dictionary
    Dim m_ErrorLocal As String
    Dim m_UsarCache As String

    On Error GoTo errores
    p_Error = ""

    m_UsarCache = "No"
    On Error Resume Next
    m_UsarCache = CStr(Application.TempVars("Publicabilidad_Usar_Cache"))
    On Error GoTo errores
    If UCase$(Trim$(m_UsarCache)) = "NO" Then
        CachePublicabilidad_CargarInformeDesdeCache = EnumSiNo.No
        Exit Function
    End If

    If p_IDEdicion = "" Or Not IsNumeric(p_IDEdicion) Then
        p_Error = "No se ha podido determinar el IDEdición"
        Err.Raise 1000
    End If


    m_ErrorLocal = ""
    Call CachePublicabilidad_AsegurarSchema(m_ErrorLocal)
    If m_ErrorLocal <> "" Then
        p_Error = m_ErrorLocal
        Err.Raise 1000
    End If

    Set db = getdb(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    m_EdSQL = "SELECT Publicable, ChecksJson " & _
              "FROM " & CACHE_PUBLICABILIDAD_TABLA & " " & _
              "WHERE IDEdicion=" & CStr(CLng(p_IDEdicion)) & " " & _
              "AND Tipo=" & CachePublicabilidad_SqlText(CACHE_PUBLICABILIDAD_TIPO_EDICION) & " " & _
              "AND IDRiesgo=0 " & _
              "AND AlgoritmoVersion=" & CStr(CACHE_PUBLICABILIDAD_ALGORITMO_VERSION) & ";"

    Set rcdEd = db.OpenRecordset(m_EdSQL)
    If rcdEd.EOF Then
        rcdEd.Close
        Set rcdEd = Nothing
        CachePublicabilidad_CargarInformeDesdeCache = EnumSiNo.No
        Exit Function
    End If

    p_EdicionPublicable = CBool(Nz(rcdEd.Fields("Publicable"), False))
    Set p_EdicionChecksArr = New Collection
    If Not IsNull(rcdEd.Fields("ChecksJson")) Then
        m_ChecksJson = CStr(Nz(rcdEd.Fields("ChecksJson"), ""))
        If m_ChecksJson <> "" Then
            Set m_ChecksObj = ParseJson(m_ChecksJson)
            If TypeName(m_ChecksObj) = "Collection" Then
                Set p_EdicionChecksArr = m_ChecksObj
            End If
        End If
    End If

    rcdEd.Close
    Set rcdEd = Nothing

    m_SQL = "SELECT IDRiesgo, CodigoRiesgo, Descripcion, Veredicto, ChecksJson " & _
            "FROM " & CACHE_PUBLICABILIDAD_TABLA & " " & _
            "WHERE IDEdicion=" & CStr(CLng(p_IDEdicion)) & " " & _
            "AND Tipo=" & CachePublicabilidad_SqlText(CACHE_PUBLICABILIDAD_TIPO_RIESGO) & " " & _
            "AND AlgoritmoVersion=" & CStr(CACHE_PUBLICABILIDAD_ALGORITMO_VERSION) & " " & _
            "ORDER BY CodigoRiesgo;"

    Set rcd = db.OpenRecordset(m_SQL)
    If rcd.EOF Then
        rcd.Close
        Set rcd = Nothing
        CachePublicabilidad_CargarInformeDesdeCache = EnumSiNo.No
        Exit Function
    End If

    Set p_RiesgosArr = New Collection
    p_Total = 0
    p_Pub = 0
    p_NoPub = 0
    p_NoAplica = 0

    Do While Not rcd.EOF
        p_Total = p_Total + 1

        If IsNull(rcd.Fields("Veredicto")) Then
            m_Veredicto = EnumPublicabilidadVeredicto.NoAplica
        Else
            m_Veredicto = CLng(rcd.Fields("Veredicto"))
        End If

        Select Case m_Veredicto
            Case EnumPublicabilidadVeredicto.Publicable
                p_Pub = p_Pub + 1
            Case EnumPublicabilidadVeredicto.NoPublicable
                p_NoPub = p_NoPub + 1
            Case Else
                p_NoAplica = p_NoAplica + 1
        End Select

        Set m_R = New Scripting.Dictionary
        m_R.CompareMode = TextCompare
        m_R.Add "idRiesgo", CLng(Nz(rcd.Fields("IDRiesgo"), 0))
        m_R.Add "codigo", CStr(Nz(rcd.Fields("CodigoRiesgo"), ""))
        m_R.Add "descripcion", CStr(Nz(rcd.Fields("Descripcion"), ""))
        m_R.Add "veredicto", TextoVeredictoPublicabilidad(m_Veredicto)

        Set m_ChecksObj = Nothing
        Set m_ChecksObj = New Collection
        If Not IsNull(rcd.Fields("ChecksJson")) Then
            m_ChecksJson = CStr(Nz(rcd.Fields("ChecksJson"), ""))
            If m_ChecksJson <> "" Then
                Set m_ChecksObj = ParseJson(m_ChecksJson)
            End If
        End If
        m_R.Add "checks", m_ChecksObj

        p_RiesgosArr.Add m_R
        rcd.MoveNext
    Loop

    rcd.Close
    Set rcd = Nothing

    CachePublicabilidad_CargarInformeDesdeCache = EnumSiNo.Sí
    Exit Function
errores:
    CachePublicabilidad_CargarInformeDesdeCache = EnumSiNo.No
    If Err.Number <> 1000 Then
        p_Error = "El método CachePublicabilidad_CargarInformeDesdeCache ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub AgregarCheckPublicabilidad( _
                                        ByRef p_Checks As Scripting.Dictionary, _
                                        ByRef p_Index As Long, _
                                        ByVal p_Id As String, _
                                        ByVal p_Texto As String, _
                                        ByVal p_Estado As EnumPublicabilidadCheckEstado, _
                                        Optional ByVal p_Detalle As String = "" _
                                        )
    Dim m_Check As Scripting.Dictionary

    Set m_Check = New Scripting.Dictionary
    m_Check.CompareMode = TextCompare
    m_Check.Add "id", p_Id
    m_Check.Add "texto", p_Texto
    m_Check.Add "estado", p_Estado
    If p_Detalle <> "" Then
        m_Check.Add "detalle", p_Detalle
    End If

    p_Checks.Add p_Id, m_Check
    p_Index = p_Index + 1
End Sub

Public Function EvaluarPublicabilidadEdicion( _
                                            ByVal p_Edicion As Edicion, _
                                            ByRef p_Checks As Scripting.Dictionary, _
                                            ByRef p_EdicionPublicable As Boolean, _
                                            Optional p_db As DAO.Database = Nothing, _
                                            Optional ByRef p_Error As String _
                                            ) As EnumSiNo
    Dim m_Index As Long
    Dim m_EnUTE As String
    Dim m_NecesitaEvidenciasSum As EnumSiNo
    Dim m_DetalleSuministradores As String

    On Error GoTo errores
    p_Error = ""

    If p_Edicion Is Nothing Then
        p_Error = "Se ha de indicar la edición"
        Err.Raise 1000
    End If

    Set p_Checks = New Scripting.Dictionary
    p_Checks.CompareMode = TextCompare
    m_Index = 1
    p_EdicionPublicable = True

    m_EnUTE = UCase$(Left$(Nz(p_Edicion.Proyecto.EnUTE, ""), 1))
    If m_EnUTE = "S" Then
        If p_Edicion.TieneAnexoEvidenciaUTE = EnumSiNo.Sí Then
            AgregarCheckPublicabilidad p_Checks, m_Index, "ute_evidence", "Evidencia de acuerdo con la otra empresa (UTE)", EnumPublicabilidadCheckEstado.Cumple
        Else
            AgregarCheckPublicabilidad p_Checks, m_Index, "ute_evidence", "Evidencia de acuerdo con la otra empresa (UTE)", EnumPublicabilidadCheckEstado.NoCumple, "Falta evidencia de acuerdo con la otra empresa"
            p_EdicionPublicable = False
        End If
    Else
        AgregarCheckPublicabilidad p_Checks, m_Index, "ute_evidence", "Evidencia de acuerdo con la otra empresa (UTE)", EnumPublicabilidadCheckEstado.NoAplica
    End If

    m_NecesitaEvidenciasSum = p_Edicion.EvidenciasSuministradoresRequeridas
    If m_NecesitaEvidenciasSum = EnumSiNo.Sí Then
        If p_Edicion.EvidenciasSuministradoresCompletadas = EnumSiNo.Sí Then
            AgregarCheckPublicabilidad p_Checks, m_Index, "supplier_evidence", "Evidencia de acuerdo con todos los suministradores", EnumPublicabilidadCheckEstado.Cumple
        Else
            m_DetalleSuministradores = SuministradoresHelper.DetalleSuministradoresSinEvidenciaEnEdicion(p_Edicion.IDEdicion, EnumSiNo.No, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_DetalleSuministradores = "" Then
                m_DetalleSuministradores = "Falta evidencia de acuerdo con suministradores"
            End If
            AgregarCheckPublicabilidad p_Checks, m_Index, "supplier_evidence", "Evidencia de acuerdo con todos los suministradores", EnumPublicabilidadCheckEstado.NoCumple, m_DetalleSuministradores
            p_EdicionPublicable = False
        End If
    Else
        AgregarCheckPublicabilidad p_Checks, m_Index, "supplier_evidence", "Evidencia de acuerdo con todos los suministradores", EnumPublicabilidadCheckEstado.NoAplica
    End If

    EvaluarPublicabilidadEdicion = EnumSiNo.Sí
    Exit Function

errores:
    EvaluarPublicabilidadEdicion = EnumSiNo.No
    If Err.Number <> 1000 Then
        p_Error = "Error en EvaluarPublicabilidadEdicion: " & Err.Description
    End If
End Function

Public Function CalcularPublicabilidadEdicion( _
                                            ByVal p_Edicion As Edicion, _
                                            ByRef p_EdicionPublicable As Boolean, _
                                            Optional p_db As DAO.Database = Nothing, _
                                            Optional ByRef p_Error As String _
                                            ) As EnumSiNo
    Dim m_ColRiesgos As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Riesgo As riesgo
    Dim m_Datos As tPublicabilidadRiesgoDatos
    Dim m_Checks As Scripting.Dictionary
    Dim m_Veredicto As EnumPublicabilidadVeredicto
    Dim m_ErrorLocal As String
    Dim m_EdicionChecks As Scripting.Dictionary
    Dim m_Publicable As EnumSiNo
    On Error GoTo errores
    p_Error = ""

    If p_Edicion Is Nothing Then
        p_Error = "Se ha de indicar la edición"
        Err.Raise 1000
    End If
    If p_Edicion.IDEdicion = "" Then
        p_Error = "No se ha podido determinar el IDEdicion"
        Err.Raise 1000
    End If

    m_ErrorLocal = ""
    If EvaluarPublicabilidadEdicion(p_Edicion, m_EdicionChecks, p_EdicionPublicable, p_db, m_ErrorLocal) = EnumSiNo.No Then
        p_Error = m_ErrorLocal
        Err.Raise 1000
    End If

    m_ErrorLocal = ""
    Set m_ColRiesgos = getRiesgosPorEdicionPub(p_Edicion.IDEdicion, p_db, EnumSiNo.Sí, m_ErrorLocal)
    If m_ErrorLocal <> "" Then
        p_Error = m_ErrorLocal
        Err.Raise 1000
    End If

    If Not m_ColRiesgos Is Nothing Then
        For Each m_ID In m_ColRiesgos
            Set m_Riesgo = m_ColRiesgos(m_ID)

            m_ErrorLocal = ""
            If ConstruirDatosPublicabilidadRiesgo(m_Riesgo, m_Datos, p_db, m_ErrorLocal) = EnumSiNo.No Then
                p_Error = "Error al evaluar publicabilidad (" & m_Riesgo.CodigoRiesgo & "): " & m_ErrorLocal
                Err.Raise 1000
            End If

            m_ErrorLocal = ""
            Call EvaluarPublicabilidadRiesgo(m_Datos, m_Checks, m_Veredicto, m_ErrorLocal)
            If m_ErrorLocal <> "" Then
                p_Error = "Error al evaluar publicabilidad (" & m_Riesgo.CodigoRiesgo & "): " & m_ErrorLocal
                Err.Raise 1000
            End If

            If m_Veredicto = EnumPublicabilidadVeredicto.NoPublicable Then
                p_EdicionPublicable = False
            End If

            Set m_Checks = Nothing
            Set m_Riesgo = Nothing
        Next
    End If

    CalcularPublicabilidadEdicion = EnumSiNo.Sí
    Exit Function

errores:
    CalcularPublicabilidadEdicion = EnumSiNo.No
    If Err.Number <> 1000 Then
        p_Error = "Error en CalcularPublicabilidadEdicion: " & Err.Description
    End If
End Function

Public Function GenerarInformePublicabilidadEdicionInteractivoHTML( _
                                                                    ByVal p_Edicion As Edicion, _
                                                                    Optional ByVal p_hWnd As Long, _
                                                                    Optional ByRef p_Error As String, _
                                                                    Optional ByVal p_GenerarPDF As Boolean = False _
                                                                    ) As String

    Dim m_ColRiesgos As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Riesgo As riesgo

    Dim m_Datos As tPublicabilidadRiesgoDatos
    Dim m_Checks As Scripting.Dictionary
    Dim m_Veredicto As EnumPublicabilidadVeredicto
    Dim m_ErrorLocal As String
    Dim m_UsarCache As String

    Dim m_Root As Scripting.Dictionary
    Dim m_EdicionJSON As Scripting.Dictionary
    Dim m_Resumen As Scripting.Dictionary
    Dim m_RiesgosArr As Collection

    Dim m_Total As Long
    Dim m_Pub As Long
    Dim m_NoPub As Long
    Dim m_NoAplica As Long
    Dim m_EdicionPublicable As Boolean
    Dim m_EdicionChecks As Scripting.Dictionary
    Dim m_EdicionChecksArr As Collection
    Dim m_Key As Variant
    Dim m_Chk As Scripting.Dictionary

    Dim m_JSON As String
    Dim m_HTML As String
    Dim m_URL As String
    Dim m_URL_Apertura As String
    Dim m_URL_PDF As String

    On Error GoTo errores
    p_Error = ""

    If p_Edicion Is Nothing Then
        p_Error = "Se ha de indicar la edición"
        Err.Raise 1000
    End If
    If p_Edicion.IDEdicion = "" Then
        p_Error = "No se ha podido determinar el IDEdición"
        Err.Raise 1000
    End If

    m_UsarCache = "No"
    On Error Resume Next
    m_UsarCache = CStr(Application.TempVars("Publicabilidad_Usar_Cache"))
    On Error GoTo errores

    If UCase$(Trim$(m_UsarCache)) = "NO" Then
        m_ErrorLocal = ""
        Set m_ColRiesgos = getRiesgosPorEdicionPub(p_Edicion.IDEdicion, getdb(), EnumSiNo.Sí, m_ErrorLocal)
        If m_ErrorLocal <> "" Then
            p_Error = m_ErrorLocal
            Err.Raise 1000
        End If
        If m_ColRiesgos Is Nothing Then
            p_Error = "La edición no tiene riesgos"
            Err.Raise 1000
        End If

        Set m_RiesgosArr = New Collection
        m_Total = 0
        m_Pub = 0
        m_NoPub = 0
        m_NoAplica = 0
        m_EdicionPublicable = True

        m_ErrorLocal = ""
        If EvaluarPublicabilidadEdicion(p_Edicion, m_EdicionChecks, m_EdicionPublicable, , m_ErrorLocal) = EnumSiNo.No Then
            p_Error = m_ErrorLocal
            Err.Raise 1000
        End If

        For Each m_ID In m_ColRiesgos
            Set m_Riesgo = m_ColRiesgos(m_ID)
            m_Total = m_Total + 1

            m_ErrorLocal = ""
            If ConstruirDatosPublicabilidadRiesgo(m_Riesgo, m_Datos, , m_ErrorLocal) = EnumSiNo.No Then
                p_Error = "Error al evaluar publicabilidad (" & m_Riesgo.CodigoRiesgo & "): " & m_ErrorLocal
                Err.Raise 1000
            End If

            m_ErrorLocal = ""
            Call EvaluarPublicabilidadRiesgo(m_Datos, m_Checks, m_Veredicto, m_ErrorLocal)
            If m_ErrorLocal <> "" Then
                p_Error = "Error al evaluar publicabilidad (" & m_Riesgo.CodigoRiesgo & "): " & m_ErrorLocal
                Err.Raise 1000
            End If

            Select Case m_Veredicto
                Case EnumPublicabilidadVeredicto.Publicable
                    m_Pub = m_Pub + 1
                Case EnumPublicabilidadVeredicto.NoPublicable
                    m_NoPub = m_NoPub + 1
                    m_EdicionPublicable = False
                Case Else
                    m_NoAplica = m_NoAplica + 1
            End Select

            m_RiesgosArr.Add ConstruirRiesgoJSON_Publicabilidad(m_Riesgo, m_Veredicto, m_Checks)

            Set m_Checks = Nothing
            Set m_Riesgo = Nothing
        Next

        ' Prepend edition pseudo-risk
        Dim m_EdicionPseudo As Scripting.Dictionary
        Set m_EdicionPseudo = ConstruirPseudoRiesgoEdicion(m_EdicionChecksArr, m_EdicionPublicable)
        Dim m_RiesgosArrWithEdicion As Collection
        Set m_RiesgosArrWithEdicion = New Collection
        m_RiesgosArrWithEdicion.Add m_EdicionPseudo
        Dim m_RItem As Variant
        For Each m_RItem In m_RiesgosArr
            m_RiesgosArrWithEdicion.Add m_RItem
        Next
        Set m_RiesgosArr = m_RiesgosArrWithEdicion
    Else
        m_ErrorLocal = ""
        If CachePublicabilidad_CargarInformeDesdeCache( _
                                                    p_Edicion.IDEdicion, _
                                                    m_RiesgosArr, _
                                                    m_Total, _
                                                    m_Pub, _
                                                    m_NoPub, _
                                                    m_NoAplica, _
                                                    m_EdicionPublicable, _
                                                    m_EdicionChecksArr, _
                                                    m_ErrorLocal _
                                                    ) = EnumSiNo.No Then
            If m_ErrorLocal <> "" Then
                p_Error = m_ErrorLocal
                Err.Raise 1000
            End If

            m_ErrorLocal = ""
            If CachePublicabilidad_RecalcularEdicion(p_Edicion, , m_ErrorLocal) = EnumSiNo.No Then
                If m_ErrorLocal <> "" Then
                    p_Error = m_ErrorLocal
                    Err.Raise 1000
                End If
            End If

            m_ErrorLocal = ""
            If CachePublicabilidad_CargarInformeDesdeCache( _
                                                        p_Edicion.IDEdicion, _
                                                        m_RiesgosArr, _
                                                        m_Total, _
                                                        m_Pub, _
                                                        m_NoPub, _
                                                        m_NoAplica, _
                                                        m_EdicionPublicable, _
                                                        m_EdicionChecksArr, _
                                                        m_ErrorLocal _
                                                        ) = EnumSiNo.No Then
                If m_ErrorLocal <> "" Then
                    p_Error = m_ErrorLocal
                    Err.Raise 1000
                End If

                m_ErrorLocal = ""
                Set m_ColRiesgos = getRiesgosPorEdicionPub(p_Edicion.IDEdicion, getdb(), EnumSiNo.Sí, m_ErrorLocal)
                If m_ErrorLocal <> "" Then
                    p_Error = m_ErrorLocal
                    Err.Raise 1000
                End If
                If m_ColRiesgos Is Nothing Then
                    p_Error = "La edición no tiene riesgos"
                    Err.Raise 1000
                End If

                Set m_RiesgosArr = New Collection
                m_Total = 0
                m_Pub = 0
                m_NoPub = 0
                m_NoAplica = 0
                m_EdicionPublicable = True

                m_ErrorLocal = ""
                If EvaluarPublicabilidadEdicion(p_Edicion, m_EdicionChecks, m_EdicionPublicable, , m_ErrorLocal) = EnumSiNo.No Then
                    p_Error = m_ErrorLocal
                    Err.Raise 1000
                End If

                For Each m_ID In m_ColRiesgos
                    Set m_Riesgo = m_ColRiesgos(m_ID)
                    m_Total = m_Total + 1

                    m_ErrorLocal = ""
                    If ConstruirDatosPublicabilidadRiesgo(m_Riesgo, m_Datos, , m_ErrorLocal) = EnumSiNo.No Then
                        p_Error = "Error al evaluar publicabilidad (" & m_Riesgo.CodigoRiesgo & "): " & m_ErrorLocal
                        Err.Raise 1000
                    End If

                    m_ErrorLocal = ""
                    Call EvaluarPublicabilidadRiesgo(m_Datos, m_Checks, m_Veredicto, m_ErrorLocal)
                    If m_ErrorLocal <> "" Then
                        p_Error = "Error al evaluar publicabilidad (" & m_Riesgo.CodigoRiesgo & "): " & m_ErrorLocal
                        Err.Raise 1000
                    End If

                    Select Case m_Veredicto
                        Case EnumPublicabilidadVeredicto.Publicable
                            m_Pub = m_Pub + 1
                        Case EnumPublicabilidadVeredicto.NoPublicable
                            m_NoPub = m_NoPub + 1
                            m_EdicionPublicable = False
                        Case Else
                            m_NoAplica = m_NoAplica + 1
                    End Select

                    m_RiesgosArr.Add ConstruirRiesgoJSON_Publicabilidad(m_Riesgo, m_Veredicto, m_Checks)

                    Set m_Checks = Nothing
                    Set m_Riesgo = Nothing
                Next

                ' Prepend edition pseudo-risk (cache recalc path)
                Dim m_EdicionPseudo2 As Scripting.Dictionary
                Set m_EdicionPseudo2 = ConstruirPseudoRiesgoEdicion(m_EdicionChecksArr, m_EdicionPublicable)
                Dim m_RiesgosArrWithEdicion2 As Collection
                Set m_RiesgosArrWithEdicion2 = New Collection
                m_RiesgosArrWithEdicion2.Add m_EdicionPseudo2
                Dim m_RItem2 As Variant
                For Each m_RItem2 In m_RiesgosArr
                    m_RiesgosArrWithEdicion2.Add m_RItem2
                Next
                Set m_RiesgosArr = m_RiesgosArrWithEdicion2
            End If
        End If
    End If

    Set m_EdicionJSON = New Scripting.Dictionary
    m_EdicionJSON.CompareMode = TextCompare
    m_EdicionJSON.Add "idEdicion", p_Edicion.IDEdicion
    m_EdicionJSON.Add "nombre", p_Edicion.Edicion
    m_EdicionJSON.Add "proyecto", Nz(p_Edicion.Proyecto.NombreProyecto, Nz(p_Edicion.Proyecto.Proyecto, ""))

    Set m_Resumen = New Scripting.Dictionary
    m_Resumen.CompareMode = TextCompare
    m_Resumen.Add "total", m_Total
    m_Resumen.Add "publicables", m_Pub
    m_Resumen.Add "noPublicables", m_NoPub
    m_Resumen.Add "noAplica", m_NoAplica
    m_Resumen.Add "edicionPublicable", m_EdicionPublicable

    Set m_Root = New Scripting.Dictionary
    m_Root.CompareMode = TextCompare
    m_Root.Add "edicion", m_EdicionJSON
    m_Root.Add "resumen", m_Resumen
    If m_EdicionChecksArr Is Nothing Then
        Set m_EdicionChecksArr = New Collection
        If Not m_EdicionChecks Is Nothing Then
            For Each m_Key In m_EdicionChecks
                Set m_Chk = m_EdicionChecks(m_Key)
                m_EdicionChecksArr.Add ConstruirCheckJSON(m_Chk)
            Next
        End If
    End If
    m_Root.Add "edition_checks", m_EdicionChecksArr
    m_Root.Add "riesgos", m_RiesgosArr

    m_JSON = ConvertToJson(m_Root, 2)
    m_JSON = Replace(m_JSON, "</", "<\/")

    m_HTML = ConstruirTemplateInformePublicabilidadEdicion()
    m_HTML = Replace(m_HTML, "__DATA_JSON__", m_JSON)

    m_URL = GuardarInformePublicabilidadEdicion_UTF8(p_Edicion, m_HTML, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    m_URL_Apertura = m_URL
    If p_GenerarPDF Then
        m_URL_PDF = RutaPDFInformePublicabilidadDesdeHTML(m_URL)
        If ConvertirHTMLaPDFConEdge(m_URL, m_URL_PDF, p_Error) Then
            m_URL_Apertura = m_URL_PDF
        Else
            Err.Raise 1000
        End If
    End If

    If p_hWnd = 0 Then
        On Error Resume Next
        p_hWnd = Application.hWndAccessApp
        On Error GoTo errores
    End If

    Ejecutar p_hWnd, "open", m_URL_Apertura, "", "", 1
    GenerarInformePublicabilidadEdicionInteractivoHTML = m_URL_Apertura
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en GenerarInformePublicabilidadEdicionInteractivoHTML: " & Err.Description
    End If
End Function
Private Function RutaPDFInformePublicabilidadDesdeHTML(ByVal p_URLHTML As String) As String
    RutaPDFInformePublicabilidadDesdeHTML = Replace(p_URLHTML, ".html", ".pdf", 1, -1, vbTextCompare)
End Function

Private Function RutaSalidaInformePublicabilidad( _
                                                ByVal p_URLHTML As String, _
                                                ByVal p_GenerarPDF As Boolean _
                                                ) As String
    If p_GenerarPDF Then
        RutaSalidaInformePublicabilidad = RutaPDFInformePublicabilidadDesdeHTML(p_URLHTML)
    Else
        RutaSalidaInformePublicabilidad = p_URLHTML
    End If
End Function

Public Function PublicabilidadEdicion_RutaPDFDesdeHTMLParaTest(ByVal p_URLHTML As String) As String
    PublicabilidadEdicion_RutaPDFDesdeHTMLParaTest = RutaPDFInformePublicabilidadDesdeHTML(p_URLHTML)
End Function

Public Function PublicabilidadEdicion_RutaSalidaParaTest( _
                                                        ByVal p_URLHTML As String, _
                                                        ByVal p_GenerarPDF As Boolean _
                                                        ) As String
    PublicabilidadEdicion_RutaSalidaParaTest = RutaSalidaInformePublicabilidad(p_URLHTML, p_GenerarPDF)
End Function
Private Function ConstruirRiesgoJSON_Publicabilidad( _
                                                    ByVal p_Riesgo As riesgo, _
                                                    ByVal p_Veredicto As EnumPublicabilidadVeredicto, _
                                                    ByVal p_Checks As Scripting.Dictionary _
                                                    ) As Scripting.Dictionary

    Dim m_R As Scripting.Dictionary
    Dim m_ChecksArr As Collection
    Dim m_Key As Variant
    Dim m_Chk As Scripting.Dictionary

    Set m_R = New Scripting.Dictionary
    m_R.CompareMode = TextCompare
    m_R.Add "idRiesgo", p_Riesgo.IDRiesgo
    m_R.Add "codigo", p_Riesgo.CodigoRiesgo
    m_R.Add "descripcion", p_Riesgo.DescripcionParaLista
    m_R.Add "veredicto", TextoVeredictoPublicabilidad(p_Veredicto)

    Set m_ChecksArr = New Collection
    If Not p_Checks Is Nothing Then
        For Each m_Key In p_Checks
            Set m_Chk = p_Checks(m_Key)
            m_ChecksArr.Add ConstruirCheckJSON(m_Chk)
        Next
    End If

    m_R.Add "checks", m_ChecksArr
    Set ConstruirRiesgoJSON_Publicabilidad = m_R
End Function
Private Function ConstruirPseudoRiesgoEdicion( _
    ByVal p_ChecksArr As Collection, _
    ByVal p_EdicionPublicable As Boolean _
) As Scripting.Dictionary
    Dim m_R As Scripting.Dictionary
    Dim m_Chk As Scripting.Dictionary
    Dim m_Key As Variant

    Set m_R = New Scripting.Dictionary
    m_R.CompareMode = TextCompare
    m_R.Add "idRiesgo", "__EDICION__"
    m_R.Add "codigo", "__EDICION__"
    m_R.Add "descripcion", "Gestión de Riesgos"
    m_R.Add "veredicto", IIf(p_EdicionPublicable, "Publicable", "NoPublicable")

    Set m_R("checks") = New Collection
    If Not p_ChecksArr Is Nothing Then
        For Each m_Key In p_ChecksArr
            If TypeName(p_ChecksArr(m_Key)) = "Dictionary" Then
                Set m_Chk = p_ChecksArr(m_Key)
                m_R("checks").Add ConstruirCheckJSON(m_Chk)
            End If
        Next
    End If

    Set ConstruirPseudoRiesgoEdicion = m_R
End Function
Private Function ConstruirCheckJSON(ByVal p_Check As Scripting.Dictionary) As Scripting.Dictionary
    Dim m_O As Scripting.Dictionary
    Dim m_Detalle As String
    Dim m_Estado As EnumPublicabilidadCheckEstado

    Set m_O = New Scripting.Dictionary
    m_O.CompareMode = TextCompare

    m_O.Add "id", CStr(p_Check("id"))
    m_O.Add "texto", CStr(p_Check("texto"))
    m_Estado = p_Check("estado")
    m_O.Add "estado", TextoEstadoCheckPublicabilidad(m_Estado)
    m_Detalle = ""
    If p_Check.Exists("detalle") Then
        m_Detalle = CStr(p_Check("detalle"))
    End If
    m_O.Add "detalle", m_Detalle

    Set ConstruirCheckJSON = m_O
End Function

Private Function TextoVeredictoPublicabilidad(ByVal p_Veredicto As EnumPublicabilidadVeredicto) As String
    Select Case p_Veredicto
        Case EnumPublicabilidadVeredicto.Publicable
            TextoVeredictoPublicabilidad = "Publicable"
        Case EnumPublicabilidadVeredicto.NoPublicable
            TextoVeredictoPublicabilidad = "NoPublicable"
        Case Else
            TextoVeredictoPublicabilidad = "NoAplica"
    End Select
End Function

Private Function TextoEstadoCheckPublicabilidad(ByVal p_Estado As EnumPublicabilidadCheckEstado) As String
    Select Case p_Estado
        Case EnumPublicabilidadCheckEstado.Cumple
            TextoEstadoCheckPublicabilidad = "Cumple"
        Case EnumPublicabilidadCheckEstado.NoCumple
            TextoEstadoCheckPublicabilidad = "NoCumple"
        Case Else
            TextoEstadoCheckPublicabilidad = "NoAplica"
    End Select
End Function
Private Function GuardarInformePublicabilidadEdicion_UTF8( _
                                                        ByVal p_Edicion As Edicion, _
                                                        ByVal p_HTML As String, _
                                                        ByRef p_Error As String _
                                                        ) As String

    Dim m_Ruta As String
    Dim m_Stream As Object
    Dim m_Nombre As String

    On Error GoTo errores
    p_Error = ""

    m_Nombre = "Publicabilidad_Edicion_" & SanitizarNombreArchivo(Nz(p_Edicion.Proyecto.NombreProyecto, p_Edicion.Proyecto.Proyecto) & "_Ed_" & p_Edicion.Edicion)
    m_Ruta = Environ$("TEMP") & "\" & m_Nombre & ".html"

    Set m_Stream = CreateObject("ADODB.Stream")
    m_Stream.Type = 2
    m_Stream.Charset = "utf-8"
    m_Stream.Open
    m_Stream.WriteText p_HTML
    m_Stream.SaveToFile m_Ruta, 2
    m_Stream.Close

    GuardarInformePublicabilidadEdicion_UTF8 = m_Ruta
    Exit Function

errores:
    p_Error = "Error al guardar el HTML: " & Err.Description
End Function

Private Function SanitizarNombreArchivo(ByVal p_Texto As String) As String
    Dim m_T As String
    m_T = Nz(p_Texto, "")
    m_T = Replace(m_T, "/", "_")
    m_T = Replace(m_T, "\", "_")
    m_T = Replace(m_T, ":", "_")
    m_T = Replace(m_T, "*", "_")
    m_T = Replace(m_T, "?", "_")
    m_T = Replace(m_T, """", "_")
    m_T = Replace(m_T, "<", "_")
    m_T = Replace(m_T, ">", "_")
    m_T = Replace(m_T, "|", "_")
    SanitizarNombreArchivo = m_T
End Function

Private Function ConstruirTemplateInformePublicabilidadEdicion() As String
    Dim m_Lineas As Collection
    Dim m_Linea As Variant
    Dim m_HTML As String

    Set m_Lineas = New Collection

    m_Lineas.Add "<!DOCTYPE html>"
    m_Lineas.Add "<html lang='es'>"
    m_Lineas.Add "<head>"
    m_Lineas.Add "  <meta charset='UTF-8'>"
    m_Lineas.Add "  <meta name='viewport' content='width=device-width, initial-scale=1' />"
    m_Lineas.Add "  <title>Gestión de Riesgos - Publicabilidad</title>"
    m_Lineas.Add "  <style>"
    m_Lineas.Add "    :root { --tele-blue:#0066FF; --tele-white:#F2F4FF; --pure-white:#FFFFFF; --grey-9:#031A34; --grey-6:#58617A; --grey-2:#D1D5E4; }"
    m_Lineas.Add "    body { font-family:'Segoe UI', Arial, sans-serif; background-color:var(--tele-white); color:var(--grey-6); margin:0; }"
    m_Lineas.Add "    header { background-color:var(--tele-blue); color:white; padding:25px 0; border-bottom:4px solid #0055D4; }"
    m_Lineas.Add "    .container { max-width:1200px; margin:0 auto; padding:0 24px; }"
    m_Lineas.Add "    .header-container { display:flex; justify-content:space-between; align-items:center; gap:18px; }"
    m_Lineas.Add "    .header-text h1 { font-size:26px; margin:0; font-weight:800; }"
    m_Lineas.Add "    .header-text p { margin:0; opacity:0.85; font-size:14px; }"
    m_Lineas.Add "    .header-info { text-align:right; }"
    m_Lineas.Add "    .hdr-badge { display:inline-block; margin-top:8px; padding:7px 12px; border-radius:999px; font-weight:900; font-size:12px; letter-spacing:0.4px; text-transform:uppercase; border:1px solid rgba(255,255,255,0.45); }"
    m_Lineas.Add "    .hdr-badge.publicable { background:rgba(232,245,233,0.18); }"
    m_Lineas.Add "    .hdr-badge.no-publicable { background:rgba(255,235,238,0.18); }"
    m_Lineas.Add "    .main-container { margin-top:35px; padding-bottom:60px; }"
    m_Lineas.Add "    .card { background:white; padding:18px; border-radius:18px; border:1px solid var(--grey-2); }"
    m_Lineas.Add "    h2 { color:var(--grey-9); border-bottom:2px solid var(--tele-blue); padding-bottom:8px; margin:0 0 14px 0; }"
    m_Lineas.Add "    .verdict-card { display:flex; align-items:center; justify-content:space-between; gap:15px; margin-bottom:14px; }"
    m_Lineas.Add "    .verdict-badge { padding:8px 14px; border-radius:999px; font-weight:bold; font-size:12px; letter-spacing:0.5px; text-transform:uppercase; }"
    m_Lineas.Add "    .verdict-publicable { background:#E8F5E9; color:#1B5E20; }"
    m_Lineas.Add "    .verdict-no-publicable { background:#FFEBEE; color:#B71C1C; }"
    m_Lineas.Add "    .verdict-no-aplica { background:#ECEFF1; color:#455A64; }"
    m_Lineas.Add "    .checklist { display:grid; gap:12px; }"
    m_Lineas.Add "    .check-item { border:1px solid var(--grey-2); border-radius:12px; padding:12px 14px; display:flex; justify-content:space-between; gap:12px; align-items:flex-start; }"
    m_Lineas.Add "    .check-left { flex:1; }"
    m_Lineas.Add "    .check-text { font-weight:600; color:var(--grey-9); }"
    m_Lineas.Add "    .check-detail { margin-top:4px; font-size:12px; color:var(--grey-6); }"
    m_Lineas.Add "    .check-state { font-size:11px; font-weight:bold; padding:4px 8px; border-radius:999px; text-transform:uppercase; white-space:nowrap; }"
    m_Lineas.Add "    .state-cumple { background:#E8F5E9; color:#1B5E20; }"
    m_Lineas.Add "    .state-no-cumple { background:#FFEBEE; color:#B71C1C; }"
    m_Lineas.Add "    .state-no-aplica { background:#ECEFF1; color:#455A64; }"
    m_Lineas.Add "    .layout { display:grid; grid-template-columns: 360px 1fr; gap:18px; align-items:start; }"
    m_Lineas.Add "    @media (max-width: 980px) { .layout { grid-template-columns:1fr; } }"
    m_Lineas.Add "    .sidebar { position:sticky; top:16px; }"
    m_Lineas.Add "    .sidebar-top { display:grid; gap:10px; margin-bottom:12px; }"
    m_Lineas.Add "    .search { width:100%; box-sizing:border-box; padding:10px 12px; border-radius:12px; border:1px solid var(--grey-2); outline:none; }"
    m_Lineas.Add "    .filters { display:flex; flex-wrap:wrap; gap:8px; }"
    m_Lineas.Add "    .chip { border:1px solid var(--grey-2); background:#F3F5FC; padding:8px 10px; border-radius:999px; cursor:pointer; font-weight:700; font-size:12px; color:var(--grey-6); }"
    m_Lineas.Add "    .chip.active { background:var(--pure-white); color:var(--tele-blue); border-color:var(--tele-blue); }"
    m_Lineas.Add "    .risk-list { display:grid; gap:10px; }"
    m_Lineas.Add "    .risk-item { border:1px solid var(--grey-2); border-radius:14px; padding:12px; cursor:pointer; background:white; display:grid; gap:6px; }"
    m_Lineas.Add "    .risk-item.active { outline:2px solid var(--tele-blue); border-color:transparent; }"
    m_Lineas.Add "    .risk-item.no-publicable { border-left:6px solid #B71C1C; background:#FFF7F8; }"
    m_Lineas.Add "    .risk-item.publicable { border-left:6px solid #1B5E20; }"
    m_Lineas.Add "    .risk-item.no-aplica { border-left:6px solid #455A64; background:#FAFBFD; }"
    m_Lineas.Add "    .risk-head { display:flex; justify-content:space-between; gap:10px; align-items:center; }"
    m_Lineas.Add "    .risk-code { font-weight:800; color:var(--grey-9); }"
    m_Lineas.Add "    .risk-desc { font-size:12px; color:var(--grey-6); line-height:1.35; }"
    m_Lineas.Add "    .mini-badge { padding:4px 8px; border-radius:999px; font-weight:800; font-size:10px; text-transform:uppercase; letter-spacing:0.4px; }"
    m_Lineas.Add "    .mini-publicable { background:#E8F5E9; color:#1B5E20; }"
    m_Lineas.Add "    .mini-no-publicable { background:#FFEBEE; color:#B71C1C; }"
    m_Lineas.Add "    .mini-no-aplica { background:#ECEFF1; color:#455A64; }"
    m_Lineas.Add "    .list-links { display:flex; flex-wrap:wrap; gap:8px; margin-top:10px; }"
    m_Lineas.Add "    .link { border:1px solid var(--grey-2); background:#F3F5FC; padding:8px 10px; border-radius:999px; cursor:pointer; font-weight:800; font-size:12px; color:var(--grey-9); }"
    m_Lineas.Add "    .link:hover { border-color:var(--tele-blue); }"
    m_Lineas.Add "    .edition-card { border:1px solid var(--grey-2); border-radius:14px; padding:12px; cursor:pointer; background:white; display:grid; gap:6px; }"
    m_Lineas.Add "    .edition-card.no-publicable { border-left:6px solid #B71C1C; background:#FFF7F8; }"
    m_Lineas.Add "    .edition-card.publicable { border-left:6px solid #1B5E20; }"
    m_Lineas.Add "    .edition-card.active { outline:2px solid var(--tele-blue); border-color:transparent; }"
    m_Lineas.Add "    .edition-head { display:flex; justify-content:space-between; gap:10px; align-items:center; }"
    m_Lineas.Add "    .edition-label { font-weight:800; color:var(--grey-9); font-size:13px; }"
    m_Lineas.Add "    footer { text-align:center; padding:40px; color:var(--grey-6); font-size:12px; }"
    m_Lineas.Add "  </style>"
    m_Lineas.Add "</head>"
    m_Lineas.Add "<body>"
    m_Lineas.Add "  <header>"
    m_Lineas.Add "    <div class='container header-container'>"
    m_Lineas.Add "      <div class='header-text'>"
    m_Lineas.Add "        <h1>GESTIÓN DE RIESGOS - Publicabilidad</h1>"
    m_Lineas.Add "      </div>"
    m_Lineas.Add "      <div class='header-info'>"
    m_Lineas.Add "        <p style='margin:0;'><strong id='hdrProyecto'>-</strong></p>"
    m_Lineas.Add "        <p style='margin:0;'>Fecha: <span id='hdrFecha'>-</span></p>"
    m_Lineas.Add "        <div id='hdrEdicionBadge' class='hdr-badge'>-</div>"
    m_Lineas.Add "      </div>"
    m_Lineas.Add "    </div>"
    m_Lineas.Add "  </header>"
    m_Lineas.Add "  <div class='container main-container'>"
    m_Lineas.Add "    <div class='layout'>"
    m_Lineas.Add "      <div class='sidebar'>"
    m_Lineas.Add "        <div class='card sidebar-top'>"
    m_Lineas.Add "          <div class='verdict-card'>"
    m_Lineas.Add "            <div>"
    m_Lineas.Add "              <div style='font-weight:800;color:var(--grey-9);'>Edición</div>"
    m_Lineas.Add "              <div style='font-size:13px;color:var(--grey-6);' id='lblEdicion'>-</div>"
    m_Lineas.Add "              <div style='font-size:12px;color:var(--grey-6);margin-top:6px;' id='lblResumen'>-</div>"
    m_Lineas.Add "            </div>"
    m_Lineas.Add "            <div class='verdict-badge' id='badgeEdicion'>-</div>"
    m_Lineas.Add "          </div>"
    m_Lineas.Add "          <input id='txtBuscar' class='search' placeholder='Buscar por código o texto...' />"
    m_Lineas.Add "          <div class='filters'>"
    m_Lineas.Add "            <button class='chip active' data-filter='Todos'>Todos</button>"
    m_Lineas.Add "            <button class='chip' data-filter='Publicable'>Publicables</button>"
    m_Lineas.Add "            <button class='chip' data-filter='NoPublicable'>No publicables</button>"
    m_Lineas.Add "            <button class='chip' data-filter='NoAplica'>No aplica</button>"
    m_Lineas.Add "          </div>"
    m_Lineas.Add "        </div>"
    m_Lineas.Add "        <div class='card edition-item' id='editionCard' data-id='__EDICION__'>"
        m_Lineas.Add "          <div class='edition-head'>"
        m_Lineas.Add "            <div class='edition-label'>Gestión de Riesgos</div>"
        m_Lineas.Add "            <div class='mini-badge' id='editionMiniBadge'>-</div>"
        m_Lineas.Add "          </div>"
        m_Lineas.Add "          <div class='risk-desc' id='editionSummary'>-</div>"
        m_Lineas.Add "        </div>"
        m_Lineas.Add "        <div class='risk-list' id='riskList'></div>"
    m_Lineas.Add "      </div>"
    m_Lineas.Add "      <div class='detail'>"
    m_Lineas.Add "        <div class='card' id='detailPanel'></div>"
    m_Lineas.Add "      </div>"
    m_Lineas.Add "    </div>"
    m_Lineas.Add "  </div>"
    m_Lineas.Add "  <footer><p>© <span id='hdrYear'></span> Telefónica - Aplicación GESTIÓN DE RIESGOS</p></footer>"
    m_Lineas.Add ""
    m_Lineas.Add "  <script id='data' type='application/json'>__DATA_JSON__</script>"
    m_Lineas.Add ""
    m_Lineas.Add "  <script>"
    m_Lineas.Add "    function estadoClass(estado) { if (estado === 'Cumple') return 'state-cumple'; if (estado === 'NoCumple') return 'state-no-cumple'; return 'state-no-aplica'; }"
    m_Lineas.Add "    function veredictoClase(ver) { if (ver === 'Publicable') return 'verdict-publicable'; if (ver === 'NoPublicable') return 'verdict-no-publicable'; return 'verdict-no-aplica'; }"
    m_Lineas.Add "    function veredictoTexto(ver) { if (ver === 'Publicable') return 'Publicable'; if (ver === 'NoPublicable') return 'No publicable'; return 'No aplica'; }"
    m_Lineas.Add "    function miniClase(ver) { if (ver === 'Publicable') return 'mini-badge mini-publicable'; if (ver === 'NoPublicable') return 'mini-badge mini-no-publicable'; return 'mini-badge mini-no-aplica'; }"
    m_Lineas.Add "    function riesgoClase(ver) { if (ver === 'Publicable') return 'publicable'; if (ver === 'NoPublicable') return 'no-publicable'; return 'no-aplica'; }"
    m_Lineas.Add "    function safe(s) { s = (s ?? '').toString(); return s.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll(String.fromCharCode(34),'&quot;'); }"
    m_Lineas.Add ""
    m_Lineas.Add "    const data = JSON.parse(document.getElementById('data').textContent || '{}');"
    m_Lineas.Add "    const riesgosAll = Array.isArray(data.riesgos) ? data.riesgos : [];"
    m_Lineas.Add "    let filtro = 'Todos';"
    m_Lineas.Add "    let q = '';"
    m_Lineas.Add "    let seleccionado = null;"
    m_Lineas.Add ""
    m_Lineas.Add "    function aplicarFiltros(r) {"
    m_Lineas.Add "      const texto = ((r.codigo || '') + ' ' + (r.descripcion || '')).toLowerCase();"
    m_Lineas.Add "      const okQ = q.length === 0 || texto.includes(q);"
    m_Lineas.Add "      const okF = filtro === 'Todos'"
    m_Lineas.Add "        || (filtro === 'Publicable' && r.veredicto === 'Publicable')"
    m_Lineas.Add "        || (filtro === 'NoPublicable' && r.veredicto === 'NoPublicable')"
    m_Lineas.Add "        || (filtro === 'NoAplica' && r.veredicto === 'NoAplica');"
    m_Lineas.Add "      return okQ && okF;"
    m_Lineas.Add "    }"
    m_Lineas.Add ""
    m_Lineas.Add "    function renderCabecera() {"
    m_Lineas.Add "      const ed = data.edicion || {};"
    m_Lineas.Add "      const res = data.resumen || {};"
    m_Lineas.Add "      document.getElementById('hdrProyecto').textContent = (ed.proyecto || '-') + ' / ' + (ed.nombre || '-');"
    m_Lineas.Add "      document.getElementById('hdrFecha').textContent = new Date().toLocaleDateString('es-ES');"
    m_Lineas.Add "      document.getElementById('hdrYear').textContent = new Date().getFullYear().toString();"
    m_Lineas.Add "      document.getElementById('lblEdicion').textContent = (ed.proyecto || '-') + ' · ' + (ed.nombre || '-');"
    m_Lineas.Add "      document.getElementById('lblResumen').textContent = 'Total: ' + (res.total ?? 0) + ' · Publicables: ' + (res.publicables ?? 0) + ' · No publicables: ' + (res.noPublicables ?? 0) + ' · No aplica: ' + (res.noAplica ?? 0);"
    m_Lineas.Add "      const badge = document.getElementById('badgeEdicion');"
    m_Lineas.Add "      const edPub = !!res.edicionPublicable;"
    m_Lineas.Add "      badge.className = 'verdict-badge ' + (edPub ? 'verdict-publicable' : 'verdict-no-publicable');"
    m_Lineas.Add "      badge.textContent = edPub ? 'Publicable' : 'No publicable';"
    m_Lineas.Add ""
m_Lineas.Add "      const hdrBadge = document.getElementById('hdrEdicionBadge');"
        m_Lineas.Add "      hdrBadge.className = 'hdr-badge ' + (edPub ? 'publicable' : 'no-publicable');"
        m_Lineas.Add "      hdrBadge.textContent = 'Edición: ' + (edPub ? 'Publicable' : 'No publicable');"
        m_Lineas.Add "      const edMini = document.getElementById('editionMiniBadge');"
        m_Lineas.Add "      const edSum = document.getElementById('editionSummary');"
        m_Lineas.Add "      if (edMini) { edMini.className = 'mini-badge ' + (edPub ? 'mini-publicable' : 'mini-no-publicable'); edMini.textContent = edPub ? 'Publicable' : 'No publicable'; }"
        m_Lineas.Add "      if (edSum) { const editionChecks = Array.isArray(data.edition_checks) ? data.edition_checks : []; const noCumple = editionChecks.filter(c => c.estado === 'NoCumple').length; edSum.textContent = editionChecks.length + ' checks · ' + noCumple + ' no cumplen'; }"
        m_Lineas.Add "    }"
        m_Lineas.Add ""
        m_Lineas.Add "    function renderLista() {"
        m_Lineas.Add "      const cont = document.getElementById('riskList');"
        m_Lineas.Add "      cont.innerHTML = '';"
m_Lineas.Add "      const editionCard = document.getElementById('editionCard');"
        m_Lineas.Add "      if (editionCard) {"
        m_Lineas.Add "        const isEdSelected = (seleccionado && seleccionado.idRiesgo === '__EDICION__');"
        m_Lineas.Add "        const edPub = !!(data.resumen && data.resumen.edicionPublicable);"
        m_Lineas.Add "        editionCard.className = 'risk-item edition-card ' + riesgoClase(edPub ? 'Publicable' : 'NoPublicable') + (isEdSelected ? ' active' : '');"
        m_Lineas.Add "        editionCard.addEventListener('click', () => seleccionar('__EDICION__'));"
        m_Lineas.Add "      }"
        m_Lineas.Add "      const riesgos = riesgosAll.filter(r => r.idRiesgo !== '__EDICION__' && aplicarFiltros(r));"
        m_Lineas.Add "      if (riesgos.length === 0) {"
        m_Lineas.Add "        cont.innerHTML = ""<div class='card'>No hay riesgos que coincidan con el filtro.</div>"";"
        m_Lineas.Add "        return;"
        m_Lineas.Add "      }"
    m_Lineas.Add "      for (const r of riesgos) {"
    m_Lineas.Add "        const div = document.createElement('div');"
    m_Lineas.Add "        const extra = ' ' + riesgoClase(r.veredicto);"
    m_Lineas.Add "        div.className = 'risk-item' + extra + (seleccionado && seleccionado.idRiesgo === r.idRiesgo ? ' active' : '');"
    m_Lineas.Add "        div.dataset.id = r.idRiesgo;"
    m_Lineas.Add "        div.innerHTML = ""<div class='risk-head'><div class='risk-code'>"" + safe(r.codigo) + ""</div><div class='"" + miniClase(r.veredicto) + ""'>"" + safe(veredictoTexto(r.veredicto)) + ""</div></div><div class='risk-desc'>"" + safe(r.descripcion || '') + ""</div>"";"
    m_Lineas.Add "        div.addEventListener('click', () => seleccionar(r.idRiesgo));"
    m_Lineas.Add "        cont.appendChild(div);"
    m_Lineas.Add "      }"
    m_Lineas.Add "    }"
    m_Lineas.Add ""
m_Lineas.Add "    function renderDetalle(r) {"
        m_Lineas.Add "      const panel = document.getElementById('detailPanel');"
        m_Lineas.Add "      if (!r || r.idRiesgo === '__EDICION__') {"
        m_Lineas.Add "        const res = data.resumen || {};"
        m_Lineas.Add "        const ed = data.edicion || {};"
        m_Lineas.Add "        const edPub = !!res.edicionPublicable;"
        m_Lineas.Add "        const noPub = riesgosAll.filter(x => x.veredicto === 'NoPublicable');"
        m_Lineas.Add "        const editionChecks = Array.isArray(data.edition_checks) ? data.edition_checks : [];"
    m_Lineas.Add "        let html = '';"
    m_Lineas.Add "        html += ""<div class='verdict-card'>"";"
    m_Lineas.Add "        html += ""<div><h2 style='margin:0 0 6px 0;'>Resultado de la edición</h2>"";"
    m_Lineas.Add "        html += ""<div style='color:var(--grey-6); font-size:13px;'><strong>"" + safe(ed.proyecto || '-') + ""</strong> · "" + safe(ed.nombre || '-') + ""</div></div>"";"
    m_Lineas.Add "        html += ""<div class='verdict-badge "" + (edPub ? 'verdict-publicable' : 'verdict-no-publicable') + ""'>"" + (edPub ? 'Publicable' : 'No publicable') + ""</div>"";"
    m_Lineas.Add "        html += ""</div>"";"
    m_Lineas.Add "        if (editionChecks.length > 0) {"
    m_Lineas.Add "          html += ""<div style='margin-top:12px;'>"";"
    m_Lineas.Add "          html += ""<h2>Checks de Edición</h2>"";"
    m_Lineas.Add "          html += ""<div class='checklist'>"";"
    m_Lineas.Add "          for (const c of editionChecks) {"
    m_Lineas.Add "            const est = (c.estado || 'NoAplica');"
    m_Lineas.Add "            const det = (c.detalle || '');"
    m_Lineas.Add "            html += ""<div class='check-item'>"";"
    m_Lineas.Add "            html += ""<div class='check-left'><div class='check-text'>"" + safe(c.texto || '') + ""</div>"";"
    m_Lineas.Add "            if (det) html += ""<div class='check-detail'>"" + safe(det) + ""</div>"";"
    m_Lineas.Add "            html += ""</div>"";"
    m_Lineas.Add "            html += ""<div class='check-state "" + estadoClass(est) + ""'>"" + safe(est === 'NoCumple' ? 'No cumple' : (est === 'Cumple' ? 'Cumple' : 'No aplica')) + ""</div>"";"
    m_Lineas.Add "            html += ""</div>"";"
    m_Lineas.Add "          }"
    m_Lineas.Add "          html += ""</div>"";"
    m_Lineas.Add "          html += ""</div>"";"
    m_Lineas.Add "        }"
    m_Lineas.Add "        if (!edPub && noPub.length > 0) {"
    m_Lineas.Add "          html += ""<div style='color:var(--grey-9); font-weight:800; margin-top:10px;'>Riesgos que bloquean la publicación</div>"";"
    m_Lineas.Add "          html += ""<div class='list-links'>"";"
    m_Lineas.Add "          for (const rp of noPub) {"
    m_Lineas.Add "            html += ""<button class='link' data-id='"" + safe(rp.idRiesgo) + ""'>"" + safe(rp.codigo) + ""</button>"";"
    m_Lineas.Add "          }"
    m_Lineas.Add "          html += ""</div>"";"
    m_Lineas.Add "          html += ""<div style='color:var(--grey-6); font-size:12px; margin-top:10px;'>Pulsa un riesgo para ver sus comprobaciones.</div>"";"
    m_Lineas.Add "        } else {"
    m_Lineas.Add "          html += ""<div style='color:var(--grey-6);'>Pulsa en un riesgo de la lista izquierda para ver sus comprobaciones.</div>"";"
    m_Lineas.Add "        }"
    m_Lineas.Add "        panel.innerHTML = html;"
m_Lineas.Add "        for (const btn of panel.querySelectorAll('button[data-id]')) {"
        m_Lineas.Add "          btn.addEventListener('click', () => seleccionar(btn.getAttribute('data-id')));"
        m_Lineas.Add "        }"
        m_Lineas.Add "        if (!seleccionado || seleccionado.idRiesgo !== '__EDICION__') { seleccionado = { idRiesgo: '__EDICION__', codigo: '__EDICION__', descripcion: 'Gestión de Riesgos' }; }"
        m_Lineas.Add "        renderLista();"
        m_Lineas.Add "        return;"
        m_Lineas.Add "      }"
    m_Lineas.Add ""
    m_Lineas.Add "      const verClase = veredictoClase(r.veredicto);"
    m_Lineas.Add "      const verText = veredictoTexto(r.veredicto);"
    m_Lineas.Add "      const checks = Array.isArray(r.checks) ? r.checks : [];"
    m_Lineas.Add "      let html = '';"
    m_Lineas.Add "      html += ""<div class='verdict-card'>"";"
    m_Lineas.Add "      html += ""<div><h2 style='margin:0 0 6px 0;'>Publicabilidad</h2><div style='color:var(--grey-6); font-size:13px;'><strong>"" + safe(r.codigo) + ""</strong> · "" + safe(r.descripcion || '') + ""</div></div>"";"
    m_Lineas.Add "      html += ""<div class='verdict-badge "" + verClase + ""'>"" + safe(verText) + ""</div>"";"
    m_Lineas.Add "      html += ""</div>"";"
    m_Lineas.Add "      html += ""<div class='checklist'>"";"
    m_Lineas.Add "      for (const c of checks) {"
    m_Lineas.Add "        const est = (c.estado || 'NoAplica');"
    m_Lineas.Add "        const det = (c.detalle || '');"
    m_Lineas.Add "        html += ""<div class='check-item'>"";"
    m_Lineas.Add "        html += ""<div class='check-left'><div class='check-text'>"" + safe(c.texto || '') + ""</div>"";"
    m_Lineas.Add "        if (det) html += ""<div class='check-detail'>"" + safe(det) + ""</div>"";"
    m_Lineas.Add "        html += ""</div>"";"
    m_Lineas.Add "        html += ""<div class='check-state "" + estadoClass(est) + ""'>"" + safe(est === 'NoCumple' ? 'No cumple' : (est === 'Cumple' ? 'Cumple' : 'No aplica')) + ""</div>"";"
    m_Lineas.Add "        html += ""</div>"";"
    m_Lineas.Add "      }"
    m_Lineas.Add "      html += ""</div>"";"
    m_Lineas.Add "      panel.innerHTML = html;"
    m_Lineas.Add "    }"
    m_Lineas.Add ""
    m_Lineas.Add "    function seleccionar(idRiesgo) {"
    m_Lineas.Add "      const r = riesgosAll.find(x => x.idRiesgo === idRiesgo) || null;"
    m_Lineas.Add "      seleccionado = r;"
    m_Lineas.Add "      renderLista();"
    m_Lineas.Add "      renderDetalle(r);"
    m_Lineas.Add "      if (r && r.idRiesgo) location.hash = 'riesgo-' + encodeURIComponent(r.idRiesgo);"
    m_Lineas.Add "    }"
    m_Lineas.Add ""
    m_Lineas.Add "    document.getElementById('txtBuscar').addEventListener('input', (e) => {"
    m_Lineas.Add "      q = (e.target.value || '').trim().toLowerCase();"
    m_Lineas.Add "      renderLista();"
    m_Lineas.Add "    });"
    m_Lineas.Add ""
    m_Lineas.Add "    for (const btn of document.querySelectorAll('.chip')) {"
    m_Lineas.Add "      btn.addEventListener('click', () => {"
    m_Lineas.Add "        for (const b of document.querySelectorAll('.chip')) b.classList.remove('active');"
    m_Lineas.Add "        btn.classList.add('active');"
    m_Lineas.Add "        filtro = btn.dataset.filter || 'Todos';"
    m_Lineas.Add "        renderLista();"
    m_Lineas.Add "      });"
    m_Lineas.Add "    }"
    m_Lineas.Add ""
    m_Lineas.Add "    renderCabecera();"
    m_Lineas.Add "    renderLista();"
    m_Lineas.Add ""
m_Lineas.Add "    const noPub = riesgosAll.filter(x => x.veredicto === 'NoPublicable');"
        m_Lineas.Add "    if (location.hash && location.hash.startsWith('#riesgo-')) {"
        m_Lineas.Add "      const id = decodeURIComponent(location.hash.substring('#riesgo-'.length));"
        m_Lineas.Add "      seleccionar(id);"
        m_Lineas.Add "    } else {"
        m_Lineas.Add "      seleccionar('__EDICION__');"
        m_Lineas.Add "    }"
        m_Lineas.Add "  </script>"
    m_Lineas.Add "</body>"
    m_Lineas.Add "</html>"

    m_HTML = ""
    For Each m_Linea In m_Lineas
        m_HTML = m_HTML & CStr(m_Linea) & vbCrLf
    Next
    If Right$(m_HTML, 2) = vbCrLf Then
        m_HTML = Left$(m_HTML, Len(m_HTML) - 2)
    End If

    ConstruirTemplateInformePublicabilidadEdicion = m_HTML
End Function

Public Function URLInformePublicabilidad( _
                                            Optional p_Edicion As Edicion, _
                                            Optional ByRef p_Error As String, _
                                            Optional ByVal p_GenerarPDF As Boolean = False _
                                            ) As String

   
    
    On Error GoTo errores
    If p_Edicion Is Nothing Then
        p_Error = "Faltan datos para el informe de publicabilidad"
        Err.Raise 1000
    End If
    URLInformePublicabilidad = GenerarInformePublicabilidadEdicionInteractivoHTML(p_Edicion:=p_Edicion, p_Error:=p_Error, p_GenerarPDF:=p_GenerarPDF)
    If p_Error <> "" Then Err.Raise 1000

    

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método URLInformePublicabilidad ha producido el error num " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If

End Function


Public Function getRiesgosPorEdicionPub( _
                                        p_IDEdicion As String, _
                                        Optional p_db As DAO.Database, _
                                        Optional p_PorPriorizacion As EnumSiNo, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_OrderBY As String
    Dim db As DAO.Database
    On Error GoTo errores
    If p_IDEdicion = "" Then
        Exit Function
    End If
    If p_PorPriorizacion = EnumSiNo.Sí Then
       m_OrderBY = "ORDER BY Priorizacion;"
    Else
        m_OrderBY = "ORDER BY IDRiesgo;"
    End If
    If p_db Is Nothing Then
        Set db = getdb()
    Else
        Set db = p_db
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDEdicion=" & p_IDEdicion & " AND FechaRetirado IS NULL " & _
            m_OrderBY
    
    Set rcdDatos = db.OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgo = New riesgo
            For Each m_Campo In m_ObjRiesgo.ColCampos
                m_ObjRiesgo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getRiesgosPorEdicionPub Is Nothing Then
                Set getRiesgosPorEdicionPub = New Scripting.Dictionary
                getRiesgosPorEdicionPub.CompareMode = TextCompare
            End If
            If Not getRiesgosPorEdicionPub.Exists(CStr(m_ObjRiesgo.IDRiesgo)) Then
                getRiesgosPorEdicionPub.Add CStr(m_ObjRiesgo.IDRiesgo), m_ObjRiesgo
            End If
            Set m_ObjRiesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRiesgosPorEdicionPub ha devuelto el error: " & Err.Description
    End If
End Function



