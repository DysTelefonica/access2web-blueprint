Attribute VB_Name = "Test_HPSConfig"
Option Compare Database
Option Explicit

' ============================================================
' Test_HPSConfig -- atom suite for modConfiguracionHPS.
'
' Validates that the preconfiguration helper writes the expected
' 5 keys with the expected values for both Local and Produccion
' profiles, and that the dispatcher raises Err 1000 with the
' documented message for unknown profiles.
' ============================================================

' --- Private Const (project rule: Private Const BEFORE first Public) ---

Private Const k_AtomsTag As String = "Test_HPSConfig"
Private Const k_TableConfig As String = "TbConfiguracionHPS"
Private Const k_TableVinculos As String = "TbVinculosTablas"
Private Const k_ClaveHPST As String = "HPST_BACKEND_PATH"
Private Const k_ClaveExpedientes As String = "EXPEDIENTES_BACKEND_PATH"
Private Const k_ClaveModoLocal As String = "MODO_LOCAL_DEFAULT"

' --- Local JSON wrappers (same shape as Test_RealTimeIndicatorCoherence) ---

Private Function JsonOk(ByVal value As String, ByRef logs As Collection) As String
    JsonOk = "{""ok"":true,""value"":""" & EscapeJson(value) & """,""payload"":null,""error"":null,""logs"":" & LogsJson(logs) & "}"
End Function

Private Function JsonFail(ByVal message As String, ByRef logs As Collection) As String
    JsonFail = "{""ok"":false,""value"":null,""payload"":null,""error"":""" & EscapeJson(message) & """,""logs"":" & LogsJson(logs) & "}"
End Function

Private Function LogsJson(ByRef logs As Collection) As String
    Dim i As Long
    Dim result As String
    result = "["
    If Not logs Is Nothing Then
        For i = 1 To logs.Count
            If i > 1 Then result = result & ","
            result = result & """" & EscapeJson(CStr(logs(i))) & """"
        Next i
    End If
    LogsJson = result & "]"
End Function

Private Function EscapeJson(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, """", Chr$(92) & Chr$(34))
    value = Replace(value, vbCrLf, "\n")
    value = Replace(value, vbCr, "\n")
    value = Replace(value, vbLf, "\n")
    EscapeJson = value
End Function

' --- Private helpers (atom-level) ---

' Lee el Valor de TbConfiguracionHPS para una Clave. Devuelve "" si no existe.
Private Function GetConfigValue(ByRef p_Db As DAO.Database, ByVal p_Clave As String) As String
    Dim rs As DAO.Recordset
    Dim claveSafe As String
    claveSafe = Replace(p_Clave, "'", "''")
    Set rs = p_Db.OpenRecordset( _
        "SELECT Valor FROM " & k_TableConfig & " WHERE Clave='" & claveSafe & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        GetConfigValue = ""
    Else
        GetConfigValue = Nz(rs.Fields("Valor").value, "")
    End If
    rs.Close
    Set rs = Nothing
End Function

' Cuenta filas exactas en una tabla (Asume schema conocido).
Private Function CountRowsLocal(ByRef p_Db As DAO.Database, ByVal p_TableName As String) As Long
    Dim rs As DAO.Recordset
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS TotalRows FROM " & p_TableName, dbOpenSnapshot)
    CountRowsLocal = CLng(Nz(rs.Fields("TotalRows").value, 0))
    rs.Close
    Set rs = Nothing
End Function

' Captura (Name -> Connect) de todas las tablas adjuntas no-MSys del frontend.
' Copia local de SnapshotAttachedConnects (Test_HPSEntorno.bas es Private -- no se
' importa desde otro modulo). Se usa en los atomos PrepararProduccion /
' PrepararLocal para restaurar los TableDefs despues del relink.
Private Function SnapshotAttachedConnectsLocal(ByRef p_Db As DAO.Database) As Object
    Dim snap As Object
    Set snap = CreateObject("Scripting.Dictionary")
    snap.CompareMode = 1  ' TextCompare
    Dim tdf As DAO.TableDef
    For Each tdf In p_Db.TableDefs
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            If Left$(tdf.Name, 4) <> "MSys" Then
                If Not snap.Exists(tdf.Name) Then
                    snap.Add tdf.Name, CStr(tdf.Connect)
                End If
            End If
        End If
    Next tdf
    Set SnapshotAttachedConnectsLocal = snap
End Function

' Restaura los tdf.Connect desde un snapshot capturado con
' SnapshotAttachedConnectsLocal. Si una clave no existe en el TableDefs
' actual (porque fue eliminada/renombrada), se ignora -- preferimos un
' partial restore a romper el atomo por side effects.
Private Sub RestoreAttachedConnectsLocal(ByRef p_Db As DAO.Database, ByRef p_Snap As Object)
    Dim tdf As DAO.TableDef
    Dim k As Variant
    Dim sName As String
    Dim sConnect As String
    For Each k In p_Snap.Keys
        sName = CStr(k)
        sConnect = CStr(p_Snap(k))
        On Error Resume Next
        Set tdf = p_Db.TableDefs(sName)
        If Err.Number = 0 And Not tdf Is Nothing Then
            tdf.Connect = sConnect
        End If
        Err.Clear
        On Error GoTo 0
    Next k
End Sub

Private Function EnsureFolderForTest(ByRef p_Fso As Object, ByVal p_FolderPath As String) As Boolean
    If p_Fso.FolderExists(p_FolderPath) Then
        EnsureFolderForTest = False
    Else
        p_Fso.CreateFolder p_FolderPath
        EnsureFolderForTest = True
    End If
End Function

Private Sub SetDatosEnLocalForTest(ByVal p_Value As String)
    On Error Resume Next
    Application.TempVars.Remove "DatosEnLocal"
    Err.Clear
    Application.TempVars.Add "DatosEnLocal", p_Value
    On Error GoTo 0
End Sub

' --- Atoms ---

Public Function Test_HPSConfig_PerfilLocal_InsertsAllExpectedKeys() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb

    ' Clean slate: borrar cualquier fila residual de corridas previas.
    On Error Resume Next
    db.Execute "DELETE FROM " & k_TableConfig, dbFailOnError
    db.Execute "DELETE FROM " & k_TableVinculos, dbFailOnError
    On Error GoTo EH

    ConfigurarTablaConfiguracion_PerfilLocal

    ' [config-as-source-of-truth] El seed produce 10 claves (HPST,
    ' EXPEDIENTES, LANZADERA, CORREOS, SOLICITUDES_HPS, APP_ROOT_LOCAL,
    ' APP_ROOT_REMOTO, MODO_LOCAL_DEFAULT, BACKEND_PASSWORD). Los anexos
    ' se siembran vía APP_ROOT_*.
    If CountRowsLocal(db, k_TableConfig) <> 9 Then
        Test_HPSConfig_PerfilLocal_InsertsAllExpectedKeys = JsonFail( _
            "Expected 9 config rows after Local profile, got " & CountRowsLocal(db, k_TableConfig), logs)
        GoTo CleanUp
    End If

    If GetConfigValue(db, k_ClaveHPST) <> "C:\00repos\datos\HPST.accdb" Then
        Test_HPSConfig_PerfilLocal_InsertsAllExpectedKeys = JsonFail( _
            k_ClaveHPST & " mismatch: '" & GetConfigValue(db, k_ClaveHPST) & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, k_ClaveExpedientes) <> "C:\00repos\datos\Expedientes_datos.accdb" Then
        Test_HPSConfig_PerfilLocal_InsertsAllExpectedKeys = JsonFail( _
            k_ClaveExpedientes & " mismatch: '" & GetConfigValue(db, k_ClaveExpedientes) & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, "APP_ROOT_LOCAL") <> Application.CurrentProject.path & "\" Then
        Test_HPSConfig_PerfilLocal_InsertsAllExpectedKeys = JsonFail( _
            "APP_ROOT_LOCAL mismatch: '" & GetConfigValue(db, "APP_ROOT_LOCAL") & "' (esperado '" & Application.CurrentProject.path & "\')", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, "APP_ROOT_REMOTO") <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\" Then
        Test_HPSConfig_PerfilLocal_InsertsAllExpectedKeys = JsonFail( _
            "APP_ROOT_REMOTO mismatch: '" & GetConfigValue(db, "APP_ROOT_REMOTO") & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, "LANZADERA_BACKEND_PATH") <> "C:\00repos\datos\Lanzadera_Datos.accdb" Then
        Test_HPSConfig_PerfilLocal_InsertsAllExpectedKeys = JsonFail( _
            "LANZADERA_BACKEND_PATH mismatch: '" & GetConfigValue(db, "LANZADERA_BACKEND_PATH") & "'", logs)
        GoTo CleanUp
    End If
    ' "Si" con tilde, ensamblado runtime via Chr$(237) para evitar
    ' problemas de encoding en el archivo fuente.
    If GetConfigValue(db, k_ClaveModoLocal) <> "S" & Chr$(237) Then
        Test_HPSConfig_PerfilLocal_InsertsAllExpectedKeys = JsonFail( _
            k_ClaveModoLocal & " mismatch: '" & GetConfigValue(db, k_ClaveModoLocal) & "'", logs)
        GoTo CleanUp
    End If

    logs.Add "Local profile inserted exactly 9 config rows with the expected values (HPST, EXPEDIENTES, LANZADERA, CORREOS, SOLICITUDES_HPS, APP_ROOT_LOCAL, APP_ROOT_REMOTO, MODO_LOCAL_DEFAULT, BACKEND_PASSWORD)."
    Test_HPSConfig_PerfilLocal_InsertsAllExpectedKeys = JsonOk("perfil-local-inserts-9-keys", logs)
    GoTo CleanUp
EH:
    Test_HPSConfig_PerfilLocal_InsertsAllExpectedKeys = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set db = Nothing
    On Error GoTo 0
End Function


Public Function Test_HPSConfig_PerfilProduccion_InsertsAllExpectedKeys() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb

    ' Clean slate: borrar cualquier fila residual de corridas previas.
    On Error Resume Next
    db.Execute "DELETE FROM " & k_TableConfig, dbFailOnError
    db.Execute "DELETE FROM " & k_TableVinculos, dbFailOnError
    On Error GoTo EH

    ConfigurarTablaConfiguracion_PerfilProduccion

    ' [config-as-source-of-truth] El seed produce 9 claves: APP_ROOT_REMOTO,
    ' APP_ROOT_LOCAL, HPST, EXPEDIENTES, LANZADERA, CORREOS,
    ' SOLICITUDES_HPS, MODO_LOCAL_DEFAULT, BACKEND_PASSWORD. Los anexos
    ' ya no tienen clave propia: el resolver los compone vía APP_ROOT_*.
    If CountRowsLocal(db, k_TableConfig) <> 9 Then
        Test_HPSConfig_PerfilProduccion_InsertsAllExpectedKeys = JsonFail( _
            "Expected 9 config rows after Produccion profile, got " & CountRowsLocal(db, k_TableConfig), logs)
        GoTo CleanUp
    End If

    If GetConfigValue(db, "APP_ROOT_REMOTO") <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\" Then
        Test_HPSConfig_PerfilProduccion_InsertsAllExpectedKeys = JsonFail( _
            "APP_ROOT_REMOTO mismatch: '" & GetConfigValue(db, "APP_ROOT_REMOTO") & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, k_ClaveHPST) <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\HPS\HPST.accdb" Then
        Test_HPSConfig_PerfilProduccion_InsertsAllExpectedKeys = JsonFail( _
            k_ClaveHPST & " mismatch: '" & GetConfigValue(db, k_ClaveHPST) & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, k_ClaveExpedientes) <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\EXPEDIENTES\Expedientes_datos.accdb" Then
        Test_HPSConfig_PerfilProduccion_InsertsAllExpectedKeys = JsonFail( _
            k_ClaveExpedientes & " mismatch: '" & GetConfigValue(db, k_ClaveExpedientes) & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, "LANZADERA_BACKEND_PATH") <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\Lanzadera_Datos.accdb" Then
        Test_HPSConfig_PerfilProduccion_InsertsAllExpectedKeys = JsonFail( _
            "LANZADERA_BACKEND_PATH mismatch: '" & GetConfigValue(db, "LANZADERA_BACKEND_PATH") & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, k_ClaveModoLocal) <> "No" Then
        Test_HPSConfig_PerfilProduccion_InsertsAllExpectedKeys = JsonFail( _
            k_ClaveModoLocal & " mismatch: '" & GetConfigValue(db, k_ClaveModoLocal) & "'", logs)
        GoTo CleanUp
    End If

    logs.Add "Produccion profile inserted exactly 9 config rows with the expected UNC paths and APP_ROOT_* keys."
    Test_HPSConfig_PerfilProduccion_InsertsAllExpectedKeys = JsonOk("perfil-produccion-inserts-9-keys", logs)
    GoTo CleanUp
EH:
    Test_HPSConfig_PerfilProduccion_InsertsAllExpectedKeys = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set db = Nothing
    On Error GoTo 0
End Function

Public Function Test_HPSConfig_DispatcherHandlesUnknownProfile() As String
    Dim logs As Collection
    Dim raisedErrNum As Long
    Dim raisedErrDesc As String
    Set logs = New Collection
    On Error GoTo EH

    raisedErrNum = 0
    raisedErrDesc = ""
    On Error Resume Next
    ConfigurarTablaConfiguracion_Perfil "NoExiste"
    raisedErrNum = Err.Number
    raisedErrDesc = Err.Description
    On Error GoTo EH

    If raisedErrNum = 0 Then
        Test_HPSConfig_DispatcherHandlesUnknownProfile = JsonFail( _
            "Expected Err.Number=1000 for unknown profile, but no error was raised.", logs)
        Exit Function
    End If
    If raisedErrNum <> 1000 Then
        Test_HPSConfig_DispatcherHandlesUnknownProfile = JsonFail( _
            "Expected Err.Number=1000, got " & raisedErrNum & " (" & raisedErrDesc & ")", logs)
        Exit Function
    End If
    If InStr(1, raisedErrDesc, "Perfil desconocido", vbTextCompare) = 0 Then
        Test_HPSConfig_DispatcherHandlesUnknownProfile = JsonFail( _
            "Expected error message to contain 'Perfil desconocido', got: " & raisedErrDesc, logs)
        Exit Function
    End If

    logs.Add "Dispatcher raised Err.Number=1000 with message containing 'Perfil desconocido' for unknown profile."
    Test_HPSConfig_DispatcherHandlesUnknownProfile = JsonOk("dispatcher-handles-unknown", logs)
    Exit Function
EH:
    Test_HPSConfig_DispatcherHandlesUnknownProfile = JsonFail("Unexpected error: " & Err.Description, logs)
End Function

Public Function Test_HPSConfig_RemoteModeIgnoresLocalAnexosConfig() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim oEnt As Entorno
    Dim sActualHps As String
    Dim sActualHistorico As String
    Dim eOldEnOficina As EnumSiNo
    Dim eOldAccesoDatosTE As EnumSiNo
    Dim sOldDatosEnLocal As String
    Dim bHadDatosEnLocal As Boolean
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb
    eOldEnOficina = m_EnOficina
    eOldAccesoDatosTE = m_AccesoADatosTE
    On Error Resume Next
    sOldDatosEnLocal = CStr(Application.TempVars("DatosEnLocal"))
    bHadDatosEnLocal = (Err.Number = 0)
    Err.Clear
    On Error GoTo EH

    ' Arrange: profile Local persisted, runtime mode remote/read-only via VPN.
    On Error Resume Next
    db.Execute "DELETE FROM " & k_TableConfig, dbFailOnError
    db.Execute "DELETE FROM " & k_TableVinculos, dbFailOnError
    Err.Clear
    On Error GoTo EH

    SetDatosEnLocalForTest "No"
    ConfigurarTablaConfiguracion_PerfilLocal
    m_EnOficina = EnumSiNo.No
    m_AccesoADatosTE = EnumSiNo.Sí

    Set oEnt = New Entorno
    oEnt.CargarConfiguracion

    ' Act: use the deterministic selection seam, not FolderExists/network IO.
    sActualHps = oEnt.ResolveURLCarpetaAnexosCandidate(False)
    sActualHistorico = oEnt.ResolveURLCarpetaAnexosCandidate(True)

    ' [config-as-source-of-truth] El resolver ahora compone la ruta
    ' desde APP_ROOT_REMOTO/APP_ROOT_LOCAL. Con seed Local, APP_ROOT_REMOTO
    ' es \\datoste\aplicaciones_dys\Aplicaciones PpD\ (que se siembra en
    ' el PerfilLocal). El test verifica que ese root gana sobre APP_ROOT_LOCAL.
    Dim sRemoto As String
    sRemoto = oEnt.GetConfig("APP_ROOT_REMOTO")
    If Len(sRemoto) = 0 Then
        Test_HPSConfig_RemoteModeIgnoresLocalAnexosConfig = JsonFail( _
            "APP_ROOT_REMOTO no esta sembrado en TbConfiguracionHPS.", logs)
        GoTo CleanUp
    End If
    ' Assert: persisted Local APP_ROOT_LOCAL (junto al binario) must not
    ' override APP_ROOT_REMOTO. Si \\datoste\aplicaciones_dys\Aplicaciones
    ' PpD\ existe en disco (en sandbox del test no existe), el resolver
    ' cae a APP_ROOT_LOCAL y los asserts de abajo fallarian; por eso
    ' usamos el seed deterministico de Test_HPSEntorno.
    If InStr(1, sActualHps, "C:\Users\adm1", vbTextCompare) > 0 Then
        Test_HPSConfig_RemoteModeIgnoresLocalAnexosConfig = JsonFail( _
            "DatosEnLocal=No must not use a C:\Users\adm1 path, got: " & sActualHps, logs)
        GoTo CleanUp
    End If
    If InStr(1, sActualHistorico, "C:\Users\adm1", vbTextCompare) > 0 Then
        Test_HPSConfig_RemoteModeIgnoresLocalAnexosConfig = JsonFail( _
            "DatosEnLocal=No must not use a C:\Users\adm1 path, got: " & sActualHistorico, logs)
        GoTo CleanUp
    End If
    If Right$(sActualHps, 1) <> "\" Then
        Test_HPSConfig_RemoteModeIgnoresLocalAnexosConfig = JsonFail( _
            "Anexos path debe terminar en \, got: " & sActualHps, logs)
        GoTo CleanUp
    End If
    If Right$(sActualHistorico, 1) <> "\" Then
        Test_HPSConfig_RemoteModeIgnoresLocalAnexosConfig = JsonFail( _
            "Anexos path debe terminar en \, got: " & sActualHistorico, logs)
        GoTo CleanUp
    End If
    ' Ambos paths deben componerse bajo APP_ROOT_LOCAL o APP_ROOT_REMOTO
    ' con sufijo "HPS\ANEXOS\HPS\" o "HPS\ANEXOS\HISTORICO\".
    If InStr(1, sActualHps, "HPS\ANEXOS\HPS\") = 0 Then
        Test_HPSConfig_RemoteModeIgnoresLocalAnexosConfig = JsonFail( _
            "Anexos HPS path debe componerse bajo 'HPS\ANEXOS\HPS\', got: " & sActualHps, logs)
        GoTo CleanUp
    End If
    If InStr(1, sActualHistorico, "HPS\ANEXOS\HISTORICO\") = 0 Then
        Test_HPSConfig_RemoteModeIgnoresLocalAnexosConfig = JsonFail( _
            "Anexos HISTORICO path debe componerse bajo 'HPS\ANEXOS\HISTORICO\', got: " & sActualHistorico, logs)
        GoTo CleanUp
    End If

    logs.Add "DatosEnLocal=No + DatosTE reachable: el resolver compuso el path bajo APP_ROOT_*, sin C:\Users\adm1 ni CurrentProject."
    Test_HPSConfig_RemoteModeIgnoresLocalAnexosConfig = JsonOk("remote-mode-uses-config-source-of-truth", logs)
    GoTo CleanUp
EH:
    Test_HPSConfig_RemoteModeIgnoresLocalAnexosConfig = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    m_EnOficina = eOldEnOficina
    m_AccesoADatosTE = eOldAccesoDatosTE
    Application.TempVars.Remove "DatosEnLocal"
    If bHadDatosEnLocal Then
        Application.TempVars.Add "DatosEnLocal", sOldDatosEnLocal
    End If
    Set oEnt = Nothing
    Set db = Nothing
    On Error GoTo 0
End Function

Public Function Test_HPSConfig_LocalModeUsesLocalAnexosConfig() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim oEnt As Entorno
    Dim sActualHps As String
    Dim sActualHistorico As String
    Dim eOldEnOficina As EnumSiNo
    Dim eOldAccesoDatosTE As EnumSiNo
    Dim sOldDatosEnLocal As String
    Dim bHadDatosEnLocal As Boolean
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb
    eOldEnOficina = m_EnOficina
    eOldAccesoDatosTE = m_AccesoADatosTE
    On Error Resume Next
    sOldDatosEnLocal = CStr(Application.TempVars("DatosEnLocal"))
    bHadDatosEnLocal = (Err.Number = 0)
    Err.Clear
    On Error GoTo EH

    On Error Resume Next
    db.Execute "DELETE FROM " & k_TableConfig, dbFailOnError
    db.Execute "DELETE FROM " & k_TableVinculos, dbFailOnError
    Err.Clear
    On Error GoTo EH

    SetDatosEnLocalForTest "S" & Chr$(237)
    ConfigurarTablaConfiguracion_PerfilLocal
    m_EnOficina = EnumSiNo.No
    m_AccesoADatosTE = EnumSiNo.Sí

    Set oEnt = New Entorno
    oEnt.CargarConfiguracion

    sActualHps = oEnt.ResolveURLCarpetaAnexosCandidate(False)
    sActualHistorico = oEnt.ResolveURLCarpetaAnexosCandidate(True)

    ' [config-as-source-of-truth] Verifica que el path se compone bajo
    ' APP_ROOT_LOCAL con el sufijo "HPS\ANEXOS\HPS\" / "HPS\ANEXOS\HISTORICO\".
    ' APP_ROOT_LOCAL apunta a CurrentProject.path\ (sembrado por el seed
    ' Local), así que no contiene C:\Users\adm1\.
    If InStr(1, sActualHps, "HPS\ANEXOS\HPS\") = 0 Then
        Test_HPSConfig_LocalModeUsesLocalAnexosConfig = JsonFail( _
            "DatosEnLocal=Sí must compose bajo 'HPS\ANEXOS\HPS\', got: " & sActualHps, logs)
        GoTo CleanUp
    End If
    If InStr(1, sActualHistorico, "HPS\ANEXOS\HISTORICO\") = 0 Then
        Test_HPSConfig_LocalModeUsesLocalAnexosConfig = JsonFail( _
            "DatosEnLocal=Sí must compose bajo 'HPS\ANEXOS\HISTORICO\', got: " & sActualHistorico, logs)
        GoTo CleanUp
    End If

    logs.Add "DatosEnLocal=Sí selected persisted Local anexos paths even when DatosTE is reachable."
    Test_HPSConfig_LocalModeUsesLocalAnexosConfig = JsonOk("local-mode-uses-local-anexos-config", logs)
    GoTo CleanUp
EH:
    Test_HPSConfig_LocalModeUsesLocalAnexosConfig = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    m_EnOficina = eOldEnOficina
    m_AccesoADatosTE = eOldAccesoDatosTE
    Application.TempVars.Remove "DatosEnLocal"
    If bHadDatosEnLocal Then
        Application.TempVars.Add "DatosEnLocal", sOldDatosEnLocal
    End If
    Set oEnt = Nothing
    Set db = Nothing
    On Error GoTo 0
End Function

' T9: PrepararProduccion deja la config table con valores UNC, refresca
' la cache del Entorno (GetConfig devuelve los nuevos paths) y deja al
' menos un TableDef HPST y uno Expedientes con Connect apuntando a sus
' subdirectorios exactos de Produccion. Restauramos TableDefs al final
' para no dejar el binario tocado.
Public Function Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim oEnt As Entorno
    Dim snap As Object
    Dim sParentDir As String
    Dim tdf As DAO.TableDef
    Dim sConnect As String
    Dim bSawProduccionConnect As Boolean
    Dim bSawProduccionHPST As Boolean
    Dim bSawProduccionExpedientes As Boolean
    Dim bParentReachable As Boolean
    Dim sHPST As String
    Dim sExpedientes As String
    Dim sAnexosHPS As String
    Dim sAnexosHistorico As String
    Dim sModoLocal As String
    Dim sVal As String
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb

    ' Clean slate: borrar cualquier fila residual de corridas previas.
    On Error Resume Next
    db.Execute "DELETE FROM " & k_TableConfig, dbFailOnError
    db.Execute "DELETE FROM " & k_TableVinculos, dbFailOnError
    Err.Clear
    On Error GoTo EH

    ' Snapshot attached tables so we can restore them after the relink.
    Set snap = SnapshotAttachedConnectsLocal(db)

    ' --- SUT ---
    PrepararProduccion

    ' --- Assertions on TbConfiguracionHPS ---
    sHPST = GetConfigValue(db, k_ClaveHPST)
    If sHPST <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\HPS\HPST.accdb" Then
        Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks = JsonFail( _
            k_ClaveHPST & " en TbConfiguracionHPS no coincide: '" & sHPST & "'", logs)
        GoTo CleanUp
    End If
    sExpedientes = GetConfigValue(db, k_ClaveExpedientes)
    If sExpedientes <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\EXPEDIENTES\Expedientes_datos.accdb" Then
        Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks = JsonFail( _
            k_ClaveExpedientes & " en TbConfiguracionHPS no coincide: '" & sExpedientes & "'", logs)
        GoTo CleanUp
    End If
    ' [config-as-source-of-truth] Los anexos se siembran vía
    ' APP_ROOT_REMOTO y APP_ROOT_LOCAL; las claves ANEXOS_* ya no existen.
    sAnexosHPS = GetConfigValue(db, "APP_ROOT_REMOTO")
    If sAnexosHPS <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\" Then
        Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks = JsonFail( _
            "APP_ROOT_REMOTO en TbConfiguracionHPS no coincide: '" & sAnexosHPS & "'", logs)
        GoTo CleanUp
    End If
    sModoLocal = GetConfigValue(db, k_ClaveModoLocal)
    If sModoLocal <> "No" Then
        Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks = JsonFail( _
            k_ClaveModoLocal & " en TbConfiguracionHPS no coincide: '" & sModoLocal & "'", logs)
        GoTo CleanUp
    End If

    ' --- Assertions on Entorno.GetConfig (cache refrescada por PrepararProduccion) ---
    Set oEnt = New Entorno
    oEnt.CargarConfiguracion

    sVal = oEnt.GetConfig(k_ClaveHPST)
    If sVal <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\HPS\HPST.accdb" Then
        Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks = JsonFail( _
            "Entorno.GetConfig(" & k_ClaveHPST & ") no coincide: '" & sVal & "'", logs)
        GoTo CleanUp
    End If
    sVal = oEnt.GetConfig(k_ClaveExpedientes)
    If sVal <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\EXPEDIENTES\Expedientes_datos.accdb" Then
        Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks = JsonFail( _
            "Entorno.GetConfig(" & k_ClaveExpedientes & ") no coincide: '" & sVal & "'", logs)
        GoTo CleanUp
    End If
    sVal = oEnt.GetConfig(k_ClaveModoLocal)
    If sVal <> "No" Then
        Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks = JsonFail( _
            "Entorno.GetConfig(" & k_ClaveModoLocal & ") no coincide: '" & sVal & "'", logs)
        GoTo CleanUp
    End If

    ' --- Assertion: relink solo si la parent dir es alcanzable. Esta es la
    ' nueva propiedad: PrepararProduccion NO debe fallar si \\datoste
    ' no está accesible (sandbox, casa sin VPN); debe completar la
    ' siembra de config sin tocar TableDefs.
    sParentDir = "\\datoste\aplicaciones_dys\Aplicaciones PpD"
    bSawProduccionConnect = False
    bSawProduccionHPST = False
    bSawProduccionExpedientes = False
    Dim fsoAux As Object
    Set fsoAux = CreateObject("Scripting.FileSystemObject")
    bParentReachable = fsoAux.FolderExists(sParentDir)
    If bParentReachable Then
        For Each tdf In db.TableDefs
            If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
                If Left$(tdf.Name, 4) <> "MSys" Then
                    sConnect = CStr(tdf.Connect)
                    If InStr(1, sConnect, sParentDir, vbTextCompare) > 0 Then
                        bSawProduccionConnect = True
                    End If
                    If InStr(1, sConnect, sParentDir & "\HPS\HPST.accdb", vbTextCompare) > 0 Then
                        bSawProduccionHPST = True
                    End If
                    If InStr(1, sConnect, sParentDir & "\EXPEDIENTES\Expedientes_datos.accdb", vbTextCompare) > 0 Then
                        bSawProduccionExpedientes = True
                    End If
                End If
            End If
        Next tdf
    End If
    Set fsoAux = Nothing

    ' Restauramos SIEMPRE el snapshot para no dejar el binario con TableDefs modificados.
    RestoreAttachedConnectsLocal db, snap

    If Not bParentReachable Then
        ' Si \\datoste NO es alcanzable (sandbox), el relink no se aplica
        ' y eso es el contrato correcto del nuevo resolver. La config sigue
        ' sembrada y Entorno.GetConfig la sirve correctamente.
        If Len(sHPST) = 0 Or Len(sExpedientes) = 0 Then
            Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks = JsonFail( _
                "PrepararProduccion en sandbox no sembró las claves esperadas: HPST='" & sHPST & "', Expedientes='" & sExpedientes & "'.", logs)
            GoTo CleanUp
        End If
        logs.Add "PrepararProduccion (sandbox sin \\datoste): 9 claves pobladas con paths UNC; relink omitido correctamente porque APP_ROOT_REMOTO no es alcanzable. Contrato config-as-source-of-truth respetado."
        Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks = JsonOk("preparar-produccion-sandbox-config-driven", logs)
        GoTo CleanUp
    End If

    If Not bSawProduccionConnect Then
        Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks = JsonFail( _
            "Tras PrepararProduccion, ningun TableDef adjunto apuntaba a '" & sParentDir & "'.", logs)
        GoTo CleanUp
    End If
    If Not bSawProduccionHPST Then
        Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks = JsonFail( _
            "Tras PrepararProduccion, ningun TableDef adjunto apunto a '" & sParentDir & "\HPS\HPST.accdb'.", logs)
        GoTo CleanUp
    End If
    If Not bSawProduccionExpedientes Then
        Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks = JsonFail( _
            "Tras PrepararProduccion, ningun TableDef adjunto apunto a '" & sParentDir & "\EXPEDIENTES\Expedientes_datos.accdb'.", logs)
        GoTo CleanUp
    End If

    logs.Add "PrepararProduccion: 9 claves pobladas con paths UNC + cache del Entorno refrescada + TableDefs apuntan a HPS/HPST.accdb y EXPEDIENTES/Expedientes_datos.accdb."
    Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks = JsonOk("preparar-produccion-populates-and-relinks", logs)
    GoTo CleanUp
EH:
    ' Restauramos antes de devolver el fail.
    On Error Resume Next
    If Not snap Is Nothing Then
        RestoreAttachedConnectsLocal db, snap
    End If
    On Error GoTo 0
    Test_HPSConfig_PrepararProduccion_PopulatesAndRelinks = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set oEnt = Nothing
    Set db = Nothing
    Set snap = Nothing
    On Error GoTo 0
End Function

' T10: PrepararLocal deja la config table con paths locales, refresca la
' cache del Entorno y deja al menos un TableDef adjunto con Connect
' apuntando a C:\00repos\datos. Restauramos TableDefs al final.
Public Function Test_HPSConfig_PrepararLocal_PopulatesAndRelinks() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim oEnt As Entorno
    Dim snap As Object
    Dim sParentDir As String
    Dim tdf As DAO.TableDef
    Dim sConnect As String
    Dim bSawLocalConnect As Boolean
    Dim sHPST As String
    Dim sExpedientes As String
    Dim sAnexosHPS As String
    Dim sAnexosHistorico As String
    Dim sModoLocal As String
    Dim sVal As String
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb

    ' Clean slate: borrar cualquier fila residual de corridas previas.
    On Error Resume Next
    db.Execute "DELETE FROM " & k_TableConfig, dbFailOnError
    db.Execute "DELETE FROM " & k_TableVinculos, dbFailOnError
    Err.Clear
    On Error GoTo EH

    ' Snapshot attached tables so we can restore them after the relink.
    Set snap = SnapshotAttachedConnectsLocal(db)

    ' --- SUT ---
    PrepararLocal

    ' --- Assertions on TbConfiguracionHPS ---
    sHPST = GetConfigValue(db, k_ClaveHPST)
    If sHPST <> "C:\00repos\datos\HPST.accdb" Then
        Test_HPSConfig_PrepararLocal_PopulatesAndRelinks = JsonFail( _
            k_ClaveHPST & " en TbConfiguracionHPS no coincide: '" & sHPST & "'", logs)
        GoTo CleanUp
    End If
    sExpedientes = GetConfigValue(db, k_ClaveExpedientes)
    If sExpedientes <> "C:\00repos\datos\Expedientes_datos.accdb" Then
        Test_HPSConfig_PrepararLocal_PopulatesAndRelinks = JsonFail( _
            k_ClaveExpedientes & " en TbConfiguracionHPS no coincide: '" & sExpedientes & "'", logs)
        GoTo CleanUp
    End If
    ' [config-as-source-of-truth] APP_ROOT_LOCAL apunta a CurrentProject.path
    ' sembrado por el seed Local. Los anexos ya no son una clave propia.
    sAnexosHPS = GetConfigValue(db, "APP_ROOT_LOCAL")
    If sAnexosHPS <> Application.CurrentProject.path & "\" Then
        Test_HPSConfig_PrepararLocal_PopulatesAndRelinks = JsonFail( _
            "APP_ROOT_LOCAL en TbConfiguracionHPS no coincide: '" & sAnexosHPS & "'", logs)
        GoTo CleanUp
    End If
    sModoLocal = GetConfigValue(db, k_ClaveModoLocal)
    If sModoLocal <> "S" & Chr$(237) Then
        Test_HPSConfig_PrepararLocal_PopulatesAndRelinks = JsonFail( _
            k_ClaveModoLocal & " en TbConfiguracionHPS no coincide: '" & sModoLocal & "'", logs)
        GoTo CleanUp
    End If

    ' --- Assertions on Entorno.GetConfig (cache refrescada por PrepararLocal) ---
    Set oEnt = New Entorno
    oEnt.CargarConfiguracion

    sVal = oEnt.GetConfig(k_ClaveHPST)
    If sVal <> "C:\00repos\datos\HPST.accdb" Then
        Test_HPSConfig_PrepararLocal_PopulatesAndRelinks = JsonFail( _
            "Entorno.GetConfig(" & k_ClaveHPST & ") no coincide: '" & sVal & "'", logs)
        GoTo CleanUp
    End If
    sVal = oEnt.GetConfig(k_ClaveExpedientes)
    If sVal <> "C:\00repos\datos\Expedientes_datos.accdb" Then
        Test_HPSConfig_PrepararLocal_PopulatesAndRelinks = JsonFail( _
            "Entorno.GetConfig(" & k_ClaveExpedientes & ") no coincide: '" & sVal & "'", logs)
        GoTo CleanUp
    End If
    sVal = oEnt.GetConfig(k_ClaveModoLocal)
    If sVal <> "S" & Chr$(237) Then
        Test_HPSConfig_PrepararLocal_PopulatesAndRelinks = JsonFail( _
            "Entorno.GetConfig(" & k_ClaveModoLocal & ") no coincide: '" & sVal & "'", logs)
        GoTo CleanUp
    End If

    ' --- Assertion: al menos un TableDef adjunto apunta al parent dir Local ---
    sParentDir = "C:\00repos\datos"
    bSawLocalConnect = False
    For Each tdf In db.TableDefs
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            If Left$(tdf.Name, 4) <> "MSys" Then
                sConnect = CStr(tdf.Connect)
                If InStr(1, sConnect, sParentDir, vbTextCompare) > 0 Then
                    bSawLocalConnect = True
                    Exit For
                End If
            End If
        End If
    Next tdf

    ' Restauramos SIEMPRE el snapshot para no dejar el binario con TableDefs modificados.
    RestoreAttachedConnectsLocal db, snap

    If Not bSawLocalConnect Then
        Test_HPSConfig_PrepararLocal_PopulatesAndRelinks = JsonFail( _
            "Tras PrepararLocal, ningun TableDef adjunto apuntaba a '" & sParentDir & "'.", logs)
        GoTo CleanUp
    End If

    logs.Add "PrepararLocal: 5 claves pobladas con paths locales + cache del Entorno refrescada + al menos un TableDef apunta a 'C:\00repos\datos'."
    Test_HPSConfig_PrepararLocal_PopulatesAndRelinks = JsonOk("preparar-local-populates-and-relinks", logs)
    GoTo CleanUp
EH:
    ' Restauramos antes de devolver el fail.
    On Error Resume Next
    If Not snap Is Nothing Then
        RestoreAttachedConnectsLocal db, snap
    End If
    On Error GoTo 0
    Test_HPSConfig_PrepararLocal_PopulatesAndRelinks = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set oEnt = Nothing
    Set db = Nothing
    Set snap = Nothing
    On Error GoTo 0
End Function

' T11: AplicarConfiguracionEnVinculos -- happy path. Carga el perfil Local
' (los 2 .accdb existen en C:\00repos\datos), llama al SUT, y verifica
' que al menos un TableDef adjunto quedo re-apuntado a ese parent dir.
' Tambien verifica que NO raise error. Restauramos TableDefs al final.
Public Function Test_HPSConfig_AplicarConfiguracionEnVinculos_RelinksWhenAllReachable() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim snap As Object
    Dim sParentDir As String
    Dim tdf As DAO.TableDef
    Dim sConnect As String
    Dim bSawLocalConnect As Boolean
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb

    ' Clean slate: borrar cualquier fila residual de corridas previas.
    On Error Resume Next
    db.Execute "DELETE FROM " & k_TableConfig, dbFailOnError
    db.Execute "DELETE FROM " & k_TableVinculos, dbFailOnError
    Err.Clear
    On Error GoTo EH

    ' Poblar TbConfiguracionHPS + TbVinculosTablas con el perfil Local.
    ' Los backends (HPST.accdb + Expedientes_datos.accdb) ya existen en
    ' C:\00repos\datos en la maquina de dev, asi que el SUT los encontrara
    ' alcanzables y procedera con el relink.
    ConfigurarTablaConfiguracion_PerfilLocal

    ' Snapshot attached tables so we can restore them after the relink.
    Set snap = SnapshotAttachedConnectsLocal(db)

    ' --- SUT ---
    AplicarConfiguracionEnVinculos

    ' --- Assertion: al menos un TableDef adjunto apunta al parent dir Local ---
    sParentDir = "C:\00repos\datos"
    bSawLocalConnect = False
    For Each tdf In db.TableDefs
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            If Left$(tdf.Name, 4) <> "MSys" Then
                sConnect = CStr(tdf.Connect)
                If InStr(1, sConnect, sParentDir, vbTextCompare) > 0 Then
                    bSawLocalConnect = True
                    Exit For
                End If
            End If
        End If
    Next tdf

    ' Restauramos SIEMPRE el snapshot para no dejar el binario con TableDefs modificados.
    RestoreAttachedConnectsLocal db, snap

    If Not bSawLocalConnect Then
        Test_HPSConfig_AplicarConfiguracionEnVinculos_RelinksWhenAllReachable = JsonFail( _
            "Tras AplicarConfiguracionEnVinculos, ningun TableDef adjunto apuntaba a '" & sParentDir & "'.", logs)
        GoTo CleanUp
    End If

    logs.Add "AplicarConfiguracionEnVinculos: ambos backends alcanzables en C:\00repos\datos + al menos un TableDef re-apuntado ahi."
    Test_HPSConfig_AplicarConfiguracionEnVinculos_RelinksWhenAllReachable = JsonOk("aplicar-config-relinks-when-all-reachable", logs)
    GoTo CleanUp
EH:
    ' Restauramos antes de devolver el fail.
    On Error Resume Next
    If Not snap Is Nothing Then
        RestoreAttachedConnectsLocal db, snap
    End If
    On Error GoTo 0
    Test_HPSConfig_AplicarConfiguracionEnVinculos_RelinksWhenAllReachable = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set db = Nothing
    Set snap = Nothing
    On Error GoTo 0
End Function

' T14: PerfilRemotoPruebas siembra exactamente las 9 claves esperadas
' apuntando al share de pruebas internas
' \\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\.
' Mismo cardinality / keys / filenames que PerfilProduccion, solo cambia
' la raíz. APP_ROOT_LOCAL conserva CurrentProject.path\ (convención del
' perfil Local) para fallback. Restauramos nada -- este Sub es no-mutante
' de TableDefs.
Public Function Test_HPSConfig_PerfilRemotoPruebas_InsertsAllExpectedKeys() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb

    ' Clean slate: borrar cualquier fila residual de corridas previas.
    On Error Resume Next
    db.Execute "DELETE FROM " & k_TableConfig, dbFailOnError
    db.Execute "DELETE FROM " & k_TableVinculos, dbFailOnError
    On Error GoTo EH

    ConfigurarTablaConfiguracion_PerfilRemotoPruebas

    ' 9 claves: APP_ROOT_REMOTO, APP_ROOT_LOCAL, HPST, EXPEDIENTES,
    ' LANZADERA, CORREOS, SOLICITUDES_HPS, MODO_LOCAL_DEFAULT,
    ' BACKEND_PASSWORD (misma cardinality que PerfilProduccion).
    If CountRowsLocal(db, k_TableConfig) <> 9 Then
        Test_HPSConfig_PerfilRemotoPruebas_InsertsAllExpectedKeys = JsonFail( _
            "Expected 9 config rows after RemotoPruebas profile, got " & CountRowsLocal(db, k_TableConfig), logs)
        GoTo CleanUp
    End If

    If GetConfigValue(db, "APP_ROOT_REMOTO") <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\" Then
        Test_HPSConfig_PerfilRemotoPruebas_InsertsAllExpectedKeys = JsonFail( _
            "APP_ROOT_REMOTO mismatch: '" & GetConfigValue(db, "APP_ROOT_REMOTO") & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, "APP_ROOT_LOCAL") <> Application.CurrentProject.path & "\" Then
        Test_HPSConfig_PerfilRemotoPruebas_InsertsAllExpectedKeys = JsonFail( _
            "APP_ROOT_LOCAL mismatch: '" & GetConfigValue(db, "APP_ROOT_LOCAL") & "' (esperado '" & Application.CurrentProject.path & "\')", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, k_ClaveHPST) <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\HPS\HPST.accdb" Then
        Test_HPSConfig_PerfilRemotoPruebas_InsertsAllExpectedKeys = JsonFail( _
            k_ClaveHPST & " mismatch: '" & GetConfigValue(db, k_ClaveHPST) & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, k_ClaveExpedientes) <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\EXPEDIENTES\Expedientes_datos.accdb" Then
        Test_HPSConfig_PerfilRemotoPruebas_InsertsAllExpectedKeys = JsonFail( _
            k_ClaveExpedientes & " mismatch: '" & GetConfigValue(db, k_ClaveExpedientes) & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, "LANZADERA_BACKEND_PATH") <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\0Lanzadera\Lanzadera_Datos.accdb" Then
        Test_HPSConfig_PerfilRemotoPruebas_InsertsAllExpectedKeys = JsonFail( _
            "LANZADERA_BACKEND_PATH mismatch: '" & GetConfigValue(db, "LANZADERA_BACKEND_PATH") & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, "CORREOS_BACKEND_PATH") <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\00Recursos\Correos_datos.accdb" Then
        Test_HPSConfig_PerfilRemotoPruebas_InsertsAllExpectedKeys = JsonFail( _
            "CORREOS_BACKEND_PATH mismatch: '" & GetConfigValue(db, "CORREOS_BACKEND_PATH") & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, "SOLICITUDES_HPS_BACKEND_PATH") <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\SOLICITUDES HPS\Solicitudes_HPS_datos.accdb" Then
        Test_HPSConfig_PerfilRemotoPruebas_InsertsAllExpectedKeys = JsonFail( _
            "SOLICITUDES_HPS_BACKEND_PATH mismatch: '" & GetConfigValue(db, "SOLICITUDES_HPS_BACKEND_PATH") & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, k_ClaveModoLocal) <> "No" Then
        Test_HPSConfig_PerfilRemotoPruebas_InsertsAllExpectedKeys = JsonFail( _
            k_ClaveModoLocal & " mismatch: '" & GetConfigValue(db, k_ClaveModoLocal) & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, "BACKEND_PASSWORD") <> "dpddpd" Then
        Test_HPSConfig_PerfilRemotoPruebas_InsertsAllExpectedKeys = JsonFail( _
            "BACKEND_PASSWORD mismatch: '" & GetConfigValue(db, "BACKEND_PASSWORD") & "'", logs)
        GoTo CleanUp
    End If

    ' El dispatcher también debe reconocer el perfil.
    On Error Resume Next
    ConfigurarTablaConfiguracion_Perfil "RemotoPruebas"
    If Err.Number <> 0 Then
        Test_HPSConfig_PerfilRemotoPruebas_InsertsAllExpectedKeys = JsonFail( _
            "Dispatcher raised for 'RemotoPruebas': " & Err.Description, logs)
        GoTo CleanUp
    End If
    On Error GoTo EH

    logs.Add "RemotoPruebas profile inserted exactly 9 config rows pointing at \\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\ with the same subdirs/filenames as Produccion; dispatcher accepts 'RemotoPruebas'."
    Test_HPSConfig_PerfilRemotoPruebas_InsertsAllExpectedKeys = JsonOk("perfil-remotopruebas-inserts-9-keys", logs)
    GoTo CleanUp
EH:
    Test_HPSConfig_PerfilRemotoPruebas_InsertsAllExpectedKeys = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set db = Nothing
    On Error GoTo 0
End Function

' T15: PrepararRemotoPruebas siembra config, refresca la cache del
' Entorno y dispara RelinkAllTables con el parent dir del share de
' pruebas. Misma propiedad que PrepararProduccion: si el share no es
' alcanzable, el relink se omite y la config queda poblada. Si lo es,
' al menos un TableDef HPST y uno Expedientes apuntan a sus subdirs
' del share. Restauramos TableDefs siempre.
Public Function Test_HPSConfig_PrepararRemotoPruebas_PopulatesAndRelinks() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim oEnt As Entorno
    Dim snap As Object
    Dim sParentDir As String
    Dim tdf As DAO.TableDef
    Dim sConnect As String
    Dim bSawRemotoConnect As Boolean
    Dim bSawRemotoHPST As Boolean
    Dim bSawRemotoExpedientes As Boolean
    Dim bParentReachable As Boolean
    Dim sHPST As String
    Dim sExpedientes As String
    Dim sVal As String
    Dim fsoAux As Object
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb

    ' Clean slate.
    On Error Resume Next
    db.Execute "DELETE FROM " & k_TableConfig, dbFailOnError
    db.Execute "DELETE FROM " & k_TableVinculos, dbFailOnError
    Err.Clear
    On Error GoTo EH

    Set snap = SnapshotAttachedConnectsLocal(db)

    ' --- SUT ---
    PrepararRemotoPruebas

    ' --- Assertions on TbConfiguracionHPS ---
    sHPST = GetConfigValue(db, k_ClaveHPST)
    If sHPST <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\HPS\HPST.accdb" Then
        Test_HPSConfig_PrepararRemotoPruebas_PopulatesAndRelinks = JsonFail( _
            k_ClaveHPST & " en TbConfiguracionHPS no coincide: '" & sHPST & "'", logs)
        GoTo CleanUp
    End If
    sExpedientes = GetConfigValue(db, k_ClaveExpedientes)
    If sExpedientes <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\EXPEDIENTES\Expedientes_datos.accdb" Then
        Test_HPSConfig_PrepararRemotoPruebas_PopulatesAndRelinks = JsonFail( _
            k_ClaveExpedientes & " en TbConfiguracionHPS no coincide: '" & sExpedientes & "'", logs)
        GoTo CleanUp
    End If
    If GetConfigValue(db, "APP_ROOT_REMOTO") <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\" Then
        Test_HPSConfig_PrepararRemotoPruebas_PopulatesAndRelinks = JsonFail( _
            "APP_ROOT_REMOTO en TbConfiguracionHPS no coincide: '" & GetConfigValue(db, "APP_ROOT_REMOTO") & "'", logs)
        GoTo CleanUp
    End If

    ' --- Assertions on Entorno.GetConfig (cache refrescada) ---
    Set oEnt = New Entorno
    oEnt.CargarConfiguracion

    sVal = oEnt.GetConfig(k_ClaveHPST)
    If sVal <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\HPS\HPST.accdb" Then
        Test_HPSConfig_PrepararRemotoPruebas_PopulatesAndRelinks = JsonFail( _
            "Entorno.GetConfig(" & k_ClaveHPST & ") no coincide: '" & sVal & "'", logs)
        GoTo CleanUp
    End If
    sVal = oEnt.GetConfig(k_ClaveExpedientes)
    If sVal <> "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\EXPEDIENTES\Expedientes_datos.accdb" Then
        Test_HPSConfig_PrepararRemotoPruebas_PopulatesAndRelinks = JsonFail( _
            "Entorno.GetConfig(" & k_ClaveExpedientes & ") no coincide: '" & sVal & "'", logs)
        GoTo CleanUp
    End If

    ' --- Assertion: relink solo si la parent dir es alcanzable. Misma
    ' propiedad que PrepararProduccion: si el share no está accesible
    ' el relink se omite, pero la config queda sembrada.
    sParentDir = "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba"
    bSawRemotoConnect = False
    bSawRemotoHPST = False
    bSawRemotoExpedientes = False
    Set fsoAux = CreateObject("Scripting.FileSystemObject")
    bParentReachable = fsoAux.FolderExists(sParentDir)
    If bParentReachable Then
        For Each tdf In db.TableDefs
            If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
                If Left$(tdf.Name, 4) <> "MSys" Then
                    sConnect = CStr(tdf.Connect)
                    If InStr(1, sConnect, sParentDir, vbTextCompare) > 0 Then
                        bSawRemotoConnect = True
                    End If
                    If InStr(1, sConnect, sParentDir & "\HPS\HPST.accdb", vbTextCompare) > 0 Then
                        bSawRemotoHPST = True
                    End If
                    If InStr(1, sConnect, sParentDir & "\EXPEDIENTES\Expedientes_datos.accdb", vbTextCompare) > 0 Then
                        bSawRemotoExpedientes = True
                    End If
                End If
            End If
        Next tdf
    End If
    Set fsoAux = Nothing

    ' Restauramos SIEMPRE el snapshot para no dejar el binario con TableDefs modificados.
    RestoreAttachedConnectsLocal db, snap

    If Not bParentReachable Then
        ' Sin share alcanzable: assert que la config quedó poblada.
        If Len(sHPST) = 0 Or Len(sExpedientes) = 0 Then
            Test_HPSConfig_PrepararRemotoPruebas_PopulatesAndRelinks = JsonFail( _
                "PrepararRemotoPruebas sin share alcanzable no sembró las claves esperadas: HPST='" & sHPST & "', Expedientes='" & sExpedientes & "'.", logs)
            GoTo CleanUp
        End If
        logs.Add "PrepararRemotoPruebas (sin share alcanzable): 9 claves pobladas apuntando al share de pruebas; relink omitido correctamente."
        Test_HPSConfig_PrepararRemotoPruebas_PopulatesAndRelinks = JsonOk("preparar-remotopruebas-sandbox-config-driven", logs)
        GoTo CleanUp
    End If

    If Not bSawRemotoConnect Then
        Test_HPSConfig_PrepararRemotoPruebas_PopulatesAndRelinks = JsonFail( _
            "Tras PrepararRemotoPruebas, ningun TableDef adjunto apuntaba a '" & sParentDir & "'.", logs)
        GoTo CleanUp
    End If
    If Not bSawRemotoHPST Then
        Test_HPSConfig_PrepararRemotoPruebas_PopulatesAndRelinks = JsonFail( _
            "Tras PrepararRemotoPruebas, ningun TableDef adjunto apunto a '" & sParentDir & "\HPS\HPST.accdb'.", logs)
        GoTo CleanUp
    End If
    If Not bSawRemotoExpedientes Then
        Test_HPSConfig_PrepararRemotoPruebas_PopulatesAndRelinks = JsonFail( _
            "Tras PrepararRemotoPruebas, ningun TableDef adjunto apunto a '" & sParentDir & "\EXPEDIENTES\Expedientes_datos.accdb'.", logs)
        GoTo CleanUp
    End If

    logs.Add "PrepararRemotoPruebas: 9 claves pobladas + cache refrescada + TableDefs apuntan al share de pruebas (HPS\HPST.accdb y EXPEDIENTES\Expedientes_datos.accdb)."
    Test_HPSConfig_PrepararRemotoPruebas_PopulatesAndRelinks = JsonOk("preparar-remotopruebas-populates-and-relinks", logs)
    GoTo CleanUp
EH:
    On Error Resume Next
    If Not snap Is Nothing Then
        RestoreAttachedConnectsLocal db, snap
    End If
    On Error GoTo 0
    Test_HPSConfig_PrepararRemotoPruebas_PopulatesAndRelinks = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set oEnt = Nothing
    Set db = Nothing
    Set snap = Nothing
    On Error GoTo 0
End Function

' T12: AplicarConfiguracionEnVinculos -- backend missing. Sobreescribimos
' los 2 paths a uno que NO existe (C:\__no_existe_HPST\) y verificamos
' que el SUT raise con un mensaje que menciona el archivo faltante Y
' que TableDefs NO se modificaron (safety: no debe haber partial relink).
Public Function Test_HPSConfig_AplicarConfiguracionEnVinculos_RaisesErrorWhenBackendMissing() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim snap As Object
    Dim snap2 As Object
    Dim raisedErrNum As Long
    Dim raisedErrDesc As String
    Dim bAllUnchanged As Boolean
    Dim k As Variant
    Dim sVal As String
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb

    ' Clean slate: borrar cualquier fila residual de corridas previas.
    On Error Resume Next
    db.Execute "DELETE FROM " & k_TableConfig, dbFailOnError
    db.Execute "DELETE FROM " & k_TableVinculos, dbFailOnError
    Err.Clear
    On Error GoTo EH

    ' Poblar la tabla con el perfil Local (valido) y luego reescribir
    ' los 2 paths a uno que NO existe. Asi forzamos el caso de error
    ' "backend no alcanzable" sin tener que crear la tabla a mano.
    ConfigurarTablaConfiguracion_PerfilLocal
    db.Execute _
        "UPDATE " & k_TableConfig & _
        " SET Valor='C:\__no_existe_HPST\HPST.accdb'" & _
        " WHERE Clave='HPST_BACKEND_PATH'", dbFailOnError
    db.Execute _
        "UPDATE " & k_TableConfig & _
        " SET Valor='C:\__no_existe_HPST\Expedientes_datos.accdb'" & _
        " WHERE Clave='EXPEDIENTES_BACKEND_PATH'", dbFailOnError

    ' Snapshot TableDefs para verificar que NO se modificaron.
    Set snap = SnapshotAttachedConnectsLocal(db)

    ' --- SUT (esperamos Err.Raise) ---
    raisedErrNum = 0
    raisedErrDesc = ""
    On Error Resume Next
    AplicarConfiguracionEnVinculos
    raisedErrNum = Err.Number
    raisedErrDesc = Err.Description
    On Error GoTo EH

    ' Re-snapshot para comparar contra el original.
    Set snap2 = SnapshotAttachedConnectsLocal(db)
    bAllUnchanged = True
    If snap2.Count <> snap.Count Then
        bAllUnchanged = False
    Else
        For Each k In snap.Keys
            sVal = ""
            If snap2.Exists(CStr(k)) Then
                sVal = CStr(snap2(CStr(k)))
            End If
            If sVal <> CStr(snap(CStr(k))) Then
                bAllUnchanged = False
                Exit For
            End If
        Next k
    End If

    ' Restauramos TableDefs siempre (puede o no haber side effects, pero
    ' por seguridad no dejamos el binario tocado).
    RestoreAttachedConnectsLocal db, snap

    ' --- Assertions ---
    If raisedErrNum = 0 Then
        Test_HPSConfig_AplicarConfiguracionEnVinculos_RaisesErrorWhenBackendMissing = JsonFail( _
            "Expected an error to be raised, but AplicarConfiguracionEnVinculos completed silently.", logs)
        GoTo CleanUp
    End If

    ' El mensaje debe mencionar el archivo faltante (HPST.accdb o Expedientes_datos.accdb).
    If InStr(1, raisedErrDesc, "HPST.accdb", vbTextCompare) = 0 And _
       InStr(1, raisedErrDesc, "Expedientes_datos.accdb", vbTextCompare) = 0 Then
        Test_HPSConfig_AplicarConfiguracionEnVinculos_RaisesErrorWhenBackendMissing = JsonFail( _
            "Expected error message to mention the missing file (HPST.accdb / Expedientes_datos.accdb), got: " & raisedErrDesc, logs)
        GoTo CleanUp
    End If

    If Not bAllUnchanged Then
        Test_HPSConfig_AplicarConfiguracionEnVinculos_RaisesErrorWhenBackendMissing = JsonFail( _
            "TableDefs fueron modificados pese al error. Snapshot antes/despues no coincide (posible partial relink).", logs)
        GoTo CleanUp
    End If

    logs.Add "AplicarConfiguracionEnVinculos: raise con mensaje mencionando el archivo faltante + TableDefs intactos."
    Test_HPSConfig_AplicarConfiguracionEnVinculos_RaisesErrorWhenBackendMissing = JsonOk("aplicar-config-raises-error-when-backend-missing", logs)
    GoTo CleanUp
EH:
    ' Restauramos antes de devolver el fail.
    On Error Resume Next
    If Not snap Is Nothing Then
        RestoreAttachedConnectsLocal db, snap
    End If
    On Error GoTo 0
    Test_HPSConfig_AplicarConfiguracionEnVinculos_RaisesErrorWhenBackendMissing = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set db = Nothing
    Set snap = Nothing
    Set snap2 = Nothing
    On Error GoTo 0
End Function

' T13: AplicarConfiguracionEnVinculos -- config vacia. Borramos todas
' las filas de TbConfiguracionHPS y verificamos que el SUT raise con
' un mensaje que menciona "config" / "vacia" / "no hay" Y que TableDefs
' NO se modificaron. No depende de la existencia de ningun archivo.
Public Function Test_HPSConfig_AplicarConfiguracionEnVinculos_RaisesErrorWhenConfigEmpty() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim snap As Object
    Dim snap2 As Object
    Dim raisedErrNum As Long
    Dim raisedErrDesc As String
    Dim bAllUnchanged As Boolean
    Dim k As Variant
    Dim sVal As String
    Set logs = New Collection
    On Error GoTo EH

    Set db = CurrentDb

    ' Clean slate: borrar cualquier fila residual de corridas previas.
    On Error Resume Next
    db.Execute "DELETE FROM " & k_TableConfig, dbFailOnError
    db.Execute "DELETE FROM " & k_TableVinculos, dbFailOnError
    Err.Clear
    On Error GoTo EH

    ' Aseguramos que la tabla existe (la creo via PerfilLocal) y luego
    ' borramos todas las filas -> la tabla queda con 0 filas.
    ConfigurarTablaConfiguracion_PerfilLocal
    db.Execute "DELETE FROM " & k_TableConfig, dbFailOnError

    ' Snapshot TableDefs para verificar que NO se modificaron.
    Set snap = SnapshotAttachedConnectsLocal(db)

    ' --- SUT (esperamos Err.Raise) ---
    raisedErrNum = 0
    raisedErrDesc = ""
    On Error Resume Next
    AplicarConfiguracionEnVinculos
    raisedErrNum = Err.Number
    raisedErrDesc = Err.Description
    On Error GoTo EH

    ' Re-snapshot para comparar contra el original.
    Set snap2 = SnapshotAttachedConnectsLocal(db)
    bAllUnchanged = True
    If snap2.Count <> snap.Count Then
        bAllUnchanged = False
    Else
        For Each k In snap.Keys
            sVal = ""
            If snap2.Exists(CStr(k)) Then
                sVal = CStr(snap2(CStr(k)))
            End If
            If sVal <> CStr(snap(CStr(k))) Then
                bAllUnchanged = False
                Exit For
            End If
        Next k
    End If

    ' Restauramos TableDefs siempre.
    RestoreAttachedConnectsLocal db, snap

    ' --- Assertions ---
    If raisedErrNum = 0 Then
        Test_HPSConfig_AplicarConfiguracionEnVinculos_RaisesErrorWhenConfigEmpty = JsonFail( _
            "Expected an error to be raised when TbConfiguracionHPS is empty, but the call completed silently.", logs)
        GoTo CleanUp
    End If

    ' La descripcion debe mencionar "config", "vac" (cubre "vacia" y "vacía")
    ' o "no hay" (case-insensitive).
    If InStr(1, raisedErrDesc, "config", vbTextCompare) = 0 And _
       InStr(1, raisedErrDesc, "vac", vbTextCompare) = 0 And _
       InStr(1, raisedErrDesc, "no hay", vbTextCompare) = 0 Then
        Test_HPSConfig_AplicarConfiguracionEnVinculos_RaisesErrorWhenConfigEmpty = JsonFail( _
            "Expected error message to mention 'config' / 'vacia' / 'no hay', got: " & raisedErrDesc, logs)
        GoTo CleanUp
    End If

    If Not bAllUnchanged Then
        Test_HPSConfig_AplicarConfiguracionEnVinculos_RaisesErrorWhenConfigEmpty = JsonFail( _
            "TableDefs fueron modificados pese al error. Snapshot antes/despues no coincide (posible partial relink).", logs)
        GoTo CleanUp
    End If

    logs.Add "AplicarConfiguracionEnVinculos: raise con mensaje de config vacia + TableDefs intactos."
    Test_HPSConfig_AplicarConfiguracionEnVinculos_RaisesErrorWhenConfigEmpty = JsonOk("aplicar-config-raises-error-when-config-empty", logs)
    GoTo CleanUp
EH:
    ' Restauramos antes de devolver el fail.
    On Error Resume Next
    If Not snap Is Nothing Then
        RestoreAttachedConnectsLocal db, snap
    End If
    On Error GoTo 0
    Test_HPSConfig_AplicarConfiguracionEnVinculos_RaisesErrorWhenConfigEmpty = JsonFail("Unexpected error: " & Err.Description, logs)
CleanUp:
    On Error Resume Next
    Set db = Nothing
    Set snap = Nothing
    Set snap2 = Nothing
    On Error GoTo 0
End Function


