Attribute VB_Name = "Test_Helper_ExpedienteAnexos"
Option Compare Database
Option Explicit
' Tests atómicos para Helper_ExpedienteAnexos (PRUEBA-003 REFAC-3b slice 1).
' Cobertura: BR-18-01 (ValidarEliminacionAnexo), BR-18-02 (EliminarAnexo).
'
' Convenciones:
'   - 0-arg Public Function
'   - JSON return: BuildJsonOk / BuildJsonFail
'   - Fixture-first: cada test que toca datos siembra su propio escenario
'   - IDs deterministicos con prefijo TEST-REFAC-3B- para evitar colisiones
'   - Teardown defensivo en orden inverso de FK (anexo antes que parent)
'
' Regla NO-UI (user feedback 2026-06-16):
'   - Los tests NUNCA deben mostrar UI de Access (MsgBox, dialogs, etc).
'   - Cualquier error en tiempo de ejecucion se convierte a BuildJsonFail via
'     HandleError, no se muestra dialog.
'   - El cleanup es best-effort con On Error Resume Next (cero dialog).
'   - Si el setup falla (no se puede crear archivo/carpeta), el test falla
'     claro, no queda en estado intermedio.

' BR-18-01: ValidarEliminacionAnexo con row sembrada -> True
' Siembra parent en TbExpedientes + 1 row en TbExpedientesAnexos. Llama al
' helper, valida True + sin error. Teardown defensivo.
Public Function Test_Helper_ExpedienteAnexos_ValidarEliminacionAnexo_Valido_DevuelveTrue() As String
    Dim logs(0 To 4) As String
    Dim db As DAO.Database
    Dim motivo As String
    Dim result As Boolean
    Const TEST_PADRE As Long = 99990001
    Const TEST_ANEXO As Long = 99990002

    ' Default: fail. Solo se override si los asserts pasan.
    Test_Helper_ExpedienteAnexos_ValidarEliminacionAnexo_Valido_DevuelveTrue = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    Set db = getdb()

    ' Teardown preventivo
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesAnexos WHERE IDDocumento=" & TEST_ANEXO
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
    Err.Clear
    On Error GoTo HandleError

    ' Seed: parent + 1 anexo
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico) VALUES (" & TEST_PADRE & ", 'TEST-REFAC-3B-VAL')", dbFailOnError
    db.Execute "INSERT INTO TbExpedientesAnexos (IDDocumento, IDExpediente, NombreDocumento) VALUES (" & TEST_ANEXO & ", " & TEST_PADRE & ", 'TEST-REFAC-3B-VAL.pdf')", dbFailOnError
    logs(0) = "Arrange: seed parent (" & TEST_PADRE & ") + 1 anexo (" & TEST_ANEXO & ")"

    ' Act
    result = Helper_ExpedienteAnexos.ValidarEliminacionAnexo(TEST_ANEXO, Nothing, motivo)
    logs(1) = "Act: ValidarEliminacionAnexo called"

    ' Assert
    If result <> True Then
        Test_Helper_ExpedienteAnexos_ValidarEliminacionAnexo_Valido_DevuelveTrue = BuildJsonFail("expected True, got: " & result & " motivo: " & motivo, logs)
        GoTo Cleanup
    End If
    logs(2) = "Assert: True + no motivo"

    Test_Helper_ExpedienteAnexos_ValidarEliminacionAnexo_Valido_DevuelveTrue = BuildJsonOk("ok", logs)
    GoTo Cleanup

Cleanup:
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesAnexos WHERE IDDocumento=" & TEST_ANEXO
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
    Exit Function

HandleError:
    ' Error inesperado: convierte a fail, va al cleanup defensivo.
    Test_Helper_ExpedienteAnexos_ValidarEliminacionAnexo_Valido_DevuelveTrue = BuildJsonFail("unexpected error: " & Err.Description & " (source: " & Err.Source & ")", logs)
    Resume Cleanup
End Function

' BR-18-02: EliminarAnexo con row + file en disco -> True, row gone, file gone
' Setup: crea temp folder UNICO + expediente folder + file (asserts en cada paso).
' Act: llama al helper con p_PathBase apuntando al temp folder (spec deviation).
' Assert: retorna True, row borrada de BD, file borrado de disco.
' Teardown: cleanup defensivo best-effort (cero dialog).
Public Function Test_Helper_ExpedienteAnexos_EliminarAnexo_Valido_BorraFilaYBorraArchivo() As String
    Dim logs(0 To 6) As String
    Dim db As DAO.Database
    Dim tempBase As String
    Dim expFolder As String
    Dim filePath As String
    Dim errMsg As String
    Dim result As Boolean
    Dim rs As DAO.Recordset
    Const TEST_PADRE As Long = 99990100
    Const TEST_ANEXO As Long = 99990101
    Const TEST_FILE As String = "TEST-REFAC-3B-DEL.pdf"

    ' Default: fail. Solo se override si los asserts pasan.
    Test_Helper_ExpedienteAnexos_EliminarAnexo_Valido_BorraFilaYBorraArchivo = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    Set db = getdb()

    ' Setup: temp folder UNICO por test (incluye TEST_ANEXO en el path para
    ' evitar colisiones con runs paralelos / leftover de tests fallidos previos).
    tempBase = Environ$("TEMP") & "\TEST-REFAC-3B-" & TEST_ANEXO & "\"

    ' Pre-clean cualquier leftover de runs previos
    On Error Resume Next
    Kill tempBase & "*.*"
    RmDir StripTrailingSlash(tempBase)
    Err.Clear
    On Error GoTo HandleError

    ' Crear parent folder. MkDir puede fallar si no se puede escribir en TEMP
    ' (permisos, disco lleno, etc). Si falla, HandleError captura.
    MkDir tempBase
    expFolder = tempBase & Format$(TEST_PADRE, "00000")
    MkDir expFolder
    filePath = expFolder & "\" & TEST_FILE

    ' Crear el archivo via Open (VBA nativo, sin FileSystemObject)
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Output As #ff
    Print #ff, "test content for TEST-REFAC-3B"
    Close #ff
    logs(0) = "Arrange: tempBase=" & tempBase & " file=" & filePath

    ' Teardown preventivo (BD only, filesystem ya esta limpio)
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesAnexos WHERE IDDocumento=" & TEST_ANEXO
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
    Err.Clear
    On Error GoTo HandleError

    ' Seed: parent + 1 anexo (dbFailOnError para que falle ruidosamente si FK rota)
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico) VALUES (" & TEST_PADRE & ", 'TEST-REFAC-3B-DEL')", dbFailOnError
    db.Execute "INSERT INTO TbExpedientesAnexos (IDDocumento, IDExpediente, NombreDocumento) VALUES (" & TEST_ANEXO & ", " & TEST_PADRE & ", '" & TEST_FILE & "')", dbFailOnError
    logs(1) = "Arrange: seed parent + 1 anexo row + file on disk"

    ' Act: con p_PathBase = tempBase (spec deviation para sandbox determinista)
    result = Helper_ExpedienteAnexos.EliminarAnexo(TEST_ANEXO, Nothing, tempBase, errMsg)
    logs(2) = "Act: EliminarAnexo called with p_PathBase=" & tempBase

    ' Assert: result is True
    If result <> True Then
        Test_Helper_ExpedienteAnexos_EliminarAnexo_Valido_BorraFilaYBorraArchivo = BuildJsonFail("expected True, got: " & result & " err: " & errMsg, logs)
        GoTo Cleanup
    End If
    logs(3) = "Assert: True"

    ' Assert: row is gone (BD delete = source of truth)
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS N FROM TbExpedientesAnexos WHERE IDDocumento=" & TEST_ANEXO, dbReadOnly)
    If rs!N <> 0 Then
        rs.Close
        Test_Helper_ExpedienteAnexos_EliminarAnexo_Valido_BorraFilaYBorraArchivo = BuildJsonFail("expected row gone, found " & rs!N, logs)
        GoTo Cleanup
    End If
    rs.Close
    logs(4) = "Assert: row deleted from TbExpedientesAnexos"

    ' Assert: file is gone (best-effort delete, verificado en happy path)
    If Dir(filePath) <> "" Then
        Test_Helper_ExpedienteAnexos_EliminarAnexo_Valido_BorraFilaYBorraArchivo = BuildJsonFail("expected file gone, still exists: " & filePath, logs)
        GoTo Cleanup
    End If
    logs(5) = "Assert: file deleted from disk"

    ' All assertions passed
    Test_Helper_ExpedienteAnexos_EliminarAnexo_Valido_BorraFilaYBorraArchivo = BuildJsonOk("ok", logs)
    GoTo Cleanup

Cleanup:
    ' Teardown best-effort. Cero dialog permitido: todos los errores
    ' suprimidos por On Error Resume Next. Si algo falla, queda en el
    ' filesystem y se limpia en el siguiente run.
    On Error Resume Next
    db.Execute "DELETE FROM TbExpedientesAnexos WHERE IDDocumento=" & TEST_ANEXO
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_PADRE
    Err.Clear
    If Len(filePath) > 0 And Dir(filePath) <> "" Then Kill filePath
    Err.Clear
    If Len(expFolder) > 0 And Dir(expFolder, vbDirectory) <> "" Then RmDir StripTrailingSlash(expFolder)
    Err.Clear
    If Len(tempBase) > 0 And Dir(tempBase, vbDirectory) <> "" Then RmDir StripTrailingSlash(tempBase)
    Exit Function

HandleError:
    ' Error inesperado: convierte a fail con descripcion + source, va al cleanup
    ' defensivo. NUNCA muestra dialog (regla del user 2026-06-16).
    Test_Helper_ExpedienteAnexos_EliminarAnexo_Valido_BorraFilaYBorraArchivo = BuildJsonFail("unexpected error: " & Err.Description & " (source: " & Err.Source & ")", logs)
    Resume Cleanup
End Function

' Quita el trailing backslash de un path. RmDir no lo acepta.
Private Function StripTrailingSlash(ByVal p_Path As String) As String
    If Len(p_Path) > 0 And Right$(p_Path, 1) = "\" Then
        StripTrailingSlash = Left$(p_Path, Len(p_Path) - 1)
    Else
        StripTrailingSlash = p_Path
    End If
End Function
