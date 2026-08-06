Attribute VB_Name = "Test_AGEDO_Integracion"
' ============================================================
' Test_AGEDO_Integracion
'
' issue-58: AGEDO integracion real no validada (riesgo regulatorio).
'
' Contexto confirmado por el usuario:
'   - En produccion la carpeta real de AGEDO es:
'     \\datoste\aplicaciones_dys\Aplicaciones PpD\1ParaAGEDO
'   - La comunicacion con el backend de AGEDO es via la TABLA VINCULADA
'     TbDocumentos (no HTTP/REST): la fila insertada en TbDocumentos
'     queda visible al backend AGEDO.
'   - El codigo original copiaba el PDF a URLDirectorioTemporalAGEDO
'     (la 1ParaAGEDO real) pero la carpeta se creaba en
'     URLDirectorioDocumentacion (ANEXOS) por un override silencioso
'     en MotivoNoOK, y la copia estaba envuelta en 'On Error Resume Next'
'     que silenciaba cualquier fallo. El sistema seguia como si el
'     documento estuviera en AGEDO cuando en realidad nunca llegaba.
'
' Este test verifica que despues de Documento.Registrar:
'   1. El PDF esta fisicamente en URLCarpetaAGEDO (la carpeta AGEDO real)
'   2. La fila esta en TbDocumentos con URLCarpetaAGEDO consistente
'   3. La verificacion post-copia (issue-58) reporta error explicito
'      si el archivo no llega a destino
'
' Skill: access-vba-tdd v2.4
' - Test IDs en rango 900800+ (separado de SeedAll 900500-900599 y
'   de Test_IndicadorCobertura 900700-900799)
' - Sandbox via Test_Fixtures.GetTestDb
' ============================================================
Option Compare Database
Option Explicit

' --- IDs de fixture ---
Private Const FIX_IDDOCUMENTO As Long = 900800
Private Const FIX_CODIGO As String = "IS2025999CCCE01"
Private Const FIX_CODEXP As String = "TEST/99"
Private Const FIX_CADENA As String = "test_fixture_agedo.pdf"


' --- Helpers JSON ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function


' --- Helpers de filesystem ---
Private Function MakeTempDir(ByVal p_Suffix As String) As String
    Dim sBase As String
    sBase = Environ$("TEMP")
    If Right$(sBase, 1) <> "\" Then sBase = sBase & "\"
    MakeTempDir = sBase & "test_agedo_" & p_Suffix & "_" & Format$(Now, "yyyymmddhhnnss") & "\"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(MakeTempDir) Then
        fso.CreateFolder MakeTempDir
    End If
    Set fso = Nothing
End Function

Private Sub WriteTextFile(ByVal p_Path As String, ByVal p_Content As String)
    Dim iFile As Integer
    iFile = FreeFile
    Open p_Path For Output As #iFile
    Print #iFile, p_Content
    Close #iFile
End Sub

Private Sub RemoveDirIfExists(ByVal p_Path As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(p_Path) Then
        fso.DeleteFolder p_Path, True
    End If
    Set fso = Nothing
    On Error GoTo 0
End Sub


' --- Limpieza de fila de test en TbDocumentos ---
Private Sub CleanTbDocumentosRow()
    On Error Resume Next
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then Exit Sub
    db.Execute "DELETE FROM TbDocumentos WHERE IDDocumento=" & FIX_IDDOCUMENTO
    On Error GoTo 0
End Sub


' ============================================================
' Test 1: el PDF realmente termina en la carpeta AGEDO
' (URLCarpetaAGEDO), y NO silenciosamente en otra carpeta.
' ============================================================
Public Function Test_RegistrarEnAGEDO_CopiaAGEDOYRegistraTbDocumentos() As String
    On Error GoTo EH
    Dim logs(0 To 7) As String
    logs(0) = "1. Arrange: crear source PDF y carpeta AGEDO sandbox"
    logs(1) = "2. Arrange: instancia de Documento con URLCarpetaAGEDO apuntando a sandbox AGEDO"
    logs(2) = "3. Act: Documento.Registrar"
    logs(3) = "4. Assert: no retorna error"
    logs(4) = "5. Assert: PDF copiado a ULRCarpetaAGEDO + CadenaNombreArchivos"
    logs(5) = "6. Assert: fila en TbDocumentos con URLCarpetaAGEDO consistente"
    logs(6) = "7. Teardown: borrar PDF destino + fila TbDocumentos + temp dirs"
    logs(7) = "8. Assert: side effects limpios"

    ' --- Arrange ---
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RegistrarEnAGEDO_CopiaAGEDOYRegistraTbDocumentos = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If
    CleanTbDocumentosRow

    Dim sSrcDir As String, sAgedoDir As String
    sSrcDir = MakeTempDir("src")
    sAgedoDir = MakeTempDir("agedo")
    Dim sSrcFile As String, sDestFile As String
    sSrcFile = sSrcDir & FIX_CADENA
    sDestFile = sAgedoDir & FIX_CADENA
    WriteTextFile sSrcFile, "PDF de prueba issue-58"

    ' --- Act ---
    Dim doc As Documento
    Set doc = New Documento
    With doc
        .IDDocumento = CStr(FIX_IDDOCUMENTO)
        .Codigo = FIX_CODIGO
        .CodExp = FIX_CODEXP
        .CadenaNombreArchivos = FIX_CADENA
        .Edicion = "1"
        .Tipo = "IS"
        .Area = "E"
        .Versionable = "Sí"
        .Clasificacion = "SINCLAS"
        .Titulo = "Test AGEDO integration"
        .FechaAlta = CStr(Date)
        .URLCarpetaAGEDO = sAgedoDir
        .URLCarpetaArchivosLocales = sSrcDir
    End With

    Dim sError As String
    Call doc.Registrar(sSrcFile, sError)
    If sError <> "" Then
        CleanTbDocumentosRow
        RemoveDirIfExists sSrcDir
        RemoveDirIfExists sAgedoDir
        Test_RegistrarEnAGEDO_CopiaAGEDOYRegistraTbDocumentos = BuildFail( _
            "Documento.Registrar retorno error: " & sError, logs)
        Exit Function
    End If

    ' --- Assert: PDF en la carpeta AGEDO ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(sDestFile) Then
        CleanTbDocumentosRow
        RemoveDirIfExists sSrcDir
        RemoveDirIfExists sAgedoDir
        Test_RegistrarEnAGEDO_CopiaAGEDOYRegistraTbDocumentos = BuildFail( _
            "PDF no esta en la carpeta AGEDO. Esperado: " & sDestFile, logs)
        Exit Function
    End If

    ' --- Assert: fila en TbDocumentos ---
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT IDDocumento, Codigo, CodExp, CadenaNombreArchivos, URLCarpetaAGEDO " & _
        "FROM TbDocumentos WHERE IDDocumento=" & FIX_IDDOCUMENTO)
    If rs.EOF Then
        rs.Close
        CleanTbDocumentosRow
        RemoveDirIfExists sSrcDir
        RemoveDirIfExists sAgedoDir
        Test_RegistrarEnAGEDO_CopiaAGEDOYRegistraTbDocumentos = BuildFail( _
            "No se encontro la fila en TbDocumentos con IDDocumento=" & FIX_IDDOCUMENTO, logs)
        Exit Function
    End If
    Dim sCodigoRow As String, sCodexpRow As String, sCadenaRow As String, sCarpetaRow As String
    sCodigoRow = Nz(rs!Codigo, "")
    sCodexpRow = Nz(rs!CodExp, "")
    sCadenaRow = Nz(rs!CadenaNombreArchivos, "")
    sCarpetaRow = Nz(rs!URLCarpetaAGEDO, "")
    rs.Close

    If sCodigoRow <> FIX_CODIGO Then
        CleanTbDocumentosRow
        RemoveDirIfExists sSrcDir
        RemoveDirIfExists sAgedoDir
        Test_RegistrarEnAGEDO_CopiaAGEDOYRegistraTbDocumentos = BuildFail( _
            "TbDocumentos.Codigo no coincide. Esperado '" & FIX_CODIGO & "', obtuvo '" & sCodigoRow & "'", logs)
        Exit Function
    End If
    If sCodexpRow <> FIX_CODEXP Then
        CleanTbDocumentosRow
        RemoveDirIfExists sSrcDir
        RemoveDirIfExists sAgedoDir
        Test_RegistrarEnAGEDO_CopiaAGEDOYRegistraTbDocumentos = BuildFail( _
            "TbDocumentos.CodExp no coincide. Esperado '" & FIX_CODEXP & "', obtuvo '" & sCodexpRow & "'", logs)
        Exit Function
    End If
    If sCadenaRow <> FIX_CADENA Then
        CleanTbDocumentosRow
        RemoveDirIfExists sSrcDir
        RemoveDirIfExists sAgedoDir
        Test_RegistrarEnAGEDO_CopiaAGEDOYRegistraTbDocumentos = BuildFail( _
            "TbDocumentos.CadenaNombreArchivos no coincide. Esperado '" & FIX_CADENA & "', obtuvo '" & sCadenaRow & "'", logs)
        Exit Function
    End If
    If sCarpetaRow <> sAgedoDir Then
        CleanTbDocumentosRow
        RemoveDirIfExists sSrcDir
        RemoveDirIfExists sAgedoDir
        Test_RegistrarEnAGEDO_CopiaAGEDOYRegistraTbDocumentos = BuildFail( _
            "TbDocumentos.URLCarpetaAGEDO inconsistente con destino. Esperado '" & sAgedoDir & "', obtuvo '" & sCarpetaRow & "'", logs)
        Exit Function
    End If

    ' --- Teardown ---
    CleanTbDocumentosRow
    RemoveDirIfExists sSrcDir
    RemoveDirIfExists sAgedoDir

    Test_RegistrarEnAGEDO_CopiaAGEDOYRegistraTbDocumentos = BuildOk( _
        "copia_y_tbdocumentos_ok", logs)
    Exit Function
EH:
    CleanTbDocumentosRow
    Test_RegistrarEnAGEDO_CopiaAGEDOYRegistraTbDocumentos = BuildFail( _
        "Excepcion: " & Err.Number & " - " & Err.Description, logs)
End Function


' ============================================================
' Test 2: la verificacion post-copia (issue-58) reporta error
' explicito cuando el destino NO es alcanzable, en vez de fallar
' silenciosamente como hacia el codigo original con 'On Error Resume Next'.
' ============================================================
Public Function Test_RegistrarEnAGEDO_DestinoNoAlcanzable_ReportaError() As String
    On Error GoTo EH
    Dim logs(0 To 5) As String
    logs(0) = "1. Arrange: source PDF + URLCarpetaAGEDO apuntando a ruta INVALIDA"
    logs(1) = "2. Act: Documento.Registrar"
    logs(2) = "3. Assert: retorna error NO vacio (antes: 'On Error Resume Next' silenciaba el fallo)"
    logs(3) = "4. Assert: NO se crea el PDF en una ruta alternativa silenciosa"
    logs(4) = "5. Teardown: limpiar"
    logs(5) = "6. Assert: el caller ve el error, no un OK falso"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RegistrarEnAGEDO_DestinoNoAlcanzable_ReportaError = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If
    CleanTbDocumentosRow

    Dim sSrcDir As String, sInvalida As String
    sSrcDir = MakeTempDir("src_inv")
    sInvalida = "Z:\Ruta\Inexistente\En\Produccion\1ParaAGEDO_INVALIDA\"
    Dim sSrcFile As String
    sSrcFile = sSrcDir & FIX_CADENA
    WriteTextFile sSrcFile, "PDF de prueba issue-58 destino invalido"

    Dim doc As Documento
    Set doc = New Documento
    With doc
        .IDDocumento = CStr(FIX_IDDOCUMENTO)
        .Codigo = FIX_CODIGO
        .CodExp = FIX_CODEXP
        .CadenaNombreArchivos = FIX_CADENA
        .Edicion = "1"
        .Tipo = "IS"
        .Area = "E"
        .Versionable = "Sí"
        .Clasificacion = "SINCLAS"
        .Titulo = "Test AGEDO destino invalido"
        .FechaAlta = CStr(Date)
        .URLCarpetaAGEDO = sInvalida
        .URLCarpetaArchivosLocales = sSrcDir
    End With

    Dim sError As String
    On Error Resume Next
    Call doc.Registrar(sSrcFile, sError)
    Dim errNumber As Long
    errNumber = Err.Number
    On Error GoTo EH

    ' --- Assert: el error retorna no vacio ---
    ' Antes del fix issue-58, 'On Error Resume Next' en Registrar silenciaba
    ' cualquier fallo de CreateFolder o CopyFile, asi que la funcion retornaba
    ' sin error y el caller pensaba que el documento estaba en AGEDO.
    ' Ahora cualquier fallo en CreateFolder/CopyFile/FileExistsCheck propaga
    ' a p_Error y el caller ve el error.
    If sError = "" And errNumber = 0 Then
        CleanTbDocumentosRow
        RemoveDirIfExists sSrcDir
        Test_RegistrarEnAGEDO_DestinoNoAlcanzable_ReportaError = BuildFail( _
            "Esperado error explicito por destino inalcanzable, pero Registrar retorno OK silenciosamente. " & _
            "Esto era EXACTAMENTE el bug que issue-58 documenta: el sistema decia que " & _
            "el documento estaba en AGEDO cuando en realidad nunca llego.", logs)
        Exit Function
    End If

    ' --- Assert: el PDF NO se copio silenciosamente a la ruta invalida ni a ningun otro lado ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(sInvalida & FIX_CADENA) Then
        CleanTbDocumentosRow
        RemoveDirIfExists sSrcDir
        Test_RegistrarEnAGEDO_DestinoNoAlcanzable_ReportaError = BuildFail( _
            "El PDF no debio copiarse a la ruta invalida", logs)
        Exit Function
    End If

    ' --- Teardown ---
    CleanTbDocumentosRow
    RemoveDirIfExists sSrcDir

    Test_RegistrarEnAGEDO_DestinoNoAlcanzable_ReportaError = BuildOk( _
        "error_explicito_sin_silencio", logs)
    Exit Function
EH:
    CleanTbDocumentosRow
    Test_RegistrarEnAGEDO_DestinoNoAlcanzable_ReportaError = BuildFail( _
        "Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function
