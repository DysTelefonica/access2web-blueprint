Attribute VB_Name = "Test_DocumentTemplateMapping"
Option Compare Database
Option Explicit

' ============================================================================
' MODULE: Test_DocumentTemplateMapping.bas
' PURPOSE: Read-only contract test between real tbMapeoCampos rows and real
'          Word template FormFields. This slice must not mutate configuration,
'          business data, mappings, or templates.
' ============================================================================

Private Const TEMPLATE_PC As String = "PC"
Private Const TEMPLATE_PCSUB As String = "PCSUB"
Private Const TEMPLATE_CDCA As String = "CDCA"
Private Const TEMPLATE_CDCASUB As String = "CDCASUB"
Private Const MAX_FAILURE_LOGS As Long = 80

Public Function Test_DTM_AllMappedFields() As String
    ' Legacy aggregator. Kept for backwards compatibility; splits into 4 atoms in the
    ' strict TDD manifest. Individual atoms: Test_DTM_PC_AllMappedFields etc.
    Test_DTM_AllMappedFields = Test_DocumentTemplateMapping_AllTemplatesHaveMappedFields()
End Function

Public Function Test_DTM_PC_AllMappedFields() As String
    Test_DTM_PC_AllMappedFields = Test_DocumentTemplateMapping_SingleTemplate(TEMPLATE_PC)
End Function

Public Function Test_DTM_PCSUB_AllMappedFields() As String
    Test_DTM_PCSUB_AllMappedFields = Test_DocumentTemplateMapping_SingleTemplate(TEMPLATE_PCSUB)
End Function

Public Function Test_DTM_CDCA_AllMappedFields() As String
    Test_DTM_CDCA_AllMappedFields = Test_DocumentTemplateMapping_SingleTemplate(TEMPLATE_CDCA)
End Function

Public Function Test_DTM_CDCASUB_AllMappedFields() As String
    Test_DTM_CDCASUB_AllMappedFields = Test_DocumentTemplateMapping_SingleTemplate(TEMPLATE_CDCASUB)
End Function

Public Function Test_DTM_StrictMissingFieldValidation() As String
    On Error GoTo EH

    Dim logs() As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim svc As DocumentoServicio
    Dim rutaPlantilla As String
    Dim rutaDestino As String
    Dim missingField As String
    Dim caughtNumber As Long
    Dim caughtDescription As String
    Dim missingContext As String
    Dim cleanupError As String
    Dim okResult As Boolean
    Dim resultMessage As String

    logs = NewLogArray()
    rutaPlantilla = ResolveTemplatePath(TEMPLATE_PC)
    rutaDestino = CurrentProject.Path & "\tmp_strict_missing_field_validation.docx"
    missingField = "__CONDOR_TEST_CAMPO_WORD_AUSENTE__"

    AppendLog logs, "Arrange: open a real Word template read-only; do not mutate tbMapeoCampos or templates."
    AppendLog logs, "Arrange: missing mapped field -> " & missingField

    If Not FileExists(rutaPlantilla) Then
        resultMessage = "Template file not found: " & rutaPlantilla
        GoTo CleanExit
    End If

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0
    wordApp.AutomationSecurity = 3
    Set wordDoc = wordApp.Documents.Open(FileName:=rutaPlantilla, ConfirmConversions:=False, ReadOnly:=True, _
        AddToRecentFiles:=False, PasswordDocument:="", Revert:=False, Visible:=False, OpenAndRepair:=False, _
        NoEncodingDialog:=True)
    Set svc = New DocumentoServicio

    AppendLog logs, "Act: call DocumentoServicio.ValidarCampoWordMapeado with a controlled absent field."

    On Error GoTo CaptureValidationError
    svc.ValidarCampoWordMapeado wordDoc, missingField, TEMPLATE_PC, "campoTablaPrueba", rutaPlantilla, rutaDestino
    On Error GoTo EH

    resultMessage = "ValidarCampoWordMapeado did not raise for a missing mapped Word field."
    GoTo CleanExit

CaptureValidationError:
    caughtNumber = Err.Number
    caughtDescription = Err.Description
    Err.Clear
    On Error GoTo EH

    AppendLog logs, "Assert: captured error " & CStr(caughtNumber) & " -> " & caughtDescription

    If caughtNumber <> 513 Then
        resultMessage = "Expected validation error 513, got " & CStr(caughtNumber) & ": " & caughtDescription
        GoTo CleanExit
    End If

    missingContext = MissingStrictValidationContext(caughtDescription, TEMPLATE_PC, "campoTablaPrueba", _
        missingField, rutaPlantilla, rutaDestino)
    If Len(missingContext) > 0 Then
        resultMessage = "Strict missing-field error lacks context: " & missingContext & ". Error: " & caughtDescription
        GoTo CleanExit
    End If

    okResult = True
    resultMessage = "strict missing field validation error includes plantilla, campoTabla, campoWord, rutaPlantilla and rutaDestino"

CleanExit:
    CleanupWordResources wordDoc, wordApp, cleanupError
    If Len(cleanupError) > 0 Then AppendLog logs, "Cleanup warning: " & cleanupError
    Set svc = Nothing

    If okResult Then
        Test_DTM_StrictMissingFieldValidation = TestHelper.BuildJsonOk(resultMessage, logs)
    Else
        Test_DTM_StrictMissingFieldValidation = TestHelper.BuildJsonFail(resultMessage, logs)
    End If
    Exit Function

EH:
    AppendLog logs, "ERR " & Err.Number & " - " & Err.Description
    resultMessage = Err.Description
    Resume CleanExit
End Function

Public Function Test_DocumentTemplateMapping_AllTemplatesHaveMappedFields() As String
    On Error GoTo EH

    Dim logs() As String
    Dim checkedCount As Long
    Dim failureCount As Long
    Dim firstFailure As String

    logs = NewLogArray()
    AppendLog logs, "Arrange: read real tbMapeoCampos via getdb() using DAO snapshot recordsets only."
    AppendLog logs, "Arrange: open real Word templates read-only/invisible; no mapping or business data writes."
    AppendLog logs, "Act: validate PC, PCSUB, CDCA and CDCASUB mapped FormFields, including _extN continuation fields."

    ValidateTemplateMappings TEMPLATE_PC, ResolveTemplatePath(TEMPLATE_PC), checkedCount, failureCount, firstFailure, logs
    ValidateTemplateMappings TEMPLATE_PCSUB, ResolveTemplatePath(TEMPLATE_PCSUB), checkedCount, failureCount, firstFailure, logs
    ValidateTemplateMappings TEMPLATE_CDCA, ResolveTemplatePath(TEMPLATE_CDCA), checkedCount, failureCount, firstFailure, logs
    ValidateTemplateMappings TEMPLATE_CDCASUB, ResolveTemplatePath(TEMPLATE_CDCASUB), checkedCount, failureCount, firstFailure, logs

    AppendLog logs, "Assert: mapped Word fields checked: " & CStr(checkedCount)

    If failureCount > 0 Then
        Test_DocumentTemplateMapping_AllTemplatesHaveMappedFields = TestHelper.BuildJsonFail( _
            CStr(failureCount) & " document-template mapping contract failure(s). First: " & firstFailure, logs)
        Exit Function
    End If

    If checkedCount = 0 Then
        Test_DocumentTemplateMapping_AllTemplatesHaveMappedFields = TestHelper.BuildJsonFail( _
            "No tbMapeoCampos rows were checked for PC/PCSUB/CDCA/CDCASUB.", logs)
        Exit Function
    End If

    Test_DocumentTemplateMapping_AllTemplatesHaveMappedFields = TestHelper.BuildJsonOk(CStr(checkedCount), logs)
    Exit Function

EH:
    AppendLog logs, "ERR " & Err.Number & " - " & Err.Description
    Test_DocumentTemplateMapping_AllTemplatesHaveMappedFields = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_DocumentTemplateMapping_SingleTemplate(ByVal plantilla As String) As String
    ' One-template slice. Reads FormFields directly from word/document.xml inside
    ' the .docx (via PowerShell Expand-Archive, no Word COM), validates them
    ' against tbMapeoCampos. Designed to fit within Dysflow's per-call timeout.
    On Error GoTo EH

    Dim logs() As String
    Dim checkedCount As Long
    Dim failureCount As Long
    Dim firstFailure As String
    Dim rutaPlantilla As String
    Dim xmlContent As String
    Dim formFields As Object

    logs = NewLogArray()
    rutaPlantilla = ResolveTemplatePath(plantilla)
    AppendLog logs, "Arrange: read word/document.xml from plantilla=" & plantilla & " path=" & rutaPlantilla

    If Not FileExists(rutaPlantilla) Then
        Test_DocumentTemplateMapping_SingleTemplate = TestHelper.BuildJsonFail( _
            plantilla & ": template file not found at " & rutaPlantilla, logs)
        Exit Function
    End If

    xmlContent = ExtractWordDocumentXml(rutaPlantilla)
    If Len(xmlContent) = 0 Then
        Test_DocumentTemplateMapping_SingleTemplate = TestHelper.BuildJsonFail( _
            plantilla & ": failed to extract word/document.xml (extraction timed out or returned empty)", logs)
        Exit Function
    End If
    AppendLog logs, "Arrange: extracted " & Len(xmlContent) & " chars of XML"

    Set formFields = LoadFormFieldNamesFromDocxXml(rutaPlantilla)
    If formFields Is Nothing Then
        Test_DocumentTemplateMapping_SingleTemplate = TestHelper.BuildJsonFail( _
            plantilla & ": failed to load FormField names from XML", logs)
        Exit Function
    End If
    AppendLog logs, "Arrange: parsed " & formFields.Count & " FormField names from XML"

    ValidateMappingsFromXml plantilla, rutaPlantilla, formFields, checkedCount, failureCount, firstFailure, logs

    AppendLog logs, "Assert: mapped Word fields checked: " & CStr(checkedCount)

    If failureCount > 0 Then
        Test_DocumentTemplateMapping_SingleTemplate = TestHelper.BuildJsonFail( _
            plantilla & ": " & CStr(failureCount) & " document-template mapping contract failure(s). First: " & firstFailure, logs)
        Exit Function
    End If

    If checkedCount = 0 Then
        Test_DocumentTemplateMapping_SingleTemplate = TestHelper.BuildJsonFail( _
            plantilla & ": no tbMapeoCampos rows were checked.", logs)
        Exit Function
    End If

    Test_DocumentTemplateMapping_SingleTemplate = TestHelper.BuildJsonOk(plantilla & "=" & CStr(checkedCount), logs)
    Exit Function

EH:
    AppendLog logs, "ERR " & Err.Number & " - " & Err.Description
    Test_DocumentTemplateMapping_SingleTemplate = TestHelper.BuildJsonFail(plantilla & ": " & Err.Description, logs)
End Function

' Validate mappings by reading tbMapeoCampos and comparing against the dictionary
' of FormField names extracted from word/document.xml. No Word COM.
Private Sub ValidateMappingsFromXml(ByVal plantilla As String, ByVal rutaPlantilla As String, _
    ByVal formFields As Object, ByRef checkedCount As Long, ByRef failureCount As Long, _
    ByRef firstFailure As String, ByRef logs() As String)

    On Error GoTo EH

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim rowCount As Long

    Set db = getdb()
    sql = "SELECT nombrePlantilla, nombreCampoTabla, nombreCampoWord, numExtensiones " & _
          "FROM tbMapeoCampos " & _
          "WHERE nombrePlantilla = " & TestHelper.SqlStr(plantilla) & " " & _
          "ORDER BY nombreCampoTabla, nombreCampoWord"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    Do While Not rs.EOF
        rowCount = rowCount + 1
        ValidateMappedField plantilla, Nz(rs!nombreCampoTabla, vbNullString), Nz(rs!nombreCampoWord, vbNullString), _
            CLng(Nz(rs!numExtensiones, 0)), rutaPlantilla, formFields, checkedCount, failureCount, firstFailure, logs
        rs.MoveNext
    Loop

    If rowCount = 0 Then
        AddFailure failureCount, firstFailure, logs, plantilla, "(mapping)", "(mapping)", 0, rutaPlantilla, "no tbMapeoCampos rows for template"
    Else
        AppendLog logs, "Assert: " & plantilla & " mapping rows checked: " & CStr(rowCount)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0
    Exit Sub

EH:
    AddFailure failureCount, firstFailure, logs, plantilla, "(runtime)", "(runtime)", 0, rutaPlantilla, _
        "runtime error " & Err.Number & ": " & Err.Description
    Resume CleanExit
End Sub

Private Sub ValidateTemplateMappings(ByVal plantilla As String, ByVal rutaPlantilla As String, _
    ByRef checkedCount As Long, ByRef failureCount As Long, ByRef firstFailure As String, ByRef logs() As String)

    On Error GoTo EH

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim formFields As Object
    Dim sql As String
    Dim rowCount As Long
    Dim cleanupError As String

    AppendLog logs, "Arrange: " & plantilla & " template path -> " & rutaPlantilla

    If Not FileExists(rutaPlantilla) Then
        AddFailure failureCount, firstFailure, logs, plantilla, "(template)", "(template)", 0, rutaPlantilla, "template file not found"
        Exit Sub
    End If

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0
    wordApp.AutomationSecurity = 3
    Set wordDoc = wordApp.Documents.Open(FileName:=rutaPlantilla, ConfirmConversions:=False, ReadOnly:=True, _
        AddToRecentFiles:=False, PasswordDocument:="", Revert:=False, Visible:=False, OpenAndRepair:=False, _
        NoEncodingDialog:=True)
    Set formFields = LoadFormFieldNames(wordDoc)

    Set db = getdb()
    sql = "SELECT nombrePlantilla, nombreCampoTabla, nombreCampoWord, numExtensiones " & _
          "FROM tbMapeoCampos " & _
          "WHERE nombrePlantilla = " & TestHelper.SqlStr(plantilla) & " " & _
          "ORDER BY nombreCampoTabla, nombreCampoWord"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    Do While Not rs.EOF
        rowCount = rowCount + 1
        ValidateMappedField plantilla, Nz(rs!nombreCampoTabla, vbNullString), Nz(rs!nombreCampoWord, vbNullString), _
            CLng(Nz(rs!numExtensiones, 0)), rutaPlantilla, formFields, checkedCount, failureCount, firstFailure, logs
        rs.MoveNext
    Loop

    If rowCount = 0 Then
        AddFailure failureCount, firstFailure, logs, plantilla, "(mapping)", "(mapping)", 0, rutaPlantilla, "no tbMapeoCampos rows for template"
    Else
        AppendLog logs, "Assert: " & plantilla & " mapping rows checked: " & CStr(rowCount)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0
    CleanupWordResources wordDoc, wordApp, cleanupError
    If Len(cleanupError) > 0 Then AppendLog logs, "Cleanup warning: " & plantilla & " -> " & cleanupError
    Set formFields = Nothing
    Exit Sub

EH:
    AddFailure failureCount, firstFailure, logs, plantilla, "(runtime)", "(runtime)", 0, rutaPlantilla, _
        "runtime error " & Err.Number & ": " & Err.Description
    Resume CleanExit
End Sub

Private Sub ValidateMappedField(ByVal plantilla As String, ByVal campoTabla As String, _
    ByVal campoWord As String, ByVal numExtensiones As Long, ByVal rutaPlantilla As String, _
    ByVal formFields As Object, ByRef checkedCount As Long, ByRef failureCount As Long, _
    ByRef firstFailure As String, ByRef logs() As String)

    Dim extIndex As Long
    Dim extField As String

    If Len(Trim$(campoWord)) = 0 Then
        AddFailure failureCount, firstFailure, logs, plantilla, campoTabla, campoWord, 0, rutaPlantilla, "empty nombreCampoWord"
        Exit Sub
    End If

    checkedCount = checkedCount + 1
    If Not formFields.Exists(campoWord) Then
        AddFailure failureCount, firstFailure, logs, plantilla, campoTabla, campoWord, 0, rutaPlantilla, "base FormField missing"
    End If

    For extIndex = 1 To numExtensiones
        extField = campoWord & "_ext" & CStr(extIndex)
        checkedCount = checkedCount + 1
        If Not formFields.Exists(extField) Then
            AddFailure failureCount, firstFailure, logs, plantilla, campoTabla, extField, extIndex, rutaPlantilla, "extension FormField missing"
        End If
    Next extIndex
End Sub

Private Function LoadFormFieldNames(ByVal wordDoc As Object) As Object
    ' Legacy COM-based loader kept for callers that still pass an open Word doc.
    Dim names As Object
    Dim formField As Object

    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = 1

    For Each formField In wordDoc.FormFields
        If Len(Trim$(formField.Name)) > 0 Then names(formField.Name) = True
    Next formField

    Set LoadFormFieldNames = names
End Function

' Read the .docx as a zip and parse word/document.xml directly. No Word COM, so
' no 30+ s COM overhead per template. Used by ValidateTemplateMappings when called
' from Test_DocumentTemplateMapping_SingleTemplate (the strict atom path).
' Returns a Scripting.Dictionary with FormField names as keys (TextCompare).
Private Function LoadFormFieldNamesFromDocxXml(ByVal rutaPlantilla As String) As Object
    Dim names As Object
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = 1

    Dim xmlContent As String
    xmlContent = ExtractWordDocumentXml(rutaPlantilla)
    If Len(xmlContent) = 0 Then
        Set LoadFormFieldNamesFromDocxXml = names
        Exit Function
    End If

    Dim searchFrom As Long
    searchFrom = 1
    Do
        Dim startPos As Long
        startPos = InStr(searchFrom, xmlContent, "<w:ffData", vbTextCompare)
        If startPos = 0 Then Exit Do

        Dim endPos As Long
        endPos = InStr(startPos, xmlContent, "</w:ffData>", vbTextCompare)
        If endPos = 0 Then
            searchFrom = startPos + 8
        Else
            Dim ffBlock As String
            ffBlock = Mid$(xmlContent, startPos, endPos - startPos + Len("</w:ffData>"))
            Dim nameStart As Long
            nameStart = InStr(1, ffBlock, "w:name w:val=""", vbTextCompare)
            If nameStart > 0 Then
                Dim valStart As Long
                valStart = nameStart + Len("w:name w:val=""")
                Dim valEnd As Long
                valEnd = InStr(valStart, ffBlock, """", vbTextCompare)
                If valEnd > 0 Then
                    Dim fieldName As String
                    fieldName = Mid$(ffBlock, valStart, valEnd - valStart)
                    If Len(Trim$(fieldName)) > 0 And Not names.Exists(fieldName) Then
                        names(fieldName) = True
                    End If
                End If
            End If
            searchFrom = endPos + Len("</w:ffData>")
        End If
    Loop

    Set LoadFormFieldNamesFromDocxXml = names
End Function

' Extract word/document.xml from a .docx using PowerShell Expand-Archive.
' Avoids Word COM entirely. WScript.Shell.Run with bWaitOnReturn=True hangs when
' called from inside Access COM embedded in the Dysflow runner, so we use
' Shell.Application.Exec which is non-blocking and poll for completion.
Private Function ExtractWordDocumentXml(ByVal docxPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempDir As String
    tempDir = Environ$("TEMP") & "\condor_dtm_" & Format(Now, "yyyymmddhhnnssfff") & "\"
    On Error Resume Next
    fso.DeleteFolder tempDir, True
    On Error GoTo 0
    fso.CreateFolder tempDir

    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")
    Dim psCmd As String
    psCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ""Expand-Archive -Path '" & _
        docxPath & "' -DestinationPath '" & tempDir & "' -Force"""
    On Error Resume Next
    Dim ps As Object
    Set ps = shellApp.Shell(psCmd, 0, False, True)
    On Error GoTo 0

    Dim xmlPath As String
    xmlPath = tempDir & "word\document.xml"
    Dim waitStart As Double
    waitStart = Timer
    Do
        DoEvents
        If fso.FileExists(xmlPath) Then Exit Do
        If Not ps Is Nothing Then
            If ps.Status = 1 Then Exit Do  ' 1=Done
        End If
        If Timer - waitStart > 15 Then Exit Do
    Loop

    Dim result As String
    If fso.FileExists(xmlPath) Then
        Dim ts As Object
        Set ts = fso.OpenTextFile(xmlPath, 1, False)
        result = ts.ReadAll
        ts.Close
    End If

    On Error Resume Next
    fso.DeleteFolder tempDir, True
    On Error GoTo 0

    ExtractWordDocumentXml = result
End Function

Private Function MissingStrictValidationContext(ByVal errorText As String, ByVal plantilla As String, _
    ByVal campoTabla As String, ByVal campoWord As String, ByVal rutaPlantilla As String, _
    ByVal rutaDestino As String) As String

    If InStr(1, errorText, "Plantilla='" & plantilla & "'", vbTextCompare) = 0 Then
        MissingStrictValidationContext = "plantilla"
        Exit Function
    End If

    If InStr(1, errorText, "CampoTabla='" & campoTabla & "'", vbTextCompare) = 0 Then
        MissingStrictValidationContext = "campoTabla"
        Exit Function
    End If

    If InStr(1, errorText, "CampoWord='" & campoWord & "'", vbTextCompare) = 0 Then
        MissingStrictValidationContext = "campoWord"
        Exit Function
    End If

    If InStr(1, errorText, "RutaPlantilla='" & rutaPlantilla & "'", vbTextCompare) = 0 Then
        MissingStrictValidationContext = "rutaPlantilla"
        Exit Function
    End If

    If InStr(1, errorText, "RutaDestino='" & rutaDestino & "'", vbTextCompare) = 0 Then
        MissingStrictValidationContext = "rutaDestino"
        Exit Function
    End If

    MissingStrictValidationContext = vbNullString
End Function

Private Function ResolveTemplatePath(ByVal plantilla As String) As String
    Dim fileName As String

    Select Case UCase$(Trim$(plantilla))
        Case TEMPLATE_PC
            fileName = "PC.docx"
        Case TEMPLATE_PCSUB
            fileName = "PC_SUB.docx"
        Case TEMPLATE_CDCA
            fileName = "CD_CA.docx"
        Case TEMPLATE_CDCASUB
            fileName = "CD_CA_SUB.docx"
        Case Else
            fileName = plantilla & ".docx"
    End Select

    ResolveTemplatePath = CurrentProject.Path & "\resources\plantillas\" & fileName
End Function

Private Sub AddFailure(ByRef failureCount As Long, ByRef firstFailure As String, ByRef logs() As String, _
    ByVal plantilla As String, ByVal campoTabla As String, ByVal campoWord As String, _
    ByVal extensionIndex As Long, ByVal rutaPlantilla As String, ByVal reason As String)

    Dim detail As String

    failureCount = failureCount + 1
    detail = "Drift: plantilla=" & plantilla & _
        "; campoTabla=" & campoTabla & _
        "; nombreCampoWord=" & campoWord & _
        "; extension=" & CStr(extensionIndex) & _
        "; rutaPlantilla=" & rutaPlantilla & _
        "; reason=" & reason

    If Len(firstFailure) = 0 Then firstFailure = detail
    If failureCount <= MAX_FAILURE_LOGS Then AppendLog logs, detail
End Sub

Private Function FileExists(ByVal path As String) As Boolean
    FileExists = (Len(Dir$(path, vbNormal)) > 0)
End Function

Private Function NewLogArray() As String()
    Dim result() As String
    ReDim result(0 To 0)
    result(0) = vbNullString
    NewLogArray = result
End Function

Private Sub AppendLog(ByRef logs() As String, ByVal message As String)
    Dim upperBound As Long

    On Error GoTo InitArray
    upperBound = UBound(logs)
    If upperBound = 0 And Len(logs(0)) = 0 Then
        logs(0) = message
    Else
        ReDim Preserve logs(0 To upperBound + 1)
        logs(upperBound + 1) = message
    End If
    Exit Sub

InitArray:
    ReDim logs(0 To 0)
    logs(0) = message
End Sub

Private Sub CleanupWordResources(ByRef wordDoc As Object, ByRef wordApp As Object, ByRef cleanupError As String)
    Dim doc As Object

    cleanupError = vbNullString
    On Error Resume Next

    If Not wordDoc Is Nothing Then
        wordDoc.Close SaveChanges:=False
        If Err.Number <> 0 And Len(cleanupError) = 0 Then cleanupError = Err.Description
        Err.Clear
    End If

    If Not wordApp Is Nothing Then
        For Each doc In wordApp.Documents
            doc.Close SaveChanges:=False
            If Err.Number <> 0 And Len(cleanupError) = 0 Then cleanupError = Err.Description
            Err.Clear
        Next doc

        wordApp.Quit SaveChanges:=False
        If Err.Number <> 0 And Len(cleanupError) = 0 Then cleanupError = Err.Description
        Err.Clear
    End If

    Set doc = Nothing
    Set wordDoc = Nothing
    Set wordApp = Nothing
    On Error GoTo 0
End Sub
