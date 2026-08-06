Attribute VB_Name = "Test_ManifestContracts"
Option Compare Database
Option Explicit

Private Const MAIN_MANIFEST_PATH As String = "tests\tests.vba.json"
Private Const SMOKE_MANIFEST_PATH As String = "tests\tests.vba.smoke.json"
Private Const SLICES_MANIFEST_PATH As String = "tests\tests.vba.slices.json"

Public Function Test_Manifest_MainAtomic_NoRunAll() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: leer " & MAIN_MANIFEST_PATH
    logs(1) = "2. Act: extraer procedures"
    logs(2) = "3. Assert: ningún procedure termina en _RunAll"

    Dim errMsg As String
    Dim manifestText As String
    manifestText = Test_Helper.ReadProjectFile(MAIN_MANIFEST_PATH, errMsg)
    If errMsg <> "" Then
        Test_Manifest_MainAtomic_NoRunAll = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    Dim procedures As Collection
    Dim offenders As New Collection
    Dim i As Long
    Set procedures = Test_Helper.ExtractProcedures(manifestText)
    For i = 1 To procedures.Count
        If Test_Helper.IsRunAllProcedure(CStr(procedures(i))) Then offenders.Add CStr(procedures(i))
    Next i

    logs(3) = "Procedures revisados: " & procedures.Count
    If offenders.Count > 0 Then
        Test_Manifest_MainAtomic_NoRunAll = Test_Helper.BuildJsonFail(MAIN_MANIFEST_PATH & " contiene agregadores: " & Test_Helper.JoinCollection(offenders), logs)
        Exit Function
    End If

    Test_Manifest_MainAtomic_NoRunAll = Test_Helper.BuildJsonOk("main_manifest_atomic", logs)
End Function

Public Function Test_Manifest_SmokeManifest_OnlyRunAll() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: leer " & SMOKE_MANIFEST_PATH
    logs(1) = "2. Act: extraer procedures"
    logs(2) = "3. Assert: todos los procedures terminan en _RunAll"

    Dim errMsg As String
    Dim manifestText As String
    manifestText = Test_Helper.ReadProjectFile(SMOKE_MANIFEST_PATH, errMsg)
    If errMsg <> "" Then
        Test_Manifest_SmokeManifest_OnlyRunAll = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    Dim procedures As Collection
    Dim offenders As New Collection
    Dim i As Long
    Set procedures = Test_Helper.ExtractProcedures(manifestText)
    For i = 1 To procedures.Count
        If Not Test_Helper.IsRunAllProcedure(CStr(procedures(i))) Then offenders.Add CStr(procedures(i))
    Next i

    logs(3) = "Procedures smoke revisados: " & procedures.Count
    If offenders.Count > 0 Then
        Test_Manifest_SmokeManifest_OnlyRunAll = Test_Helper.BuildJsonFail(SMOKE_MANIFEST_PATH & " contiene no-agregadores: " & Test_Helper.JoinCollection(offenders), logs)
        Exit Function
    End If

    Test_Manifest_SmokeManifest_OnlyRunAll = Test_Helper.BuildJsonOk("smoke_manifest_only_runall", logs)
End Function

Public Function Test_Manifest_NoDuplicateProcedures() As String
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: leer manifests principal y smoke"
    logs(1) = "2. Act: extraer procedures"
    logs(2) = "3. Assert: no hay procedures duplicados entre manifests"

    Dim errMsg As String
    Dim mainText As String
    Dim smokeText As String
    mainText = Test_Helper.ReadProjectFile(MAIN_MANIFEST_PATH, errMsg)
    If errMsg <> "" Then
        Test_Manifest_NoDuplicateProcedures = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    smokeText = Test_Helper.ReadProjectFile(SMOKE_MANIFEST_PATH, errMsg)
    If errMsg <> "" Then
        Test_Manifest_NoDuplicateProcedures = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    Dim mainProcedures As Collection
    Dim smokeProcedures As Collection
    Dim duplicates As New Collection
    Dim i As Long
    Set mainProcedures = Test_Helper.ExtractProcedures(mainText)
    Set smokeProcedures = Test_Helper.ExtractProcedures(smokeText)

    For i = 1 To smokeProcedures.Count
        If Test_Helper.CollectionContains(mainProcedures, CStr(smokeProcedures(i))) Then duplicates.Add CStr(smokeProcedures(i))
    Next i

    logs(3) = "Principal: " & mainProcedures.Count & " procedures"
    logs(4) = "Smoke: " & smokeProcedures.Count & " procedures"
    If duplicates.Count > 0 Then
        Test_Manifest_NoDuplicateProcedures = Test_Helper.BuildJsonFail("Procedures duplicados entre manifests: " & Test_Helper.JoinCollection(duplicates), logs)
        Exit Function
    End If

    Test_Manifest_NoDuplicateProcedures = Test_Helper.BuildJsonOk("manifest_no_duplicate_procedures", logs)
End Function

Public Function Test_Manifest_Slices_NoDuplicateNames() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: leer " & SLICES_MANIFEST_PATH
    logs(1) = "2. Act: extraer nombres de slices"
    logs(2) = "3. Assert: no hay nombres duplicados"

    Dim errMsg As String
    Dim slicesText As String
    slicesText = Test_Helper.ReadProjectFile(SLICES_MANIFEST_PATH, errMsg)
    If errMsg <> "" Then
        Test_Manifest_Slices_NoDuplicateNames = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    Dim names As Collection
    Dim seen As Object
    Dim duplicates As New Collection
    Dim i As Long
    Set names = Test_Helper.ExtractJsonPropertyValues(slicesText, "name")
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    For i = 1 To names.Count
        If seen.Exists(CStr(names(i))) Then
            duplicates.Add CStr(names(i))
        Else
            seen.Add CStr(names(i)), True
        End If
    Next i

    logs(3) = "Slices declarados: " & names.Count
    If duplicates.Count > 0 Then
        Test_Manifest_Slices_NoDuplicateNames = Test_Helper.BuildJsonFail("Slices duplicados: " & Test_Helper.JoinCollection(duplicates), logs)
        Exit Function
    End If

    Test_Manifest_Slices_NoDuplicateNames = Test_Helper.BuildJsonOk("slices_unique_names", logs)
End Function

Public Function Test_Manifest_Slices_SourceManifest_IsMain() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: leer " & SLICES_MANIFEST_PATH
    logs(1) = "2. Act: extraer sourceManifest"
    logs(2) = "3. Assert: sourceManifest apunta al manifest principal"

    Dim errMsg As String
    Dim slicesText As String
    slicesText = Test_Helper.ReadProjectFile(SLICES_MANIFEST_PATH, errMsg)
    If errMsg <> "" Then
        Test_Manifest_Slices_SourceManifest_IsMain = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    Dim values As Collection
    Set values = Test_Helper.ExtractJsonPropertyValues(slicesText, "sourceManifest")
    logs(3) = "sourceManifest encontrados: " & values.Count

    If values.Count = 0 Then
        Test_Manifest_Slices_SourceManifest_IsMain = Test_Helper.BuildJsonFail("No existe sourceManifest en " & SLICES_MANIFEST_PATH, logs)
        Exit Function
    End If
    Dim actualSourceManifest As String
    actualSourceManifest = CStr(values(1))
    actualSourceManifest = Replace(actualSourceManifest, "\\", "\")
    actualSourceManifest = Replace(actualSourceManifest, "/", "\")

    If actualSourceManifest <> MAIN_MANIFEST_PATH Then
        Test_Manifest_Slices_SourceManifest_IsMain = Test_Helper.BuildJsonFail("sourceManifest esperado=" & MAIN_MANIFEST_PATH & "; actual=" & CStr(values(1)), logs)
        Exit Function
    End If

    Test_Manifest_Slices_SourceManifest_IsMain = Test_Helper.BuildJsonOk("slices_source_manifest_main", logs)
End Function
