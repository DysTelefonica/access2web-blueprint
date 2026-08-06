Attribute VB_Name = "Test_ManifestContracts"
Option Compare Database
Option Explicit

' ============================================================
' Test_ManifestContracts — Contratos ejecutables de manifests VBA
'
' Skill: access-vba-tdd v2.2
' Objetivo: asegurar que la suite principal sea atómica y que
' los agregadores RunAll vivan separados en smoke.
' ============================================================

Private Const MAIN_MANIFEST_PATH As String = "tests\tests.vba.json"
Private Const SMOKE_MANIFEST_PATH As String = "tests\tests.vba.smoke.json"
Private Const SLICES_MANIFEST_PATH As String = "tests\tests.vba.slices.json"

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Private Function ReadProjectFile(ByVal fileName As String, Optional ByRef p_Error As String = "") As String
    On Error GoTo EH
    p_Error = ""

    Dim path As String
    path = CurrentProject.Path & "\" & fileName

    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then
        p_Error = "Manifest no encontrado: " & path
        ReadProjectFile = ""
        Exit Function
    End If

    Set ts = fso.OpenTextFile(path, 1, False)
    ReadProjectFile = ts.ReadAll
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Exit Function

EH:
    p_Error = "ReadProjectFile: " & Err.Number & " - " & Err.Description
    ReadProjectFile = ""
End Function

Private Function ExtractProcedures(ByVal manifestText As String) As Collection
    Dim rx As Object
    Dim matches As Object
    Dim m As Object
    Dim procedures As New Collection

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = """procedure""\s*:\s*""([^""]+)"""

    Set matches = rx.Execute(manifestText)
    For Each m In matches
        procedures.Add CStr(m.SubMatches(0))
    Next m

    Set ExtractProcedures = procedures
End Function

Private Function IsRunAllProcedure(ByVal procedureName As String) As Boolean
    IsRunAllProcedure = (Right$(procedureName, 7) = "_RunAll")
End Function

Private Function JoinCollection(ByVal values As Collection) As String
    Dim i As Long
    Dim result As String
    For i = 1 To values.Count
        If result <> "" Then result = result & ", "
        result = result & CStr(values(i))
    Next i
    JoinCollection = result
End Function

Private Function CollectionContains(ByVal values As Collection, ByVal target As String) As Boolean
    Dim i As Long
    For i = 1 To values.Count
        If CStr(values(i)) = target Then
            CollectionContains = True
            Exit Function
        End If
    Next i
    CollectionContains = False
End Function

Private Function ExtractSlicePropertyValues(ByVal manifestText As String, ByVal propertyName As String) As Collection
    Dim rx As Object
    Dim matches As Object
    Dim m As Object
    Dim values As New Collection

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = Chr$(34) & propertyName & Chr$(34) & "\s*:\s*" & Chr$(34) & "([^" & Chr$(34) & "]+)" & Chr$(34)

    Set matches = rx.Execute(manifestText)
    For Each m In matches
        values.Add CStr(m.SubMatches(0))
    Next m

    Set ExtractSlicePropertyValues = values
End Function

Private Function ValidateMandatorySliceFilters(ByVal filters As Collection) As String
    Dim required As Variant
    required = Array("control-cambios", "informe-html", "batch", "cache", "subcontratistas", "publicabilidad", "suministradores")

    Dim i As Long
    Dim missing As New Collection
    For i = LBound(required) To UBound(required)
        If Not CollectionContains(filters, CStr(required(i))) Then
            missing.Add CStr(required(i))
        End If
    Next i

    If missing.Count > 0 Then
        ValidateMandatorySliceFilters = "Faltan filtros obligatorios en " & SLICES_MANIFEST_PATH & ": " & JoinCollection(missing)
    Else
        ValidateMandatorySliceFilters = ""
    End If
End Function

Public Function Test_Manifest_MainAtomic_NoRunAll() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: leer " & MAIN_MANIFEST_PATH
    logs(1) = "2. Act: extraer procedures"
    logs(2) = "3. Assert: ningún procedure termina en _RunAll"

    Dim errMsg As String
    Dim manifestText As String
    manifestText = ReadProjectFile(MAIN_MANIFEST_PATH, errMsg)
    If errMsg <> "" Then
        Test_Manifest_MainAtomic_NoRunAll = BuildFail(errMsg, logs)
        Exit Function
    End If

    Dim procedures As Collection
    Dim offenders As New Collection
    Dim i As Long
    Set procedures = ExtractProcedures(manifestText)
    For i = 1 To procedures.Count
        If IsRunAllProcedure(CStr(procedures(i))) Then offenders.Add CStr(procedures(i))
    Next i

    logs(3) = "Procedures revisados: " & procedures.Count
    If offenders.Count > 0 Then
        Test_Manifest_MainAtomic_NoRunAll = BuildFail(MAIN_MANIFEST_PATH & " contiene agregadores: " & JoinCollection(offenders), logs)
        Exit Function
    End If

    Test_Manifest_MainAtomic_NoRunAll = BuildOk("main_manifest_atomic", logs)
End Function

Public Function Test_Manifest_SmokeOnly_RunAll() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: leer " & SMOKE_MANIFEST_PATH
    logs(1) = "2. Act: extraer procedures"
    logs(2) = "3. Assert: todos los procedures terminan en _RunAll"

    Dim errMsg As String
    Dim manifestText As String
    manifestText = ReadProjectFile(SMOKE_MANIFEST_PATH, errMsg)
    If errMsg <> "" Then
        Test_Manifest_SmokeOnly_RunAll = BuildFail(errMsg, logs)
        Exit Function
    End If

    Dim procedures As Collection
    Dim offenders As New Collection
    Dim i As Long
    Set procedures = ExtractProcedures(manifestText)
    For i = 1 To procedures.Count
        If Not IsRunAllProcedure(CStr(procedures(i))) Then offenders.Add CStr(procedures(i))
    Next i

    logs(3) = "Procedures smoke revisados: " & procedures.Count
    If offenders.Count > 0 Then
        Test_Manifest_SmokeOnly_RunAll = BuildFail(SMOKE_MANIFEST_PATH & " contiene no-agregadores: " & JoinCollection(offenders), logs)
        Exit Function
    End If

    Test_Manifest_SmokeOnly_RunAll = BuildOk("smoke_manifest_only_runall", logs)
End Function

Public Function Test_Manifest_NoDuplicateProcedures() As String
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: leer manifests principal y smoke"
    logs(1) = "2. Act: extraer procedures"
    logs(2) = "3. Assert: no hay procedures duplicados entre manifests"

    Dim errMsg As String
    Dim mainText As String
    Dim smokeText As String
    mainText = ReadProjectFile(MAIN_MANIFEST_PATH, errMsg)
    If errMsg <> "" Then
        Test_Manifest_NoDuplicateProcedures = BuildFail(errMsg, logs)
        Exit Function
    End If

    smokeText = ReadProjectFile(SMOKE_MANIFEST_PATH, errMsg)
    If errMsg <> "" Then
        Test_Manifest_NoDuplicateProcedures = BuildFail(errMsg, logs)
        Exit Function
    End If

    Dim mainProcedures As Collection
    Dim smokeProcedures As Collection
    Dim duplicates As New Collection
    Dim i As Long
    Set mainProcedures = ExtractProcedures(mainText)
    Set smokeProcedures = ExtractProcedures(smokeText)

    For i = 1 To smokeProcedures.Count
        If CollectionContains(mainProcedures, CStr(smokeProcedures(i))) Then duplicates.Add CStr(smokeProcedures(i))
    Next i

    logs(3) = "Principal: " & mainProcedures.Count & " procedures"
    logs(4) = "Smoke: " & smokeProcedures.Count & " procedures"
    If duplicates.Count > 0 Then
        Test_Manifest_NoDuplicateProcedures = BuildFail("Procedures duplicados entre manifests: " & JoinCollection(duplicates), logs)
        Exit Function
    End If

    Test_Manifest_NoDuplicateProcedures = BuildOk("manifest_no_duplicate_procedures", logs)
End Function

Public Function Test_Manifest_Slices_MandatoryFilters() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: leer " & SLICES_MANIFEST_PATH
    logs(1) = "2. Act: extraer propiedades ""filter"""
    logs(2) = "3. Assert: existen filtros obligatorios por slice"

    Dim errMsg As String
    Dim slicesText As String
    slicesText = ReadProjectFile(SLICES_MANIFEST_PATH, errMsg)
    If errMsg <> "" Then
        Test_Manifest_Slices_MandatoryFilters = BuildFail(errMsg, logs)
        Exit Function
    End If

    Dim filters As Collection
    Set filters = ExtractSlicePropertyValues(slicesText, "filter")
    logs(3) = "Slices declarados: " & filters.Count

    If filters.Count = 0 Then
        Test_Manifest_Slices_MandatoryFilters = BuildFail(SLICES_MANIFEST_PATH & " no declara filtros", logs)
        Exit Function
    End If

    errMsg = ValidateMandatorySliceFilters(filters)
    If errMsg <> "" Then
        Test_Manifest_Slices_MandatoryFilters = BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_Manifest_Slices_MandatoryFilters = BuildOk("slices_mandatory_filters", logs)
End Function

Public Function Test_Manifest_Slices_NoDuplicateNames() As String
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: leer " & SLICES_MANIFEST_PATH
    logs(1) = "2. Act: extraer propiedades ""name"""
    logs(2) = "3. Assert: no hay nombres de slice duplicados"

    Dim errMsg As String
    Dim slicesText As String
    slicesText = ReadProjectFile(SLICES_MANIFEST_PATH, errMsg)
    If errMsg <> "" Then
        Test_Manifest_Slices_NoDuplicateNames = BuildFail(errMsg, logs)
        Exit Function
    End If

    Dim names As Collection
    Dim seen As New Collection
    Dim duplicates As New Collection
    Dim i As Long
    Set names = ExtractSlicePropertyValues(slicesText, "name")

    For i = 1 To names.Count
        If CollectionContains(seen, CStr(names(i))) Then
            duplicates.Add CStr(names(i))
        Else
            seen.Add CStr(names(i))
        End If
    Next i

    logs(3) = "Slices declarados: " & names.Count
    logs(4) = "Duplicados detectados: " & duplicates.Count

    If duplicates.Count > 0 Then
        Test_Manifest_Slices_NoDuplicateNames = BuildFail("Nombres de slice duplicados: " & JoinCollection(duplicates), logs)
        Exit Function
    End If

    Test_Manifest_Slices_NoDuplicateNames = BuildOk("slices_unique_names", logs)
End Function
