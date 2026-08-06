Attribute VB_Name = "Test_Punto01Terminologia"
Option Compare Database
Option Explicit

' --- Test_Punto01_NoConsultarEnFormsObjetivo ---
' Test atómico del Punto 01 del acta 2026-06-25: "consultar" → "mostrar" en la UI.
'
' Verifica estáticamente que ninguna forma (.cls) o archivo de layout (.form.txt)
' de src/forms contenga la cadena "consultar" como palabra completa. La verificación
' es case-insensitive y respeta límites de palabra (no matchea "consultarse",
' "consultarAlgo", etc.) para evitar falsos positivos en identificadores
' compuestos.
'
' Contrato: docs/uat/contrato-staging-2026-06-22.html §#req-cal-01.
' Acta: docs/uat/acta-reunion-calidad-2026-06-25.html Punto 01.
'
' Resultado:
'   ok=true  → 0 apariciones de "consultar" en forms objetivo (cumple)
'   ok=false → N>0 apariciones (lista archivos:línea:contenido en error)
'
' Tags en manifest: fase-a, punto-01, terminologia, ui-naming,
'                   consultar-mostrar, static-check, documentation
Public Function Test_Punto01_NoConsultarEnFormsObjetivo() As String
    On Error GoTo ErrHandler

    ' Per skill access-vba-tdd-sandbox + decisión del usuario 2026-07-22:
    ' path hard-pineado para evitar dependencia de Application.CurrentProject.Path
    ' en el runtime del test runner (mismo patrón que ForceLocalBackend bypass).
    Const FORMS_PATH As String = "C:\00repos\codigo\00_GESTION_RIESGOS_staging\src\forms"

    Dim logs(0 To 5) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 1. Verificar que el directorio existe
    If Not fso.FolderExists(FORMS_PATH) Then
        logs(0) = "1. Arrange: Forms path hard-pineado = " & FORMS_PATH
        logs(1) = "2. Act: FolderExists(FORMS_PATH)"
        logs(2) = "3. Assert: NO existe — entorno inválido"
        Test_Punto01_NoConsultarEnFormsObjetivo = BuildJsonFail( _
            "Forms path no existe: " & FORMS_PATH, logs)
        Exit Function
    End If

    ' 2. Iterar todos los Form_*.cls y Form_*.form.txt buscando "consultar"
    Dim hits As Collection
    Set hits = New Collection

    Dim filesScanned As Long
    filesScanned = 0

    Dim folder As Object
    Set folder = fso.GetFolder(FORMS_PATH)

    Dim file As Object
    For Each file In folder.Files
        Dim ext As String
        ext = LCase$(fso.GetExtensionName(file.Path))
        If ext = "cls" Or ext = "txt" Then
            ' Solo archivos que empiecen con "Form_" (excluye cls de classes/, txt sueltos)
            If LCase$(Left$(fso.GetFileName(file.Path), 5)) = "form_" Then
                filesScanned = filesScanned + 1
                ScanFileForConsultar file.Path, hits
            End If
        End If
    Next file

    ' 3. Reportar resultado
    logs(0) = "1. Arrange: Forms path hard-pineado = " & FORMS_PATH
    logs(1) = "2. Act: Scan " & filesScanned & " archivos (.cls + .form.txt) buscando 'consultar' (case-insensitive, word-boundary)"
    logs(2) = "3. Assert: hits = " & hits.Count

    If hits.Count > 0 Then
        Dim msg As String
        msg = "Encontradas " & hits.Count & " apariciones de 'consultar' en forms objetivo (esperado: 0). Lista: " _
            & vbCrLf & JoinCollection(hits, vbCrLf)
        logs(3) = "FAIL: " & msg
        Test_Punto01_NoConsultarEnFormsObjetivo = BuildJsonFail(msg, logs)
        Exit Function
    End If

    logs(3) = "4. PASS"
    Test_Punto01_NoConsultarEnFormsObjetivo = BuildJsonOk( _
        "punto01_consultar_mostrar_pass", logs)
    Exit Function

ErrHandler:
    logs(0) = "ERR: " & Err.Number & " - " & Err.Description
    Test_Punto01_NoConsultarEnFormsObjetivo = BuildJsonFail( _
        "Test_Punto01_NoConsultarEnFormsObjetivo: " & Err.Number & " - " & Err.Description, logs)
End Function

' --- ScanFileForConsultar (helper privado) ---
' Lee el archivo línea por línea y, por cada hit de "consultar" como palabra
' completa (no substring de identificador más largo), añade un string
' "<archivo>:<número-línea>: <contenido>" a la colección hits.
Private Sub ScanFileForConsultar(ByVal p_Path As String, ByRef p_Hits As Collection)
    Const BUF_SIZE As Long = 4096
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim ts As Object
    Set ts = fso.OpenTextFile(p_Path, 1, False, -2)  ' -2 = System default (lee bytes)

    Dim line As String
    Dim lineNum As Long
    lineNum = 0
    Do While Not ts.AtEndOfStream
        lineNum = lineNum + 1
        line = ts.ReadLine
        If MatchWordConsultar(line) Then
            p_Hits.Add fso.GetFileName(p_Path) & ":" & lineNum & ": " & Trim$(line)
        End If
    Loop
    ts.Close
End Sub

' --- MatchWordConsultar (helper privado) ---
' Devuelve True si la cadena contiene la palabra "consultar" con bordes de
' palabra. Case-insensitive. Implementación sin RegExp (no garantizado en
' todas las versiones de VBA) — usa comparaciones de caracteres adyacentes.
'
' Bordes de palabra: inicio/fin de la cadena, o cualquier carácter que no sea
' letra (a-zA-Z) ni dígito (0-9) ni underscore (_). Esto excluye
' identificadores como "consultarX", "Xconsultar", "m_ConsultarAlgo",
' "sub_consultar" pero matchea "consultar", "consultar ", " (consultar)",
' "consultar.", "consultar,", etc.
Private Function MatchWordConsultar(ByVal p_Line As String) As Boolean
    MatchWordConsultar = False

    Dim needle As String
    needle = "consultar"
    Dim needleLen As Long
    needleLen = Len(needle)

    Dim lowerLine As String
    lowerLine = LCase$(p_Line)
    Dim lineLen As Long
    lineLen = Len(lowerLine)

    If lineLen < needleLen Then Exit Function

    Dim i As Long
    For i = 1 To lineLen - needleLen + 1
        If Mid$(lowerLine, i, needleLen) = needle Then
            ' Comprobar borde izquierdo
            Dim leftOk As Boolean
            If i = 1 Then
                leftOk = True
            Else
                leftOk = Not IsWordChar(Mid$(lowerLine, i - 1, 1))
            End If

            ' Comprobar borde derecho
            Dim rightOk As Boolean
            If i + needleLen - 1 = lineLen Then
                rightOk = True
            Else
                rightOk = Not IsWordChar(Mid$(lowerLine, i + needleLen, 1))
            End If

            If leftOk And rightOk Then
                MatchWordConsultar = True
                Exit Function
            End If
        End If
    Next i
End Function

' --- IsWordChar (helper privado) ---
' True si el carácter es letra, dígito o underscore (parte de un identificador).
Private Function IsWordChar(ByVal p_Char As String) As Boolean
    IsWordChar = False
    If Len(p_Char) = 0 Then Exit Function
    Dim c As String
    c = p_Char
    If c >= "a" And c <= "z" Then IsWordChar = True: Exit Function
    If c >= "A" And c <= "Z" Then IsWordChar = True: Exit Function
    If c >= "0" And c <= "9" Then IsWordChar = True: Exit Function
    If c = "_" Then IsWordChar = True: Exit Function
End Function

' --- JoinCollection (helper privado) ---
' Concatena los elementos de una Collection usando un separador.
Private Function JoinCollection(ByVal p_Coll As Collection, ByVal p_Sep As String) As String
    Dim s As String
    Dim i As Long
    For i = 1 To p_Coll.Count
        If i > 1 Then s = s & p_Sep
        s = s & p_Coll(i)
    Next i
    JoinCollection = s
End Function
