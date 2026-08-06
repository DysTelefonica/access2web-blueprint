Attribute VB_Name = "Test_IndicadorRiesgosRepositorio"
' ============================================================
' Test_IndicadorRiesgosRepositorio
'
' issue-61: defensa contra inyeccion SQL por p_ListaIDsCsv.
' Cubre el helper ValidarListaIDsCsv (pure logic) y la
' propagacion de Err 1000 desde IndicadorRiesgosV2_SqlDetalle
' cuando el caller le pasa un payload malicioso.
'
' Skill: access-vba-tdd v2.4 — pure logic, sin fixtures, sin DB.
' ============================================================
Option Compare Database
Option Explicit

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function


' ============================================================
' ValidarListaIDsCsv — CASOS VALIDOS
' ============================================================

Public Function Test_IndicadorRiesgos_ValidarCsv_Valido_SoloUnId() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input '42'"
    logs(1) = "2. Act: ValidarListaIDsCsv('42')"
    logs(2) = "3. Assert: devuelve '42' normalizado"

    Dim result As String
    result = ValidarListaIDsCsv("42")

    If result <> "42" Then
        Test_IndicadorRiesgos_ValidarCsv_Valido_SoloUnId = BuildFail("Esperado '42', obtenido '" & result & "'", logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Valido_SoloUnId = BuildOk(result, logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Valido_SoloUnId = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Valido_Multiples() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input '1, 2, 3'"
    logs(1) = "2. Act: ValidarListaIDsCsv('1, 2, 3')"
    logs(2) = "3. Assert: devuelve '1, 2, 3' con espacios"

    Dim result As String
    result = ValidarListaIDsCsv("1, 2, 3")

    If result <> "1, 2, 3" Then
        Test_IndicadorRiesgos_ValidarCsv_Valido_Multiples = BuildFail("Esperado '1, 2, 3', obtenido '" & result & "'", logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Valido_Multiples = BuildOk(result, logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Valido_Multiples = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Valido_TrimeaEspacios() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input '  10 , 20  '"
    logs(1) = "2. Act: ValidarListaIDsCsv con espacios"
    logs(2) = "3. Assert: devuelve '10, 20' normalizado"

    Dim result As String
    result = ValidarListaIDsCsv("  10 , 20  ")

    If result <> "10, 20" Then
        Test_IndicadorRiesgos_ValidarCsv_Valido_TrimeaEspacios = BuildFail("Esperado '10, 20', obtenido '" & result & "'", logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Valido_TrimeaEspacios = BuildOk(result, logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Valido_TrimeaEspacios = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Valido_DropeaVacios() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input '1,, 2, , 3'"
    logs(1) = "2. Act: ValidarListaIDsCsv con separadores extra"
    logs(2) = "3. Assert: descarta vacios y devuelve '1, 2, 3'"

    Dim result As String
    result = ValidarListaIDsCsv("1,, 2, , 3")

    If result <> "1, 2, 3" Then
        Test_IndicadorRiesgos_ValidarCsv_Valido_DropeaVacios = BuildFail("Esperado '1, 2, 3', obtenido '" & result & "'", logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Valido_DropeaVacios = BuildOk(result, logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Valido_DropeaVacios = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function


' ============================================================
' ValidarListaIDsCsv — CASOS VACIOS (devuelve "" sin error)
' ============================================================

Public Function Test_IndicadorRiesgos_ValidarCsv_Vacio_String() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input ''"
    logs(1) = "2. Act: ValidarListaIDsCsv('')"
    logs(2) = "3. Assert: devuelve '' sin error"

    Dim result As String
    result = ValidarListaIDsCsv("")

    If result <> "" Then
        Test_IndicadorRiesgos_ValidarCsv_Vacio_String = BuildFail("Esperado vacio, obtenido '" & result & "'", logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Vacio_String = BuildOk(result, logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Vacio_String = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Vacio_SoloEspacios() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input '   '"
    logs(1) = "2. Act: ValidarListaIDsCsv con solo espacios"
    logs(2) = "3. Assert: devuelve ''"

    Dim result As String
    result = ValidarListaIDsCsv("   ")

    If result <> "" Then
        Test_IndicadorRiesgos_ValidarCsv_Vacio_SoloEspacios = BuildFail("Esperado vacio, obtenido '" & result & "'", logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Vacio_SoloEspacios = BuildOk(result, logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Vacio_SoloEspacios = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Vacio_SoloSeparadores() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input ',,, , ,'"
    logs(1) = "2. Act: ValidarListaIDsCsv con solo comas"
    logs(2) = "3. Assert: devuelve '' (sin IDs validos)"

    Dim result As String
    result = ValidarListaIDsCsv(",,, , ,")

    If result <> "" Then
        Test_IndicadorRiesgos_ValidarCsv_Vacio_SoloSeparadores = BuildFail("Esperado vacio, obtenido '" & result & "'", logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Vacio_SoloSeparadores = BuildOk(result, logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Vacio_SoloSeparadores = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function


' ============================================================
' ValidarListaIDsCsv — CASOS INVALIDOS (deben elevar Err 1000)
' ============================================================

Private Function TryValidarCsv(ByVal p_Input As String, ByRef p_Result As String, ByRef p_ErrNumber As Long, ByRef p_ErrDesc As String) As Boolean
    On Error GoTo TryValidarCsv_EH
    p_Result = ValidarListaIDsCsv(p_Input)
    p_ErrNumber = 0
    p_ErrDesc = ""
    TryValidarCsv = True
    Exit Function
TryValidarCsv_EH:
    p_ErrNumber = Err.Number
    p_ErrDesc = Err.Description
    p_Result = ""
    TryValidarCsv = False
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Rechaza_InyeccionOr() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input '1) OR 1=1 --'"
    logs(1) = "2. Act: ValidarListaIDsCsv con payload SQL clasico"
    logs(2) = "3. Assert: eleva Err 1000"

    Dim result As String
    Dim errNum As Long, errDesc As String
    If TryValidarCsv("1) OR 1=1 --", result, errNum, errDesc) Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_InyeccionOr = BuildFail("Debio elevar Err 1000, devolvio '" & result & "'", logs)
        Exit Function
    End If

    If errNum <> 1000 Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_InyeccionOr = BuildFail("Esperado Err 1000, obtuvo Err " & errNum & ": " & errDesc, logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Rechaza_InyeccionOr = BuildOk("err_1000", logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Rechaza_InyeccionOr = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Rechaza_ComillaApostrofe() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input ""1'; DROP TABLE TbProyectos; --"""
    logs(1) = "2. Act: ValidarListaIDsCsv con intento de DROP TABLE"
    logs(2) = "3. Assert: eleva Err 1000"

    Dim result As String
    Dim errNum As Long, errDesc As String
    If TryValidarCsv("1'; DROP TABLE TbProyectos; --", result, errNum, errDesc) Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_ComillaApostrofe = BuildFail("Debio elevar Err 1000, devolvio '" & result & "'", logs)
        Exit Function
    End If

    If errNum <> 1000 Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_ComillaApostrofe = BuildFail("Esperado Err 1000, obtuvo Err " & errNum & ": " & errDesc, logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Rechaza_ComillaApostrofe = BuildOk("err_1000", logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Rechaza_ComillaApostrofe = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Rechaza_UnionSelect() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input '1 UNION SELECT password FROM TbUsuarios'"
    logs(1) = "2. Act: ValidarListaIDsCsv con intento de UNION"
    logs(2) = "3. Assert: eleva Err 1000"

    Dim result As String
    Dim errNum As Long, errDesc As String
    If TryValidarCsv("1 UNION SELECT password FROM TbUsuarios", result, errNum, errDesc) Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_UnionSelect = BuildFail("Debio elevar Err 1000, devolvio '" & result & "'", logs)
        Exit Function
    End If

    If errNum <> 1000 Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_UnionSelect = BuildFail("Esperado Err 1000, obtuvo Err " & errNum & ": " & errDesc, logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Rechaza_UnionSelect = BuildOk("err_1000", logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Rechaza_UnionSelect = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Rechaza_Alfabetico() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input 'abc'"
    logs(1) = "2. Act: ValidarListaIDsCsv con texto alfabetico"
    logs(2) = "3. Assert: eleva Err 1000"

    Dim result As String
    Dim errNum As Long, errDesc As String
    If TryValidarCsv("abc", result, errNum, errDesc) Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_Alfabetico = BuildFail("Debio elevar Err 1000, devolvio '" & result & "'", logs)
        Exit Function
    End If

    If errNum <> 1000 Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_Alfabetico = BuildFail("Esperado Err 1000, obtuvo Err " & errNum & ": " & errDesc, logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Rechaza_Alfabetico = BuildOk("err_1000", logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Rechaza_Alfabetico = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Rechaza_Decimal() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input '1.5'"
    logs(1) = "2. Act: ValidarListaIDsCsv con decimal"
    logs(2) = "3. Assert: eleva Err 1000"

    Dim result As String
    Dim errNum As Long, errDesc As String
    If TryValidarCsv("1.5", result, errNum, errDesc) Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_Decimal = BuildFail("Debio elevar Err 1000, devolvio '" & result & "'", logs)
        Exit Function
    End If

    If errNum <> 1000 Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_Decimal = BuildFail("Esperado Err 1000, obtuvo Err " & errNum & ": " & errDesc, logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Rechaza_Decimal = BuildOk("err_1000", logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Rechaza_Decimal = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Rechaza_Negativo() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input '-1'"
    logs(1) = "2. Act: ValidarListaIDsCsv con negativo"
    logs(2) = "3. Assert: eleva Err 1000"

    Dim result As String
    Dim errNum As Long, errDesc As String
    If TryValidarCsv("-1", result, errNum, errDesc) Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_Negativo = BuildFail("Debio elevar Err 1000, devolvio '" & result & "'", logs)
        Exit Function
    End If

    If errNum <> 1000 Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_Negativo = BuildFail("Esperado Err 1000, obtuvo Err " & errNum & ": " & errDesc, logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Rechaza_Negativo = BuildOk("err_1000", logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Rechaza_Negativo = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Rechaza_NotacionCientifica() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input '1e5'"
    logs(1) = "2. Act: ValidarListaIDsCsv con notacion cientifica"
    logs(2) = "3. Assert: eleva Err 1000 (no es entero puro)"

    Dim result As String
    Dim errNum As Long, errDesc As String
    If TryValidarCsv("1e5", result, errNum, errDesc) Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_NotacionCientifica = BuildFail("Debio elevar Err 1000, devolvio '" & result & "'", logs)
        Exit Function
    End If

    If errNum <> 1000 Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_NotacionCientifica = BuildFail("Esperado Err 1000, obtuvo Err " & errNum & ": " & errDesc, logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Rechaza_NotacionCientifica = BuildOk("err_1000", logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Rechaza_NotacionCientifica = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Rechaza_Hexadecimal() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input '0x1F'"
    logs(1) = "2. Act: ValidarListaIDsCsv con prefijo hexadecimal"
    logs(2) = "3. Assert: eleva Err 1000"

    Dim result As String
    Dim errNum As Long, errDesc As String
    If TryValidarCsv("0x1F", result, errNum, errDesc) Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_Hexadecimal = BuildFail("Debio elevar Err 1000, devolvio '" & result & "'", logs)
        Exit Function
    End If

    If errNum <> 1000 Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_Hexadecimal = BuildFail("Esperado Err 1000, obtuvo Err " & errNum & ": " & errDesc, logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Rechaza_Hexadecimal = BuildOk("err_1000", logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Rechaza_Hexadecimal = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Rechaza_ComentarioSQL() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input '1 /* comentario'"
    logs(1) = "2. Act: ValidarListaIDsCsv con intento de comentario SQL"
    logs(2) = "3. Assert: eleva Err 1000"

    Dim result As String
    Dim errNum As Long, errDesc As String
    If TryValidarCsv("1 /* comentario", result, errNum, errDesc) Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_ComentarioSQL = BuildFail("Debio elevar Err 1000, devolvio '" & result & "'", logs)
        Exit Function
    End If

    If errNum <> 1000 Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_ComentarioSQL = BuildFail("Esperado Err 1000, obtuvo Err " & errNum & ": " & errDesc, logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Rechaza_ComentarioSQL = BuildOk("err_1000", logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Rechaza_ComentarioSQL = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_ValidarCsv_Rechaza_MezclaValidoInvalido() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: input '1, 2, abc, 4' (mezcla)"
    logs(1) = "2. Act: ValidarListaIDsCsv con ID invalido entre validos"
    logs(2) = "3. Assert: eleva Err 1000 (no acepta parcialmente)"

    Dim result As String
    Dim errNum As Long, errDesc As String
    If TryValidarCsv("1, 2, abc, 4", result, errNum, errDesc) Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_MezclaValidoInvalido = BuildFail("Debio elevar Err 1000, devolvio '" & result & "'", logs)
        Exit Function
    End If

    If errNum <> 1000 Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_MezclaValidoInvalido = BuildFail("Esperado Err 1000, obtuvo Err " & errNum & ": " & errDesc, logs)
        Exit Function
    End If

    ' Ademas: el mensaje debe mencionar el valor problematico
    If InStr(1, errDesc, "abc", vbTextCompare) = 0 Then
        Test_IndicadorRiesgos_ValidarCsv_Rechaza_MezclaValidoInvalido = BuildFail("Mensaje debe mencionar el valor invalido 'abc'. Obtenido: " & errDesc, logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_ValidarCsv_Rechaza_MezclaValidoInvalido = BuildOk("err_1000", logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_ValidarCsv_Rechaza_MezclaValidoInvalido = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function


' ============================================================
' IndicadorRiesgosV2_SqlDetalle — INTEGRACION
' El helper se invoca en cada Case del Select; cualquier payload
' malicioso debe elevar Err 1000 y NO llegar al SQL builder.
' ============================================================

Public Function Test_IndicadorRiesgos_SqlDetalle_RechazaInyeccion_ItIdentificados() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: itIdentificados con CSV malicioso"
    logs(1) = "2. Act: IndicadorRiesgosV2_SqlDetalle con payload SQL"
    logs(2) = "3. Assert: eleva Err 1000, no retorna SQL con payload"

    Dim sql As String
    On Error Resume Next
    sql = IndicadorRiesgosV2_SqlDetalle(itIdentificados, DateSerial(2026, 1, 1), DateSerial(2026, 12, 31), "1) OR 1=1 --")
    Dim errNum As Long
    errNum = Err.Number
    On Error GoTo EH

    If errNum <> 1000 Then
        Test_IndicadorRiesgos_SqlDetalle_RechazaInyeccion_ItIdentificados = BuildFail("Esperado Err 1000, obtuvo Err " & errNum & ". SQL='" & sql & "'", logs)
        Exit Function
    End If

    If InStr(1, sql, "OR 1=1", vbTextCompare) > 0 Then
        Test_IndicadorRiesgos_SqlDetalle_RechazaInyeccion_ItIdentificados = BuildFail("El SQL no debe contener el payload. SQL='" & sql & "'", logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_SqlDetalle_RechazaInyeccion_ItIdentificados = BuildOk("err_1000_propagado", logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_SqlDetalle_RechazaInyeccion_ItIdentificados = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_SqlDetalle_RechazaInyeccion_ItMaterializados() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: itMaterializados con CSV malicioso"
    logs(1) = "2. Act: IndicadorRiesgosV2_SqlDetalle con payload SQL"
    logs(2) = "3. Assert: eleva Err 1000, no retorna SQL con payload"

    Dim sql As String
    On Error Resume Next
    sql = IndicadorRiesgosV2_SqlDetalle(itMaterializados, DateSerial(2026, 1, 1), DateSerial(2026, 12, 31), "5 UNION SELECT * FROM TbUsuarios")
    Dim errNum As Long
    errNum = Err.Number
    On Error GoTo EH

    If errNum <> 1000 Then
        Test_IndicadorRiesgos_SqlDetalle_RechazaInyeccion_ItMaterializados = BuildFail("Esperado Err 1000, obtuvo Err " & errNum & ". SQL='" & sql & "'", logs)
        Exit Function
    End If

    If InStr(1, sql, "UNION SELECT", vbTextCompare) > 0 Then
        Test_IndicadorRiesgos_SqlDetalle_RechazaInyeccion_ItMaterializados = BuildFail("El SQL no debe contener el payload. SQL='" & sql & "'", logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_SqlDetalle_RechazaInyeccion_ItMaterializados = BuildOk("err_1000_propagado", logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_SqlDetalle_RechazaInyeccion_ItMaterializados = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function

Public Function Test_IndicadorRiesgos_SqlDetalle_CsvValidoGeneraSqlConIds() As String
    On Error GoTo EH
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: itIdentificados con CSV '7, 8, 9'"
    logs(1) = "2. Act: IndicadorRiesgosV2_SqlDetalle"
    logs(2) = "3. Assert: SQL no vacio y contiene los IDs normalizados"
    logs(3) = "4. Assert: SQL NO contiene la cadena original con espacios arbitrarios del caller"

    Dim sql As String
    sql = IndicadorRiesgosV2_SqlDetalle(itIdentificados, DateSerial(2026, 1, 1), DateSerial(2026, 12, 31), "7, 8, 9")

    If Len(sql) = 0 Then
        Test_IndicadorRiesgos_SqlDetalle_CsvValidoGeneraSqlConIds = BuildFail("SQL no debe estar vacio para CSV valido", logs)
        Exit Function
    End If

    If InStr(1, sql, "In (7, 8, 9)", vbTextCompare) = 0 Then
        Test_IndicadorRiesgos_SqlDetalle_CsvValidoGeneraSqlConIds = BuildFail("El SQL debe contener 'In (7, 8, 9)'. SQL='" & sql & "'", logs)
        Exit Function
    End If

    ' No debe colarse ningun caracter no-entero del input original
    If InStr(1, sql, "1=1", vbTextCompare) > 0 Or InStr(1, sql, "UNION", vbTextCompare) > 0 Then
        Test_IndicadorRiesgos_SqlDetalle_CsvValidoGeneraSqlConIds = BuildFail("El SQL contiene un payload sospechoso. SQL='" & sql & "'", logs)
        Exit Function
    End If

    Test_IndicadorRiesgos_SqlDetalle_CsvValidoGeneraSqlConIds = BuildOk("sql_ok", logs)
    Exit Function
EH:
    Test_IndicadorRiesgos_SqlDetalle_CsvValidoGeneraSqlConIds = BuildFail("Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function
