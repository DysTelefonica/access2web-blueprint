Attribute VB_Name = "Test_modProyectoCuadroHelper"
Option Compare Database
Option Explicit

' ============================================================
' Test_modProyectoCuadroHelper - TDD atoms for modProyectoCuadroHelper
'
' Helper: ObtenerCeldaFechaUltimaEdicion, ObtenerCeldaNumeroEdicion
'   Signatures:
'     Public Function ObtenerCeldaFechaUltimaEdicion( _
'                                         ByVal p_FechaPublicacion As String, _
'                                         ByRef p_Error As String) As String
'     Public Function ObtenerCeldaNumeroEdicion( _
'                                         ByVal p_Edicion As Edicion, _
'                                         ByRef p_Error As String) As String
'
' Arquitectura: helpers puros sobre datos primitivos y referencias del modelo.
' Regla semantica (Punto 17 del acta + decision 2026-07-08):
'   - "Fecha Edicion" / "F.Ultima Ed." = SIEMPRE FechaPublicacion
'   - Si no hay FechaPublicacion: celda vacia (sin "-----" ni label)
'   - Si no hay EdicionUltimaPublicada: celda "Nº Edicion" vacia
'
' Skill: access-vba-tdd v2.4.3
' SDD:    pendiente archivar (Punto 17 acta 2026-06-25)
' ============================================================

' --- JSON helpers wrapper (project convention) ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' ----------------------------------------------------------------
' ATOM 1: Happy - FechaUltimaEdicion con valor (proyecto con edicion publicada)
'   p_FechaPublicacion = "15/03/2026" -> helper devuelve "15/03/2026"
' ----------------------------------------------------------------
Public Function Test_ProyectoCuadro_Happy_FechaUltimaEdicionConValor() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    Dim m_Fecha As String
    Dim m_Resultado As String
    Dim m_Error As String

    logs(0) = "1. Arrange: p_FechaPublicacion='15/03/2026'"
    logs(1) = "2. Act: ObtenerCeldaFechaUltimaEdicion"
    logs(2) = "3. Assert: devuelve '15/03/2026' (la fecha de publicacion)"

    m_Fecha = "15/03/2026"
    m_Resultado = modProyectoCuadroHelper.ObtenerCeldaFechaUltimaEdicion(m_Fecha, m_Error)

    If m_Resultado <> m_Fecha Then
        logs(3) = "3. Assert FAIL: expected '" & m_Fecha & "', got '" & m_Resultado & "'"
        Test_ProyectoCuadro_Happy_FechaUltimaEdicionConValor = _
            BuildFail("La celda debe devolver la fecha de publicacion", logs)
        Exit Function
    End If

    If Len(m_Error) > 0 Then
        logs(3) = "3. Assert FAIL: p_Error poblado: " & m_Error
        Test_ProyectoCuadro_Happy_FechaUltimaEdicionConValor = _
            BuildFail("p_Error debe estar vacio cuando el resultado es correcto", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: '" & m_Resultado & "' con p_Error vacio"
    Test_ProyectoCuadro_Happy_FechaUltimaEdicionConValor = _
        BuildOk(m_Resultado, logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_ProyectoCuadro_Happy_FechaUltimaEdicionConValor = _
        BuildFail("ObtenerCeldaFechaUltimaEdicion raised: " & Err.description, logs)
End Function

' ----------------------------------------------------------------
' ATOM 2: Sad - FechaUltimaEdicion vacia (proyecto sin edicion publicada)
'   p_FechaPublicacion = "" -> helper devuelve "" (regla semantica: no mostrar nada)
' ----------------------------------------------------------------
Public Function Test_ProyectoCuadro_Sad_FechaUltimaEdicionVacia() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    Dim m_Resultado As String
    Dim m_Error As String

    logs(0) = "1. Arrange: p_FechaPublicacion='' (sin edicion publicada)"
    logs(1) = "2. Act: ObtenerCeldaFechaUltimaEdicion"
    logs(2) = "3. Assert: devuelve '' (regla semantica: no mostrar nada, NO '-----')"

    m_Resultado = modProyectoCuadroHelper.ObtenerCeldaFechaUltimaEdicion("", m_Error)

    If m_Resultado <> "" Then
        logs(3) = "3. Assert FAIL: expected '' (celda vacia), got '" & m_Resultado & "'"
        Test_ProyectoCuadro_Sad_FechaUltimaEdicionVacia = _
            BuildFail("La celda debe quedar vacia cuando no hay FechaPublicacion (regla semantica de Natalia)", logs)
        Exit Function
    End If

    If InStr(1, m_Resultado, "-", vbTextCompare) > 0 Then
        logs(3) = "3. Assert FAIL: contiene guion (parece '-----')"
        Test_ProyectoCuadro_Sad_FechaUltimaEdicionVacia = _
            BuildFail("NO debe haber '-----' como placeholder, debe quedar vacia", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: '' sin placeholder"
    Test_ProyectoCuadro_Sad_FechaUltimaEdicionVacia = _
        BuildOk(m_Resultado, logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_ProyectoCuadro_Sad_FechaUltimaEdicionVacia = _
        BuildFail("ObtenerCeldaFechaUltimaEdicion raised: " & Err.description, logs)
End Function

' ----------------------------------------------------------------
' ATOM 3: Happy - NumeroEdicion con valor (proyecto con edicion publicada)
'   p_Edicion = New Edicion (con .Edicion = "3") -> helper devuelve "3"
' ----------------------------------------------------------------
Public Function Test_ProyectoCuadro_Happy_NumeroEdicionConValor() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    Dim m_Edicion As Edicion
    Dim m_Resultado As String
    Dim m_Error As String

    logs(0) = "1. Arrange: p_Edicion = New Edicion, .Edicion='3'"
    logs(1) = "2. Act: ObtenerCeldaNumeroEdicion"
    logs(2) = "3. Assert: devuelve '3' (el numero de la edicion)"

    Set m_Edicion = New Edicion
    m_Edicion.Edicion = "3"
    m_Resultado = modProyectoCuadroHelper.ObtenerCeldaNumeroEdicion(m_Edicion, m_Error)

    If m_Resultado <> "3" Then
        logs(3) = "3. Assert FAIL: expected '3', got '" & m_Resultado & "'"
        Test_ProyectoCuadro_Happy_NumeroEdicionConValor = _
            BuildFail("La celda debe devolver el numero de la edicion", logs)
        Exit Function
    End If

    If Len(m_Error) > 0 Then
        logs(3) = "3. Assert FAIL: p_Error poblado: " & m_Error
        Test_ProyectoCuadro_Happy_NumeroEdicionConValor = _
            BuildFail("p_Error debe estar vacio cuando el resultado es correcto", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: '3' con p_Error vacio"
    Set m_Edicion = Nothing
    Test_ProyectoCuadro_Happy_NumeroEdicionConValor = _
        BuildOk(m_Resultado, logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_ProyectoCuadro_Happy_NumeroEdicionConValor = _
        BuildFail("ObtenerCeldaNumeroEdicion raised: " & Err.description, logs)
End Function

' ----------------------------------------------------------------
' ATOM 4: Sad - NumeroEdicion vacio (proyecto sin edicion publicada)
'   p_Edicion = Nothing -> helper devuelve "" (regla semantica: no mostrar nada)
' ----------------------------------------------------------------
Public Function Test_ProyectoCuadro_Sad_NumeroEdicionVacio() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    Dim m_Resultado As String
    Dim m_Error As String

    logs(0) = "1. Arrange: p_Edicion = Nothing (proyecto sin edicion publicada)"
    logs(1) = "2. Act: ObtenerCeldaNumeroEdicion"
    logs(2) = "3. Assert: devuelve '' (regla semantica: no mostrar nada)"

    m_Resultado = modProyectoCuadroHelper.ObtenerCeldaNumeroEdicion(Nothing, m_Error)

    If m_Resultado <> "" Then
        logs(3) = "3. Assert FAIL: expected '' (celda vacia), got '" & m_Resultado & "'"
        Test_ProyectoCuadro_Sad_NumeroEdicionVacio = _
            BuildFail("La celda debe quedar vacia cuando no hay edicion publicada", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: '' cuando no hay edicion"
    Test_ProyectoCuadro_Sad_NumeroEdicionVacio = _
        BuildOk(m_Resultado, logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_ProyectoCuadro_Sad_NumeroEdicionVacio = _
        BuildFail("ObtenerCeldaNumeroEdicion raised: " & Err.description, logs)
End Function
