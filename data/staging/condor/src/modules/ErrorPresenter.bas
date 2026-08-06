Attribute VB_Name = "ErrorPresenter"

Option Explicit
' ==========================================================================
' MÓDULO: ErrorPresenter.bas
' RESPONSABILIDAD: Único punto de presentación de errores al usuario.
'                 Todos los forms deben usar este módulo, nunca acceder
'                 a g_objLastError directamente.
' ==========================================================================

' Muestra el error al usuario y opcionalmente lo loguea
' Este método es SEGURO — nunca lanza errores (wrapped internally)
Public Sub MostrarError(ByVal formName As String, Optional ByVal logThisError As Boolean = True)
    On Error Resume Next  ' SILENT — this method never throws
    
    If g_objLastError Is Nothing Then
        MsgBox "Error inesperado. Contacte con el administrador.", vbCritical, "CONDOR"
        Exit Sub
    End If
    
    Dim msg As String
    Dim estilo As VbMsgBoxStyle
    Dim titulo As String
    
    ' VALIDACIÓN DE NEGOCIO (NO ERROR TÉCNICO)
    ' Regla UX: los faltantes funcionales deben mostrarse como advertencia amigable.
    If EsValidacionFuncional(g_objLastError) Then
        ' Alineado con comportamiento histórico (main): validación funcional en warning.
        titulo = "Validación: Datos Requeridos"
        estilo = vbExclamation + vbOKOnly
        msg = FormatearDescripcionValidacion(g_objLastError.description)
    ElseIf Err.Number = vbObjectError + 512 Or g_objLastError.Number <> 0 Then
        ' Error de aplicación CONDOR — mostrar descripción completa
        titulo = "CONDOR - Error en " & formName
        estilo = vbCritical
        msg = g_objLastError.FullDescription
    Else
        ' Error VBA genérico
        titulo = "CONDOR - Error en " & formName
        estilo = vbCritical
        msg = "Error " & Err.Number & ": " & Err.description
    End If
    
    MsgBox msg, estilo, titulo
    
    ' Loguear si hay error pendiente y se solicita
    If logThisError And Not (g_objLastError Is Nothing) Then
        Call ErrorLogger.LogToTable(g_objLastError)
    End If
    
    ' Limpiar el error global después de mostrar
    Set g_objLastError = Nothing
End Sub

' Clasifica si un CondorError 513 corresponde a validación funcional de usuario
' (faltan campos / precondiciones de flujo) y por tanto debe mostrarse como warning.
' Mantiene como error técnico los 513 operativos (rutas, plantillas, conexión, etc.).
Private Function EsValidacionFuncional(ByVal ce As CondorError) As Boolean
    Dim src As String
    Dim desc As String

    On Error GoTo Fallback

    If ce Is Nothing Then Exit Function
    If ce.Number <> 513 Then Exit Function

    src = UCase$(Trim$(Nz(ce.source, "")))
    desc = UCase$(Trim$(Nz(ce.description, "")))

    ' 1) Fuente explícita de validación de formularios/subformularios
    If src = "VALIDACIÓN" Or src = "VALIDACION" Then
        EsValidacionFuncional = True
        Exit Function
    End If

    ' 2) Reglas de negocio y validaciones de flujo
    If InStr(1, src, "WORKFLOWSERVICIO.CUMPLEPASOA", vbTextCompare) > 0 Then
        EsValidacionFuncional = True
        Exit Function
    End If

    If InStr(1, src, "WORKFLOWSERVICIO.EJECUTARTRANSICION", vbTextCompare) > 0 Then
        EsValidacionFuncional = True
        Exit Function
    End If

    If InStr(1, src, "REGLA DE NEGOCIO", vbTextCompare) > 0 Then
        EsValidacionFuncional = True
        Exit Function
    End If

    ' 3) Heurística por mensaje funcional (faltan campos / completar / guardar)
    If InStr(1, desc, "FALTAN CAMPOS OBLIGATORIOS", vbTextCompare) > 0 Or _
       InStr(1, desc, "DEBE CUMPLIMENTAR", vbTextCompare) > 0 Or _
       InStr(1, desc, "DEBE COMPLETAR", vbTextCompare) > 0 Or _
       InStr(1, desc, "DEBE ESTAR COMPLETA", vbTextCompare) > 0 Or _
       InStr(1, desc, "GUARDE PRIMERO", vbTextCompare) > 0 Or _
       InStr(1, desc, "TRANSICIÓN NO PERMITIDA PARA SU ROL", vbTextCompare) > 0 Or _
       InStr(1, desc, "TRANSICION NO PERMITIDA PARA SU ROL", vbTextCompare) > 0 Or _
       InStr(1, desc, "ES OBLIGATORIO ADJUNTAR", vbTextCompare) > 0 Then
        EsValidacionFuncional = True
        Exit Function
    End If

    EsValidacionFuncional = False
    Exit Function

Fallback:
    EsValidacionFuncional = False
End Function

' ==========================================
'  FORMATEO DE DESCRIPCIÓN DE VALIDACIÓN 513
' ==========================================
' Alineado con el formato clásico que existía antes de centralizar en ErrorPresenter.
Private Function FormatearDescripcionValidacion(ByVal descripcion As String) As String
    Dim lineas() As String, i As Long
    Dim grupos As Object ' Scripting.Dictionary late-bound
    Dim nombreGrupo As String, item As String, posCierre As Long
    Dim clave As Variant, sb As String

    On Error GoTo FinSeguro

    Set grupos = CreateObject("Scripting.Dictionary")
    grupos.CompareMode = 1 ' TextCompare

    lineas = Split(Nz(descripcion, ""), vbCrLf)
    For i = LBound(lineas) To UBound(lineas)
        Dim ln As String
        ln = Trim(lineas(i))
        If ln <> "" Then
            nombreGrupo = "Otros"
            item = ln

            ' Limpieza defensiva de prefijos típicos
            If Left$(ln, 1) = "?" Or Left$(ln, 1) = "•" Or Left$(ln, 1) = "-" Then
                ln = Trim$(Mid$(ln, 2))
                item = ln
            End If

            If Left$(ln, 1) = "(" Then
                posCierre = InStr(2, ln, ")")
                If posCierre > 2 Then
                    nombreGrupo = Trim(Mid$(ln, 2, posCierre - 2))
                    item = Trim(Mid$(ln, posCierre + 1))
                    If Left$(item, 1) = "-" Then item = Trim(Mid$(item, 2))
                End If
            End If

            If Not grupos.Exists(nombreGrupo) Then
                grupos.Add nombreGrupo, New Collection
            End If
            grupos(nombreGrupo).Add item
        End If
    Next i

    sb = "Faltan campos obligatorios:" & vbCrLf & vbCrLf
    For Each clave In grupos.Keys
        sb = sb & "[" & CStr(clave) & "]" & vbCrLf
        For i = 1 To grupos(clave).count
            sb = sb & "  • " & CStr(grupos(clave).item(i)) & vbCrLf
        Next i
        sb = sb & vbCrLf
    Next clave

    sb = sb & "Completá los campos indicados y volvé a intentar."

    FormatearDescripcionValidacion = sb
    Exit Function

FinSeguro:
    FormatearDescripcionValidacion = "Faltan campos obligatorios." & vbCrLf & "Completá los datos requeridos y volvé a intentar."
End Function

' Verifica si hay un error pendiente por procesar
Public Function HayErrorPendiente() As Boolean
    On Error Resume Next
    HayErrorPendiente = Not (g_objLastError Is Nothing)
End Function

' Limpia el error global sin mostrarlo (para casos donde se recuperó el error)
Public Sub LimpiarError()
    On Error Resume Next
    Set g_objLastError = Nothing
End Sub


