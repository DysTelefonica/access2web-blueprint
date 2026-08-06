Option Compare Database
Option Explicit

' Form_FormEjercito — REWORK (2026-06-26)
' Thin UI wiring for the Ejercito Alta/Edición form.
'
' Implements access-vba-e2e-methodology rule #1: forms are thin UI wiring.

' === Module-level state (all declarations at top per vba-access §10.1) ==============

Public Event Alta()
Public Event Editado()

Private m_Error As String


' === Local helpers (all at top per vba-access §10.1) =================================

Private Function Ejercito_ExtraerValores() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Ejercito") = CStr(Nz(Me.Ejercito.Value, ""))
    d("DESCRIPCION") = CStr(Nz(Me.DESCRIPCION.Value, ""))
    Set Ejercito_ExtraerValores = d
End Function

Private Function Ejercito_EsAdministrador() As Boolean
    If EsAdministrador = EnumSiNo.Sí Then
        Ejercito_EsAdministrador = True
    Else
        Ejercito_EsAdministrador = False
    End If
End Function

Private Function Ejercito_Modo() As String
    If m_ObjEjercitoActivo Is Nothing Then
        Ejercito_Modo = "alta"
    Else
        Ejercito_Modo = "edicion"
    End If
End Function

Private Function Ejercito_IDEntidad() As String
    If m_ObjEjercitoActivo Is Nothing Then
        Ejercito_IDEntidad = ""
    Else
        If m_ObjEjercitoActivo.IDEjercito <> "" Then
            Ejercito_IDEntidad = CStr(m_ObjEjercitoActivo.IDEjercito)
        Else
            Ejercito_IDEntidad = ""
        End If
    End If
End Function

Private Sub Ejercito_RenderResultadoInicializar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim titulo As String
    titulo = CStr(p_Payload("titulo"))
    Me.lblTitulo.Caption = titulo

    If p_Payload.Exists("entidad") Then
        Dim entity As Object
        Set entity = p_Payload("entidad")
        If Not entity Is Nothing Then
            If entity.Exists("Ejercito") Then Me.Ejercito.Value = CStr(entity("Ejercito"))
            If entity.Exists("DESCRIPCION") Then Me.DESCRIPCION.Value = CStr(entity("DESCRIPCION"))
        End If
    End If

    FormInteraction_EstablecerPropiedad Me, "ComandoRegistrar", "Enabled", Ejercito_EsAdministrador()
End Sub


' === Form lifecycle / event handlers ===================================================

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Ajustar Me

    Dim modo As String
    modo = Ejercito_Modo()

    Dim idEntidad As String
    idEntidad = Ejercito_IDEntidad()

    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim initJson As String
    initJson = modEjercitoAltaHelper.EjercitoAlta_Abrir_Inicializar(modo, idEntidad, p_Entidad, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim initParsed As Object
    Set initParsed = JsonConverter.ParseJson(initJson)
    If initParsed("ok") = True Then
        Ejercito_RenderResultadoInicializar initParsed("payload")
    End If

    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    On Error Resume Next
    CorreoAlAdministrador m_Error
    On Error GoTo 0
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub cmdSalir_Click()
    On Error Resume Next
    Dim cerrarJson As String
    cerrarJson = modEjercitoAltaHelper.EjercitoAlta_Cerrar(m_Error)
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub ComandoAyuda_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    AbrirAyuda m_Error
    If m_Error <> "" Then Err.Raise 1000

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoRegistrar_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    MostrarPopupProgreso "Procesando...", "Guardando..."

    m_Error = ""

    Dim valores As Object
    Set valores = Ejercito_ExtraerValores()

    Dim p_Error As String
    Dim json As String
    json = modEjercitoAltaHelper.EjercitoAlta_Registrar(valores, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("ok") <> True Then Err.Raise 1001, , "helper returned ok=false"

    Dim modo As String
    modo = CStr(parsed("payload")("modo"))

    If modo = "alta" Then
        RaiseEvent Alta
    ElseIf modo = "edicion" Then
        RaiseEvent Editado
    End If

    VBA.DoEvents
    CerrarPopupProgreso
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
