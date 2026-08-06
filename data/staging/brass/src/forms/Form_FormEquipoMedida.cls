Option Compare Database
Option Explicit
Private m_Error As String

Private m_ObjEquipoAlInicio As EquipoMedida
Public Event EquipoModificado()
Public Event EquipoAlta()


Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim m_Titulo As String
    Dim ctl As Control
    Dim m_Valor As String
    On Error GoTo errores
    
    
    If m_ObjEquipoMedidaActivo Is Nothing Then
        m_Titulo = "ALTA DE EQUIPO DE MEDIDA"
        Me.lblTitulo1.Caption = m_Titulo
        Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
        Me.Caption = Me.lblTitulo1.Caption
        Exit Function
    End If
    Set m_ObjEquipoAlInicio = Constructor.getEquipoMedida(m_ObjEquipoMedidaActivo.IDEquipoMedida, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjEquipoAlInicio Is Nothing Then
        p_Error = "No se encuentra el equipo de medida registrado"
        Err.Raise 1000
    End If
    m_Titulo = "EDICIÓN DE EQUIPO DE MEDIDA"
    Me.lblTitulo1.Caption = m_Titulo
    Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
    Me.Caption = Me.lblTitulo1.Caption
    For Each ctl In Me.Controls
        If InStr(1, ctl.Tag, "DATO") <> 0 Then
            m_Valor = m_ObjEquipoAlInicio.getPropiedad(ctl.Name, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_Valor <> "" Then
                ctl.Value = m_Valor
            End If
        End If
    Next
    
    EstablecerDatos = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Function RellenarObjeto(Optional ByRef p_Error As String) As String
    
    Dim ctl As Control
    Dim m_Valor As String
    On Error GoTo errores
    
    
    If m_ObjEquipoMedidaActivo Is Nothing Then
        Set m_ObjEquipoMedidaActivo = New EquipoMedida
    End If
    For Each ctl In Me.Controls
        If InStr(1, ctl.Tag, "DATO") <> 0 Then
            m_Valor = Nz(ctl.Value, "")
            m_ObjEquipoMedidaActivo.SetPropiedad ctl.Name, m_Valor, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    Next
    
    
    RellenarObjeto = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarObjeto ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    
    Dim ctl As Control
    Dim m_Valor As String
    Dim m_valorAlInicio As String
    On Error GoTo errores
    
    If m_ObjEquipoAlInicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If
    For Each ctl In Me.Controls
        If InStr(1, ctl.Tag, "DATO") <> 0 Then
            m_valorAlInicio = m_ObjEquipoAlInicio.getPropiedad(ctl.Name, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            m_Valor = Nz(ctl.Value, "")
            If m_valorAlInicio <> m_Valor Then
                HaHabidoCambios = True
                Exit Function
            End If
        End If
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Sub ComandoRegistrar_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If Not HaHabidoCambios(m_Error) Then
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No ha habido cambios"
        Err.Raise 1000
    End If
    RellenarObjeto m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjEquipoAlInicio Is Nothing Then
        m_ObjEquipoMedidaActivo.ParaAlta = EnumSino.No
    Else
        m_ObjEquipoMedidaActivo.ParaAlta = EnumSino.Sí
    End If
    m_ObjEquipoMedidaActivo.Registrar m_ObjEquipoAlInicio, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjEquipoAlInicio Is Nothing Then
        RaiseEvent EquipoAlta
    Else
        RaiseEvent EquipoModificado
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub Form_Load()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me
    
    Me.Caption = Me.lblTitulo1.Caption
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub








