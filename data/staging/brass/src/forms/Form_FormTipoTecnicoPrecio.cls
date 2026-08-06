Option Compare Database
Option Explicit

Public Event Alta(m_ObjTipoTecnicoPrecio As TipoTecnicoPrecio)
Public Event Edicion(m_ObjTipoTecnicoPrecio As TipoTecnicoPrecio)
Private m_ObjTipoTecnicoPrecioAlInicio As TipoTecnicoPrecio
Private m_Error As String

Public Function HaHabidoCambios(Optional ByRef p_Error) As EnumSino
    
    On Error GoTo errores
    If m_ObjTipoTecnicoPrecioAlInicio Is Nothing Then
        HaHabidoCambios = EnumSino.Sí
        Exit Function
    End If
    If m_ObjTipoTecnicoPrecioAlInicio.FechaInicial <> Nz(Me.FechaInicial, "") Then
        HaHabidoCambios = EnumSino.Sí
        Exit Function
    End If
    If m_ObjTipoTecnicoPrecioAlInicio.FechaFinal <> Nz(Me.FechaFinal, "") Then
        HaHabidoCambios = EnumSino.Sí
        Exit Function
    End If
    If m_ObjTipoTecnicoPrecioAlInicio.PrecioHoraLab <> Nz(Me.PrecioHoraLab, "") Then
        HaHabidoCambios = EnumSino.Sí
        Exit Function
    End If
    If m_ObjTipoTecnicoPrecioAlInicio.PrecioHoraExtra <> Nz(Me.PrecioHoraExtra, "") Then
        HaHabidoCambios = EnumSino.Sí
        Exit Function
    End If
    HaHabidoCambios = EnumSino.No
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    If Not m_ObjTipoTecnicoPrecioActivo Is Nothing Then
        If m_ObjTipoTecnicoPrecioActivo.ParaAlta = EnumSino.No Then
            Set m_ObjTipoTecnicoPrecioAlInicio = Constructor.getTipoTecnicoPrecio(m_ObjTipoTecnicoPrecioActivo.IDTipoTecnicoFecha, , , p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjTipoTecnicoPrecioAlInicio Is Nothing Then
                p_Error = "No se puede obtener los datos del precio del perfil"
                Err.Raise 1000
            End If
             With m_ObjTipoTecnicoPrecioAlInicio
                Me.TIPOTECNICO = .TIPOTECNICO
                If IsDate(.FechaInicial) Then
                    Me.FechaInicial = .FechaInicial
                End If
                If IsDate(.FechaFinal) Then
                    Me.FechaFinal = .FechaFinal
                End If
                If IsNumeric(.PrecioHoraLab) Then
                    Me.PrecioHoraLab = .PrecioHoraLab
                End If
                If IsNumeric(.PrecioHoraExtra) Then
                    Me.PrecioHoraExtra = .PrecioHoraExtra
                End If
            End With
        Else
            Me.TIPOTECNICO = m_ObjTipoTecnicoPrecioActivo.TIPOTECNICO
            Set m_ObjTipoTecnicoPrecioActivo = New TipoTecnicoPrecio
            m_ObjTipoTecnicoPrecioActivo.TIPOTECNICO = Me.TIPOTECNICO
            m_ObjTipoTecnicoPrecioActivo.ParaAlta = EnumSino.Sí
        End If
    Else
        Set m_ObjTipoTecnicoPrecioActivo = New TipoTecnicoPrecio
        m_ObjTipoTecnicoPrecioActivo.ParaAlta = EnumSino.Sí
    End If
    
   Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Sub ComandoRegistrar_Click()
    
    Dim m_ConCambios As EnumSino
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    m_ConCambios = HaHabidoCambios(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ConCambios = EnumSino.No Then
        m_Error = "Ningún cambio que registrar"
        Err.Raise 1000
    End If
    
    With m_ObjTipoTecnicoPrecioActivo
        .TIPOTECNICO = Nz(Me.TIPOTECNICO, "")
        .FechaInicial = Nz(Me.FechaInicial, "")
        .FechaFinal = Nz(Me.FechaFinal, "")
        .PrecioHoraLab = Nz(Me.PrecioHoraLab, "")
        .PrecioHoraExtra = Nz(Me.PrecioHoraExtra, "")
        .Registrar m_ObjTipoTecnicoPrecioAlInicio, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    If m_ObjTipoTecnicoPrecioAlInicio Is Nothing Then
        RaiseEvent Alta(m_ObjTipoTecnicoPrecioActivo)
    Else
        RaiseEvent Edicion(m_ObjTipoTecnicoPrecioActivo)
    End If
    
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
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
    
    AjustarTamaño Me

    Me.Caption = Me.lblTitulo1.Caption
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
