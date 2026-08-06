Option Compare Database
Option Explicit

Private m_Error As String
Private m_ObjTecnicoAlInicio As Tecnico

Public Event Alta(m_ObjTecnico As Tecnico)
Public Event Editado(m_ObjTecnico As Tecnico)
Public Event Borrado()

Private Sub ComandoRegistrar_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    
    If m_ObjTecnicoActivo Is Nothing Then
        m_Error = "No se conoce el técnico activo"
        Err.Raise 1000
    End If
    RellenarDatosDeColeDeFormulario m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjTecnicoActivo.ParaAlta = EnumSino.No Then
        If Not HaHabidoCambios(m_Error) Then
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            m_Error = "Ningún cambio que guardar"
            Err.Raise 1000
        End If
    End If
    m_ObjTecnicoActivo.Registrar m_ObjTecnicoAlInicio, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjTecnicoActivo.ParaAlta = EnumSino.Sí Then
       RaiseEvent Alta(m_ObjTecnicoActivo)
        
    ElseIf m_ObjTecnicoActivo.ParaAlta = EnumSino.No Then
       RaiseEvent Editado(m_ObjTecnicoActivo)
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
'    Me.InsideHeight = 3675
'    Me.InsideWidth = 9510
    AjustarTamaño Me
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    Me.Repaint
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
End Sub

Private Function RellenarComboTipos(Optional ByRef p_Error As String) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim cmb As ComboBox
    On Error GoTo errores
    
    Set cmb = Me.Tipo
    cmb.RowSource = ""
    m_SQL = "SELECT TbTipoTecnicoPrecios.TIPOTECNICO " & _
            "FROM TbTipoTecnicoPrecios " & _
            "WHERE FechaFinal Is Null " & _
            "ORDER BY TbTipoTecnicoPrecios.TIPOTECNICO;"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            cmb.AddItem .Fields("TIPOTECNICO")
            .MoveNext
        Loop
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboTipos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim cmb As ComboBox
    Dim ctl As Control
    Dim m_Titulo As String
    Dim m_NombreCampo As String
    Dim m_Valor As String
    
    On Error GoTo errores
    RellenarComboTipos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjTecnicoActivo Is Nothing Then
        m_Titulo = "Alta de Técnico"
        Me.lblTitulo1.Caption = m_Titulo
        Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
        Me.Caption = Me.lblTitulo1.Caption
        Set m_ObjTecnicoActivo = New Tecnico
        m_ObjTecnicoActivo.ParaAlta = EnumSino.Sí
        Exit Function
    End If
    m_ObjTecnicoActivo.ParaAlta = EnumSino.No
    Set m_ObjTecnicoAlInicio = Constructor.getTecnico(m_ObjTecnicoActivo.Alias, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjTecnicoAlInicio Is Nothing Then
        p_Error = "Técnico no reconocido en el Sistema"
        Err.Raise 1000
    End If
    For Each ctl In Me.Controls
        If InStr(1, ctl.Tag, "DATO") <> 0 Then
            m_NombreCampo = ctl.Name
            m_Valor = m_ObjTecnicoAlInicio.getPropiedad(m_NombreCampo, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_Valor <> "" Then
                ctl.Value = m_Valor
            End If
        End If
        m_NombreCampo = ""
        m_Valor = ""
    Next
    m_Titulo = "Edición de " & m_ObjTecnicoAlInicio.Nombre
    Me.lblTitulo1.Caption = m_Titulo
    Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
    Me.Caption = Me.lblTitulo1.Caption
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Function RellenarDatosDeColeDeFormulario(Optional ByRef p_Error As String) As String
    
    Dim m_Dato As String
    Dim m_NombreCampo As String
    Dim ctl As Control
    
    On Error GoTo errores
    For Each ctl In Me.Controls
        If InStr(1, Nz(ctl.Tag, ""), "DATO") > 0 Then
            m_NombreCampo = ctl.Name
            'If m_nombreCampo = "CausaRaiz" Then Stop
            m_Dato = Nz(ctl.Value, "")
            m_ObjTecnicoActivo.SetPropiedad m_NombreCampo, m_Dato, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarDatosDeColeDeFormulario ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean

    On Error GoTo errores
    
    If m_ObjTecnicoAlInicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If
    With m_ObjTecnicoAlInicio
        If .Alias <> Nz(Me.Alias, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        
        If .Tipo <> Nz(Me.Tipo, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Nombre <> Nz(Me.Nombre, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Telefono <> Nz(Me.Telefono, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Email <> Nz(Me.Email, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .FechaAlta <> Nz(Me.FechaAlta, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .FechaBaja <> Nz(Me.FechaBaja, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function









