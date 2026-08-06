Option Compare Database
Option Explicit
Private m_ObjProyectodelRiesgo As Proyecto
Private m_ObjNCAlInicio As nc
Public Event NCRegistrada()
Private m_Error As String


Private Sub ComandoRegistrar_Click()
    
    Dim m_FechaRef As Date
    
    On Error GoTo errores
    
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    If Not HaHabidoCambios(m_Error) Then
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No hay cambios que guardar"
        Err.Raise 1000
    End If
    RellenarDatosDeColeDeFormulario m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_ObjNCActiva.Registrar m_ObjNCAlInicio, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjNCAlInicio Is Nothing Then
        With m_ObjRiesgoMaterializadoActivo
            
            .FechaDecison = CStr(m_FechaRef)
            
        
            .RegistrarParaNC m_ObjNCActiva.IDNoConformidad, CStr(m_FechaRef), m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            
        End With
   
    End If
    RaiseEvent NCRegistrada
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
   
    Dim m_Titulo As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me

    If m_ObjRiesgoMaterializadoActivo Is Nothing Then
        m_Error = "No se sabe qué riesgo materializado está activo"
        Err.Raise 1000
    End If
    If EsTecnico = EnumSiNo.Sí Then
        Me.AllowEdits = False
    Else
        Me.AllowEdits = True
    End If
    EstablecerCombos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    
    
    On Error GoTo errores
    
    If m_ObjRiesgoMaterializadoActivo Is Nothing Then
        p_Error = "No se sabe el riesgo origen"
        Err.Raise 1000
    End If
    With m_ObjRiesgoMaterializadoActivo
        Set m_ObjNCActiva = .nc
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjProyectodelRiesgo = .Proyecto
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjProyectodelRiesgo Is Nothing Then
            p_Error = "No se conoce la gestión de riesgos"
            Err.Raise 1000
        End If
    End With
    If m_ObjNCActiva Is Nothing Then
        Me.Caption = "Alta de No Conformidad"
        EstablecerDatoPrevios p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    Set m_ObjNCAlInicio = Constructor.getNC(m_ObjNCActiva.IDNoConformidad, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    With m_ObjNCActiva
        Me.Caption = "Detalle de No Conformidad"
        If .Descripcion <> "" Then
            Me.Descripcion = .Descripcion
        End If
        If .Causa <> "" Then
            Me.Causa = .Causa
        End If
        If .ACR <> "" Then
            Me.ACR = .ACR
        End If
        If .Expediente <> "" Then
            Me.Expediente = .Expediente
        End If
        If .Tipo <> "" Then
            Me.Tipo = .Tipo
        End If
        If .Proyecto <> "" Then
            Me.Proyecto = .Proyecto
        End If
        If .Vehiculo <> "" Then
            Me.Vehiculo = .Vehiculo
        End If
        If .EntidadResponsable <> "" Then
            Me.EntidadResponsable = .EntidadResponsable
        End If
        If .FechaApertura <> "" Then
            Me.FechaApertura = .FechaApertura
        End If
        If .FPREVCIERRE <> "" Then
            Me.FPREVCIERRE = .FPREVCIERRE
        End If
        If .ResponsableTelefonica <> "" Then
            Me.ResponsableTelefonica = .ResponsableTelefonica
        End If
        If .Notas <> "" Then
            Me.Notas = .Notas
        End If
        If .RequiereControlEficacia <> "" Then
            Me.RequiereControlEficacia = .RequiereControlEficacia
        End If
        
        If .FechaPrevistaControlEficacia <> "" Then
            Me.FechaPrevistaControlEficacia = .FechaPrevistaControlEficacia
        End If
    End With
    If Me.AllowEdits = True Then
        Me.ComandoRegistrar.Enabled = True
    Else
        Me.ComandoRegistrar.Enabled = False
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Establecerdatos ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Private Function EstablecerDatoPrevios( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_FechaRef As Date
    On Error GoTo errores
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    Me.FechaApertura = CStr(m_FechaRef)
    
    
    Me.Expediente = m_ObjRiesgoMaterializadoActivo.Proyecto.Proyecto
    p_Error = m_ObjRiesgoMaterializadoActivo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    With m_ObjRiesgoMaterializadoActivo.riesgo
        If .DetectadoPor <> "" Then
            Me.ResponsableTelefonica = .DetectadoPor
        End If
    End With
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatoPrevios ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function EstablecerCombos(Optional ByRef p_Error As String) As String
    
   
    On Error GoTo errores
    
    p_Error = ""
    Me.Expediente.RowSource = ""
    RellenarComboCached Me.Expediente, "TbNoConformidades", "Expediente", getdb(), p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Tipo.RowSource = ""
    EstablecerComboTipo Me.Tipo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Proyecto.RowSource = ""
    RellenarComboCached Me.Proyecto, "TbNoConformidades", "PROYECTO", getdb(), p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Vehiculo.RowSource = ""
    RellenarComboCached Me.Vehiculo, "TbNoConformidades", "VEHICULO", getdb(), p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.ResponsableTelefonica.RowSource = ""
    EstablecerComboResponsables Me.ResponsableTelefonica, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerCombos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.description
    End If
End Function

Public Function EstablecerComboResponsables( _
                                            cmb As ComboBox, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim m_Col As Scripting.Dictionary
    Dim m_Tecnico As Usuario
    Dim m_ID As Variant
    
    
    On Error GoTo errores
    
    cmb.RowSource = ""
    Set m_Col = m_ObjEntorno.ColUsuariosTecnicosActivos
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 100
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_Tecnico = m_Col(m_ID)
        cmb.AddItem m_Tecnico.Nombre
        Set m_Tecnico = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerComboResponsables ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function EstablecerComboTipo( _
                                            cmb As ComboBox, _
                                            Optional ByRef p_Error As String _
                                            ) As String
   
    Dim m_Col As Scripting.Dictionary
    
    Dim m_Cod As Variant
    Dim m_Tipo As String
    
    
    On Error GoTo errores
    
    cmb.RowSource = ""
    Set m_Col = m_ObjEntorno.ColTipos
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 100
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_Cod In m_Col
        m_Tipo = m_Col(m_Cod)
        cmb.AddItem m_Cod & ";" & m_Tipo
        
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerComboTipo ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean

    On Error GoTo errores
    If m_ObjNCAlInicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If
    With m_ObjNCAlInicio
        If .Descripcion <> Nz(Me.Descripcion, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Causa <> Nz(Me.Causa, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .ACR <> Nz(Me.ACR, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Expediente <> Nz(Me.Expediente, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Tipo <> Nz(Me.Tipo, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Proyecto <> Nz(Me.Proyecto, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Vehiculo <> Nz(Me.Vehiculo, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .EntidadResponsable <> Nz(Me.EntidadResponsable, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .FechaApertura <> Nz(Me.FechaApertura, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .FPREVCIERRE <> Nz(Me.FPREVCIERRE, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .ResponsableTelefonica <> Nz(Me.ResponsableTelefonica, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Notas <> Nz(Me.Notas, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .RequiereControlEficacia <> Nz(Me.RequiereControlEficacia, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        
        
        If .FechaPrevistaControlEficacia <> Nz(Me.FechaPrevistaControlEficacia, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Private Function RellenarDatosDeColeDeFormulario(Optional ByRef p_Error As String) As String
    
    Dim m_Dato As Variant
    Dim m_NombreCampo As String
    Dim ctl As Control
    
    On Error GoTo errores
    If m_ObjNCActiva Is Nothing Then
        Set m_ObjNCActiva = New nc
        
    End If
    For Each ctl In Me.Controls
        If Nz(ctl.Tag, "") = "DATOS" Then
            m_NombreCampo = ctl.Name
            'Debug.Print m_NombreCampo
            m_Dato = Nz(ctl.value, "")
            
            m_ObjNCActiva.SetPropiedad m_NombreCampo, m_Dato, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    Next
    m_ObjNCActiva.Juridica = m_ObjProyectodelRiesgo.Juridica
    m_ObjNCActiva.EsNoConformidad = True
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarDatosDeColeDeFormulario ha devuelto el error: " & vbNewLine & Err.description
        
    End If
End Function

Private Sub RequiereControlEficacia_Change()
    If Me.AllowEdits = False Then
        Exit Sub
    End If
    If Nz(Me.RequiereControlEficacia, "") <> "Sí" Then
        Me.FechaPrevistaControlEficacia = Null
        Me.FechaPrevistaControlEficacia.Enabled = False
    Else
        Me.FechaPrevistaControlEficacia.Enabled = True
    End If
End Sub
