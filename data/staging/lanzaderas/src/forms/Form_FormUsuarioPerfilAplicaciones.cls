Option Compare Database
Option Explicit
Private m_Error As String

Private m_ObjAplicacionSeleccionada As Aplicacion


Private Sub ComandoAltaPerfil_Click()
    
    Dim m_ObjUsuarioAplicacionPermisos As UsuarioAplicacionPermisos
    Dim Admin As String
    Dim Calidad As String
    Dim CalidadAvisos As String
    Dim Secretaria As String
    Dim Tecnico As String
    Dim Economia As String
    Dim PartePermisos As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Nz(Me.ComboAplicacion, "") = "" Then
        m_Error = "Seleccione un elemento de la lista"
        Err.Raise 1000
    End If
    Set m_ObjAplicacionSeleccionada = Constructor.getAplicacion(Me.ComboAplicacion.Column(0), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjAplicacionSeleccionada Is Nothing Then
        m_Error = "Seleccione un elemento de la lista"
        Err.Raise 1000
    End If
    If Me.EsUsuarioAdministrador = True Then
        Admin = "Sí"
    Else
        Admin = "No"
    End If
    If Me.EsUsuarioCalidad = True Then
        Calidad = "Sí"
    Else
        Calidad = "No"
    End If
    If Me.EsUsuarioEconomia = True Then
        Economia = "Sí"
    Else
        Economia = "No"
    End If
    If Me.EsUsuarioSecretaria = True Then
        Secretaria = "Sí"
    Else
        Secretaria = "No"
    End If
    If Me.EsUsuarioTecnico = True Then
        Tecnico = "Sí"
    Else
        Tecnico = "No"
    End If
    If Me.EsUsuarioCalidadAvisos = True Then
        CalidadAvisos = "Sí"
    Else
        CalidadAvisos = "No"
    End If
    If Admin = "No" And Calidad = "No" And Economia = "No" And Secretaria = "No" And Tecnico = "No" And CalidadAvisos = "No" Then
        m_Error = "Para dar de alta una aplicación, al menos ha de haber uno de los permisos"
        Err.Raise 1000
    End If
    Set m_ObjUsuarioAplicacionPermisos = Constructor.getUsuarioAplicacionPermisos(m_ObjUsuarioActivo.CorreoUsuario, _
                        m_ObjAplicacionSeleccionada.IDAplicacion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjUsuarioAplicacionPermisos Is Nothing Then
        Set m_ObjUsuarioAplicacionPermisos = New UsuarioAplicacionPermisos
        m_ObjUsuarioAplicacionPermisos.CorreoUsuario = m_ObjUsuarioActivo.CorreoUsuario
    End If
    With m_ObjUsuarioAplicacionPermisos
        
        .IDAplicacion = m_ObjAplicacionSeleccionada.IDAplicacion
        .EsUsuarioAdministrador = Admin
        .EsUsuarioCalidad = Calidad
        .EsUsuarioEconomia = Economia
        .EsUsuarioSecretaria = Secretaria
        .EsUsuarioTecnico = Tecnico
        .EsUsuarioCalidadAvisos = CalidadAvisos
        .RegistrarPermisosCompletos m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Set m_ObjUsuarioActivo = Constructor.getUsuario(m_ObjUsuarioActivo.ID, , , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAltaPerfil_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminar_Click()
    Dim m_ObjUsuarioAplicacionPermisos As UsuarioAplicacionPermisos
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    If m_ObjAplicacionSeleccionada Is Nothing Then
        m_Error = "Seleccione un elemento de la lista"
        Err.Raise 1000
    End If
    Set m_ObjUsuarioAplicacionPermisos = Constructor.getUsuarioAplicacionPermisos(m_ObjUsuarioActivo.CorreoUsuario, _
                        m_ObjAplicacionSeleccionada.IDAplicacion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjUsuarioAplicacionPermisos Is Nothing Then
        m_Error = "No registrado"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Pregunta = MsgBox("¿Desea eliminar los permisos de la aplicación seleccionada?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Permisos")
    If Pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjUsuarioAplicacionPermisos.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioActivo = Constructor.getUsuario(m_ObjUsuarioActivo.ID, , , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Private Sub Form_Open(Cancel As Integer)
   
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If m_ObjUsuarioActivo Is Nothing Then
        m_Error = "No hay un usuario activo"
        Err.Raise 1000
    End If
    If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado = EnumSino.Sí Then
        Me.AllowEdits = True
    Else
        Me.AllowEdits = False
    End If
    EstablecerCombo m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.ComandoEliminar.Enabled = False
    If Me.AllowEdits = True Then
        Me.ComandoAltaPerfil.Enabled = True
    Else
        Me.ComandoAltaPerfil.Enabled = False
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Public Function EstablecerLista(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_ID As Variant
    Dim m_IDApp As String
    Dim m_ObjCol As Scripting.Dictionary
    Dim m_objAplicacion As Aplicacion
    Dim m_ObjUsuarioAplicacionPermisos As UsuarioAplicacionPermisos
    Dim Admin As String
    Dim Calidad As String
    Dim CalidadAvisos As String
    Dim Secretaria As String
    Dim Tecnico As String
    Dim Economia As String
    Dim PartePermisos As String
    On Error GoTo errores
    
    p_Error = ""
    Set lst = Me.ListaUsuarios
    
    lst.RowSource = "IDAplicacion;Aplicación;Admin;Calidad;Economia;Técnico;Secretaría;Cal Avisos"
    
    Set m_ObjCol = m_ObjUsuarioActivo.ColAplicacionesPermisos
    p_Error = m_ObjUsuarioActivo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjCol Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_ObjCol
        Set m_ObjUsuarioAplicacionPermisos = m_ObjCol(m_ID)
        With m_ObjUsuarioAplicacionPermisos
            If .EsUsuarioAdministradorCalculado = EnumSino.Sí Then
                Admin = "Sí"
            Else
                Admin = "No"
            End If
            If .EsUsuarioTecnicoCalculado = EnumSino.Sí Then
                Tecnico = "Sí"
            Else
                Tecnico = "No"
            End If
            If .EsUsuarioCalidadAvisosCalculado = EnumSino.Sí Then
                CalidadAvisos = "Sí"
            Else
                CalidadAvisos = "No"
            End If
            If .EsUsuarioCalidadCalculado = EnumSino.Sí Then
                Calidad = "Sí"
            Else
                Calidad = "No"
            End If
            If .EsUsuarioEconomiaCalculado = EnumSino.Sí Then
                Economia = "Sí"
            Else
                Economia = "No"
            End If
            If .EsUsuarioSecretariaCalculado = EnumSino.Sí Then
                Secretaria = "Sí"
            Else
                Secretaria = "No"
            End If
            'lst.RowSource = "IDAplicacion;Aplicación;Admin;Calidad;Economia;Técnico;Secretaría;Cal Avisos"
            PartePermisos = Admin & ";" & Calidad & ";" & Economia & ";" & Tecnico & ";" & Secretaria & ";" & CalidadAvisos
            lst.AddItem .IDAplicacion & ";" & .Aplicacion.NombreAplicacion & ";" & PartePermisos
        End With
        Set m_ObjUsuarioAplicacionPermisos = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function



Public Function EstablecerCombo(Optional ByRef p_Error As String) As String
    
    Dim cmb As ComboBox
    Dim m_ID As Variant
    Dim m_objAplicacion As Aplicacion
    
    On Error GoTo errores
    
    Set cmb = Me.ComboAplicacion
    cmb.RowSource = ""
    If m_ObjEntorno.ColAplicaciones Is Nothing Then
        p_Error = m_ObjEntorno.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
   
    For Each m_ID In m_ObjEntorno.ColAplicaciones
        Set m_objAplicacion = m_ObjEntorno.ColAplicaciones(m_ID)
        If m_objAplicacion.ConIconoEnLanzaderaCalculado = EnumSino.Sí Then
            cmb.AddItem m_objAplicacion.IDAplicacion & ";" & m_objAplicacion.NombreAplicacion
            Set m_objAplicacion = Nothing
        End If
    Next
    Set m_objAplicacion = m_ObjEntorno.ColAplicaciones("12")
    cmb.AddItem m_objAplicacion.IDAplicacion & ";" & m_objAplicacion.NombreAplicacion
    Set m_objAplicacion = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerCombo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Sub ListaUsuarios_Click()
    
    Dim m_IDAplicacion As String
    
    On Error GoTo errores
    
    m_Error = ""
   
    Me.ComandoEliminar.Enabled = False
    m_IDAplicacion = Nz(Me.ListaUsuarios, "")
    
    If m_IDAplicacion = "" Then
        Set m_ObjAplicacionSeleccionada = Nothing
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    If m_ObjAplicacionSeleccionada Is Nothing Then
        Set m_ObjAplicacionSeleccionada = Constructor.getAplicacion(m_IDAplicacion, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_IDAplicacion <> m_ObjAplicacionSeleccionada.IDAplicacion Then
            Set m_ObjAplicacionSeleccionada = Constructor.getAplicacion(m_IDAplicacion, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If Me.AllowEdits = True Then
        Me.ComandoEliminar.Enabled = True
        Me.ComandoAltaPerfil.Enabled = True
    Else
        Me.ComandoEliminar.Enabled = False
        Me.ComandoAltaPerfil.Enabled = False
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaUsuarios_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    
End Sub