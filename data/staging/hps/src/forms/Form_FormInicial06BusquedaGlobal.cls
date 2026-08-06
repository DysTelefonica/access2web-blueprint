Attribute VB_Name = "Form_FormInicial06BusquedaGlobal"
Option Compare Database
Option Explicit
Private m_Error As String

Private m_ObjUsuarioSeleccionado As UsuarioHPS
Private m_ObjExpedienteSeleccionado As Expediente
Private ApartadoSeleccionado As String
Private IDSeleccionado As String
Public m_PalabraClaveAlInicio As String

Private Sub CmdSalir_Click()
    Forms("FormInicial00Principal").PalabraClave = Null
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoLimpiarEmpresaUsuario_Click()
    On Error Resume Next
    Me.PalabraClave = Null
    Me.ListaResultados.RowSource = "ID;Apartado;Nombre Campo;Valor"
    Me.ListaResultados.Requery
End Sub

Private Sub ComandoVerDatoSeleccionado_Click()
    
    On Error GoTo errores
    m_Error = ""
    IDSeleccionado = Nz(Me.ListaResultados.Column(0), "")
    ApartadoSeleccionado = Nz(Me.ListaResultados.Column(1), "")
    
    If ApartadoSeleccionado = "UsuarioS SICA" Then
        FiltrarUsuarioSICA IDSeleccionado, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf ApartadoSeleccionado = "Usuarios Históricos" Then
        FiltrarUsuarioHistorico IDSeleccionado, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf ApartadoSeleccionado = "Anexos" Then
        FiltrarAnexos IDSeleccionado, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf ApartadoSeleccionado = "Anexos SICA" Then
        FiltrarAnexosSICA IDSeleccionado, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf ApartadoSeleccionado = "Observaciones" Then
        FiltrarObservaciones IDSeleccionado, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf ApartadoSeleccionado = "Observaciones Históricas" Then
        FiltrarObservacionesHistoricas IDSeleccionado, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        FiltrarUsuario IDSeleccionado, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar ver el dato se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CopiarEstiloBoton Me.ComandoVerDatoSeleccionado, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    CopiarEstiloBoton Me.cmdSalir, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.ListaResultados.RowSource = "ID;Apartado;Nombre Campo;Valor"
    Me.ListaResultados.Requery
    Me.AllowEdits = True
     VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
     Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ListaResultados_Click()
    
    
    On Error GoTo errores
    m_Error = ""
    
    Me.ComandoVerDatoSeleccionado.Enabled = False
    IDSeleccionado = Nz(Me.ListaResultados.Column(0), "")
    ApartadoSeleccionado = Nz(Me.ListaResultados.Column(1), "")
    If IDSeleccionado = "" Then
        
        Exit Sub
    End If
    Me.ComandoVerDatoSeleccionado.Enabled = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al abrir el formulario se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Function FiltrarUsuario(p_ID As String, Optional ByRef p_Error As String) As String
    
    Dim frm As Form
    Dim rst As DAO.Recordset
    On Error GoTo errores
    
    If Not FormularioAbierto("FormInicial00Principal") Then
        DoCmd.OpenForm "FormInicial00Principal"
    End If
    Set frm = Forms("FormInicial00Principal")
    Form_FormInicial01Lateral.lblUsuariosHPS_Click
    
    Set frm = frm.Controls("SubFormCentral").Form
    Set frm = frm.Controls("SubFormResultados").Form
    Set rst = frm.RecordsetClone 'Clonamos el recordset
    rst.FindFirst "[ID] =" & p_ID   'Buscamos
    If Not rst.NoMatch Then 'Si  se encuentra
        frm.Bookmark = rst.Bookmark 'Nos vamos al registro encontrado
        'frm.IDExpediente.SetFocus 'Activamos el campo ID (ahora es cuando se produce el scroll)
        Set m_ObjUsuarioActivo = Constructor.getUsuarioHPS(p_ID, , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Form_FormInicial02ConsultasPrincipal.ComandoEditarUsuario_Click
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FiltrarUsuario ha producido el error: " & Err.Description
    End If
End Function
Public Function FiltrarObservaciones(p_IDObservacion As String, Optional ByRef p_Error As String) As String
    
    Dim m_ID As String
    Dim m_ObjUsuario As UsuarioHPS
    Dim m_ObjObServacion As Observacion
    On Error GoTo errores
    
    If p_IDObservacion = "" Then
        Err.Raise 1000
    End If
    Set m_ObjObServacion = Constructor.getObservacion(p_IDObservacion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjObServacion Is Nothing Then
        Exit Function
    End If
    If Not m_ObjObServacion.Usuario Is Nothing Then
        m_ID = m_ObjObServacion.Usuario.ID
        FiltrarObservaciones = FiltrarUsuario(m_ID, p_Error)
    End If
    
    
    
    
   
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FiltrarObservaciones ha producido el error: " & Err.Description
    End If
End Function
Public Function FiltrarObservacionesHistoricas(p_IDObservacion As String, Optional ByRef p_Error As String) As String
    
    Dim m_ID As String
    Dim m_ObjUsuarioHistorico As UsuarioHistorico
    Dim m_ObjObServacionHistorica As ObservacionHistorica
    On Error GoTo errores
    
    If p_IDObservacion = "" Then
        Err.Raise 1000
    End If
    Set m_ObjObServacionHistorica = Constructor.getObservacionHistorica(p_IDObservacion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjObServacionHistorica Is Nothing Then
        Exit Function
    End If
    If Not m_ObjObServacionHistorica.Usuario Is Nothing Then
        m_ID = m_ObjObServacionHistorica.Usuario.ID
        FiltrarObservacionesHistoricas = FiltrarUsuarioHistorico(m_ID, p_Error)
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FiltrarObservacionesHistoricas ha producido el error: " & Err.Description
    End If
End Function
Public Function FiltrarUsuarioHistorico(p_ID As String, Optional ByRef p_Error As String) As String
    
    Dim frm As Form
    Dim frmLateral As Form
    Dim rst As DAO.Recordset
    On Error GoTo errores
    
    If Not FormularioAbierto("FormInicial00Principal") Then
        DoCmd.OpenForm "FormInicial00Principal"
    End If
    Set frm = Forms("FormInicial00Principal")
    Set frmLateral = frm.Controls("SubFormLateral").Form
    If frm.Controls("SubFormCentral").SourceObject <> "FormInicial08UsuariosHistoricos" Then
        Form_FormInicial01Lateral.lblUsuariosHistoricos_Click
    End If
    Set frm = frm.Controls("SubFormCentral").Form
    Set frm = frm.Controls("SubFormResultados").Form
    Set rst = frm.RecordsetClone 'Clonamos el recordset
    rst.FindFirst "[ID] =" & p_ID   'Buscamos
    If Not rst.NoMatch Then 'Si  se encuentra
        frm.Bookmark = rst.Bookmark 'Nos vamos al registro encontrado
        frm.IDExpediente.SetFocus 'Activamos el campo ID (ahora es cuando se produce el scroll)
        Set m_ObjUsuarioHistoricoActivo = Constructor.getUsuarioHistorico(p_ID, , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Form_FormInicial08UsuariosHistoricos.ComandoVerUsuarioHistorico_Click
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FiltrarUsuarioHistorico ha producido el error: " & Err.Description
    End If
End Function
Public Function FiltrarAnexos(p_ID As String, Optional ByRef p_Error As String) As String
    
    
    On Error GoTo errores
    
    Set m_ObjAnexoUsuarioHPSActivo = getAnexoUsuarioHPS(p_ID, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjAnexoUsuarioHPSActivo Is Nothing Then
        Exit Function
    End If
    If Not fso.FileExists(m_ObjAnexoUsuarioHPSActivo.URLAnexo) Then
        p_Error = "No es alcanzable en estos momentos el anexo seleccionado" & vbNewLine & m_ObjAnexoUsuarioHPSActivo.URLAnexo
        Err.Raise 1000
    End If
    Ejecutar Me.hWnd, "open", m_ObjAnexoUsuarioHPSActivo.URLAnexo, "", "", 1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FiltrarUsuario ha producido el error: " & Err.Description
    End If
End Function
Public Function FiltrarAnexosSICA(p_ID As String, Optional ByRef p_Error As String) As String
    
    
    On Error GoTo errores
    
    Set m_ObjAnexoUsuarioSICAActivo = getAnexoUsuarioSICA(p_ID, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjAnexoUsuarioSICAActivo Is Nothing Then
        Exit Function
    End If
    If Not fso.FileExists(m_ObjAnexoUsuarioSICAActivo.URLAnexo) Then
        p_Error = "No es alcanzable en estos momentos el anexo SICA seleccionado" & vbNewLine & m_ObjAnexoUsuarioSICAActivo.URLAnexo
        Err.Raise 1000
    End If
    Ejecutar Me.hWnd, "open", m_ObjAnexoUsuarioSICAActivo.URLAnexo, "", "", 1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FiltrarAnexosSICA ha producido el error: " & Err.Description
    End If
End Function
Public Function FiltrarUsuarioSICA(p_ID As String, Optional ByRef p_Error As String) As String
    
    Dim frm As Form
    Dim frmLateral As Form
    Dim rst As DAO.Recordset
    On Error GoTo errores
    
    If Not FormularioAbierto("FormInicial00Principal") Then
        DoCmd.OpenForm "FormInicial00Principal"
    End If
    Set frm = Forms("FormInicial00Principal")
    Set frmLateral = frm.Controls("SubFormLateral").Form
    If frm.Controls("SubFormCentral").SourceObject <> "FormInicial07UsuariosSICA" Then
        Form_FormInicial01Lateral.lblUsuariosSICA_Click
    End If
    Set frm = frm.Controls("SubFormCentral").Form
    Set frm = frm.Controls("SubFormResultados").Form
    Set rst = frm.RecordsetClone 'Clonamos el recordset
    rst.FindFirst "[ID] ='" & p_ID & "'"   'Buscamos
    If Not rst.NoMatch Then 'Si  se encuentra
        frm.Bookmark = rst.Bookmark 'Nos vamos al registro encontrado
        frm.ID.SetFocus 'Activamos el campo ID (ahora es cuando se produce el scroll)
        Form_FormInicial07UsuariosSICA.ComandoEditarUsuarioSICA_Click
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FiltrarUsuarioSICA ha producido el error: " & Err.Description
    End If
End Function


Private Sub ListaResultados_DblClick(Cancel As Integer)
    On Error Resume Next
    Application.Echo False
    If ComandoVerDatoSeleccionado.Enabled Then
        ComandoVerDatoSeleccionado_Click
    End If
    Application.Echo True
End Sub

Public Sub PalabraClave_Exit(Cancel As Integer)
    Dim Palabra As String
    On Error GoTo errores

    Palabra = Nz(Me.PalabraClave, "")
    Me.AllowEdits = False
    If Palabra = "" Then
        RellenarListaConBusquedaGlobal Palabra, Me.ListaResultados, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Sub
    End If
    If Palabra <> m_PalabraClaveAlInicio Or Me.ListaResultados.ListCount = 1 Then
        VBA.DoEvents
        DoCmd.Hourglass True
        VBA.DoEvents
        Me.ComandoVerDatoSeleccionado.Enabled = False
        'Palabra = Nz(Me.ListaResultados, "")
        RellenarListaConBusquedaGlobal Palabra, Me.ListaResultados, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If

        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
    End If
    m_PalabraClaveAlInicio = Palabra
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al realizar la búsqueda se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub PalabraClave_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Palabra As String
    On Error GoTo errores
   
    If KeyCode = vbKeyReturn Then
        If Me.AllowEdits = False Then
            Exit Sub
        End If
        Me.ComandoLimpiarEmpresaUsuario.Enabled = True
        Me.ComandoLimpiarEmpresaUsuario.SetFocus
        Me.PalabraClave.SetFocus
        Palabra = Nz(Me.PalabraClave, "")
        Me.AllowEdits = False
        If Palabra = "" Then
            RellenarListaConBusquedaGlobal Palabra, Me.ListaResultados, m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Sub
        End If
        If Palabra <> m_PalabraClaveAlInicio Or Me.ListaResultados.ListCount = 1 Then
            VBA.DoEvents
            DoCmd.Hourglass True
            VBA.DoEvents
            Me.ComandoVerDatoSeleccionado.Enabled = False
            'Palabra = Nz(Me.ListaResultados, "")
            RellenarListaConBusquedaGlobal Palabra, Me.ListaResultados, m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
    
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
        End If
        m_PalabraClaveAlInicio = Palabra
        Me.AllowEdits = True
        
        KeyCode = 0
    End If
     Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al realizar la PalabraClave_KeyDown se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Function RellenarListaConBusquedaGlobal( _
                                                p_Palabra As String, _
                                                ByRef lst As ListBox, _
                                                Optional ByRef _
                                                p_Error As String _
                                                ) As String
    
    Dim Apartado As String
    Dim NombreCampo As String
    Dim ValorCampo As String
    Dim ID As String
    Dim ColResultado As Scripting.Dictionary
    Dim VarItem As Variant
    On Error GoTo errores
    
    
    
    lst.RowSource = "ID;Apartado;Nombre Campo;Valor"
    If p_Palabra = "" Then
        Exit Function
    End If
    
    'Usuarios
    'ColResultado.Add Nz(rcdDatos.Fields("ID").value, "") & "|" &  Apartado & "|" & m_NombreCampo & "|" & m_Valor
    'Expedientes
    'ColResultado.Add Nz(rcdDatos.Fields("Cod_Exp").value, "") & "|" & Apartado & "|" & m_NombreCampo & "|" & m_Valor
    
    Set ColResultado = Constructor.getColResultadoBusqueda(p_Palabra, p_Error)
    
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If ColResultado Is Nothing Then
        Exit Function
    End If
    For Each VarItem In ColResultado.Keys
        dato = Split(VarItem, "|")
        If UBound(dato) = 3 Then
            ID = dato(0)
            Apartado = dato(1)
            NombreCampo = dato(2)
            ValorCampo = dato(3)
            lst.AddItem Replace(ID, ";", ":") & ";" & Replace(Apartado, ";", ":") & ";" & Replace(NombreCampo, ";", ":") & ";" & _
                        Replace(ValorCampo, ";", ":")
        End If
    Next
    If lst.ListCount = 2 Then
        lst.Selected(1) = True
        ListaResultados_Click
    End If
    Set ColResultado = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarListaConBusquedaGlobal ha devuelto el error: " & Err.Description
    End If
End Function
