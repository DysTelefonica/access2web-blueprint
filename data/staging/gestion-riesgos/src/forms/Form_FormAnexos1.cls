Attribute VB_Name = "Form_FormAnexos1"
Option Compare Database
Option Explicit
Private m_objEdicion As Edicion
Private m_ObjProyecto As Proyecto
Private m_ObjRiesgo As riesgo

Private m_Error As String

Private blnPermitidoEscribir As Boolean
Private m_ObjAnexoSeleccionado As Anexo
Private m_ParaProyecto As EnumSiNo
Private m_ParaEdicion As EnumSiNo
Private m_ParaRiesgo As EnumSiNo
Private m_EnUTE As String
Public Event AnexoAñadido(ByVal p_ObjAnexo As Anexo)
Public Event AnexoEliminado(ByVal IDAnexo As String)
Private m_ObjAnexo As Anexo
Private m_ObjDELAnexo As Object
Private m_ObjColAnexos As Scripting.Dictionary
Dim m_URLAnexoSeleccionado As String
Const ALTO_FORM_CONUTE As Long = 9145
Const ALTO_FORM_SINUTE As Long = 7836
Const SUPERIOR_ANEXO_CONUTE As Long = 1679
Const SUPERIOR_ANEXO_SINUTE As Long = 1049
Const ALTO_CUADRO_PPAL_CONUTE As Long = 7753
Const ALTO_CUADRO_PPAL_SINUTE As Long = 6464
Const SUPERIOR_LISTA_CONUTE As Long = 2384
Const SUPERIOR_LISTA_SINUTE As Long = 959
Const SUPERIOR_TITULO_CONUTE As Long = 7709
Const SUPERIOR_TITULO_SINUTE As Long = 6284
Const SUPERIOR_CUADRO2_CONUTE As Long = 8459
Const SUPERIOR_CUADRO2_SINUTE As Long = 7150
Const SUPERIOR_SALIR_CONUTE As Long = 8518
Const SUPERIOR_SALIR_SINUTE As Long = 7228


Const SUPERIOR_EVIDENCIA_UTE = 1049
Const SUPERIOR_TITULO_ANEXO = 1679
Const SUPERIOR_LISTA = 2384
Const SUPERIOR_TITULO = 7709
Const SUPERIOR_CUADRO2 = 8459
Const SUPERIOR_SALIR = 8518
Const ALTO_FONDO = 7753



Public Function EstablecerLista(Optional ByRef p_Error As String) As String
    
    
    Dim lst As ListBox
    Dim m_IDAnexo As Variant
    Dim m_ObjAnexo As Anexo
    Dim m_ObjProyecto As Proyecto
    Dim m_ObjEdicionDeAnexo As Edicion
    Dim m_EdicionDeAnexo As String
    Dim m_FechaAnexo As String
    Dim m_Tipo As String
    Dim m_CodRiesgo As String
    Dim m_EdicionAnexado As String
    
    On Error GoTo errores
    
    Set lst = Me.ListaDocumentos
    Me.Titulo = ""
    lst.RowSource = "Tipo;IDAnexo;Título;Riesgo;Edic_Anexado;Fecha"
    Me.ComandoEliminarAnexo.Enabled = False
    Me.ImagenAnexo.Visible = False
    If m_ParaProyecto = EnumSiNo.Sí Then
        Set m_ObjColAnexos = m_ObjDELAnexo.ColAnexosTotales
    Else
        Set m_ObjColAnexos = m_ObjDELAnexo.ColAnexos
    End If
    
    p_Error = m_ObjDELAnexo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjColAnexos Is Nothing Then
        Exit Function
    End If
    For Each m_IDAnexo In m_ObjColAnexos.keys
        If Nz(m_IDAnexo, "") <> "" Then
            'If CStr(m_IDAnexo) = "48" Then Stop
            Set m_ObjAnexo = m_ObjColAnexos(m_IDAnexo)
            With m_ObjAnexo
                If .IDEdicion <> "" Then
                    Set m_ObjEdicionDeAnexo = .Edicion
                    p_Error = m_ObjEdicionDeAnexo.Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If Not m_ObjEdicionDeAnexo Is Nothing Then
                        m_EdicionDeAnexo = m_ObjEdicionDeAnexo.Edicion
                    Else
                        m_EdicionDeAnexo = ""
                    End If
                    Set m_ObjEdicionDeAnexo = Nothing
                Else
                    m_EdicionDeAnexo = ""
                End If
                m_FechaAnexo = .FechaAnexo
                If IsDate(m_FechaAnexo) Then
                    m_FechaAnexo = Format(m_FechaAnexo, "dd/mm/yyyy")
                Else
                    m_FechaAnexo = ""
                End If
                m_Tipo = .Tipo
                If m_Tipo = "R" Then
                    m_CodRiesgo = m_ObjAnexo.riesgo.CodigoRiesgo
                    m_EdicionDeAnexo = m_ObjAnexo.riesgo.EdicionEnLaQueNace.Edicion
                Else
                    m_CodRiesgo = ""
                End If
                lst.AddItem m_Tipo & ";" & .IDAnexo & ";" & .Titulo & ";" & m_CodRiesgo & ";" & m_EdicionDeAnexo & ";" & m_FechaAnexo
            End With
            Set m_ObjAnexo = Nothing
        End If
    Next
    lst.Requery
    ListaDocumentos_Click
    Set lst = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Sub ComandoAnexar_Click()
    
    Dim m_URL As String
    Dim m_EvidenciaUTE As String
    Dim m_Titulo As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set m_ObjAnexo = New Anexo
    If m_ParaEdicion = EnumSiNo.Sí Then
        If m_objEdicion Is Nothing Then
            m_Error = "No se conoce la edición"
            Err.Raise 1000
        End If
        If Me.EvidenciaUTE = "Sí" Then
            
            m_Titulo = "ED " & m_objEdicion.Edicion & " " & "Acuerdo tratamiento Riesgos con miembros de UTE"
            Me.TituloDoc = m_Titulo
            Me.TituloDoc.Enabled = False
        Else
            If Nz(Me.TituloDoc, "") = "" Then
                m_Error = "Escriba primero una pequeña descripción del anexo"
                Err.Raise 1000
            End If
        End If
    Else
        If Nz(Me.TituloDoc, "") = "" Then
            m_Error = "Escriba primero una pequeña descripción del anexo"
            Err.Raise 1000
        End If
    End If
    m_Titulo = Nz(Me.TituloDoc, "")
    m_URL = Seleccionar(True, "Seleccione el archivo a anexar", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_URL = "" Then
        Set m_ObjAnexo = Nothing
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    
    
    With m_ObjAnexo
        .EvidenciaUTE = Nz(Me.EvidenciaUTE, "")
        .Titulo = m_Titulo
        If m_ParaEdicion = EnumSiNo.Sí Then
            If Not m_objEdicion Is Nothing Then
                .IDEdicion = m_objEdicion.IDEdicion
            End If
            If .EvidenciaUTE = "Sí" Then
                If Not m_objEdicion.AnexoEvidenciaUTE Is Nothing Then
                    VBA.DoEvents
                    DoCmd.Hourglass False
                    VBA.DoEvents
                    pregunta = MsgBox("Ya existe un anexo para esta edición y para evicencia de UTE" & vbNewLine & _
                                    "¿Desea sustituirlo?", vbExclamation + vbYesNo + vbDefaultButton2, "Anexo existente")
                    If pregunta <> vbYes Then
                        Exit Sub
                    End If
                End If
            End If
        End If
        If m_ParaProyecto = EnumSiNo.Sí Then
            If Not m_ObjProyecto Is Nothing Then
                .IDProyecto = m_ObjProyecto.IDProyecto
            End If
        End If
        If m_ParaRiesgo = EnumSiNo.Sí Then
            If Not m_ObjRiesgo Is Nothing Then
                .IDRiesgo = m_ObjRiesgo.IDRiesgo
            End If
        End If
        .Registrar m_ObjDELAnexo, m_URL, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    If Me.EvidenciaUTE = "Sí" Then
        If Not m_objEdicion.AnexoEvidenciaUTE Is Nothing Then
            m_objEdicion.AnexoEvidenciaUTE.EliminarAnexo
        End If
    End If
    RaiseEvent AnexoAñadido(m_ObjAnexo)
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.SetFocus
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAnexar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    
    
End Sub
Private Sub CmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    AbrirAyuda m_Error
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
        m_Error = "Al ComandoAyuda_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoEliminarAnexo_Click()
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    On Error GoTo errores
    If m_ObjAnexoSeleccionado Is Nothing Then
        Set m_ObjAnexoSeleccionado = Constructor.getAnexo(Nz(Me.ListaDocumentos.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjAnexoSeleccionado Is Nothing Then
            m_Error = "Se ha de seleccionar un documento de la lista"
            Err.Raise 1000
        End If
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea eliminar el anexo seleccionado?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar anexo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_ObjAnexoSeleccionado.EliminarAnexo m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RaiseEvent AnexoEliminado(m_ObjAnexoSeleccionado.IDAnexo)
    Set m_ObjAnexoSeleccionado = Nothing
    Me.SetFocus
    If m_ParaProyecto = EnumSiNo.Sí Then
        If m_ObjProyecto Is Nothing Then
            m_Error = "No se conoce la gestión de riesgos activa"
            Err.Raise 1000
        End If
        Set m_ObjDELAnexo = Constructor.getProyecto(m_ObjProyecto.IDProyecto, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    ElseIf m_ParaEdicion = EnumSiNo.Sí Then
        If m_objEdicion Is Nothing Then
            m_Error = "No se conoce la edición activa"
            Err.Raise 1000
        End If
        Set m_ObjDELAnexo = m_objEdicion
    ElseIf m_ParaRiesgo = EnumSiNo.Sí Then
        If m_ObjRiesgoActivo Is Nothing Then
            m_Error = "No se conoce el riesgo activo"
            Err.Raise 1000
        End If
        Set m_ObjDELAnexo = m_ObjRiesgoActivo
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
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarAnexo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
Private Function AdaptarTamañoFormulario( _
                                            Optional ByRef p_Error As String _
                                            ) As String

    On Error GoTo errores
    
    If m_EnUTE = "Sí" Then
        Me.lblEvidenciaUTE.Visible = True
        Me.EvidenciaUTE.Visible = True
    Else
        Me.lblEvidenciaUTE.Visible = False
        Me.EvidenciaUTE.Visible = False
    End If

    AdaptarTamañoFormulario = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AdaptarTamañoFormulario ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function



Private Sub Form_Open(Cancel As Integer)
    
    Dim m_Argumento As String
    Dim m_Titulo As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me
    Me.AllowEdits = True
    Me.EvidenciaUTE.Enabled = False
    Me.ComandoEliminarAnexo.Enabled = False
    Me.ComandoAnexar.Enabled = False
    
    'm_Argumento="Riesgo"
    'm_Argumento="Edición"
    'm_Argumento="Proyecto"
    m_Error = ""
    m_Argumento = Nz(Me.openArgs(), "")
    Me.ImagenAnexo.Visible = False
    If m_Argumento = "Proyecto" Then
        If m_ObjProyectoActivo Is Nothing Then
            m_Error = "No se conoce el proyecto activo"
            Err.Raise 1000
        End If
        Set m_ObjProyecto = m_ObjProyectoActivo
        Set m_ObjDELAnexo = m_ObjProyecto
        m_Titulo = "ANEXOS PARA: " & UCase(m_ObjDELAnexo.Proyecto)
        m_ParaProyecto = EnumSiNo.Sí
        m_ParaEdicion = EnumSiNo.No
        m_ParaRiesgo = EnumSiNo.No
    ElseIf m_Argumento = "Edición" Then
        If m_ObjEdicionActiva Is Nothing Then
            m_Error = "No se conoce la edición activa"
            Err.Raise 1000
        End If
        Set m_objEdicion = m_ObjEdicionActiva
        Set m_ObjProyecto = m_objEdicion.Proyecto
        Set m_ObjDELAnexo = m_objEdicion
        If m_ObjProyecto.EnUTE = "Sí" Then
            m_EnUTE = "Sí"
            Me.EvidenciaUTE.Enabled = True
        Else
            m_EnUTE = "No"
        End If
        m_Titulo = "ANEXOS PARA LA EDICIÓN : " & m_ObjDELAnexo.Edicion
        m_ParaProyecto = EnumSiNo.No
        m_ParaEdicion = EnumSiNo.Sí
        m_ParaRiesgo = EnumSiNo.No
    ElseIf m_Argumento = "Riesgo" Then
        If m_ObjRiesgoActivo Is Nothing Then
            m_Error = "No se conoce el riesgo activo"
            Err.Raise 1000
        End If
        Set m_ObjRiesgo = m_ObjRiesgoActivo
        Set m_objEdicion = m_ObjRiesgo.Edicion
        Set m_ObjProyecto = m_objEdicion.Proyecto
        Set m_ObjDELAnexo = m_ObjRiesgoActivo
        m_Titulo = "ANEXOS PARA EL RIESGO : " & UCase(m_ObjDELAnexo.CodigoRiesgo)
        m_ParaProyecto = EnumSiNo.No
        m_ParaEdicion = EnumSiNo.No
        m_ParaRiesgo = EnumSiNo.Sí
    Else
        m_Error = "Se ha abierto el formulario con parámteros insuficientes"
        Err.Raise 1000
    End If
    If m_ObjDELAnexo.EsActivo = EnumSiNo.Sí Then
        m_Error = m_ObjDELAnexo.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        blnPermitidoEscribir = True
        Me.ComandoAnexar.Enabled = True
    Else
        m_Error = m_ObjDELAnexo.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        blnPermitidoEscribir = False
        Me.ComandoAnexar.Enabled = False
    End If
    AdaptarTamañoFormulario m_Error
    If m_Error <> "" Then
        Err.Raise 10000
    End If
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Me.lblTitulo1.Caption = m_Titulo
    
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ImagenAnexo_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_URLAnexoSeleccionado = "" Then
        m_Error = "Seleccione un elemento de la lista"
        Err.Raise 1000
    End If
    AbrirEnLocal m_URLAnexoSeleccionado, Me.hWnd, m_Error
    If m_Error <> "" Then
        Err.Raise 100
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ImagenAnexo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ListaDocumentos_Click()
     
     
    Dim m_IDAnexoSeleccionado As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.ComandoEliminarAnexo.Enabled = False
    Me.ImagenAnexo.Visible = False
    
    m_IDAnexoSeleccionado = Nz(Me.ListaDocumentos.Column(1), "")
    If m_IDAnexoSeleccionado = "" Then
        m_URLAnexoSeleccionado = ""
        Set m_ObjAnexoSeleccionado = Nothing
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_ObjAnexoSeleccionado Is Nothing Then
        Set m_ObjAnexoSeleccionado = Constructor.getAnexo(m_IDAnexoSeleccionado, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_IDAnexoSeleccionado <> m_ObjAnexoSeleccionado.IDAnexo Then
            Set m_ObjAnexoSeleccionado = Constructor.getAnexo(m_IDAnexoSeleccionado, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_ObjAnexoSeleccionado Is Nothing Then
        Me.Titulo = ""
        Me.ListaDocumentos.RemoveItem (Me.ListaDocumentos.Column(0))
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.Titulo = m_ObjAnexoSeleccionado.Titulo
    m_URLAnexoSeleccionado = m_ObjAnexoSeleccionado.URLAnexo
    m_Error = m_ObjAnexoSeleccionado.Error
    If m_Error <> "" Then
        m_URLAnexoSeleccionado = ""
        Err.Raise 1000
    End If
    If m_URLAnexoSeleccionado <> "" Then
        Me.ImagenAnexo.Visible = True
    Else
        Me.ImagenAnexo.Visible = False
    End If
    If blnPermitidoEscribir = True Then
        Me.ComandoEliminarAnexo.Enabled = True
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ListaDocumentos_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub





Private Sub ListaDocumentos_DblClick(Cancel As Integer)
    On Error Resume Next
    Application.Echo False
    If Me.ImagenAnexo.Visible = True Then
        ImagenAnexo_Click
    End If
    Application.Echo True
End Sub
