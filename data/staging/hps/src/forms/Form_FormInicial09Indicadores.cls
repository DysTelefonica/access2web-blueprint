Attribute VB_Name = "Form_FormInicial09Indicadores"
Option Compare Database
Option Explicit
Private WithEvents m_Frm As Form_FormUsuario07SICAEdicion
Private WithEvents m_FrmEdicionHPS As Form_FormUsuario02Edicion
Private WithEvents m_Indicadores As clsIndicadoresBus
Private m_UltimoTipoEmpresa As String
Private m_UltimoNombreEmpresa As String
Private m_Error As String

Private Sub cmdActualizarIndicadores_Click()
    Dim frmIndicadores As Form
    
    On Error GoTo errores
    m_Error = ""
    
    ' 1. Feedback visual inmediato
    DoCmd.Hourglass True
    SincronizarCachesLocalesFrontend p_Error:=m_Error
    ' 2. Ejecutar sincronización incremental rápida de cachés locales.
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' 3. Refrescar contadores/botones numéricos tras la sincronización.
    EstablecerIndicadores m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' 4. Refrescar la rejilla del subformulario actual para mostrar los cambios
    ' (Si un usuario ya no está caducado, desaparecerá de la lista aquí)
    If Nz(Me.SubFormIndicadores.SourceObject, "") <> "" Then
        Set frmIndicadores = Me.SubFormIndicadores.Form
        If Not frmIndicadores Is Nothing Then
            frmIndicadores.Requery
        End If
    End If
    
    DoCmd.Hourglass False
    MsgBox "Sincronización completada. Tablas locales actualizadas.", vbInformation, "Indicadores"
    
    Exit Sub

errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Error al pulsar actualizar: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    MsgBox m_Error, vbCritical, "Error"
End Sub

' --- IMPORTANTE: Añade esto al Form_Load para mantener la estética ---
' Private Sub Form_Load()
'     ... (código existente) ...
'     CopiarEstiloBoton Me.cmdActualizarIndicadores, m_Error
'     ...
' End Sub

Public Sub ComandoAPuntoDeCaducar_Click()
    
    On Error GoTo errores
    m_Error = ""
    m_UltimoTipoEmpresa = ""
    m_UltimoNombreEmpresa = ""
    Me.ComboNombreEmpresa = Null
    Me.ComboTipoEmpresaFiltro = Null
    EstablecerFuenteDatosIndicadores EnumTipoIndicador.HPSAPuntoDeCaducar, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.lblFiltroAplicado.Caption = m_ObjEntorno.colFiltroEnIndicadordes(Me.ComandoAPuntoDeCaducar.Name)
    HacerVisible EnumSiNo.No
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAPuntoDeCaducar_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Sub ComandoCaducados_Click()
    
    On Error GoTo errores
    m_Error = ""
    m_UltimoTipoEmpresa = ""
    m_UltimoNombreEmpresa = ""
    Me.ComboNombreEmpresa = Null
    Me.ComboTipoEmpresaFiltro = Null
    EstablecerFuenteDatosIndicadores EnumTipoIndicador.HPSACaducadas, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.lblFiltroAplicado.Caption = m_ObjEntorno.colFiltroEnIndicadordes(Me.ComandoCaducados.Name)
    HacerVisible EnumSiNo.No
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoCaducados_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoCopiaAPortaPapelesFiltrado_Click()
    
    
    Dim m_Texto As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    
    m_Texto = CopiaAPortapapelesFiltrado(m_Error)
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox(m_Texto, vbInformation, "Información")
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoCopiaAPortaPapelesFiltrado_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEditarUsuario_Click()
    
    
    On Error GoTo errores
    
    m_Error = ""
    EstablecerUsuarioActivoIndicador m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjUsuarioActivo Is Nothing And m_ObjUsuarioSICAActivo Is Nothing Then
        m_Error = "Se ha de seleccionar el usuario primero"
        Err.Raise 1000
    End If
    If Not m_ObjUsuarioActivo Is Nothing Then
        If FormularioAbierto("FormUsuario00Principal") Then
            DoCmd.Close acForm, "FormUsuario00Principal", acSaveNo
        End If
        DoCmd.OpenForm "FormUsuario00Principal"
    Else
        If FormularioAbierto("FormUsuario07SICAEdicion") Then
            DoCmd.Close acForm, "FormUsuario07SICAEdicion", acSaveNo
        End If
        DoCmd.OpenForm "FormUsuario07SICAEdicion"
        Set m_Frm = Forms("FormUsuario07SICAEdicion")
    End If
    
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El ComandoEditarUsuario_Click ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Sub ComandoEnSolicitud_Click()

    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_UltimoTipoEmpresa = ""
    m_UltimoNombreEmpresa = ""
    Me.ComboNombreEmpresa = Null
    Me.ComboTipoEmpresaFiltro = Null
    EstablecerFuenteDatosIndicadores EnumTipoIndicador.HPSASolicitandose, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.lblFiltroAplicado.Caption = m_ObjEntorno.colFiltroEnIndicadordes(Me.ComandoEnSolicitud.Name)
    HacerVisible EnumSiNo.No
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEnSolicitud_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoListaEnExcel_Click()
    
    Dim m_URLExcelGenerado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_URLExcelGenerado = ConsultaAExcel(Me.SubFormIndicadores.Form, , , EnumSiNo.Sí, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Ejecutar Me.hWnd, "open", m_URLExcelGenerado, "", "", 1
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método ComandoListaEnExcel_Click ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
 
Public Sub ComandoPendientesConvocatoria1_Click()
    
    Dim m_EnumTipoIndicador As EnumTipoIndicador
    Dim m_Texto As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_UltimoTipoEmpresa = ""
    m_UltimoNombreEmpresa = ""
    Me.ComboNombreEmpresa = Null
    Me.ComboTipoEmpresaFiltro = Null
    m_EnumTipoIndicador = EnumTipoIndicador.PtesPrimeraConvocatoriaCurso
    EstablecerFuenteDatosIndicadores m_EnumTipoIndicador, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Me.lblFiltroAplicado.Caption = m_ObjEntorno.colFiltroEnIndicadordes(Me.ComandoPendientesConvocatoria1.Name)
    HacerVisible EnumSiNo.No
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPendientesConvocatoria1_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Sub ComandoPendientesConvocatoria2_Click()
    
    Dim m_EnumTipoIndicador As EnumTipoIndicador
    Dim m_Texto As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_UltimoTipoEmpresa = ""
    m_UltimoNombreEmpresa = ""
    Me.ComboNombreEmpresa = Null
    Me.ComboTipoEmpresaFiltro = Null
    m_EnumTipoIndicador = EnumTipoIndicador.PtesSegundaConvocatoriaCurso
    EstablecerFuenteDatosIndicadores m_EnumTipoIndicador, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Me.lblFiltroAplicado.Caption = m_ObjEntorno.colFiltroEnIndicadordes(Me.ComandoPendientesConvocatoria2.Name)
    HacerVisible EnumSiNo.No
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents

    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPendientesConvocatoria2_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Sub ComandoPendientesCorreoJefeSeguridad_Click()
    
    Dim m_EnumTipoIndicador As EnumTipoIndicador
    Dim m_Texto As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_UltimoTipoEmpresa = ""
    m_UltimoNombreEmpresa = ""
    Me.ComboNombreEmpresa = Null
    Me.ComboTipoEmpresaFiltro = Null
    m_EnumTipoIndicador = EnumTipoIndicador.PtesEnvioCorreoJefeSeguridadCurso
    
    EstablecerFuenteDatosIndicadores m_EnumTipoIndicador, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Me.lblFiltroAplicado.Caption = m_ObjEntorno.colFiltroEnIndicadordes(Me.ComandoPendientesCorreoJefeSeguridad.Name)
    HacerVisible EnumSiNo.No
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPendientesCorreoJefeSeguridad_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Sub ComandoPendientesCursoHPS_Click()
    
    Dim m_EnumTipoIndicador As EnumTipoIndicador
    Dim m_Texto As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    m_UltimoTipoEmpresa = ""
    m_UltimoNombreEmpresa = ""
    Me.ComboNombreEmpresa = Null
    Me.ComboTipoEmpresaFiltro = Null
    m_EnumTipoIndicador = EnumTipoIndicador.PendientesCursoHPS
    EstablecerFuenteDatosIndicadores m_EnumTipoIndicador, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    HacerVisible EnumSiNo.Sí
    
    Me.lblFiltroAplicado.Caption = m_ObjEntorno.colFiltroEnIndicadordes(Me.ComandoPendientesCursoHPS.Name)
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPendientesCursoHPS_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Function CopiaAPortapapeles( _
                                        p_EnumTipoIndicador As EnumTipoIndicador, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim m_Texto As String
    Dim m_CadenaCorreos As String
    Dim m_Col As Scripting.Dictionary
    
    
    On Error GoTo errores
    If p_EnumTipoIndicador = EnumTipoIndicador.PendientesCursoHPS Then
        m_Texto = "Copiado al portapapeles los que están pendientes de curso"
        Set m_Col = m_ObjEntorno.ColUsuariosHPSRequierenCurso
    ElseIf p_EnumTipoIndicador = EnumTipoIndicador.PtesPrimeraConvocatoriaCurso Then
        m_Texto = "Copiado al portapapeles los que están pendientes de curso y aún no se ha enviado la primera convocatoria"
        Set m_Col = m_ObjEntorno.ColRequiereCursoFaltaPrimeraConvocatoria
    ElseIf p_EnumTipoIndicador = EnumTipoIndicador.PtesSegundaConvocatoriaCurso Then
        m_Texto = "Copiado al portapapeles los que están pendientes de curso y ha pasado al menos 15 días y no se ha enviado segunda convocatoria"
        Set m_Col = m_ObjEntorno.ColRequiereCursoFaltaSegundaConvocatoria
    ElseIf p_EnumTipoIndicador = EnumTipoIndicador.PtesEnvioCorreoJefeSeguridadCurso Then
        m_Texto = "Copiado al portapapeles los que están pendientes de curso y han pasado al menos 15 días de la segunda convocatoria aún no se ha enviado el Correo al Jefe de Seguridad"
        Set m_Col = m_ObjEntorno.ColRequiereCursoFaltaCorreoJefeSeguridad
    Else
        p_Error = "Tipo de indicador no contemplado para copia al portapapeles"
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        p_Error = "Nada que copiar al portapapeles"
        Err.Raise 1000
    End If
    m_CadenaCorreos = getCadenaCorreosUsuarios(m_Col, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    SetText m_CadenaCorreos
    CopiaAPortapapeles = m_Texto
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiaAPortapapeles ha devuelto el error: " & Err.Description
    End If
End Function

Private Function HacerVisible( _
                            Optional p_Visible As EnumSiNo, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim blnVisible As Boolean
    On Error GoTo errores
    If p_Visible = EnumSiNo.Sí Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    Me.lbl1.Visible = blnVisible
    Me.ComboTipoEmpresaFiltro.Visible = blnVisible
    Me.ComboNombreEmpresa.Visible = blnVisible
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HacerVisible ha devuelto el error: " & Err.Description
    End If
End Function
Private Function CopiaAPortapapelesFiltrado( _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim m_CadenaCorreos As String
    Dim frm As Form
    Dim rcdDatos As DAO.Recordset
    Dim m_UsuarioHPS As UsuarioHPS
    Dim m_ID As String
    
    On Error GoTo errores
    
    Set frm = Me.SubFormIndicadores.Form
    frm.Requery
    Set rcdDatos = frm.RecordsetClone 'Clonamos el recordset
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            p_Error = "Ningún elemento filtrado"
            Err.Raise 1000
        End If
        .MoveFirst
        Do While Not .EOF
            m_ID = .Fields("ID")
            Set m_UsuarioHPS = Constructor.getUsuarioHPS(m_ID, , p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Not m_UsuarioHPS Is Nothing Then
                If InStr(1, m_UsuarioHPS.Correo_e, "@") <> 0 Then
                    If m_CadenaCorreos = "" Then
                        m_CadenaCorreos = m_UsuarioHPS.Correo_e
                    Else
                        m_CadenaCorreos = m_CadenaCorreos & ";" & m_UsuarioHPS.Correo_e
                    End If
                    
                End If
            End If
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    SetText m_CadenaCorreos
    CopiaAPortapapelesFiltrado = "Copiados los correos de los usuarios filtrados al portapapeles "
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiaAPortapapelesFiltrado ha devuelto el error: " & Err.Description
    End If
End Function

Private Sub ComboNombreEmpresa_AfterUpdate()
    
    Dim m_NombreEmpresa As String
    Dim m_TipoEmpresa As String
    Dim m_SQLActual As String
    
    On Error GoTo errores
    m_Error = ""
    m_TipoEmpresa = Nz(Me.ComboTipoEmpresaFiltro, "")
    m_NombreEmpresa = Nz(Me.ComboNombreEmpresa.Column(1), "")
    
    If m_TipoEmpresa = "" Or m_NombreEmpresa = "" Then
        m_UltimoTipoEmpresa = ""
        m_UltimoNombreEmpresa = ""
        Exit Sub
    End If
    If Me.ComboTipoEmpresaFiltro = m_UltimoTipoEmpresa And Me.ComboNombreEmpresa = m_UltimoNombreEmpresa Then
        Exit Sub
    End If
    m_UltimoTipoEmpresa = m_TipoEmpresa
    m_UltimoNombreEmpresa = m_NombreEmpresa
    EstablecerFuenteDatosIndicadoresConEmpresas m_TipoEmpresa, m_NombreEmpresa, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método ComboNombreEmpresa_AfterUpdate ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComboTipoEmpresaFiltro_Change()
    
    On Error GoTo errores
    m_Error = ""
    If Nz(Me.ComboTipoEmpresaFiltro, "") = "" Then
        Me.ComboNombreEmpresa.RowSource = ""
        Exit Sub
    End If
    Me.ComboNombreEmpresa = Null
    Me.ComboNombreEmpresa.RowSource = ""
    If Me.ComboTipoEmpresaFiltro = "Empresa Usuario" Or Me.ComboNombreEmpresa = "Jurídica Contratación" Then
        RellenarComboEmpresas Me.ComboNombreEmpresa, EnumSiNo.No, EnumSiNo.No, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        RellenarComboEmpresas Me.ComboNombreEmpresa, EnumSiNo.Sí, EnumSiNo.No, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método ComboTipoEmpresaFiltro_Change ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()


    On Error GoTo errores
    m_Error = ""

    CopiarDatosAIndicadores m_Error, p_CacheAcabDeRegenerar:=True
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    EstablecerIndicadores m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    Me.ComboTipoEmpresaFiltro.RowSource = "Empresa Usuario;Empresa Tramitadora;Juridica Contratación"
    Me.ComboNombreEmpresa.RowSource = ""
    Me.ComandoEditarUsuario.Enabled = False
    CopiarEstiloBoton Me.ComandoEditarUsuario, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CopiarEstiloBoton Me.cmdActualizarIndicadores
    Set m_Indicadores = modIndicadores.m_Indicadores
    If FormularioAbierto("FormUsuario02Edicion") Then
        Set m_FrmEdicionHPS = Forms("FormUsuario02Edicion")
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Form_Load ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")

End Sub

Private Function EstablecerUsuarioActivoIndicador( _
                                                    Optional p_Error As String _
                                                    ) As String
    Dim m_ID As String
    Dim frm As Form
    
    On Error GoTo errores
    Set m_ObjUsuarioActivo = Nothing
    Set m_ObjObservacionActiva = Nothing
    Set m_ObjAnexoUsuarioHPSActivo = Nothing
   
    Set m_ObjHPSNACActiva = Nothing
    Set m_ObjHPSOTANActiva = Nothing
    Set m_ObjHPSESAActiva = Nothing
    Set m_ObjHPSUEActiva = Nothing
    Set m_ObjUsuarioSICAActivo = Nothing
    If Me.SubFormIndicadores.SourceObject = "FormInicial09IndicadoresUsuariosSICADatos" Then
        Set frm = Me.SubFormIndicadores.Form
        m_ID = Nz(frm.ID, "")
        If m_ID = "" Then
            Exit Function
        End If
        Set m_ObjUsuarioSICAActivo = getUsuarioSICA(m_ID, , , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Set frm = Me.SubFormIndicadores.Form
        m_ID = Nz(frm.ID, "")
        If m_ID = "" Then
            Exit Function
        End If
        Set m_ObjUsuarioActivo = getUsuarioHPS(m_ID, , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerUsuarioActivoIndicador ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Private Sub m_Frm_EdicionRegistro()

    On Error GoTo errores
    m_Error = ""

    Me.SubFormIndicadores.SetFocus
    EstablecerIndicadores m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.SubFormIndicadores.SourceObject = ""

    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método m_Frm_EdicionRegistro ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_FrmEdicionHPS_EdicionRegistro()

    On Error GoTo errores
    m_Error = ""

    Me.SubFormIndicadores.SetFocus
    EstablecerIndicadores m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.SubFormIndicadores.SourceObject = ""

    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método m_FrmEdicionHPS_EdicionRegistro ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Activate()
    On Error GoTo errores
    m_Error = ""
    EstablecerIndicadores m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Form_Activate ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_Indicadores_LifecycleMoveCompleted(p_Direction As String, p_UserId As String)
    On Error Resume Next
    m_Error = ""
    EstablecerIndicadores m_Error
    If Nz(Me.SubFormIndicadores.SourceObject, "") <> "" Then
        Me.SubFormIndicadores.Form.Requery
    End If
    On Error GoTo 0
End Sub
