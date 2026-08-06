Attribute VB_Name = "FUNCIONES UTILES"
Option Compare Database
Option Explicit
Private s_contadorPasos As Long
Public Function ActualizaDatosLocalEmpresas( _
                                            p_NombreEmpresaCambiado As String, _
                                            p_NombreEmpresaInicial As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    If p_NombreEmpresaCambiado = "" Or p_NombreEmpresaInicial = "" Then
        Exit Function
    End If
    
    m_SQL = "UPDATE TbUsuariosEntidades SET EmpresaUsuario = '" & p_NombreEmpresaCambiado & "' " & _
            "WHERE (((TbUsuariosEntidades.EmpresaUsuario)='" & p_NombreEmpresaInicial & "'));"
    
    getdbHPS().Execute m_SQL
    
    m_SQL = "UPDATE TbUsuariosEntidades SET EmpresaTramitadora = '" & p_NombreEmpresaCambiado & "' " & _
            "WHERE (((TbUsuariosEntidades.EmpresaTramitadora)='" & p_NombreEmpresaInicial & "'));"
    
    getdbHPS().Execute m_SQL
    m_SQL = "UPDATE TbUsuariosEntidades SET JuridicaContrato = '" & p_NombreEmpresaCambiado & "' " & _
            "WHERE (((TbUsuariosEntidades.JuridicaContrato)='" & p_NombreEmpresaInicial & "'));"
    
    getdbHPS().Execute m_SQL
   
     
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizaDatosLocalEmpresas ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function
Public Function PintarIndicadores( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    On Error GoTo errores
    
    ' 1. SINCRONIZACIÓN (CRÍTICO)
    ' Como hemos optimizado el sistema para usar caché en la lectura,
    ' ahora debemos obligar a refrescar esa caché cuando modificamos datos.
    ' Esto borra las variables internas del Entorno, consulta la BD y actualiza los Longs.
    If Not m_ObjEntorno Is Nothing Then
        m_ObjEntorno.CalcularTotalesPendientes m_ObjUsuarioConectado, EsTecnico
    End If
    
    ' 2. ACTUALIZACIÓN DE INTERFAZ
    ' Ahora que el Entorno tiene los datos frescos, pedimos a los formularios que se repinten.
    
    ' A) Formulario de Tareas (Lista detallada)
    If FormularioAbierto("FormTareasTramitadorPendientes") Then
        ' Este formulario suele hacer sus propias consultas para llenar la lista,
        ' pero llamamos a EstablecerDatos para asegurar coherencia total.
        Form_FormTareasTramitadorPendientes.EstablecerDatos
    End If
    
    ' B) Formulario Principal (Menú / Botones con contadores)
    If FormularioAbierto("Form0BDOpciones") Then
        ' Este método ahora es muy rápido, solo lee las variables Long del Entorno
        ' que acabamos de actualizar en el paso 1.
        Form_Form0BDOpciones.EstablecerIndicadores
    End If
       
    PintarIndicadores = "OK"
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PintarIndicadores ha devuelto el error: " & vbNewLine & Err.Description
    End If
    ' Opcional: Registrar el error pero no detener la ejecución si es solo refresco visual
    Debug.Print "Error en PintarIndicadores: " & p_Error
End Function

Public Function RegistrarActualizacion( _
                                        p_Solicitud As solicitud, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    
    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    Dim m_UltOp As UltimoCambioOperaciones
    Dim m_EstadoCalculadoTexto As String
    On Error GoTo errores
    If p_Solicitud Is Nothing Then
        Exit Function
    End If
    If p_Solicitud.IDSolicitud = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbSolicitudes " & _
            "WHERE IDSolicitud=" & p_Solicitud.IDSolicitud & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            If Nz(.Fields("UsuarioCreacion"), "") = "" Then
                .Fields("UsuarioCreacion") = m_ObjUsuarioConectado.UsuarioRed
                .Fields("FechaCreacion") = Now()
            Else
                .Fields("UsuarioUltimoCambio") = m_ObjUsuarioConectado.UsuarioRed
                .Fields("FechaUltimoCambio") = Now()
            End If
            m_EstadoCalculadoTexto = p_Solicitud.EstadoCalculadoTexto
            If p_Solicitud.Estado <> m_EstadoCalculadoTexto Then
                p_Solicitud.Estado = m_EstadoCalculadoTexto
                .Fields("Estado") = p_Solicitud.Estado
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    ActualizarListaSolicitudes p_Solicitud
    Set m_UltOp = New UltimoCambioOperaciones
    With m_UltOp
        .Registrar p_Solicitud
    End With
    If EsAdministrador = EnumSiNo.Sí Then
        Set m_ObjEntorno.UltimoCambio = Nothing
        If FormularioAbierto("FormSolicitudesGestion") Then
            Forms("FormSolicitudesGestion").lblUltimaModificacion.Visible = True
            Forms("FormSolicitudesGestion").lblUltimaModificacion.Caption = m_ObjEntorno.UltimoCambio.texto
        End If
        If FormularioAbierto("FormTareasTramitadorPendientes") Then
            Forms("FormTareasTramitadorPendientes").lblUltimaModificacion.Visible = True
            Forms("FormTareasTramitadorPendientes").lblUltimaModificacion.Caption = m_ObjEntorno.UltimoCambio.texto
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistrarActualizacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
Public Function ActualizarListaSolicitudes( _
                                            p_Solicitud As solicitud, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim i As Long
    Dim m_ID As String
    Dim m_LineaInicial As String
    Dim m_LineaFinal As String
    Dim m_queryInicial As String
    Dim m_queryFinal As String
    Dim lst As ListBox
    
    On Error GoTo errores
    If p_Solicitud Is Nothing Then
        Exit Function
    End If
    If Not FormularioAbierto("FormSolicitudesGestion") Then
        Exit Function
    End If
    Set lst = Forms("FormSolicitudesGestion").ListaFiltrados
    If lst.ListCount = 1 Then
        Exit Function
    End If
    Set p_Solicitud = constructor.getSolicitud(p_IDSolicitud:=p_Solicitud.IDSolicitud, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_Solicitud Is Nothing Then
        Exit Function
    End If
    m_queryInicial = lst.RowSource
    For i = 1 To lst.ListCount - 1
        m_ID = lst.Column(0, i)
        If m_ID = p_Solicitud.IDSolicitud Then
            With lst
                m_LineaInicial = lst.Column(0, i) & ";" & lst.Column(1, i) & ";" & lst.Column(2, i) & ";" & _
                                lst.Column(3, i) & ";" & lst.Column(4, i)
            End With
            With p_Solicitud
                m_LineaFinal = m_ID & ";" & .TIPO & ";" & .DNI & ";" & .NombreCompletoCapitalizado & ";" & .EstadoTitulo
            
            End With
            m_queryFinal = Replace(m_queryInicial, m_LineaInicial, m_LineaFinal)
            lst.RowSource = m_queryFinal
            lst.Requery
            Exit For
        End If
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizarListaSolicitudes ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
Public Function getSolicitudFechasPrevistas(p_SolcitudFechas As SolicitudFechas, Optional ByRef p_Error As String) As SolicitudFechas
    
    Dim m_SolicitudFechas As SolicitudFechas
    On Error GoTo errores
    
    Set m_SolicitudFechas = New SolicitudFechas
    With m_SolicitudFechas
        .IDSolicitud = p_SolcitudFechas.IDSolicitud
        If .solicitud.CorreosAutomaticos = "No" Then
            Exit Function
        End If
        If m_FaltaAlgunParametroEnConfiguracion = True Then
            Exit Function
        End If
        If .FechaEnvioExcel = "" Then
            Exit Function
        End If
        
        If .FechaTramitacionAltaMarga = "" Then
            If .FechaCorreoRecordatorioExcel1 = "" Then
                .FechaPrevistaCorreoRecordatorioExcel1 = DateAdd("d", CDbl(m_ObjEntorno.DiasParaRecordatorioExcel1), CDate(.FechaEnvioExcel))
            Else
                If .FechaCorreoRecordatorioExcel2 = "" Then
                    .FechaPrevistaCorreoRecordatorioExcel2 = DateAdd("d", CDbl(m_ObjEntorno.DiasParaRecordatorioExcel2), CDate(.FechaEnvioExcel))
                End If
            End If
            If m_ObjEntorno.AutocancelacionPreMARGA = "Sí" Then
                .FechaPrevistaAutocancelacionPreMarga = DateAdd("d", CDbl(m_ObjEntorno.DiasCancelacionPreMARGA), CDate(.FechaEnvioExcel))
            End If
        Else
            'relleno alta tramitación marga
            If .FechaCorreoRecordatorioRellenoMarga1 = "" Then
                .FechaPrevistaCorreoRecordatorioRellenoMarga1 = DateAdd("d", CDbl(m_ObjEntorno.DiasParaRecordatorioRellenoMarga1), CDate(.FechaTramitacionAltaMarga))
            Else
                If .FechaCorreoRecordatorioRellenoMarga2 = "" Then
                    .FechaPrevistaCorreoRecordatorioRellenoMarga2 = DateAdd("d", CDbl(m_ObjEntorno.DiasParaRecordatorioRellenoMarga2), CDate(.FechaEnvioExcel))
                End If
            End If
            If m_ObjEntorno.AutocancelacionMARGA = "Sí" Then
                .FechaPrevistaAutocancelacionMarga = DateAdd("d", CDbl(m_ObjEntorno.DiasCancelacionMARGA), CDate(.FechaTramitacionAltaMarga))
            End If
        End If
    End With
    Set getSolicitudFechasPrevistas = m_SolicitudFechas
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSolicitudFechasPrevistas ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function



Sub CopiarAlPortapapeles(texto As String)
    Dim hGlobalMemory As LongPtr
    Dim lpGlobalMemory As LongPtr
    Dim hwnd As LongPtr
    Dim lngReturnValue As LongPtr

    ' Abrir el portapapeles
    lngReturnValue = OpenClipboard(hwnd)
    If lngReturnValue = 0 Then Exit Sub

    ' Vaciar el portapapeles
    lngReturnValue = EmptyClipboard
    If lngReturnValue = 0 Then
        CloseClipboard
        Exit Sub
    End If

    ' Asignar memoria global para el texto
    hGlobalMemory = GlobalAlloc(GMEM_MOVEABLE, Len(texto) + 1)
    If hGlobalMemory = 0 Then
        CloseClipboard
        Exit Sub
    End If

    ' Bloquear la memoria global
    lpGlobalMemory = GlobalLock(hGlobalMemory)
    If lpGlobalMemory = 0 Then
        CloseClipboard
        Exit Sub
    End If

    ' Copiar el texto a la memoria global
    lstrcpy lpGlobalMemory, texto

    ' Desbloquear la memoria global
    GlobalUnlock hGlobalMemory

    ' Establecer los datos del portapapeles
    lngReturnValue = SetClipboardData(CF_TEXT, hGlobalMemory)
    If lngReturnValue = 0 Then
        CloseClipboard
        Exit Sub
    End If

    ' Cerrar el portapapeles
    CloseClipboard
End Sub


Function ConvertirImagenABase64( _
                                p_URLImagen As String, _
                                Optional p_EnPortaPapeles As EnumSiNo = EnumSiNo.Sí _
                                ) As String
    Dim FileStream As Object
    Dim BinaryData
    Dim Base64Data As String
    
    
    ' Lee el archivo como binario
    Set FileStream = CreateObject("ADODB.Stream")
    
    FileStream.Type = 1 ' Tipo binario
    FileStream.Open
    FileStream.LoadFromFile p_URLImagen
    BinaryData = FileStream.Read
    FileStream.Close

    ' Convierte a base64
    Set FileStream = CreateObject("MSXML2.DOMDocument").createElement("Base64Data")
    FileStream.DataType = "bin.base64"
    FileStream.nodeTypedValue = BinaryData
    Base64Data = FileStream.Text
    Set FileStream = Nothing
    If p_EnPortaPapeles = EnumSiNo.Sí Then
        CopiarAlPortapapeles Base64Data
        Exit Function
    End If
    
    ConvertirImagenABase64 = Base64Data
End Function


Function CrearArchivoHTML( _
                            ByVal p_URLFinal As String, _
                            ByVal p_TextoHTML As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_archivo As Object
    
    On Error GoTo errores
  
    If fso.FileExists(p_URLFinal) Then
        If FicheroAbierto(p_URLFinal) Then
            p_Error = "Parece haber abierto un fichero con el mismo nombre"
            Err.Raise 1000
        End If
        fso.DeleteFile p_URLFinal, True
    End If
    ' Crear el m_archivo en la ruta especificada
    Set m_archivo = fso.CreateTextFile(p_URLFinal, True) ' True permite sobrescribir si el m_archivo existe
    
    ' Escribir el contenido en el m_archivo
    m_archivo.Write p_TextoHTML
    
    ' Cerrar el m_archivo
    m_archivo.Close
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CrearArchivoHTML ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function


Public Function RellenarComboTramitadores( _
                                            cmb As ComboBox, _
                                            Optional p_ConTodos As EnumSiNo = EnumSiNo.Sí, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_ID As Variant
    Dim m_Usuario As Usuario
    On Error GoTo errores
    
    
    cmb.RowSource = ""
    If p_ConTodos = EnumSiNo.Sí Then
        cmb.AddItem "Todos"
    End If
    If m_ObjEntorno.Coltramitadores Is Nothing Then
        Exit Function
    End If
    
    
    For Each m_ID In m_ObjEntorno.Coltramitadores
        Set m_Usuario = m_ObjEntorno.Coltramitadores(m_ID)
        cmb.AddItem m_Usuario.Nombre
        Set m_Usuario = Nothing
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboTramitadores a devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function RellenarComboMotivos( _
                                        cmb As ComboBox, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_Motivo As Variant
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    
    
    cmb.RowSource = ""
    Set m_Col = m_ObjEntorno.ColMotivos
    If m_Col Is Nothing Then
        Exit Function
    End If
    
    
    For Each m_Motivo In m_Col
        
        cmb.AddItem m_Motivo
        
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboMotivos a devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function AbrirEnLocal( _
                                p_URLFinal As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_URLLocal As String
    Dim m_Hwn As Long
    
    On Error Resume Next
    m_Hwn = Application.Screen.ActiveForm.hwnd
    If Err.Number <> 0 Then
        Err.Clear
        m_Hwn = 1
    End If
    On Error GoTo errores
    If Not fso.FileExists(p_URLFinal) Then
        p_Error = "No es accesible la ruta del archivo que se pretende abrir" & vbNewLine & p_URLFinal
        Err.Raise 1000
    End If
    If InStr(1, p_URLFinal, m_ObjEntorno.URLDirectorioLocal) <> 0 Then
        Ejecutar m_Hwn, "open", p_URLFinal, "", "", 1
        Exit Function
    End If
    m_URLLocal = m_ObjEntorno.URLDirectorioLocal & fso.GetFileName(p_URLFinal)
    If fso.FileExists(m_URLLocal) Then
        If FicheroAbierto(m_URLLocal) Then
            p_Error = "Tiene el archivo abierto"
            Err.Raise 1000
        End If
        fso.DeleteFile m_URLLocal, True
    End If
    fso.CopyFile p_URLFinal, m_URLLocal, True
    Ejecutar m_Hwn, "open", m_URLLocal, "", "", 1
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AbrirEnLocal ha producido el error nº: " & Err.Number & vbNewLine & _
                    "Detalle: " & Err.Description
    End If
End Function
Public Function AbrirAyuda( _
                            Optional ByRef p_Error As String _
                            ) As String

    
    Dim m_URLNombreArchivo As String
    On Error Resume Next
    If Application.Screen.ActiveForm Is Nothing Then
        If Err.Number <> 0 Then
            Exit Function
        End If
        Exit Function
    End If
    If Err.Number <> 0 Then
        Exit Function
    End If
    Err.Clear
    On Error GoTo errores
    m_URLNombreArchivo = m_ObjEntorno.URLDirectorioDocumentacionAyuda & Application.Screen.ActiveForm.Name & ".pdf"
    
    If Not fso.FileExists(m_URLNombreArchivo) Then
        p_Error = "No se ha generado la ayuda para este formulario aún"
        Err.Raise 1000
        
    End If
    AbrirEnLocal m_URLNombreArchivo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AbrirAyuda ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Sub Avance(ByRef p_Linea As Variant)
    Dim frm As Form
    On Error Resume Next
    
    Set frm = Screen.ActiveForm
    If Not frm Is Nothing Then
        If Not frm.Controls("lblEstado") Is Nothing Then
            frm.Controls("lblEstado").Caption = p_Linea
        End If
    End If
    
    If FormularioAbierto("frmSplash") Then
        With Forms("frmSplash")
            Dim totalPasos As Long
            Dim anchoMaximo As Long, nuevoAncho As Long
            
            Dim tempEntorno As New Entorno
            totalPasos = tempEntorno.ColItems.Count
            Set tempEntorno = Nothing
            
            s_contadorPasos = s_contadorPasos + 1
            
            anchoMaximo = .lblProgresoFondo.Width
            
            If totalPasos > 0 Then
                ' Cálculo proporcional
                nuevoAncho = (s_contadorPasos / totalPasos) * anchoMaximo
                
                ' SALVAGUARDA: Asegurarse de que el nuevo ancho no supere el máximo.
                If nuevoAncho > anchoMaximo Then
                    nuevoAncho = anchoMaximo
                End If
            End If
            
            .lblProgresoBarra.Width = nuevoAncho
        End With
    End If
    
    VBA.DoEvents
End Sub

Public Function AvanceCerrar( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    Dim frm As Form
    
    
    On Error GoTo errores
    
    
    Set frm = Application.Screen.ActiveForm
    If frm Is Nothing Then
        Exit Function
    End If
    On Error Resume Next
    If Not lbl Is Nothing Then
        lbl.Visible = False
        Set lbl = Nothing
    End If
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AvanceCerrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function CorreoAlAdministrador( _
                                        p_MensajeError As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                        
   
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Mensaje As String
    Dim m_Asunto As String
    Dim m_Nombre As String
    Dim m_NombreFormulario As String
    Dim m_NombreControl As String
    Dim m_TextoEnOficina As String
    Dim m_IDCorreo As String
    Dim m_Destinatarios As String
    Dim m_NombreEquipo As String
    
    On Error GoTo errores
    
    If p_MensajeError = "" Then
        p_Error = "No hay mensaje que enviar"
        Err.Raise 1000
    End If
    If m_EnOficina = Empty Then
        m_TextoEnOficina = "En Oficina Desconocido"
    Else
        If m_EnOficina = EnumSiNo.Sí Then
            m_TextoEnOficina = "En Oficina"
        Else
            m_TextoEnOficina = "Fuera de Oficina"
        End If
    End If
    On Error Resume Next
    m_NombreEquipo = VBA.Environ("COMPUTERNAME")
    m_NombreFormulario = Screen.ActiveForm.Name
    If Err.Number <> 0 Then
        m_NombreFormulario = "Formulario Desconocido"
        Err.Clear
        
    End If
    m_NombreControl = Screen.ActiveControl.Name
    If Err.Number <> 0 Then
        m_NombreControl = "Control Desconocido"
        Err.Clear
        
    End If
    
    m_Nombre = getNombreUsuarioConectado()
    
    m_Asunto = "Error en SOLICITUD HPS " & m_Nombre & " " & m_TextoEnOficina
    
    m_Mensaje = "FORMULARIO del ERROR: " & m_NombreFormulario
    m_Mensaje = m_Mensaje & "<BR> </BR>" & vbNewLine
    m_Mensaje = m_Mensaje & "NOMBRE DEL CONTROL: " & m_NombreControl
    m_Mensaje = m_Mensaje & "<BR> </BR>" & vbNewLine
    m_Mensaje = m_Mensaje & "NOMBRE EQUIPO: " & m_NombreEquipo
    m_Mensaje = m_Mensaje & "<BR> </BR>" & vbNewLine
    m_Mensaje = m_Mensaje & "DETALLE: " & p_MensajeError
    
    
    
    
    
    m_Destinatarios = "ardelperal@gmail.com;andres.romandelperal@telefonica.com"
   
    
    m_IDCorreo = DameID("TbCorreosEnviados", "IDCorreo", getdbCorreo())
    
    m_SQL = "TbCorreosEnviados"
    Set rcdDatos = getdbCorreo().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDCorreo") = m_IDCorreo
            .Fields("Aplicacion") = "HPS_SOLICITUD"
            .Fields("Originador") = m_ObjUsuarioConectado.UsuarioRed
            .Fields("Destinatarios") = m_Destinatarios
            .Fields("Asunto") = m_Asunto
            .Fields("Cuerpo") = m_Mensaje
            .Fields("FechaGrabacion") = Ahora
            
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoAlAdministrador ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function Seleccionar( _
                            p_EsArchivo As Boolean, _
                            Optional p_Titulo As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_ObjfDialog As Object
    Dim varFile As Variant
    
    On Error GoTo errores
    
    If p_Titulo = "" Then
        p_Titulo = "Seleccione el archivo"
    End If
    If p_EsArchivo = True Then
        Set m_ObjfDialog = Application.FileDialog(msoFileDialogFilePicker)
    Else
        Set m_ObjfDialog = Application.FileDialog(msoFileDialogFolderPicker)
    End If
    With m_ObjfDialog
        .Show
        If p_EsArchivo Then
            .AllowMultiSelect = False
            .InitialFileName = m_ObjEntorno.URLArchivoUltimo
            .Title = p_Titulo
            .Filters.Clear
            .Filters.Add "All Files", "*.*"
        End If
        For Each varFile In .SelectedItems
            Seleccionar = CStr(varFile)
        Next
    End With
    If p_EsArchivo Then
        m_ObjEntorno.URLArchivoUltimo = CStr(varFile)
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Seleccionar ha producido el error : " & vbNewLine & Err.Description
    End If
End Function

Public Function FormularioAbierto(Nombre As String) As Boolean
   
   
    Dim strEstado As String
    On Error GoTo errores
    strEstado = SysCmd(acSysCmdGetObjectState, acForm, Nombre)
    If strEstado = "0" Then
        FormularioAbierto = False
    Else
        FormularioAbierto = True
    End If
    Exit Function
errores:
    FormularioAbierto = False
End Function
Public Sub Ajustar(ByRef frmFormulario As Form)
    Dim i As Integer

    On Error Resume Next

    ' ajusto el ancho del formulario teniendo en cuenta si tiene o no selector de registros
    If Not frmFormulario.RecordSelectors Then
        frmFormulario.InsideWidth = frmFormulario.Width
    Else
        frmFormulario.InsideWidth = frmFormulario.Width + 250
    End If

    ' si se abre en vista formulario simple
    If frmFormulario.DefaultView = 0 Then
        'ajusto el alto incluyendo las distintas secciones, encabezado, pie, grupos...
        ' como no sé el número de secciones del formulario, me salgo al producirse un error
        frmFormulario.InsideHeight = 0
        For i = 0 To 100
            frmFormulario.InsideHeight = frmFormulario.InsideHeight + frmFormulario.Section(i).Height
        Next
    End If



End Sub        ' Ajustar
Public Function DameID( _
                        p_NOmbreTabla As String, _
                        p_NombreCampoID As String, _
                        Optional ByRef p_db As DAO.Database, _
                        Optional ByRef p_Error As String _
                        ) As String
   
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim lngOrdinalMaximo As Long
    On Error GoTo errores
    
    If p_db Is Nothing Then
        Set p_db = CurrentDb()
    End If
    m_SQL = "SELECT Max(" & p_NOmbreTabla & "." & p_NombreCampoID & ") AS Maximo " & _
            "FROM " & p_NOmbreTabla & ";"
    Set rcdDatos = p_db.OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("Maximo"), "")) Then
                lngOrdinalMaximo = .Fields("Maximo")
            End If
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameID = CStr(lngOrdinalMaximo + 1)
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameID ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Function FicheroAbierto(strURLArchivo As String) As Boolean
    
    Dim infNumeroFichero As Integer, intNumeroError As Integer
    On Error Resume Next
    infNumeroFichero = FreeFile()
    Open strURLArchivo For Input Lock Read As #infNumeroFichero
    Close infNumeroFichero          ' Close the file.
    intNumeroError = Err.Number
    On Error GoTo 0        ' Turn error checking back on.
    Select Case intNumeroError
        Case 0
         FicheroAbierto = False
    
        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case Else
            FicheroAbierto = True
    End Select
End Function
Public Function EjecutarShell( _
                                p_Comando As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim ManejadorProceso As Long
    Dim IDProceso As Long
    Dim lpExitCode As Long
    
    On Error GoTo errores
    If p_Comando = "" Then
        p_Error = "No se ha indicado el comando"
        Err.Raise 1000
    End If
    IDProceso = shell(p_Comando, vbHide)
    ManejadorProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, IDProceso)
    ' Mientras lp_ExitCode = STATUS_PENDING, se ejecuta el do
    Do
        Call GetExitCodeProcess(ManejadorProceso, lpExitCode)
        DoEvents
    Loop While lpExitCode = STATUS_PENDING
    Call CloseHandle(ManejadorProceso)
    
    EjecutarShell = "OK"
    Exit Function
errores:
    
    If Err.Number <> 1000 Then
        p_Error = "El método EjecutarShell ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function


Public Function EnOficina(Optional ByRef p_Error As String) As EnumSiNo
    
    Dim strIPS As String
    On Error GoTo errores
    strIPS = GetIPAddresses
    If InStr(1, strIPS, SubRedOficina) = 0 Then
        EnOficina = EnumSiNo.no
    Else
        EnOficina = EnumSiNo.Sí
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnOficina ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function



Public Function getURLFinal( _
                                p_URLLocal As String, _
                                p_IDSolicitud As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    On Error GoTo errores
    If p_URLLocal = "" Then
        p_Error = "Se ha de indicar una URL válida"
        Err.Raise 1000
    End If
    If p_IDSolicitud = "" Then
        p_Error = "Se ha de indicar un ID"
        Err.Raise 1000
    End If
    
    getURLFinal = m_ObjEntorno.URLDirectorioDocumentacion & "Solicitudes\" & Format(p_IDSolicitud, "000000") & "\" & fso.GetFileName(p_URLLocal)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getURLFinal ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function RegistrarLogGeneral( _
                                    p_Accion As String, _
                                    Optional p_Descripcion As String, _
                                    Optional p_Solicitud As solicitud, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim m_LogOp As LogGeneralOperaciones
    Dim m_Log As LogGeneral
    Dim m_CorreoAutomaticoResultante As String
    
    On Error GoTo errores
    m_CorreoAutomaticoResultante = CorreoAutomaticoResultante(p_Solicitud, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_Log = New LogGeneral
    With m_Log
        .Accion = p_Accion
        .Descripcion = p_Descripcion
        .CorreosAutomaticos = m_CorreoAutomaticoResultante
        If Not p_Solicitud Is Nothing Then
            If p_Solicitud.DNI <> "" Then
                m_Log.DNI = p_Solicitud.DNI
            End If
            .IDSolicitud = p_Solicitud.IDSolicitud
        End If
    End With
    
    Set m_LogOp = New LogGeneralOperaciones
    With m_LogOp
        Set .Log = m_Log
        .Registrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistrarLogGeneral ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getHora() As String
    Dim xmlHttp As Object
    Dim URL As String
    Dim response As String
    Dim json As Object
    Dim hora As String
    Dim dato As Variant
    Dim flag As String
    On Error GoTo errores
    ' URL de la API para obtener la hora en Madrid, España
    URL = "http://worldtimeapi.org/api/timezone/Europe/Madrid"
    
    ' Crear el objeto XMLHTTP
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Hacer la solicitud a la API
    xmlHttp.Open "GET", URL, False
    xmlHttp.Send
    
    ' Obtener la respuesta
    response = xmlHttp.responseText
    
    ' Analizar la respuesta JSON
    Set json = JsonConverter.ParseJson(response)
    
    ' Extraer la hora de la respuesta JSON
    hora = json("datetime")
    dato = Split(hora, "T")
    flag = dato(1)
    dato = Split(flag, ".")
    hora = dato(0)
    ' Retornar la hora
    getHora = hora
    Exit Function
errores:
    getHora = Format(Now(), "hh:mm:ss")
    
End Function
Public Function Ahora() As String
    Dim m_Ahora
   
    m_Ahora = getHora()
    If Not IsDate(m_Ahora) Then
        m_Ahora = Date
    End If
    m_Ahora = Date & " " & m_Ahora
    ' Retornar la hora
    Ahora = m_Ahora
    
    
End Function
Function EsCorreoValido(ByVal Correo As String) As Boolean
    Dim dato As Variant
    Dim m_Correo As Variant
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Configuración de la expresión regular
    regex.Pattern = "^[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$"
    regex.IgnoreCase = True
    regex.Global = False
    dato = Split(Correo, ";")
    For Each m_Correo In dato
        If Not regex.Test(m_Correo) Then
            EsCorreoValido = False
            Exit Function
        End If
    Next
    ' Verificar si el correo coincide con el patrón
    EsCorreoValido = True
End Function
Function FormatearNombreCompleto(p_NombreCompleto As String) As String
    ' Capitalizar cada palabra
    Dim NombreCompleto As String
    Dim dato As Variant
    Dim m_Trozo As Variant
    Dim m_TrozoCapitalizado As String
    
    dato = Split(p_NombreCompleto, " ")
    For Each m_Trozo In dato
        m_TrozoCapitalizado = Capitalizar(CStr(m_Trozo))
        If NombreCompleto = "" Then
            NombreCompleto = m_TrozoCapitalizado
        Else
            NombreCompleto = NombreCompleto & " " & m_TrozoCapitalizado
        End If
    Next
    
    
    
    FormatearNombreCompleto = NombreCompleto
End Function

Function Capitalizar(texto As String) As String
    ' Convierte solo la primera letra en mayúscula y el resto en minúscula
    Dim dato As Variant
    Dim m_Texto As Variant
    Dim m_TextoCompleto As String
    dato = Split(texto, " ")
    For Each m_Texto In dato
        If m_TextoCompleto = "" Then
            m_TextoCompleto = UCase(Left(m_Texto, 1)) & LCase(Mid(m_Texto, 2))
        Else
            m_TextoCompleto = m_TextoCompleto & " " & UCase(Left(m_Texto, 1)) & LCase(Mid(m_Texto, 2))
        End If
    Next
    Capitalizar = m_TextoCompleto
End Function


Public Function CorreoAlServidor( _
                                p_Correo As Correo, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_FechaEnvio As String
    Dim m_URL As String
    Dim m_Cuerpo As String
    Dim m_CadenaVariables As String
    Dim varItem As Variant
    Dim dato As Variant
    Dim MiMensaje As Object
    Dim MiConfiguracion As Object
    
    On Error GoTo errores
    
    m_Cuerpo = getMensajeDeCorreo(p_Correo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set MiConfiguracion = CreateObject("CDO.Configuration")
    With MiConfiguracion
        .Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "10.73.54.85"
        .Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 5
        .Fields.Update
    End With
    Set MiMensaje = CreateObject("CDO.Message")
    With MiMensaje
        Set .Configuration = MiConfiguracion
        .FROM = p_Correo.Aplicacion & ".DySN@telefonica.com"
        If InStr(1, p_Correo.DESTINATARIOS, "@") <> 0 Then
            .To = p_Correo.DESTINATARIOS
        End If
        If InStr(1, p_Correo.DestinatariosConCopia, "@") <> 0 Then
            .Cc = p_Correo.DestinatariosConCopia
        End If
        If p_Correo.DestinatariosConCopiaOculta <> "" Then
            .BCC = p_Correo.DestinatariosConCopiaOculta
        End If
        .Subject = p_Correo.Asunto
      .HTMLBody = m_Cuerpo
        If Not IsNull(p_Correo.URLAdjunto) Then
            If InStr(1, p_Correo.URLAdjunto, ";") <> 0 Then
                dato = Split(p_Correo.URLAdjunto, ";")
                For Each varItem In dato
                    m_URL = CStr(varItem)
                    If fso.FileExists(m_URL) Then
                        .AddAttachment m_URL
                    End If
                Next
            Else
                If fso.FileExists(p_Correo.URLAdjunto) Then
                    .AddAttachment p_Correo.URLAdjunto
                End If
            End If
        End If
        
        On Error Resume Next
       .Send
       On Error GoTo errores
       'm_FechaEnvio = Ahora
        If Err.Number = 0 Then
            m_FechaEnvio = Ahora
            CorreoAlServidor = m_FechaEnvio
         Else
            
            CorreoAlServidor = ""
        End If
        
    End With
    
    Set MiMensaje = Nothing
    Set MiConfiguracion = Nothing
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoAlServidor ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getMensajeDeCorreo( _
                                    p_Correo As Correo, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    Dim m_CadenaRecursos As String
    
    Dim m_NombrePlantilla As String
    Dim m_URLPlantilla As String
    Dim dato As Variant
    
    
    On Error GoTo errores
    
    m_CadenaRecursos = p_Correo.cadenaRecursos
    m_NombrePlantilla = p_Correo.nombrePlantilla
    m_URLPlantilla = m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida & m_NombrePlantilla
    If Not fso.FileExists(m_URLPlantilla) Then
        p_Error = "No se puede obtener la plantilla del HTML"
        Err.Raise 1000
    End If
   
    

    getMensajeDeCorreo = getTextoDeHTML(m_URLPlantilla, , m_CadenaRecursos, EnumSiNo.Sí, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getMensajeDeCorreo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function GenerarCorreo( _
                            Optional ByRef p_Error As String _
                            ) As String
                            
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoAlServidor1 ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function GenerarHTML( _
                            Optional p_contenidoHTML As String, _
                            Optional p_URLPlantillaHTML As String, _
                            Optional p_ColVariables As Scripting.Dictionary, _
                            Optional p_RutaDestino As String, _
                            Optional p_ParaCorreo As EnumSiNo = EnumSiNo.Sí, _
                            Optional ByRef p_Error As String) As String
    
    
    
    
    Dim stream As ADODB.stream
    
    On Error GoTo errores
    
    If p_contenidoHTML = "" Then
        If Not fso.FileExists(p_URLPlantillaHTML) Or p_ColVariables Is Nothing Then
            p_Error = "Sin no se da el TextoHTM se ha de dar la plantilla y la colección de variables"
            Err.Raise 1000
        End If
    End If
    If p_contenidoHTML = "" Then
        p_contenidoHTML = getTextoDeHTML(p_URLPlantillaHTML, p_ColVariables, , p_ParaCorreo, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If p_RutaDestino = "" Then
        p_RutaDestino = m_ObjEntorno.URLDirectorioLocal & fso.GetBaseName(fso.GetTempName()) & ".html"
        If fso.FileExists(p_RutaDestino) Then
            fso.DeleteFile p_RutaDestino, True
        End If
    End If
    
    
    Set stream = New ADODB.stream
    With stream
        .Type = 2 ' 2 indica texto
        .Charset = "UTF-8"
        .Open
        .WriteText p_contenidoHTML
        .SaveToFile p_RutaDestino, 2 ' 2 para sobrescribir si existe
        .Close
    End With
    Set stream = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarHTML ha devuelto el error: " & vbNewLine & Err.Description
    End If
   
End Function

Public Function getTextoDeHTML( _
                                p_URLPlantillaHTML As String, _
                                Optional p_ColVariables As Scripting.Dictionary, _
                                Optional p_ColVariablesTexto As String, _
                                Optional p_ParaCorreo As EnumSiNo = EnumSiNo.Sí, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    
    Dim m_contenidoHTML As String
    Dim m_clave As Variant
    Dim m_Valor As String
    Dim stream As ADODB.stream
    Dim m_MetaCharset As String
    Const m_CharsetISO As String = "<meta charset='ISO-8859-1'>"
    Const m_CharsetUTF As String = "<meta charset='UTF-8'>"
    
    
    If p_ParaCorreo <> EnumSiNo.Sí And p_ParaCorreo <> EnumSiNo.no Then
        p_ParaCorreo = EnumSiNo.Sí
    End If
    If p_ParaCorreo = EnumSiNo.Sí Then
        m_MetaCharset = m_CharsetISO
    Else
        m_MetaCharset = m_CharsetUTF
    End If
    If p_ColVariables Is Nothing And p_ColVariablesTexto <> "" Then
        Set p_ColVariables = getColVariablesDeTexto(p_ColVariablesTexto, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
       
    End If
    
    
    Set stream = New ADODB.stream
    With stream
        .Type = 2 ' 2 indica texto
        .Charset = "UTF-8"
        .Open
        .LoadFromFile (p_URLPlantillaHTML)
        m_contenidoHTML = .ReadText
        .Close
    
    End With
    Set stream = Nothing
    ' Recorrer la colección de variables y reemplazar en el HTML
    For Each m_clave In p_ColVariables
        m_Valor = p_ColVariables(m_clave)
        m_contenidoHTML = Replace(m_contenidoHTML, m_clave, m_Valor)
        
    Next
    If p_ParaCorreo = EnumSiNo.no Then
        m_clave = m_CharsetISO
        m_Valor = m_CharsetUTF
    Else
        m_clave = m_CharsetUTF
        m_Valor = m_CharsetISO
    End If
    m_contenidoHTML = Replace(m_contenidoHTML, m_clave, m_Valor)
    getTextoDeHTML = m_contenidoHTML
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getTextoDeHTML ha devuelto el error: " & vbNewLine & Err.Description
    End If
   
   
End Function
Public Function getColVariablesDeTexto( _
                                        p_ColVariablesTexto As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    
    Dim parser As ParametrosParser
    Set parser = New ParametrosParser
    
    On Error GoTo errores
    
    If p_ColVariablesTexto = "" Then
        Set getColVariablesDeTexto = New Scripting.Dictionary
        Exit Function
    End If
    
    ' Usar ParametrosParser en lugar del parseo manual
    Set getColVariablesDeTexto = parser.ParsearCadenaRecursos(p_ColVariablesTexto)
    
    ' Verificar si hubo errores
    If parser.Error <> "" Then
        p_Error = parser.Error
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColVariablesDeTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Set getColVariablesDeTexto = New Scripting.Dictionary
End Function

' [MODIFICAR EN FUNCIONES UTILES.bas]
Public Function ObtenerHTMLEnLocal( _
                                    p_URLPlantillaHTML As String, _
                                    Optional p_IDSolicitud As String, _
                                    Optional p_Solicitud As solicitud, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_ColVariables As Scripting.Dictionary
    Dim m_RutaDestino As String
    
    On Error GoTo errores
    If p_Solicitud Is Nothing Then
        Set p_Solicitud = constructor.getSolicitud(p_IDSolicitud:=p_IDSolicitud, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If p_Solicitud Is Nothing Then
            p_Error = "No se puede obtener la solicitud"
            Err.Raise 1000
        End If
    End If
    
    
    With p_Solicitud
        Set m_ColVariables = .ColParametrosParaHTML
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    
    m_RutaDestino = m_ObjEntorno.URLDirectorioLocal & fso.GetBaseName(fso.GetTempName()) & ".html"
    
    ' Generamos el HTML (texto) en local
    GenerarHTML , p_URLPlantillaHTML, m_ColVariables, m_RutaDestino, EnumSiNo.no, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' --- NUEVO BLOQUE: COPIAR IMÁGENES ---
    ' Antes de visualizar, nos aseguramos que la carpeta IMG está al lado del HTML generado
    SincronizarImagenesLocal p_Error
    ' -------------------------------------
    
    m_URLHTMLActivo = m_RutaDestino
    If FormularioAbierto("FormWeb") Then
        DoCmd.Close acForm, "FormWeb", acSaveNo
    End If
    DoCmd.OpenForm "FormWeb"
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ObtenerHTMLEnLocal ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function RellenarComboGrado( _
                                    p_Tipo As String, _
                                    ByRef p_Cbo As ComboBox, _
                                    Optional ByRef p_Error As String _
                                    ) As String

    Dim m_ObjColGrados As Scripting.Dictionary
    Dim m_Grado As Variant
    On Error GoTo errores
    
    p_Cbo.RowSource = ""
    If Not m_ObjEntorno.colGradosHPS.Exists(p_Tipo) Then
        p_Error = "No hay ningún listado de grados para el tipo " & p_Tipo
        Err.Raise 1000
    End If
    Set m_ObjColGrados = m_ObjEntorno.colGradosHPS(p_Tipo)
    If Not m_ObjColGrados Is Nothing Then
        For Each m_Grado In m_ObjColGrados
            p_Cbo.AddItem m_Grado
        Next
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboGrado ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function RegistrarTbMotivoHPS( _
                                    Optional ByRef p_Error As String _
                                    ) As String

    
    Dim m_SQL As String
    On Error GoTo errores
    
    m_SQL = "INSERT INTO TbMotivoHPS ( Motivo_HPS ) " & _
            "SELECT distinct TbMotivoHPS.Motivo_HPS " & _
            "FROM TbMotivoHPS;"
    getdbHPS().Execute m_SQL
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistrarTbMotivoHPS ha devuelto el error: " & Err.Description
    End If

End Function
Public Function MeterUnMotivoEnListaMOtivos( _
                                            ByRef p_Motivo As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String

    
    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    
    On Error GoTo errores
    If p_Motivo = "" Then
        Exit Function
    End If
    p_Motivo = Trim(p_Motivo)
    
    m_SQL = "SELECT TbMotivoHPS.* " & _
            "FROM TbMotivoHPS " & _
            "WHERE (((TbMotivoHPS.MotivoHPS)='" & p_Motivo & "'));"
    Set rcdDatos = getdbHPS().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            .AddNew
                .Fields("MotivoHPS") = p_Motivo
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MeterUnMotivoEnListaMOtivos ha devuelto el error: " & Err.Description
    End If

End Function

Public Function DameCabeceraHTML( _
                                p_Titulo As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_Mensaje As String
    
    On Error GoTo errores
    
    
    m_Mensaje = "<!DOCTYPE html>" & vbNewLine
    m_Mensaje = m_Mensaje & "<html lang=""es"">" & vbNewLine
    m_Mensaje = m_Mensaje & "<head>" & vbNewLine
        m_Mensaje = m_Mensaje & "<title>" & p_Titulo & "</title>" & vbNewLine
        m_Mensaje = m_Mensaje & "<meta charset=""ISO-8859-1"" />" & vbNewLine
        'm_Mensaje = m_Mensaje & "<meta charset=""UTF-8"">" & vbnewline
        
        m_Mensaje = m_Mensaje & "<style type=""text/css"">" & vbNewLine
            m_Mensaje = m_Mensaje & m_ObjEntorno.CSS & vbNewLine
        m_Mensaje = m_Mensaje & "</style>" & vbNewLine
    m_Mensaje = m_Mensaje & "</head>" & vbNewLine
    m_Mensaje = m_Mensaje & "<body>" & vbNewLine
    DameCabeceraHTML = m_Mensaje
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameCabeceraHTML ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Private Function BorraHTMLs( _
                            Optional ByRef p_Error As String) As String
    
    Dim fichero As File
    
    On Error GoTo errores
    
    For Each fichero In fso.GetFolder(m_ObjEntorno.URLDirectorioLocal).Files
        If fso.GetExtensionName(fichero.Path) = "html" Or fso.GetExtensionName(fichero.Path) = "htm" Then
            If Not FicheroAbierto(fichero.Path) Then
                fso.DeleteFile fichero.Path
            End If
        End If
    Next
   
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método BorraHTMLs ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Function DameUntxtYHtml(Optional ByRef p_Error As String) As String
    Dim i As Integer
    Dim m_URLTXT As String
    Dim m_URLHTML As String
    Dim m_NombreHTML As String
    Dim m_Nombretxt As String
    Dim m_URLDirLocal As String
    On Error GoTo errores
    
    m_URLDirLocal = m_ObjEntorno.URLDirectorioLocal
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    BorraHTMLs p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    For i = 1 To 50
        m_Nombretxt = "HTML" & i & ".txt"
        m_NombreHTML = "HTML" & i & ".html"
        m_URLTXT = m_URLDirLocal & m_Nombretxt
        m_URLHTML = m_URLDirLocal & m_NombreHTML
        If Not fso.FileExists(m_URLTXT) And Not fso.FileExists(m_URLHTML) Then
            DameUntxtYHtml = m_URLHTML
            Exit Function
        End If
        
    Next
    p_Error = "No se ha podido obtener ningún html"
    Err.Raise 1000
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTMLInformeCompletoDPD ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function GenerarArchivoConHTML( _
                                    Optional p_Mensaje As String, _
                                    Optional ByRef p_Error As String) As String
    
    
    
    
    Dim stream As ADODB.stream
    Dim m_URL As String
    Dim m_Hwd As Long
    On Error GoTo errores
    m_URL = DameUntxtYHtml(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_URL = "" Then
        p_Error = "No se ha podido obtener el HTML"
        Err.Raise 1000
    End If
    
    
    
    Set stream = New ADODB.stream
    With stream
        .Type = 2 ' 2 indica texto
        .Charset = "UTF-8"
        .Open
        .WriteText p_Mensaje
        
        .SaveToFile m_URL, 2 ' 2 para sobrescribir si existe
        .Close
    End With
    Set stream = Nothing
    
    
    GenerarArchivoConHTML = m_URL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarArchivoConHTML ha devuelto el error: " & vbNewLine & Err.Description
    End If
   
End Function
Public Function HTMLLog( _
                            p_IDSolicitud As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    Dim m_Mensaje As String
    Dim m_Cabecera As String
    Dim m_HTMLTablaLogdeSolicitud As String
    
    
    On Error GoTo errores
    
    
    Avance "HTMLLog ....."
   
    m_Cabecera = DameCabeceraHTML("Informe de Acciones", p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_HTMLTablaLogdeSolicitud = HTMLTablaLogdeSolicitud(p_IDSolicitud, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_Mensaje = m_Cabecera & vbNewLine
    
    m_Mensaje = m_Mensaje & m_HTMLTablaLogdeSolicitud & vbNewLine
    
    m_Mensaje = m_Mensaje & "</body>" & vbNewLine
    m_Mensaje = m_Mensaje & "</html>" & vbNewLine
    HTMLLog = m_Mensaje
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTMLLog ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function



Public Function HTMLTablaLogdeSolicitud( _
                                            p_IDSolicitud As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
   
    Dim m_Mensaje As String
    Dim m_Log As LogGeneral
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    
    Dim m_Registrador As String
    On Error GoTo errores
    
    If p_IDSolicitud = "" Then
       Exit Function
    End If
    Set m_Col = constructor.getLogsGenerales(p_IDSolicitud:=p_IDSolicitud, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_Mensaje = m_Mensaje & "<table>" & vbNewLine
        m_Mensaje = m_Mensaje & "<tr>" & vbNewLine
            m_Mensaje = m_Mensaje & "<td colspan='5' class=""ColespanArriba""> REGISTROS LOGS DE ACTIVIDADES DE SOLICITUD " & p_IDSolicitud & " </td>"
        m_Mensaje = m_Mensaje & "</tr>" & vbNewLine
        m_Mensaje = m_Mensaje & "<tr>" & vbNewLine
            m_Mensaje = m_Mensaje & "<td class=""Cabecera"" > FECHA</td>" & vbNewLine
            m_Mensaje = m_Mensaje & "<td class=""Cabecera"" > ACCIÓN</td>" & vbNewLine
            m_Mensaje = m_Mensaje & "<td class=""Cabecera"" > DESCRIPCIÓN</td>" & vbNewLine
            m_Mensaje = m_Mensaje & "<td class=""Cabecera"" > ACTIVADO CORREO AUTOMÁTICO</td>" & vbNewLine
            m_Mensaje = m_Mensaje & "<td class=""Cabecera"" > REGISTRADOR</td>" & vbNewLine
            
        m_Mensaje = m_Mensaje & "</tr>" & vbNewLine
        If Not m_Col Is Nothing Then
            For Each m_ID In m_Col
                Set m_Log = m_Col(m_ID)
               
                If m_Log.UsuarioObj Is Nothing Then
                    m_Registrador = "Desconocido"
                Else
                    m_Registrador = m_Log.UsuarioObj.Nombre
                End If
                m_Mensaje = m_Mensaje & "<tr>" & vbNewLine
                    m_Mensaje = m_Mensaje & "<td> " & m_Log.FECHA & "</td>" & vbNewLine
                    m_Mensaje = m_Mensaje & "<td> " & m_Log.Accion & "</td>" & vbNewLine
                    m_Mensaje = m_Mensaje & "<td> " & IIf(m_Log.Descripcion = "", "&nbsp;", m_Log.Descripcion) & "</td>" & vbNewLine
                    m_Mensaje = m_Mensaje & "<td class=""centrado""> " & m_Log.CorreosAutomaticos & "</td>" & vbNewLine
                    m_Mensaje = m_Mensaje & "<td> " & m_Registrador & "</td>" & vbNewLine
                m_Mensaje = m_Mensaje & "</tr>" & vbNewLine
            Next
        End If
    m_Mensaje = m_Mensaje & "</table>" & vbNewLine
        
        
        
    
    HTMLTablaLogdeSolicitud = m_Mensaje
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTMLTablaLogdeSolicitud ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function InformeDelRegistro( _
                                    p_IDSolicitud As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    Dim m_Mensaje As String
   
    
    On Error GoTo errores
    m_Mensaje = HTMLLog(p_IDSolicitud, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_URLHTMLActivo = GenerarArchivoConHTML(m_Mensaje, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If FormularioAbierto("FormWeb") Then
        DoCmd.Close acForm, "FormWeb", acSaveNo
    End If
    DoCmd.OpenForm "FormWeb"
    InformeDelRegistro = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método InformeDelRegistro ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function CorreoAutomaticoResultante( _
                                            p_Solicitud As solicitud, _
                                            Optional ByRef p_Error As String _
                                            ) As String
   
    On Error GoTo errores
    
    If m_ObjEntorno.CorreosAutomaticos = "No" Then
        CorreoAutomaticoResultante = "No"
        Exit Function
    End If
    If Not p_Solicitud Is Nothing Then
        CorreoAutomaticoResultante = p_Solicitud.CorreosAutomaticos
    Else
        CorreoAutomaticoResultante = "Sí"
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoAutomaticoResultante ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Function GenerarCIFValido() As String
    Dim LetrasIniciales As String
    Dim LetrasControl As String
    Dim LetraInicial As String
    Dim Numeros As String
    Dim SumaPares As Integer
    Dim SumaImpares As Integer
    Dim total As Integer
    Dim DigitoControl As String
    Dim i As Integer
    Dim Digito As Integer

    ' Definir letras iniciales y de control
    LetrasIniciales = "ABCDEFGHIJNPQRSUVW"
    LetrasControl = "JABCDEFGHI"

    ' Seleccionar una letra inicial aleatoria
    LetraInicial = Mid(LetrasIniciales, Int((Len(LetrasIniciales) * Rnd) + 1), 1)

    ' Generar los 7 números aleatorios intermedios
    For i = 1 To 7
        Numeros = Numeros & CStr(Int(10 * Rnd))
    Next i

    ' Calcular la suma de pares e impares
    SumaPares = 0
    SumaImpares = 0
    For i = 1 To 7 Step 2
        SumaPares = SumaPares + Val(Mid(Numeros, i + 1, 1))
    Next i
    For i = 1 To 7 Step 2
        Digito = Val(Mid(Numeros, i, 1)) * 2
        SumaImpares = SumaImpares + (Digito \ 10) + (Digito Mod 10)
    Next i

    ' Calcular el dígito de control
    total = SumaPares + SumaImpares
    DigitoControl = CStr((10 - (total Mod 10)) Mod 10)

    ' Determinar el carácter de control según la letra inicial
    Select Case LetraInicial
        Case "A", "B", "E", "H"
            ' El carácter de control debe ser numérico
            GenerarCIFValido = LetraInicial & Numeros & DigitoControl
        Case "K", "P", "Q", "S"
            ' El carácter de control debe ser una letra
            GenerarCIFValido = LetraInicial & Numeros & Mid(LetrasControl, DigitoControl + 1, 1)
        Case Else
            ' Puede ser letra o número
            If Rnd > 0.5 Then
                GenerarCIFValido = LetraInicial & Numeros & DigitoControl
            Else
                GenerarCIFValido = LetraInicial & Numeros & Mid(LetrasControl, DigitoControl + 1, 1)
            End If
    End Select
End Function
Function GenerarDNINIEValido() As String
    Dim LetrasDNI As String
    Dim LetrasNIE As String
    Dim TipoDocumento As String
    Dim NumeroBase As String
    Dim LetraControl As String
    Dim NumeroCalculado As Long

    ' Letras válidas para DNI y NIE
    LetrasDNI = "TRWAGMYFPDXBNJZSQVHLCKE"
    LetrasNIE = "XYZ"

    ' Determinar si generar un DNI o un NIE (50% probabilidad para cada uno)
    If Rnd > 0.5 Then
        TipoDocumento = "DNI"
    Else
        TipoDocumento = "NIE"
    End If

    ' Generar el número base
    If TipoDocumento = "DNI" Then
        ' Generar un número aleatorio de 8 dígitos para DNI
        NumeroBase = Format(Int((10 ^ 8) * Rnd), "00000000")
    Else
        ' Generar un número NIE (1 letra inicial + 7 dígitos)
        NumeroBase = Mid(LetrasNIE, Int((Len(LetrasNIE) * Rnd) + 1), 1) & _
                     Format(Int((10 ^ 7) * Rnd), "0000000")
    End If

    ' Convertir NIE a número si es necesario (X=0, Y=1, Z=2)
    If TipoDocumento = "NIE" Then
        Select Case Left(NumeroBase, 1)
            Case "X"
                NumeroCalculado = Val("0" & Mid(NumeroBase, 2))
            Case "Y"
                NumeroCalculado = Val("1" & Mid(NumeroBase, 2))
            Case "Z"
                NumeroCalculado = Val("2" & Mid(NumeroBase, 2))
        End Select
    Else
        NumeroCalculado = Val(NumeroBase)
    End If

    ' Calcular la letra de control
    LetraControl = Mid(LetrasDNI, (NumeroCalculado Mod 23) + 1, 1)

    ' Construir el DNI o NIE válido
    GenerarDNINIEValido = NumeroBase & LetraControl
End Function
Function ValidarCIF(CIF As String) As Boolean
    Dim LetrasIniciales As String
    Dim LetrasFinales As String
    Dim SumaPar As Integer
    Dim SumaImpar As Integer
    Dim total As Integer
    Dim Resto As Integer
    Dim DigitoControl As String
    Dim DigitosDelCIF As String
    Dim i As Integer
    
    ' Definir las letras iniciales y finales válidas
    LetrasIniciales = "ABCDEFGHKLMNPQS"
    LetrasFinales = "JABCDEFGHI"
    
    ' Validar longitud y formato inicial
    If Len(CIF) <> 9 Then
        ValidarCIF = False
        Exit Function
    End If
    
    If InStr(1, LetrasIniciales, Left(CIF, 1)) = 0 Then
        ValidarCIF = False
        Exit Function
    End If
    DigitosDelCIF = Mid(CIF, 2, 7)
    ' Calcular suma de dígitos pares e impares
    SumaPar = 0
    SumaImpar = 0
    For i = 1 To 7
        Resto = i Mod 2
        If Resto = 0 Then
            SumaPar = SumaPar + CInt(Mid(DigitosDelCIF, i, 1))
        End If
        
    Next i
    Dim Impar As Integer
    For i = 1 To 7 Step 1
        
        Resto = i Mod 2
        If Resto <> 0 Then
            Impar = CInt(Mid(DigitosDelCIF, i + 1, 1)) * 2
            SumaImpar = SumaImpar + Impar
        End If
        
    Next i
    
    ' Calcular el dígito de control
    total = SumaPar + SumaImpar
    Resto = total Mod 10
    If Resto <> 0 Then
        Resto = 10 - Resto
    End If
    
    ' Comparar dígito de control
    If IsNumeric(Right(CIF, 1)) Then
        DigitoControl = CStr(Resto)
    Else
        DigitoControl = Mid(LetrasFinales, Resto + 1, 1)
    End If
    
    If Right(CIF, 1) = DigitoControl Then
        ValidarCIF = True
    Else
        ValidarCIF = False
    End If
End Function
Public Function getDirectorioOneDrive(Optional ByRef p_Error As String) As String
    Dim fso As Object
    Dim carpetaRaiz As Object
    Dim subCarpeta As Object
    Dim rutaEncontrada As String
    Dim encontrado As Boolean
    Dim dato As Variant
    On Error GoTo errores
    
    rutaEncontrada = Environ("OneDrive")
    If InStr(1, rutaEncontrada, "OneDrive") <> 0 Then
        dato = Split(rutaEncontrada, "OneDrive")
        getDirectorioOneDrive = dato(0) & "OneDrive"
        Exit Function
    End If
    
    
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Obtener la carpeta raíz de C:\
    Set carpetaRaiz = fso.GetFolder("C:\")
    
    ' Inicializar variables
    encontrado = False
    rutaEncontrada = ""
    
    ' Recorrer las subcarpetas en la raíz de C:\
    For Each subCarpeta In carpetaRaiz.SubFolders
        If InStr(1, subCarpeta.Name, "OneDrive", vbTextCompare) > 0 Then
            rutaEncontrada = subCarpeta.Path
            encontrado = True
            Exit For
        
        End If
    Next subCarpeta
    
    ' Mostrar el resultado
    If encontrado Then
        getDirectorioOneDrive = rutaEncontrada
   
    End If
    
    ' Liberar objetos
    Set subCarpeta = Nothing
    Set carpetaRaiz = Nothing
    Set fso = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDirectorioOneDrive ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function


Private Function getDirectorioOneDriveTelefonicaApps(Optional ByRef p_Error As String) As String
    Dim fso As Object
    Dim carpeta As String
    Dim m_RutaOneDrive As String
    
    On Error GoTo errores
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    m_RutaOneDrive = getDirectorioOneDrive(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RutaOneDrive = "" Then
        Exit Function
    End If
    carpeta = m_RutaOneDrive & "\Telefonica\Aplicaciones_dys.TMETF - Aplicaciones PpD\"
    If Not fso.FolderExists(carpeta) Then
        Exit Function
    End If
    getDirectorioOneDriveTelefonicaApps = carpeta
    
    ' Liberar objetos
   
    
    Set fso = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDirectorioOneDriveTelefonicaApps ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function
Public Function getDirectorioOneDriveApps(Optional ByRef p_Error As String) As String
    Dim fso As Object
    Dim carpeta As String
    Dim m_RutaOneDrive As String
    
    On Error GoTo errores
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    m_RutaOneDrive = getDirectorioOneDrive(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RutaOneDrive = "" Then
        Exit Function
    End If
    'C:\OneDrive\OneDrive - Telefonica\00LABORAL\Aplicaciones PpD
    carpeta = m_RutaOneDrive & " - Telefonica\00LABORAL\Aplicaciones PpD\"
    If Not fso.FolderExists(carpeta) Then
        Exit Function
    End If
    getDirectorioOneDriveApps = carpeta
    
    ' Liberar objetos
   
    
    Set fso = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDirectorioOneDriveApps ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function


Public Function getRutaAplicacionesLocal(Optional ByRef p_Error As String) As String
    Dim fso As Object
    Dim m_RutaOneDrive As String
    Dim m_RutaOneDriveTelefonica As String
    
    On Error GoTo errores
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    'm_RutaOneDriveTelefonica = getDirectorioOneDriveTelefonicaApps(p_Error)
    m_RutaOneDriveTelefonica = getDirectorioOneDriveApps(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If fso.FolderExists(m_RutaOneDriveTelefonica) Then
        Set fso = Nothing
        getRutaAplicacionesLocal = m_RutaOneDriveTelefonica
        Exit Function
    End If
    
    m_RutaOneDrive = getDirectorioOneDrive(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RutaOneDrive = "" Then
        Set fso = Nothing
        Exit Function
    End If
   
    getRutaAplicacionesLocal = m_RutaOneDrive
    
    ' Liberar objetos
   
    
    Set fso = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRutaAplicacionesLocal ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function
Public Function GenerarConsultaSolicitudes( _
                                            Optional p_Col As Scripting.Dictionary, _
                                            Optional ByRef p_Error As String _
                                            ) As String
   
    
    Dim m_ID As Variant
    Dim m_Solicitud As New solicitud
    Dim m_SolicitudFechas As New SolicitudFechas
    Dim m_ColCampos As Scripting.Dictionary
    Dim m_Objeto As String
    Dim m_Campo As Variant
    Dim m_NombreCampo As String
    Dim intFila As Integer
    Dim columna As Integer
    Dim m_NombreArchivo As String
    Dim m_URLExcel As String
    Dim appExcel As Excel.Application
    Dim wbLibro As Excel.Workbook
    Dim wbHoja As Excel.Worksheet
    Dim m_Valor As String
    On Error GoTo errores
    
    
    
    If p_Col Is Nothing Then
        Exit Function
    End If
    Set m_ColCampos = New Scripting.Dictionary
    For Each m_Campo In m_Solicitud.ColCampos
        If CStr(m_Campo) = "IDExpediente" Then m_Campo = "Expediente"
        If CStr(m_Campo) = "IDEmpresaUsuario" Then m_Campo = "EmpresaUsuario"
        If CStr(m_Campo) = "IDEmpresaTramitadora" Then m_Campo = "EmpresaTramitadora"
        If Not m_ColCampos.Exists(m_Campo) Then
            m_ColCampos.Add m_Campo, "SO"
        End If
    Next
    For Each m_Campo In m_SolicitudFechas.ColCampos
        If Not m_ColCampos.Exists(m_Campo) Then
            m_ColCampos.Add m_Campo, "SOFECHAS"
        End If
    Next
    
    Avance "Abriendo excel para rellenar ..."
    m_NombreArchivo = fso.GetTempName & ".xlsx"
    m_URLExcel = m_ObjEntorno.URLDirectorioLocal & m_NombreArchivo
    If fso.FileExists(m_URLExcel) Then
        If FicheroAbierto(m_URLExcel) Then
            p_Error = "Tiene una consulta abierta"
            Err.Raise 1000
        End If
        fso.DeleteFile m_URLExcel, True
    End If
    Set appExcel = New Excel.Application
    appExcel.Visible = False
    Set wbLibro = appExcel.Workbooks.Add
    wbLibro.SaveAs m_URLExcel
    Set wbHoja = wbLibro.Worksheets(1)
    intFila = 1
    columna = 0
    'CABECERA
    With wbHoja
        For Each m_Campo In m_ColCampos
            columna = columna + 1
            
            
            .Cells(intFila, columna).value = m_Campo
        Next
        intFila = intFila + 1
        For Each m_ID In p_Col
            Set m_Solicitud = p_Col(m_ID)
            Set m_SolicitudFechas = m_Solicitud.SolicitudFechas
            columna = 0
            For Each m_Campo In m_ColCampos
                
                If CStr(m_Campo) = "Expediente" Then
                    If Application.TempVars("ExpedienteUnificado") = "No" Then
                        If Not m_Solicitud.Expediente Is Nothing Then
                            m_Valor = m_Solicitud.Expediente.TextoExpediente
                        End If
                    Else
                        If Not m_Solicitud.Expediente Is Nothing Then
                            m_Valor = m_Solicitud.Expediente.TextoExpediente
                        End If
                    End If
                    
                
                ElseIf CStr(m_Campo) = "EmpresaUsuario" Then
                    If Not m_Solicitud.EmpresaUsuario Is Nothing Then
                        m_Valor = m_Solicitud.EmpresaUsuario.Nombre
                    End If
                ElseIf CStr(m_Campo) = "EmpresaTramitadora" Then
                    If Not m_Solicitud.EmpresaTramitadora Is Nothing Then
                        m_Valor = m_Solicitud.EmpresaTramitadora.Nombre
                    End If

                    
                Else
                    m_Objeto = m_ColCampos(m_Campo)
                    If m_Objeto = "SO" Then
                        m_Valor = m_Solicitud.getPropiedad(m_Campo, p_Error)
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                    Else
                        m_Valor = m_SolicitudFechas.getPropiedad(m_Campo, p_Error)
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                    End If
                End If
               
                
                If InStr(1, m_Campo, "Fecha") <> 0 And IsDate(m_Valor) Then
                    m_Valor = Format(m_Valor, "mm/dd/yyyy")
                End If
                columna = columna + 1
                If m_Valor <> "" Then
                    .Cells(intFila, columna).value = m_Valor
                End If
            Next
            intFila = intFila + 1
            Set m_Solicitud = Nothing
        Next
        
      
   
    End With
    
    AjustarCeldas p_Hoja:=wbHoja, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    ConvertirATabla p_Hoja:=wbHoja, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    For columna = 1 To m_ColCampos.Count
        With wbHoja
            .Columns(columna).EntireColumn.AutoFit
        End With
    Next
    wbLibro.Close True
    Set wbLibro = Nothing
    appExcel.Quit
    Set appExcel = Nothing
    
    
    GenerarConsultaSolicitudes = m_URLExcel
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Consultas.GenerarConsultaSolicitudes ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
     If Not wbLibro Is Nothing Then
        wbLibro.Close False
        Set wbLibro = Nothing
    End If
    If Not appExcel Is Nothing Then
        appExcel.Quit
        Set appExcel = Nothing
    End If
    
End Function
Public Function ConvertirATabla( _
                                        p_Hoja As Excel.Worksheet, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim m_Rango As Excel.Range
    On Error GoTo errores
    Set m_Rango = p_Hoja.Range("A1").CurrentRegion
    p_Hoja.Application.CutCopyMode = False
    p_Hoja.ListObjects.Add(xlSrcRange, m_Rango, , xlYes).Name = "Tabla1"
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ConvertirATabla ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function AjustarCeldas( _
                                p_Hoja As Excel.Worksheet, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    On Error GoTo errores
    
   
    With p_Hoja.Cells
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AjustarCeldas ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Public Function ActualizarEstados(Optional ByRef p_Error As String) As String
    Dim m_Col As Scripting.Dictionary
    Dim m_Solicitud As solicitud
    Dim m_ID As Variant
    Dim m_SQL As String
    Dim m_EstadoCalculado As String
    Dim m_Estado As String
    On Error GoTo errores
    Set m_Col = constructor.getSolicitudes(p_Error:=p_Error)
    If p_Error <> "" Then
        Exit Function
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_Solicitud = m_Col(m_ID)
        m_Estado = m_Solicitud.Estado
        m_EstadoCalculado = m_Solicitud.EstadoCalculadoTexto
'        If m_EstadoCalculado = "" Then
'            Stop
'        End If
        If m_Estado <> m_EstadoCalculado Then
            m_SQL = "UPDATE TbSolicitudes SET Estado = '" & m_EstadoCalculado & "' " & _
                    "WHERE IDSolicitud=" & m_Solicitud.IDSolicitud & ";"
            getdb().Execute m_SQL
            Debug.Print "IDSolicitud: " & m_Solicitud.IDSolicitud & " Estado:" & m_Estado & vbTab & "Estado Calculado: " & m_EstadoCalculado
        Else
            Debug.Print "IDSolicitud: " & m_Solicitud.IDSolicitud & " " & m_Estado
        End If
        Set m_Solicitud = Nothing
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizarEstados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
Public Function DatosConfInsuficientes( _
                                        p_Configuracion As Configuracion, _
                                        Optional p_AlInicio As EnumSiNo = EnumSiNo.Sí, _
                                        Optional ByRef p_Error As String) As Boolean
    
    Dim m_URL As String
    On Error GoTo errores
    
    If p_Configuracion Is Nothing Then
        Set p_Configuracion = constructor.getConfiguracion(p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If p_Configuracion Is Nothing Then
        DatosConfInsuficientes = True
        Exit Function
    End If
    
    With p_Configuracion
        If .CorreosAutomaticos = "" Then
            DatosConfInsuficientes = True
            Exit Function
        End If
        If .AutocancelacionPreMARGA = "" Then
            DatosConfInsuficientes = True
            Exit Function
        End If
        If .AutocancelacionMARGA = "" Then
            DatosConfInsuficientes = True
            Exit Function
        End If
        If .AutocancelacionMARGA = "Sí" Then
            If .DiasCancelacionMARGA = "" Then
                DatosConfInsuficientes = True
                Exit Function
            End If
        End If
        If .AutocancelacionPreMARGA = "Sí" Then
            If .DiasCancelacionPreMARGA = "" Then
                DatosConfInsuficientes = True
                Exit Function
            End If
        End If
        If .CorreosAutomaticos = "No" Then
            DatosConfInsuficientes = False
            Exit Function
        End If
        If m_ObjEntorno.DatosParaCorreoAutomaticoRegistrados <> EnumSiNo.Sí Then
            DatosConfInsuficientes = True
            Exit Function
        End If
        If p_AlInicio = EnumSiNo.Sí Then
            DatosConfInsuficientes = False
            Exit Function
        End If
        If .VersionPlantillasHTML = "" Then
            DatosConfInsuficientes = True
            Exit Function
        End If
        m_URL = m_ObjEntorno.URLDirectorioPlantillasHTML & .VersionPlantillasHTML
        If Not fso.FolderExists(m_URL) Then
            DatosConfInsuficientes = True
            Exit Function
        End If
        m_URL = m_ObjEntorno.URLDirectorioPlantillasExcel & .VersionPlantillasEXCEL
        If Not fso.FolderExists(m_URL) Then
            DatosConfInsuficientes = True
            Exit Function
        End If
        DatosConfInsuficientes = False
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DatosConfInsuficientes ha devuelto el error: " & vbNewLine & Err.Description
    End If
   
End Function

Public Function AsociarUsuarioHPSConSolicitud( _
                                            p_ID As String, _
                                            p_IDSolicitud As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_SQL As String
    On Error GoTo errores
    If p_ID = "" Or p_IDSolicitud = "" Then
        p_Error = "Falta IDUsuario o IDSolicitud"
        Err.Raise 1000
    End If
    
    m_SQL = "UPDATE TbUsuarios SET IDSolicitud = " & p_IDSolicitud & " " & _
            "WHERE ID=" & p_ID & ";"
    getdbHPS().Execute m_SQL
    m_SQL = "UPDATE TbUsuariosEntidades SET IDSolicitud = " & p_IDSolicitud & " " & _
            "WHERE ID=" & p_ID & ";"
    getdbHPS().Execute m_SQL
    
    
   
     
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AsociarUsuarioHPSConSolicitud ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function CorreoLimpio( _
                                p_CorreoInicial As String, _
                                Optional ByRef p_Error As String) As String
                            
    Dim m_Correo As Variant
    Dim m_Resultado As String
    Dim dato As Variant
    
    On Error GoTo errores
    If p_CorreoInicial = "" Then
        Exit Function
    End If
    dato = Split(p_CorreoInicial, ";")
    For Each m_Correo In dato
        If m_Resultado = "" Then
            m_Resultado = m_Correo
        Else
            If InStr(1, m_Resultado, m_Correo) = 0 Then
                m_Resultado = m_Resultado & ";" & m_Correo
            End If
        End If
    Next
    CorreoLimpio = m_Resultado
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoLimpio ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
' [AÑADIR EN FUNCIONES UTILES.bas]
' Sincroniza la carpeta de imágenes de la plantilla al directorio local para la previsualización
Private Sub SincronizarImagenesLocal(Optional ByRef p_Error As String)
    Dim m_RutaOrigenIMG As String
    Dim m_RutaDestinoIMG As String
    
    On Error GoTo errores
    
    ' 1. Calculamos la ruta origen (Donde están las plantillas HTML + la carpeta IMG)
    ' Usamos la propiedad del Entorno que ya apunta a la versión correcta: ...\HTML\Vxx\
    m_RutaOrigenIMG = m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida & "IMG"
    
    ' 2. Calculamos la ruta destino (Tu carpeta local temporal)
    m_RutaDestinoIMG = m_ObjEntorno.URLDirectorioLocal & "IMG"
    
    ' 3. Validamos que exista el origen
    If fso.FolderExists(m_RutaOrigenIMG) Then
        ' Si no existe la carpeta destino, la creamos, o copiamos el contenido
        ' fso.CopyFolder sobreescribe si se pone True
        On Error Resume Next ' Prevenimos error si hay archivos abiertos/bloqueados
        fso.CopyFolder m_RutaOrigenIMG, m_RutaDestinoIMG, True
        On Error GoTo errores
    End If
    
    Exit Sub
errores:
    ' No detenemos el proceso si falla la copia de imágenes, pero lo notificamos en debug
    Debug.Print "Error al sincronizar imágenes locales: " & Err.Description
    ' p_Error = "Error copiando recursos gráficos: " & Err.Description
End Sub

Public Sub ResetAvanceCounter()
    ' Responsabilidad: Pone a cero el contador de pasos del splash.
    ' Debe ser llamado desde el evento Form_Open del frmSplash.
    s_contadorPasos = 0
End Sub

Public Sub GestionarRibbon(ByVal mostrar As Boolean)
    ' RESPONSABILIDAD: Muestra u oculta la cinta de opciones de Access.
    On Error Resume Next ' Si hay algún problema, no debe detener el arranque.
    
    If mostrar Then
        DoCmd.ShowToolbar "Ribbon", acToolbarYes
    Else
        DoCmd.ShowToolbar "Ribbon", acToolbarNo
    End If
End Sub
