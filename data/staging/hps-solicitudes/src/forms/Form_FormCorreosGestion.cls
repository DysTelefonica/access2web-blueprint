Option Compare Database
Option Explicit

Private WithEvents m_Form As Form_FormCorreo
Public m_CorreoSeleccionado As Correo
Private m_CorreoServicio As CorreoServicio
Public Event CorreoModificado()
Private m_Error As String



Private Sub cmdCorreoPruebas_Click()
    DoCmd.OpenForm "FormCorreoPruebas"
End Sub

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


' En Form_FormCorreosGestion.cls

Private Sub ComandoInformeControl_Click()
    Dim m_Cuerpo As String
    Dim m_RutaArchivo As String
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    
    ' 1. Generar el HTML con los datos actuales
    ' Llamamos a la función que creamos en el módulo InformesOperaciones
    m_Cuerpo = InformesOperaciones.GenerarCuerpoInformeControl(m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' 2. Guardar el HTML en un archivo local temporal
    ' Usamos la función existente en FUNCIONES UTILES que gestiona nombres y rutas
    m_RutaArchivo = GenerarArchivoConHTML(m_Cuerpo, m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' 3. Abrir el archivo directamente (Lanza el navegador predeterminado)
    ' Usamos la API 'Ejecutar' declarada en Variables Globales
    Ejecutar Me.hwnd, "open", m_RutaArchivo, "", "", 1
    
    ' Log opcional: Dejar constancia de que se consultó (sin destinatarios porque es local)
    RegistrarLogGeneral "Consulta Informe Control", "Visualización en pantalla por " & m_ObjUsuarioConectado.UsuarioRed, Nothing, m_Error
    
    DoCmd.Hourglass False
    Exit Sub
    
errores:
    DoCmd.Hourglass False
    MsgBox "Error al generar el informe: " & m_Error & vbNewLine & Err.Description, vbCritical
End Sub

Private Sub ComandoLimpiarEstado_Click()
    If Me.Estado = "Todos" Then Exit Sub
    Me.Estado = "Todos"
    Filtrar
End Sub

Private Sub ComandoLimpiarSolicitud_Click()
    If Me.solicitud = "0" Then Exit Sub
    Me.solicitud = "0"
    Filtrar
End Sub

' Eventos de cambio en los combos
Private Sub Estado_AfterUpdate()
    Filtrar
End Sub



Private Sub EstablecerCombos()
    Dim m_SQL As String
    Dim rcd As DAO.Recordset
    
    On Error Resume Next
    
    ' 1. Configurar Combo de ESTADO
    Me.Estado.RowSourceType = "Value List"
    Me.Estado.RowSource = ""
    Me.Estado.AddItem "Todos"
    Me.Estado.AddItem "Enviado"
    Me.Estado.AddItem "Pendiente"
    Me.Estado.AddItem "Programado"
    Me.Estado.value = "Todos"
    
    ' 2. Configurar Combo de SOLICITUD
    ' ---------------------------------------------------------
    ' CONFIGURACIÓN VISUAL CRÍTICA PARA OCULTAR ID
    ' ---------------------------------------------------------
    With Me.solicitud
        .RowSourceType = "Value List"
        .RowSource = ""
        .ColumnCount = 2          ' Tiene 2 columnas: ID y Texto
        .ColumnWidths = "0cm;10cm" ' 0cm oculta la primera (ID), 10cm muestra la segunda (Nombre)
        .BoundColumn = 1          ' El valor real sigue siendo el ID (Columna 1)
    End With
    
    ' Añadir opción por defecto
    ' Nota: El espacio antes de "Todas" es un truco para que salga arriba si se ordenara por texto,
    ' pero aquí controlamos nosotros el orden de inserción.
    Me.solicitud.AddItem "0; Todas las solicitudes"
    
    ' 3. Obtener datos ordenados
    ' Ordenamos por Nombre y Apellidos, y si hay repes, el ID más alto (reciente) primero
    m_SQL = "SELECT IDSolicitud, " & _
            "Nombre & ' ' & Apellido1 & ' ' & Apellido2 & ' (' & IDSolicitud & ')' AS NombreCompleto " & _
            "FROM TbSolicitudes " & _
            "ORDER BY Nombre ASC, Apellido1 ASC, Apellido2 ASC, IDSolicitud DESC;"
            
    Set rcd = getdb().OpenRecordset(m_SQL)
    
    If Not rcd.EOF Then
        rcd.MoveFirst
        Do While Not rcd.EOF
            ' Añadimos: ID (Oculto) ; Texto Visible
            Me.solicitud.AddItem rcd!IDSolicitud & ";" & rcd!NombreCompleto
            rcd.MoveNext
        Loop
    End If
    
    rcd.Close
    Set rcd = Nothing
    
    ' Seleccionar "Todas" por defecto
    Me.solicitud.value = "0"
End Sub

Public Function Filtrar(Optional ByRef p_Error As String) As String
    Dim m_Col As Scripting.Dictionary
    Dim m_Correo As Correo
    Dim m_ID As Variant
    Dim lst As ListBox
    Dim m_EstadoSel As String
    Dim m_SolicitudSel As String
    Dim m_NombreSolicitud As String
    
    On Error GoTo errores
    
    DoCmd.Hourglass True
    
    Set lst = Me.ListaFiltrados
    lst.RowSource = ""
    
    ' AÑADIR CABECERAS (Si la propiedad "Encabezados de columna" es Sí en diseño)
    ' Las columnas deben ser: IDCorreo; Estado; Solicitud; Asunto/Accion; F.Orden; F.Envio
    lst.AddItem "ID;Estado;Solicitud;Acción;F. Programada;F. Envío"
    
    ' Recoger valores de la UI
    m_EstadoSel = Nz(Me.Estado, "Todos")
    If m_EstadoSel = "Todos" Then m_EstadoSel = ""
    
    m_SolicitudSel = Nz(Me.solicitud, "0")
    
    ' LLAMADA AL SERVICIO
    If m_CorreoServicio Is Nothing Then Set m_CorreoServicio = New CorreoServicio
    Set m_Col = m_CorreoServicio.ObtenerCorreosGestion(m_EstadoSel, m_SolicitudSel, p_Error)
    
    If p_Error <> "" Then
        MsgBox p_Error, vbCritical
        DoCmd.Hourglass False
        Exit Function
    End If
    
    ' PINTAR RESULTADOS
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_Correo = m_Col(m_ID)
            
            ' Obtener nombre bonito de la solicitud si existe
            m_NombreSolicitud = "Desconocido"
            If Not m_Correo.solicitud Is Nothing Then
                m_NombreSolicitud = m_Correo.solicitud.NombreCompletoCapitalizado & " (" & m_Correo.IDSolicitud & ")"
            Else
                m_NombreSolicitud = "Sol. ID: " & m_Correo.IDSolicitud
            End If
            
            ' Construir la línea para el ListBox con el orden solicitado
            ' IDCorreo (Oculto normalmente); ESTADO; SOLICITUD; ACCION; FECHAS
            lst.AddItem m_Correo.IDCORREO & ";" & _
                        m_Correo.EstadoTexto & ";" & _
                        m_NombreSolicitud & ";" & _
                        m_Correo.Accion & ";" & _
                        m_Correo.FechaOrdenEnvio & ";" & _
                        m_Correo.FechaEnvio
        Next
    End If
    
    ' Restamos 1 al conteo por la fila de cabecera
    Dim total As Long
    total = lst.ListCount - 1
    If total < 0 Then total = 0
    
    Me.lblListaDeResultados.Caption = "Registros filtrados (" & total & ")"
    Me.ComandoEditar.Enabled = False
    Me.ComandoEliminar.Enabled = False
    DoCmd.Hourglass False
    
    Exit Function
errores:
    DoCmd.Hourglass False
    MsgBox "Error UI en Filtrar: " & Err.Description
End Function

Private Sub Solicitud_AfterUpdate()
    Filtrar
End Sub


Private Sub ComandoActualizar_Click()

    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
   
    Filtrar
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.AllowEdits = False
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub ComandoEditar_Click()
    
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    If m_CorreoSeleccionado Is Nothing Then
        m_Error = "Seleccione un Correo de la lista"
        Err.Raise 1000
    End If
    
    If FormularioAbierto("FormCorreo") Then
        DoCmd.Close acForm, "FormCorreo", acSaveNo
    End If
    
    ' Pasamos el objeto seleccionado a la variable global
    Set m_ObjCorreoActivo = m_CorreoSeleccionado
    
    ' ABRIMOS EL FORMULARIO PASANDO LA "CONTRASEÑA" EN OpenArgs
    DoCmd.OpenForm "FormCorreo", OpenArgs:="MODO_GESTION"
    
    If FormularioAbierto("FormCorreo") Then
        Set m_Form = Forms("FormCorreo")
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    Me.AllowEdits = True
    AvanceCerrar
    DoCmd.Hourglass False
    MsgBox m_Error, vbCritical, "Error"
End Sub

Private Sub ComandoEliminar_Click()
     
    Dim m_SolOp As CorreoOperaciones
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
   
    m_Error = ""
    Me.AllowEdits = False
   
    ' 1. Validar que haya selección
    If m_CorreoSeleccionado Is Nothing Then
        m_Error = "Seleccione un Correo de la lista"
        Err.Raise 1000
    End If
    
    ' 2. NUEVA VALIDACIÓN: Impedir borrar si ya fue enviado
    If IsDate(m_CorreoSeleccionado.FechaEnvio) Then
        Me.AllowEdits = True
        DoCmd.Hourglass False
        MsgBox "No es posible eliminar un correo que ya ha sido enviado." & vbNewLine & _
               "Solo se pueden eliminar correos pendientes o programados.", vbExclamation, "Acción denegada"
        Exit Sub
    End If

    Me.AllowEdits = True
    DoCmd.Hourglass False
    
    ' 3. Confirmación
    pregunta = MsgBox("¿Desea realmente eliminar definitivamente el Correo Seleccionado?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Correo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    
    ' 4. Ejecución del borrado
    Set m_SolOp = New CorreoOperaciones
    With m_SolOp
        Set .Correo = m_CorreoSeleccionado
        .Eliminar p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    ' 5. Refresco
    ComandoActualizar_Click
    Filtrar p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.ComandoEditar.Enabled = False
    Me.ComandoEliminar.Enabled = False
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    'On Error Resume Next
    Me.AllowEdits = True
    AvanceCerrar
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Estado_BeforeUpdate(Cancel As Integer)
    Filtrar
End Sub

Private Sub ListaFiltrados_DblClick(Cancel As Integer)
    On Error Resume Next
    Me.AllowEdits = False
    If Me.ComandoEditar.Enabled = True Then
        ComandoEditar_Click
    End If

    Me.AllowEdits = True
End Sub




Private Sub Form_Open(Cancel As Integer)
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""

    ' 1. Ajustes visuales iniciales
    Ajustar Me
    Me.Caption = m_TituloFormulario
    ' Si tienes una etiqueta lblTitulo oculta que usas para el caption real:
    If Me.lblTitulo.Caption <> "" Then Me.Caption = Me.lblTitulo.Caption
    
    Me.AllowEdits = False
    
    ' 2. Inicializar el Servicio (Capa de Negocio)
    Set m_CorreoServicio = New CorreoServicio
    
    ' 3. Rellenar Combos (Capa de Presentación)
    ' IMPORTANTE: Hacemos esto antes de filtrar para tener valores por defecto
    EstablecerCombos
    
    ' 4. Cargar datos iniciales
    Filtrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' 5. Finalizar carga
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    'On Error Resume Next
    Me.AllowEdits = True
    AvanceCerrar
    DoCmd.Hourglass False
    MsgBox m_Error, vbCritical, "Error"
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ListaFiltrados_Click()

    Dim m_SOl As Correo
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""


    Me.ComandoEditar.Enabled = False
    Me.ComandoEliminar.Enabled = False
    
    Me.AllowEdits = False
    If Nz(Me.ListaFiltrados.Column(0), "") = "" Then
        Set m_CorreoSeleccionado = Nothing
        
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If m_CorreoSeleccionado Is Nothing Then
        Set m_CorreoSeleccionado = constructor.getCorreo(p_IDCorreo:=Me.ListaFiltrados.Column(0), p_Error:=m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If Me.ListaFiltrados.Column(0) <> m_CorreoSeleccionado.IDCORREO Then
            Set m_CorreoSeleccionado = constructor.getCorreo(p_IDCorreo:=Me.ListaFiltrados.Column(0), p_Error:=m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If

        End If
    End If
    
    
    If m_CorreoSeleccionado Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    Me.ComandoEditar.Enabled = True
   
    
    ' Lógica visual para el botón eliminar
    If IsDate(m_CorreoSeleccionado.FechaEnvio) Then
        Me.ComandoEliminar.Enabled = False ' Desactivar si ya se envió
    Else
        Me.ComandoEliminar.Enabled = True  ' Activar si está pendiente
    End If
    Me.AllowEdits = True
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
   
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaFiltrados_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")

End Sub

Private Sub m_Form_Alta()
    Filtrar
End Sub

Private Sub m_Form_Editado()
    Filtrar
    
End Sub