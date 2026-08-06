Option Compare Database
Option Explicit

Private m_SolicitudOrigen As solicitud
Private m_Error As String

' =========================================================================================
' INICIALIZACIÓN Y CARGA DE COMBOS
' =========================================================================================
Private Sub Form_Load()
    Dim m_SQL As String
    Dim rcd As DAO.Recordset
    Dim key As Variant
    
    On Error GoTo errores
    
    Ajustar Me
    Me.Caption = "CONSOLA DE ENVÍO DE CORREOS DE PRUEBA"
    
    ' -------------------------------------------------------------------------
    ' 1. Cargar Combo de Solicitudes (Modo Desconectado / Value List)
    ' -------------------------------------------------------------------------
    With Me.cboSolicitudOrigen
        .RowSourceType = "Value List"
        .RowSource = ""
        .ColumnCount = 2
        .ColumnWidths = "0cm;10cm" ' Ocultamos ID, mostramos texto
        .BoundColumn = 1
    End With
    
    ' Construimos la consulta igual que en FormCorreosGestion
    ' Traemos Nombre, Apellidos e ID para identificarlo fácilmente
    m_SQL = "SELECT IDSolicitud, " & _
            "Nombre & ' ' & Apellido1 & ' ' & Apellido2 & ' (' & IDSolicitud & ')' AS NombreCompleto " & _
            "FROM TbSolicitudes " & _
            "ORDER BY IDSolicitud DESC;"
            
    ' Usamos el helper getdb() para obtener el recordset
    Set rcd = getdb().OpenRecordset(m_SQL)
    
    If Not rcd.EOF Then
        rcd.MoveFirst
        Do While Not rcd.EOF
            ' Añadimos: ID ; Texto Visible
            Me.cboSolicitudOrigen.AddItem rcd!IDSolicitud & ";" & rcd!NombreCompleto
            rcd.MoveNext
        Loop
    End If
    
    rcd.Close
    Set rcd = Nothing
    
    ' -------------------------------------------------------------------------
    ' 2. Cargar Combo de Tipos de Correo (Desde el Entorno)
    ' -------------------------------------------------------------------------
    Me.cboTipoCorreo.RowSourceType = "Value List"
    
    ' Iteramos sobre la colección de tipos definida en Entorno
    For Each key In m_ObjEntorno.ColTiposCorreos.Keys
        Me.cboTipoCorreo.AddItem key & ";" & m_ObjEntorno.ColTiposCorreos(key)
    Next
    
    ' Valores por defecto
    Me.txtFechaOrden.Enabled = True
    Me.txtFechaOrden = Null ' Vacío por defecto (Envío inmediato)
    
    Exit Sub
errores:
    If Not rcd Is Nothing Then rcd.Close
    MsgBox "Error al cargar formulario: " & Err.Description
End Sub

Private Sub cmdSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

' =========================================================================================
' LÓGICA DE HIDRATACIÓN (CARGA DE DATOS)
' =========================================================================================

' Al cambiar la solicitud o el tipo, recargamos los datos propuestos
Private Sub cboSolicitudOrigen_AfterUpdate()
    CargarDatosSimulados
End Sub

Private Sub cboTipoCorreo_AfterUpdate()
    CargarDatosSimulados
End Sub

Private Sub cmdCargarDatos_Click()
    CargarDatosSimulados
End Sub

Private Sub CargarDatosSimulados()
    Dim idSol As String
    Dim tipoCorreo As EnumTipoEnvioCorreo
    
    On Error GoTo errores
    m_Error = ""
    
    ' Validaciones previas
    If IsNull(Me.cboSolicitudOrigen) Then Exit Sub
    If IsNull(Me.cboTipoCorreo) Then Exit Sub
    
    DoCmd.Hourglass True
    
    idSol = Me.cboSolicitudOrigen.Column(0)
    tipoCorreo = CLng(Me.cboTipoCorreo.Column(0))
    
    ' 1. Obtener la Solicitud Real (Hidratación)
    Set m_SolicitudOrigen = constructor.getSolicitud(idSol, m_Error)
    If m_Error <> "" Then Err.Raise 1000, , m_Error
    
    ' 2. Rellenar los campos del formulario con la lógica de negocio
    ' Aquí usamos los métodos de la clase Solicitud para que calcule como si fuera real
    
    Me.txtDestinatarios = m_SolicitudOrigen.DestinatariosCalculados
    Me.txtCopia = m_SolicitudOrigen.DestinatariosConCopiaCalculados(tipoCorreo)
    Me.txtAsunto = "PRUEBA: " & m_SolicitudOrigen.AsuntoCalculado(tipoCorreo)
    
    ' Nota: No tocamos txtFechaOrden aquí, dejamos que el usuario decida si poner fecha o no
    
    ' Feedback
    Me.lblInfo.Caption = "Datos cargados desde Solicitud " & idSol & " para fase " & Me.cboTipoCorreo.Column(1)
    
    DoCmd.Hourglass False
    Exit Sub
    
errores:
    DoCmd.Hourglass False
    MsgBox "Error cargando datos: " & Err.Description, vbCritical
End Sub

' =========================================================================================
' ACCIÓN DE PREVISUALIZAR
' =========================================================================================
Private Sub cmdVerHTML_Click()
    Dim rutaPlantilla As String
    Dim nombrePlantilla As String
    Dim tipoCorreo As EnumTipoEnvioCorreo
    
    On Error GoTo errores
    
    If m_SolicitudOrigen Is Nothing Then
        MsgBox "Seleccione primero una solicitud origen.", vbExclamation
        Exit Sub
    End If
    
    DoCmd.Hourglass True
    
    ' 1. Determinar plantilla
    tipoCorreo = CLng(Me.cboTipoCorreo.Column(0))
    nombrePlantilla = m_SolicitudOrigen.NombrePlantillaParaCorreoCalculada(tipoCorreo)
    rutaPlantilla = m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida & nombrePlantilla
    
    ' 2. Llamar a la función de visualización local (usando la solicitud cargada para los parámetros)
    ' Nota: Al pasar m_SolicitudOrigen, el generador usará sus datos (||Nombre||, etc.) para el HTML
    ObtenerHTMLEnLocal rutaPlantilla, , m_SolicitudOrigen, m_Error
    
    If m_Error <> "" Then Err.Raise 1000
    
    DoCmd.Hourglass False
    Exit Sub
    
errores:
    DoCmd.Hourglass False
    MsgBox "Error previsualizando: " & m_Error & vbNewLine & Err.Description, vbCritical
End Sub

' =========================================================================================
' ACCIÓN DE ENVIAR (GENERAR EN COLA)
' =========================================================================================
Private Sub cmdEnviarPrueba_Click()
    Dim m_CorreoGen As Correo
    Dim tipoCorreo As EnumTipoEnvioCorreo
    Dim FechaEnvio As Variant
    Dim sRecursos As String
    Dim sPlantilla As String
    Dim sAccion As String
    Dim sAdjunto As String
    
    On Error GoTo errores
    m_Error = ""
    
    ' Validaciones
    If IsNull(Me.txtDestinatarios) Then
        MsgBox "Indique al menos un destinatario.", vbExclamation
        Exit Sub
    End If
    If m_SolicitudOrigen Is Nothing Then
        MsgBox "Debe cargar una solicitud origen primero.", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Se va a generar un registro en TbCorreosEnviados con ID -1 (Prueba)." & vbNewLine & _
              "Esto NO modificará la solicitud original." & vbNewLine & vbNewLine & _
              "¿Proceder?", vbQuestion + vbYesNo, "Confirmar Prueba") = vbNo Then Exit Sub
              
    DoCmd.Hourglass True
    
    ' 1. Activar Flag de Modo Nuevo (Desacoplado)
    m_ObjEntorno.ModoNuevoEnvioCorreo = True
    
    ' 2. Preparar Datos
    tipoCorreo = CLng(Me.cboTipoCorreo.Column(0))
    
    ' Recalcular estos datos desde la solicitud origen por si acaso,
    ' aunque el HTML final dependerá de la cadena de recursos generada
    sRecursos = m_SolicitudOrigen.ColParametrosParaHTMLTexto
    sPlantilla = m_SolicitudOrigen.NombrePlantillaParaCorreoCalculada(tipoCorreo)
    sAccion = "PRUEBA: " & m_SolicitudOrigen.AccionParaCorreoCalculada(tipoCorreo)
    
    ' Adjunto (si aplica)
    If tipoCorreo = EnumTipoEnvioCorreo.EnvioExcel Then
        sAdjunto = m_SolicitudOrigen.URLPlantillaExcel
    End If
    
    ' --- LÓGICA DE FECHA (Si hay fecha en el cuadro de texto, es programado) ---
    If IsDate(Me.txtFechaOrden) Then
        FechaEnvio = CDate(Me.txtFechaOrden)
    Else
        FechaEnvio = Null
    End If
    ' ---------------------------------------------------------------------------
    
    ' 3. Llamada al Generador Genérico (Automiatizacion.bas)
    Set m_CorreoGen = Automiatizacion.EnviarCorreoGenerico( _
        p_DESTINATARIOS:=Me.txtDestinatarios, _
        p_asunto:=Me.txtAsunto, _
        p_NombrePlantilla:=sPlantilla, _
        p_CadenaRecursos:=sRecursos, _
        p_Accion:=sAccion, _
        p_TipoCorreo:=tipoCorreo, _
        p_DestinatariosCopia:=Nz(Me.txtCopia, ""), _
        p_URLAdjunto:=sAdjunto, _
        p_IDSolicitudVinculada:=CONST_IDSOLICITUD_PRUEBA, _
        p_FechaOrden:=FechaEnvio, _
        p_Error:=m_Error)
        
    ' 4. Restaurar Flag (Importante)
    m_ObjEntorno.ModoNuevoEnvioCorreo = False
    
    DoCmd.Hourglass False
    
    If m_Error <> "" Then
        Err.Raise 1000
    Else
        MsgBox "Correo de prueba generado correctamente." & vbNewLine & _
               "ID Correo: " & m_CorreoGen.IDCORREO & vbNewLine & _
               "Estado: " & m_CorreoGen.EstadoTexto, vbInformation
    End If
    
    Exit Sub

errores:
    m_ObjEntorno.ModoNuevoEnvioCorreo = False ' Asegurar que se restaura siempre
    DoCmd.Hourglass False
    MsgBox "Error al enviar prueba: " & m_Error & vbNewLine & Err.Description, vbCritical
End Sub
