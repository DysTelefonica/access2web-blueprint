Option Compare Database
Option Explicit
Private m_Error As String
Private m_NombreFormulario As String
Private m_ObjAplicacionAlInicio As Aplicacion
Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean


    On Error GoTo errores

    With m_ObjAplicacionAlInicio
        If .NombreAplicacion <> Nz(Me.NombreAplicacion, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .NombreCorto <> Nz(Me.NombreCorto, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .NombreEjecutable <> Nz(Me.NombreEjecutable, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Pass <> Nz(Pass, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .EjecucionEnOficina <> Nz(Me.EjecucionEnOficina, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .EnPruebas <> Nz(Me.EnPruebas, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .NombreIconoLanzadera <> Nz(Me.NombreIconoLanzadera, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .NombreArchivoDatos <> Nz(Me.NombreArchivoDatos, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .NombreIcono <> Nz(Me.NombreIcono, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .NombreIconoLanzadera <> Nz(Me.NombreIconoLanzadera, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .NombreCarpeta <> Nz(Me.NombreCarpeta, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .NombreCarpetaTemporal <> Nz(Me.NombreCarpetaTemporal, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .NombreCarpetaDocumentacion <> Nz(Me.NombreCarpetaDocumentacion, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .NombreDirectorioIconos <> Nz(Me.NombreDirectorioIconos, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .NombreDirectorioAyuda <> Nz(Me.NombreDirectorioAyuda, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .NombreDirectorioRecursos <> Nz(Me.NombreDirectorioRecursos, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If

End Function
Private Function EstablecerDatos(Optional ByRef p_Error As String) As String


    On Error GoTo errores


    If m_ObjAplicacionActiva Is Nothing Then
        Set m_ObjAplicacionActiva = New Aplicacion
        Exit Function
    End If
    With m_ObjAplicacionActiva
        If .NombreAplicacion <> "" Then
            Me.NombreAplicacion = .NombreAplicacion
        End If
        If .NombreCorto <> "" Then
            Me.NombreCorto = .NombreCorto
        End If
        If .NombreEjecutable <> "" Then
            Me.NombreEjecutable = .NombreEjecutable
        End If
        If .Pass <> "" Then
            Me.Pass = .Pass
        End If
        If .EjecucionEnOficina <> "" Then
            Me.EjecucionEnOficina = .EjecucionEnOficina
        End If
        If .EnPruebas <> "" Then
            Me.EnPruebas = .EnPruebas
        End If
        
        If .ConIconoEnLanzadera <> "" Then
            Me.ConIconoEnLanzadera = .ConIconoEnLanzadera
        End If
        If .NombreArchivoDatos <> "" Then
            Me.NombreArchivoDatos = .NombreArchivoDatos
        End If
        If .NombreIcono <> "" Then
            Me.NombreIcono = .NombreIcono
        End If
        If .NombreIconoParaArbol <> "" Then
            Me.NombreIconoParaArbol = .NombreIconoParaArbol
        End If
        
        If .NombreIconoLanzadera <> "" Then
            Me.NombreIconoLanzadera = .NombreIconoLanzadera
        End If
        If .NombreCarpeta <> "" Then
            Me.NombreCarpeta = .NombreCarpeta
        End If
        
        If .NombreCarpetaTemporal <> "" Then
            Me.NombreCarpetaTemporal = .NombreCarpetaTemporal
        End If
        If .NombreCarpetaDocumentacion <> "" Then
            Me.NombreCarpetaDocumentacion = .NombreCarpetaDocumentacion
        End If
        If .NombreDirectorioIconos <> "" Then
            Me.NombreDirectorioIconos = .NombreDirectorioIconos
        End If
        If .NombreDirectorioAyuda <> "" Then
            Me.NombreDirectorioAyuda = .NombreDirectorioAyuda
        End If
        If .NombreDirectorioRecursos <> "" Then
            Me.NombreDirectorioRecursos = .NombreDirectorioRecursos
        End If
        If .IDAplicacion <> "" Then
            Me.IDAplicacion = .IDAplicacion
            Set m_ObjAplicacionAlInicio = Constructor.getAplicacion(m_ObjAplicacionActiva.IDAplicacion, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End With

    If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado = EnumSino.Sí Then
        Me.ComandoRegistrar.Enabled = True
        Me.AllowEdits = True
    Else
        Me.ComandoRegistrar.Enabled = False
        Me.AllowEdits = False
    End If


    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If

End Function

Private Sub ComandoRegistrar_Click()

    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Not m_ObjAplicacionAlInicio Is Nothing Then
        If Not HaHabidoCambios(m_Error) Then
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            m_Error = "Ningún cambio que registrar"
            Err.Raise 1000
        End If
    End If
    If m_ObjAplicacionActiva Is Nothing Then
        Set m_ObjAplicacionActiva = New Aplicacion
    End If
    With m_ObjAplicacionActiva
        .NombreAplicacion = Nz(Me.NombreAplicacion, "")
        .NombreCorto = Nz(Me.NombreCorto, "")
        .NombreEjecutable = Nz(Me.NombreEjecutable, "")
        .Pass = Nz(Me.Pass, "")
        .EjecucionEnOficina = Nz(Me.EjecucionEnOficina, "")
        .EnPruebas = Nz(Me.EnPruebas, "")
        .ConIconoEnLanzadera = Nz(Me.ConIconoEnLanzadera, "")
        .NombreArchivoDatos = Nz(Me.NombreArchivoDatos, "")
        .NombreIcono = Nz(Me.NombreIcono, "")
        .NombreIconoParaArbol = Nz(Me.NombreIconoParaArbol, "")
        .NombreIconoLanzadera = Nz(Me.NombreIconoLanzadera, "")
        .NombreCarpeta = Nz(Me.NombreCarpeta, "")
        .NombreCarpetaTemporal = Nz(Me.NombreCarpetaTemporal, "")
        .NombreCarpetaDocumentacion = Nz(Me.NombreCarpetaDocumentacion, "")
        .NombreDirectorioIconos = Nz(Me.NombreDirectorioIconos, "")
        .NombreDirectorioAyuda = Nz(Me.NombreDirectorioAyuda, "")
        .NombreDirectorioRecursos = Nz(Me.NombreDirectorioRecursos, "")
        .Registrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With

    Set m_ObjAplicacionAlInicio = Constructor.getAplicacion(m_ObjAplicacionActiva.IDAplicacion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.IDAplicacion = m_ObjAplicacionAlInicio.IDAplicacion


    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    m_NombreFormulario = Me.Name
    

    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
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
    DoCmd.Close acForm, m_NombreFormulario, acSaveNo
End Sub