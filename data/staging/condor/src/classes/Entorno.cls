VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Entorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public URLArchivoUltimo As String

' --- Propiedades privadas del módulo ---
Private m_sCSS As String
Private m_sURLAchivoIni As String
Private m_sURLArchivoCSS As String
Private m_sURLDirectorioAplicaciones As String
Private m_sURLDirectorioAplicacion As String
Private m_sURLDirectorioDocumentacion As String
Private m_sURLDirectorioDocumentacionAyuda As String
Private m_sURLDirectorioLocal As String
Private m_sURLPlantillaCDCA As String
Private m_sURLPlantillaCDCASUB As String
Private m_sURLPlantillaPC As String
Private m_sURLPlantillaPCSUB As String
Private m_sVersionAplicacion As String
Private m_sURLVideo As String
Private m_sURLAyudaWorkflow As String
Private m_sURLManual As String
Private m_sTituloUsuarioConectado As String
Private m_sNombreJefeCalidad As String
Private m_sDirUTE As String

' --- Variables para HTMLs cacheados (Spec-149) ---
Private m_sURLDirectorioHTML As String
Private m_sURLHTMLCapturaFechaRAC As String
Private m_sURLHTMLCapturaComentariosRAC As String
Private m_sURLHTMLCapturaRechazoCalidad As String

' --- Variables de Caché ---
Private m_colResponsablesTecnicos As Scripting.Dictionary
Private m_colResponsablesCalidad As Scripting.Dictionary
Private m_colSuministradores As Scripting.Dictionary
Private m_colEstados As Scripting.Dictionary

' --- Password de Base de Datos (Spec-PWD-01) ---
Private m_sPasswordDB As String
Private m_objColItems As Scripting.Dictionary
Public Property Get PasswordDB() As String
    On Error GoTo Errores
    
    If m_sPasswordDB = "" Then
        ' Lee desde .ini; fallback legacy "dpddpd" si no existe la key
        m_sPasswordDB = LeerIni("PasswordDB", "dpddpd")
    End If
    PasswordDB = m_sPasswordDB
    Exit Property
    
Errores:
    ' Fallback hardcoded: si .ini falla, usamos legacy
    PasswordDB = "dpddpd"
End Property

' --- Propiedad para la colección de items (uso interno para EVE) ---

Public Property Get DirUTE() As String
    On Error GoTo Errores
    
    If m_sDirUTE = "" Then
        ' Leemos el valor desde el archivo .ini
        m_sDirUTE = LeerIni("DirUTE", "Calle Julián Camarillo 8, 4º planta, 28037 Madrid") ' Valor por defecto
    End If
    DirUTE = m_sDirUTE
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.DirUTE"
    errObj.Raise
End Property
' ==========================================================================
' PROPIEDADES DE CACHÉ (Lazy Loading)
' ==========================================================================
Public Property Get NombreJefeCalidad() As String
    On Error GoTo Errores
    
    If m_sNombreJefeCalidad = "" Then
        ' Leemos el valor desde el archivo .ini
        m_sNombreJefeCalidad = LeerIni("NombreJefeCalidad", "Ana Rubio Canales") ' Valor por defecto por si falla
    End If
    NombreJefeCalidad = m_sNombreJefeCalidad
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.NombreJefeCalidad"
    errObj.Raise
End Property
Public Property Get estados() As Scripting.Dictionary
    On Error GoTo Errores
    
    If m_colEstados Is Nothing Then
        Debug.Print "Caché de Estados vacía. Cargando desde la Base de Datos..."
        Set m_colEstados = EstadoRepositorio.getTodosLosEstados()
    End If
    
    Set estados = m_colEstados
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.Estados"
    errObj.Raise
End Property

Public Property Get ResponsablesTecnicos() As Scripting.Dictionary
    Dim UsuarioServicio As New UsuarioServicio
    On Error GoTo Errores
    
    If m_colResponsablesTecnicos Is Nothing Then
        Debug.Print "Caché de Técnicos vacía. Cargando desde la Base de Datos..."
        Set m_colResponsablesTecnicos = UsuarioServicio.getResponsablesTecnicos()
        Debug.Print "Técnicos cargados en caché: " & m_colResponsablesTecnicos.count & " usuarios."
    End If
    
    Set ResponsablesTecnicos = m_colResponsablesTecnicos
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.ResponsablesTecnicos"
    errObj.Raise
End Property

Public Property Get ResponsablesCalidad() As Scripting.Dictionary
    Dim UsuarioServicio As New UsuarioServicio
    On Error GoTo Errores
    
    If m_colResponsablesCalidad Is Nothing Then
        Debug.Print "Caché de Calidad vacía. Cargando desde la Base de Datos..."
        Set m_colResponsablesCalidad = UsuarioServicio.getResponsablesCalidad()
    End If
    
    Set ResponsablesCalidad = m_colResponsablesCalidad
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.ResponsablesCalidad"
    errObj.Raise
End Property

Public Property Get Suministradores() As Scripting.Dictionary
    Dim gestorSuministradores As New SuministradorServicio
    On Error GoTo Errores
    
    If m_colSuministradores Is Nothing Then
        Debug.Print "Caché de Suministradores vacía. Cargando desde la Base de Datos..."
        Set m_colSuministradores = gestorSuministradores.getSuministradores("")
    End If
    
    Set Suministradores = m_colSuministradores
    Exit Property
    
Errores:
    ' En caso de error, devolvemos una colección vacía para evitar fallos en los consumidores
    Set Suministradores = New Scripting.Dictionary
End Property

' ==========================================================================
' MÉTODOS PARA INVALIDAR CACHÉ
' ==========================================================================

Public Sub RefreshEstados()
    Set m_colEstados = Nothing
End Sub

Public Sub RefreshResponsablesTecnicos()
    Set m_colResponsablesTecnicos = Nothing
End Sub

Public Sub RefreshResponsablesCalidad()
    Set m_colResponsablesCalidad = Nothing
End Sub

Public Sub RefreshSuministradores()
    Set m_colSuministradores = Nothing
End Sub

' ==========================================================================
' PROPIEDADES DE CONFIGURACIÓN Y RUTAS (Lazy Loading)
' ==========================================================================

Public Property Get URLAchivoIni() As String
    On Error GoTo Errores
    
    If m_sURLAchivoIni = "" Then
        m_sURLAchivoIni = Application.CurrentProject.path & "\" & fso.GetBaseName(Application.CurrentProject.name) & ".ini"
        If Not fso.FileExists(m_sURLAchivoIni) Then
            Err.Raise 513, "Entorno.URLAchivoIni", "No se encuentra el archivo .ini en: " & m_sURLAchivoIni
        End If
    End If
    URLAchivoIni = m_sURLAchivoIni
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.URLAchivoIni"
    errObj.Raise
End Property

Public Property Get VersionAplicacion() As String
    On Error GoTo Errores
    
    If m_sVersionAplicacion = "" Then
        m_sVersionAplicacion = LeerIni("VersionAplicacion", "#ERR#")
    End If
    VersionAplicacion = m_sVersionAplicacion
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.VersionAplicacion"
    errObj.Raise
End Property

Public Property Get URLDirectorioAplicaciones() As String
    On Error GoTo Errores

    If m_sURLDirectorioAplicaciones = "" Then
        Select Case Application.TempVars("DatosEnLocal")
            Case "Sí": m_sURLDirectorioAplicaciones = m_URLRutaAplicacionesLocal
            Case "No": m_sURLDirectorioAplicaciones = m_URLRutaAplicacionesRemotas
            Case Else: Err.Raise 513, "Entorno.URLDirectorioAplicaciones", "No se ha definido si el origen de datos es Local o Remoto."
        End Select
        
        If Not fso.FolderExists(m_sURLDirectorioAplicaciones) Then
            Err.Raise 513, "Entorno.URLDirectorioAplicaciones", "No se puede acceder a la ruta de aplicaciones: " & m_sURLDirectorioAplicaciones
        End If
    End If
    
    URLDirectorioAplicaciones = m_sURLDirectorioAplicaciones
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.URLDirectorioAplicaciones"
    errObj.Raise
End Property

Public Property Get URLDirectorioAplicacion() As String
    On Error GoTo Errores

    If m_sURLDirectorioAplicacion = "" Then
        If Application.TempVars("DatosEnLocal") = "Sí" Then
            m_sURLDirectorioAplicacion = m_URLRutaAplicacionLocal
        Else
            m_sURLDirectorioAplicacion = m_URLRutaAplicacionRemota
        End If
        
        If Not fso.FolderExists(m_sURLDirectorioAplicacion) Then
            Err.Raise 513, "Entorno.URLDirectorioAplicacion", "No se puede acceder a la ruta de la aplicación: " & m_sURLDirectorioAplicacion
        End If
    End If
    
    URLDirectorioAplicacion = m_sURLDirectorioAplicacion
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.URLDirectorioAplicacion"
    errObj.Raise
End Property

Public Property Get URLArchivoCSS() As String
    On Error GoTo Errores

    If m_sURLArchivoCSS = "" Then
        If fso.FileExists(Me.URLDirectorioAplicacion & "recursos\CSS.txt") Then
            m_sURLArchivoCSS = Me.URLDirectorioAplicacion & "recursos\CSS.txt"
        Else
            m_sURLArchivoCSS = Me.URLDirectorioAplicaciones & "CSS.txt"
        End If
        
        If Not fso.FileExists(m_sURLArchivoCSS) Then
            Err.Raise 513, "Entorno.URLArchivoCSS", "No se encuentra el archivo CSS en: " & m_sURLArchivoCSS
        End If
    End If
    URLArchivoCSS = m_sURLArchivoCSS
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.URLArchivoCSS"
    errObj.Raise
End Property

Public Property Get css() As String
    Dim objDatos As Object ' TextStream
    On Error GoTo Errores

    If m_sCSS = "" Then
        Set objDatos = fso.OpenTextFile(Me.URLArchivoCSS, 1) ' ForReading
        m_sCSS = objDatos.ReadAll
        objDatos.Close
    End If
    css = m_sCSS
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.CSS"
    errObj.Raise
End Property

Public Property Get URLDirectorioDocumentacion() As String
    On Error GoTo Errores

    If m_sURLDirectorioDocumentacion = "" Then
        m_sURLDirectorioDocumentacion = Me.URLDirectorioAplicacion & "ANEXOS\"
        If Not fso.FolderExists(m_sURLDirectorioDocumentacion) Then
            Err.Raise 513, "Entorno.URLDirectorioDocumentacion", "No se encuentra la carpeta de documentación: " & m_sURLDirectorioDocumentacion
        End If
    End If
    URLDirectorioDocumentacion = m_sURLDirectorioDocumentacion
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.URLDirectorioDocumentacion"
    errObj.Raise
End Property

Public Property Get URLDirectorioDocumentacionAyuda() As String
    On Error GoTo Errores
   
    If m_sURLDirectorioDocumentacionAyuda = "" Then
        m_sURLDirectorioDocumentacionAyuda = Me.URLDirectorioAplicacion & "Doc Ayuda\"
        If Not fso.FolderExists(m_sURLDirectorioDocumentacionAyuda) Then
            Err.Raise 513, "Entorno.URLDirectorioDocumentacionAyuda", "No se encuentra la carpeta de ayuda: " & m_sURLDirectorioDocumentacionAyuda
        End If
    End If
    URLDirectorioDocumentacionAyuda = m_sURLDirectorioDocumentacionAyuda
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.URLDirectorioDocumentacionAyuda"
    errObj.Raise
End Property

Public Property Get URLDirectorioLocal() As String
    Dim m_URL As String
    On Error GoTo Errores

    If m_sURLDirectorioLocal = "" Then
        m_URL = Environ("APPDATA") & "\Aplicaciones DYSN\"
        If Not fso.FolderExists(m_URL) Then fso.CreateFolder m_URL
        
        m_sURLDirectorioLocal = m_URL & "CONDOR\"
        If Not fso.FolderExists(m_sURLDirectorioLocal) Then fso.CreateFolder m_sURLDirectorioLocal
        
        m_sURLDirectorioLocal = m_sURLDirectorioLocal & "docs\"
        
        If fso.FolderExists(m_sURLDirectorioLocal) Then
            On Error Resume Next
            EliminarContenidoCarpeta m_sURLDirectorioLocal
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo Errores
        Else
            fso.CreateFolder m_sURLDirectorioLocal
        End If
        
        If Right(m_sURLDirectorioLocal, 1) <> "\" Then m_sURLDirectorioLocal = m_sURLDirectorioLocal & "\"
    End If
    URLDirectorioLocal = m_sURLDirectorioLocal
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.URLDirectorioLocal"
    errObj.Raise
End Property

Public Property Get URLPlantillaCDCA() As String
    On Error GoTo Errores
    
    If m_sURLPlantillaCDCA = "" Then
        m_sURLPlantillaCDCA = Me.URLDirectorioAplicacion & "recursos\Plantillas\CD_CA.docx"
        If Not fso.FileExists(m_sURLPlantillaCDCA) Then
            Err.Raise 513, "Entorno.URLPlantillaCDCA", "No se encuentra la plantilla CDCA: " & m_sURLPlantillaCDCA
        End If
    End If
    URLPlantillaCDCA = m_sURLPlantillaCDCA
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.URLPlantillaCDCA"
    errObj.Raise
End Property

Public Property Get URLPlantillaCDCASUB() As String
    On Error GoTo Errores
    
    If m_sURLPlantillaCDCASUB = "" Then
        m_sURLPlantillaCDCASUB = Me.URLDirectorioAplicacion & "recursos\Plantillas\CD_CA_SUB.docx"
        If Not fso.FileExists(m_sURLPlantillaCDCASUB) Then
            Err.Raise 513, "Entorno.URLPlantillaCDCASUB", "No se encuentra la plantilla CDCASUB: " & m_sURLPlantillaCDCASUB
        End If
    End If
    URLPlantillaCDCASUB = m_sURLPlantillaCDCASUB
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.URLPlantillaCDCASUB"
    errObj.Raise
End Property

Public Property Get URLPlantillaPC() As String
    On Error GoTo Errores

    If m_sURLPlantillaPC = "" Then
        m_sURLPlantillaPC = Me.URLDirectorioAplicacion & "recursos\Plantillas\PC.docx"
        If Not fso.FileExists(m_sURLPlantillaPC) Then
            Err.Raise 513, "Entorno.URLPlantillaPC", "No se encuentra la plantilla PC: " & m_sURLPlantillaPC
        End If
    End If
    URLPlantillaPC = m_sURLPlantillaPC
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.URLPlantillaPC"
    errObj.Raise
End Property

Public Property Get URLPlantillaPCSUB() As String
    On Error GoTo Errores

    If m_sURLPlantillaPCSUB = "" Then
        m_sURLPlantillaPCSUB = Me.URLDirectorioAplicacion & "recursos\Plantillas\PC_SUB.docx"
        If Not fso.FileExists(m_sURLPlantillaPCSUB) Then
            Err.Raise 513, "Entorno.URLPlantillaPCSUB", "No se encuentra la plantilla PCSUB: " & m_sURLPlantillaPCSUB
        End If
    End If
    URLPlantillaPCSUB = m_sURLPlantillaPCSUB
    Exit Property

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.URLPlantillaPCSUB"
    errObj.Raise
End Property

Public Property Get URLVideo() As String
    On Error GoTo Errores

    If m_sURLVideo = "" Then
        m_sURLVideo = Me.URLDirectorioDocumentacionAyuda & "Video.mp4"
'        If Not fso.FileExists(m_sURLVideo) Then
'            Err.Raise 513, "Entorno.URLVideo", "No se encuentra el archivo de vídeo: " & m_sURLVideo
'        End If
    End If
    URLVideo = m_sURLVideo
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.URLVideo"
    errObj.Raise
End Property
Public Property Get URLAyudaWorkflow() As String
    On Error GoTo Errores

    If m_sURLAyudaWorkflow = "" Then
        m_sURLAyudaWorkflow = Me.URLDirectorioDocumentacionAyuda & "workflow.html"
'        If Not fso.FileExists(m_sURLAyudaWorkflow) Then
'            Err.Raise 513, "Entorno.URLAyudaWorkflow", "No se encuentra el archivo de vídeo: " & m_sURLAyudaWorkflow
'        End If
    End If
    URLAyudaWorkflow = m_sURLAyudaWorkflow
    Exit Property
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.URLAyudaWorkflow"
    errObj.Raise
End Property

Public Property Get URLManual() As String
    On Error GoTo Errores

    If m_sURLManual = "" Then
        m_sURLManual = Me.URLDirectorioDocumentacionAyuda & "Manual de Formación.pdf"
'        If Not fso.FileExists(m_sURLManual) Then
'            Err.Raise 513, "Entorno.URLManual", "No se encuentra el archivo de vídeo: " & m_sURLManual
'        End If
    End If
    URLManual = m_sURLManual
    Exit Property

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.URLManual"
    errObj.Raise
End Property

' ==========================================================================
' PROPIEDADES PARA HTMLs CACHEADOS (Spec-149)
' ==========================================================================
Public Property Get URLDirectorioHTML() As String
    On Error GoTo Errores

    If m_sURLDirectorioHTML = "" Then
        m_sURLDirectorioHTML = Me.URLDirectorioAplicacion & "recursos\html\"
        If Not fso.FolderExists(m_sURLDirectorioHTML) Then
            m_sURLDirectorioHTML = ""
        End If
    End If
    
    URLDirectorioHTML = m_sURLDirectorioHTML
    Exit Property
    
Errores:
    m_sURLDirectorioHTML = ""
    URLDirectorioHTML = ""
End Property

Public Property Get URLHTMLCapturaFechaRAC() As String
    On Error GoTo Errores

    If m_sURLHTMLCapturaFechaRAC = "" Then
        If Me.URLDirectorioHTML <> "" Then
            m_sURLHTMLCapturaFechaRAC = Me.URLDirectorioHTML & "captura_fecha_rac.html"
        End If
    End If
    
    URLHTMLCapturaFechaRAC = m_sURLHTMLCapturaFechaRAC
    Exit Property
    
Errores:
    m_sURLHTMLCapturaFechaRAC = ""
    URLHTMLCapturaFechaRAC = ""
End Property

Public Property Get URLHTMLCapturaComentariosRAC() As String
    On Error GoTo Errores

    If m_sURLHTMLCapturaComentariosRAC = "" Then
        If Me.URLDirectorioHTML <> "" Then
            m_sURLHTMLCapturaComentariosRAC = Me.URLDirectorioHTML & "captura_comentarios_rac.html"
        End If
    End If
    
    URLHTMLCapturaComentariosRAC = m_sURLHTMLCapturaComentariosRAC
    Exit Property
    
Errores:
    m_sURLHTMLCapturaComentariosRAC = ""
    URLHTMLCapturaComentariosRAC = ""
End Property

Public Property Get URLHTMLCapturaRechazoCalidad() As String
    On Error GoTo Errores

    If m_sURLHTMLCapturaRechazoCalidad = "" Then
        If Me.URLDirectorioHTML <> "" Then
            m_sURLHTMLCapturaRechazoCalidad = Me.URLDirectorioHTML & "captura_rechazo_calidad.html"
        End If
    End If
    
    URLHTMLCapturaRechazoCalidad = m_sURLHTMLCapturaRechazoCalidad
    Exit Property
    
Errores:
    m_sURLHTMLCapturaRechazoCalidad = ""
    URLHTMLCapturaRechazoCalidad = ""
End Property

Public Property Get TituloUsuarioConectado() As String
    Dim m_Colectivo As String
    
    ' Esta propiedad no necesita manejo de errores porque se basa en variables globales ya cargadas
    
    ' --- INICIO DE LA MEJORA ---
    ' Refrescamos la variable interna para que siempre se recalcule si cambia el estado de suplantación
    m_sTituloUsuarioConectado = ""
    
    If m_sTituloUsuarioConectado = "" Then
        ' Determinamos el rol del USUARIO ACTIVO (que puede ser el suplantado)
        Select Case rolUsuario
            Case rol.Administrador: m_Colectivo = "Administrador"
            Case rol.Calidad: m_Colectivo = "Calidad"
            Case rol.Tecnico: m_Colectivo = "Técnico"
            Case Else: m_Colectivo = "Desconocido"
        End Select

        ' Construimos el título basándonos en si estamos suplantando o no
        If g_blnImpersonando Then
            ' Formato para cuando se está suplantando
            m_sTituloUsuarioConectado = "Actuando como: " & m_ObjUsuarioActivo.nombre & _
                " (Técnico) | Sesión Real: " & m_ObjUsuarioReal.nombre & " (" & m_Colectivo & ")"
        Else
            ' Formato normal
            m_sTituloUsuarioConectado = "Usuario conectado: " & m_ObjUsuarioActivo.nombre & _
                    " (" & m_Colectivo & ")"
        End If
        
        If Application.TempVars("EnPruebas") = "Sí" Then
            m_sTituloUsuarioConectado = m_sTituloUsuarioConectado & " EN PRUEBAS"
        End If
    End If
    ' --- FIN DE LA MEJORA ---
    
    TituloUsuarioConectado = m_sTituloUsuarioConectado
End Property

Public Property Get ColItems() As Scripting.Dictionary
    ' Propiedad de solo lectura para la inicialización. No necesita manejo de errores.
    If m_objColItems Is Nothing Then
        Set m_objColItems = New Scripting.Dictionary
        With m_objColItems
            .Add "URLAchivoIni", ""
            .Add "VersionAplicacion", ""
            .Add "URLDirectorioAplicaciones", ""
            .Add "URLArchivoCSS", ""
            .Add "CSS", ""
            .Add "URLDirectorioAplicacion", ""
            .Add "URLDirectorioDocumentacion", ""
            .Add "URLDirectorioLocal", ""
            .Add "URLDirectorioDocumentacionAyuda", ""
            .Add "URLPlantillaCDCA", ""
            .Add "URLPlantillaCDCASUB", ""
            .Add "URLPlantillaPC", ""
            .Add "URLPlantillaPCSUB", ""
            .Add "URLVideo", ""
            .Add "URLAyudaWorkflow", ""
            .Add "URLManual", ""
            .Add "TituloUsuarioConectado", ""
            .Add "NombreJefeCalidad", ""
            .Add "DirUTE", ""
            .Add "URLDirectorioHTML", ""
        End With
    End If
    Set ColItems = m_objColItems
End Property

Public Function getPropiedad(ByVal m_NombrePropiedad As String) As Variant
    On Error GoTo Errores
    
    getPropiedad = CallByName(Me, m_NombrePropiedad, VbGet)
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Entorno.getPropiedad"
    errObj.Raise
End Function




