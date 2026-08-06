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
Private m_eDatosEnLocal As EnumSino
Private m_objColTecnicosActivos As Scripting.Dictionary
Private m_objColPerfilesActivos As Scripting.Dictionary

Private m_sURLDirectorioLocal As String
Private m_sURLDirectorioRecursos As String
Private m_sCSS As String
Private m_sURLArchivoCSS As String
Private m_sURLPlantillaInforme As String
Private m_sURLDirectorioPartesFirmados As String
Private m_sURLDirectorioCertificaciones As String
Private m_sURLDirectorioCalibraciones As String
Private m_sURLDirectorioAnexos As String
Private m_sURLDirectorioFacturas As String
Private m_sURLDirectorioDocumentacionAyuda As String
Private m_sURLDirectorioGastos As String
Private m_sURLDirectorioPlanificaciones As String
Private m_sURLDirImportacioneslogs As String
Private m_sURLPlantillaInformeAverias As String
Private m_sURLArchivoUltimo As String
Private m_sURLAchivoIni As String
Private m_sURLDirectorioDocumentacion As String
Private m_sVerAplicacion As String
Private m_sURLDirectorioAplicaciones As String
Private m_sURLDirectorioAplicacion As String
Private m_sDiasParaAvisoDePlanificaciones As String
Private m_objColPropiedadesEntorno As Scripting.Dictionary
Private m_objColTipoAnexos As Scripting.Dictionary
Private m_objColEquiposMedidaActivos As Scripting.Dictionary
Private m_objColEquiposMedida As Scripting.Dictionary
Private m_objCentros As Scripting.Dictionary
Private m_ObjUsuarioConectadoMaquina As Usuario
Private fso As New FileSystemObject
Private m_sURLArchivoSSID As String
Private m_Error As String
Public Error As String

Public Property Get UsuarioConectadoMaquina() As Usuario
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjUsuarioConectadoMaquina Is Nothing Then
        Set UsuarioConectadoMaquina = m_ObjUsuarioConectadoMaquina
        Exit Property
    End If
    Set m_ObjUsuarioConectadoMaquina = Constructor.getUsuarioConectadoPorMaquina(m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Set UsuarioConectadoMaquina = m_ObjUsuarioConectadoMaquina
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.UsuarioConectadoMaquina ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error


End Property

Public Property Get URLArchivoSSID() As String
    
    Dim m_URLDirectorioLocal As String
    On Error GoTo errores
    Me.Error = ""
    If m_sURLArchivoSSID <> "" Then
        URLArchivoSSID = m_sURLArchivoSSID
        Exit Property
    End If
    m_URLDirectorioLocal = Me.URLDirectorioLocal
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_sURLArchivoSSID = fso.GetParentFolderName(m_URLDirectorioLocal) & "\ssid.txt"
    
    URLArchivoSSID = m_sURLArchivoSSID
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLArchivoSSID ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get DatosEnLocal() As EnumSino
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not Application.TempVars("DatosEnLocal") = "Sí" Then
        m_eDatosEnLocal = EnumSino.No
    Else
        m_eDatosEnLocal = EnumSino.Sí
    End If
    DatosEnLocal = m_eDatosEnLocal
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.DatosEnLocal ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get URLDirectorioAplicaciones() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioAplicaciones <> "" Then
        URLDirectorioAplicaciones = m_sURLDirectorioAplicaciones
        Exit Property
    End If
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_sURLDirectorioAplicaciones = m_URLRutaAplicacionesLocal
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_sURLDirectorioAplicaciones = m_URLRutaAplicacionesRemotas
    Else
        m_Error = "No se sabe si se está usando en local o en remoto"
        Err.Raise 1000
    End If
    
   
    If Not fso.FolderExists(m_sURLDirectorioAplicaciones) Then
        m_Error = "No es alcanzable m_sURLDirectorioAplicaciones " & m_sURLDirectorioAplicaciones
        Err.Raise 1000
    End If
    
    URLDirectorioAplicaciones = m_sURLDirectorioAplicaciones
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioAplicaciones ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLDirectorioAplicacion() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLDirectorioAplicacion <> "" Then
        URLDirectorioAplicacion = m_sURLDirectorioAplicacion
        Exit Property
    End If
    m_sURLDirectorioAplicacion = Me.URLDirectorioAplicaciones & "BRASS\"
    
    If Not fso.FolderExists(m_sURLDirectorioAplicacion) Then
        m_Error = "No es alcanzable URLDirectorioAplicacion " & Me.URLDirectorioAplicacion
        Err.Raise 1000
    End If
    URLDirectorioAplicacion = m_sURLDirectorioAplicacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioAplicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Get URLDirectorioLocal() As String
    
    Dim m_URLDirectorioAplicaciones As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioLocal <> "" Then
        URLDirectorioLocal = m_sURLDirectorioLocal
        Exit Property
    End If
    m_URLDirectorioAplicaciones = Environ("APPDATA") & "\" & "Aplicaciones DYSN\"
    If Not fso.FolderExists(m_URLDirectorioAplicaciones) Then
        fso.CreateFolder m_URLDirectorioAplicaciones
    End If
    m_sURLDirectorioLocal = m_URLDirectorioAplicaciones & "Brass\"
    If Not fso.FolderExists(m_sURLDirectorioLocal) Then
        fso.CreateFolder m_sURLDirectorioLocal
    End If
    URLDirectorioLocal = m_sURLDirectorioLocal
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioLocal ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get CSS() As String
    
    Dim ObjDatos As TextStream
    Dim m_URL As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sCSS <> "" Then
        CSS = m_sCSS
        Exit Property
    End If
    m_URL = Me.URLArchivoCSS
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_URL) Then
        m_Error = "Ne es alcanzable el archivo URLArchivoCSS" & vbNewLine & Me.URLArchivoCSS
        Err.Raise 1000
    End If
    Set ObjDatos = fso.OpenTextFile(m_URL, ForReading, False, TristateMixed)
    While Not ObjDatos.AtEndOfStream
          m_sCSS = m_sCSS & ObjDatos.ReadAll
    Wend
    Set ObjDatos = Nothing
    

    CSS = m_sCSS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CSS ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLArchivoCSS() As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLArchivoCSS <> "" Then
        URLArchivoCSS = m_sURLArchivoCSS
        Exit Property
    End If
    m_sURLArchivoCSS = Me.URLDirectorioAplicacion & "CSS.txt"
    If Not fso.FileExists(m_sURLArchivoCSS) Then
        m_Error = "No es alcanzable URLArchivoCSS " & Me.URLArchivoCSS
        Err.Raise 1000
    End If
    
    URLArchivoCSS = m_sURLArchivoCSS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLArchivoCSS ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLDirectorioAnexos() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLDirectorioAnexos <> "" Then
        URLDirectorioAnexos = m_sURLDirectorioAnexos
        Exit Property
    End If
    m_sURLDirectorioAnexos = Me.URLDirectorioAplicacion & "ANEXOS\"
    
    If Not fso.FolderExists(m_sURLDirectorioAnexos) Then
        m_Error = "No es alcanzable URLDirectorioAnexos " & Me.URLDirectorioAnexos
        Err.Raise 1000
    End If
    URLDirectorioAnexos = m_sURLDirectorioAnexos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioAnexos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get URLDirectorioDocumentacion() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLDirectorioDocumentacion <> "" Then
        URLDirectorioDocumentacion = m_sURLDirectorioDocumentacion
        Exit Property
    End If
    m_sURLDirectorioDocumentacion = Me.URLDirectorioAnexos & "ANEXOS\"
    
    If Not fso.FolderExists(m_sURLDirectorioDocumentacion) Then
        m_Error = "No es alcanzable URLDirectorioDocumentacion " & Me.URLDirectorioDocumentacion
        Err.Raise 1000
    End If
    URLDirectorioDocumentacion = m_sURLDirectorioDocumentacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioDocumentacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get URLDirectorioRecursos() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLDirectorioRecursos <> "" Then
        URLDirectorioRecursos = m_sURLDirectorioRecursos
        Exit Property
    End If
    m_sURLDirectorioRecursos = Me.URLDirectorioAplicacion & "recursos\"
    
    If Not fso.FolderExists(m_sURLDirectorioRecursos) Then
        m_Error = "No es alcanzable URLDirectorioRecursos " & Me.URLDirectorioRecursos
        Err.Raise 1000
    End If
    URLDirectorioRecursos = m_sURLDirectorioRecursos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioRecursos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get URLPlantillaInforme() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLPlantillaInforme <> "" Then
        URLPlantillaInforme = m_sURLPlantillaInforme
        Exit Property
    End If
    m_sURLPlantillaInforme = Me.URLDirectorioRecursos & "PlantillaInformePartesTrabajo.xlsx"
    
    If Not fso.FileExists(m_sURLPlantillaInforme) Then
        m_Error = "No es alcanzable URLPlantillaInforme " & Me.URLPlantillaInforme
        Err.Raise 1000
    End If
    URLPlantillaInforme = m_sURLPlantillaInforme
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLPlantillaInforme ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get URLDirectorioPartesFirmados() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLDirectorioPartesFirmados <> "" Then
        URLDirectorioPartesFirmados = m_sURLDirectorioPartesFirmados
        Exit Property
    End If
    m_sURLDirectorioPartesFirmados = Me.URLDirectorioAnexos & "Partes Firmados\"
    
    If Not fso.FolderExists(m_sURLDirectorioPartesFirmados) Then
        m_Error = "No es alcanzable URLDirectorioPartesFirmados " & Me.URLDirectorioPartesFirmados
        Err.Raise 1000
    End If
    URLDirectorioPartesFirmados = m_sURLDirectorioPartesFirmados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioPartesFirmados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get URLDirectorioCertificaciones() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLDirectorioCertificaciones <> "" Then
        URLDirectorioCertificaciones = m_sURLDirectorioCertificaciones
        Exit Property
    End If
    m_sURLDirectorioCertificaciones = Me.URLDirectorioDocumentacion & "Certificaciones\"
    
    If Not fso.FolderExists(m_sURLDirectorioCertificaciones) Then
        m_Error = "No es alcanzable URLDirectorioCertificaciones " & Me.URLDirectorioCertificaciones
        Err.Raise 1000
    End If
    URLDirectorioCertificaciones = m_sURLDirectorioCertificaciones
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioCertificaciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error


End Property

Public Property Get URLDirectorioCalibraciones() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLDirectorioCalibraciones <> "" Then
        URLDirectorioCalibraciones = m_sURLDirectorioCalibraciones
        Exit Property
    End If
    m_sURLDirectorioCalibraciones = Me.URLDirectorioAnexos & "Calibraciones\"
    
    If Not fso.FolderExists(m_sURLDirectorioCalibraciones) Then
        m_Error = "No es alcanzable URLDirectorioCalibraciones " & Me.URLDirectorioCalibraciones
        Err.Raise 1000
    End If
    URLDirectorioCalibraciones = m_sURLDirectorioCalibraciones
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioCalibraciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get URLDirectorioFacturas() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLDirectorioFacturas <> "" Then
        URLDirectorioFacturas = m_sURLDirectorioFacturas
        Exit Property
    End If
    m_sURLDirectorioFacturas = Me.URLDirectorioAnexos & "Facturas\"
    
    If Not fso.FolderExists(m_sURLDirectorioFacturas) Then
        m_Error = "No es alcanzable URLDirectorioFacturas " & Me.URLDirectorioFacturas
        Err.Raise 1000
    End If
    URLDirectorioFacturas = m_sURLDirectorioFacturas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioFacturas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get URLDirectorioDocumentacionAyuda() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLDirectorioDocumentacionAyuda <> "" Then
        URLDirectorioDocumentacionAyuda = m_sURLDirectorioDocumentacionAyuda
        Exit Property
    End If
    m_sURLDirectorioDocumentacionAyuda = Me.URLDirectorioAnexos & "documentos de ayuda\"
    
    If Not fso.FolderExists(m_sURLDirectorioDocumentacionAyuda) Then
        m_Error = "No es alcanzable URLDirectorioDocumentacionAyuda " & Me.URLDirectorioDocumentacionAyuda
        Err.Raise 1000
    End If
    URLDirectorioDocumentacionAyuda = m_sURLDirectorioDocumentacionAyuda
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioDocumentacionAyuda ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get URLDirectorioGastos() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLDirectorioGastos <> "" Then
        URLDirectorioGastos = m_sURLDirectorioGastos
        Exit Property
    End If
    m_sURLDirectorioGastos = Me.URLDirectorioDocumentacion & "Gastos\"
    
    If Not fso.FolderExists(m_sURLDirectorioGastos) Then
        m_Error = "No es alcanzable URLDirectorioGastos " & Me.URLDirectorioGastos
        Err.Raise 1000
    End If
    URLDirectorioGastos = m_sURLDirectorioGastos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioGastos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error


End Property

Public Property Get URLDirectorioPlanificaciones() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLDirectorioPlanificaciones <> "" Then
        URLDirectorioPlanificaciones = m_sURLDirectorioPlanificaciones
        Exit Property
    End If
    m_sURLDirectorioPlanificaciones = Me.URLDirectorioAnexos & "Planificaciones\"
    
    If Not fso.FolderExists(m_sURLDirectorioPlanificaciones) Then
        m_Error = "No es alcanzable URLDirectorioPlanificaciones " & Me.URLDirectorioPlanificaciones
        Err.Raise 1000
    End If
    URLDirectorioPlanificaciones = m_sURLDirectorioPlanificaciones
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioPlanificaciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get URLPlantillaInformeAverias() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLPlantillaInformeAverias <> "" Then
        URLPlantillaInformeAverias = m_sURLPlantillaInformeAverias
        Exit Property
    End If
    m_sURLPlantillaInformeAverias = Me.URLDirectorioAnexos & "PlantillaAverías.docx"
    
    If Not fso.FileExists(m_sURLPlantillaInformeAverias) Then
        m_Error = "No es alcanzable URLPlantillaInformeAverias " & Me.URLPlantillaInformeAverias
        Err.Raise 1000
    End If
    URLPlantillaInformeAverias = m_sURLPlantillaInformeAverias
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLPlantillaInformeAverias ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get URLAchivoIni() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLAchivoIni <> "" Then
        URLAchivoIni = m_sURLAchivoIni
        Exit Property
    End If
    If Me.DatosEnLocal = EnumSino.Sí Then
        m_sURLAchivoIni = Application.CurrentProject.Path & "\" & fso.GetBaseName(Application.CurrentProject.Name) & ".ini"
    Else
        m_sURLAchivoIni = Me.URLDirectorioRecursos & fso.GetBaseName(Application.CurrentProject.Name) & ".ini"
    End If
    If Not fso.FileExists(m_sURLAchivoIni) Then
        m_Error = "No es alcanzable URLAchivoIni " & Me.URLAchivoIni
        Err.Raise 1000
    End If
    URLAchivoIni = m_sURLAchivoIni
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLAchivoIni ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get VersionAplicacion() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sVerAplicacion <> "" Then
        VersionAplicacion = m_sVerAplicacion
        Exit Property
    End If
    m_sVerAplicacion = LeerIni("VersionAplicacion", "")
    VersionAplicacion = m_sVerAplicacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.VersionAplicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get ColPerfilesActivos() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColPerfilesActivos Is Nothing Then
        Set ColPerfilesActivos = m_objColPerfilesActivos
        Exit Property
    End If
    Set m_objColPerfilesActivos = Constructor.getTipoTecnicoPrecios(EnumSino.Sí, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPerfilesActivos = m_objColPerfilesActivos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPerfilesActivos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get ColTecnicosActivos() As Scripting.Dictionary
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColTecnicosActivos Is Nothing Then
        Set ColTecnicosActivos = m_objColTecnicosActivos
        Exit Property
    End If
    Set m_objColTecnicosActivos = Constructor.getTecnicosPorTipo(, EnumSino.Sí, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColTecnicosActivos = m_objColTecnicosActivos
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColTecnicosActivos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Set ColTecnicosActivos(p_Valor As Variant)

    Set m_objColTecnicosActivos = p_Valor
    
End Property

Public Property Get DiasParaAvisoDePlanificaciones() As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If m_sDiasParaAvisoDePlanificaciones <> "" Then
        DiasParaAvisoDePlanificaciones = m_sDiasParaAvisoDePlanificaciones
        Exit Property
    End If
    m_sDiasParaAvisoDePlanificaciones = "15"
    DiasParaAvisoDePlanificaciones = m_sDiasParaAvisoDePlanificaciones
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.DiasParaAvisoDePlanificaciones ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get ColTipoAnexos() As Scripting.Dictionary

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColTipoAnexos Is Nothing Then
        Set ColTipoAnexos = m_objColTipoAnexos
        Exit Property
    End If
    Set m_objColTipoAnexos = New Scripting.Dictionary
    m_objColTipoAnexos.CompareMode = TextCompare
    With Me.ColTipoAnexos
        .Add "1", "Evento"
        .Add "2", "Actividad"
        .Add "3", "Material"
        .Add "4", "MaterialSeg"
        .Add "5", "Subcontratacion"
        .Add "6", "Gasto"
    End With
    Set ColTipoAnexos = m_objColTipoAnexos
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColTipoAnexos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Get ColEquiposMedidaActivos() As Scripting.Dictionary

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
'    If Not m_objColEquiposMedidaActivos Is Nothing Then
'        Set ColEquiposMedidaActivos = m_objColEquiposMedidaActivos
'        Exit Property
'    End If
    Set m_objColEquiposMedidaActivos = Constructor.getEquiposMedidaActivos(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColEquiposMedidaActivos = m_objColEquiposMedidaActivos
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEquiposMedidaActivos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property
Public Property Set ColEquiposMedidaActivos(Valor As Variant)

    
    Set m_objColEquiposMedidaActivos = Valor
    
    
    

End Property

Public Property Get ColEquiposMedida() As Scripting.Dictionary

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColEquiposMedida Is Nothing Then
        Set ColEquiposMedida = m_objColEquiposMedida
        Exit Property
    End If
    Set m_objColEquiposMedida = Constructor.getEquiposMedida(, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColEquiposMedida = m_objColEquiposMedida
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEquiposMedida ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Set ColEquiposMedida(Valor As Variant)

    
    Set m_objColEquiposMedida = Valor
    
    
    

End Property

Public Property Get Centros() As Scripting.Dictionary

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objCentros Is Nothing Then
        Set Centros = m_objCentros
        Exit Property
    End If
    Set m_objCentros = Constructor.getCentros(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Centros = m_objCentros
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.Centros ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Set Centros(Valor As Variant)
    Set m_objCentros = Valor
End Property

Public Property Get ColPropiedadesEntorno() As Scripting.Dictionary
    
    On Error GoTo errores
    
    m_Error = ""
    Me.Error = ""
    
    If Not m_objColPropiedadesEntorno Is Nothing Then
        Set ColPropiedadesEntorno = m_objColPropiedadesEntorno
        Exit Property
    End If
    Set m_objColPropiedadesEntorno = New Scripting.Dictionary
    m_objColPropiedadesEntorno.CompareMode = TextCompare
    
    With m_objColPropiedadesEntorno
        
        .Add "ColTecnicosActivos", "o"
        .Add "ColPerfilesActivos", "o"
        .Add "ColTipoAnexos", "o"
        .Add "DatosEnLocal", "s"
        .Add "URLDirectorioLocal", ""
        .Add "URLDirectorioRecursos", ""
        .Add "CSS", ""
        .Add "URLArchivoCSS", ""
        .Add "URLPlantillaInforme", ""
        .Add "URLDirectorioPartesFirmados", ""
        .Add "URLDirectorioCertificaciones", ""
        .Add "URLDirectorioCalibraciones", ""
        .Add "URLDirectorioAnexos", ""
        .Add "URLDirectorioFacturas", ""
        .Add "URLDirectorioDocumentacionAyuda", ""
        .Add "URLDirectorioGastos", ""
        .Add "URLDirectorioPlanificaciones", ""
        .Add "URLPlantillaInformeAverias", ""
        .Add "URLAchivoIni", ""
        .Add "URLDirectorioDocumentacion", ""
        .Add "VersionAplicacion", ""
        .Add "URLDirectorioAplicaciones", ""
        .Add "URLDirectorioAplicacion", ""
        .Add "DiasParaAvisoDePlanificaciones", ""
        .Add "ColEquiposMedidaActivos", "o"
        .Add "ColEquiposMedida", "o"
    End With
    Set ColPropiedadesEntorno = m_objColPropiedadesEntorno
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El metodo Entorno.ColPropiedadesEntorno ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Function getValorPropiedad( _
                                    p_NombreCampo As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Variant
    
    On Error GoTo errores
    m_Error = ""
    p_Error = ""
    
   
    
    If p_NombreCampo = "ColTecnicosActivos" Then
        Set getValorPropiedad = Me.ColTecnicosActivos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColPerfilesActivos" Then
        Set getValorPropiedad = Me.ColPerfilesActivos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColTipoAnexos" Then
        Set getValorPropiedad = Me.ColTipoAnexos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "DatosEnLocal" Then
        getValorPropiedad = Me.DatosEnLocal
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLDirectorioLocal" Then
        getValorPropiedad = Me.URLDirectorioLocal
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLDirectorioRecursos" Then
        getValorPropiedad = Me.URLDirectorioRecursos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombreCampo = "CSS" Then
        getValorPropiedad = Me.CSS
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    
    If p_NombreCampo = "URLArchivoCSS" Then
        getValorPropiedad = Me.URLArchivoCSS
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    
    If p_NombreCampo = "URLPlantillaInforme" Then
        getValorPropiedad = Me.URLPlantillaInforme
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLDirectorioPartesFirmados" Then
        getValorPropiedad = Me.URLDirectorioPartesFirmados
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLDirectorioCertificaciones" Then
        getValorPropiedad = Me.URLDirectorioCertificaciones
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombreCampo = "URLDirectorioCalibraciones" Then
        getValorPropiedad = Me.URLDirectorioCalibraciones
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLDirectorioAnexos" Then
        getValorPropiedad = Me.URLDirectorioAnexos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLDirectorioFacturas" Then
        getValorPropiedad = Me.URLDirectorioFacturas
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLDirectorioDocumentacionAyuda" Then
        getValorPropiedad = Me.URLDirectorioDocumentacionAyuda
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLDirectorioGastos" Then
        getValorPropiedad = Me.URLDirectorioGastos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLDirectorioPlanificaciones" Then
        getValorPropiedad = Me.URLDirectorioPlanificaciones
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombreCampo = "URLPlantillaInformeAverias" Then
        getValorPropiedad = Me.URLPlantillaInformeAverias
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLAchivoIni" Then
        getValorPropiedad = Me.URLAchivoIni
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLDirectorioDocumentacion" Then
        getValorPropiedad = Me.URLDirectorioDocumentacion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombreCampo = "VersionAplicacion" Then
        getValorPropiedad = Me.VersionAplicacion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLDirectorioAplicaciones" Then
        getValorPropiedad = Me.URLDirectorioAplicaciones
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLDirectorioAplicacion" Then
        getValorPropiedad = Me.URLDirectorioAplicacion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "DiasParaAvisoDePlanificaciones" Then
        getValorPropiedad = Me.DiasParaAvisoDePlanificaciones
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColEquiposMedidaActivos" Then
        Set getValorPropiedad = Me.ColEquiposMedidaActivos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColEquiposMedida" Then
        Set getValorPropiedad = Me.ColEquiposMedida
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    p_Error = "La propiedad " & p_NombreCampo & " no esta registrada en Entorno.getValorPropiedad "
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Entorno.getValorPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function RellenarColesConCampos(Optional ByRef p_Error As String) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_IDEnum As EnumCampos
    Dim m_NombreCampoParaInforme As String
    Dim m_NombreCampoParaDato As String
    On Error GoTo errores
    Set m_ColParaNombreCamposInformeTabla = New Scripting.Dictionary
    Set m_ColParaEnumNombreCamposEnInforme = New Scripting.Dictionary
    m_SQL = "TbAuxNombreCamposParaInforme"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                m_IDEnum = .Fields("IDNombreCampo")
                m_NombreCampoParaInforme = .Fields("NombreParaInforme")
                m_NombreCampoParaDato = .Fields("NombreCampoDeTabla")
                m_ColParaEnumNombreCamposEnInforme.Add CStr(m_IDEnum), m_NombreCampoParaInforme
                m_ColParaNombreCamposInformeTabla.Add m_NombreCampoParaInforme, m_NombreCampoParaDato
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarColesConCampos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Sub Class_Initialize()
    
    'If m_ColParaEnumNombreCamposEnInforme Is Nothing Or m_ColParaNombreCamposInformeTabla Is Nothing Then
        RellenarColesConCampos
    'End If
End Sub


