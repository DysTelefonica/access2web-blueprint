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

Public EnDesarrollo As EnumSiNo
Public URLUltimoArchivo As String
Public OrdenAscendenteEnListaProyectos As EnumSiNo

Private m_sURLDirectorioAplicaciones As String
Private m_sURLDirectorioAplicacion As String
Private m_sURLDirectorioTemporalAGEDO As String
Private m_sURLDirectorioDocumentacion As String
Private m_sURLDirectorioAyuda As String
Private m_sURLDirectorioRecursos As String
Private m_sURLPlantillaExcel As String
Private m_sURLPlantillaWord As String

Private m_sURLPlantillaEvolucionRiesgosExcel As String
Private m_sURLAchivoIni As String
Private m_sJPMesesAvisoEntreEdiciones As Integer
Private m_sJPDiasPreviosParaElAviso As Integer
Private m_sCalDiaInicialMesAviso As Integer
Private m_sURLDirectorioLocal As String
Private m_sVersionAplicacion As String
Private m_sValorAlarma As String
Private m_sValorObjetivo As String
Private m_sURLIconoCarpetaCompletaCerrada32 As String
Private m_sURLIconoCarpetaCompletaAbierta32 As String
Private m_sURLIconoCarpetaIncompletaCerrada32 As String
Private m_sURLIconoCarpetaIncompletaAbierta32 As String
Private m_sURLIconoItemCompleto32 As String
Private m_sURLIconoItemInCompleto32 As String

Private m_sURLVideo As String
Private m_sURLDirIconos As String
Private m_sURLIconoProyecto As String
Private m_sURLIconoOK As String
Private m_sURLIconoNoOK As String
Private m_sURLDirectorioAplicacionesRecursos As String
Private m_sURLIconoPDF As String
Private m_sURLIconoZIP As String
Private m_sURLDirectorioDocumentosCambio As String
Private m_sURLDirectorioDocumentosRequisitosCambio As String
Private m_sTituloUsuarioConectado As String
Private m_eusuarioConectadoEsAdministrador As EnumSiNo
Private m_eusuarioConectadoEsDeCalidad As EnumSiNo
Private m_eusuarioConectadoEsTecnico As EnumSiNo

Private m_sCorreoCalidad As String
Private m_sCSS As String
Private m_objColDatosPlazoCalidadCosteVulnerabilidad As Scripting.Dictionary
Private m_objColMitigacionValores As Scripting.Dictionary
Private m_objColOrigenRiesgosValores As Scripting.Dictionary
Private m_objColDatosVulnerabilidades As Scripting.Dictionary
Private m_objColRiesgoEstados As Scripting.Dictionary
Private m_sCabeceraHTML As String
Private m_sFinHTML As String
Private m_ObjColProyectosTotales As Scripting.Dictionary
Private m_objColTipoTareas As Scripting.Dictionary
Private m_objColAreaImpactoCalidad As Scripting.Dictionary
Private m_objColAreaImpactoCoste As Scripting.Dictionary
Private m_objColAreaImpactoPlazo As Scripting.Dictionary
Private m_ObjColRiesgosBibliotecasFamilias As Scripting.Dictionary
Private m_sDestinatariosCorreoAdministradores As String
Private m_objColUsuariosAdministradores As Scripting.Dictionary
Private m_objColUsuariosCalidad As Scripting.Dictionary
Private m_objColUsuariosCalidadAvisos As Scripting.Dictionary
Private m_objColUsuariosTecnicos As Scripting.Dictionary
Private m_objColUsuariosTecnicosActivos As Scripting.Dictionary
Private m_objColUsuariosTecnicosConAlgunaGestion As Scripting.Dictionary
Private m_objColJPs As Scripting.Dictionary

Private m_objColEdicionesCaducadas As Scripting.Dictionary
Private m_objColEdicionesAPuntoDeCaducar As Scripting.Dictionary
Private m_objColTareasTipo As Scripting.Dictionary
Private m_objColImpactoNumericoParaGraficas As Scripting.Dictionary
Private m_objColEstadoNumericoParaGraficas As Scripting.Dictionary
Private m_objColExplicacionesArbolTareas As Scripting.Dictionary
Private m_eVerSoloRiesgosNoRetirados As EnumSiNo
Private m_eVerRiesgosDescripcion As EnumSiNo
Private m_objRiesgoBibliotecaNoExiste As RiesgoBiblioteca

Private m_objColCambiosExplicaciones As Scripting.Dictionary
Private m_ObjUsuarioConectadoMaquina As Usuario
Private m_sURLArchivoSSID As String
Private m_objColTipos As Scripting.Dictionary

Private m_objColImagenes As Scripting.Dictionary



Private m_objColCarenciasExplicaciones As Scripting.Dictionary
Private m_sCodigoBibliotecaParaRetipificar As String
Private m_sURLArchivoCSS As String
Private m_objColVersiones As Scripting.Dictionary
Private m_objVersionesAplicacion As Scripting.Dictionary
Private m_objColEstados As Scripting.Dictionary
Private m_objColEstadosTexto As Scripting.Dictionary
Private m_objColEstadosTitulo As Scripting.Dictionary
Private m_objColEstadosTituloMejorado As Scripting.Dictionary
Private m_ListaUsuarios As String
Private m_sURLAyudaIndicadores As String

Public Error As String
Private m_Error As String
Private m_objColItems As Scripting.Dictionary
Public Property Get ColEstados() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColEstados Is Nothing Then
        Set ColEstados = m_objColEstados
        Exit Property
    End If
    Set m_objColEstados = New Scripting.Dictionary
    m_objColEstados.CompareMode = TextCompare
    With m_objColEstados
        .Add CStr(EnumEstados.Oferta), "Oferta"
        .Add CStr(EnumEstados.Adjudicada), "Adjudicada"
        .Add CStr(EnumEstados.EnEjecucion), "EnEjecucion"
        .Add CStr(EnumEstados.EnGarantia), "EnGarantia"
        .Add CStr(EnumEstados.Cerrado), "Cerrado"
        .Add CStr(EnumEstados.Desestimado), "Desestimado"
        .Add CStr(EnumEstados.Perdido), "Perdido"
        .Add CStr(EnumEstados.NoAplica), "NoAPlica"
        .Add CStr(EnumEstados.Desconocido), "Desconocido"
    End With
    Set ColEstados = m_objColEstados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEstados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColEstadosTexto() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColEstadosTexto Is Nothing Then
        Set ColEstadosTexto = m_objColEstadosTexto
        Exit Property
    End If
    Set m_objColEstadosTexto = New Scripting.Dictionary
    m_objColEstadosTexto.CompareMode = TextCompare
    With m_objColEstadosTexto
        .Add "Oferta", EnumEstados.Oferta
        .Add "Adjudicada", EnumEstados.Adjudicada
        .Add "EnEjecucion", EnumEstados.EnEjecucion
        .Add "EnGarantia", EnumEstados.EnGarantia
        .Add "Cerrado", EnumEstados.Cerrado
        .Add "Desestimado", EnumEstados.Desestimado
        .Add "Perdido", EnumEstados.Perdido
        .Add "NoAPlica", EnumEstados.NoAplica
        .Add "Desconocido", EnumEstados.Desconocido
    End With
    Set ColEstadosTexto = m_objColEstadosTexto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEstadosTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColEstadosTitulo() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColEstadosTitulo Is Nothing Then
        Set ColEstadosTitulo = m_objColEstadosTitulo
        Exit Property
    End If
    Set m_objColEstadosTitulo = New Scripting.Dictionary
    m_objColEstadosTitulo.CompareMode = TextCompare
    With m_objColEstadosTitulo
        .Add CStr(EnumEstados.Oferta), "Oferta"
        .Add CStr(EnumEstados.Adjudicada), "Adjudicada"
        .Add CStr(EnumEstados.EnEjecucion), "Ejecución"
        .Add CStr(EnumEstados.EnGarantia), "Garantía"
        .Add CStr(EnumEstados.Cerrado), "Cerrado"
        .Add CStr(EnumEstados.Desestimado), "Desestimado"
        .Add CStr(EnumEstados.Perdido), "Perdido"
        .Add CStr(EnumEstados.NoAplica), "NoAPlica"
        .Add CStr(EnumEstados.Desconocido), "Desconocido"
    End With
    Set ColEstadosTitulo = m_objColEstadosTitulo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEstadosTitulo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get ColEstadosTituloMejorado() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColEstadosTituloMejorado Is Nothing Then
        Set ColEstadosTituloMejorado = m_objColEstadosTituloMejorado
        Exit Property
    End If
    Set m_objColEstadosTituloMejorado = New Scripting.Dictionary
    m_objColEstadosTituloMejorado.CompareMode = TextCompare
    With m_objColEstadosTituloMejorado
       
        .Add CStr(EnumEstados.Oferta), "En fase de oferta"
        .Add CStr(EnumEstados.Adjudicada), "Adjudicada"
        .Add CStr(EnumEstados.EnEjecucion), "En fase de ejecución"
        .Add CStr(EnumEstados.EnGarantia), "En garantía"
        .Add CStr(EnumEstados.Cerrado), "Cerrado"
        .Add CStr(EnumEstados.Desestimado), "Desestimado"
        .Add CStr(EnumEstados.Perdido), "Perdido"
        .Add CStr(EnumEstados.NoAplica), "NoAPlica"
        .Add CStr(EnumEstados.Desconocido), "Desconocido"
    End With
    Set ColEstadosTituloMejorado = m_objColEstadosTituloMejorado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEstadosTituloMejorado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get URLDirectorioDocumentosRequisitosCambio() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioDocumentosRequisitosCambio <> "" Then
        URLDirectorioDocumentosRequisitosCambio = m_sURLDirectorioDocumentosRequisitosCambio
        Exit Property
    End If
    m_sURLDirectorioDocumentosRequisitosCambio = Me.URLDirectorioAplicaciones & "CONTROL CAMBIOS\ANEXOS\Requisitos\"
    
    If Not fso.FolderExists(m_sURLDirectorioDocumentosRequisitosCambio) Then
        m_Error = "No es alcanzable m_sURLDirectorioDocumentosRequisitosCambio " & m_sURLDirectorioDocumentosRequisitosCambio
        Err.Raise 1000
    End If
    
    URLDirectorioDocumentosRequisitosCambio = m_sURLDirectorioDocumentosRequisitosCambio
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioDocumentosRequisitosCambio ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLDirectorioDocumentosCambio() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioDocumentosCambio <> "" Then
        URLDirectorioDocumentosCambio = m_sURLDirectorioDocumentosCambio
        Exit Property
    End If
    m_sURLDirectorioDocumentosCambio = Me.URLDirectorioAplicaciones & "CONTROL CAMBIOS\ANEXOS\Cambios\"
    
    If Not fso.FolderExists(m_sURLDirectorioDocumentosCambio) Then
        m_Error = "No es alcanzable m_sURLDirectorioDocumentosCambio " & m_sURLDirectorioDocumentosCambio
        Err.Raise 1000
    End If
    
    URLDirectorioDocumentosCambio = m_sURLDirectorioDocumentosCambio
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioDocumentosCambio ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLIconoPDF() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLIconoPDF <> "" Then
        URLIconoPDF = m_sURLIconoPDF
        Exit Property
    End If
    m_sURLIconoPDF = Me.URLDirIconos & "PDF_32x32.png"
    
    If Not fso.FileExists(m_sURLIconoPDF) Then
        m_Error = "No es alcanzable m_sURLIconoPDF " & m_sURLIconoPDF
        Err.Raise 1000
    End If
    
    URLIconoPDF = m_sURLIconoPDF
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLIconoPDF ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get URLIconoZIP() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLIconoZIP <> "" Then
        URLIconoZIP = m_sURLIconoZIP
        Exit Property
    End If
    m_sURLIconoZIP = Me.URLDirIconos & "ZIP_32x32.png"
    
    If Not fso.FileExists(m_sURLIconoZIP) Then
        m_Error = "No es alcanzable m_sURLIconoZIP " & m_sURLIconoZIP
        Err.Raise 1000
    End If
    
    URLIconoZIP = m_sURLIconoZIP
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLIconoZIP ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get URLDirectorioAplicacionesRecursos() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioAplicacionesRecursos <> "" Then
        URLDirectorioAplicacionesRecursos = m_sURLDirectorioAplicacionesRecursos
        Exit Property
    End If
    m_sURLDirectorioAplicacionesRecursos = Me.URLDirectorioAplicaciones & "00Recursos\"
    
    If Not fso.FolderExists(m_sURLDirectorioAplicacionesRecursos) Then
        m_Error = "No es alcanzable m_sURLDirectorioAplicacionesRecursos " & m_sURLDirectorioAplicacionesRecursos
        Err.Raise 1000
    End If
    
    URLDirectorioAplicacionesRecursos = m_sURLDirectorioAplicacionesRecursos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioAplicacionesRecursos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get URLIconoNoOK() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLIconoNoOK <> "" Then
        URLIconoNoOK = m_sURLIconoNoOK
        Exit Property
    End If
    m_sURLIconoNoOK = Me.URLDirIconos & "NoOK.png"
    
    If Not fso.FileExists(m_sURLIconoNoOK) Then
        m_Error = "No es alcanzable m_sURLIconoNoOK " & m_sURLIconoNoOK
        Err.Raise 1000
    End If
    
    URLIconoNoOK = m_sURLIconoNoOK
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLIconoNoOK ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLIconoOK() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLIconoOK <> "" Then
        URLIconoOK = m_sURLIconoOK
        Exit Property
    End If
    m_sURLIconoOK = Me.URLDirIconos & "OK.png"
    
    If Not fso.FileExists(m_sURLIconoOK) Then
        m_Error = "No es alcanzable m_sURLIconoOK " & m_sURLIconoOK
        Err.Raise 1000
    End If
    
    URLIconoOK = m_sURLIconoOK
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLIconoOK ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLIconoProyecto() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLIconoProyecto <> "" Then
        URLIconoProyecto = m_sURLIconoProyecto
        Exit Property
    End If
    m_sURLIconoProyecto = Me.URLDirIconos & "Proyecto.ico"
    
    If Not fso.FileExists(m_sURLIconoProyecto) Then
        m_Error = "No es alcanzable m_sURLIconoProyecto " & m_sURLIconoProyecto
        Err.Raise 1000
    End If
    
    URLIconoProyecto = m_sURLIconoProyecto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLIconoProyecto ha devuelto el error:" & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get URLDirIconos() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirIconos <> "" Then
        URLDirIconos = m_sURLDirIconos
        Exit Property
    End If
    m_sURLDirIconos = Me.URLDirectorioRecursos & "Iconos" & "\"
    
    If Not fso.FolderExists(m_sURLDirIconos) Then
        m_Error = "No es alcanzable m_sURLDirIconos " & m_sURLDirIconos
        Err.Raise 1000
    End If
    
    URLDirIconos = m_sURLDirIconos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirIconos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get URLIconoItemInCompleto32() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLIconoItemInCompleto32 <> "" Then
        URLIconoItemInCompleto32 = m_sURLIconoItemInCompleto32
        Exit Property
    End If
    m_sURLIconoItemInCompleto32 = Me.URLDirIconos & "Item_Incompleto_32x32.ico"
    
    If Not fso.FileExists(m_sURLIconoItemInCompleto32) Then
        m_Error = "No es alcanzable m_sURLIconoItemInCompleto32 " & m_sURLIconoItemInCompleto32
        Err.Raise 1000
    End If
    
    URLIconoItemInCompleto32 = m_sURLIconoItemInCompleto32
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLIconoItemInCompleto32 ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLIconoItemCompleto32() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLIconoItemCompleto32 <> "" Then
        URLIconoItemCompleto32 = m_sURLIconoItemCompleto32
        Exit Property
    End If
    m_sURLIconoItemCompleto32 = Me.URLDirIconos & "Item_Completo_32x32.ico"
    
    If Not fso.FileExists(m_sURLIconoItemCompleto32) Then
        m_Error = "No es alcanzable m_sURLIconoItemCompleto32 " & m_sURLIconoItemCompleto32
        Err.Raise 1000
    End If
    
    URLIconoItemCompleto32 = m_sURLIconoItemCompleto32
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLIconoItemCompleto32 ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get URLIconoCarpetaIncompletaAbierta32() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLIconoCarpetaIncompletaAbierta32 <> "" Then
        URLIconoCarpetaIncompletaAbierta32 = m_sURLIconoCarpetaIncompletaAbierta32
        Exit Property
    End If
    m_sURLIconoCarpetaIncompletaAbierta32 = Me.URLDirIconos & "CarpetaIncompletaAbierta_32x32.ico"
    
    If Not fso.FileExists(m_sURLIconoCarpetaIncompletaAbierta32) Then
        m_Error = "No es alcanzable m_sURLIconoCarpetaIncompletaAbierta32 " & m_sURLIconoCarpetaIncompletaAbierta32
        Err.Raise 1000
    End If
    
    URLIconoCarpetaIncompletaAbierta32 = m_sURLIconoCarpetaIncompletaAbierta32
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLIconoCarpetaIncompletaAbierta32 ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property



Public Property Get URLIconoCarpetaIncompletaCerrada32() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLIconoCarpetaIncompletaCerrada32 <> "" Then
        URLIconoCarpetaIncompletaCerrada32 = m_sURLIconoCarpetaIncompletaCerrada32
        Exit Property
    End If
    m_sURLIconoCarpetaIncompletaCerrada32 = Me.URLDirIconos & "CarpetaIncompletaCerrada_32x32.ico"
    
    If Not fso.FileExists(m_sURLIconoCarpetaIncompletaCerrada32) Then
        m_Error = "No es alcanzable m_sURLIconoCarpetaIncompletaCerrada32 " & m_sURLIconoCarpetaIncompletaCerrada32
        Err.Raise 1000
    End If
    
    URLIconoCarpetaIncompletaCerrada32 = m_sURLIconoCarpetaIncompletaCerrada32
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLIconoCarpetaIncompletaCerrada32 ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLIconoCarpetaCompletaAbierta32() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLIconoCarpetaCompletaAbierta32 <> "" Then
        URLIconoCarpetaCompletaAbierta32 = m_sURLIconoCarpetaCompletaAbierta32
        Exit Property
    End If
    m_sURLIconoCarpetaCompletaAbierta32 = Me.URLDirIconos & "CarpetaCompletaAbierta_32x32.ico"
    
    If Not fso.FileExists(m_sURLIconoCarpetaCompletaAbierta32) Then
        m_Error = "No es alcanzable m_sURLIconoCarpetaCompletaAbierta32 " & m_sURLIconoCarpetaCompletaAbierta32
        Err.Raise 1000
    End If
    
    URLIconoCarpetaCompletaAbierta32 = m_sURLIconoCarpetaCompletaAbierta32
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLIconoCarpetaCompletaAbierta32 ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get URLIconoCarpetaCompletaCerrada32() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLIconoCarpetaCompletaCerrada32 <> "" Then
        URLIconoCarpetaCompletaCerrada32 = m_sURLIconoCarpetaCompletaCerrada32
        Exit Property
    End If
    m_sURLIconoCarpetaCompletaCerrada32 = Me.URLDirIconos & "CarpetaCompletaCerrada_32x32.ico"
    
    If Not fso.FileExists(m_sURLIconoCarpetaCompletaCerrada32) Then
        m_Error = "No es alcanzable m_sURLIconoCarpetaCompletaCerrada32 " & m_sURLIconoCarpetaCompletaCerrada32
        Err.Raise 1000
    End If
    
    URLIconoCarpetaCompletaCerrada32 = m_sURLIconoCarpetaCompletaCerrada32
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLIconoCarpetaCompletaCerrada32 ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get ValorObjetivo() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sValorObjetivo <> "" Then
        ValorObjetivo = m_sValorObjetivo
        Exit Property
    End If
    m_sValorObjetivo = "15"
    
    ValorObjetivo = m_sValorObjetivo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ValorObjetivo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get ValorAlarma() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sValorAlarma <> "" Then
        ValorAlarma = m_sValorAlarma
        Exit Property
    End If
    m_sValorAlarma = "15"
    
    ValorAlarma = m_sValorAlarma
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ValorAlarma ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get VersionAplicacion() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sVersionAplicacion <> "" Then
        VersionAplicacion = m_sVersionAplicacion
        Exit Property
    End If
    m_sVersionAplicacion = LeerIni("VersionAplicacion", "#ERR#")
    
    VersionAplicacion = m_sVersionAplicacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.VersionAplicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get URLDirectorioLocal() As String
    
    Dim m_URL As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioLocal <> "" Then
        URLDirectorioLocal = m_sURLDirectorioLocal
        Exit Property
    End If
    m_URL = Environ("APPDATA") & "\Aplicaciones DYSN\"
    If Not fso.FolderExists(m_URL) Then
        fso.CreateFolder m_URL
    End If
    m_sURLDirectorioLocal = m_URL & "GESTION RIESGOS\"
    If Not fso.FolderExists(m_sURLDirectorioLocal) Then
        fso.CreateFolder m_sURLDirectorioLocal
    End If
    
    
    URLDirectorioLocal = m_sURLDirectorioLocal
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioLocal ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get CalDiaInicialMesAviso() As Integer
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sCalDiaInicialMesAviso = Nz(Application.TempVars("CalDiaInicialMesAviso"), 2)
    
    CalDiaInicialMesAviso = m_sCalDiaInicialMesAviso
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CalDiaInicialMesAviso ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get JPDiasPreviosParaElAviso() As Integer
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sJPDiasPreviosParaElAviso = Nz(Application.TempVars("JPDiasPreviosParaElAviso"), 15)
    
    JPDiasPreviosParaElAviso = m_sJPDiasPreviosParaElAviso
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.JPDiasPreviosParaElAviso ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get JPMesesAvisoEntreEdiciones() As Integer
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sJPMesesAvisoEntreEdiciones = Nz(Application.TempVars("JPMesesAvisoEntreEdiciones"), 3)
    
    JPMesesAvisoEntreEdiciones = m_sJPMesesAvisoEntreEdiciones
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.JPMesesAvisoEntreEdiciones ha devuelto el error: " & vbNewLine & Err.Description
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
    m_sURLAchivoIni = Application.CurrentProject.Path & "\" & fso.GetBaseName(Application.CurrentProject.Name) & ".ini"
    If Not fso.FileExists(m_sURLAchivoIni) Then
        m_Error = "No es alcanzable m_sURLAchivoIni " & vbNewLine & m_sURLAchivoIni
        m_sURLAchivoIni = ""
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


Public Property Get URLPlantillaEvolucionRiesgosExcel() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLPlantillaEvolucionRiesgosExcel <> "" Then
        URLPlantillaEvolucionRiesgosExcel = m_sURLPlantillaEvolucionRiesgosExcel
        Exit Property
    End If
    m_sURLPlantillaEvolucionRiesgosExcel = Me.URLDirectorioRecursos & "PlantillaEvolucionRiesgos.xlsx"
    
    If Not fso.FileExists(m_sURLPlantillaEvolucionRiesgosExcel) Then
        m_Error = "No es alcanzable m_sURLPlantillaEvolucionRiesgosExcel " & m_sURLPlantillaEvolucionRiesgosExcel
        Err.Raise 1000
    End If
    
    URLPlantillaEvolucionRiesgosExcel = m_sURLPlantillaExcel
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLPlantillaEvolucionRiesgosExcel ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get URLAyudaIndicadores() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLAyudaIndicadores <> "" Then
        URLAyudaIndicadores = m_sURLAyudaIndicadores
        Exit Property
    End If
    m_sURLAyudaIndicadores = Me.URLDirectorioAyuda & "Indicadores\Indicadores.html"
    
    If Not fso.FileExists(m_sURLAyudaIndicadores) Then
        m_Error = "No es alcanzable m_sURLAyudaIndicadores " & m_sURLAyudaIndicadores
        Err.Raise 1000
    End If
    
    URLAyudaIndicadores = m_sURLAyudaIndicadores
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLAyudaIndicadores ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get URLVideo() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLVideo <> "" Then
        URLVideo = m_sURLVideo
        Exit Property
    End If
    m_sURLVideo = Me.URLDirectorioAyuda & "Video.mp4"
    
    If Not fso.FileExists(m_sURLVideo) Then
        m_Error = "No es alcanzable m_sURLVideo " & m_sURLVideo
        Err.Raise 1000
    End If
    
    URLVideo = m_sURLPlantillaExcel
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLVideo ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get URLPlantillaExcel() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLPlantillaExcel <> "" Then
        URLPlantillaExcel = m_sURLPlantillaExcel
        Exit Property
    End If
    m_sURLPlantillaExcel = Me.URLDirectorioRecursos & "Plantilla.xlsx"
    
    If Not fso.FileExists(m_sURLPlantillaExcel) Then
        m_Error = "No es alcanzable m_sURLPlantillaExcel " & m_sURLPlantillaExcel
        Err.Raise 1000
    End If
    
    URLPlantillaExcel = m_sURLPlantillaExcel
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLPlantillaExcel ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
    Public Property Get URLPlantillaWord() As String
    On Error GoTo errores
    If m_sURLPlantillaWord <> "" Then
        URLPlantillaWord = m_sURLPlantillaWord
        Exit Property
    End If
    m_sURLPlantillaWord = Me.URLDirectorioRecursos & "PlantillaConMarcadores.docx"
    If Not fso.FileExists(m_sURLPlantillaWord) Then
        m_Error = "No es alcanzable m_sURLPlantillaWord " & m_sURLPlantillaWord
        Err.Raise 1000
    End If
    URLPlantillaWord = m_sURLPlantillaWord
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLPlantillaWord ha devuelto el error:" & vbNewLine & Err.Description
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
        m_Error = "No es alcanzable m_sURLDirectorioRecursos " & m_sURLDirectorioRecursos
        Err.Raise 1000
    End If
    
    URLDirectorioRecursos = m_sURLDirectorioRecursos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioRecursos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLDirectorioAyuda() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioAyuda <> "" Then
        URLDirectorioAyuda = m_sURLDirectorioAyuda
        Exit Property
    End If
    m_sURLDirectorioAyuda = Me.URLDirectorioDocumentacion & "DOC AYUDA\"
    
    If Not fso.FolderExists(m_sURLDirectorioAyuda) Then
        m_Error = "No es alcanzable m_sURLDirectorioAyuda " & m_sURLDirectorioAyuda
        Err.Raise 1000
    End If
    
    URLDirectorioAyuda = m_sURLDirectorioDocumentacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioAyuda ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLDirectorioDocumentacion() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioDocumentacion <> "" Then
        URLDirectorioDocumentacion = m_sURLDirectorioDocumentacion
        Exit Property
    End If
    m_sURLDirectorioDocumentacion = Me.URLDirectorioAplicacion & "ANEXOS\"
    
    If Not fso.FolderExists(m_sURLDirectorioDocumentacion) Then
        m_Error = "No es alcanzable m_sURLDirectorioDocumentacion " & m_sURLDirectorioDocumentacion
        Err.Raise 1000
    End If
    
    URLDirectorioDocumentacion = m_sURLDirectorioDocumentacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioDocumentacion ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLDirectorioTemporalAGEDO() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioTemporalAGEDO <> "" Then
        URLDirectorioTemporalAGEDO = m_sURLDirectorioTemporalAGEDO
        Exit Property
    End If
    
    ' Normalizar bases con trailing backslash
    Dim baseRemota As String: baseRemota = m_URLRutaAplicacionesRemotas
    Dim baseLocal As String: baseLocal = m_URLRutaAplicacionesLocal
    If Right(baseRemota, 1) <> "\" Then baseRemota = baseRemota & "\"
    If Right(baseLocal, 1) <> "\" Then baseLocal = baseLocal & "\"
    
    If Application.TempVars("EnPruebas") = "Sí" Then
        If m_BackendActivo = "PROD" Then
            m_sURLDirectorioTemporalAGEDO = baseRemota & "00Recursos\Pruebas\1ParaAGEDO\"
        Else
            m_sURLDirectorioTemporalAGEDO = baseLocal & "00Recursos\Pruebas\1ParaAGEDO\"
        End If
    Else
        If m_BackendActivo = "PROD" Then
            m_sURLDirectorioTemporalAGEDO = baseRemota & "1ParaAGEDO\"
        Else
            m_sURLDirectorioTemporalAGEDO = baseLocal & "1ParaAGEDO\"
        End If
    End If
    
    
    
    If Application.TempVars("EnPruebas") = "Sí" Then
        If m_BackendActivo = "PROD" Then
            m_sURLDirectorioTemporalAGEDO = m_URLRutaAplicacionesRemotas & "00Recursos\Pruebas\1ParaAGEDO\"
        Else
            m_sURLDirectorioTemporalAGEDO = m_URLRutaAplicacionesLocal & "00Recursos\Pruebas\1ParaAGEDO\"
        End If


    Else
        If m_BackendActivo = "PROD" Then
            m_sURLDirectorioTemporalAGEDO = m_URLRutaAplicacionesRemotas & "1ParaAGEDO\"
        Else
            m_sURLDirectorioTemporalAGEDO = m_URLRutaAplicacionesLocal & "1ParaAGEDO\"
        End If



    End If
    If Not fso.FolderExists(m_sURLDirectorioTemporalAGEDO) Then
        m_Error = "No es alcanzable m_sURLDirectorioTemporalAGEDO " & m_sURLDirectorioTemporalAGEDO
        Err.Raise 1000
    End If
    
    URLDirectorioTemporalAGEDO = m_sURLDirectorioTemporalAGEDO
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioTemporalAGEDO ha devuelto el error:" & vbNewLine & Err.Description
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
    
    ' Usar la ruta directa de Variables Globales — ya resuelta por m_BackendActivo
    If m_BackendActivo = "PROD" Then
        m_sURLDirectorioAplicacion = m_RutaDirApp_PROD
    Else
        m_sURLDirectorioAplicacion = m_RutaDirApp_LOCAL
    End If
    
    If m_sURLDirectorioAplicacion = "" Then
        m_Error = "No está definida la ruta del directorio de aplicación"
        Err.Raise 1000
    End If
    
    If Not fso.FolderExists(m_sURLDirectorioAplicacion) Then
        m_Error = "No es alcanzable m_sURLDirectorioAplicacion " & m_sURLDirectorioAplicacion
        Err.Raise 1000
    End If
    
    ' Normalizar: asegurar que termina en backslash
    If Right(m_sURLDirectorioAplicacion, 1) <> "\" Then
        m_sURLDirectorioAplicacion = m_sURLDirectorioAplicacion & "\"
    End If
    
    URLDirectorioAplicacion = m_sURLDirectorioAplicacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioAplicacion ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get URLDirectorioAplicaciones() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirectorioAplicaciones <> "" Then
        URLDirectorioAplicaciones = m_sURLDirectorioAplicaciones
        Exit Property
    End If
    If m_BackendActivo = "PROD" Then
        m_sURLDirectorioAplicaciones = m_URLRutaAplicacionesRemotas
    Else
        m_sURLDirectorioAplicaciones = m_URLRutaAplicacionesLocal
    End If
    If Not fso.FolderExists(m_sURLDirectorioAplicaciones) Then
        m_Error = "No es alcanzable m_sURLDirectorioAplicaciones " & m_sURLDirectorioAplicaciones
        Err.Raise 1000
    End If
    
    ' Normalizar: asegurar que termina en backslash
    If Right(m_sURLDirectorioAplicaciones, 1) <> "\" Then
        m_sURLDirectorioAplicaciones = m_sURLDirectorioAplicaciones & "\"
    End If
    
    URLDirectorioAplicaciones = m_sURLDirectorioAplicaciones
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirectorioAplicaciones ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


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

Public Property Get VerSoloRiesgosNoRetirados() As EnumSiNo
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eVerSoloRiesgosNoRetirados <> 0 Then
        VerSoloRiesgosNoRetirados = m_eVerSoloRiesgosNoRetirados
        Exit Property
    End If
    m_SQL = "SELECT * FROM TbConfiguracionVisionRiesgos " & _
            "WHERE Usuario='" & m_ObjUsuarioConectado.UsuarioRed & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If Nz(.Fields("VerSoloNoRetirados"), "") = "Sí" Then
                m_eVerSoloRiesgosNoRetirados = EnumSiNo.Sí
            ElseIf Nz(.Fields("VerSoloNoRetirados"), "") = "No" Then
                m_eVerSoloRiesgosNoRetirados = EnumSiNo.No
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    VerSoloRiesgosNoRetirados = m_eVerSoloRiesgosNoRetirados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.VerSoloRiesgosNoRetirados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get VerRiesgosDescripcion() As EnumSiNo
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eVerSoloRiesgosNoRetirados <> 0 Then
        VerRiesgosDescripcion = m_eVerRiesgosDescripcion
        Exit Property
    End If
    m_SQL = "SELECT * FROM TbConfiguracionVisionRiesgos " & _
            "WHERE Usuario='" & m_ObjUsuarioConectado.UsuarioRed & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If Nz(.Fields("VerDescripcion"), "") = "Sí" Then
                m_eVerRiesgosDescripcion = EnumSiNo.Sí
            ElseIf Nz(.Fields("VerDescripcion"), "") = "No" Then
                m_eVerRiesgosDescripcion = EnumSiNo.No
            End If
        End If
    End With
    
    VerRiesgosDescripcion = m_eVerRiesgosDescripcion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.VerRiesgosDescripcion ha devuelto el error: " & vbNewLine & Err.Description
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


Public Property Let VerSoloRiesgosNoRetirados(NuevoValor As EnumSiNo)
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ID As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If NuevoValor <> EnumSiNo.Sí And NuevoValor <> EnumSiNo.No Then
        m_SQL = "DELETE * FROM TbConfiguracionVisionRiesgos " & _
            "WHERE Usuario='" & m_ObjUsuarioConectado.UsuarioRed & "';"
        getdb().Execute m_SQL
        m_eVerSoloRiesgosNoRetirados = 0
        Exit Property
    End If
   
    
     m_SQL = "SELECT * FROM TbConfiguracionVisionRiesgos " & _
            "WHERE Usuario='" & m_ObjUsuarioConectado.UsuarioRed & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            m_ID = DameID("TbConfiguracionVisionRiesgos", "IDConfiguracion", getdb(), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            .AddNew
                .Fields("IDConfiguracion") = m_ID
                .Fields("Usuario") = m_ObjUsuarioConectado.UsuarioRed
                If NuevoValor = EnumSiNo.Sí Then
                    .Fields("VerSoloNoRetirados") = "Sí"
                Else
                    .Fields("VerSoloNoRetirados") = "No"
                End If
            .Update
        Else
            .Edit
                If NuevoValor = EnumSiNo.Sí Then
                    .Fields("VerSoloNoRetirados") = "Sí"
                Else
                    .Fields("VerSoloNoRetirados") = "No"
                End If
                
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_eVerSoloRiesgosNoRetirados = NuevoValor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.VerSoloRiesgosNoRetirados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Let VerRiesgosDescripcion(NuevoValor As EnumSiNo)
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ID As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If NuevoValor <> EnumSiNo.Sí And NuevoValor <> EnumSiNo.No Then
        m_SQL = "DELETE * FROM TbConfiguracionVisionRiesgos " & _
            "WHERE Usuario='" & m_ObjUsuarioConectado.UsuarioRed & "';"
        getdb().Execute m_SQL
        m_eVerRiesgosDescripcion = 0
        Exit Property
    End If
   
    
     m_SQL = "SELECT * FROM TbConfiguracionVisionRiesgos " & _
            "WHERE Usuario='" & m_ObjUsuarioConectado.UsuarioRed & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            m_ID = DameID("TbConfiguracionVisionRiesgos", "IDConfiguracion", getdb(), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            .AddNew
                .Fields("IDConfiguracion") = m_ID
                .Fields("Usuario") = m_ObjUsuarioConectado.UsuarioRed
                If NuevoValor = EnumSiNo.Sí Then
                    .Fields("VerDescripcion") = "Sí"
                Else
                    .Fields("VerDescripcion") = "No"
                End If
            .Update
        Else
            .Edit
                If NuevoValor = EnumSiNo.Sí Then
                    .Fields("VerDescripcion") = "Sí"
                Else
                    .Fields("VerDescripcion") = "No"
                End If
                
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_eVerRiesgosDescripcion = NuevoValor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.VerRiesgosDescripcion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get TituloUsuarioConectado() As String
    
    Dim m_Colectivo As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sTituloUsuarioConectado <> "" Then
        TituloUsuarioConectado = m_sTituloUsuarioConectado
        Exit Property
    End If
    
    If m_ObjEntorno.ColUsuariosAdministradores.Exists(m_ObjUsuarioConectado.UsuarioRed) Then
        m_Colectivo = "Administrador"
    ElseIf m_ObjEntorno.ColUsuariosCalidad.Exists(m_ObjUsuarioConectado.UsuarioRed) Then
        m_Colectivo = "Calidad"
    Else
        m_Colectivo = "Técnico"
    End If
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_sTituloUsuarioConectado = "Usuario conectado: " & m_ObjUsuarioConectado.Nombre & " ( " & m_Colectivo & " ) EN PRUEBAS"
    Else
        m_sTituloUsuarioConectado = "Usuario conectado: " & m_ObjUsuarioConectado.Nombre & " ( " & m_Colectivo & " )"
    End If
    
    TituloUsuarioConectado = m_sTituloUsuarioConectado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.TituloUsuarioConectado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get ColUsuariosAdministradores() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosAdministradores Is Nothing Then
        Set ColUsuariosAdministradores = m_objColUsuariosAdministradores
        Exit Property
    End If
    Set m_objColUsuariosAdministradores = Constructor.getUsuariosAdministradores(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosAdministradores = m_objColUsuariosAdministradores
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColUsuariosAdministradores ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColUsuariosAdministradores(valor As Variant)
    
    Set m_objColUsuariosAdministradores = valor
    
End Property

Public Property Get ColUsuariosCalidad() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosCalidad Is Nothing Then
        Set ColUsuariosCalidad = m_objColUsuariosCalidad
        Exit Property
    End If
    Set m_objColUsuariosCalidad = Constructor.getUsuariosCalidad(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosCalidad = m_objColUsuariosCalidad
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColUsuariosCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColUsuariosCalidad(valor As Variant)
    
    Set m_objColUsuariosCalidad = valor
    
End Property
Public Property Get ColUsuariosTecnicos() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosTecnicos Is Nothing Then
        Set ColUsuariosTecnicos = m_objColUsuariosTecnicos
        Exit Property
    End If
    Set m_objColUsuariosTecnicos = Constructor.getUsuariosTecnicos(p_SoloActivos:=EnumSiNo.No, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosTecnicos = m_objColUsuariosTecnicos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColUsuariosTecnicos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColUsuariosTecnicos(valor As Variant)
    
    Set m_objColUsuariosTecnicos = valor
    
End Property
Public Property Get ListaUsuarios() As String
    Dim m_Usuario As Variant
    Dim m_ObjUsuario As Usuario
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_ListaUsuarios <> "" Then
        ListaUsuarios = m_ListaUsuarios
        Exit Property
    End If
    For Each m_Usuario In Me.ColUsuariosTecnicosConAlgunaGestion
        If Nz(m_Usuario, "") <> "" Then
            Set m_ObjUsuario = m_ObjEntorno.ColUsuariosTecnicosConAlgunaGestion(m_Usuario)
            If Not m_ObjUsuario Is Nothing Then
                If m_ListaUsuarios = "" Then
                    m_ListaUsuarios = m_ObjUsuario.Nombre
                Else
                    m_ListaUsuarios = m_ListaUsuarios & ";" & m_ObjUsuario.Nombre
                End If
                
                Set m_ObjUsuario = Nothing
            End If
        End If
    Next
    ListaUsuarios = m_ListaUsuarios
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ListaUsuarios ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let ListaUsuarios(valor As String)
    
    m_ListaUsuarios = valor
    
End Property
Public Property Get ColUsuariosTecnicosActivos() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosTecnicosActivos Is Nothing Then
        Set ColUsuariosTecnicosActivos = m_objColUsuariosTecnicosActivos
        Exit Property
    End If
    Set m_objColUsuariosTecnicosActivos = Constructor.getUsuariosTecnicos(p_SoloActivos:=EnumSiNo.Sí, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosTecnicosActivos = m_objColUsuariosTecnicosActivos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColUsuariosTecnicosActivos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColUsuariosTecnicosActivos(valor As Variant)
    
    Set m_objColUsuariosTecnicosActivos = valor
    
End Property

Public Property Get ColUsuariosTecnicosConAlgunaGestion() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosTecnicosConAlgunaGestion Is Nothing Then
        Set ColUsuariosTecnicosConAlgunaGestion = m_objColUsuariosTecnicosConAlgunaGestion
        Exit Property
    End If
    Set m_objColUsuariosTecnicosConAlgunaGestion = Constructor.getUsuariosTecnicosConAlgunaGestion(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosTecnicosConAlgunaGestion = m_objColUsuariosTecnicosConAlgunaGestion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColUsuariosTecnicosConAlgunaGestion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColUsuariosTecnicosConAlgunaGestion(valor As Variant)
    
    Set m_objColUsuariosTecnicosConAlgunaGestion = valor
    
End Property


Public Property Get ColUsuariosCalidadAvisos() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosCalidadAvisos Is Nothing Then
        Set ColUsuariosCalidadAvisos = m_objColUsuariosCalidadAvisos
        Exit Property
    End If
    Set m_objColUsuariosCalidadAvisos = Constructor.getUsuariosCalidadAvisos(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosCalidadAvisos = m_objColUsuariosCalidadAvisos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColUsuariosCalidadAvisos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColUsuariosCalidadAvisos(valor As Variant)
    
    Set m_objColUsuariosCalidadAvisos = valor
    
End Property

Public Property Get ColJPs() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColJPs Is Nothing Then
        Set ColJPs = m_objColJPs
        Exit Property
    End If
    Set m_objColJPs = Constructor.getColJPs(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColJPs = m_objColJPs
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColJPs ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColJPs(valor As Variant)
    
    Set m_objColJPs = valor
    
End Property
Public Property Get ColVersiones() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColVersiones Is Nothing Then
        Set ColVersiones = m_objColVersiones
        Exit Property
    End If
    Set m_objColVersiones = Constructor.getCCVersiones(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColVersiones = m_objColVersiones
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColVersiones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get UsuarioConectadoEsAdministrador() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_eusuarioConectadoEsAdministrador <> 0 Then
         UsuarioConectadoEsAdministrador = m_eusuarioConectadoEsAdministrador
        Exit Property
    End If
    
    If Me.ColUsuariosAdministradores Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eusuarioConectadoEsAdministrador = EnumSiNo.No
        UsuarioConectadoEsAdministrador = m_eusuarioConectadoEsAdministrador
        Exit Property
    End If
    If Me.ColUsuariosAdministradores.Exists(m_ObjUsuarioConectado.UsuarioRed) Then
        m_eusuarioConectadoEsAdministrador = EnumSiNo.Sí
    Else
        m_eusuarioConectadoEsAdministrador = EnumSiNo.No
    End If
    UsuarioConectadoEsAdministrador = m_eusuarioConectadoEsAdministrador
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método General.UsuarioConectadoEsAdministrador ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get UsuarioConectadoEsDeCalidad() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_eusuarioConectadoEsDeCalidad <> 0 Then
        UsuarioConectadoEsDeCalidad = m_eusuarioConectadoEsDeCalidad
        Exit Property
    End If
    If Me.ColUsuariosCalidad Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eusuarioConectadoEsDeCalidad = EnumSiNo.No
        UsuarioConectadoEsDeCalidad = m_eusuarioConectadoEsDeCalidad
        Exit Property
    End If
    If Me.ColUsuariosCalidad.Exists(m_ObjUsuarioConectado.UsuarioRed) Then
        m_eusuarioConectadoEsDeCalidad = EnumSiNo.Sí
    Else
        m_eusuarioConectadoEsDeCalidad = EnumSiNo.No
    End If
    UsuarioConectadoEsDeCalidad = m_eusuarioConectadoEsDeCalidad
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método General.usuarioConectadoEsDeCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get UsuarioConectadoEsTecnico() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_eusuarioConectadoEsTecnico <> 0 Then
        UsuarioConectadoEsTecnico = m_eusuarioConectadoEsTecnico
        Exit Property
    End If
    If Me.ColUsuariosTecnicos Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eusuarioConectadoEsTecnico = EnumSiNo.No
        UsuarioConectadoEsTecnico = m_eusuarioConectadoEsTecnico
        Exit Property
    End If
    If Me.ColUsuariosTecnicos.Exists(m_ObjUsuarioConectado.UsuarioRed) Then
        m_eusuarioConectadoEsTecnico = EnumSiNo.Sí
    Else
        m_eusuarioConectadoEsTecnico = EnumSiNo.No
    End If
    UsuarioConectadoEsTecnico = m_eusuarioConectadoEsTecnico
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método General.UsuarioConectadoEsTecnico ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get CodigoBibliotecaParaRetipificar() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sCodigoBibliotecaParaRetipificar <> "" Then
        CodigoBibliotecaParaRetipificar = m_sCodigoBibliotecaParaRetipificar
        Exit Property
    End If
    m_sCodigoBibliotecaParaRetipificar = "R-PR-30"
    CodigoBibliotecaParaRetipificar = m_sCodigoBibliotecaParaRetipificar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CodigoBibliotecaParaRetipificar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColDatosPlazoCalidadCosteVulnerabilidad() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColDatosPlazoCalidadCosteVulnerabilidad Is Nothing Then
        Set ColDatosPlazoCalidadCosteVulnerabilidad = m_objColDatosPlazoCalidadCosteVulnerabilidad
        Exit Property
    End If
    Set m_objColDatosPlazoCalidadCosteVulnerabilidad = New Scripting.Dictionary
    m_objColDatosPlazoCalidadCosteVulnerabilidad.CompareMode = TextCompare
    With m_objColDatosPlazoCalidadCosteVulnerabilidad
        .Add "1", "Muy Bajo"
        .Add "2", "Bajo"
        .Add "3", "Medio"
        .Add "4", "Alto"
        .Add "5", "Muy Alto"
    End With
    Set ColDatosPlazoCalidadCosteVulnerabilidad = m_objColDatosPlazoCalidadCosteVulnerabilidad
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColDatosPlazoCalidadCosteVulnerabilidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColMitigacionValores() As Scripting.Dictionary
        
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColMitigacionValores Is Nothing Then
        Set ColMitigacionValores = m_objColMitigacionValores
        Exit Property
    End If
    Set m_objColMitigacionValores = New Scripting.Dictionary
    m_objColMitigacionValores.CompareMode = TextCompare
    With m_objColMitigacionValores
        .Add "Aceptar", "Cuando decidimos Aceptar el riesgo y no hacer nada. Requiere una autorización por parte del RAC que se lleva a cabo por medio de Calidad"
        .Add "Evitar", "Cuando vamos a orientar todos nuestros planes futuros para intentar evitar el Riesgo."
        .Add "Reducir", "Cuando vamos a orientar todos nuestros planes futuros para intentar reducir las posibilidades de que el Riesgo se produzca"
        .Add "Transferir", "Cuando decidimos que la gestión del trabajo con este riesgo lo va a llevar otra empresa"
        
    End With
    Set ColMitigacionValores = m_objColMitigacionValores
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColMitigacionValores ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColOrigenRiesgosValores() As Scripting.Dictionary
        
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColOrigenRiesgosValores Is Nothing Then
        Set ColOrigenRiesgosValores = m_objColOrigenRiesgosValores
        Exit Property
    End If
    Set m_objColOrigenRiesgosValores = New Scripting.Dictionary
    m_objColOrigenRiesgosValores.CompareMode = TextCompare
    With m_objColOrigenRiesgosValores
        .Add "Ejecución", "Son los riesgos detectados por Telefónica durante la propia ejecución del proyecto"
        .Add "Oferta", "Son los riesgos detectados durante la fase de oferta del proyecto"
        .Add "Pedido", "Son riesgos puntuales derivados de un pedido en específico"
        .Add "Suministrador", "Son los riesgos detectados por un suministrador y que incorporamos a nuestra propia gestión de riesgos por considerarlo relevantes para el proyecto"
        .Add "UTE", "Son los riesgos detectados por empresas miembros de la UTE que no sean Telefónica"
    End With
    Set ColOrigenRiesgosValores = m_objColOrigenRiesgosValores
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColOrigenRiesgosValores ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get CorreoCalidad() As String
    
    Dim m_Usuario As Usuario
    Dim m_ID As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sCorreoCalidad <> "" Then
        CorreoCalidad = m_sCorreoCalidad
        Exit Property
    End If
    If Me.ColUsuariosCalidad Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_ID In Me.ColUsuariosCalidad.keys
        If IsObject(Me.ColUsuariosCalidad(m_ID)) Then
            Set m_Usuario = Me.ColUsuariosCalidad(m_ID)
            If m_Usuario.CorreoUsuario <> "" Then
                If m_sCorreoCalidad = "" Then
                    m_sCorreoCalidad = m_Usuario.CorreoUsuario
                Else
                    m_sCorreoCalidad = m_sCorreoCalidad & ";" & m_Usuario.CorreoUsuario
                End If
            End If
        End If
    Next
    CorreoCalidad = m_sCorreoCalidad
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CorreoCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get URLArchivoCSS() As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLArchivoCSS <> "" Then
        URLArchivoCSS = m_sURLArchivoCSS
        Exit Property
    End If
    m_sURLArchivoCSS = Me.URLDirectorioAplicaciones & "CSS.txt"
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

Public Property Get css() As String
    
    Dim sCorp As String
    Dim p1 As Long
    Dim p2 As Long
    Dim pClose As Long
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sCSS <> "" Then
        css = m_sCSS
        Exit Property
    End If
    sCorp = GetEstilosCSS_Corporativos()
    p1 = InStr(1, sCorp, "<style", vbTextCompare)
    If p1 > 0 Then
        pClose = InStr(p1, sCorp, ">", vbTextCompare)
        p2 = InStr(1, sCorp, "</style>", vbTextCompare)
        If pClose > 0 And p2 > pClose Then
            m_sCSS = Mid$(sCorp, pClose + 1, p2 - (pClose + 1))
        End If
    End If
    If m_sCSS = "" Then
        m_sCSS = Replace(Replace(sCorp, "<style>", ""), "</style>", "")
    End If
    css = m_sCSS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CSS ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get CabeceraHTML() As String
    
    'variable #titulo
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sCabeceraHTML <> "" Then
        CabeceraHTML = m_sCabeceraHTML
        Exit Property
    End If
    m_sCabeceraHTML = "<!DOCTYPE html>" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<html lang=""es"">" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<head>" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<meta charset=""UTF-8"">" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<title>#titulo</title>" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<style type=""text/css"">" & vbNewLine
        m_sCabeceraHTML = m_sCabeceraHTML & m_ObjEntorno.css & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "</style>" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "</head>" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<body>" & vbNewLine

    CabeceraHTML = m_sCabeceraHTML
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CabeceraHTML ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FinHTML() As String
    
    'variable #titulo
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sFinHTML <> "" Then
        FinHTML = m_sFinHTML
        Exit Property
    End If
    m_sFinHTML = m_sFinHTML & "</body>" & vbNewLine
    m_sFinHTML = m_sFinHTML & "</html>" & vbNewLine
    FinHTML = m_sFinHTML
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.FinHTML ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColDatosVulnerabilidades() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColDatosVulnerabilidades Is Nothing Then
        Set ColDatosVulnerabilidades = m_objColDatosVulnerabilidades
        Exit Property
    End If
    Set m_objColDatosVulnerabilidades = New Scripting.Dictionary
    m_objColDatosVulnerabilidades.CompareMode = TextCompare
    With m_objColDatosVulnerabilidades
        .Add "Muy Bajo", "Muy Bajo(<=25%)"
        .Add "Bajo", "Bajo(>=25% y <45%)"
        .Add "Medio", "Medio(>=45% y <65%)"
        .Add "Alto", "Alto(>=65% y <85%)"
        .Add "Muy Alto", "Muy Alto(>=85%)"
    End With
    Set ColDatosVulnerabilidades = m_objColDatosVulnerabilidades
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColDatosVulnerabilidades ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColProyectosTotales() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColProyectosTotales Is Nothing Then
        Set ColProyectosTotales = m_ObjColProyectosTotales
        Exit Property
    End If
    
    Set m_ObjColProyectosTotales = Constructor.getProyectos(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColProyectosTotales = m_ObjColProyectosTotales
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColProyectosTotales ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Property Set ColProyectosTotales(p_Valor As Variant)
    Set m_ObjColProyectosTotales = p_Valor
End Property
Public Property Get ColRiesgosBibliotecasFamilias() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColRiesgosBibliotecasFamilias Is Nothing Then
        Set ColRiesgosBibliotecasFamilias = m_ObjColRiesgosBibliotecasFamilias
        Exit Property
    End If
    
    Set m_ObjColRiesgosBibliotecasFamilias = Constructor.getRiesgosBibliotecasFamilias(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColRiesgosBibliotecasFamilias = m_ObjColRiesgosBibliotecasFamilias
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColRiesgosBibliotecasFamilias ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Property Set ColRiesgosBibliotecasFamilias(p_Valor As Variant)
    Set m_ObjColRiesgosBibliotecasFamilias = p_Valor
End Property
Public Property Get DestinatariosCorreoAdministradores() As String
    
    Dim m_UsuarioRed As Variant
    Dim m_ObjUsuario As Usuario
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sDestinatariosCorreoAdministradores <> "" Then
        DestinatariosCorreoAdministradores = m_sDestinatariosCorreoAdministradores
        Exit Property
    End If
    If Me.ColUsuariosAdministradores Is Nothing Then
        Exit Property
    End If
    For Each m_UsuarioRed In Me.ColUsuariosAdministradores.keys
        If Nz(m_UsuarioRed, "") <> "" Then
            Set m_ObjUsuario = Me.ColUsuariosAdministradores(m_UsuarioRed)
            If m_ObjUsuario.CorreoUsuario <> "" Then
                If m_sDestinatariosCorreoAdministradores = "" Then
                    m_sDestinatariosCorreoAdministradores = m_ObjUsuario.CorreoUsuario
                Else
                    m_sDestinatariosCorreoAdministradores = m_sDestinatariosCorreoAdministradores & ";" & m_ObjUsuario.CorreoUsuario
                End If
            End If
            Set m_ObjUsuario = Nothing
        End If
    Next
    DestinatariosCorreoAdministradores = m_sDestinatariosCorreoAdministradores
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.DestinatariosCorreoAdministradores ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get ColExplicacionesArbolTareas() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColExplicacionesArbolTareas Is Nothing Then
        Set ColExplicacionesArbolTareas = m_objColExplicacionesArbolTareas
        Exit Property
    End If
    
    Set m_objColExplicacionesArbolTareas = Constructor.getListaExplicaciones(m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColExplicacionesArbolTareas = m_objColExplicacionesArbolTareas
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColExplicacionesArbolTareas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get ColTipoTareas() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColTipoTareas Is Nothing Then
        Set ColTipoTareas = m_objColTipoTareas
        Exit Property
    End If
    Set m_objColTipoTareas = New Scripting.Dictionary
    m_objColTipoTareas.CompareMode = TextCompare
    With m_objColTipoTareas
        .Add CStr(EnumTipoTareas.ParaAceptadosRetirados), "Aceptados Retirados"
        .Add CStr(EnumTipoTareas.ParaMaterializados), "Materializados"
        .Add CStr(EnumTipoTareas.ParaProyectos), "Proyectos"
        .Add CStr(EnumTipoTareas.ParaRetipificados), "Retipificados"
        .Add CStr(EnumTipoTareas.ParaPreparadasParaPublicar), "Preparados para Publicar"
        
    End With
    Set ColTipoTareas = m_objColTipoTareas
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColTipoTareas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColTareasTipo() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColTareasTipo Is Nothing Then
        Set ColTareasTipo = m_objColTareasTipo
        Exit Property
    End If
    Set m_objColTareasTipo = New Scripting.Dictionary
    With m_objColTareasTipo
        .Add "Aceptados Retirados", EnumTipoTareas.ParaAceptadosRetirados
        .Add "Materializados", EnumTipoTareas.ParaMaterializados
        .Add "Proyectos", EnumTipoTareas.ParaProyectos
        .Add "Retipificados", EnumTipoTareas.ParaRetipificados
        .Add "Preparados para Publicar", EnumTipoTareas.ParaPreparadasParaPublicar
    End With
    Set ColTareasTipo = m_objColTareasTipo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColTareasTipo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColImpactoNumericoParaGraficas() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColImpactoNumericoParaGraficas Is Nothing Then
        Set ColImpactoNumericoParaGraficas = m_objColImpactoNumericoParaGraficas
        Exit Property
    End If
    Set m_objColImpactoNumericoParaGraficas = New Scripting.Dictionary
    With m_objColImpactoNumericoParaGraficas
        .Add "Acept./Retir.", 0
        .Add "Muy Bajo", 1
        .Add "Bajo", 2
        .Add "Medio", 3
        .Add "Alto", 4
        .Add "Muy Alto", 5
    End With
    Set ColImpactoNumericoParaGraficas = m_objColImpactoNumericoParaGraficas
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColImpactoNumericoParaGraficas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get ColEstadoNumericoParaGraficas() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColEstadoNumericoParaGraficas Is Nothing Then
        Set ColEstadoNumericoParaGraficas = m_objColEstadoNumericoParaGraficas
        Exit Property
    End If
    Set m_objColEstadoNumericoParaGraficas = New Scripting.Dictionary
    m_objColEstadoNumericoParaGraficas.CompareMode = TextCompare
    With m_objColEstadoNumericoParaGraficas
        .Add "Detectado", 1
        .Add "Planificado", 2
        .Add "Activo", 3
        .Add "Materializado", 4
        .Add "Aceptado", 5
        .Add "Retirado", 6
        .Add "Cerrado", 7
        
    End With
    Set ColEstadoNumericoParaGraficas = m_objColEstadoNumericoParaGraficas
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEstadoNumericoParaGraficas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get ColAreaImpactoCalidad() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColAreaImpactoCalidad Is Nothing Then
        Set ColAreaImpactoCalidad = m_objColAreaImpactoCalidad
        Exit Property
    End If
    
    Set m_objColAreaImpactoCalidad = Constructor.getListaImpactosPorTipo("Calidad", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAreaImpactoCalidad = m_objColAreaImpactoCalidad
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColAreaImpactoCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get ColAreaImpactoPlazo() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColAreaImpactoPlazo Is Nothing Then
        Set ColAreaImpactoPlazo = m_objColAreaImpactoPlazo
        Exit Property
    End If
    
    Set m_objColAreaImpactoPlazo = Constructor.getListaImpactosPorTipo("Plazo", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAreaImpactoPlazo = m_objColAreaImpactoPlazo
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColAreaImpactoPlazo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColAreaImpactoCoste() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColAreaImpactoCoste Is Nothing Then
        Set ColAreaImpactoCoste = m_objColAreaImpactoCoste
        Exit Property
    End If
    
    Set m_objColAreaImpactoCoste = Constructor.getListaImpactosPorTipo("Coste", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAreaImpactoCoste = m_objColAreaImpactoCoste
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColAreaImpactoCoste ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get ColEdicionesCaducadas() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColEdicionesCaducadas Is Nothing Then
        Set ColEdicionesCaducadas = m_objColEdicionesCaducadas
        Exit Property
    End If
    Set m_objColEdicionesCaducadas = Constructor.getEdicionesCaducadas(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColEdicionesCaducadas = m_objColEdicionesCaducadas
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.colEdicionesCaducadas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set ColEdicionesCaducadas(ByVal objNewValue As Variant)

    Set m_objColEdicionesCaducadas = objNewValue

End Property
Public Property Get ColEdicionesApuntoDeCaducar() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColEdicionesAPuntoDeCaducar Is Nothing Then
        Set ColEdicionesApuntoDeCaducar = m_objColEdicionesAPuntoDeCaducar
        Exit Property
    End If
    Set m_objColEdicionesAPuntoDeCaducar = Constructor.getEdicionesApuntoDeCaducar(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColEdicionesApuntoDeCaducar = m_objColEdicionesAPuntoDeCaducar
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEdicionesApuntoDeCaducar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set ColEdicionesApuntoDeCaducar(ByVal objNewValue As Variant)

    Set m_objColEdicionesAPuntoDeCaducar = objNewValue

End Property
Public Property Get RiesgoBibliotecaNoExiste() As RiesgoBiblioteca

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objRiesgoBibliotecaNoExiste Is Nothing Then
        Set RiesgoBibliotecaNoExiste = m_objRiesgoBibliotecaNoExiste
        Exit Property
    End If
    Set m_objRiesgoBibliotecaNoExiste = Constructor.getRiesgoBibliotecaNoExiste(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_objRiesgoBibliotecaNoExiste Is Nothing Then
        m_Error = "No se ha podido determinar el riesgo que indica que no existe en biblioteca"
        Err.Raise 1000
    End If
    Set RiesgoBibliotecaNoExiste = m_objRiesgoBibliotecaNoExiste
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.RiesgoBibliotecaNoExiste  ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get ColCambiosExplicaciones() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColCambiosExplicaciones Is Nothing Then
        Set ColCambiosExplicaciones = m_objColCambiosExplicaciones
        Exit Property
    End If
    
    Set m_objColCambiosExplicaciones = Constructor.getCambiosExplicaciones(m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColCambiosExplicaciones = m_objColCambiosExplicaciones
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColCambiosExplicaciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get ColCarenciasExplicaciones() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColCarenciasExplicaciones Is Nothing Then
        Set ColCarenciasExplicaciones = m_objColCarenciasExplicaciones
        Exit Property
    End If
    
    Set m_objColCarenciasExplicaciones = Constructor.getCarenciasExplicaciones(m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColCarenciasExplicaciones = m_objColCarenciasExplicaciones
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColCarenciasExplicaciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get VersionesAplicacion() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objVersionesAplicacion Is Nothing Then
        Set VersionesAplicacion = m_objVersionesAplicacion
        Exit Property
    End If
    
    Set m_objVersionesAplicacion = Constructor.getVersionesAplicacion(m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set VersionesAplicacion = m_objVersionesAplicacion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.VersionesAplicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get ColTipos() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColTipos Is Nothing Then
        Set ColTipos = m_objColTipos
        Exit Property
    End If
    Set m_objColTipos = Constructor.getTiposNC(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Set ColTipos = m_objColTipos
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColTipos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property



Public Property Get ColImagenes() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColImagenes Is Nothing Then
        Set ColImagenes = m_objColImagenes
        Exit Property
    End If
    Set m_objColImagenes = New Scripting.Dictionary
    m_objColImagenes.CompareMode = TextCompare
    With m_objColImagenes
        .Add m_ObjEntorno.URLIconoCarpetaCompletaAbierta32, m_ObjEntorno.URLIconoCarpetaCompletaAbierta32
        .Add m_ObjEntorno.URLIconoCarpetaCompletaCerrada32, m_ObjEntorno.URLIconoCarpetaCompletaCerrada32
        .Add m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32, m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32
        .Add m_ObjEntorno.URLIconoCarpetaIncompletaCerrada32, m_ObjEntorno.URLIconoCarpetaIncompletaCerrada32
        .Add m_ObjEntorno.URLIconoItemInCompleto32, m_ObjEntorno.URLIconoItemInCompleto32
        .Add m_ObjEntorno.URLIconoItemCompleto32, m_ObjEntorno.URLIconoItemCompleto32
    End With
    Set ColImagenes = m_objColImagenes
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColImagenes ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColItems() As Scripting.Dictionary

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColItems Is Nothing Then
        Set ColItems = m_objColItems
        Exit Property
    End If
    Set m_objColItems = New Scripting.Dictionary
    With m_objColItems
        
        
        .Add "URLDirectorioAplicaciones", ""
        .Add "URLDirectorioAplicacion", ""
        .Add "URLDirectorioTemporalAGEDO", ""
        .Add "URLDirectorioDocumentacion", ""
        .Add "URLDirectorioAyuda", ""
        .Add "URLDirectorioRecursos", ""
        .Add "URLPlantillaExcel", ""
        .Add "URLPlantillaEvolucionRiesgosExcel", ""
        .Add "URLAchivoIni", ""
        .Add "JPMesesAvisoEntreEdiciones", ""
        .Add "JPDiasPreviosParaElAviso", ""
        .Add "CalDiaInicialMesAviso", ""
        
        .Add "URLDirectorioLocal", ""
        .Add "VersionAplicacion", ""
        .Add "URLUltimoArchivo", ""
        .Add "OrdenAscendenteEnListaProyectos", "s"
        .Add "ValorAlarma", ""
        .Add "ValorObjetivo", ""
        .Add "URLIconoCarpetaCompletaCerrada32", ""
        .Add "URLIconoCarpetaCompletaAbierta32", ""
        .Add "URLIconoCarpetaIncompletaCerrada32", ""
        .Add "URLIconoCarpetaIncompletaAbierta32", ""
        .Add "URLIconoItemInCompleto32", ""
        .Add "URLIconoItemCompleto32", ""
        
        .Add "URLDirIconos", ""
        .Add "URLIconoProyecto", ""
        .Add "TituloUsuarioConectado", ""
        .Add "UsuarioConectadoEsAdministrador", "s"
        .Add "UsuarioConectadoEsDeCalidad", "s"
        .Add "UsuarioConectadoEsTecnico", "s"
        .Add "CorreoCalidad", ""
        .Add "CSS", ""
       
        .Add "ColMitigacionValores", "o"
        .Add "ColOrigenRiesgosValores", "o"
        .Add "ColDatosVulnerabilidades", "o"
        
        
        .Add "CabeceraHTML", ""
        .Add "FinHTML", ""
        .Add "ColProyectosTotales", "o"
        .Add "ColTipoTareas", "o"
        .Add "ColAreaImpactoCalidad", "o"
        .Add "ColAreaImpactoCoste", "o"
        .Add "ColAreaImpactoPlazo", "o"
        .Add "DestinatariosCorreoAdministradores", ""
        .Add "ColUsuariosAdministradores", "o"
        .Add "ColUsuariosCalidad", "o"
        
        
        
        
        
        .Add "ColTareasTipo", "o"
        .Add "ColImpactoNumericoParaGraficas", "o"
        .Add "ColEstadoNumericoParaGraficas", "o"
        .Add "ColExplicacionesArbolTareas", "o"
        
        .Add "URLIconoOK", ""
        .Add "URLIconoNoOK", ""
        .Add "RiesgoBibliotecaNoExiste", "o"
        
        .Add "URLArchivoSSID", ""
        .Add "ColRiesgosBibliotecasFamilias", "o"
        .Add "ColImagenes", "o"
        .Add "URLArchivoCSS", ""
        .Add "URLDirectorioDocumentosCambio", ""
        .Add "URLDirectorioDocumentosRequisitosCambio", ""
        .Add "URLDirectorioAplicacionesRecursos", ""
        .Add "URLPlantillaWord", ""
        
        .Add "URLIconoPDF", ""
        .Add "ColEstados", "o"
        .Add "ColEstadosTexto", "o"
        .Add "ColEstadosTitulo", "o"
        .Add "ColEstadosTituloMejorado", "o"
        .Add "URLVideo", ""
    End With
    Set ColItems = m_objColItems
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColItems ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(p_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant
    
    
    On Error GoTo errores
    Me.Error = ""
    
    
    If p_NombrePropiedad = "URLDirectorioAplicaciones" Then
        getPropiedad = Me.URLDirectorioAplicaciones
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioAplicacion" Then
        getPropiedad = Me.URLDirectorioAplicacion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioTemporalAGEDO" Then
        getPropiedad = Me.URLDirectorioTemporalAGEDO
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioDocumentacion" Then
        getPropiedad = Me.URLDirectorioDocumentacion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioAyuda" Then
        getPropiedad = Me.URLDirectorioAyuda
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioRecursos" Then
        getPropiedad = Me.URLDirectorioRecursos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLPlantillaExcel" Then
        getPropiedad = Me.URLPlantillaExcel
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLPlantillaWord" Then
        getPropiedad = Me.URLPlantillaWord
        Exit Function
    End If
    
    If p_NombrePropiedad = "URLPlantillaEvolucionRiesgosExcel" Then
        getPropiedad = Me.URLPlantillaEvolucionRiesgosExcel
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLAchivoIni" Then
        getPropiedad = Me.URLAchivoIni
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "JPMesesAvisoEntreEdiciones" Then
        getPropiedad = Me.JPMesesAvisoEntreEdiciones
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "JPDiasPreviosParaElAviso" Then
        getPropiedad = Me.JPDiasPreviosParaElAviso
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "CalDiaInicialMesAviso" Then
        getPropiedad = Me.CalDiaInicialMesAviso
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioAyuda" Then
        getPropiedad = Me.URLDirectorioAyuda
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioLocal" Then
        getPropiedad = Me.URLDirectorioLocal
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "VersionAplicacion" Then
        getPropiedad = Me.VersionAplicacion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "URLUltimoArchivo" Then
        getPropiedad = Me.URLUltimoArchivo
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "OrdenAscendenteEnListaProyectos" Then
        getPropiedad = Me.OrdenAscendenteEnListaProyectos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ValorAlarma" Then
        getPropiedad = Me.ValorAlarma
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "ValorObjetivo" Then
        getPropiedad = Me.ValorObjetivo
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLIconoCarpetaCompletaCerrada32" Then
        getPropiedad = Me.URLIconoCarpetaCompletaCerrada32
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLIconoCarpetaCompletaAbierta32" Then
        getPropiedad = Me.URLIconoCarpetaCompletaAbierta32
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLIconoCarpetaIncompletaCerrada32" Then
        getPropiedad = Me.URLIconoCarpetaIncompletaCerrada32
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLIconoCarpetaIncompletaAbierta32" Then
        getPropiedad = Me.URLIconoCarpetaIncompletaAbierta32
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLIconoItemInCompleto32" Then
        getPropiedad = Me.URLIconoItemInCompleto32
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLIconoItemCompleto32" Then
        getPropiedad = Me.URLIconoItemCompleto32
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    
    If p_NombrePropiedad = "URLDirIconos" Then
        getPropiedad = Me.URLDirIconos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLIconoProyecto" Then
        getPropiedad = Me.URLIconoProyecto
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "TituloUsuarioConectado" Then
        getPropiedad = Me.TituloUsuarioConectado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "UsuarioConectadoEsAdministrador" Then
        getPropiedad = Me.UsuarioConectadoEsAdministrador
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "UsuarioConectadoEsDeCalidad" Then
        getPropiedad = Me.UsuarioConectadoEsDeCalidad
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "UsuarioConectadoEsTecnico" Then
        getPropiedad = Me.UsuarioConectadoEsTecnico
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "CorreoCalidad" Then
        getPropiedad = Me.CorreoCalidad
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "CSS" Then
        getPropiedad = Me.css
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    
    If p_NombrePropiedad = "ColMitigacionValores" Then
        Set getPropiedad = Me.ColMitigacionValores
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColOrigenRiesgosValores" Then
        Set getPropiedad = Me.ColOrigenRiesgosValores
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColDatosVulnerabilidades" Then
        Set getPropiedad = Me.ColDatosVulnerabilidades
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    
    
    If p_NombrePropiedad = "CabeceraHTML" Then
        getPropiedad = Me.CabeceraHTML
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "FinHTML" Then
        getPropiedad = Me.FinHTML
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColProyectosTotales" Then
        Set getPropiedad = Me.ColProyectosTotales
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColTipoTareas" Then
        Set getPropiedad = Me.ColTipoTareas
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColAreaImpactoCalidad" Then
        Set getPropiedad = Me.ColAreaImpactoCalidad
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColAreaImpactoCoste" Then
        Set getPropiedad = Me.ColAreaImpactoCoste
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColAreaImpactoPlazo" Then
        Set getPropiedad = Me.ColAreaImpactoPlazo
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "DestinatariosCorreoAdministradores" Then
        getPropiedad = Me.DestinatariosCorreoAdministradores
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColUsuariosAdministradores" Then
        Set getPropiedad = Me.ColUsuariosAdministradores
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColUsuariosCalidad" Then
        Set getPropiedad = Me.ColUsuariosCalidad
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
   
    
    
    
    
    
    If p_NombrePropiedad = "ColTareasTipo" Then
        Set getPropiedad = Me.ColTareasTipo
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColImpactoNumericoParaGraficas" Then
        Set getPropiedad = Me.ColImpactoNumericoParaGraficas
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColEstadoNumericoParaGraficas" Then
        Set getPropiedad = Me.ColEstadoNumericoParaGraficas
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColExplicacionesArbolTareas" Then
        Set getPropiedad = Me.ColExplicacionesArbolTareas
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "URLIconoOK" Then
        getPropiedad = Me.URLIconoOK
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLIconoNoOK" Then
        getPropiedad = Me.URLIconoNoOK
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "RiesgoBibliotecaNoExiste" Then
        Set getPropiedad = Me.RiesgoBibliotecaNoExiste
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "URLArchivoSSID" Then
        getPropiedad = Me.URLArchivoSSID
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColRiesgosBibliotecasFamilias" Then
        Set getPropiedad = Me.ColRiesgosBibliotecasFamilias
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColImagenes" Then
        Set getPropiedad = Me.ColImagenes
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLArchivoCSS" Then
        getPropiedad = Me.URLArchivoCSS
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioDocumentosCambio" Then
        getPropiedad = Me.URLDirectorioDocumentosCambio
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioDocumentosRequisitosCambio" Then
        getPropiedad = Me.URLDirectorioDocumentosRequisitosCambio
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirectorioAplicacionesRecursos" Then
        getPropiedad = Me.URLDirectorioAplicacionesRecursos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLIconoPDF" Then
        getPropiedad = Me.URLIconoPDF
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "ColEstados" Then
        Set getPropiedad = Me.ColEstados
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColEstadosTexto" Then
        Set getPropiedad = Me.ColEstadosTexto
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColEstadosTitulo" Then
        Set getPropiedad = Me.ColEstadosTitulo
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColEstadosTituloMejorado" Then
        Set getPropiedad = Me.ColEstadosTituloMejorado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLVideo" Then
        getPropiedad = Me.URLVideo
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    p_Error = "La propiedad " & p_NombrePropiedad & " no está codificada en Entorno.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Entorno.getPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

' =============================================================================
' ValidarRecursos — valida todos los recursos críticos y retorna errors consolidado
' Called by EVE() in Variables Globales to collect ALL resource failures at once
' =============================================================================
Public Function ValidarRecursos() As String
    Dim sb As String
    sb = ""
    Dim test As String
    Dim errTemp As String
    
    On Error Resume Next
    
    ' Iconos
    test = Me.URLDirIconos
    errTemp = m_Error
    If errTemp <> "" Then
        sb = sb & "- Directorio de iconos: " & errTemp & vbCrLf
        m_Error = ""
    End If
    
    ' Recursos
    test = Me.URLDirectorioRecursos
    errTemp = m_Error
    If errTemp <> "" Then
        sb = sb & "- Recursos: " & errTemp & vbCrLf
        m_Error = ""
    End If
    
    ' Documentacion (ANEXOS)
    test = Me.URLDirectorioDocumentacion
    errTemp = m_Error
    If errTemp <> "" Then
        sb = sb & "- Documentacion: " & errTemp & vbCrLf
        m_Error = ""
    End If
    
    ' Ayuda (DOC AYUDA)
    test = Me.URLDirectorioAyuda
    errTemp = m_Error
    If errTemp <> "" Then
        sb = sb & "- Ayuda: " & errTemp & vbCrLf
        m_Error = ""
    End If
    
    ' Directorio aplicacion
    test = Me.URLDirectorioAplicacion
    errTemp = m_Error
    If errTemp <> "" Then
        sb = sb & "- Directorio aplicacion: " & errTemp & vbCrLf
        m_Error = ""
    End If
    
    ' 00Recursos (AplicacionesRecursos)
    test = Me.URLDirectorioAplicacionesRecursos
    errTemp = m_Error
    If errTemp <> "" Then
        sb = sb & "- Recursos de aplicaciones (00Recursos): " & errTemp & vbCrLf
        m_Error = ""
    End If
    
    On Error GoTo 0
    
    ValidarRecursos = sb
End Function



