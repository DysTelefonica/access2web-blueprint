VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private fso As New Scripting.FileSystemObject

Public IDParte As String
Public FechaObtencion As String
Public FechaInicial As String
Public FechaFinal As String
Public Descripcion As String
Public NombreArchivoAdjunto As String
Public IDFactura As String
Public Centro As String

Private m_sURLParte As String
Private m_ObjColEventos As Scripting.Dictionary
Private m_ObjFactura As Factura
Private m_sIDParteCalculado As String
Private m_eAnexoCerrado As EnumSino
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get URLParte() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLParte <> "" Then
        URLParte = m_sURLParte
        Exit Property
    End If
    If Me.NombreArchivoAdjunto = "" Then
        Exit Property
    End If
    If m_ObjEntorno.URLDirectorioPartesFirmados = "" Then
        Exit Property
    End If
    
    m_sURLParte = m_ObjEntorno.URLDirectorioPartesFirmados & Me.NombreArchivoAdjunto
    
    URLParte = m_sURLParte
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Parte.URLParte ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error


End Property
Public Property Let URLParte(Valor As String)
    m_sURLParte = Valor
    
End Property
Public Property Get colEventos() As Scripting.Dictionary

    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not m_ObjColEventos Is Nothing Then
        Set colEventos = m_ObjColEventos
        Exit Property
    End If
    If Me.IDParte = "" Then
       Exit Property
    End If
    Set m_ObjColEventos = Constructor.getEventosEnParte(Me.IDParte, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colEventos = m_ObjColEventos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Parte.colEventos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set colEventos(Valor As Variant)
    Set m_ObjColEventos = Valor
End Property


Public Property Get Factura() As Factura

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjFactura Is Nothing Then
        Set Factura = m_ObjFactura
        Exit Property
    End If
    If Me.IDParte = "" Then
        Exit Property
    End If
    Set m_ObjFactura = Constructor.getFactura(Me.IDFactura, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Factura = m_ObjFactura
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Parte.Factura ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set Factura(ByVal objNewValue As Factura)

    Set m_ObjFactura = objNewValue
    
End Property

Public Property Get AnexoCerrado() As EnumSino
    
    Dim m_URL As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_eAnexoCerrado <> 0 Then
        AnexoCerrado = m_eAnexoCerrado
        Exit Property
    End If
    m_URL = Me.URLParte
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If fso.FileExists(m_URL) Then
        If FicheroAbierto(m_URL) Then
            m_eAnexoCerrado = EnumSino.No
        Else
            m_eAnexoCerrado = EnumSino.Sí
        End If
    Else
        m_eAnexoCerrado = EnumSino.Sí
    End If
    AnexoCerrado = m_eAnexoCerrado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Parte.AnexoCerrado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Function AnexarParte( _
                            p_URLLocal As String, _
                            Optional ByRef p_Error As String _
                            ) As String
   
    Dim p_URLAnexoFinal As String
    On Error GoTo errores
    
    If Not fso.FileExists(p_URLLocal) Then
        p_Error = "No es alcanzable el anexo local"
        Err.Raise 1000
    End If
    If FicheroAbierto(p_URLLocal) Then
        p_Error = "Ha de cerrar el anexo primero"
        Err.Raise 1000
    End If
    If Me.IDParte = "" Then
        p_Error = "El parte no tiene ID"
        Err.Raise 1000
    End If
    If fso.FileExists(Me.URLParte) Then
        If FicheroAbierto(Me.URLParte) Then
            p_Error = "El parte tiene ya anexado un anexo y está abierto"
            Err.Raise 1000
        End If
    End If
    
    
    p_URLAnexoFinal = m_ObjEntorno.URLDirectorioPartesFirmados & Me.IDParte & "_PA." & fso.GetExtensionName(p_URLLocal)
    If fso.FileExists(p_URLAnexoFinal) Then
        If FicheroAbierto(p_URLAnexoFinal) Then
            p_Error = "El parte tiene ya anexado un anexo y está abierto"
            Err.Raise 1000
        End If
    End If
    fso.CopyFile p_URLLocal, p_URLAnexoFinal, True
    Me.NombreArchivoAdjunto = fso.GetFileName(p_URLAnexoFinal)
    Me.URLParte = ""
    '----------
    ' registro dentro de la parte principal del parte
    '-------------
    m_SQL = "UPDATE TbPartesPpal SET TbPartesPpal.NombreArchivoAdjunto = '" & Me.NombreArchivoAdjunto & "' " & _
                "WHERE (((TbPartesPpal.IDParte)='" & Me.IDParte & "'));"
    DoCmd.SetWarnings False
    getdb().Execute m_SQL
    DoCmd.SetWarnings True
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Parte.AnexarParte ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Function GeneraWordParaFirma( _
                                    p_Factura As Factura, _
                                    p_Centro As String, _
                                    Optional ByRef p_ObjColEventos As Scripting.Dictionary, _
                                    Optional ByRef p_URLCarpeta As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    Dim appWord As Word.Application
    Dim Plantilla As Word.Document
    Dim m_FechaMinima As String
    Dim FechaMaxima As String
    Dim m_Resultado As String
    
    On Error GoTo errores
    If p_URLCarpeta = "" Then
        p_URLCarpeta = m_ObjEntorno.URLDirectorioLocal
    End If
    m_Linea = "Obtención de los eventos de la factura"
    Avance m_Linea
    If p_ObjColEventos Is Nothing Then
        Set p_ObjColEventos = Constructor.getEventosEnFacturaYCentro(p_Factura, p_Centro, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If p_ObjColEventos Is Nothing Then
        Exit Function
    End If
    m_Linea = "Obtención de los eventos de la factura ......"
    Avance m_Linea
    m_Resultado = getFechaInicialYFinalDeColeEventos(p_ObjColEventos, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    dato = Split(m_Resultado, "|")
    m_FechaMinima = dato(0)
    FechaMaxima = dato(1)
    m_Linea = "Abriendo Word ......"
    Avance m_Linea
    Set appWord = New Word.Application
    'abrimos la plantilla
    Set Plantilla = appWord.Documents.Open(m_ObjEntorno.URLPlantillaInformeAverias)
    GeneraWordParaFirma = RellenarInformeEnWord(appWord, Plantilla, p_ObjColEventos, p_Centro, p_URLCarpeta, m_FechaMinima, FechaMaxima, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    appWord.Documents(m_ObjEntorno.URLPlantillaInformeAverias).Close False
    Set Plantilla = Nothing
    appWord.Quit
    Set appWord = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GeneraWordParaFirma ha devuelto el error: " & vbNewLine & Err.Description
    End If
    If Not appWord.Documents(m_ObjEntorno.URLPlantillaInformeAverias) Is Nothing Then
        appWord.Documents(m_ObjEntorno.URLPlantillaInformeAverias).Close False
        Set Plantilla = Nothing
    End If
    If Not appWord Is Nothing Then
        appWord.Quit
        Set appWord = Nothing
    End If
End Function

Public Function Alta( _
                        p_ObjColEventos As Scripting.Dictionary, _
                        p_IDFactura As String, _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_ID As Variant
    Dim m_ObjEvento As Evento
    Dim m_IDParte As String
    Dim m_SQL As String
    Dim m_FechaMinima As String
    Dim m_FechaMaxima As String
    Dim flag As String
    Dim blnPrimeraParteRegistrada As Boolean
    Dim blnSegundaParteRegistrada As Boolean
    On Error GoTo errores
    
    
    If p_ObjColEventos Is Nothing Then
        p_Error = "No se ha introducido una cadena de IDEventos"
        Err.Raise 1000
    End If
    If p_IDFactura = "" Then
        p_Error = "Falta IDFactura"
        Err.Raise 1000
    End If
    Me.IDFactura = p_IDFactura
    
    m_SQL = "DELETE TbPartesPpal.* " & _
            "FROM TbPartesPpal " & _
            "WHERE IDFactura=" & p_IDFactura & " AND Centro='" & Me.Centro & "';"
    getdb().Execute m_SQL
    m_SQL = "DELETE TbPartesDetalle.* " & _
            "FROM TbPartesPpal INNER JOIN TbPartesDetalle ON TbPartesPpal.IDParte = TbPartesDetalle.IDParte " & _
            "WHERE IDFactura=" & p_IDFactura & " AND Centro='" & Me.Centro & "';"
    getdb().Execute m_SQL
    ' getFechaInicialYFinalDeColeEventos = m_FechaInicialMinima & "|" & m_FechaFinalMaxima
    flag = getFechaInicialYFinalDeColeEventos(p_ObjColEventos, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    dato = Split(flag, "|")
    m_FechaMinima = dato(0)
    m_FechaMaxima = dato(1)
    Me.FechaInicial = m_FechaMinima
    Me.FechaFinal = m_FechaMaxima
    
   
    
    m_SQL = "TbPartesPpal"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        Me.IDParte = Me.IDParteCalculado
        .AddNew
            .Fields("IDParte") = Me.IDParte
            .Fields("FechaObtencion") = Date
            .Fields("FechaInicial") = Me.FechaInicial
            .Fields("FechaFinal") = Me.FechaFinal
            If Me.Descripcion <> "" Then
                .Fields("Descripcion") = Me.Descripcion
            End If
            .Fields("Centro") = Me.Centro
            .Fields("IDFactura") = Me.IDFactura
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
     m_SQL = "DELETE TbPartesDetalle.* " & _
            "FROM TbPartesDetalle " & _
            "WHERE (((TbPartesDetalle.IDParte)='" & Me.IDParte & "'));"
    getdb().Execute m_SQL
    
    '------------------------
    ' RELLENO TABLA TbPartesDetalle
    '-------------------
    For Each m_ID In p_ObjColEventos
        Set m_ObjEvento = p_ObjColEventos(m_ID)
        
        
        
        m_SQL = "TbPartesDetalle"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        With rcdDatos
            .AddNew
                .Fields("IDParte") = Me.IDParte
                .Fields("IDEvento") = m_ObjEvento.IDEVENTO
                .Fields("FechaAlta") = m_ObjEvento.FECHAALTAEVENTO
                .Fields("HoraAlta") = m_ObjEvento.HORAINICIALEVENTO
                .Fields("Descripcion") = m_ObjEvento.Descripcion
                .Fields("CausaFin") = m_ObjEvento.CausaFin
                .Fields("FechaFin") = m_ObjEvento.FechaFinal
                .Fields("HoraFin") = m_ObjEvento.HORAFINALEVENTO
                .Fields("AliasCreador") = m_ObjEvento.ALIASTECNICO
                .Fields("Originador") = m_ObjEvento.Originador
                .Fields("TiempoRespuesta") = m_ObjEvento.TIEMPORESPUESTAEVENTO
                .Fields("SubSistema") = m_ObjEvento.SubSistema
                If Not m_ObjEvento.Equipo Is Nothing Then
                    .Fields("Equipo") = m_ObjEvento.Equipo.Equipo
                End If
                If m_ObjEvento.PMPR <> "" Then
                    .Fields("PMPR") = m_ObjEvento.PMPR
                End If
            .Update
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        Set m_ObjEvento = Nothing
        
    Next
    blnSegundaParteRegistrada = True
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Parte.Alta ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    If blnSegundaParteRegistrada = False And blnPrimeraParteRegistrada = True Then
        On Error Resume Next
        m_SQL = "DELETE TbPartesPpal.IDParte " & _
                    "FROM TbPartesPpal " & _
                    "WHERE (((TbPartesPpal.IDParte)='" & Me.IDParte & "'));"
        DoCmd.SetWarnings False
        getdb().Execute m_SQL
        DoCmd.SetWarnings True
    End If
    
End Function

Public Function Eliminar( _
                        Optional ByRef p_Error As String _
                        ) As String
    
   
    
    On Error GoTo errores
    If Me.IDParte = "" Then
        p_Error = "No se conoce el IDParte"
        Err.Raise 1000
    End If
    m_SQL = "DELETE * FROM TbPartesPpal " & _
            "WHERE IDParte='" & Me.IDParte & "';"
    DoCmd.SetWarnings False
    getdb().Execute m_SQL
    DoCmd.SetWarnings True
    If fso.FileExists(Me.URLParte) Then
        fso.DeleteFile Me.URLParte, True
    End If
    m_SQL = "DELETE * FROM TbPartesDetalle " & _
            "WHERE IDParte='" & Me.IDParte & "';"
    DoCmd.SetWarnings False
    getdb().Execute m_SQL
    DoCmd.SetWarnings True
    
    m_SQL = "UPDATE TbEventos SET TbEventos.IDParte = Null " & _
            "WHERE (((TbEventos.IDParte)='" & Me.IDParte & "'));"
    DoCmd.SetWarnings False
    getdb().Execute m_SQL
    DoCmd.SetWarnings True
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Parte.Eliminar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    
End Function

Public Property Get IDParteCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_sIDParteCalculado = getIDParte(m_Error)
    If m_Error <> "" Then
        Err.r1000
    End If
    IDParteCalculado = m_sIDParteCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Parte.IDParteCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Private Function getIDParte( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Alias As String
    Dim m_ParteFija As String
    Dim Maximo As Long
    On Error GoTo errores
    
    m_ParteFija = "FSP" & "_" & Left(Year(Date), 2) & Format(Month(Date), "00")
    m_SQL = "SELECT Max(Right([IDParte],2)) AS Ordinal " & _
            "FROM TbPartesPpal " & _
            "WHERE (((TbPartesPpal.IDParte) Like '" & m_ParteFija & "*'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(.Fields(0)) Then
                Maximo = .Fields(0)
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    getIDParte = m_ParteFija & Format(CStr(Maximo + 1), "00")
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getIDParte ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function



Public Property Get ColCampos() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set m_objColCampos = New Collection
    With m_objColCampos
        .Add "IDParte"
        .Add "FechaObtencion"
        .Add "FechaInicial"
        .Add "FechaFinal"
        .Add "Descripcion"
        .Add "NombreArchivoAdjunto"
        .Add "IDFactura"
        .Add "Centro"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PartesPpal.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDParte" Then
        getPropiedad = Me.IDParte
        Exit Function
    End If
    If m_NombrePropiedad = "FechaObtencion" Then
        getPropiedad = Me.FechaObtencion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicial" Then
        getPropiedad = Me.FechaInicial
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinal" Then
        getPropiedad = Me.FechaFinal
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "NombreArchivoAdjunto" Then
        getPropiedad = Me.NombreArchivoAdjunto
        Exit Function
    End If
    If m_NombrePropiedad = "IDFactura" Then
        getPropiedad = Me.IDFactura
        Exit Function
    End If
    If m_NombrePropiedad = "Centro" Then
        getPropiedad = Me.Centro
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en PartesPpal.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PartesPpal.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDParte" Then
        Me.IDParte = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaObtencion" Then
        Me.FechaObtencion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicial" Then
        Me.FechaInicial = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinal" Then
        Me.FechaFinal = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreArchivoAdjunto" Then
        Me.NombreArchivoAdjunto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDFactura" Then
        Me.IDFactura = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Centro" Then
        Me.Centro = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en PartesPpal.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PartesPpal.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function




