Attribute VB_Name = "Variables Globales"
Option Compare Database
Option Explicit

#If Win64 = 1 Then
    
    Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
            ByVal lpApplicationName As String, _
            ByVal lpKeyName As String, _
            ByVal lpDefault As String, _
            ByVal lpReturnedString As String, _
            ByVal nSize As Long, _
            ByVal lpFileName As String) As Long
    
    Public Declare PtrSafe Function Ejecutar Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hWnd As Long, ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
    Public Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    Public Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" ( _
        ByVal hProcess As Long, lpExitCode As Long) As Long
    Public Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
    
#Else
    
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
            ByVal lpApplicationName As String, _
            ByVal lpKeyName As String, _
            ByVal lpDefault As String, _
            ByVal lpReturnedString As String, _
            ByVal nSize As Long, _
            ByVal lpFileName As String) As Long
    Public Declare Function Ejecutar Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hWnd As Long, ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
    Public Declare Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    Public Declare Function GetExitCodeProcess Lib "kernel32" ( _
        ByVal hProcess As Long, lpExitCode As Long) As Long
    Public Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
#End If
Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STATUS_PENDING = &H103&
Public appWord As Word.Application
Public Const wdOrientLandscape As Long = 1
Public Const wdStory As Long = 6
Public Const wdFormatOriginalFormatting As Long = 16
Public Const wdPaneNone As Long = 0
Public Const wdNormalView As Long = 1
Public Const wdOutlineView As Long = 2
Public Const wdPrintView As Long = 3
Public Const wdSeekCurrentPageFooter As Long = 10
Public Const wdFieldEmpty As Long = -1
Public Const wdExportFormatPDF As Long = 17
Public Const wdExportOptimizeForPrint As Long = 0
Public Const wdExportAllDocument As Long = 0
Public Const wdExportDocumentContent As Long = 0
Public Const wdExportCreateNoBookmarks As Long = 0
    
Public Const xlDiagonalDown As Long = 5
Public Const xlDiagonalUp As Long = 6
Public Const xlEdgeLeft As Long = 7
Public Const xlEdgeTop As Long = 8
Public Const xlEdgeBottom As Long = 9
Public Const xlEdgeRight As Long = 10

Public Const xlInsideVertical As Long = 11
Public Const xlInsideHorizontal As Long = 12
Public Const xlContinuous As Long = 1
Public Const xlThin As Long = 2
Public Const xlMedium As Long = -4138
Public Const xlAutomatic As Long = -4105
Public Const xlNone As Long = -4142
Public Const xlDouble As Long = -4119
Public Const xlThick As Long = 4
Public Const xlToLeft As Long = -4159

Public Const xlLineMarkersStacked As Long = 66
Public Const xlValue As Long = 2
Public Const xlLinear As Long = -4132
Public Const xlCategory As Long = 1
Public Const xlSquare As Long = 1
Public Const xlHairline As Long = 1
Public Const xlUnderlineStyleNone As Long = -4142
Public Const xlCenter As Long = -4108
Public Const xlContext As Long = -5002
Public Const xlLabelPositionAbove As Long = 0
Public Const xlHorizontal As Long = -4128
Public Const xlBottom As Long = -4107
Public Const xlColumnClustered As Long = 51
Public Const xlSolid As Long = 1
Public Const xlPrintNoComments As Long = -4142
Public Const xlLandscape As Long = 2
Public Const xlPortrait As Long = 1
Public Const xlPaperA4 As Long = 9
Public Const xlDownThenOver As Long = 1
Public Const xlPrintErrorsDisplayed As Long = 0
Public Const xlGeneral As Long = 1
Public Const xlTop As Long = -4160
Public Const xlLegendPositionBottom As Long = -4107


Public Const xlLeft As Long = -4131
Public Const xlRight As Long = -4152
Public Const xlUp As Long = -4162

Public Const xlTypePDF As Long = 0
Public Const xlThemeColorDark1 As Long = 1
Public Const xlLine As Long = 4
Public Const xlLightDown As Long = 13
Public Const xlSortOnValues As Long = 0
Public Const xlAscending As Long = 1
Public Const xlSortNormal As Long = 0
Public Const xlGuess As Long = 0
Public Const xlLeftToRight As Long = 2
Public Const xlPinYin As Long = 1
Public Const xlMillions As Long = -6
Public Const xlInside As Long = 2
     
Public Const msoFileDialogFilePicker As Long = 3
Public Const msoFileDialogFolderPicker As Long = 4
Public Const msoFileDialogOpen As Long = 1
Public Const msoFileDialogSaveAs As Long = 2



'*************************************
' PARA LOS MANTENIMIENTOS PREVENTIVOS
'****************************************
Public colPreventivoEquipos As New Collection
Public colPreventivosEquiposProgramados As New Collection
Public colPreventivosProgramados As New Collection
Public colPreventivosEquiposNoProgramados As New Collection
Public colPreventivosRealizados As New Collection
Public colPreventivosRealizadosEnPlazo As New Collection
Public colPreventivosRealizadosFueraDePlazo As New Collection
Public colPreventivosNoRealizados As New Collection
Public colPreventivosNoRealizadosVencidos As New Collection
Public colPreventivosNoRealizadosEnPlazo As New Collection
Public colPreventivosNoRealizadosProximos As New Collection
        
Public Enum EnumSino
    Sí = 1
    No = 2
End Enum
Public Enum EnumCampos
    TbEventosIDEvento = 1
    TbEventosDESCRIPCION = 2
    TbEventosBUI = 3
    TbEventosSUBSISTEMA = 4
    TbEventosIDEquipo = 5
    TbEventosFECHAALTAEVENTO = 6
    TbEventosFranqueado = 7
    TbEventosPMPR = 8
    TbEventosORIGINADOR = 9
    TbEventosTIPOEVENTO = 10
    TbActividadesNOMBRE = 11
    TbActividadesActividad = 12
    TbActividadesFechaAlta = 13
    TbActividadesHorasLaborables = 14
    TbActividadesHorasExtras = 15
    TbActividadesDESCRIPCION = 16
    TbActividadesUbicacion = 17
    TbMaterialFechaIncio = 18
    TbMaterialGarantia = 19
    TbMaterialTipoAccion = 20
    TbMaterialMaterial = 21
    TbMaterialPN = 22
    TbMaterialNS = 23
    TbMaterialCOSTE = 24
    TbMaterialReparadoPor = 25
    TbMaterialDescripcion = 26
    TbMaterialEsReparacion = 27
    TbEventosFECHAFIN = 28
    TbActividadesIDActividad = 29
    TbMaterialIDMaterial = 30
End Enum
Public Enum TipoInformeHoras
    Todo = 1
    Excepciones = 2
End Enum
Public Enum EnumTipoAnexo
    Evento = 1
    Actividad = 2
    Material = 3
    MaterialSeg = 4
    Subcontratacion = 5
    Gasto = 6
End Enum
Public m_SQL As String
Public pregunta As Long
Public flag As String
Public dato As Variant
Public appexcel As excel.Application
'----------------------------------------
' Iniciar cosas generales
'----------------------------------------
Public m_ObjEntorno As Entorno
Private fso As New FileSystemObject
        
        
Public lngColorCampoEditable As Long, lngColorCampoNoEditable As Long, lngColorCabecera As Long
Public lngColorEtiquetasTexto As Long, lngColorCampoTexto As Long, lngColorEtiquetaMenuSinPulsar As Long, _
    longColorEtiquetaMenuPulsado As Long
Public m_ColEventosParaInformeDeRAC As Scripting.Dictionary
Public m_ColParaEnumNombreCamposEnInforme As Scripting.Dictionary
Public m_ColParaNombreCamposInformeTabla As Scripting.Dictionary

Public m_ColEventos As Scripting.Dictionary
Public m_ColActividades As Scripting.Dictionary
Public m_ColMateriales As Scripting.Dictionary
Public m_ColMaterialesSeguimiento As Scripting.Dictionary
Public m_ColTecnicos As Scripting.Dictionary
Public m_ColFacturas As Scripting.Dictionary
Public m_ColGastos As Scripting.Dictionary
Public m_ColSubContrataciones As Scripting.Dictionary
Public m_ColEquipos As Scripting.Dictionary
Public m_ColTipoTecnicos As Scripting.Dictionary
Public m_ColTipoTecnicoPrecios As Scripting.Dictionary


Public m_Linea As String
Public wks As DAO.Workspace
Private db As DAO.Database
Private db1 As DAO.Database
Public IDAplicacion As String
Public m_ObjUsuarioConectado As Usuario
Public m_ObjUsuarioConectadoInicialmente As Usuario

Public m_NombreCarpetaVideo As String
Public m_ObjEventoActivo As Evento
Public m_ObjTecnicoActivo As Tecnico
Public m_ObjLibranzaActiva As LIbranza
Public m_ObjTipoTecnicoActivo As TIPOTECNICO
Public m_ObjTipoTecnicoPrecioActivo As TipoTecnicoPrecio
Public m_ObjFacturaActiva As Factura
Public m_ObjAnexoActivo As Anexo
Public m_ObjGastoActivo As Gasto
Public m_ObjEntidadParaAnexoActiva As Object
Public m_ObjEquipoMedidaActivo As EquipoMedida
Public m_ObjCalibracionEquipoMedidaActiva As EquipoMedidaCalibracion
Public m_ColEquiposMedidaPorEventoActivos As Scripting.Dictionary
Public m_EnOficina As EnumSino
Public m_Win64 As EnumSino
Public m_TextoWin64 As String
Public m_URLRutaAplicacionesLocal As String
Public m_URLRutaAplicacionesRemotas As String
Public m_URLRutaAplicacionLocal As String
Public m_URLRutaAplicacionRemota As String

Public Function EVE(Optional ByRef p_Error As String) As String
    
    Dim m_Propiedad As Variant
    Dim m_ValorPropiedad As String
    Dim m_ObjValorPropiedad As Object
    Dim t1 As Variant
    Dim t2 As Variant
    Dim m_ObjColPropiedades As Scripting.Dictionary
    Dim m_UsuarioLogeadoEnOrdenador As String
    Dim m_ParteObjeto As String
    Dim m_ColMal As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Command As String
    Dim objNetwork As Object
    
    On Error GoTo errores
    
    t1 = Timer
    IDAplicacion = "6"
    #If Win64 = 1 Then
        m_Win64 = EnumSino.Sí
        m_TextoWin64 = "(64 bits)"
    #Else
         m_Win64 = EnumSino.No
         m_TextoWin64 = "(32 bits)"
    #End If
    Set m_ObjUsuarioConectado = Nothing
    Set m_ColEquiposMedidaPorEventoActivos = Nothing
    'Application.TempVars("DatosEnLocal") = "Sí"
    Application.TempVars("DatosEnLocal") = "No"
    Application.TempVars("EnDesarrollo") = "No"
    'Application.TempVars("EnDesarrollo") = "Sí"
    
    
    Set m_ObjEntorno = New Entorno
     m_URLRutaAplicacionesRemotas = "\\datoste\aplicaciones_dys\Aplicaciones PpD\"
    m_URLRutaAplicacionRemota = m_URLRutaAplicacionesRemotas & "BRASS\"
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URLRutaAplicacionesLocal = getRutaAplicacionesLocal(p_Error)
        If m_URLRutaAplicacionesLocal <> "" Then
            m_URLRutaAplicacionLocal = m_URLRutaAplicacionesLocal & "BRASS\"
        End If
    Else
       m_URLRutaAplicacionesLocal = ""
       m_URLRutaAplicacionLocal = ""
    End If
    m_Command = Nz(VBA.Command, "")
    'm_Command = "david.escuderolazaro@telefonica.com"
    If m_Command <> "" Then
        Set m_ObjUsuarioConectado = Constructor.getUsuario(, , , m_Command, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Set objNetwork = CreateObject("Wscript.Network")
        m_UsuarioLogeadoEnOrdenador = objNetwork.UserName
        If m_UsuarioLogeadoEnOrdenador = "Local1" Then m_UsuarioLogeadoEnOrdenador = "adm"
        If m_UsuarioLogeadoEnOrdenador = "adm1" Then m_UsuarioLogeadoEnOrdenador = "adm"
        Set m_ObjUsuarioConectado = Constructor.getUsuario(, m_UsuarioLogeadoEnOrdenador, , , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set objNetwork = Nothing
    End If
    If m_ObjUsuarioConectado Is Nothing Then
        p_Error = "No se ha podido determinar el usuario que está usando la herramienta"
        Err.Raise 1000
    End If
If m_ObjUsuarioConectadoInicialmente Is Nothing Then
        Set m_ObjUsuarioConectadoInicialmente = m_ObjUsuarioConectado
    End If

    m_NombreCarpetaVideo = "BRASS"
    lngColorCampoEditable = 16511723
    lngColorCampoNoEditable = 16047562
    lngColorCampoTexto = 8210719
    lngColorEtiquetasTexto = 8210719
    lngColorCabecera = 8210719
    
    lngColorEtiquetaMenuSinPulsar = 8210719
    longColorEtiquetaMenuPulsado = 26367
   
    Set m_ColEventos = Nothing
    Set m_ColActividades = Nothing
    Set m_ColMateriales = Nothing
    Set m_ColMaterialesSeguimiento = Nothing
    Set m_ColTecnicos = Nothing
    Set m_ColFacturas = Nothing
    Set m_ColGastos = Nothing
    Set m_ColSubContrataciones = Nothing
    Set m_ColEquipos = Nothing
    Set m_ColTipoTecnicos = Nothing
    Set m_ColTipoTecnicoPrecios = Nothing
    Set m_ObjEventoActivo = Nothing
    Set m_ObjTecnicoActivo = Nothing
    Set m_ObjLibranzaActiva = Nothing
    Set m_ObjTipoTecnicoActivo = Nothing
    Set m_ObjTipoTecnicoPrecioActivo = Nothing
    Set m_ObjFacturaActiva = Nothing
    
    Set m_ObjColPropiedades = m_ObjEntorno.ColPropiedadesEntorno
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    For Each m_Propiedad In m_ObjColPropiedades.Keys
       'If m_Propiedad = "URLArchivoCSS" Then Stop
        m_ParteObjeto = m_ObjColPropiedades(m_Propiedad)
        If m_ParteObjeto = "o" Then
          
            Set m_ObjValorPropiedad = m_ObjEntorno.getValorPropiedad(CStr(m_Propiedad))
            p_Error = m_ObjEntorno.Error
            If p_Error <> "" Then
                Set m_ObjValorPropiedad = Nothing
                If m_ColMal Is Nothing Then
                    Set m_ColMal = New Scripting.Dictionary
                    m_ColMal.CompareMode = TextCompare
                End If
                
                m_ColMal.Add m_Propiedad, p_Error
                p_Error = ""
                m_ObjEntorno.Error = ""
                GoTo SiguientePropiedad
            End If
          'Debug.Print m_Propiedad, "Objeto"
           
           Set m_ObjValorPropiedad = Nothing
        Else
            
            m_ValorPropiedad = m_ObjEntorno.getValorPropiedad(CStr(m_Propiedad), p_Error)
            
            If p_Error <> "" Then
                m_ValorPropiedad = ""
                If m_ColMal Is Nothing Then
                    Set m_ColMal = New Scripting.Dictionary
                    m_ColMal.CompareMode = TextCompare
                End If
                
                m_ColMal.Add m_Propiedad, p_Error
                p_Error = ""
                m_ObjEntorno.Error = ""
                GoTo SiguientePropiedad
            End If
            If m_ParteObjeto = "s" Then
                If m_ValorPropiedad = "1" Then
                    m_ValorPropiedad = "Sí"
                ElseIf m_ValorPropiedad = "2" Then
                    m_ValorPropiedad = "No"
                Else
                    m_ValorPropiedad = "#ERR"
                End If
            End If

            m_ValorPropiedad = ""
        End If
SiguientePropiedad:
        m_ObjEntorno.Error = ""
        
    Next
    If Not m_ColMal Is Nothing Then
        For Each m_ID In m_ColMal
            If p_Error = "" Then
                p_Error = m_ID & vbTab & m_ColMal(m_ID)
            Else
                p_Error = p_Error & vbNewLine & m_ID & vbTab & m_ColMal(m_ID)
            End If
            
        Next
        Err.Raise 1000
    End If
    
    
    t2 = Timer
    Debug.Print "Variables establecidas correctamente en : " & t2 - t1
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EVE ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    Debug.Print p_Error
End Function



Public Function LeerIni(Key As String, Default As Variant) As String
    Dim bufer As String * 256, Len_Value As Long
    Len_Value = GetPrivateProfileString(fso.GetBaseName(CurrentDb().Name), _
                                     Key, _
                                     Default, _
                                     bufer, _
                                     Len(bufer), _
                                     m_ObjEntorno.URLAchivoIni)
      
    LeerIni = Left$(bufer, CLng(Len_Value))
    
End Function

Private Function getRutaBackendDesdeTablasVinculadas( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    On Error GoTo errores

    Dim dbActual As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strConnect As String
    Dim lngPosIni As Long
    Dim lngPosFin As Long

    Set dbActual = CurrentDb()

    For Each tdf In dbActual.TableDefs
        If Left$(tdf.Name, 4) <> "MSys" Then
            If Nz(tdf.SourceTableName, "") <> "" Then
                strConnect = Nz(tdf.Connect, "")
                lngPosIni = InStr(1, strConnect, "DATABASE=", vbTextCompare)
                If lngPosIni > 0 Then
                    lngPosIni = lngPosIni + Len("DATABASE=")
                    lngPosFin = InStr(lngPosIni, strConnect, ";")

                    If lngPosFin = 0 Then
                        getRutaBackendDesdeTablasVinculadas = Mid$(strConnect, lngPosIni)
                    Else
                        getRutaBackendDesdeTablasVinculadas = Mid$(strConnect, lngPosIni, lngPosFin - lngPosIni)
                    End If

                    If getRutaBackendDesdeTablasVinculadas <> "" Then
                        GoTo salir
                    End If
                End If
            End If
        End If
    Next tdf

salir:
    Set tdf = Nothing
    Set dbActual = Nothing
    Exit Function

errores:
    p_Error = "El método getRutaBackendDesdeTablasVinculadas ha devuelto el error: " & vbNewLine & Err.Description
    Resume salir
End Function

Public Function getdbLanzadera( _
                                Optional ByRef p_Error As String _
                                ) As DAO.Database
    
    Dim m_URL As String
    On Error GoTo errores
    
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\Lanzadera_Datos.accdb"
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_URL = m_URLRutaAplicacionesRemotas & "0Lanzadera\Lanzadera_Datos.accdb"
    Else
        p_Error = "No se conoce el origen de los datos"
        Err.Raise 1000
    End If
   
    
    
    Set wks = DBEngine.Workspaces(0)
    Set db1 = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdbLanzadera = db1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdbLanzadera ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getdb( _
                        Optional ByRef p_Error As String _
                        ) As DAO.Database
    
    Dim m_URL As String
    Dim m_RutaBackendVinculado As String
    On Error GoTo errores

    m_RutaBackendVinculado = getRutaBackendDesdeTablasVinculadas(p_Error)
    If p_Error <> "" Then
        p_Error = ""
    End If

    If m_RutaBackendVinculado <> "" Then
        m_URL = m_RutaBackendVinculado
    ElseIf Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = m_URLRutaAplicacionLocal & "Gestion_Brass_Gestion_Datos.accdb"
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_URL = m_URLRutaAplicacionRemota & "Gestion_Brass_Gestion_Datos.accdb"
    Else
        p_Error = "No se conoce el origen de los datos"
        Err.Raise 1000
    End If
    
    Set wks = DBEngine.Workspaces(0)
    Set db = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdb = db
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdb ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


