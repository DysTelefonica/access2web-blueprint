VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Evento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private fso As New FileSystemObject

Public IDEVENTO As String
Public NODO As String
Public BUI As String
Public SubSistema As String
Public IDEquipo As String
Public PMPR As String
Public TipoEvento As String
Public Criticidad As String
Public ALIASTECNICO As String
Public Originador As String
Public FECHAALTAEVENTO As String
Public FechaFinal As String
Public HORAINICIALEVENTO As String
Public HORAFINALEVENTO As String
Public TIEMPORESPUESTAEVENTO As String
Public CausaFin As String
Public Descripcion As String
Public Contacto As String
Public FechaRegistroAlta As String
Public FechaRegistroModificacion As String
Public Franqueado As Boolean
Public IDExportacion As String
Public CodImportacion As String
Public IDParte As String
Public Notas As String
Public FechaEnInformeRAC As String

'===========================================================
' Campos SLA - Spec-001
' En la entidad usamos tipos de tabla (Boolean), no EnumSino.
' EnumSino es para la UI (ViewModel), para representar el tercer estado "sin responder".
'===========================================================
Public FechaRecepcionNotificacion As String
Public FechaInicioContactoCliente As String
Public IncidenciaAveriaOReparacion As Boolean
Public FechaInicioTiempoAdquisicion As String
Public FechaFinTiempoAdquisicion As String
Public TipoReparacion As String
Public Urgente As Boolean
Public EventoConServicioAfectado As Boolean
Public FechaRestablecimientoServicio As String
Public TipoRepInsitu As Boolean
Public TipoRepNoSMT As Boolean
Public TipoRepValvulas As Boolean

Public ParaAlta As EnumSino
Private m_objColAnexos As Scripting.Dictionary
Private m_sIDEventoCalculado As String
Private m_eTieneActividades As EnumSino
Private m_eTieneMateriales As EnumSino
Private m_eTieneActividadesConMateriales As EnumSino
Private m_eTodosLosMaterialesRecepcionados As EnumSino
Private m_ObjColActividades As Scripting.Dictionary
Private m_ObjColMateriales As Scripting.Dictionary
Private m_ObjEquipo As Equipo
Private m_ObjFactura As Factura
Private m_sFechaFacturado As String
Private m_ObjTecnico As Tecnico
Private m_DiasParaFranqueo As Double
Private m_HorasEnFranqueo As Double
Private m_sFechaFinalCalculada As String
Private m_sCausaFinCalculada As String
Private m_sNombreEquipo As String
Private m_sFranqueadoParaInforme As String
Private m_sHorasLabTotales As String
Private m_sHorasExtTotales As String
Private m_objParte As Parte
Private m_eAnexosCerrados As EnumSino
Private m_sURLAnexos As String
Private m_sURLAnexoParte As String
Private m_objEquipoMedida As EquipoMedida
Private m_objColEquiposMedida As Scripting.Dictionary
Private m_eModificable As EnumSino
Private m_eAlguienDeGetronics As EnumSino
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String



Public Property Get URLAnexoParte() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Parte Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_sURLAnexoParte = Me.Parte.URLParte
    m_Error = Me.Parte.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    URLAnexoParte = m_sURLAnexoParte
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.URLAnexoParte ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get URLAnexos() As String
    
    Dim m_Anexo As Anexo
    Dim m_ID As Variant
    Dim m_Col As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    Set m_Col = Me.ColAnexos
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_Anexo = m_Col(m_ID)
        If m_sURLAnexos = "" Then
            m_sURLAnexos = m_Anexo.URLAnexo
        Else
            m_sURLAnexos = m_sURLAnexos & ";" & m_Anexo.URLAnexo
        End If
        Set m_Anexo = Nothing
    Next
    URLAnexos = m_sURLAnexos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.URLAnexos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get IDEventoCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sIDEventoCalculado = DameIDEvento(Me.BUI, Me.FECHAALTAEVENTO)
    
    IDEventoCalculado = m_sIDEventoCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.IDEventoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColActividades() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If Not m_ObjColActividades Is Nothing Then
        Set ColActividades = m_ObjColActividades
        Exit Property
    End If
    If Me.IDEVENTO = "" Then
        Exit Property
    End If
    Set m_ObjColActividades = Constructor.getActividadesDeEvento(Me.IDEVENTO, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColActividades = m_ObjColActividades
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.colActividades ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FechaFinalCalculada() As String
    
    Dim m_IDActividad As Variant
    Dim m_Actividad As Actividad
    
    Dim m_FechaMax As String
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If m_sFechaFinalCalculada <> "" Then
        FechaFinalCalculada = m_sFechaFinalCalculada
        Exit Property
    End If
    
    If Me.ColActividades Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    For Each m_IDActividad In Me.ColActividades
        Set m_Actividad = Me.ColActividades(m_IDActividad)
        If IsDate(m_Actividad.FechaAlta) Then
            If Not IsDate(m_FechaMax) Then
                m_FechaMax = m_Actividad.FechaAlta
            Else
                If CDate(m_Actividad.FechaAlta) > CDate(m_FechaMax) Then
                    m_FechaMax = m_Actividad.FechaAlta
                End If
            End If
        End If
        Set m_Actividad = Nothing
    Next
    m_sFechaFinalCalculada = m_FechaMax
    FechaFinalCalculada = m_sFechaFinalCalculada
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.FechaFinalCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get CausaFinCalculada() As String
    
    Dim m_IDActividad As Variant
    Dim m_Actividad As Actividad
    Dim m_Causas As Scripting.Dictionary
    Dim m_Causa As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If m_sCausaFinCalculada <> "" Then
        CausaFinCalculada = m_sCausaFinCalculada
        Exit Property
    End If
    If Me.ColActividades Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    For Each m_IDActividad In Me.ColActividades
        Set m_Actividad = Me.ColActividades(m_IDActividad)
        If m_Actividad.Descripcion <> "" Then
            If m_Causas Is Nothing Then
                Set m_Causas = New Scripting.Dictionary
                m_Causas.CompareMode = TextCompare
            End If
            If Not m_Causas.Exists(m_Actividad.Descripcion) Then
                m_Causas.Add m_Actividad.Descripcion, m_Actividad.Descripcion
            End If
            
            
        End If
        Set m_Actividad = Nothing
    Next
    If Not m_Causas Is Nothing Then
        For Each m_Causa In m_Causas
            If m_sCausaFinCalculada = "" Then
                m_sCausaFinCalculada = CStr(m_Causa)
            Else
                m_sCausaFinCalculada = m_sCausaFinCalculada & vbNewLine & CStr(m_Causa)
            End If
        Next
    End If
    CausaFinCalculada = m_sCausaFinCalculada
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.CausaFinCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColMateriales() As Scripting.Dictionary
    
    Dim m_Actividad As Actividad
    Dim m_Material As Material
    Dim m_IDActividad As Variant
    Dim m_IDMaterial As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColMateriales Is Nothing Then
        Set ColMateriales = m_ObjColMateriales
        Exit Property
    End If
    
    If Me.ColActividades Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    For Each m_IDActividad In Me.ColActividades
        Set m_Actividad = Me.ColActividades(m_IDActividad)
        If m_Actividad.ColMateriales Is Nothing Then
            m_Error = m_Actividad.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            GoTo sigActividad
        Else
            m_Error = m_Actividad.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        For Each m_IDMaterial In m_Actividad.ColMateriales
            Set m_Material = m_Actividad.ColMateriales(m_IDMaterial)
            If m_ObjColMateriales Is Nothing Then
                Set m_ObjColMateriales = New Scripting.Dictionary
                m_ObjColMateriales.CompareMode = TextCompare
            End If
            If Not m_ObjColMateriales.Exists(m_Material.IDMaterial) Then
                m_ObjColMateriales.Add m_Material.IDMaterial, m_Material
            End If
            Set m_Material = Nothing
        Next
sigActividad:
        Set m_Actividad = Nothing
    Next
   
    Set ColMateriales = m_ObjColMateriales
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.ColMateriales ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get TieneActividades() As EnumSino

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eTieneActividades <> 0 Then
        TieneActividades = m_eTieneActividades
        Exit Property
    End If
    If Me.ColActividades Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneActividades = EnumSino.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneActividades = EnumSino.Sí
    End If
    TieneActividades = m_eTieneActividades
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.TieneActividades ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Modificable() As EnumSino
    
    Dim m_FechaFacturado As String
    Dim m_FechaFinal As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eModificable <> 0 Then
        Modificable = m_eModificable
        Exit Property
    End If
    m_FechaFacturado = Me.FechaFacturado
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If IsDate(m_FechaFacturado) Then
        m_eModificable = EnumSino.No
         Modificable = m_eModificable
         Exit Property
    End If
    m_FechaFinal = Me.FechaFinal
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If IsDate(m_FechaFinal) Then
        m_eModificable = EnumSino.No
        Modificable = m_eModificable
        Exit Property
    End If
    m_eModificable = EnumSino.Sí
    Modificable = m_eModificable
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.Modificable ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get TieneMateriales() As EnumSino

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eTieneMateriales <> 0 Then
        TieneMateriales = m_eTieneMateriales
        Exit Property
    End If
    If Me.ColMateriales Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneMateriales = EnumSino.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneMateriales = EnumSino.Sí
    End If
    TieneMateriales = m_eTieneMateriales
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.TieneMateriales ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get TieneActividadesConMateriales() As EnumSino

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eTieneActividadesConMateriales <> 0 Then
        TieneActividadesConMateriales = m_eTieneActividadesConMateriales
        Exit Property
    End If
    If ColMateriales Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneActividadesConMateriales = EnumSino.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneActividadesConMateriales = EnumSino.Sí
    End If
    TieneActividadesConMateriales = m_eTieneActividadesConMateriales
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.TieneActividadesConMateriales ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get TodosLosMaterialesRecepcionados() As EnumSino
    
    Dim m_IDMaterial As Variant
    Dim m_Material As Material
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eTodosLosMaterialesRecepcionados <> 0 Then
        TodosLosMaterialesRecepcionados = m_eTodosLosMaterialesRecepcionados
        Exit Property
    End If
    
    Set m_Col = Me.ColMateriales
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Property
    End If
    For Each m_IDMaterial In Me.ColMateriales
        Set m_Material = Me.ColMateriales(m_IDMaterial)
        If m_Material.Entregado = EnumSino.No Then
            m_Error = m_Material.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            m_eTodosLosMaterialesRecepcionados = EnumSino.No
            TodosLosMaterialesRecepcionados = m_eTodosLosMaterialesRecepcionados
            Exit Property
        Else
            m_Error = m_Material.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        
        Set m_Material = Nothing
    Next
    m_eTodosLosMaterialesRecepcionados = EnumSino.Sí
    TodosLosMaterialesRecepcionados = m_eTodosLosMaterialesRecepcionados
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.TodosLosMaterialesRecepcionados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Equipo() As Equipo
    
    On Error GoTo errores
    
    Me.Error = ""
    
    
    If Not m_ObjEquipo Is Nothing Then
        Set Equipo = m_ObjEquipo
        Exit Property
    End If
    If Me.IDEquipo = "" Then
        Exit Property
    End If
    
    Set m_ObjEquipo = Constructor.getEquipo(Me.IDEquipo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Equipo = m_ObjEquipo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.Equipo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Factura() As Factura
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjFactura Is Nothing Then
        Set Factura = m_ObjFactura
        Exit Property
    End If
    If Me.IDEVENTO = "" Then
        Exit Property
    End If
   
    Set m_ObjFactura = Constructor.getFacturaDeEvento(Me.IDEVENTO, m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Factura = m_ObjFactura
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.Factura ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FechaFacturado() As String

    
    On Error GoTo errores
    
    Me.Error = ""
    
    If IsDate(m_sFechaFacturado) Then
        FechaFacturado = m_sFechaFacturado
        Exit Property
    End If
    If Me.Factura Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_sFechaFacturado = Me.Factura.FechaFactura
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    FechaFacturado = m_sFechaFacturado
        
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.FechaFacturado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get DiasParaFranqueo() As Double
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If m_DiasParaFranqueo <> 0 Then
        DiasParaFranqueo = m_DiasParaFranqueo
        Exit Property
    End If
    If Not Me.Franqueado Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    If Not IsDate(Me.FECHAALTAEVENTO) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "La fecha de alta del evento no es una fecha válida"
        Err.Raise 1000
    End If
    If Not IsDate(Me.FechaFinal) Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
   
    m_DiasParaFranqueo = DateDiff("d", CDate(Me.FECHAALTAEVENTO), CDate(Me.FechaFinal))
    If m_DiasParaFranqueo = 0 Then m_DiasParaFranqueo = 1
    DiasParaFranqueo = m_DiasParaFranqueo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.DiasParaFranqueo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get HorasEnFranqueo() As Double
   
    Dim m_Horas As Double
    On Error GoTo errores
    
    Me.Error = ""
    
    If m_HorasEnFranqueo <> 0 Then
        HorasEnFranqueo = m_HorasEnFranqueo
        Exit Property
    End If
    If Me.IDEVENTO = "" Then
       Exit Property
    End If
   If Not Me.Franqueado Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    m_Horas = getHorasEnFranquear(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_HorasEnFranqueo = CStr(m_Horas)
    HorasEnFranqueo = m_HorasEnFranqueo
        
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.HorasEnFranqueo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get NombreEquipo() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If m_sNombreEquipo <> "" Then
        NombreEquipo = m_sNombreEquipo
        Exit Property
    End If
    
    If Me.Equipo Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    m_sNombreEquipo = Me.Equipo.Equipo
    If m_sNombreEquipo <> "" Then
        NombreEquipo = m_sNombreEquipo
        Exit Property
    End If
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.NombreEquipo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FranqueadoParaInforme() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sFranqueadoParaInforme <> "" Then
        FranqueadoParaInforme = m_sFranqueadoParaInforme
        Exit Property
    End If
    
    If Me.Franqueado = True Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_sFranqueadoParaInforme = "Sí"
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_sFranqueadoParaInforme = "No"
    End If
    FranqueadoParaInforme = m_sFranqueadoParaInforme
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.FranqueadoParaInforme ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get HorasLabTotales() As String
    
    Dim m_Actividad As Actividad
    Dim m_IDActividad As Variant
    Dim m_HorasLab As Double
    
    On Error GoTo errores
    
    Me.Error = ""
    If IsNumeric(m_sHorasLabTotales) Then
        HorasLabTotales = m_sHorasLabTotales
        Exit Property
    End If
    
    If Me.ColActividades Is Nothing Then
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_IDActividad In Me.ColActividades.Keys
        Set m_Actividad = Me.ColActividades(m_IDActividad)
        
        If IsNumeric(m_Actividad.HorasLaborables) Then
            m_Error = m_Actividad.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            m_HorasLab = m_HorasLab + CDbl(m_Actividad.HorasLaborables)
        End If
        Set m_Actividad = Nothing
    Next
    m_sHorasLabTotales = CStr(m_HorasLab)
    HorasLabTotales = m_sHorasLabTotales
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.HorasLabTotales ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get HorasExtTotales() As String

    Dim m_Actividad As Actividad
    Dim m_IDActividad As Variant
    Dim m_HorasExt As Double
    
    On Error GoTo errores
    If IsNumeric(m_sHorasExtTotales) Then
        HorasExtTotales = m_sHorasExtTotales
        Exit Property
    End If
    If Me.ColActividades Is Nothing Then
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_IDActividad In Me.ColActividades.Keys
        Set m_Actividad = Me.ColActividades(m_IDActividad)
        
        If IsNumeric(m_Actividad.HorasExtras) Then
            m_Error = m_Actividad.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            m_HorasExt = m_HorasExt + CDbl(m_Actividad.HorasExtras)
        End If
        Set m_Actividad = Nothing
    Next
    m_sHorasExtTotales = CStr(m_HorasExt)
    HorasExtTotales = m_sHorasExtTotales
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.HorasExtTotales ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Tecnico() As Tecnico
    
    
    On Error GoTo errores
    Me.Error = ""
    If Not m_ObjTecnico Is Nothing Then
        Set Tecnico = m_ObjTecnico
        Exit Property
    End If
    If Me.ALIASTECNICO = "" Then
        Exit Property
        
    End If
    Set m_ObjTecnico = Constructor.getTecnico(Me.ALIASTECNICO, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Tecnico = m_ObjTecnico
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.Tecnico ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Parte() As Parte

    On Error GoTo errores
    Me.Error = ""
    If Not m_objParte Is Nothing Then
        Set Parte = m_objParte
        Exit Property
    End If
    If Me.IDEVENTO = "" Then
        Exit Property
        
    End If
    Set m_objParte = Constructor.getParte(, Me.IDEVENTO, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Parte = m_objParte
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.Parte ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property



Public Property Get ColAnexos() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    If Not m_objColAnexos Is Nothing Then
        Set ColAnexos = m_objColAnexos
        Exit Property
    End If
    If Me.IDEVENTO = "" Then
        Exit Property
        
    End If
    Set m_objColAnexos = Constructor.getAnexosDeEvento(Me.IDEVENTO, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAnexos = m_objColAnexos
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.ColAnexos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set ColAnexos(ByVal objNewValue As Variant)

    Set m_objColAnexos = objNewValue

End Property

Public Property Get AnexosCerrados() As EnumSino
    
    Dim m_ColAnexosDeEvento As Scripting.Dictionary
    Dim m_ColAnexosDeActividades As Scripting.Dictionary
    
    Dim m_ID As Variant
    Dim m_Anexo As Anexo
    Dim m_Actividad As Actividad
    
    Dim m_URL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_eAnexosCerrados <> 0 Then
        AnexosCerrados = m_eAnexosCerrados
        Exit Property
    End If
    Set m_ColAnexosDeEvento = Me.ColAnexos
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ColAnexosDeEvento Is Nothing Then
        For Each m_ID In m_ColAnexosDeEvento
            Set m_Anexo = m_ColAnexosDeEvento(m_ID)
            If fso.FileExists(m_Anexo.URLAnexo) Then
                If FicheroAbierto(m_Anexo.URLAnexo) Then
                    m_eAnexosCerrados = EnumSino.No
                    AnexosCerrados = m_eAnexosCerrados
                    Exit Property
                End If
            End If
            Set m_Anexo = Nothing
        Next
    End If
    Set m_ColAnexosDeActividades = Me.ColActividades
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ColAnexosDeActividades Is Nothing Then
        For Each m_ID In m_ColAnexosDeActividades
            Set m_Actividad = m_ColAnexosDeActividades(m_ID)
            If m_Actividad.AnexosCerrados = EnumSino.No Then
                m_Error = m_Actividad.Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                m_eAnexosCerrados = EnumSino.No
                AnexosCerrados = m_eAnexosCerrados
                Exit Property
            Else
                m_Error = m_Actividad.Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
            Set m_Actividad = Nothing
        Next
    End If
    m_eAnexosCerrados = EnumSino.Sí
    AnexosCerrados = m_eAnexosCerrados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Parte.AnexosCerrados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColEquiposMedida() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not m_objColEquiposMedida Is Nothing Then
        Set ColEquiposMedida = m_objColEquiposMedida
        Exit Property
    End If
    If Me.IDEVENTO = "" Then
        Exit Property
    End If
    Set m_objColEquiposMedida = Constructor.getEventoEquiposMedida(Me.IDEVENTO, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColEquiposMedida = m_objColEquiposMedida
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.ColEquiposMedida ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColEquiposMedida(ByVal objNewValue As Variant)

    Set m_objColEquiposMedida = objNewValue

End Property

Public Property Get AlguienDeGetronics() As EnumSino
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_eAlguienDeGetronics <> Empty Then
        AlguienDeGetronics = m_eAlguienDeGetronics
        Exit Property
    End If
    m_eAlguienDeGetronics = HayAlguienDeGetronicsEnEvento(Me.IDEVENTO, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    AlguienDeGetronics = m_eAlguienDeGetronics
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.AlguienDeGetronics ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Private Function getHorasEnFranquear( _
                                        Optional ByRef p_Error As String _
                                        ) As Double
    
    Dim m_Hora As Double
    Dim m_Col As Scripting.Dictionary
    Dim m_Fecha As Variant
    Dim m_HorasEnFranquear As Double
    
    On Error GoTo errores
   
    Set m_Col = getHorasMaximasEnActividadesEventoPorDia(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        If m_Col.count > 0 Then
            For Each m_Fecha In m_Col.Keys
                m_Hora = m_Col(m_Fecha)
                m_HorasEnFranquear = m_HorasEnFranquear + m_Hora
            Next
        End If
    End If
    getHorasEnFranquear = m_HorasEnFranquear
    Exit Function
errores:
    If Err.Number <> 0 Then
        p_Error = "El método getHorasEnFranquear ha devuelto el error: " & Err.Description
    End If
End Function



Private Function MotivosNoFranqueable(Optional p_Error As String) As String
    
    Dim m_TieneAcividades As EnumSino
    Dim m_TieneMateriales As EnumSino
    Dim m_TodosMatRecepcionados As EnumSino
    'ha de no estar franqueado
    'Ha de tener al menos una actividad
    'Si tiene algún material, han de estar todos recepcionados
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.IDEVENTO = "" Then
        Exit Function
    End If
    m_TieneAcividades = Me.TieneActividades
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TieneAcividades = EnumSino.No Then
        MotivosNoFranqueable = "No se puede franquear un evento que no tenga al menos una actividad"
        Exit Function
    End If
    m_TieneMateriales = Me.TieneMateriales
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TieneMateriales = EnumSino.Sí Then
        m_TodosMatRecepcionados = Me.TodosLosMaterialesRecepcionados
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_TodosMatRecepcionados <> EnumSino.Sí Then
            MotivosNoFranqueable = "Al menos hay un material que no tiene la fecha de recepción"
            Exit Function
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Evento.MotivosNoFranqueable ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Function MotivosNoQuitarFranqueo(Optional p_Error As String) As String
    
    Dim m_Factura As Factura
    
    'ha de no estar Facturado
    'ha de estar franqueado
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.IDEVENTO = "" Then
        Exit Function
    End If
    Set m_Factura = Me.Factura
    
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Factura Is Nothing Then
        MotivosNoQuitarFranqueo = "El evento pertenece a la factura num " & m_Factura.IDFactura
        Exit Function
    End If
    If Me.Franqueado = False Then
        MotivosNoQuitarFranqueo = "El evento ha de estar franqueado"
        Exit Function
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Evento.MotivosNoQuitarFranqueo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Public Function Franquear(Optional p_CausaFin As String, Optional p_Error As String) As String
    
    Dim m_CadenaCausasFin As String
    Dim m_FechaFin As String
    Dim m_HoraFin As String
    Dim m_MotivosNoFranqueable As String
    Dim m_MotivosSLA As String
    Dim m_Causa As String
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    m_MotivosNoFranqueable = MotivosNoFranqueable(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_MotivosNoFranqueable <> "" Then
        p_Error = m_MotivosNoFranqueable
        Err.Raise 1000
    End If

    ' Validaciones SLA previas al franqueo
    m_MotivosSLA = MotivosNoFranqueableSLA(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_MotivosSLA <> "" Then
        p_Error = m_MotivosSLA
        Err.Raise 1000
    End If
    If p_CausaFin = "" Then
        p_Error = "No se ha indicado la causa fin de franqueo"
        Err.Raise 1000
    End If
    
    m_FechaFin = FechaFinalCalculada
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HoraFin = "15:00:00"
    m_SQL = "UPDATE TbEventos SET TbEventos.FechaFinal = #" & Format(m_FechaFin, "mm/dd/yyyy") & _
            "#, TbEventos.HORAFINALEVENTO = #" & Format(m_FechaFin, "mm/dd/yyyy") & _
            " 15:0:0#, TbEventos.CAUSAFIN = '" & p_CausaFin & "', TbEventos.Franqueado = True " & _
            "WHERE IDEvento='" & Me.IDEVENTO & "';"
    getdb().Execute m_SQL
    
    Me.CausaFin = p_CausaFin
    Me.FechaFinal = m_FechaFin
    Me.Franqueado = True
    Me.HORAFINALEVENTO = m_FechaFin & " 15:0:0"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Evento.Franquear ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Function MotivosNoFranqueableSLA(Optional p_Error As String) As String
    ' Delega en ValidarSLA(p_EnFranqueo:=True) para mantener las reglas de negocio en un solo lugar.
    ' Además valida que Urgente no sea Null (requisito del contrato de franqueo).
    ' Nota: ValidarSLA retorna mensajes acumulados cuando hay múltiples violaciones SLA.
    Dim m_Resultado As String
    Dim m_Mensaje As String
    Dim m_MensajeUrgente As String

    On Error GoTo errores

    p_Error = ""
    m_Mensaje = ""
    m_MensajeUrgente = ""

    ' Primero: usar ValidarSLA para todas las reglas de franqueo
    m_Resultado = ValidarSLA(p_EnFranqueo:=True, p_Mensaje:=m_Mensaje)

    ' Validación adicional: Urgente no puede ser Null al franquear
    If IsNull(Me.Urgente) Then
        m_MensajeUrgente = "TRCM: falta indicar si es urgente."
    End If

    ' Concatenar mensajes de ValidarSLA + la validación adicional de Urgente
    ' Formato determinista: mensajes separados por vbCrLf, sin vbCrLf final
    If m_MensajeUrgente <> "" Then
        If m_Mensaje <> "" Then
            MotivosNoFranqueableSLA = m_Mensaje & vbCrLf & m_MensajeUrgente
        Else
            MotivosNoFranqueableSLA = m_MensajeUrgente
        End If
    Else
        MotivosNoFranqueableSLA = m_Mensaje
    End If
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Evento.MotivosNoFranqueableSLA ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

'===========================================================
' Función: ValidarSLA
' Adaptador de dominio: convierte los campos Boolean de la entidad
' al contrato de cadenas que espera SLAValidator.ValidarCamposSLA.
' - p_EnFranqueo=False: edición normal (reglas generales + cronológicas)
' - p_EnFranqueo=True: franqueo (todas las reglas, incluyendo contrato completo)
' Retorna: "" si OK, o código de error SLA-xxx
'
' NOTA: las reglas de negocio viven en SLAValidator.ValidarCamposSLA.
'       Esta función solo orquesta la conversión de tipos de la entidad.
'===========================================================
Public Function ValidarSLA( _
    ByVal p_EnFranqueo As Boolean, _
    Optional ByRef p_Mensaje As String, _
    Optional ByRef p_Error As String) As String

    Dim m_IncidenciaAveria As String
    Dim m_EventoConServicio As String

    On Error GoTo errores

    p_Mensaje = ""
    p_Error = ""

    ' Convertir campos Boolean de la entidad al formato String
    ' que espera SLAValidator (normalizado internamente a "SI"/"NO").
    m_IncidenciaAveria = IIf(Me.IncidenciaAveriaOReparacion = True, "SI", "NO")
    m_EventoConServicio = IIf(Me.EventoConServicioAfectado = True, "SI", "NO")

    ValidarSLA = SLAValidator.ValidarCamposSLA( _
        p_EnFranqueo:=p_EnFranqueo, _
        p_FechaRecepcion:=Nz(Me.FechaRecepcionNotificacion, ""), _
        p_FechaInicioContacto:=Nz(Me.FechaInicioContactoCliente, ""), _
        p_IncidenciaAveria:=m_IncidenciaAveria, _
        p_FechaInicioAdquisicion:=Nz(Me.FechaInicioTiempoAdquisicion, ""), _
        p_FechaFinAdquisicion:=Nz(Me.FechaFinTiempoAdquisicion, ""), _
        p_TipoReparacion:=Nz(Me.TipoReparacion, ""), _
        p_Urgente:=Me.Urgente, _
        p_EventoConServicioAfectado:=m_EventoConServicio, _
        p_FechaRestablecimiento:=Nz(Me.FechaRestablecimientoServicio, ""), _
        p_Mensaje:=p_Mensaje, _
        p_Error:=p_Error)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Mensaje = "Error inesperado en ValidarSLA: " & Err.Description
    End If
    p_Error = p_Mensaje
    ValidarSLA = "SLA-999"
    p_Mensaje = p_Mensaje
End Function

Public Function QuitarFranqueo(Optional p_Error As String) As String
    
    Dim m_MotivosNoQuitarFranqueo As String
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    m_MotivosNoQuitarFranqueo = MotivosNoQuitarFranqueo(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_MotivosNoQuitarFranqueo <> "" Then
        p_Error = m_MotivosNoQuitarFranqueo
        Err.Raise 1000
    End If
    m_SQL = "UPDATE TbEventos SET TbEventos.FechaFinal = Null, TbEventos.HORAFINALEVENTO = Null, TbEventos.CAUSAFIN = Null, " & _
            "TbEventos.Franqueado = False " & _
            "WHERE (((TbEventos.IDEvento)='" & Me.IDEVENTO & "'));"
    
    getdb().Execute m_SQL
    
    Me.CausaFin = ""
    Me.FechaFinal = ""
    Me.Franqueado = False
    Me.HORAFINALEVENTO = ""
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Evento.QuitarFranqueo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function InformeRACRegistrarFecha(p_Fecha As String, Optional p_Error As String) As String
    
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    m_SQL = "UPDATE TbEventos SET TbEventos.FechaEnInformeRAC = #" & Format(p_Fecha, "mm/dd/yyyy") & "# " & _
            "WHERE (((TbEventos.IDEvento)='" & Me.IDEVENTO & "'));"
    CurrentDb().Execute m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Evento.InformeRACRegistrarFecha ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Public Function InformeRACQuitarFecha(Optional p_Error As String) As String
    
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    m_SQL = "UPDATE TbEventos SET TbEventos.FechaEnInformeRAC = NUll " & _
            "WHERE (((TbEventos.IDEvento)='" & Me.IDEVENTO & "'));"
    CurrentDb().Execute m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Evento.InformeRACQuitarFecha ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function MotivoNoOK( _
                                Optional ByRef p_ObjEventoAlInicio As Evento, _
                                Optional ByRef p_Error As String _
                                ) As String
    Dim m_ObjTecnico As Tecnico
    Dim m_ComprobarTecnico As EnumSino
    Dim m_TecnicoConPrecioActivo As EnumSino
    On Error GoTo errores
        
    With Me
        If .ParaAlta = 0 Then
            MotivoNoOK = "No se sabe si es alta o edición"
            Exit Function
        End If
        If Me.ParaAlta = EnumSino.No Then
            If p_ObjEventoAlInicio Is Nothing Then
                Set p_ObjEventoAlInicio = Constructor.getEvento(Me.IDEVENTO, p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
            If p_ObjEventoAlInicio Is Nothing Then
                MotivoNoOK = "Es la edición del evento y éste no aparece registrado previamente"
                Exit Function
            End If
            If IsDate(Me.FechaFacturado) Then
                p_Error = Me.Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                MotivoNoOK = "Es la edición del evento y éste no puede estar facturado"
                Exit Function
            Else
                p_Error = Me.Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
            If Me.Franqueado = True Then
                MotivoNoOK = "Es la edición del evento y éste no puede estar franqueado"
                Exit Function
            End If
        Else
            Me.IDEVENTO = Me.IDEventoCalculado
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Me.IDEVENTO = "" Then
                MotivoNoOK = "No se ha podido obtener el IDEvento"
                Exit Function
            End If
        End If
        If Me.NODO = "" Then
            MotivoNoOK = "Es obligatorio NODO"
            Exit Function
        End If
        If Me.BUI = "" Then
            MotivoNoOK = "Es obligatorio BUI"
            Exit Function
        End If
        If Me.SubSistema = "" Then
            MotivoNoOK = "Es obligatorio SubSistema"
            Exit Function
        End If
        If Me.IDEquipo = "" Then
            MotivoNoOK = "Es obligatorio Equipo"
            Exit Function
        End If
        If Me.TipoEvento = "" Then
            MotivoNoOK = "Es obligatorio TIPOEVENTO"
            Exit Function
        End If
        If Me.Criticidad = "" Then
            MotivoNoOK = "Es obligatorio CRITICIDAD"
            Exit Function
        End If
        If Me.ALIASTECNICO = "" Then
            MotivoNoOK = "Es obligatorio ALIASTECNICO"
            Exit Function
        End If
        If Me.Originador = "" Then
            MotivoNoOK = "Es obligatorio ORIGINADOR"
            Exit Function
        End If
        If Me.Descripcion = "" Then
            MotivoNoOK = "Es obligatorio Descripcion"
            Exit Function
        End If
        If Me.FECHAALTAEVENTO = "" Then
            Me.FECHAALTAEVENTO = Date
            
        End If
        If Me.HORAINICIALEVENTO = "" Then
            Me.HORAINICIALEVENTO = "7:30"
        End If
        If Me.ParaAlta = EnumSino.Sí Then
            m_ComprobarTecnico = EnumSino.Sí
        Else
            If p_ObjEventoAlInicio.ALIASTECNICO <> Me.ALIASTECNICO Then
                m_ComprobarTecnico = EnumSino.Sí
            End If
        End If
        If m_ComprobarTecnico = EnumSino.Sí Then
            Set m_ObjTecnico = Constructor.getTecnico(Me.ALIASTECNICO, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjTecnico Is Nothing Then
                MotivoNoOK = "El técnico no aparece registrado"
                Exit Function
            End If
            If Me.ParaAlta = EnumSino.Sí Then
                If IsDate(m_ObjTecnico.FechaBaja) Then
                    MotivoNoOK = "El técnico seleccionado está dado de baja"
                    Exit Function
                End If
                m_TecnicoConPrecioActivo = m_ObjTecnico.ConPrecioActivo
                p_Error = m_ObjTecnico.Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If m_TecnicoConPrecioActivo = EnumSino.No Then
                    MotivoNoOK = "El técnico seleccionado no tiene un precio/hora activo"
                    Exit Function
                End If
            End If
            
        End If
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function

Public Function Registrar( _
                            Optional ByRef p_ObjEventoAlInicio As Evento, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If Me.ParaAlta = 0 Then
        p_Error = "No se sabe si es alta o edición"
        Err.Raise 1000
    End If
    
    m_MotivoNoOK = MotivoNoOK(p_ObjEventoAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If Me.ParaAlta = EnumSino.Sí Then
        m_SQL = "TbEventos"
        
    Else
        m_SQL = "SELECT * FROM TbEventos " & _
            "WHERE IDEVENTO='" & Me.IDEVENTO & "';"
        
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Me.ParaAlta = EnumSino.No Then
            If .EOF Then
                p_Error = "No se ha encontrado el registro"
                Err.Raise 1000
            End If
        End If
        If Me.ParaAlta = EnumSino.Sí Then
            .AddNew
                .Fields("IDEvento") = Me.IDEVENTO
                .Fields("NODO") = Me.NODO
                .Fields("BUI") = Me.BUI
                .Fields("SUBSISTEMA") = Me.SubSistema
                .Fields("IDEquipo") = Me.IDEquipo
                If Me.PMPR <> "" Then
                    .Fields("PMPR") = Me.PMPR
                End If
                .Fields("TIPOEVENTO") = Me.TipoEvento
                .Fields("CRITICIDAD") = Me.Criticidad
                .Fields("SUBSISTEMA") = Me.SubSistema
                .Fields("ALIASTECNICO") = Me.ALIASTECNICO
                .Fields("ORIGINADOR") = Me.Originador
                .Fields("FECHAALTAEVENTO") = Me.FECHAALTAEVENTO
                .Fields("HORAINICIALEVENTO") = Me.HORAINICIALEVENTO
                If Me.TIEMPORESPUESTAEVENTO = "" Then
                    Me.TIEMPORESPUESTAEVENTO = "0"
                End If
                .Fields("TIEMPORESPUESTAEVENTO") = Me.TIEMPORESPUESTAEVENTO
                
                .Fields("DESCRIPCION") = Me.Descripcion
                If Me.Contacto <> "" Then
                    .Fields("Contacto") = Me.Contacto
                End If
                
                .Fields("FechaRegistroAlta") = Now()
                .Fields("FechaRegistroModificacion") = Now()
                If Me.Notas <> "" Then
                    .Fields("Notas") = Me.Notas
                End If
                '===========================================================
                ' Campos SLA - Spec-001
                '===========================================================
                If IsDate(Me.FechaRecepcionNotificacion) Then
                    .Fields("FechaRecepcionNotificacion") = CDate(Me.FechaRecepcionNotificacion)
                End If
                If IsDate(Me.FechaInicioContactoCliente) Then
                    .Fields("FechaInicioContactoCliente") = CDate(Me.FechaInicioContactoCliente)
                End If
                .Fields("IncidenciaAveriaOReparacion") = Me.IncidenciaAveriaOReparacion
                If IsDate(Me.FechaInicioTiempoAdquisicion) Then
                    .Fields("FechaInicioTiempoAdquisicion") = CDate(Me.FechaInicioTiempoAdquisicion)
                End If
                If IsDate(Me.FechaFinTiempoAdquisicion) Then
                    .Fields("FechaFinTiempoAdquisicion") = CDate(Me.FechaFinTiempoAdquisicion)
                End If
                If Me.TipoReparacion <> "" Then
                    .Fields("TipoReparacion") = Me.TipoReparacion
                End If
                .Fields("Urgente") = Me.Urgente
                .Fields("EventoConServicioAfectado") = Me.EventoConServicioAfectado
                If IsDate(Me.FechaRestablecimientoServicio) Then
                    .Fields("FechaRestablecimientoServicio") = CDate(Me.FechaRestablecimientoServicio)
                End If
                .Fields("TipoRepInsitu") = Me.TipoRepInsitu
                .Fields("TipoRepNoSMT") = Me.TipoRepNoSMT
                .Fields("TipoRepValvulas") = Me.TipoRepValvulas
            .Update
        Else
            .Edit
                .Fields("NODO") = Me.NODO
                .Fields("BUI") = Me.BUI
                .Fields("SUBSISTEMA") = Me.SubSistema
                .Fields("IDEquipo") = Me.IDEquipo
                If Me.PMPR <> "" Then
                    .Fields("PMPR") = Me.PMPR
                Else
                    .Fields("PMPR") = Null
                End If
                .Fields("TIPOEVENTO") = Me.TipoEvento
                .Fields("CRITICIDAD") = Me.Criticidad
                .Fields("SUBSISTEMA") = Me.SubSistema
                .Fields("ALIASTECNICO") = Me.ALIASTECNICO
                .Fields("ORIGINADOR") = Me.Originador
                .Fields("FECHAALTAEVENTO") = Me.FECHAALTAEVENTO
                .Fields("HORAINICIALEVENTO") = Me.HORAINICIALEVENTO
                If Me.TIEMPORESPUESTAEVENTO = "" Then
                    Me.TIEMPORESPUESTAEVENTO = "0"
                End If
                .Fields("TIEMPORESPUESTAEVENTO") = Me.TIEMPORESPUESTAEVENTO
                .Fields("DESCRIPCION") = Me.Descripcion
                If Me.Contacto <> "" Then
                    .Fields("Contacto") = Me.Contacto
                Else
                    .Fields("Contacto") = Null
                End If
                .Fields("FechaRegistroModificacion") = Now()
                If Me.Notas <> "" Then
                    .Fields("Notas") = Me.Notas
                Else
                    .Fields("Notas") = Null
                End If
            .Update
        
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
       
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha devuelto el error: " & Err.Description
    End If
End Function

Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
    
'    'no puede estar en una factura
'
'
'    On Error GoTo errores
'    'TODO:hacer
'    Me.Error = ""
'    If Me.EsActivo = EnumSino.no Then
'        p_Error = "La Edición a la que pertenece el Plan de Contingencia no es activa"
'        Err.Raise 1000
'    End If
'    If Me.UsuarioConectadoAutorizado = EnumSino.no Then
'        p_Error = Me.Error
'        If p_Error <> "" Then
'            Err.Raise 1000
'        End If
'        p_Error = "El perfil de su usuario no tiene autorización para realizar esta operación"
'        Err.Raise 1000
'    End If
'
'
'    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PC.ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function EquipoBienSituado( _
                                    Optional ByRef p_Error As String _
                                    ) As EnumSino
    'sirve para comprobar que el IDEquipo que elegimos tiene el mismo NODO,BUI,SUBSISTEMA que el EVENTO
    
    Dim m_Equipo As Equipo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.BUI = "" Then
        Exit Function
    End If
    If Me.NODO = "" Then
        Exit Function
    End If
    If Me.SubSistema = "" Then
        Exit Function
    End If
    If Me.IDEquipo = "" Then
        Exit Function
    End If
    Set m_Equipo = Constructor.getEquipo(Me.IDEquipo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Equipo Is Nothing Then
        EquipoBienSituado = EnumSino.No
        Exit Function
    End If
    With m_Equipo
        If .NODO <> Me.NODO Then
            EquipoBienSituado = EnumSino.No
            Exit Function
        End If
        If .BUI <> Me.BUI Then
            EquipoBienSituado = EnumSino.No
            Exit Function
        End If
        If .SubSistema <> Me.SubSistema Then
            EquipoBienSituado = EnumSino.No
            Exit Function
        End If
    End With
    
    EquipoBienSituado = EnumSino.Sí
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método EquipoBienSituado ha devuelto el error: " & Err.Description
     End If
                                
End Function

Public Function RegistrarEquipoMedida( _
                                        p_IDEquipoMedida As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ID As String
    Dim m_Fecha As String
    Dim m_CalibracionActual As EquipoMedidaCalibracion
    Dim m_IDCalibracion As String
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If p_IDEquipoMedida = "" Then
        p_Error = "No hay datos del equipo de medida"
        Err.Raise 1000
    End If
    If Me.FechaFacturado <> "" Then
        p_Error = "El evento está facturado"
        Err.Raise 1000
    End If
    If Me.Franqueado = True Then
        p_Error = "El evento está franqueado"
        Err.Raise 1000
    End If
    m_Fecha = Date
    Set m_CalibracionActual = getCalibracionPorFecha(p_IDEquipoMedida, Me.IDEVENTO, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_CalibracionActual Is Nothing Then
        m_IDCalibracion = m_CalibracionActual.IDCalibracion
    End If
    m_SQL = "SELECT  * " & _
                "FROM TbEventosEquipoMedida " & _
                "WHERE IDEvento='" & Me.IDEVENTO & _
                "' AND IDEquipoMedida=" & p_IDEquipoMedida & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        m_ID = DameID("TbEventosEquipoMedida", "IDEventoEquipoMedida")
        .AddNew
            .Fields("IDEventoEquipoMedida") = m_ID
            .Fields("IDEquipoMedida") = p_IDEquipoMedida
            If IsNumeric(m_IDCalibracion) Then
                .Fields("IDCalibracion") = m_IDCalibracion
            End If
            .Fields("IDEvento") = Me.IDEVENTO
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Set Me.ColEquiposMedida = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistrarEquipoMedida ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EliminarEquipoMedida( _
                                        p_IDEventoEquipoMedida As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                        
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    If Me.Modificable = EnumSino.No Then
        p_Error = "El evento no es editable"
        Err.Raise 1000
    End If
    If p_IDEventoEquipoMedida = "" Then
        p_Error = "No hay datos del equipo de medida"
        Err.Raise 1000
    End If
    
    m_SQL = "DELETE  * " & _
                "FROM TbEventosEquipoMedida " & _
                "WHERE IDEventoEquipoMedida=" & p_IDEventoEquipoMedida & ";"
    getdb().Execute m_SQL
    
    Set Me.ColEquiposMedida = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EliminarEquipoMedida ha devuelto el error: " & Err.Description
    End If
End Function
Public Function EliminarEquiposMedida( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                        
    Dim m_SQL As String
    Dim m_ID As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If Me.ColEquiposMedida Is Nothing Then
        Exit Function
    End If
    For Each m_ID In Me.ColEquiposMedida
        m_SQL = "DELETE  TbEventosEquipoMedida.* " & _
                "FROM TbEventosEquipoMedida " & _
                "WHERE IDEvento='" & Me.IDEVENTO & _
                "' AND IDEquipoMedida=" & m_ID & ";"
        getdb().Execute m_SQL
    Next
    Set Me.ColEquiposMedida = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EliminarEquiposMedida ha devuelto el error: " & Err.Description
    End If
End Function
Private Function getHorasMaximasEnActividadesEventoPorDia( _
                                                            Optional ByRef p_Error As String _
                                                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_Horas As Double
    Dim m_Fecha As Variant
    
    m_SQL = "SELECT TbActividades.FechaAlta, Max(TbActividades.HorasLaborables) AS MáxDeHorasLaborables " & _
                "FROM TbActividades " & _
                "WHERE IDEvento='" & Me.IDEVENTO & "' " & _
                "GROUP BY FechaAlta;"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                m_Fecha = .Fields("FechaAlta")
                If IsNumeric(Nz(.Fields("MáxDeHorasLaborables"), "")) Then
                    m_Horas = CDbl(.Fields("MáxDeHorasLaborables"))
                    If getHorasMaximasEnActividadesEventoPorDia Is Nothing Then
                        Set getHorasMaximasEnActividadesEventoPorDia = New Scripting.Dictionary
                        getHorasMaximasEnActividadesEventoPorDia.CompareMode = TextCompare
                    End If
                    If Not getHorasMaximasEnActividadesEventoPorDia.Exists(m_Fecha) Then
                        getHorasMaximasEnActividadesEventoPorDia.Add m_Fecha, m_Horas
                    End If
                    
                End If
                
                .MoveNext
            Loop

        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    
    If Err.Number <> 0 Then
        p_Error = "El método getHorasMaximasEnActividadesEventoPorDia ha devuelto el error: " & Err.Description
    End If
    
End Function

Private Function RellenarTbActividadesLotes( _
                                            p_ColTecnicos As Scripting.Dictionary, _
                                            p_ObjActividad As Actividad, _
                                            p_HorasTotales As String, _
                                            p_FechaPrimera As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_ObjColActividades As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_IDActividad As String
    Dim rcdDatos As DAO.Recordset
    Dim m_ObjActividad As Actividad
    Dim m_SQL As String
    
    On Error GoTo errores
    
    m_SQL = "DELETE * FROM TbActividadesLotes;"
    CurrentDb().Execute m_SQL
    m_SQL = "TbActividadesLotes"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    Set m_ObjColActividades = getActividadesPorLotes(p_ColTecnicos, p_ObjActividad, p_HorasTotales, p_FechaPrimera, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjColActividades Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_ObjColActividades
        Set m_ObjActividad = m_ObjColActividades(CStr(m_ID))
        m_IDActividad = m_ObjActividad.IDActividadCalculada
        m_ObjActividad.IDActividad = m_IDActividad
        With rcdDatos
            .AddNew
                .Fields("IDActividad") = m_ObjActividad.IDActividad
                .Fields("IDEvento") = m_ObjActividad.IDEVENTO
                .Fields("ALIASTECNICO") = m_ObjActividad.ALIASTECNICO
                .Fields("Actividad") = m_ObjActividad.Actividad
                .Fields("FechaAlta") = m_ObjActividad.FechaAlta
                .Fields("HorasLaborables") = m_ObjActividad.HorasLaborables
                .Fields("Ubicacion") = m_ObjActividad.Ubicacion
                .Fields("DESCRIPCION") = m_ObjActividad.Descripcion
                
            .Update
        End With
        Set m_ObjActividad = Nothing
    Next
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Evento.RellenarTbActividadesLotes ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function GenerarActividadesPorLotes( _
                                            p_ColTecnicos As Scripting.Dictionary, _
                                            p_ObjActividad As Actividad, _
                                            p_HorasTotales As String, _
                                            p_FechaPrimera As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    
    Dim m_SQL As String
    
    On Error GoTo errores
    RellenarTbActividadesLotes p_ColTecnicos, p_ObjActividad, p_HorasTotales, p_FechaPrimera, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    m_SQL = "INSERT INTO TbActividades ( IDActividad, IDEvento, ALIASTECNICO, Actividad, FechaAlta, HorasLaborables, " & _
            "Ubicacion, DESCRIPCION,Observaciones ) " & _
            "SELECT TbActividadesLotes.IDActividad, TbActividadesLotes.IDEvento, TbActividadesLotes.ALIASTECNICO, " & _
            "TbActividadesLotes.Actividad, TbActividadesLotes.FechaAlta, TbActividadesLotes.HorasLaborables, " & _
            "TbActividadesLotes.Ubicacion, TbActividadesLotes.DESCRIPCION,TbActividadesLotes.Observaciones " & _
            "FROM TbActividadesLotes;"
    CurrentDb().Execute m_SQL
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Evento.GenerarActividadesPorLotes ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Function getActividadesPorLotes( _
                                            p_ColTecnicos As Scripting.Dictionary, _
                                            p_ObjActividad As Actividad, _
                                            p_HorasTotales As String, _
                                            p_FechaPrimera As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    
    Dim m_ObjColActividades As Scripting.Dictionary
    Dim m_ObjActividad As Actividad
    Dim m_HoraActividad As String
    Dim m_ObjTecnico As Tecnico
    Dim m_Alias As String
    Dim dblHorasPorTecnico As Double
    Dim dblHorasRestantes As Double
    Dim dblHorasTotales As Double
    Dim m_ID As Variant
    Dim m_ColTecnicosfechas As Scripting.Dictionary
    Dim m_ColFechasHoras As Scripting.Dictionary
    Dim m_Var As Variant
    Dim m_Fecha As String
    Dim dblHoras As Double
    Dim i As Long
    
    On Error GoTo errores
    If p_ColTecnicos Is Nothing Then
        p_Error = "Al menos ha de haber un técnico"
        Err.Raise 1000
    End If
    If p_ObjActividad Is Nothing Then
        p_Error = "Al menos ha de haber datos generales de la actividad"
        Err.Raise 1000
    End If
    If CDbl(p_HorasTotales) = 0 Then
        p_Error = "Las horas totales son obligatorias"
        Err.Raise 1000
    End If
    dblHorasTotales = CDbl(p_HorasTotales)
    If Not IsDate(p_FechaPrimera) Then
        p_Error = "La Primera Fecha es obligatoria"
        Err.Raise 1000
    End If
    If Me.IDEVENTO = "" Then
        p_Error = "No hay datos del evento"
        Err.Raise 1000
    End If
    If Me.FechaFacturado <> "" Then
        p_Error = "El evento aparece como facturado el " & Me.FechaFacturado
        Err.Raise 1000
    End If
    If Me.Franqueado = True Then
        p_Error = "El evento aparece como franqueado"
        Err.Raise 1000
    End If
    With p_ObjActividad
        If .Actividad = "" Then
            p_Error = "La Actividad es obligatoria"
            Err.Raise 1000
        End If
        If .Ubicacion = "" Then
            p_Error = "La Ubicación es obligatoria"
            Err.Raise 1000
        End If
        If .Descripcion = "" Then
            p_Error = "Se ha de indicar la descripción"
            Err.Raise 1000
        End If
        

    End With
    dblHorasPorTecnico = dblHorasTotales / p_ColTecnicos.count
    For Each m_ID In p_ColTecnicos
        m_Alias = CStr(m_ID)
        Set m_ColFechasHoras = getColFechasHoras(m_Alias, CStr(dblHorasPorTecnico), p_FechaPrimera, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_ColFechasHoras Is Nothing Then
            If m_ColTecnicosfechas Is Nothing Then
                Set m_ColTecnicosfechas = New Scripting.Dictionary
                m_ColTecnicosfechas.CompareMode = TextCompare
            End If
            If Not m_ColTecnicosfechas.Exists(m_ID) Then
                m_ColTecnicosfechas.Add m_ID, m_ColFechasHoras
            End If
        End If
        Set m_ColFechasHoras = Nothing
    Next
    If m_ColTecnicosfechas Is Nothing Then
        Exit Function
    End If
    i = 1
    For Each m_ID In m_ColTecnicosfechas
        Set m_ColFechasHoras = m_ColTecnicosfechas(m_ID)
        If Not m_ColFechasHoras Is Nothing Then
            For Each m_Var In m_ColFechasHoras
                m_Fecha = CStr(m_Var)
                dblHoras = m_ColFechasHoras(m_Var)
                Set m_ObjActividad = New Actividad
                
                m_ObjActividad.Actividad = p_ObjActividad.Actividad
                m_ObjActividad.Ubicacion = p_ObjActividad.Ubicacion
                m_ObjActividad.IDEVENTO = Me.IDEVENTO
                m_ObjActividad.ALIASTECNICO = CStr(m_ID)
                m_ObjActividad.FechaAlta = m_Fecha
                m_ObjActividad.HorasLaborables = dblHoras
                m_ObjActividad.Descripcion = p_ObjActividad.Descripcion
                If getActividadesPorLotes Is Nothing Then
                    Set getActividadesPorLotes = New Scripting.Dictionary
                    getActividadesPorLotes.CompareMode = TextCompare
                End If
                If Not getActividadesPorLotes.Exists(CStr(i)) Then
                    getActividadesPorLotes.Add CStr(i), m_ObjActividad
                End If
                
                i = i + 1
                Set m_ObjActividad = Nothing
                
            Next
        End If
        Set m_ColFechasHoras = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Evento.getActividadesPorLotes ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Function getHorasSinRealizarTecnicoEnDia( _
                                                    p_AliasTecnico As String, _
                                                    p_Fecha As String, _
                                                    Optional p_Error As String _
                                                    ) As String
     
    Dim m_DiaLaborableParaTecnico As EnumSino
    Dim dblHorasRealizadas As Double
    Dim strHorasRealizadas As String
    On Error GoTo errores
    
    If p_AliasTecnico = "" Then
        p_Error = "Se ha de indicar el técnico"
        Err.Raise 1000
    End If
    If Not IsDate(p_Fecha) Then
        p_Error = "Se ha de indicar la fecha"
        Err.Raise 1000
    End If
    m_DiaLaborableParaTecnico = DiaLaborableParaTecnico(p_AliasTecnico, p_Fecha, p_Error)
    If p_Error <> "" Then
        Exit Function
    End If
    If m_DiaLaborableParaTecnico = EnumSino.No Then
        Exit Function
    End If
    strHorasRealizadas = DameHorasTecnico(p_AliasTecnico, p_Fecha, , "Sí")
    If Not IsNumeric(strHorasRealizadas) Then
        p_Error = "No se ha podido obtener las horas realizadas en el día"
        Err.Raise 1000
    End If
    dblHorasRealizadas = CDbl(strHorasRealizadas)
    getHorasSinRealizarTecnicoEnDia = CStr(CDbl(7.5) - dblHorasRealizadas)
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHorasSinRealizarTecnicoEnDia ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Function getColFechasHoras( _
                                    p_IDTecnico As String, _
                                    p_HorasTotales As String, _
                                    p_FechaPrimera As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    
    'getColFechasHoras.add Fecha1,Horas
    Dim m_HorasFecha As String
    Dim dblHorasRestantes As Double
    Dim dblHorasTotales As Double
    Dim dblHorasFecha As Double
    Dim m_Fecha As String
    Dim i As Long
    
    On Error GoTo errores
    If p_IDTecnico = "" Then
        p_Error = "Se ha de indicar el alias del técnico"
        Err.Raise 1000
    End If
    If Not IsNumeric(p_HorasTotales) Then
        p_Error = "Las horas totales son obligatorias"
        Err.Raise 1000
    End If
    If CDbl(p_HorasTotales) = 0 Then
        p_Error = "Las horas totales son obligatorias"
        Err.Raise 1000
    End If
    dblHorasTotales = CDbl(p_HorasTotales)
    If Not IsDate(p_FechaPrimera) Then
        p_Error = "La Primera Fecha es obligatoria"
        Err.Raise 1000
    End If
    dblHorasRestantes = dblHorasTotales
    For i = 0 To 10000
        m_Fecha = DateAdd("d", i, CDate(p_FechaPrimera))
        m_HorasFecha = getHorasSinRealizarTecnicoEnDia(p_IDTecnico, m_Fecha, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If IsNumeric(m_HorasFecha) Then
            dblHorasFecha = CDbl(m_HorasFecha)
            If getColFechasHoras Is Nothing Then
                Set getColFechasHoras = New Scripting.Dictionary
                getColFechasHoras.CompareMode = TextCompare
            End If
            If Not getColFechasHoras.Exists(m_Fecha) Then
                If dblHorasFecha > dblHorasRestantes Then
                    dblHorasFecha = dblHorasRestantes
                    getColFechasHoras.Add m_Fecha, dblHorasFecha
                    dblHorasRestantes = 0
                Else
                    getColFechasHoras.Add m_Fecha, dblHorasFecha
                    dblHorasRestantes = dblHorasRestantes - dblHorasFecha
                End If
                
                If dblHorasRestantes <= 0 Then
                    Exit Function
                End If
            End If
        
        End If
        m_HorasFecha = ""
        dblHorasFecha = 0
    Next
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Evento.getColFechasHoras ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function EliminarAnexos( _
                                Optional p_Error As String _
                                ) As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Anexo As Anexo
    Dim m_AnexosCerrados As EnumSino
    On Error GoTo errores
    
    Me.Error = ""
    
    Set m_Col = Me.ColAnexos
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    m_AnexosCerrados = Me.AnexosCerrados
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_AnexosCerrados <> EnumSino.Sí Then
        p_Error = "Al menos un anexo está abierto"
        Err.Raise 1000
    End If
    For Each m_ID In m_Col
        Set m_Anexo = m_Col(m_ID)
        m_Anexo.Eliminar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_Anexo = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EliminarAnexos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
        .Add "IDEvento"
        .Add "NODO"
        .Add "BUI"
        .Add "SUBSISTEMA"
        .Add "IDEquipo"
        .Add "PMPR"
        .Add "TIPOEVENTO"
        .Add "CRITICIDAD"
        .Add "ALIASTECNICO"
        .Add "ORIGINADOR"
        .Add "FECHAALTAEVENTO"
        .Add "HORAINICIALEVENTO"
        .Add "FechaFinal"
        .Add "HORAFINALEVENTO"
        .Add "TIEMPORESPUESTAEVENTO"
        .Add "CAUSAFIN"
        .Add "DESCRIPCION"
        .Add "CONTACTO"
        .Add "FechaRegistroAlta"
        .Add "FechaRegistroModificacion"
        .Add "Franqueado"
        .Add "IDExportacion"
        .Add "CodImportacion"
        .Add "IDParte"
        .Add "Notas"
        .Add "FechaEnInformeRAC"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDEvento" Then
        getPropiedad = Me.IDEVENTO
        Exit Function
    End If
    If m_NombrePropiedad = "NODO" Then
        getPropiedad = Me.NODO
        Exit Function
    End If
    If m_NombrePropiedad = "BUI" Then
        getPropiedad = Me.BUI
        Exit Function
    End If
    If m_NombrePropiedad = "SUBSISTEMA" Then
        getPropiedad = Me.SubSistema
        Exit Function
    End If
    If m_NombrePropiedad = "IDEquipo" Then
        getPropiedad = Me.IDEquipo
        Exit Function
    End If
    If m_NombrePropiedad = "PMPR" Then
        getPropiedad = Me.PMPR
        Exit Function
    End If
    If m_NombrePropiedad = "TIPOEVENTO" Then
        getPropiedad = Me.TipoEvento
        Exit Function
    End If
    If m_NombrePropiedad = "CRITICIDAD" Then
        getPropiedad = Me.Criticidad
        Exit Function
    End If
    If m_NombrePropiedad = "ALIASTECNICO" Then
        getPropiedad = Me.ALIASTECNICO
        Exit Function
    End If
    If m_NombrePropiedad = "ORIGINADOR" Then
        getPropiedad = Me.Originador
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAALTAEVENTO" Then
        getPropiedad = Me.FECHAALTAEVENTO
        Exit Function
    End If
    If m_NombrePropiedad = "HORAINICIALEVENTO" Then
        getPropiedad = Me.HORAINICIALEVENTO
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinal" Then
        getPropiedad = Me.FechaFinal
        Exit Function
    End If
    If m_NombrePropiedad = "HORAFINALEVENTO" Then
        getPropiedad = Me.HORAFINALEVENTO
        Exit Function
    End If
    If m_NombrePropiedad = "TIEMPORESPUESTAEVENTO" Then
        getPropiedad = Me.TIEMPORESPUESTAEVENTO
        Exit Function
    End If
    If m_NombrePropiedad = "CAUSAFIN" Then
        getPropiedad = Me.CausaFin
        Exit Function
    End If
    If m_NombrePropiedad = "DESCRIPCION" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "CONTACTO" Then
        getPropiedad = Me.Contacto
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRegistroAlta" Then
        getPropiedad = Me.FechaRegistroAlta
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRegistroModificacion" Then
        getPropiedad = Me.FechaRegistroModificacion
        Exit Function
    End If
    If m_NombrePropiedad = "Franqueado" Then
        getPropiedad = Me.Franqueado
        Exit Function
    End If
    If m_NombrePropiedad = "IDExportacion" Then
        getPropiedad = Me.IDExportacion
        Exit Function
    End If
    If m_NombrePropiedad = "CodImportacion" Then
        getPropiedad = Me.CodImportacion
        Exit Function
    End If
    If m_NombrePropiedad = "IDParte" Then
        getPropiedad = Me.IDParte
        Exit Function
    End If
    If m_NombrePropiedad = "Notas" Then
        getPropiedad = Me.Notas
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEnInformeRAC" Then
        getPropiedad = Me.FechaEnInformeRAC
        Exit Function
    End If
    If m_NombrePropiedad = "NombreEquipo" Then
        getPropiedad = Me.NombreEquipo
        Exit Function
    End If
    If m_NombrePropiedad = "FranqueadoParaInforme" Then
        getPropiedad = Me.FranqueadoParaInforme
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Evento.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Evento.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDEvento" Then
        Me.IDEVENTO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NODO" Then
        Me.NODO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "BUI" Then
        Me.BUI = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "SUBSISTEMA" Then
        Me.SubSistema = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEquipo" Then
        Me.IDEquipo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PMPR" Then
        Me.PMPR = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TIPOEVENTO" Then
        Me.TipoEvento = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CRITICIDAD" Then
        Me.Criticidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ALIASTECNICO" Then
        Me.ALIASTECNICO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ORIGINADOR" Then
        Me.Originador = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAALTAEVENTO" Then
        Me.FECHAALTAEVENTO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HORAINICIALEVENTO" Then
        Me.HORAINICIALEVENTO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HORAFINALEVENTO" Then
        Me.HORAFINALEVENTO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinal" Then
        Me.FechaFinal = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TIEMPORESPUESTAEVENTO" Then
        Me.TIEMPORESPUESTAEVENTO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CAUSAFIN" Then
        Me.CausaFin = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DESCRIPCION" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CONTACTO" Then
        Me.Contacto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRegistroAlta" Then
        Me.FechaRegistroAlta = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRegistroModificacion" Then
        Me.FechaRegistroModificacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Franqueado" Then
        Me.Franqueado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDExportacion" Then
        Me.IDExportacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodImportacion" Then
        Me.CodImportacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDParte" Then
        Me.IDParte = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Notas" Then
        Me.Notas = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEnInformeRAC" Then
        Me.FechaEnInformeRAC = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Evento.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Evento.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

'===========================================================
' ValidarCamposSLA
' Thin wrapper de dominio para validación SLA en modo edición.
' Delega en SLAValidator (helper canónico) — p_EnFranqueo=False.
' Mantenida como Public para compatibilidad con llamadores externos.
'
' Para franqueo usar ValidarSLA(p_EnFranqueo:=True) directamente.
'===========================================================
Public Function ValidarCamposSLA( _
    Optional ByRef p_Mensaje As String, _
    Optional ByRef p_Error As String) As String
    ValidarCamposSLA = ValidarSLA(p_EnFranqueo:=False, p_Mensaje:=p_Mensaje, p_Error:=p_Error)
End Function

'===========================================================
' CalcularTRES
' Devuelve las horas entre FechaRecepcionNotificacion y
' FechaInicioContactoCliente (laborables o naturales según
' g_SLA_HorasLaborables)
'===========================================================
Public Function CalcularTRES(Optional ByRef p_Error As String) As Double
    Dim m_Horas As Double
    Dim m_FechaInicio As Date
    Dim m_FechaFin As Date
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    CalcularTRES = 0
    
    If Not IsDate(Me.FechaRecepcionNotificacion) Or Not IsDate(Me.FechaInicioContactoCliente) Then
        Exit Function
    End If
    
    m_FechaInicio = CDate(Me.FechaRecepcionNotificacion)
    m_FechaFin = CDate(Me.FechaInicioContactoCliente)
    
    ' Si g_SLA_HorasLaborables = Sí ? calcular horas laborables
    ' Si g_SLA_HorasLaborables = No ? calcular horas naturales (diferencia simple)
    ' Por defecto, calculamos diferencia simple en horas
    m_Horas = DateDiff("h", m_FechaInicio, m_FechaFin)
    
    ' Si la diferencia es menor a 0, ajustamos
    If m_Horas < 0 Then m_Horas = 0
    
    CalcularTRES = CDbl(m_Horas)
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error al calcular TRES: " & Err.Description
    End If
    CalcularTRES = 0
End Function

'===========================================================
' CumpleTRES
' Devuelve True si el TRES cumple el objetivo según criticidad
' Criticidad 1 (Crítica): 1 hora
' Criticidad 3 (Alta): 3 horas
' Criticidad 5 (Media): 8 horas
'===========================================================
Public Function CumpleTRES(Optional ByRef p_Error As String) As Boolean
    Dim m_TRES As Double
    Dim m_Objetivo As Double
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    CumpleTRES = False
    
    m_TRES = CalcularTRES(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    ' Si no hay TRES calculado, no cumple
    If m_TRES = 0 Then
        Exit Function
    End If
    
    ' Determinar objetivo según criticidad
    Select Case Me.Criticidad
        Case "1"
            m_Objetivo = 1  ' Crítica: 1 hora
        Case "2"
            m_Objetivo = 2
        Case "3"
            m_Objetivo = 3  ' Alta: 3 horas
        Case "4"
            m_Objetivo = 5
        Case "5"
            m_Objetivo = 8  ' Media: 8 horas
        Case Else
            m_Objetivo = 8  ' Por defecto media
    End Select
    
    CumpleTRES = (m_TRES <= m_Objetivo)
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error al verificar CumpleTRES: " & Err.Description
    End If
    CumpleTRES = False
End Function

'===========================================================
' ActualizarSLAApartirDeFormulario
' Actualiza los campos SLA del evento desde controles de formulario
' (utilizado por FormEventoSLA)
'===========================================================
Public Sub ActualizarSLAApartirDeFormulario( _
    ByVal p_FechaRecepcionNotificacion As String, _
    ByVal p_FechaInicioContactoCliente As String, _
    ByVal p_IncidenciaAveriaOReparacion As EnumSino, _
    ByVal p_FechaInicioTiempoAdquisicion As String, _
    ByVal p_FechaFinTiempoAdquisicion As String, _
    ByVal p_TipoReparacion As String, _
    ByVal p_Urgente As EnumSino, _
    ByVal p_EventoConServicioAfectado As EnumSino, _
    ByVal p_FechaRestablecimientoServicio As String, _
    ByVal p_TipoRepInsitu As EnumSino, _
    ByVal p_TipoRepNoSMT As EnumSino, _
    ByVal p_TipoRepValvulas As EnumSino, _
    Optional ByRef p_Error As String)
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    Me.FechaRecepcionNotificacion = p_FechaRecepcionNotificacion
    Me.FechaInicioContactoCliente = p_FechaInicioContactoCliente
    Me.IncidenciaAveriaOReparacion = p_IncidenciaAveriaOReparacion
    Me.FechaInicioTiempoAdquisicion = p_FechaInicioTiempoAdquisicion
    Me.FechaFinTiempoAdquisicion = p_FechaFinTiempoAdquisicion
    Me.TipoReparacion = p_TipoReparacion
    Me.Urgente = p_Urgente
    Me.EventoConServicioAfectado = p_EventoConServicioAfectado
    Me.FechaRestablecimientoServicio = p_FechaRestablecimientoServicio
    Me.TipoRepInsitu = p_TipoRepInsitu
    Me.TipoRepNoSMT = p_TipoRepNoSMT
    Me.TipoRepValvulas = p_TipoRepValvulas
    
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error al actualizar SLA desde formulario: " & Err.Description
    End If
End Sub

'===========================================================
' GuardarSLA
' Guarda los campos SLA en la base de datos (para uso desde formulario)
'===========================================================
Public Function GuardarSLA(Optional ByRef p_Error As String) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    ' --- FRANQUEADO GUARD (SERVICE LAYER) ---
    If Me.IDEVENTO <> "" Then
        Dim strFlag As String
        strFlag = EstaFranqueado("EV", Me.IDEVENTO)
        If InStr(1, strFlag, "|") = 0 Then  ' not an error
            If strFlag = "Sí" Then
                p_Error = "No se puede guardar SLA de un evento franqueado."
                Err.Raise 1000
            End If
        End If
    End If
    ' --- END GUARD ---
    
    If Me.IDEVENTO = "" Then
        p_Error = "No hay evento seleccionado"
        Err.Raise 1000
    End If
    
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbEventos " & _
            "WHERE TbEventos.IDEvento='" & Me.IDEVENTO & "';"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    
    With rcdDatos
        If .EOF Then
            p_Error = "No se encontró el evento"
            Err.Raise 1000
        End If
        
        .Edit
            ' TRES
            If IsDate(Me.FechaRecepcionNotificacion) Then
                .Fields("FechaRecepcionNotificacion") = CDate(Me.FechaRecepcionNotificacion)
            Else
                .Fields("FechaRecepcionNotificacion") = Null
            End If
            
            If IsDate(Me.FechaInicioContactoCliente) Then
                .Fields("FechaInicioContactoCliente") = CDate(Me.FechaInicioContactoCliente)
            Else
                .Fields("FechaInicioContactoCliente") = Null
            End If
            
            ' TRCM
            If Me.IncidenciaAveriaOReparacion <> 0 Then
                .Fields("IncidenciaAveriaOReparacion") = (Me.IncidenciaAveriaOReparacion = EnumSino.Sí)
            Else
                .Fields("IncidenciaAveriaOReparacion") = Null
            End If
            
            If IsDate(Me.FechaInicioTiempoAdquisicion) Then
                .Fields("FechaInicioTiempoAdquisicion") = CDate(Me.FechaInicioTiempoAdquisicion)
            Else
                .Fields("FechaInicioTiempoAdquisicion") = Null
            End If
            
            If IsDate(Me.FechaFinTiempoAdquisicion) Then
                .Fields("FechaFinTiempoAdquisicion") = CDate(Me.FechaFinTiempoAdquisicion)
            Else
                .Fields("FechaFinTiempoAdquisicion") = Null
            End If
            
            If Me.TipoReparacion <> "" Then
                .Fields("TipoReparacion") = Me.TipoReparacion
            Else
                .Fields("TipoReparacion") = Null
            End If
            
            If Me.Urgente <> 0 Then
                .Fields("Urgente") = (Me.Urgente = EnumSino.Sí)
            Else
                .Fields("Urgente") = Null
            End If
            
            ' TRSS
            If Me.EventoConServicioAfectado <> 0 Then
                .Fields("EventoConServicioAfectado") = (Me.EventoConServicioAfectado = EnumSino.Sí)
            Else
                .Fields("EventoConServicioAfectado") = Null
            End If
            
            If IsDate(Me.FechaRestablecimientoServicio) Then
                .Fields("FechaRestablecimientoServicio") = CDate(Me.FechaRestablecimientoServicio)
            Else
                .Fields("FechaRestablecimientoServicio") = Null
            End If
            
            ' SLA-4
            If Me.TipoRepInsitu <> 0 Then
                .Fields("TipoRepInsitu") = (Me.TipoRepInsitu = EnumSino.Sí)
            Else
                .Fields("TipoRepInsitu") = Null
            End If
            
            If Me.TipoRepNoSMT <> 0 Then
                .Fields("TipoRepNoSMT") = (Me.TipoRepNoSMT = EnumSino.Sí)
            Else
                .Fields("TipoRepNoSMT") = Null
            End If
            
            If Me.TipoRepValvulas <> 0 Then
                .Fields("TipoRepValvulas") = (Me.TipoRepValvulas = EnumSino.Sí)
            Else
                .Fields("TipoRepValvulas") = Null
            End If
            
            .Fields("FechaRegistroModificacion") = Now()
        .Update
    End With
    
    rcdDatos.Close
    Set rcdDatos = Nothing
    GuardarSLA = "OK"
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error al guardar SLA: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    GuardarSLA = p_Error
End Function






