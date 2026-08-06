VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiesgoMaterializacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public ID As String
Public IDProyecto As String
Public CodigoRiesgo As String
Public Fecha As String
Public IDEdicion As String
Public EsMaterializacion As String
Public IDNC As String
Public IDPlanContingencia As String
Public ParaNC As String
Public FechaDecison As String
Public Estado As String 'es el estado que va a tener, materializado si se materializa y el que tenga el riesgo cuando se quita

Private m_sIDCaculado As String
Private m_eParaNCCalculado As EnumSiNo
Private m_objEdicion As Edicion
Private m_ObjProyecto As Proyecto
Private m_ObjRiesgoMaterializacionEstadoAnterior As RiesgoMaterializacion
Private m_eUsuarioConectadoAutorizado As EnumSiNo
Private m_eEsActivo As EnumSiNo
Private m_eDecidido As EnumSiNo
Private m_ObjNC As NC
Private m_ObjRiesgo As riesgo
Private m_objRiesgoEnEdicionActiva As riesgo
Private m_eESTADOCalculado As EnumEstadoMaterializado
Private m_eEstadoAnteriorMaterializado As EnumSiNo
Private m_eEsMaterializacionCalcuado As EnumSiNo
Private m_objColCampos As Collection

Public Error As String
Private m_Error As String

Public Property Get NC() As NC

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjNC Is Nothing Then
        Set NC = m_ObjNC
        Exit Property
    End If
    If Me.IDNC = "" Then
       Exit Property
    End If
    Set m_ObjNC = Constructor.getNC(Me.IDNC, m_Error)
    If m_Error <> "" Then
        Err.Raise 100
    End If
    Set NC = m_ObjNC
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoNC.NC ha devuelto el error: " & vbNewLine & Err.description
    End If
    Me.Error = m_Error

End Property

Public Property Set NC(NewValor As Variant)
    Set m_ObjNC = NewValor
    Me.IDNC = m_ObjNC.IDNoConformidad
End Property

Public Property Get IDCaculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDCaculado = DameID("TbRiesgosMaterializaciones", "ID", , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    IDCaculado = m_sIDCaculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoMaterializacion.IDCaculado ha devuelto el error: " & vbNewLine & Err.description
    End If
    Me.Error = m_Error
End Property


Public Property Get Edicion() As Edicion
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Not m_objEdicion Is Nothing Then
        Set Edicion = m_objEdicion
        Exit Property
    End If
    If Me.IDEdicion = "" Then
        Exit Property
    End If
    Set m_objEdicion = Constructor.getEdicion(Me.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Edicion = m_objEdicion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoMaterializacion.Edicion ha devuelto el error: " & vbNewLine & Err.description
    End If
    Me.Error = m_Error
End Property

Public Property Set Edicion(ByVal sNewValue As Variant)

    Set m_objEdicion = sNewValue

End Property

Public Property Get Proyecto() As Proyecto

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjProyecto Is Nothing Then
        Set Proyecto = m_ObjProyecto
        Exit Property
    End If
    If Me.IDProyecto = "" Then
        Exit Property
    End If
    Set m_ObjProyecto = Constructor.getProyecto(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Proyecto = m_ObjProyecto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoMaterializacion.Proyecto ha devuelto el error: " & vbNewLine & Err.description
    End If
    Me.Error = m_Error
End Property

Public Property Set Proyecto(ByVal sNewValue As Variant)

    Set Proyecto = sNewValue

End Property

Public Property Get riesgo() As riesgo

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjRiesgo Is Nothing Then
        Set riesgo = m_ObjRiesgo
        Exit Property
    End If
    If Me.IDEdicion = "" Or Me.CodigoRiesgo = "" Then
        Exit Property
    End If
    Set m_ObjRiesgo = Constructor.getRiesgo(, Me.IDEdicion, Me.CodigoRiesgo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set riesgo = m_ObjRiesgo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoMaterializacion.Riesgo ha devuelto el error: " & vbNewLine & Err.description
    End If
    Me.Error = m_Error
End Property

Public Property Get RiesgoEnEdicionActiva() As riesgo

    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objRiesgoEnEdicionActiva Is Nothing Then
        Set RiesgoEnEdicionActiva = m_objRiesgoEnEdicionActiva
        Exit Property
    End If
    If Me.Proyecto Is Nothing Then
        Exit Property
    End If
    If Me.Proyecto.EdicionActiva Is Nothing Then
        Exit Property
    End If
    
    Set m_objRiesgoEnEdicionActiva = Constructor.getRiesgo(, Me.IDEdicion, Me.riesgo.CodigoRiesgo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RiesgoEnEdicionActiva = m_objRiesgoEnEdicionActiva
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoMaterializacion.RiesgoEnEdicionActiva ha devuelto el error: " & vbNewLine & Err.description
    End If
    Me.Error = m_Error
End Property

Public Property Get RiesgoMaterializacionEstadoAnterior() As RiesgoMaterializacion
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjRiesgoMaterializacionEstadoAnterior Is Nothing Then
        Set RiesgoMaterializacionEstadoAnterior = m_ObjRiesgoMaterializacionEstadoAnterior
        Exit Property
    End If
    
    Set m_ObjRiesgoMaterializacionEstadoAnterior = Constructor.getRiesgoMaterializadoEstadoAnterior(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RiesgoMaterializacionEstadoAnterior = m_ObjRiesgoMaterializacionEstadoAnterior
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoMaterializacion.RiesgoMaterializacionEstadoAnterior ha devuelto el error: " & vbNewLine & Err.description
    End If
    Me.Error = m_Error
End Property

Public Property Get EstadoAnteriorMaterializado() As EnumSiNo
    
    Dim m_RiesgoMaterializacionEstadoAnterior As RiesgoMaterializacion
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eEstadoAnteriorMaterializado <> Empty Then
        RiesgoMaterializacionEstadoAnterior = m_eEstadoAnteriorMaterializado
        Exit Property
    End If
    Set m_RiesgoMaterializacionEstadoAnterior = Me.RiesgoMaterializacionEstadoAnterior
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RiesgoMaterializacionEstadoAnterior Is Nothing Then
        Exit Property
    End If
    m_eEstadoAnteriorMaterializado = m_RiesgoMaterializacionEstadoAnterior.EsMaterializacionCalcuado
    m_Error = m_RiesgoMaterializacionEstadoAnterior.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    EstadoAnteriorMaterializado = m_eEstadoAnteriorMaterializado
    
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoMaterializacion.EstadoAnteriorMaterializado ha devuelto el error: " & vbNewLine & Err.description
    End If
    Me.Error = m_Error
End Property

Public Property Get ParaNCCalculado() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eParaNCCalculado <> Empty Then
        ParaNCCalculado = m_eParaNCCalculado
        Exit Property
    End If
    If Me.ParaNC = "Sí" Then
        m_eParaNCCalculado = EnumSiNo.Sí
    Else
        m_eParaNCCalculado = EnumSiNo.No
    End If
    ParaNCCalculado = m_eParaNCCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoMaterializacion.ParaNCCalculado ha devuelto el error: " & vbNewLine & Err.description
    End If
    Me.Error = m_Error
End Property
Public Property Get UsuarioConectadoAutorizado() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eUsuarioConectadoAutorizado <> 0 Then
        UsuarioConectadoAutorizado = m_eUsuarioConectadoAutorizado
        Exit Property
    End If
    m_eUsuarioConectadoAutorizado = UsuarioAutorizado(Me.riesgo, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    UsuarioConectadoAutorizado = m_eUsuarioConectadoAutorizado
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoMaterializacion.UsuarioConectadoAutorizado ha devuelto el error: " & vbNewLine & Err.description
    End If
    Me.Error = m_Error
End Property

Public Property Get EsActivo() As EnumSiNo
    
    Dim m_ObjEdicionPertenece As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eEsActivo <> Empty Then
        EsActivo = m_eEsActivo
        Exit Property
    End If
    Set m_ObjEdicionPertenece = Me.riesgo.Edicion
    
    m_eEsActivo = m_ObjEdicionPertenece.EsActivo
    EsActivo = m_eEsActivo
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoMaterializacion.EsActivo ha devuelto el error: " & vbNewLine & Err.description
    End If
    Me.Error = m_Error
End Property

Public Property Get EsMaterializacionCalcuado() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eEsMaterializacionCalcuado <> Empty Then
        EsMaterializacionCalcuado = m_eEsMaterializacionCalcuado
        Exit Property
    End If
    If Me.EsMaterializacion = "Sí" Then
        m_eEsMaterializacionCalcuado = EnumSiNo.Sí
    Else
        m_eEsMaterializacionCalcuado = EnumSiNo.No
    End If
    
    
    EsMaterializacionCalcuado = m_eEsMaterializacionCalcuado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoMaterializacion.EsMaterializacionCalcuado ha devuelto el error: " & vbNewLine & Err.description
    End If
    Me.Error = m_Error
End Property


Public Property Get Decidido() As EnumSiNo
    
    Dim m_RiesgoR
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eDecidido <> Empty Then
        Decidido = m_eDecidido
        Exit Property
    End If
    If Me.EsMaterializacion <> "Sí" Then
        Exit Property
    End If
    If IsDate(Me.FechaDecison) And Me.ParaNC <> "" Then
        m_eDecidido = EnumSiNo.Sí
    Else
        m_eDecidido = EnumSiNo.No
    End If
    Decidido = m_eDecidido
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoMaterializacion.EsActivo ha devuelto el error: " & vbNewLine & Err.description
    End If
    Me.Error = m_Error
End Property

Public Function MotivoNoOKAlta( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_RiesgoMaterializacionEstadoAnterior As RiesgoMaterializacion
    On Error GoTo errores
    With Me
        If .IDProyecto = "" Then
            MotivoNoOKAlta = "Se ha de indicar el IDProyecto"
            Exit Function
        End If
        If .IDEdicion = "" Then
            MotivoNoOKAlta = "Se ha de indicar el IDEdicion"
            Exit Function
        End If
        If .CodigoRiesgo = "" Then
            MotivoNoOKAlta = "Se ha de indicar CodigoRiesgo"
            Exit Function
        End If
        If Not IsDate(.Fecha) Then
            MotivoNoOKAlta = "Se ha de indicar La fecha de la materialización o desmaterialización"
            Exit Function
        End If
        If .EsMaterializacion <> "Sí" And .EsMaterializacion <> "No" Then
            MotivoNoOKAlta = "EsMaterializacion ha de ser Sí o No"
            Exit Function
        End If
        If .Estado = "" And .EsMaterializacion = "No" Then
            MotivoNoOKAlta = "Se ha de indicar Estado"
            Exit Function
        End If
        Set m_RiesgoMaterializacionEstadoAnterior = .RiesgoMaterializacionEstadoAnterior
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_RiesgoMaterializacionEstadoAnterior Is Nothing Then
            If .EsMaterializacionCalcuado <> EnumSiNo.Sí Then
                MotivoNoOKAlta = "Es el primer registro de la materialización/desmaterialización de un Riesgo. Ha de ser Materializado"
                Exit Function
            End If
        Else
            If .EsMaterializacionCalcuado = EnumSiNo.Sí Then
                If m_RiesgoMaterializacionEstadoAnterior.EsMaterializacionCalcuado = EnumSiNo.Sí Then
                    MotivoNoOKAlta = "El estado anterior de materialización era materializado. No se puede repetir el estado"
                    Exit Function
                End If
            Else
                If m_RiesgoMaterializacionEstadoAnterior.EsMaterializacionCalcuado = EnumSiNo.No Then
                    MotivoNoOKAlta = "El estado anterior de materialización era No Materializado. No se puede repetir el estado"
                    Exit Function
                End If
            End If
        End If
        
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoMaterializacion.MotivoNoOKAlta ha devuelto el error: " & Err.description
    End If
End Function

Public Function RegistrarAlta( _
                                Optional ByRef p_Error As String _
                             ) As String
                           
    ' Es cuando se materializa un riesgo o se desmaterializa se ha de rellenar
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOKAlta As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    
    m_MotivoNoOKAlta = MotivoNoOKAlta(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOKAlta
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "TbRiesgosMaterializaciones"
    Me.ID = Me.IDCaculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not IsDate(Me.Fecha) Then
        Me.Fecha = Date
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .fields("ID") = Me.ID
            .fields("IDProyecto") = Me.IDProyecto
            .fields("CodigoRiesgo") = Me.CodigoRiesgo
            .fields("IDEdicion") = Me.IDEdicion
            .fields("Fecha") = Me.Fecha
            .fields("EsMaterializacion") = Me.EsMaterializacion
            If Me.IDPlanContingencia <> "" Then
                .fields("IDPlanContingencia") = Me.IDPlanContingencia
            End If
            If Me.EsMaterializacion = "Sí" Then
                Me.Estado = "Materializado"
            End If
            If Me.Estado <> "" Then
                .fields("Estado") = Me.Estado
            End If
            
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Me.Estado = "" Then
        Me.Estado = Me.riesgo.ESTADOCalculadoTexto
        m_SQL = "UPDATE TbRiesgos SET TbRiesgos.Estado = '" & Me.Estado & "' " & _
                "WHERE IDRiesgo=" & Me.riesgo.IDRiesgo & ";"
        getdb().Execute m_SQL
        
         m_SQL = "UPDATE TbRiesgosMaterializaciones SET TbRiesgosMaterializaciones.Estado = '" & Me.Estado & "' " & _
                "WHERE ID=" & Me.ID & ";"
        getdb().Execute m_SQL
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoMaterializacion.RegistrarAlta ha devuelto el error: " & Err.description
    End If
End Function
Public Function RegistrarCambioFechaMaterializacion( _
                                                        p_Fecha As String, _
                                                       Optional ByRef p_Error As String _
                                                    ) As String
                           
    ' Es cuando se materializa un riesgo o se desmaterializa se ha de rellenar
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    
    
    On Error GoTo errores
    
    Me.Error = ""
    If Not IsDate(p_Fecha) Then
        p_Error = "Se ha de indicar la fecha"
        Err.Raise 1000
    End If
    If Me.EsMaterializacionCalcuado <> EnumSiNo.Sí Then
        p_Error = "Está registrado como desmaterialización"
        Err.Raise 1000
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosMaterializaciones " & _
            "WHERE ID=" & Me.ID & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "No parece estar registrada esta materalización"
            Err.Raise 1000
        End If
        .Edit
            .fields("Fecha") = p_Fecha
        .Update
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoMaterializacion.RegistrarCambioFechaMaterializacion ha devuelto el error: " & Err.description
    End If
End Function

Public Function RegistrarPorDecidir( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
   
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    
    If Me.EsMaterializacionCalcuado <> EnumSiNo.Sí Then
        p_Error = "Está registrado como desmaterialización"
        Err.Raise 1000
    End If
    
    Me.FechaDecison = ""
    Me.ParaNC = ""
    Me.IDNC = ""
  
    
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosMaterializaciones " & _
            "WHERE ID=" & Me.ID & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "No parece estar registrada esta materalización"
            Err.Raise 1000
        End If
        .Edit
            .fields("IDNC") = Null
            .fields("ParaNC") = Null
            .fields("FechaDecison") = Null
        .Update
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
        
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoMaterializacion.RegistrarPorDecidir ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.description
    End If
    
End Function

Public Function RegistrarParaNC( _
                                p_IDNC As String, _
                                Optional p_FechaDecision As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
   
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ObjNC As NC
    
    On Error GoTo errores
    
    Me.Error = ""
    
   
    If p_IDNC = "" Then
        p_Error = "No se ha podido determinar la NC"
        Err.Raise 1000
    End If
    If Me.EsMaterializacionCalcuado <> EnumSiNo.Sí Then
        p_Error = "Está registrado como desmaterialización"
        Err.Raise 1000
    End If
    Me.IDNC = p_IDNC
    If Not IsDate(p_FechaDecision) Then
        p_FechaDecision = Date
    End If
    Me.FechaDecison = p_FechaDecision
    Me.ParaNC = "Sí"
    Set m_ObjNC = Constructor.getNC(Me.IDNC, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjNC Is Nothing Then
        p_Error = "No hay ninguna No Conformidad registrada con el ID " & Me.IDNC
        Err.Raise 1000
    End If
    
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosMaterializaciones " & _
            "WHERE ID=" & Me.ID & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "No parece estar registrada esta materalización"
            Err.Raise 1000
        End If
        .Edit
            .fields("IDNC") = Me.IDNC
            .fields("ParaNC") = "Sí"
            .fields("FechaDecison") = Me.FechaDecison
        .Update
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
        
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoMaterializacion.RegistrarParaNC ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.description
    End If
    
End Function
Public Function RegistrarParaNONC( _
                                Optional p_FechaDecision As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    
    If Me.EsMaterializacionCalcuado <> EnumSiNo.Sí Then
        p_Error = "Está registrado como desmaterialización"
        Err.Raise 1000
    End If
    
    If Not IsDate(p_FechaDecision) Then
        p_FechaDecision = Date
    End If
    Me.FechaDecison = p_FechaDecision
    Me.ParaNC = "No"
    Me.IDNC = ""
  
    
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosMaterializaciones " & _
            "WHERE ID=" & Me.ID & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "No parece estar registrada esta materalización"
            Err.Raise 1000
        End If
        .Edit
            .fields("IDNC") = Null
            .fields("ParaNC") = Me.ParaNC
            .fields("FechaDecison") = Me.FechaDecison
        .Update
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
        
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoMaterializacion.RegistrarParaNONC ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.description
    End If
    
End Function
Public Function VincularNC( _
                            p_ObjNC As NC, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    If Me.EsMaterializacionCalcuado <> EnumSiNo.Sí Then
        p_Error = "Está registrado como desmaterialización"
        Err.Raise 1000
    End If
    
    Me.FechaDecison = Date
    Me.IDNC = p_ObjNC.IDNoConformidad
    Me.ParaNC = "Sí"
    
    
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosMaterializaciones " & _
            "WHERE ID=" & Me.ID & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "No parece estar registrada esta materalización"
            Err.Raise 1000
        End If
        .Edit
            .fields("IDNC") = Me.IDNC
            .fields("ParaNC") = Me.ParaNC
            .fields("FechaDecison") = Me.FechaDecison
        .Update
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
        
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoMaterializacion.vincularNC ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.description
    End If
    
End Function
Public Function DesvincularNC( _
                            Optional ByRef p_Error As String _
                            ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    If Me.EsMaterializacionCalcuado <> EnumSiNo.Sí Then
        p_Error = "Está registrado como desmaterialización"
        Err.Raise 1000
    End If
    
    Me.FechaDecison = ""
    Me.IDNC = ""
    Me.ParaNC = ""
    
    
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosMaterializaciones " & _
            "WHERE ID=" & Me.ID & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "No parece estar registrada esta materalización"
            Err.Raise 1000
        End If
        .Edit
            .fields("IDNC") = Null
            .fields("ParaNC") = Null
            .fields("FechaDecison") = Null
        .Update
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoMaterializacion.DesvincularNC ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.description
    End If
End Function



Public Function Eliminar( _
                            Optional ByRef p_Error As String _
                            ) As String
                            
    
    Dim m_SQL As String
    
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_SQL = "DELETE * " & _
            "FROM TbRiesgosMaterializaciones " & _
            "WHERE ID=" & Me.ID & ";"
    getdb().Execute m_SQL
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoMaterializacion.Eliminar ha devuelto el error: " & Err.description
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
        .Add "ID"
        .Add "IDProyecto"
        .Add "IDEdicion"
        .Add "CodigoRiesgo"
        .Add "Fecha"
        .Add "EsMaterializacion"
        .Add "IDNC"
        .Add "IDPlanContingencia"
        .Add "ParaNC"
        .Add "FechaDecison"
        .Add "Estado"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoMaterializacion.ColCampos ha devuelto el error:" & vbNewLine & Err.description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_NombrePropiedad = "ID" Then
        getPropiedad = Me.ID
        Exit Function
    End If
    If m_NombrePropiedad = "IDProyecto" Then
        getPropiedad = Me.IDProyecto
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        getPropiedad = Me.IDEdicion
        Exit Function
    End If
    If m_NombrePropiedad = "CodigoRiesgo" Then
        getPropiedad = Me.CodigoRiesgo
        Exit Function
    End If
    If m_NombrePropiedad = "Fecha" Then
        getPropiedad = Me.Fecha
        Exit Function
    End If
    If m_NombrePropiedad = "EsMaterializacion" Then
        getPropiedad = Me.EsMaterializacion
        Exit Function
    End If
    If m_NombrePropiedad = "IDNC" Then
        getPropiedad = Me.IDNC
        Exit Function
    End If
    If m_NombrePropiedad = "IDPlanContingencia" Then
        getPropiedad = Me.IDPlanContingencia
        Exit Function
    End If
    If m_NombrePropiedad = "ParaNC" Then
        getPropiedad = Me.ParaNC
        Exit Function
    End If
    If m_NombrePropiedad = "FechaDecison" Then
        getPropiedad = Me.FechaDecison
        Exit Function
    End If
    If m_NombrePropiedad = "Estado" Then
        getPropiedad = Me.Estado
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en RiesgoMaterializacion.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoMaterializacion.getPropiedad ha devuelto el error:" & vbNewLine & Err.description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "ID" Then
        Me.ID = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDProyecto" Then
        Me.IDProyecto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        Me.IDEdicion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodigoRiesgo" Then
        Me.CodigoRiesgo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Fecha" Then
        Me.Fecha = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsMaterializacion" Then
        Me.EsMaterializacion = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "IDNC" Then
        Me.IDNC = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDPlanContingencia" Then
        Me.IDPlanContingencia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ParaNC" Then
        Me.ParaNC = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaDecison" Then
        Me.FechaDecison = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Estado" Then
        Me.Estado = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en RiesgoMaterializacion.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoMaterializacion.SetPropiedad ha devuelto el error:" & vbNewLine & Err.description
    End If
End Function






