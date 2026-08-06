VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Expediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public IDExpediente As String
Public IDExpedientePadre As String
Public Nemotecnico As String
Public Titulo As String
Public ImporteLicitacion As String
Public ImporteContratacion As String
Public CodProyecto As String
Public CodExp As String
Public CodExpLargo As String
Public CodS4H As String
Public FechaInicioContrato As String
Public FechaFinContrato As String
Public FechaFinGarantia As String
Public EsAM As String
Public EsLote As String
Public EsBasado As String
Public EsExpediente As String
Public Ordinal As String
Public IdGradoClasificacion As String
Public IDOrganoContratacion As String
Public IDOficinaPrograma As String
Public IDEjercito As String
Public AccesoSharepoint As String
Public Observaciones As String
Public FechaCreacion As String
Public IDUsuarioCreacion As String
Public FechaUltimoCambio As String
Public IDUsuarioUltimoCambio As String
Public Ambito As String
Public NPedido As String
Public IDResponsableCalidad As String
Public AGEDYSAplica As String
Public AGEDYSGenerico As String
Public HPSAplica As String
Public Tipo As String
Public CadenaPecal As String
Public PECAL As String
Public TipoInforme As String
Public POSTAGEDO As String
Public APLICAESTADO As String
Public FECHAINICIOLICITACION As String
Public FECHAOFERTA As String
Public FECHAADJUDICACION As String
Public FechaFirmaContrato As String
Public GARANTIAMESES As String
Public FECHACERTIFICACION As String
Public FECHAPERDIDA As String
Public FECHADESESTIMADA As String
Public Estado As String
Public CodigoActividad As String

Private m_sAdjudicado As String
Private m_eESTADOCalculado As EnumEstados
Private m_sESTADOCalculadoTexto As String
Private m_sESTADOCalculadoTitulo As String
Private m_eTIPOEnum As EnumTipoExpediente
Private m_eTieneRAC As EnumSiNo
Private m_eTienePECAL As EnumSiNo
Private m_sCadenaRACs As String
Private m_sCadenaCorreoRACs As String
Private m_sCadenaJuridicas As String

Private m_ObjJuridicas As Scripting.Dictionary
Private m_ObjRACS As Scripting.Dictionary
Private m_ObjOrganoContratacion As OrganoContratacion
Private m_ObjResponsables As Scripting.Dictionary
Private m_objColUsuarios As Collection
Private m_ObjResponsableCalidad As Usuario
Private m_eFechasIrregulares As EnumSiNo
Private m_sCodigoDocParteExpediente As String
Private m_ObjSuministradores As Scripting.Dictionary

Private m_eEnUTE As EnumSiNo
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get Suministradores() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
     If Not m_ObjSuministradores Is Nothing Then
        Set Suministradores = m_ObjSuministradores
        Exit Property
    End If
    Set m_ObjSuministradores = SuministradoresHelper.getSubcontratistasPorExpediente(p_IDExpediente:=Me.IDExpediente, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Suministradores = m_ObjSuministradores
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.Suministradores ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Sub SuministradoresVaciar()
    Set m_ObjSuministradores = Nothing
End Sub

Public Property Get CadenaRACs() As String
    
    Dim m_ID As Variant
    Dim m_RAC As RAC
    
        
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sCadenaRACs = ""
    If Me.RACs Is Nothing Then
        Exit Property
    End If
    For Each m_ID In Me.RACs
        Set m_RAC = Me.RACs(m_ID)
        If m_sCadenaRACs = "" Then
            m_sCadenaRACs = m_RAC.RAC
        Else
            m_sCadenaRACs = m_sCadenaRACs & "|" & m_RAC.RAC
        End If
        Set m_RAC = Nothing
    Next
    
    CadenaRACs = m_sCadenaRACs
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.CadenaRACs ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get CadenaCorreoRACs() As String
    
    Dim m_ID As Variant
    Dim m_RAC As RAC
    
        
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sCadenaCorreoRACs = ""
    If Me.RACs Is Nothing Then
        Exit Property
    End If
    For Each m_ID In Me.RACs
        Set m_RAC = Me.RACs(m_ID)
        If InStr(1, m_RAC.CORREO, "@") <> 0 Then
            If m_sCadenaCorreoRACs = "" Then
                m_sCadenaCorreoRACs = m_RAC.CORREO
            Else
                m_sCadenaCorreoRACs = m_sCadenaCorreoRACs & ";" & m_RAC.CORREO
            End If
        End If
        
        Set m_RAC = Nothing
    Next
    
    CadenaCorreoRACs = m_sCadenaCorreoRACs
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.CadenaCorreoRACs ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get RACs() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    Set m_ObjRACS = Constructor.getExpedienteRACS(p_IDExpediente:=Me.IDExpediente, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RACs = m_ObjRACS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.RACS ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get OrganoContratacion() As OrganoContratacion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjOrganoContratacion Is Nothing Then
        Set OrganoContratacion = m_ObjOrganoContratacion
        Exit Property
    End If
    If Me.IDOrganoContratacion = "" Then
        Exit Property
    End If
    Set m_ObjOrganoContratacion = Constructor.getOrganoContratacion(p_IDOrganoContratacion:=Me.IDOrganoContratacion, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set OrganoContratacion = m_ObjOrganoContratacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.OrganoContratacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EnUTE() As EnumSiNo
    
    Dim m_CadenaJuridicas As String
    Dim dato As Variant
    Dim m_Juridica As Variant
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eEnUTE <> Empty Then
        EnUTE = m_eEnUTE
        Exit Property
    End If
    m_CadenaJuridicas = Me.CadenaJuridicas
    If InStr(1, m_CadenaJuridicas, "|") = 0 Then
        m_eEnUTE = EnumSiNo.No
        EnUTE = m_eEnUTE
        Exit Property
    End If
    dato = Split(m_CadenaJuridicas, "|")
    For Each m_Juridica In dato
        If CStr(m_Juridica) <> "TDE" And CStr(m_Juridica) <> "TSOL" And CStr(m_Juridica) <> "TME" Then
            m_eEnUTE = EnumSiNo.Sí
            EnUTE = m_eEnUTE
            Exit Property
        End If
    Next
    m_eEnUTE = EnumSiNo.No
    EnUTE = m_eEnUTE
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.EnUTE ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Juridicas() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    Set m_ObjJuridicas = Constructor.getExpedienteJuridicas(p_IDExpediente:=Me.IDExpediente, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Juridicas = m_ObjJuridicas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.Juridicas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get CadenaJuridicas() As String
    
    Dim m_ID As Variant
    Dim m_Juridica As Suministrador
    Dim m_Juridicas As Scripting.Dictionary
        
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sCadenaJuridicas = ""
    Set m_Juridicas = Me.Juridicas
    If m_Juridicas Is Nothing Then
        
        Exit Property
    End If
    For Each m_ID In m_Juridicas
        Set m_Juridica = m_Juridicas(m_ID)
        If m_Juridica.Nemotecnico <> "" Then
            If m_sCadenaJuridicas = "" Then
                m_sCadenaJuridicas = m_Juridica.Nemotecnico
            Else
                m_sCadenaJuridicas = m_sCadenaJuridicas & "|" & m_Juridica.Nemotecnico
            End If
        End If
        
        Set m_Juridica = Nothing
    Next
    
    CadenaJuridicas = m_sCadenaJuridicas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.CadenaJuridicas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get Responsables() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    Set m_ObjResponsables = Constructor.getExpedienteResponsables(p_IDExpediente:=Me.IDExpediente, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Responsables = m_ObjResponsables
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.Responsables ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FechasIrregulares() As EnumSiNo
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    With Me
        If IsDate(.FECHADESESTIMADA) Then
            m_eFechasIrregulares = EnumSiNo.No
            FechasIrregulares = m_eFechasIrregulares
            Exit Property
        End If
        If IsDate(.FECHAPERDIDA) Then
            m_eFechasIrregulares = EnumSiNo.No
            FechasIrregulares = m_eFechasIrregulares
            Exit Property
        End If
        If IsDate(.FechaInicioContrato) And Not IsDate(.FechaFinContrato) Then
            m_eFechasIrregulares = EnumSiNo.Sí
            FechasIrregulares = m_eFechasIrregulares
            Exit Property
        End If
        If IsDate(.FechaFinContrato) And Not IsDate(.FechaInicioContrato) Then
            m_eFechasIrregulares = EnumSiNo.Sí
            FechasIrregulares = m_eFechasIrregulares
            Exit Property
        End If
        If IsDate(.FechaInicioContrato) And Not IsNumeric(.GARANTIAMESES) Then
            m_eFechasIrregulares = EnumSiNo.Sí
            FechasIrregulares = m_eFechasIrregulares
            Exit Property
        End If
        If IsDate(.FechaFirmaContrato) Then
            If Not IsDate(.FechaInicioContrato) Or Not IsDate(.FechaFinContrato) Or Not IsNumeric(.GARANTIAMESES) Then
                m_eFechasIrregulares = EnumSiNo.Sí
                FechasIrregulares = m_eFechasIrregulares
                Exit Property
            End If
        End If
       
        If IsDate(.FECHACERTIFICACION) Then
            If Not IsDate(.FechaInicioContrato) Or Not IsDate(.FechaFinContrato) Or _
                Not IsNumeric(.GARANTIAMESES) Or Not IsDate(.FechaFirmaContrato) Then
                m_eFechasIrregulares = EnumSiNo.Sí
                FechasIrregulares = m_eFechasIrregulares
                Exit Property
            End If
        End If
        
        
    End With
    m_eFechasIrregulares = EnumSiNo.No
    FechasIrregulares = m_eFechasIrregulares
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.FechasIrregulares ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Adjudicado() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sAdjudicado <> "" Then
        Adjudicado = m_sAdjudicado
        Exit Property
    End If
    If IsDate(Me.FECHADESESTIMADA) Then
        m_sAdjudicado = "No"
        Adjudicado = m_sAdjudicado
        Exit Property
    End If
    If IsDate(Me.FECHAPERDIDA) Then
        m_sAdjudicado = "No"
        Adjudicado = m_sAdjudicado
        Exit Property
    End If
    If IsDate(Me.FECHAADJUDICACION) Then
        m_sAdjudicado = "Sí"
    Else
        m_sAdjudicado = "No"
    End If
    Adjudicado = m_sAdjudicado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.Adjudicado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ESTADOCalculado() As EnumEstados
    
    Dim m_FechaFinal As String
    Dim m_Hoy As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_Hoy = Date
    With Me
        If .APLICAESTADO <> "Sí" And .APLICAESTADO <> "No" Then
            GoTo Fechas
        End If
        If .APLICAESTADO <> "Sí" Then
            m_eESTADOCalculado = EnumEstados.NoAplica
            ESTADOCalculado = m_eESTADOCalculado
            Exit Property
        End If
Fechas:
        If FechasIrregulares = EnumSiNo.Sí Then
             m_eESTADOCalculado = EnumEstados.Desconocido
            ESTADOCalculado = m_eESTADOCalculado
            Exit Property
        End If
        If IsDate(.FECHADESESTIMADA) Then
            m_eESTADOCalculado = EnumEstados.Desestimado
            ESTADOCalculado = m_eESTADOCalculado
            Exit Property
        End If
        If IsDate(.FECHAPERDIDA) Then
            m_eESTADOCalculado = EnumEstados.Perdido
            ESTADOCalculado = m_eESTADOCalculado
            Exit Property
        End If
        If Not IsDate(.FECHAADJUDICACION) Then
            m_eESTADOCalculado = EnumEstados.Oferta
            ESTADOCalculado = m_eESTADOCalculado
            Exit Property
        End If
        If IsDate(.FECHAADJUDICACION) Then
            If Not IsDate(.FechaFinContrato) Then
                m_eESTADOCalculado = EnumEstados.Adjudicada
                ESTADOCalculado = m_eESTADOCalculado
                Exit Property
            End If
        End If
        
        
        
        If IsDate(.FECHACERTIFICACION) Then
            m_FechaFinal = .FECHACERTIFICACION
        Else
            m_FechaFinal = .FechaFinContrato
        End If
        If CDate(m_Hoy) <= CDate(m_FechaFinal) Then
            m_eESTADOCalculado = EnumEstados.EnEjecucion
            ESTADOCalculado = m_eESTADOCalculado
            Exit Property
        End If
        If CDate(m_Hoy) > CDate(.FechaFinGarantia) Then
            m_eESTADOCalculado = EnumEstados.Cerrado
        Else
            m_eESTADOCalculado = EnumEstados.EnGarantia
            
        End If
        ESTADOCalculado = m_eESTADOCalculado
        Exit Property
        
    End With
    
   
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.ESTADOCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ESTADOCalculadoTexto() As String
    
    Dim m_EstadoCalculado As EnumEstados
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_EstadoCalculado = Me.ESTADOCalculado
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sESTADOCalculadoTexto = m_ObjEntorno.ColEstadosTitulo(CStr(m_EstadoCalculado))
    ESTADOCalculadoTexto = m_sESTADOCalculadoTexto
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.ESTADOCalculadoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ESTADOCalculadoTitulo() As String
    
    Dim m_EstadoCalculado As EnumEstados
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_EstadoCalculado = Me.ESTADOCalculado
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sESTADOCalculadoTitulo = m_ObjEntorno.ColEstadosTituloMejorado(CStr(m_EstadoCalculado))
    ESTADOCalculadoTitulo = m_sESTADOCalculadoTitulo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.ESTADOCalculadoTitulo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ResponsableCalidad() As Usuario
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not m_ObjResponsableCalidad Is Nothing Then
        Set ResponsableCalidad = m_ObjResponsableCalidad
        Exit Property
    End If
    If Me.IDResponsableCalidad = "" Then
        Exit Property
    End If
    Set m_ObjResponsableCalidad = Constructor.getUsuario(p_Id:=Me.IDResponsableCalidad, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ResponsableCalidad = m_ObjResponsableCalidad
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.ResponsableCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    
End Property
Public Property Get CodigoDocParteExpediente() As String
    
    Dim m_Tipo As String
    Dim m_Area As String
    Dim m_Centro As String
    Dim m_PrimerParteExp As String
    Dim m_Año As String
    Dim m_ParteVariable As String
    Dim dato As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sCodigoDocParteExpediente <> "" Then
        CodigoDocParteExpediente = m_sCodigoDocParteExpediente
        Exit Property
    End If
    If Me.CodExp = "" Then
        Exit Property
    End If
    m_Tipo = "IS"
    m_Area = "E"
    m_Centro = "CCC"
    If InStr(1, Me.CodExp, "/") = 0 Then
        m_Error = "El expediente ha de tener la forma ????/??"
        Err.Raise 1000
    End If
    dato = Split(Me.CodExp, "/")
    m_PrimerParteExp = CStr(dato(0))
    m_Año = CStr(dato(1))
    m_ParteVariable = m_Año & m_PrimerParteExp & m_Centro
    m_sCodigoDocParteExpediente = m_Tipo & m_ParteVariable & m_Area & "NN"
    CodigoDocParteExpediente = m_sCodigoDocParteExpediente
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.CodigoDocParteExpediente ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    
End Property


Private Function getCodigoDocumento( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                     
    Dim m_Tipo As String
    Dim m_Area As String
    Dim m_Centro As String
    Dim m_PrimerParteExp As String
    Dim m_Año As String
    Dim m_ParteVariable As String
    Dim m_ObjDocumento As Documento
    Dim m_ObjColDocumentos As Scripting.Dictionary
    Dim m_IDDocumento As Variant
    Dim m_EdicionMaxima As Integer
    Dim m_CodigoSinEdicion As String
    Dim dato
    On Error GoTo errores
    
    m_Tipo = "IS"
    m_Area = "E"
    m_Centro = "CCC"
    If InStr(1, Me.CodExp, "/") = 0 Then
        p_Error = "El expediente ha de tener la forma ????/??"
        Err.Raise 1000
    End If
    dato = Split(Me.CodExp, "/")
    m_PrimerParteExp = CStr(dato(0))
    m_Año = CStr(dato(1))
    m_ParteVariable = m_Año & m_PrimerParteExp & m_Centro
    m_CodigoSinEdicion = m_Tipo & m_ParteVariable & m_Area & "NN"
    Set m_ObjColDocumentos = Constructor.getDocumentosbyCodigo(m_CodigoSinEdicion, p_Error)
   If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjColDocumentos Is Nothing Then
        For Each m_IDDocumento In m_ObjColDocumentos.keys
            If IsObject(m_ObjColDocumentos(m_IDDocumento)) Then
                Set m_ObjDocumento = m_ObjColDocumentos(m_IDDocumento)
                If IsNumeric(m_ObjDocumento.Edicion) Then
                    If CInt(m_ObjDocumento.Edicion) > m_EdicionMaxima Then
                        m_EdicionMaxima = CInt(m_ObjDocumento.Edicion)
                    End If
                End If
                Set m_ObjDocumento = Nothing
            End If
        Next
    End If
    getCodigoDocumento = m_CodigoSinEdicion & Format(m_EdicionMaxima + 1, "00")
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCodigoDocumento ha devuelto el error: " & Err.Description
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
        .Add "IDExpediente"
        .Add "IDExpedientePadre"
        .Add "Nemotecnico"
        .Add "Titulo"
        .Add "ImporteLicitacion"
        .Add "ImporteContratacion"
        .Add "CodProyecto"
        .Add "CodExp"
        .Add "CodExpLargo"
        .Add "CodS4H"
        .Add "FechaInicioContrato"
        .Add "FechaFinContrato"
        .Add "FechaFinGarantia"
        .Add "EsAM"
        .Add "EsLote"
        .Add "EsBasado"
        .Add "EsExpediente"
        .Add "Ordinal"
        .Add "IdGradoClasificacion"
        .Add "IDOrganoContratacion"
        .Add "IDOficinaPrograma"
        .Add "IDEjercito"
        .Add "AccesoSharepoint"
        .Add "Observaciones"
        .Add "FechaCreacion"
        .Add "IDUsuarioCreacion"
        .Add "FechaUltimoCambio"
        .Add "IDUsuarioUltimoCambio"
        .Add "Ambito"
        .Add "NPedido"
        .Add "IDResponsableCalidad"
        .Add "TIPO"
        .Add "AGEDYSAplica"
        .Add "AGEDYSGenerico"
        .Add "HPSAplica"
        .Add "CadenaPecal"
        .Add "Pecal"
        .Add "TipoInforme"
        .Add "POSTAGEDO"
        .Add "APLICAESTADO"
        .Add "FECHAINICIOLICITACION"
        .Add "FECHAOFERTA"
        .Add "FECHAADJUDICACION"
        .Add "FECHAFIRMACONTRATO"
        .Add "GARANTIAMESES"
        .Add "FECHACERTIFICACION"
        .Add "FECHAPERDIDA"
        .Add "FECHADESESTIMADA"
        .Add "ESTADO"
        .Add "CodigoActividad"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDExpediente" Then
        getPropiedad = Me.IDExpediente
        Exit Function
    End If
    If m_NombrePropiedad = "IDExpedientePadre" Then
        getPropiedad = Me.IDExpedientePadre
        Exit Function
    End If
    If m_NombrePropiedad = "Nemotecnico" Then
        getPropiedad = Me.Nemotecnico
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo" Then
        getPropiedad = Me.Titulo
        Exit Function
    End If
    If m_NombrePropiedad = "ImporteLicitacion" Then
        getPropiedad = Me.ImporteLicitacion
        Exit Function
    End If
    If m_NombrePropiedad = "ImporteContratacion" Then
        getPropiedad = Me.ImporteContratacion
        Exit Function
    End If
    
    If m_NombrePropiedad = "CodProyecto" Then
        getPropiedad = Me.CodProyecto
        Exit Function
    End If
    If m_NombrePropiedad = "CodExp" Then
        getPropiedad = Me.CodExp
        Exit Function
    End If
    If m_NombrePropiedad = "CodExpLargo" Then
        getPropiedad = Me.CodExpLargo
        Exit Function
    End If
    If m_NombrePropiedad = "CodS4H" Then
        getPropiedad = Me.CodS4H
        Exit Function
    End If
    
    
    If m_NombrePropiedad = "EsAM" Then
        getPropiedad = Me.EsAM
        Exit Function
    End If
    If m_NombrePropiedad = "EsLote" Then
        getPropiedad = Me.EsLote
        Exit Function
    End If
    If m_NombrePropiedad = "EsBasado" Then
        getPropiedad = Me.EsBasado
        Exit Function
    End If
    If m_NombrePropiedad = "EsExpediente" Then
        getPropiedad = Me.EsExpediente
        Exit Function
    End If
    If m_NombrePropiedad = "Ordinal" Then
        getPropiedad = Me.Ordinal
        Exit Function
    End If
    If m_NombrePropiedad = "IdGradoClasificacion" Then
        getPropiedad = Me.IdGradoClasificacion
        Exit Function
    End If
    If m_NombrePropiedad = "IDOrganoContratacion" Then
        getPropiedad = Me.IDOrganoContratacion
        Exit Function
    End If
    If m_NombrePropiedad = "IDOficinaPrograma" Then
        getPropiedad = Me.IDOficinaPrograma
        Exit Function
    End If
    If m_NombrePropiedad = "IDEjercito" Then
        getPropiedad = Me.IDEjercito
        Exit Function
    End If
    If m_NombrePropiedad = "AccesoSharepoint" Then
        getPropiedad = Me.AccesoSharepoint
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCreacion" Then
        getPropiedad = Me.FechaCreacion
        Exit Function
    End If
    If m_NombrePropiedad = "IDUsuarioCreacion" Then
        getPropiedad = Me.IDUsuarioCreacion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaUltimoCambio" Then
        getPropiedad = Me.FechaUltimoCambio
        Exit Function
    End If
    If m_NombrePropiedad = "IDUsuarioUltimoCambio" Then
        getPropiedad = Me.IDUsuarioUltimoCambio
        Exit Function
    End If
    If m_NombrePropiedad = "Ambito" Then
        getPropiedad = Me.Ambito
        Exit Function
    End If
    
    If m_NombrePropiedad = "NPedido" Then
        getPropiedad = Me.NPedido
        Exit Function
    End If
    If m_NombrePropiedad = "IDResponsableCalidad" Then
        getPropiedad = Me.IDResponsableCalidad
        Exit Function
    End If
    If m_NombrePropiedad = "TIPO" Then
        getPropiedad = Me.Tipo
        Exit Function
    End If
    If m_NombrePropiedad = "AGEDYSAplica" Then
        getPropiedad = Me.AGEDYSAplica
        Exit Function
    End If
    If m_NombrePropiedad = "AGEDYSGenerico" Then
        getPropiedad = Me.AGEDYSGenerico
        Exit Function
    End If
    If m_NombrePropiedad = "HPSAplica" Then
        getPropiedad = Me.HPSAplica
        Exit Function
    End If
    If m_NombrePropiedad = "CadenaPecal" Then
        getPropiedad = Me.CadenaPecal
        Exit Function
    End If
    If m_NombrePropiedad = "Pecal" Then
        getPropiedad = Me.PECAL
        Exit Function
    End If
    If m_NombrePropiedad = "TipoInforme" Then
        getPropiedad = Me.TipoInforme
        Exit Function
    End If
    If m_NombrePropiedad = "POSTAGEDO" Then
        getPropiedad = Me.POSTAGEDO
        Exit Function
    End If
    If m_NombrePropiedad = "APLICAESTADO" Then
        getPropiedad = Me.APLICAESTADO
        Exit Function
    End If
    If m_NombrePropiedad = "POSTAGEDO" Then
        getPropiedad = Me.POSTAGEDO
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAINICIOLICITACION" Then
        getPropiedad = Me.FECHAINICIOLICITACION
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAOFERTA" Then
        getPropiedad = Me.FECHAOFERTA
        Exit Function
    End If
     If m_NombrePropiedad = "FECHADESESTIMADA" Then
        getPropiedad = Me.FECHADESESTIMADA
        Exit Function
    End If
     If m_NombrePropiedad = "FECHAPERDIDA" Then
        getPropiedad = Me.FECHAPERDIDA
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAADJUDICACION" Then
        getPropiedad = Me.FECHAADJUDICACION
        Exit Function
    End If
     If m_NombrePropiedad = "FechaInicioContrato" Then
        getPropiedad = Me.FechaInicioContrato
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinContrato" Then
        getPropiedad = Me.FechaFinContrato
        Exit Function
    End If
    If m_NombrePropiedad = "GARANTIAMESES" Then
        getPropiedad = Me.GARANTIAMESES
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAFIRMACONTRATO" Then
        getPropiedad = Me.FechaFirmaContrato
        Exit Function
    End If
    If m_NombrePropiedad = "FECHACERTIFICACION" Then
        getPropiedad = Me.FECHACERTIFICACION
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinGarantia" Then
        getPropiedad = Me.FechaFinGarantia
        Exit Function
    End If
    If m_NombrePropiedad = "ESTADO" Then
        getPropiedad = Me.Estado
        Exit Function
    End If
    If m_NombrePropiedad = "CodigoActividad" Then
        getPropiedad = Me.CodigoActividad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Expediente.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Expediente.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDExpediente" Then
        Me.IDExpediente = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDExpedientePadre" Then
        Me.IDExpedientePadre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Nemotecnico" Then
        Me.Nemotecnico = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo" Then
        Me.Titulo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ImporteLicitacion" Then
        Me.ImporteLicitacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ImporteContratacion" Then
        Me.ImporteContratacion = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "CodProyecto" Then
        Me.CodProyecto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodExp" Then
        Me.CodExp = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodExpLargo" Then
        Me.CodExpLargo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodS4H" Then
        Me.CodS4H = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "EsAM" Then
        Me.EsAM = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsLote" Then
        Me.EsLote = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsBasado" Then
        Me.EsBasado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsExpediente" Then
        Me.EsExpediente = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Ordinal" Then
        Me.Ordinal = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IdGradoClasificacion" Then
        Me.IdGradoClasificacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDOrganoContratacion" Then
        Me.IDOrganoContratacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDOficinaPrograma" Then
        Me.IDOficinaPrograma = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEjercito" Then
        Me.IDEjercito = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AccesoSharepoint" Then
        Me.AccesoSharepoint = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCreacion" Then
        Me.FechaCreacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDUsuarioCreacion" Then
        Me.IDUsuarioCreacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaUltimoCambio" Then
        Me.FechaUltimoCambio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDUsuarioUltimoCambio" Then
        Me.IDUsuarioUltimoCambio = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "Ambito" Then
        Me.Ambito = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NPedido" Then
        Me.NPedido = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDResponsableCalidad" Then
        Me.IDResponsableCalidad = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "TIPO" Then
        Me.Tipo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AGEDYSAplica" Then
        Me.AGEDYSAplica = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AGEDYSGenerico" Then
        Me.AGEDYSGenerico = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPSAplica" Then
        Me.HPSAplica = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CadenaPecal" Then
        Me.CadenaPecal = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Pecal" Then
        Me.PECAL = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TipoInforme" Then
        Me.TipoInforme = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "POSTAGEDO" Then
        Me.POSTAGEDO = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "APLICAESTADO" Then
        Me.APLICAESTADO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAINICIOLICITACION" Then
        Me.FECHAINICIOLICITACION = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAOFERTA" Then
        Me.FECHAOFERTA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHADESESTIMADA" Then
        Me.FECHADESESTIMADA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAPERDIDA" Then
        Me.FECHAPERDIDA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAADJUDICACION" Then
        Me.FECHAADJUDICACION = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicioContrato" Then
        Me.FechaInicioContrato = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinContrato" Then
        Me.FechaFinContrato = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "GARANTIAMESES" Then
        Me.GARANTIAMESES = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAFIRMACONTRATO" Then
        Me.FechaFirmaContrato = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHACERTIFICACION" Then
        Me.FECHACERTIFICACION = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinGarantia" Then
        Me.FechaFinGarantia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ESTADO" Then
        Me.Estado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodigoActividad" Then
        Me.CodigoActividad = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Expediente.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Expediente.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function



