VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public ID As String
Public DNI As String
Public Nombre As String
Public Apellido_1 As String
Public Apellido_2 As String
Public Telefono As String
Public Correo_e As String
Public F_Nacimiento As String
Public LugarNacimiento As String
Public Motivo_HPS As String
Public IDExpediente As String
Public F_Curso As String
Public CursoEnVigor As String
Public Requiere_Curso As String
Public FechaPrimeraConvocatoria As String
Public FechaSegundaConvocatoria As String
Public FechaCorreoNoCurso As String
Public F_Baja As String
Public LugarPrestacionServicio As String
Public IDEmpresaUsuario As String
Public IDEmpresaHPS As String
Public IDJuridicaContrato As String
Public FAvisoConcesion As String
Public RequiereComunicacionConcesion As String
Public FechaHPSConcesionMinima As String
Public CadenaContratistas As String


Private m_objSuministradorUsuarioObj As Suministrador
Private m_objSuministradorTramitadoraObj As Suministrador
Private m_objSuministradorContratoObj As Suministrador
Private m_ObjExpediente As Expediente

Private m_ObjColObservacionesHistoricas As Scripting.Dictionary
Private m_objColObservacionesHistoricasCurso As Scripting.Dictionary
Private m_ObjColAnexos As Scripting.Dictionary
Private m_sURLCarpetaAnexos As String

Private m_sNombreCompleto As String

Private m_objUsuarioSICA As UsuarioSICA
Private m_objColCamposSimplificado As Collection
Private m_sCadenaContratistasCalculados As String


Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get CadenaContratistasCalculados() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.Expediente Is Nothing Then
        Exit Property
    End If
    
    m_sCadenaContratistasCalculados = Me.Expediente.CadenaContratistas
    m_Error = Me.Expediente.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    CadenaContratistasCalculados = m_sCadenaContratistasCalculados
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.CadenaContratistasCalculados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Get Expediente() As Expediente
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjExpediente Is Nothing Then
        Set Expediente = m_ObjExpediente
        Exit Property
    End If
    If Me.IDExpediente = "" Then
        
        Exit Property
    End If
    Set m_ObjExpediente = Constructor.getExpediente(Me.IDExpediente, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Expediente = m_ObjExpediente
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHistorico.expediente ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get SuministradorUsuarioObj() As Suministrador

    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objSuministradorUsuarioObj Is Nothing Then
        Set SuministradorUsuarioObj = m_objSuministradorUsuarioObj
        Exit Property
    End If
    If Me.IDEmpresaUsuario = "" Then
        Exit Property
    End If
    Set m_objSuministradorUsuarioObj = Constructor.getSuministrador(Me.IDEmpresaUsuario, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set SuministradorUsuarioObj = m_objSuministradorUsuarioObj
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHistorico.SuministradorUsuarioObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get SuministradorTramitadoraObj() As Suministrador

    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objSuministradorTramitadoraObj Is Nothing Then
        Set SuministradorTramitadoraObj = m_objSuministradorTramitadoraObj
        Exit Property
    End If
    If Me.IDEmpresaHPS = "" Then
        Exit Property
    End If
    Set m_objSuministradorTramitadoraObj = Constructor.getSuministrador(Me.IDEmpresaHPS, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set SuministradorTramitadoraObj = m_objSuministradorTramitadoraObj
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHistorico.SuministradorTramitadoraObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property
Public Property Get SuministradorContratoObj() As Suministrador

    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objSuministradorContratoObj Is Nothing Then
        Set SuministradorContratoObj = m_objSuministradorContratoObj
        Exit Property
    End If
    If Me.IDJuridicaContrato = "" Then
        Exit Property
    End If
    Set m_objSuministradorContratoObj = Constructor.getSuministrador(Me.IDJuridicaContrato, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set SuministradorContratoObj = m_objSuministradorContratoObj
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHistorico.SuministradorContratoObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get UsuarioSICA() As UsuarioSICA
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
     
    If Not m_objUsuarioSICA Is Nothing Then
       Set UsuarioSICA = m_objUsuarioSICA
       Exit Property
    End If
    If Me.ID = "" Then
        Exit Property
    End If
    Set m_objUsuarioSICA = Constructor.getUsuarioSICA(, , Me.ID, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioSICA = m_objUsuarioSICA
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHistorico.UsuarioSICA ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set UsuarioSICA(ByVal objNewValue As UsuarioSICA)

    Set m_objUsuarioSICA = objNewValue

End Property




Public Property Get ColAnexos() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColAnexos Is Nothing Then
        Set ColAnexos = m_ObjColAnexos
        Exit Property
    End If
    If Me.ID = "" Then
        Exit Property
    End If
    
    Set m_ObjColAnexos = Constructor.getAnexosUsuarioHistorico(Me.ID, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAnexos = m_ObjColAnexos
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHistorico.colAnexos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set ColAnexos(ByVal objNewValue As Variant)

    Set m_ObjColAnexos = objNewValue

End Property

Public Property Get URLCarpetaAnexos() As String
    
    Dim m_URLCarpetaAnexosGlobales As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLCarpetaAnexos <> "" Then
        URLCarpetaAnexos = m_sURLCarpetaAnexos
        Exit Property
    End If
    If Me.DNI = "" Then
        Exit Property
    End If
    m_URLCarpetaAnexosGlobales = m_ObjEntorno.URLCarpetaAnexosHistoricos
    If m_URLCarpetaAnexosGlobales = "" Then
        Exit Property
    End If
    If Right(m_URLCarpetaAnexosGlobales, 1) <> "\" Then
        m_URLCarpetaAnexosGlobales = m_URLCarpetaAnexosGlobales & "\"
    End If
    m_sURLCarpetaAnexos = m_URLCarpetaAnexosGlobales & Me.DNI & "\"
    URLCarpetaAnexos = m_sURLCarpetaAnexos
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHistorico.URLCarpetaAnexos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get NombreCompleto() As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sNombreCompleto <> "" Then
        NombreCompleto = m_sNombreCompleto
        Exit Property
    End If
    If Me.Nombre = "" And Me.Apellido_1 = "" And Me.Apellido_2 = "" Then
        Exit Property
    End If
    m_sNombreCompleto = Me.Nombre & " " & Me.Apellido_1 & " " & Me.Apellido_2
    NombreCompleto = m_sNombreCompleto
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHistorico.NombreCompleto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get ColObservacionesHistoricas() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColObservacionesHistoricas Is Nothing Then
        Set ColObservacionesHistoricas = m_ObjColObservacionesHistoricas
        Exit Property
    End If
    If Me.ID = "" Then
        Exit Property
    End If
    
    Set m_ObjColObservacionesHistoricas = Constructor.getObservacionesHistoricas(Me.ID, "General", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColObservacionesHistoricas = m_ObjColObservacionesHistoricas
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHistorico.ColObservacionesHistoricas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set ColObservacionesHistoricas(ByVal objNewValue As Variant)
    
    Set m_ObjColObservacionesHistoricas = objNewValue

End Property
Public Property Get ColObservacionesHistoricasCurso() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColObservacionesHistoricasCurso Is Nothing Then
        Set ColObservacionesHistoricasCurso = m_objColObservacionesHistoricasCurso
        Exit Property
    End If
    If Me.ID = "" Then
        Exit Property
    End If
    
    Set m_objColObservacionesHistoricasCurso = Constructor.getObservacionesHistoricas(Me.ID, "Curso", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColObservacionesHistoricasCurso = m_objColObservacionesHistoricasCurso
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHistorico.ColObservacionesHistoricasCurso ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set ColObservacionesHistoricasCurso(ByVal objNewValue As Variant)
    
    Set m_objColObservacionesHistoricasCurso = objNewValue

End Property

Public Function PasarActual( _
                                Optional ByRef p_Error As String _
                                ) As String
                                    
    
    
    Dim m_ObjColObservaciones As Scripting.Dictionary
    Dim m_ObjUsuario As UsuarioHPS
    Dim m_ID As Variant
    Dim m_ObjObServacion As Observacion
    Dim m_ObjObServacionHistorica As ObservacionHistorica
    Dim m_UsuarioSICA As UsuarioSICA
    
    On Error GoTo errores
    
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        p_Error = "El Sistema está en modo sólo lectura"
        Err.Raise 1000
    End If
    
    Set m_ObjUsuario = Constructor.getUsuarioHPS(m_ObjUsuarioHistoricoActivo.ID, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_UsuarioSICA = Me.UsuarioSICA
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjUsuario Is Nothing Then
        
        m_ObjUsuario.Borrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Set m_ObjUsuario = New UsuarioHPS
    With m_ObjUsuario
        .ID = m_ObjUsuarioHistoricoActivo.ID
        .DNI = m_ObjUsuarioHistoricoActivo.DNI
        .Nombre = m_ObjUsuarioHistoricoActivo.Nombre
        .Apellido_1 = m_ObjUsuarioHistoricoActivo.Apellido_1
        .Apellido_2 = m_ObjUsuarioHistoricoActivo.Apellido_2
        .Telefono = m_ObjUsuarioHistoricoActivo.Telefono
        .Correo_e = m_ObjUsuarioHistoricoActivo.Correo_e
        .F_Nacimiento = m_ObjUsuarioHistoricoActivo.F_Nacimiento
        .LugarNacimiento = m_ObjUsuarioHistoricoActivo.LugarNacimiento
        .Motivo_HPS = m_ObjUsuarioHistoricoActivo.Motivo_HPS
        .IDExpediente = m_ObjUsuarioHistoricoActivo.IDExpediente
        .F_Curso = m_ObjUsuarioHistoricoActivo.F_Curso
        .CursoEnVigor = m_ObjUsuarioHistoricoActivo.CursoEnVigor
        .Requiere_Curso = m_ObjUsuarioHistoricoActivo.Requiere_Curso
        .FechaPrimeraConvocatoria = m_ObjUsuarioHistoricoActivo.FechaPrimeraConvocatoria
        .FechaSegundaConvocatoria = m_ObjUsuarioHistoricoActivo.FechaSegundaConvocatoria
        .FechaCorreoNoCurso = m_ObjUsuarioHistoricoActivo.FechaCorreoNoCurso
        .LugarPrestacionServicio = m_ObjUsuarioHistoricoActivo.LugarPrestacionServicio
        .IDEmpresaUsuario = m_ObjUsuarioHistoricoActivo.IDEmpresaUsuario
        .IDEmpresaHPS = m_ObjUsuarioHistoricoActivo.IDEmpresaHPS
        .IDJuridicaContrato = m_ObjUsuarioHistoricoActivo.IDJuridicaContrato
        .FAvisoConcesion = m_ObjUsuarioHistoricoActivo.FAvisoConcesion
        .RequiereComunicacionConcesion = m_ObjUsuarioHistoricoActivo.RequiereComunicacionConcesion
        .FechaHPSConcesionMinima = m_ObjUsuarioHistoricoActivo.FechaHPSConcesionMinima
        .CadenaContratistas = m_ObjUsuarioHistoricoActivo.CadenaContratistas
    End With
    m_ObjUsuario.ParaAlta = EnumSiNo.Sí
    m_ObjUsuario.ID = m_ObjUsuarioHistoricoActivo.ID
    m_ObjUsuario.Registrar , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set m_ObjUsuario = Constructor.getUsuarioHPS(m_ObjUsuario.ID, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjUsuario Is Nothing Then
        p_Error = "No se ha podido grabar el usuario en la tabla de Actuales"
        Err.Raise 1000
    End If
    Set m_ObjColObservaciones = m_ObjUsuarioHistoricoActivo.ColObservacionesHistoricas
    If Not m_ObjColObservaciones Is Nothing Then
        For Each m_ID In m_ObjColObservaciones
            Set m_ObjObServacionHistorica = m_ObjColObservaciones(m_ID)
            Set m_ObjObServacion = New Observacion
            With m_ObjObServacion
                .ID = m_ObjUsuarioHistoricoActivo.ID
                .Fecha = m_ObjObServacionHistorica.Fecha
                .IDObservacion = m_ObjObServacionHistorica.IDObservacion
                .Observacion = m_ObjObServacionHistorica.Observacion
                .TIPO = m_ObjObServacionHistorica.TIPO
                .ParaAlta = EnumSiNo.Sí
                .Registrar p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End With
            Set m_ObjObServacionHistorica = Nothing
            Set m_ObjObServacion = Nothing
        Next
    End If
    Set m_ObjColObservaciones = m_ObjUsuarioHistoricoActivo.ColObservacionesHistoricasCurso
    If Not m_ObjColObservaciones Is Nothing Then
        For Each m_ID In m_ObjColObservaciones
            Set m_ObjObServacionHistorica = m_ObjColObservaciones(m_ID)
            Set m_ObjObServacion = New Observacion
            
            With m_ObjObServacion
                .ID = m_ObjUsuarioHistoricoActivo.ID
                .Fecha = m_ObjObServacionHistorica.Fecha
                .IDObservacion = m_ObjObServacionHistorica.IDObservacion
                .Observacion = m_ObjObServacionHistorica.Observacion
                .TIPO = m_ObjObServacionHistorica.TIPO
                .ParaAlta = EnumSiNo.Sí
                .Registrar p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End With
            Set m_ObjObServacionHistorica = Nothing
            Set m_ObjObServacion = Nothing
        Next
    End If
    
    PasarAnexosAActual p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_UsuarioSICA Is Nothing Then
        m_UsuarioSICA.PasarAActual m_ObjUsuario.ID, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Borrar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHistorico.PasarActual ha devuelto el error: " & Err.Description
    End If
End Function

Private Function PasarAnexosAHistoricos( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                            
    
    
End Function

Private Function PasarAnexosAActual( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                            
    
    
    Dim m_ObjColAnexosHistoricos As Scripting.Dictionary
    Dim m_ObjAnexoUsuarioHistorico As AnexoUsuarioHistorico
    Dim m_ObjAnexoUsuarioHPS As AnexoUsuarioHPS
    
    Dim m_UsuarioHPS As UsuarioHPS
    
    Dim m_ID As Variant
    Dim m_URLDirActual As String
    Dim m_URLDirHistorico As String
    Dim m_URLEnActual As String
    On Error GoTo errores
    
    
    Set m_ObjColAnexosHistoricos = Me.ColAnexos
    If m_ObjColAnexosHistoricos Is Nothing Then
        Exit Function
    End If
    Set m_UsuarioHPS = getUsuarioHPS(Me.ID, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_UsuarioHPS Is Nothing Then
        p_Error = "No se pueden trasladar los anexos Históricos de usuario Histórico si no se ha creado antes el UsuarioHPS"
        Err.Raise 1000
    End If
    m_URLDirActual = m_UsuarioHPS.URLCarpetaAnexosActuales
    m_URLDirHistorico = Me.URLCarpetaAnexos
    
    For Each m_ID In m_ObjColAnexosHistoricos
        If m_ID = Empty Then GoTo siguiente
        Set m_ObjAnexoUsuarioHistorico = m_ObjColAnexosHistoricos(m_ID)
        If fso.FileExists(m_ObjAnexoUsuarioHistorico.URLAnexo) Then
            If FicheroAbierto(m_ObjAnexoUsuarioHistorico.URLAnexo) Then
                p_Error = "Está abierto el anexo: " & vbNewLine & m_ObjAnexoUsuarioHistorico.URLAnexo
                Err.Raise 1000
            End If
        End If
        m_URLEnActual = m_URLEnActual & fso.GetFileName(m_ObjAnexoUsuarioHistorico.URLAnexo)
        If fso.FileExists(m_URLEnActual) Then
            If FicheroAbierto(m_URLEnActual) Then
                p_Error = "Está abierto el anexo: " & vbNewLine & m_URLEnActual
                Err.Raise 1000
            End If
        End If
        Set m_ObjAnexoUsuarioHistorico = Nothing
siguiente:
    Next
    For Each m_ID In m_ObjColAnexosHistoricos
        If m_ID = Empty Then GoTo siguiente
        Set m_ObjAnexoUsuarioHistorico = m_ObjColAnexosHistoricos(m_ID)
        Set m_ObjAnexoUsuarioHPS = New AnexoUsuarioHPS
        With m_ObjAnexoUsuarioHPS
            .IDAnexo = m_ObjAnexoUsuarioHistorico.IDAnexo
            .idUsuario = m_UsuarioHPS.ID
            .NombreAnexo = m_ObjAnexoUsuarioHistorico.NombreAnexo
            .EsHistorico = "No"
            .URLLocal = m_ObjAnexoUsuarioHistorico.URLAnexo
            .Registrar p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
        End With
        m_ObjAnexoUsuarioHistorico.Borrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjAnexoUsuarioHistorico = Nothing
        Set m_ObjAnexoUsuarioHPS = Nothing
siguiente1:
    Next
    Set Me.ColAnexos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHistorico.PasarAnexosAActual ha devuelto el error: " & Err.Description
    End If
    
    
End Function
Public Function Registrar( _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        p_Error = "El Sistema está en modo sólo lectura"
        Err.Rais
    End If
    
    m_SQL = "TbUsuariosHistoricos"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("ID") = Me.ID
            .Fields("DNI") = Me.DNI
            .Fields("Nombre") = Me.Nombre
            .Fields("Apellido_1") = Me.Apellido_1
            If Me.Apellido_2 <> "" Then
                .Fields("Apellido_2") = Me.Apellido_2
            End If
            .Fields("Telefono") = Me.Telefono
            .Fields("Correo_e") = Me.Correo_e
            .Fields("F_Nacimiento") = Me.F_Nacimiento
            If Me.LugarNacimiento <> "" Then
                .Fields("LugarNacimiento") = Me.LugarNacimiento
            End If
            .Fields("Motivo_HPS") = Me.Motivo_HPS
            .Fields("IDExpediente") = Me.IDExpediente
            If IsDate(Me.F_Curso) Then
                .Fields("F_Curso") = Me.F_Curso
            End If
            If Me.CursoEnVigor = "Sí" Or Me.CursoEnVigor = "No" Then
                .Fields("CursoEnVigor") = Me.CursoEnVigor
            End If
            
            If Me.Requiere_Curso = "Sí" Or Me.Requiere_Curso = "No" Then
                .Fields("Requiere_Curso") = Me.Requiere_Curso
            End If
            If IsDate(Me.FechaPrimeraConvocatoria) Then
                .Fields("FechaPrimeraConvocatoria") = Me.FechaPrimeraConvocatoria
            End If
            If IsDate(Me.FechaSegundaConvocatoria) Then
                .Fields("FechaSegundaConvocatoria") = Me.FechaSegundaConvocatoria
            End If
            If IsDate(Me.FechaCorreoNoCurso) Then
                .Fields("FechaCorreoNoCurso") = Me.FechaCorreoNoCurso
            End If
            If IsDate(Me.F_Baja) Then
                .Fields("F_Baja") = Me.F_Baja
            End If
            If Me.LugarPrestacionServicio <> "" Then
                .Fields("LugarPrestacionServicio") = Me.LugarPrestacionServicio
            End If
            If Me.IDEmpresaUsuario <> "" Then
                .Fields("IDEmpresaUsuario") = Me.IDEmpresaUsuario
            End If
            If Me.IDEmpresaHPS <> "" Then
                .Fields("IDEmpresaHPS") = Me.IDEmpresaHPS
            End If
            If Me.IDJuridicaContrato <> "" Then
                .Fields("IDJuridicaContrato") = Me.IDJuridicaContrato
            End If
            If IsDate(Me.FAvisoConcesion) Then
                .Fields("FAvisoConcesion") = Me.FAvisoConcesion
            End If
            
            If Me.RequiereComunicacionConcesion = "Sí" Or Me.RequiereComunicacionConcesion = "No" Then
                .Fields("RequiereComunicacionConcesion") = Me.RequiereComunicacionConcesion
            End If
            If IsDate(Me.FechaHPSConcesionMinima) Then
                .Fields("FechaHPSConcesionMinima") = Me.FechaHPSConcesionMinima
            End If
            If Nz(Me.CadenaContratistas, "") <> "" Then
                .Fields("CadenaContratistas") = Me.CadenaContratistas
            End If
        .Update
    
    
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    cache_usuario_historico_regenerar p_ID:=Me.ID, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHistorico.registrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
    Dim m_SQL As String
    
    On Error GoTo errores
    
    m_SQL = "DELETE * FROM TbUsuariosHistoricos " & _
            "WHERE ID=" & Me.ID & ";"
    getdb().Execute m_SQL
    cache_usuario_historico_regenerar p_ID:=Me.ID, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHistorico.Borrar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function CadenaContratistasGrabar( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                        
    
    Dim m_SQL As String
   
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    Me.CadenaContratistas = Me.CadenaContratistasCalculados
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.CadenaContratistas <> "" Then
        m_SQL = "UPDATE TbUsuariosHistoricos SET CadenaContratistas = '" & Me.CadenaContratistas & "' " & _
                "WHERE ID=" & Me.ID & ";"
    Else
        m_SQL = "UPDATE TbUsuariosHistoricos SET CadenaContratistas = Null " & _
                "WHERE ID=" & Me.ID & ";"
    End If
    getdb().Execute m_SQL
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHistorico.CadenaContratistasGrabar ha devuelto el error: " & Err.Description
    End If
End Function
Public Property Get ColCamposSimplificado() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColCamposSimplificado Is Nothing Then
        Set ColCamposSimplificado = m_objColCamposSimplificado
        Exit Property
    End If
    Set m_objColCamposSimplificado = New Collection
    With m_objColCamposSimplificado
        .Add "ID"
        .Add "DNI"
        .Add "Nombre"
        .Add "Apellido_1"
        .Add "Apellido_2"
        .Add "Telefono"
        .Add "Correo_e"
        .Add "F_Nacimiento"
        .Add "LugarNacimiento"
        .Add "Motivo_HPS"
        .Add "IDExpediente"
        .Add "F_Curso"
        .Add "CursoEnVigor"
        .Add "Requiere_Curso"
        .Add "FechaPrimeraConvocatoria"
        .Add "FechaSegundaConvocatoria"
        .Add "FechaCorreoNoCurso"
        .Add "F_Baja"
        .Add "LugarPrestacionServicio"
        .Add "IDEmpresaUsuario"
        .Add "IDEmpresaHPS"
        .Add "IDJuridicaContrato"
        .Add "FAvisoConcesion"
        .Add "RequiereComunicacionConcesion"
        .Add "FechaHPSConcesionMinima"
        .Add "CadenaContratistas"
    End With
    Set ColCamposSimplificado = m_objColCamposSimplificado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método UsuarioHistorico.ColCamposSimplificado ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
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
        .Add "DNI"
        .Add "Nombre"
        .Add "Apellido_1"
        .Add "Apellido_2"
        .Add "Telefono"
        .Add "Correo_e"
        .Add "F_Nacimiento"
        .Add "LugarNacimiento"
        .Add "Motivo_HPS"
        .Add "IDExpediente"
        .Add "F_Curso"
        .Add "CursoEnVigor"
        .Add "Requiere_Curso"
        .Add "FechaPrimeraConvocatoria"
        .Add "FechaSegundaConvocatoria"
        .Add "FechaCorreoNoCurso"
        .Add "F_Baja"
        .Add "LugarPrestacionServicio"
        .Add "IDEmpresaUsuario"
        .Add "IDEmpresaHPS"
        .Add "IDJuridicaContrato"
        .Add "FAvisoConcesion"
        .Add "RequiereComunicacionConcesion"
        .Add "FechaHPSConcesionMinima"
        .Add "IDEmpresaUsuario"
        .Add "IDEmpresaHPS"
        .Add "IDJuridicaContrato"
        .Add "CadenaContratistas"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método UsuarioHistorico.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "ID" Then
        getPropiedad = Me.ID
        Exit Function
    End If
    If m_NombrePropiedad = "DNI" Then
        getPropiedad = Me.DNI
        Exit Function
    End If
    If m_NombrePropiedad = "Nombre" Then
        getPropiedad = Me.Nombre
        Exit Function
    End If
    If m_NombrePropiedad = "Apellido_1" Then
        getPropiedad = Me.Apellido_1
        Exit Function
    End If
    If m_NombrePropiedad = "Apellido_2" Then
        getPropiedad = Me.Apellido_2
        Exit Function
    End If
    If m_NombrePropiedad = "Telefono" Then
        getPropiedad = Me.Telefono
        Exit Function
    End If
    If m_NombrePropiedad = "Correo_e" Then
        getPropiedad = Me.Correo_e
        Exit Function
    End If
    If m_NombrePropiedad = "F_Nacimiento" Then
        getPropiedad = Me.F_Nacimiento
        Exit Function
    End If
    If m_NombrePropiedad = "LugarNacimiento" Then
        getPropiedad = Me.LugarNacimiento
        Exit Function
    End If
    If m_NombrePropiedad = "Motivo_HPS" Then
        getPropiedad = Me.Motivo_HPS
        Exit Function
    End If
    If m_NombrePropiedad = "IDExpediente" Then
        getPropiedad = Me.IDExpediente
        Exit Function
    End If
    If m_NombrePropiedad = "F_Curso" Then
        getPropiedad = Me.F_Curso
        Exit Function
    End If
    If m_NombrePropiedad = "CursoEnVigor" Then
        getPropiedad = Me.CursoEnVigor
        Exit Function
    End If
    If m_NombrePropiedad = "Requiere_Curso" Then
        getPropiedad = Me.Requiere_Curso
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrimeraConvocatoria" Then
        getPropiedad = Me.FechaPrimeraConvocatoria
        Exit Function
    End If
    If m_NombrePropiedad = "FechaSegundaConvocatoria" Then
        getPropiedad = Me.FechaSegundaConvocatoria
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCorreoNoCurso" Then
        getPropiedad = Me.FechaCorreoNoCurso
        Exit Function
    End If
    
    If m_NombrePropiedad = "F_Baja" Then
        getPropiedad = Me.F_Baja
        Exit Function
    End If
    If m_NombrePropiedad = "LugarPrestacionServicio" Then
        getPropiedad = Me.LugarPrestacionServicio
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaUsuario" Then
        getPropiedad = Me.IDEmpresaUsuario
        Exit Function
    End If
    If m_NombrePropiedad = "IDJuridicaContrato" Then
        getPropiedad = Me.IDJuridicaContrato
        Exit Function
    End If
    If m_NombrePropiedad = "IDJuridicaContrato" Then
        getPropiedad = Me.IDJuridicaContrato
        Exit Function
    End If
    If m_NombrePropiedad = "FAvisoConcesion" Then
        getPropiedad = Me.FAvisoConcesion
        Exit Function
    End If
    If m_NombrePropiedad = "RequiereComunicacionConcesion" Then
        getPropiedad = Me.RequiereComunicacionConcesion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaHPSConcesionMinima" Then
        getPropiedad = Me.FechaHPSConcesionMinima
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaUsuario" Then
        getPropiedad = Me.IDEmpresaUsuario
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaHPS" Then
        getPropiedad = Me.IDEmpresaHPS
        Exit Function
    End If
    If m_NombrePropiedad = "IDJuridicaContrato" Then
        getPropiedad = Me.IDJuridicaContrato
        Exit Function
    End If
    If m_NombrePropiedad = "CadenaContratistas" Then
        getPropiedad = Me.CadenaContratistas
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en UsuarioHistorico.getPropiedad"
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHistorico.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "DNI" Then
        Me.DNI = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Nombre" Then
        Me.Nombre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Apellido_1" Then
        Me.Apellido_1 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Apellido_2" Then
        Me.Apellido_2 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Telefono" Then
        Me.Telefono = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Correo_e" Then
        Me.Correo_e = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "F_Nacimiento" Then
        Me.F_Nacimiento = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "LugarNacimiento" Then
        Me.LugarNacimiento = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Motivo_HPS" Then
        Me.Motivo_HPS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDExpediente" Then
        Me.IDExpediente = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "F_Curso" Then
        Me.F_Curso = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CursoEnVigor" Then
        Me.CursoEnVigor = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Requiere_Curso" Then
        Me.Requiere_Curso = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrimeraConvocatoria" Then
        Me.FechaPrimeraConvocatoria = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaSegundaConvocatoria" Then
        Me.FechaSegundaConvocatoria = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCorreoNoCurso" Then
        Me.FechaCorreoNoCurso = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "F_Baja" Then
        Me.F_Baja = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "LugarPrestacionServicio" Then
        Me.LugarPrestacionServicio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaUsuario" Then
        Me.IDEmpresaUsuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaHPS" Then
        Me.IDEmpresaHPS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDJuridicaContrato" Then
        Me.IDJuridicaContrato = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FAvisoConcesion" Then
        Me.FAvisoConcesion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RequiereComunicacionConcesion" Then
        Me.RequiereComunicacionConcesion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaHPSConcesionMinima" Then
        Me.FechaHPSConcesionMinima = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaUsuario" Then
        Me.IDEmpresaUsuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaHPS" Then
        Me.IDEmpresaHPS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDJuridicaContrato" Then
        Me.IDJuridicaContrato = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CadenaContratistas" Then
        Me.CadenaContratistas = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en UsuarioHistorico.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHistorico.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function







