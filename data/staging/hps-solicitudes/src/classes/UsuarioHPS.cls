VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioHPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public id As String
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
Public FAvisoConcesion As String
Public RequiereComunicacionConcesion As String
Public FechaHPSConcesionMinima As String



Private m_objEmpresaUsuario As Suministrador
Private m_ObjExpediente As Expediente
Private m_sNombreCompleto As String

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get EmpresaUsuario() As Suministrador

    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objEmpresaUsuario Is Nothing Then
        Set EmpresaUsuario = m_objEmpresaUsuario
        Exit Property
    End If
    If Me.IDEmpresaUsuario = "" Then
        Exit Property
    End If
    Set m_objEmpresaUsuario = constructor.getSuministrador(Me.IDEmpresaUsuario, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EmpresaUsuario = m_objEmpresaUsuario
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.EmpresaUsuario ha devuelto el error: " & vbNewLine & Err.Description
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
    Set m_ObjExpediente = constructor.getExpediente(Me.IDExpediente, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Expediente = m_ObjExpediente
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioHPS.expediente ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set Expediente(ByVal objNewValue As Expediente)

    Set m_ObjExpediente = objNewValue

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
        m_Error = "El método UsuarioHPS.NombreCompleto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
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
        .Add "FAvisoConcesion"
        .Add "RequiereComunicacionConcesion"
        .Add "FechaHPSConcesionMinima"
        
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método UsuarioHPS.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "ID" Then
        getPropiedad = Me.id
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
    If m_NombrePropiedad = "LugarNacimiento" Then
        getPropiedad = Me.LugarNacimiento
        Exit Function
    End If
    If m_NombrePropiedad = "F_Nacimiento" Then
        getPropiedad = Me.F_Nacimiento
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
    If m_NombrePropiedad = "IDEmpresaHPS" Then
        getPropiedad = Me.IDEmpresaHPS
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
    
    p_Error = "Propiedad no reconocida en UsuarioHPS.getPropiedad"
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHPS.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function setPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "ID" Then
        Me.id = m_ValorPropiedad
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
    
    p_Error = "Propiedad no reconocida en UsuarioHPS.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHPS.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function






