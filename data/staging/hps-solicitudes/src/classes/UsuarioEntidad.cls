VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioEntidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public id As String
Public DNI As String
Public Nombre As String
Public Apellido1 As String
Public Apellido_2 As String
Public Telefono As String
Public EmpresaUsuario As String
Public EmpresaTramitadora As String
Public JuridicaContrato As String
Public Correo_e As String
Public F_Nacimiento As String
Public LugarNacimiento As String
Public Motivo_HPS As String
Public IDExpediente As String
Public CodExp As String
Public F_Curso As String
Public CursoEnVigor As String
Public F_Baja As String
Public FAvisoConcesion As String
Public RequiereComunicacionConcesion As String
Public FechaHPSConcesionMinima As String
Public Curso_Realizado As String
Public Requiere_Curso As String
Public FechaPrimeraConvocatoria As String
Public FechaSegundaConvocatoria As String
Public FechaCorreoNoCurso As String
Public Requiere_PrimeraConvocatoriaCurso As String
Public Requiere_SegundaConvocatoriaCurso As String
Public Requiere_CorreoJefeSeguridadCurso As String
Public Observaciones As String
Public IDSolicitud As String
Public HPS_NAC_SIN_DATOS As String
Public HPS_NAC_F_Concesion As String
Public HPS_NAC_F_Caducidad As String
Public HPS_NAC_Grado As String
Public HPS_NAC_F_Solicitud As String
Public HPS_NAC_Especialidad As String
Public HPS_NAC_Renovacion As String
Public HPS_NAC_F_Baja As String
Public HPS_NAC_Activo As String
Public HPS_NAC_ApuntoDeCaducar As String
Public HPS_NAC_Baja As String
Public HPS_NAC_Caducado As String
Public HPS_NAC_PendienteRenovacion As String
Public HPS_NAC_Solicitado As String
Public HPS_NAC_ESTADO As String
Public HPS_NAC_Irregular As String
Public HPS_NAC_MOTIVO_IRREGULAR As String
Public HPS_NAC_MESES_PARA_RENOVAR As String
Public HPS_OTAN_SIN_DATOS As String
Public HPS_OTAN_F_Concesion As String
Public HPS_OTAN_F_Caducidad As String
Public HPS_OTAN_Grado As String
Public HPS_OTAN_F_Solicitud As String
Public HPS_OTAN_ESPECIALIDAD As String
Public HPS_OTAN_Renovacion As String
Public HPS_OTAN_F_Baja As String
Public HPS_OTAN_Activo As String
Public HPS_OTAN_ApuntoDeCaducar As String
Public HPS_OTAN_Baja As String
Public HPS_OTAN_Caducado As String
Public HPS_OTAN_PendienteRenovacion As String
Public HPS_OTAN_Solicitado As String
Public HPS_OTAN_ESTADO As String
Public HPS_OTAN_MOTIVO_IRREGULAR As String
Public HPS_OTAN_Irregular As String
Public HPS_OTAN_MESES_PARA_RENOVAR As String
Public HPS_ESA_SIN_DATOS As String
Public HPS_ESA_F_Concesion As String
Public HPS_ESA_F_Caducidad As String
Public HPS_ESA_Grado As String
Public HPS_ESA_F_Solicitud As String
Public HPS_ESA_ESPECIALIDAD As String
Public HPS_ESA_Renovacion As String
Public HPS_ESA_F_Baja As String
Public HPS_ESA_Activo As String
Public HPS_ESA_ApuntoDeCaducar As String
Public HPS_ESA_Baja As String
Public HPS_ESA_Caducado As String
Public HPS_ESA_PendienteRenovacion As String
Public HPS_ESA_Solicitado As String
Public HPS_ESA_ESTADO As String
Public HPS_ESA_Irregular As String
Public HPS_ESA_MOTIVO_IRREGULAR As String
Public HPS_ESA_MESES_PARA_RENOVAR As String
Public HPS_UE_SIN_DATOS As String
Public HPS_UE_F_Concesion As String
Public HPS_UE_F_Caducidad As String
Public HPS_UE_Grado As String
Public HPS_UE_F_Solicitud As String
Public HPS_UE_ESPECIALIDAD As String
Public HPS_UE_Renovacion As String
Public HPS_UE_F_Baja As String
Public HPS_UE_Activo As String
Public HPS_UE_ApuntoDeCaducar As String
Public HPS_UE_Baja As String
Public HPS_UE_Caducado As String
Public HPS_UE_PendienteRenovacion As String
Public HPS_UE_Solicitado As String
Public HPS_UE_ESTADO As String
Public HPS_UE_Irregular As String
Public HPS_UE_MOTIVO_IRREGULAR As String
Public HPS_UE_MESES_PARA_RENOVAR As String
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

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
        .Add "Apellido1"
        .Add "Apellido_2"
        .Add "Telefono"
        .Add "EmpresaUsuario"
        .Add "EmpresaTramitadora"
        .Add "JuridicaContrato"
        .Add "Correo_e"
        .Add "F_Nacimiento"
        .Add "LugarNacimiento"
        .Add "Motivo_HPS"
        .Add "IDExpediente"
        .Add "CodExp"
        .Add "F_Curso"
        .Add "CursoEnVigor"
        .Add "F_Baja"
        .Add "FAvisoConcesion"
        .Add "RequiereComunicacionConcesion"
        .Add "FechaHPSConcesionMinima"
        .Add "Curso_Realizado"
        .Add "Requiere_Curso"
        .Add "FechaPrimeraConvocatoria"
        .Add "FechaSegundaConvocatoria"
        .Add "FechaCorreoNoCurso"
        .Add "Requiere_PrimeraConvocatoriaCurso"
        .Add "Requiere_SegundaConvocatoriaCurso"
        .Add "Requiere_CorreoJefeSeguridadCurso"
        .Add "Observaciones"
        .Add "IDSolicitud"
        .Add "HPS_NAC_SIN_DATOS"
        .Add "HPS_NAC_F_Concesion"
        .Add "HPS_NAC_F_Caducidad"
        .Add "HPS_NAC_Grado"
        .Add "HPS_NAC_F_Solicitud"
        .Add "HPS_NAC_Especialidad"
        .Add "HPS_NAC_Renovacion"
        .Add "HPS_NAC_F_Baja"
        .Add "HPS_NAC_Activo"
        .Add "HPS_NAC_ApuntoDeCaducar"
        .Add "HPS_NAC_Baja"
        .Add "HPS_NAC_Caducado"
        .Add "HPS_NAC_PendienteRenovacion"
        .Add "HPS_NAC_Solicitado"
        .Add "HPS_NAC_ESTADO"
        .Add "HPS_NAC_Irregular"
        .Add "HPS_NAC_MOTIVO_IRREGULAR"
        .Add "HPS_NAC_MESES_PARA_RENOVAR"
        .Add "HPS_OTAN_SIN_DATOS"
        .Add "HPS_OTAN_F_Concesion"
        .Add "HPS_OTAN_F_Caducidad"
        .Add "HPS_OTAN_Grado"
        .Add "HPS_OTAN_F_Solicitud"
        .Add "HPS_OTAN_Especialidad"
        .Add "HPS_OTAN_Renovacion"
        .Add "HPS_OTAN_F_Baja"
        .Add "HPS_OTAN_Activo"
        .Add "HPS_OTAN_ApuntoDeCaducar"
        .Add "HPS_OTAN_Baja"
        .Add "HPS_OTAN_Caducado"
        .Add "HPS_OTAN_PendienteRenovacion"
        .Add "HPS_OTAN_Solicitado"
        .Add "HPS_OTAN_ESTADO"
        .Add "HPS_OTAN_MOTIVO_IRREGULAR"
        .Add "HPS_OTAN_Irregular"
        .Add "HPS_OTAN_MESES_PARA_RENOVAR"
        .Add "HPS_ESA_SIN_DATOS"
        .Add "HPS_ESA_F_Concesion"
        .Add "HPS_ESA_F_Caducidad"
        .Add "HPS_ESA_Grado"
        .Add "HPS_ESA_F_Solicitud"
        .Add "HPS_ESA_Especialidad"
        .Add "HPS_ESA_Renovacion"
        .Add "HPS_ESA_F_Baja"
        .Add "HPS_ESA_Activo"
        .Add "HPS_ESA_ApuntoDeCaducar"
        .Add "HPS_ESA_Baja"
        .Add "HPS_ESA_Caducado"
        .Add "HPS_ESA_PendienteRenovacion"
        .Add "HPS_ESA_Solicitado"
        .Add "HPS_ESA_ESTADO"
        .Add "HPS_ESA_Irregular"
        .Add "HPS_ESA_MOTIVO_IRREGULAR"
        .Add "HPS_ESA_MESES_PARA_RENOVAR"
        .Add "HPS_UE_SIN_DATOS"
        .Add "HPS_UE_F_Concesion"
        .Add "HPS_UE_F_Caducidad"
        .Add "HPS_UE_Grado"
        .Add "HPS_UE_F_Solicitud"
        .Add "HPS_UE_Especialidad"
        .Add "HPS_UE_Renovacion"
        .Add "HPS_UE_F_Baja"
        .Add "HPS_UE_Activo"
        .Add "HPS_UE_ApuntoDeCaducar"
        .Add "HPS_UE_Baja"
        .Add "HPS_UE_Caducado"
        .Add "HPS_UE_PendienteRenovacion"
        .Add "HPS_UE_Solicitado"
        .Add "HPS_UE_ESTADO"
        .Add "HPS_UE_Irregular"
        .Add "HPS_UE_MOTIVO_IRREGULAR"
        .Add "HPS_UE_MESES_PARA_RENOVAR"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método UsuarioEntidad.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "Apellido1" Then
        getPropiedad = Me.Apellido1
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
    If m_NombrePropiedad = "EmpresaUsuario" Then
        getPropiedad = Me.EmpresaUsuario
        Exit Function
    End If
    If m_NombrePropiedad = "EmpresaTramitadora" Then
        getPropiedad = Me.EmpresaTramitadora
        Exit Function
    End If
    If m_NombrePropiedad = "JuridicaContrato" Then
        getPropiedad = Me.JuridicaContrato
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
    If m_NombrePropiedad = "CodExp" Then
        getPropiedad = Me.CodExp
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
    If m_NombrePropiedad = "F_Baja" Then
        getPropiedad = Me.F_Baja
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
    If m_NombrePropiedad = "Curso_Realizado" Then
        getPropiedad = Me.Curso_Realizado
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
    If m_NombrePropiedad = "Requiere_PrimeraConvocatoriaCurso" Then
        getPropiedad = Me.Requiere_PrimeraConvocatoriaCurso
        Exit Function
    End If
    If m_NombrePropiedad = "Requiere_SegundaConvocatoriaCurso" Then
        getPropiedad = Me.Requiere_SegundaConvocatoriaCurso
        Exit Function
    End If
    If m_NombrePropiedad = "Requiere_CorreoJefeSeguridadCurso" Then
        getPropiedad = Me.Requiere_CorreoJefeSeguridadCurso
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    If m_NombrePropiedad = "IDSolicitud" Then
        getPropiedad = Me.IDSolicitud
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_SIN_DATOS" Then
        getPropiedad = Me.HPS_NAC_SIN_DATOS
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_F_Concesion" Then
        getPropiedad = Me.HPS_NAC_F_Concesion
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_F_Caducidad" Then
        getPropiedad = Me.HPS_NAC_F_Caducidad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Grado" Then
        getPropiedad = Me.HPS_NAC_Grado
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_F_Solicitud" Then
        getPropiedad = Me.HPS_NAC_F_Solicitud
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Especialidad" Then
        getPropiedad = Me.HPS_NAC_Especialidad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Renovacion" Then
        getPropiedad = Me.HPS_NAC_Renovacion
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_F_Baja" Then
        getPropiedad = Me.HPS_NAC_F_Baja
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Activo" Then
        getPropiedad = Me.HPS_NAC_Activo
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_ApuntoDeCaducar" Then
        getPropiedad = Me.HPS_NAC_ApuntoDeCaducar
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Baja" Then
        getPropiedad = Me.HPS_NAC_Baja
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Caducado" Then
        getPropiedad = Me.HPS_NAC_Caducado
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_PendienteRenovacion" Then
        getPropiedad = Me.HPS_NAC_PendienteRenovacion
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Solicitado" Then
        getPropiedad = Me.HPS_NAC_Solicitado
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_ESTADO" Then
        getPropiedad = Me.HPS_NAC_ESTADO
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Irregular" Then
        getPropiedad = Me.HPS_NAC_Irregular
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_MOTIVO_IRREGULAR" Then
        getPropiedad = Me.HPS_NAC_MOTIVO_IRREGULAR
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_MESES_PARA_RENOVAR" Then
        getPropiedad = Me.HPS_NAC_MESES_PARA_RENOVAR
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_SIN_DATOS" Then
        getPropiedad = Me.HPS_OTAN_SIN_DATOS
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_F_Concesion" Then
        getPropiedad = Me.HPS_OTAN_F_Concesion
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_F_Caducidad" Then
        getPropiedad = Me.HPS_OTAN_F_Caducidad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Grado" Then
        getPropiedad = Me.HPS_OTAN_Grado
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_F_Solicitud" Then
        getPropiedad = Me.HPS_OTAN_F_Solicitud
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Especialidad" Then
        getPropiedad = Me.HPS_OTAN_ESPECIALIDAD
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Renovacion" Then
        getPropiedad = Me.HPS_OTAN_Renovacion
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_F_Baja" Then
        getPropiedad = Me.HPS_OTAN_F_Baja
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Activo" Then
        getPropiedad = Me.HPS_OTAN_Activo
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_ApuntoDeCaducar" Then
        getPropiedad = Me.HPS_OTAN_ApuntoDeCaducar
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Baja" Then
        getPropiedad = Me.HPS_OTAN_Baja
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Caducado" Then
        getPropiedad = Me.HPS_OTAN_Caducado
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_PendienteRenovacion" Then
        getPropiedad = Me.HPS_OTAN_PendienteRenovacion
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Solicitado" Then
        getPropiedad = Me.HPS_OTAN_Solicitado
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_ESTADO" Then
        getPropiedad = Me.HPS_OTAN_ESTADO
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_MOTIVO_IRREGULAR" Then
        getPropiedad = Me.HPS_OTAN_MOTIVO_IRREGULAR
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Irregular" Then
        getPropiedad = Me.HPS_OTAN_Irregular
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_MESES_PARA_RENOVAR" Then
        getPropiedad = Me.HPS_OTAN_MESES_PARA_RENOVAR
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_SIN_DATOS" Then
        getPropiedad = Me.HPS_ESA_SIN_DATOS
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_F_Concesion" Then
        getPropiedad = Me.HPS_ESA_F_Concesion
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_F_Caducidad" Then
        getPropiedad = Me.HPS_ESA_F_Caducidad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Grado" Then
        getPropiedad = Me.HPS_ESA_Grado
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_F_Solicitud" Then
        getPropiedad = Me.HPS_ESA_F_Solicitud
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Especialidad" Then
        getPropiedad = Me.HPS_ESA_ESPECIALIDAD
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Renovacion" Then
        getPropiedad = Me.HPS_ESA_Renovacion
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_F_Baja" Then
        getPropiedad = Me.HPS_ESA_F_Baja
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Activo" Then
        getPropiedad = Me.HPS_ESA_Activo
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_ApuntoDeCaducar" Then
        getPropiedad = Me.HPS_ESA_ApuntoDeCaducar
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Baja" Then
        getPropiedad = Me.HPS_ESA_Baja
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Caducado" Then
        getPropiedad = Me.HPS_ESA_Caducado
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_PendienteRenovacion" Then
        getPropiedad = Me.HPS_ESA_PendienteRenovacion
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Solicitado" Then
        getPropiedad = Me.HPS_ESA_Solicitado
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_ESTADO" Then
        getPropiedad = Me.HPS_ESA_ESTADO
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Irregular" Then
        getPropiedad = Me.HPS_ESA_Irregular
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_MOTIVO_IRREGULAR" Then
        getPropiedad = Me.HPS_ESA_MOTIVO_IRREGULAR
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_MESES_PARA_RENOVAR" Then
        getPropiedad = Me.HPS_ESA_MESES_PARA_RENOVAR
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_SIN_DATOS" Then
        getPropiedad = Me.HPS_UE_SIN_DATOS
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_F_Concesion" Then
        getPropiedad = Me.HPS_UE_F_Concesion
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_F_Caducidad" Then
        getPropiedad = Me.HPS_UE_F_Caducidad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Grado" Then
        getPropiedad = Me.HPS_UE_Grado
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_F_Solicitud" Then
        getPropiedad = Me.HPS_UE_F_Solicitud
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Especialidad" Then
        getPropiedad = Me.HPS_UE_ESPECIALIDAD
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Renovacion" Then
        getPropiedad = Me.HPS_UE_Renovacion
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_F_Baja" Then
        getPropiedad = Me.HPS_UE_F_Baja
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Activo" Then
        getPropiedad = Me.HPS_UE_Activo
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_ApuntoDeCaducar" Then
        getPropiedad = Me.HPS_UE_ApuntoDeCaducar
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Baja" Then
        getPropiedad = Me.HPS_UE_Baja
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Caducado" Then
        getPropiedad = Me.HPS_UE_Caducado
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_PendienteRenovacion" Then
        getPropiedad = Me.HPS_UE_PendienteRenovacion
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Solicitado" Then
        getPropiedad = Me.HPS_UE_Solicitado
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_ESTADO" Then
        getPropiedad = Me.HPS_UE_ESTADO
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Irregular" Then
        getPropiedad = Me.HPS_UE_Irregular
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_MOTIVO_IRREGULAR" Then
        getPropiedad = Me.HPS_UE_MOTIVO_IRREGULAR
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_MESES_PARA_RENOVAR" Then
        getPropiedad = Me.HPS_UE_MESES_PARA_RENOVAR
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en UsuarioEntidad.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioEntidad.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "Apellido1" Then
        Me.Apellido1 = m_ValorPropiedad
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
    If m_NombrePropiedad = "EmpresaUsuario" Then
        Me.EmpresaUsuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EmpresaTramitadora" Then
        Me.EmpresaTramitadora = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "JuridicaContrato" Then
        Me.JuridicaContrato = m_ValorPropiedad
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
    If m_NombrePropiedad = "CodExp" Then
        Me.CodExp = m_ValorPropiedad
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
    If m_NombrePropiedad = "F_Baja" Then
        Me.F_Baja = m_ValorPropiedad
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
    If m_NombrePropiedad = "Curso_Realizado" Then
        Me.Curso_Realizado = m_ValorPropiedad
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
    If m_NombrePropiedad = "Requiere_PrimeraConvocatoriaCurso" Then
        Me.Requiere_PrimeraConvocatoriaCurso = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Requiere_SegundaConvocatoriaCurso" Then
        Me.Requiere_SegundaConvocatoriaCurso = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Requiere_CorreoJefeSeguridadCurso" Then
        Me.Requiere_CorreoJefeSeguridadCurso = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDSolicitud" Then
        Me.IDSolicitud = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_SIN_DATOS" Then
        Me.HPS_NAC_SIN_DATOS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_F_Concesion" Then
        Me.HPS_NAC_F_Concesion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_F_Caducidad" Then
        Me.HPS_NAC_F_Caducidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Grado" Then
        Me.HPS_NAC_Grado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_F_Solicitud" Then
        Me.HPS_NAC_F_Solicitud = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Especialidad" Then
        Me.HPS_NAC_Especialidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Renovacion" Then
        Me.HPS_NAC_Renovacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_F_Baja" Then
        Me.HPS_NAC_F_Baja = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Activo" Then
        Me.HPS_NAC_Activo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_ApuntoDeCaducar" Then
        Me.HPS_NAC_ApuntoDeCaducar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Baja" Then
        Me.HPS_NAC_Baja = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Caducado" Then
        Me.HPS_NAC_Caducado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_PendienteRenovacion" Then
        Me.HPS_NAC_PendienteRenovacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Solicitado" Then
        Me.HPS_NAC_Solicitado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_ESTADO" Then
        Me.HPS_NAC_ESTADO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_Irregular" Then
        Me.HPS_NAC_Irregular = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_MOTIVO_IRREGULAR" Then
        Me.HPS_NAC_MOTIVO_IRREGULAR = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_NAC_MESES_PARA_RENOVAR" Then
        Me.HPS_NAC_MESES_PARA_RENOVAR = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_SIN_DATOS" Then
        Me.HPS_OTAN_SIN_DATOS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_F_Concesion" Then
        Me.HPS_OTAN_F_Concesion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_F_Caducidad" Then
        Me.HPS_OTAN_F_Caducidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Grado" Then
        Me.HPS_OTAN_Grado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_F_Solicitud" Then
        Me.HPS_OTAN_F_Solicitud = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Especialidad" Then
        Me.HPS_OTAN_ESPECIALIDAD = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Renovacion" Then
        Me.HPS_OTAN_Renovacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_F_Baja" Then
        Me.HPS_OTAN_F_Baja = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Activo" Then
        Me.HPS_OTAN_Activo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_ApuntoDeCaducar" Then
        Me.HPS_OTAN_ApuntoDeCaducar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Baja" Then
        Me.HPS_OTAN_Baja = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Caducado" Then
        Me.HPS_OTAN_Caducado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_PendienteRenovacion" Then
        Me.HPS_OTAN_PendienteRenovacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Solicitado" Then
        Me.HPS_OTAN_Solicitado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_ESTADO" Then
        Me.HPS_OTAN_ESTADO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_MOTIVO_IRREGULAR" Then
        Me.HPS_OTAN_MOTIVO_IRREGULAR = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_Irregular" Then
        Me.HPS_OTAN_Irregular = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_OTAN_MESES_PARA_RENOVAR" Then
        Me.HPS_OTAN_MESES_PARA_RENOVAR = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_SIN_DATOS" Then
        Me.HPS_ESA_SIN_DATOS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_F_Concesion" Then
        Me.HPS_ESA_F_Concesion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_F_Caducidad" Then
        Me.HPS_ESA_F_Caducidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Grado" Then
        Me.HPS_ESA_Grado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_F_Solicitud" Then
        Me.HPS_ESA_F_Solicitud = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Especialidad" Then
        Me.HPS_ESA_ESPECIALIDAD = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Renovacion" Then
        Me.HPS_ESA_Renovacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_F_Baja" Then
        Me.HPS_ESA_F_Baja = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Activo" Then
        Me.HPS_ESA_Activo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_ApuntoDeCaducar" Then
        Me.HPS_ESA_ApuntoDeCaducar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Baja" Then
        Me.HPS_ESA_Baja = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Caducado" Then
        Me.HPS_ESA_Caducado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_PendienteRenovacion" Then
        Me.HPS_ESA_PendienteRenovacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Solicitado" Then
        Me.HPS_ESA_Solicitado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_ESTADO" Then
        Me.HPS_ESA_ESTADO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_Irregular" Then
        Me.HPS_ESA_Irregular = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_MOTIVO_IRREGULAR" Then
        Me.HPS_ESA_MOTIVO_IRREGULAR = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_ESA_MESES_PARA_RENOVAR" Then
        Me.HPS_ESA_MESES_PARA_RENOVAR = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_SIN_DATOS" Then
        Me.HPS_UE_SIN_DATOS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_F_Concesion" Then
        Me.HPS_UE_F_Concesion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_F_Caducidad" Then
        Me.HPS_UE_F_Caducidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Grado" Then
        Me.HPS_UE_Grado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_F_Solicitud" Then
        Me.HPS_UE_F_Solicitud = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Especialidad" Then
        Me.HPS_UE_ESPECIALIDAD = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Renovacion" Then
        Me.HPS_UE_Renovacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_F_Baja" Then
        Me.HPS_UE_F_Baja = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Activo" Then
        Me.HPS_UE_Activo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_ApuntoDeCaducar" Then
        Me.HPS_UE_ApuntoDeCaducar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Baja" Then
        Me.HPS_UE_Baja = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Caducado" Then
        Me.HPS_UE_Caducado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_PendienteRenovacion" Then
        Me.HPS_UE_PendienteRenovacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Solicitado" Then
        Me.HPS_UE_Solicitado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_ESTADO" Then
        Me.HPS_UE_ESTADO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_Irregular" Then
        Me.HPS_UE_Irregular = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_MOTIVO_IRREGULAR" Then
        Me.HPS_UE_MOTIVO_IRREGULAR = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPS_UE_MESES_PARA_RENOVAR" Then
        Me.HPS_UE_MESES_PARA_RENOVAR = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en UsuarioEntidad.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioEntidad.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

