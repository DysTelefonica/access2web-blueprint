VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PMAccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDAccionMitigacion As String
Public IDMitigacion As String
Public CodAccion As String
Public Accion As String
Public ResponsableAccion As String
Public FechaInicio As String
Public FechaFinPrevista As String
Public FechaFinReal As String
Public EsUltimaAccion As String
Public Estado As String
Public NombreIcono As String

Private m_sNombreIconoCalculado As String


Private m_sIDAccionMitigacionCalculada As String
Private m_sCodigoAccionNueva As String
Private m_ObjMitigacion As PM
Private m_objPMAccionEdicionAnterior As PMAccion
Private m_objPMAccionEdicionSiguiente As PMAccion
Private m_eUsuarioConectadoAutorizado As EnumSiNo
Private m_eEsActivo As EnumSiNo
Private m_eNecesitaReplanificacion As EnumSiNo
Private m_eBorrable As EnumSiNo

Private m_objColCampos As Collection
Private m_objColCamposParaCambios As Collection
Public Error As String
Private m_Error As String

Public Property Get NombreIconoCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sNombreIconoCalculado = getNombreIconoAccionArbolDeRiesgos(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    NombreIconoCalculado = m_sNombreIconoCalculado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PCAccion.NombreIconoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Borrable() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eBorrable <> 0 Then
        Borrable = m_eBorrable
        Exit Property
    End If
    m_eBorrable = Me.Mitigacion.Borrable
    m_Error = Me.Mitigacion.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Borrable = m_eBorrable
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PMAccion.Borrable ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get NecesitaReplanificacion() As EnumSiNo
    
    Dim m_FechaRef As Date
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_FechaRef = Date
    'm_FechaRef = "12/12/2023"
    If Not IsDate(Me.FechaFinPrevista) Or IsDate(Me.FechaFinReal) Then
        m_eNecesitaReplanificacion = EnumSiNo.No
        NecesitaReplanificacion = m_eNecesitaReplanificacion
        Exit Property
    End If
    If CDate(Me.FechaFinPrevista) < m_FechaRef Then
        m_eNecesitaReplanificacion = EnumSiNo.Sí
    Else
        m_eNecesitaReplanificacion = EnumSiNo.No
    End If
    
    NecesitaReplanificacion = m_eNecesitaReplanificacion
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PMAccion.NecesitaReplanificacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get IDAccionMitigacionCalculada() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDAccionMitigacionCalculada = DameID("TbRiesgosPlanMitigacionDetalle", "IDAccionMitigacion", , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDAccionMitigacionCalculada = m_sIDAccionMitigacionCalculada
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PMAccion.IDMitigacionCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get CodigoAccionNueva() As String
    
    Dim m_IDAccionMitigacion As Variant
    Dim m_ObjPMAccion As PMAccion
    Dim m_CodPMAccion As String
    Dim m_intOrdinalMaxCod As Integer
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.Mitigacion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se sabe a qué Plan de mitigación Pertenece"
        Err.Raise 1000
    End If
    If Not Me.Mitigacion.colAcciones Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        For Each m_IDAccionMitigacion In Me.Mitigacion.colAcciones.keys
            Set m_ObjPMAccion = Me.Mitigacion.colAcciones(m_IDAccionMitigacion)
            m_CodPMAccion = m_ObjPMAccion.CodAccion
            If IsNumeric(Right(m_CodPMAccion, Len(m_CodPMAccion) - 3)) Then
                If CInt(Right(m_CodPMAccion, Len(m_CodPMAccion) - 3)) > m_intOrdinalMaxCod Then
                    m_intOrdinalMaxCod = CInt(Right(m_CodPMAccion, Len(m_CodPMAccion) - 3))
                End If
            End If
            Set m_ObjPMAccion = Nothing
        Next
    End If
    
    m_sCodigoAccionNueva = "AC" & Format(m_intOrdinalMaxCod + 1, "000")
    CodigoAccionNueva = m_sCodigoAccionNueva
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.CodigoAccionNueva ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property




Public Property Get Mitigacion() As PM

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjMitigacion Is Nothing Then
        Set Mitigacion = m_ObjMitigacion
        Exit Property
    End If
    If Me.IDMitigacion = "" Then
        m_Error = "La acción de Mitigacion no se sabe a qué plan pertence"
        Err.Raise 1000
    End If
    Set m_ObjMitigacion = Constructor.getPM(Me.IDMitigacion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Mitigacion = m_ObjMitigacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PMAccion.Mitigacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set Mitigacion(ByVal sNewValue As PM)

    Set m_ObjMitigacion = sNewValue

End Property


Public Property Get PMAccionEdicionAnterior() As PMAccion
    
    Dim m_ObjPMAnterior As PM
    Dim m_IDAcion As Variant
    Dim m_ObjPMAccion As PMAccion
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objPMAccionEdicionAnterior Is Nothing Then
        Set PMAccionEdicionAnterior = m_objPMAccionEdicionAnterior
        Exit Property
    End If
    If Me.Mitigacion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se conoce a qué Plan de mitigación pertenece la acción"
        Err.Raise 1000
    End If
    Set m_ObjPMAnterior = Me.Mitigacion.PMEdicionAnterior
    m_Error = Me.Mitigacion.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjPMAnterior Is Nothing Then
        Exit Property
    End If
    If m_ObjPMAnterior.colAcciones Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_IDAcion In m_ObjPMAnterior.colAcciones.keys
        Set m_ObjPMAccion = m_ObjPMAnterior.colAcciones(m_IDAcion)
        If m_ObjPMAccion.CodAccion = Me.CodAccion Then
            Set m_objPMAccionEdicionAnterior = m_ObjPMAccion
            Exit For
        End If
        Set m_ObjPMAccion = Nothing
    Next
    Set PMAccionEdicionAnterior = m_objPMAccionEdicionAnterior
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PMAccion.PMAccionEdicionAnterior ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PMAccionEdicionSiguiente() As PMAccion
    
    Dim m_ObjPMSiguiente As PM
    Dim m_IDAcion As Variant
    Dim m_ObjPMAccion As PMAccion
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objPMAccionEdicionSiguiente Is Nothing Then
        Set PMAccionEdicionSiguiente = m_objPMAccionEdicionSiguiente
        Exit Property
    End If
    If Me.Mitigacion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se conoce a qué riesgo Plan de mitigación pertenece la acción"
        Err.Raise 1000
    End If
    Set m_ObjPMSiguiente = Me.Mitigacion.PMEdicionSiguiente
    m_Error = Me.Mitigacion.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjPMSiguiente Is Nothing Then
        Exit Property
    End If
    If m_ObjPMSiguiente.colAcciones Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_IDAcion In m_ObjPMSiguiente.colAcciones.keys
        Set m_ObjPMAccion = m_ObjPMSiguiente.colAcciones(m_IDAcion)
        If m_ObjPMAccion.CodAccion = Me.CodAccion Then
            Set m_objPMAccionEdicionSiguiente = m_ObjPMAccion
            Exit For
        End If
        Set m_ObjPMAccion = Nothing
    Next
    Set PMAccionEdicionSiguiente = m_objPMAccionEdicionSiguiente
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PMAccion.PMAccionEdicionSiguiente ha devuelto el error: " & vbNewLine & Err.Description
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
    m_eUsuarioConectadoAutorizado = UsuarioAutorizado(Me, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    UsuarioConectadoAutorizado = m_eUsuarioConectadoAutorizado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PMAccion.UsuarioConectadoAutorizado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get EsActivo() As EnumSiNo
    
    Dim m_objEdicion As Edicion
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eEsActivo <> 0 Then
        EsActivo = m_eEsActivo
        Exit Property
    End If
    Set m_objEdicion = Me.Mitigacion.riesgo.Edicion
    m_eEsActivo = m_objEdicion.EsActivo
    EsActivo = m_eEsActivo
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PMAccion.EsActivo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Function MotivoNoOK( _
                                Optional ByRef p_ObjPMAccionAlInicio As PMAccion, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    
    On Error GoTo errores
        
    With Me
        
        If .Mitigacion Is Nothing Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "No se conoce la Mitigacion"
            Exit Function
        End If
        If .Mitigacion.riesgo Is Nothing Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "No se conoce el Riesgo"
            Exit Function
        End If
        If .Mitigacion.riesgo.Estado = "" Then
            .Mitigacion.riesgo.Estado = .Mitigacion.riesgo.ESTADOCalculadoTexto
            p_Error = .Mitigacion.riesgo.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If .Mitigacion.riesgo.FichaRiesgoCompleta = EnumSiNo.No Then
            MotivoNoOK = "No se puede editar ni dar de alta una acción de Mitigacion de un Riesgo al que le falta " & _
                        "completar los datos"
            Exit Function
        End If
        
        If .EsActivo <> EnumSiNo.Sí Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "El Plan de mitigación al que pertenece esta acción no está activo"
            Exit Function
        End If
        If .UsuarioConectadoAutorizado <> EnumSiNo.Sí Then
            MotivoNoOK = "El perfil de su usuario no tiene autorización para realizar esta operación"
            Exit Function
        End If
        If .Accion = "" Then
            MotivoNoOK = "Obligatorio el disparador del plan"
            Exit Function
        End If
        If p_ObjPMAccionAlInicio Is Nothing Then
            
            .CodAccion = .CodigoAccionNueva
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If .CodAccion = "" Then
                MotivoNoOK = "No se ha podido establecer el Código de la acción"
                Exit Function
            End If
            
            
        Else
            If .CodAccion = "" Then
                MotivoNoOK = "No se ha podido establecer el Código de la acción"
                Exit Function
            End If
        End If
        If .ResponsableAccion = "" Then
            MotivoNoOK = "Falta la ResponsableAccion"
            Exit Function
        End If
        If IsDate(.FechaInicio) And Not IsDate(.FechaFinPrevista) Then
            MotivoNoOK = "Si se introduce la fecha de inicio hemos de poner la de fin prevista"
            Exit Function
        End If
        If IsDate(.FechaInicio) And IsDate(.FechaFinPrevista) Then
            If CDate(.FechaInicio) > CDate(.FechaFinPrevista) Then
                MotivoNoOK = "La fecha fin previata ha de ser posterior a la de inicio"
                Exit Function
            End If
        End If
        If IsDate(.FechaFinReal) And Not IsDate(.FechaInicio) Then
            MotivoNoOK = "Está rellena la fecha Fin Real y no la de Inicio"
            Exit Function
        End If
        If IsDate(.FechaFinReal) And IsDate(.FechaInicio) Then
            If CDate(.FechaFinReal) < CDate(.FechaInicio) Then
                MotivoNoOK = "La fecha fin real ha de ser posterior a la de inicio"
                Exit Function
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
                            Optional ByRef p_ObjPMAccionAlInicio As PMAccion, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_MotivoNoOK As String
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    
    m_MotivoNoOK = MotivoNoOK(p_ObjPMAccionAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.CodAccion = "" Then
        Me.CodAccion = Me.CodigoAccionNueva
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If p_ObjPMAccionAlInicio Is Nothing Then
        m_SQL = "TbRiesgosPlanMitigacionDetalle"
        Me.IDAccionMitigacion = Me.IDAccionMitigacionCalculada
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        m_SQL = "SELECT * FROM TbRiesgosPlanMitigacionDetalle " & _
            "WHERE IDAccionMitigacion=" & Me.IDAccionMitigacion & ";"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not p_ObjPMAccionAlInicio Is Nothing Then
            If .EOF Then
                p_Error = "No se ha encontrado el registro"
                Err.Raise 1000
            End If
        End If
        If p_ObjPMAccionAlInicio Is Nothing Then
            .AddNew
                .Fields("IDAccionMitigacion") = Me.IDAccionMitigacion
                .Fields("IDMitigacion") = Me.IDMitigacion
                .Fields("CodAccion") = Me.CodAccion
                .Fields("Accion") = Me.Accion
                .Fields("ResponsableAccion") = Me.ResponsableAccion
                
                If IsDate(Me.FechaInicio) Then
                    .Fields("FechaInicio") = Me.FechaInicio
                End If
                If IsDate(Me.FechaFinPrevista) Then
                    .Fields("FechaFinPrevista") = Me.FechaFinPrevista
                End If
                If IsDate(Me.FechaFinReal) Then
                    .Fields("FechaFinReal") = Me.FechaFinReal
                End If
                
                
            .Update
        Else
            .Edit
                .Fields("CodAccion") = Me.CodAccion
                .Fields("Accion") = Me.Accion
                .Fields("ResponsableAccion") = Me.ResponsableAccion
                
                If IsDate(Me.FechaInicio) Then
                    .Fields("FechaInicio") = Me.FechaInicio
                Else
                    .Fields("FechaInicio") = Null
                End If
                If IsDate(Me.FechaFinPrevista) Then
                    .Fields("FechaFinPrevista") = Me.FechaFinPrevista
                Else
                    .Fields("FechaFinPrevista") = Null
                End If
                If IsDate(Me.FechaFinReal) Then
                    .Fields("FechaFinReal") = Me.FechaFinReal
                Else
                    .Fields("FechaFinReal") = Null
                End If
                
                .Fields("Estado") = Null
            .Update
        
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Set Me.Mitigacion.colAcciones = Nothing
    Set Me.Mitigacion.riesgo.ColPMs = Nothing
    RegistrarActualizacionRiesgo p_Riesgo:=Me.Mitigacion.riesgo, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarActualizacionEdicion p_Edicion:=Me.Mitigacion.riesgo.Edicion, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha devuelto el error: " & Err.Description
    End If
End Function

Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                            
    '       1)- Sólo puede dar de alta una Edición el Autorizado (administradores,Calidad y JPs del Proyecto)
    '       2)- No puede provenir de un riesgo que pertenece a una Edición ya publicada
    Dim m_Particula As String
    
    On Error GoTo errores
    
    If Me.UsuarioConectadoAutorizado = EnumSiNo.No Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "El perfil de su usuario no tiene autorización para realizar esta operación"
        Err.Raise 1000
    End If
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "La Edición ya no permite modificaciones"
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PMAccion.ValidacionEliminar ha producido el error " & vbNewLine & Err.Description
    End If
End Function
Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String

    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "DELETE * FROM TbRiesgosPlanMitigacionDetalle " & _
            "WHERE IDAccionMitigacion=" & Me.IDAccionMitigacion & ";"
    getdb().Execute m_SQL
    
    Set Me.Mitigacion.colAcciones = Nothing
    RegistrarActualizacionRiesgo p_Riesgo:=Me.Mitigacion.riesgo, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarActualizacionEdicion p_Edicion:=Me.Mitigacion.riesgo.Edicion, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PMAccion.Borrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function GrabarNombreIcono( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    
    Me.NombreIcono = Me.NombreIconoCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "UPDATE TbRiesgosPlanMitigacionDetalle SET NombreIcono = '" & Me.NombreIcono & "' " & _
            "WHERE IDAccionMitigacion=" & Me.IDAccionMitigacion & ";"
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PMAccion.GrabarNombreIcono ha devuelto el error: " & Err.Description
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
        .Add "IDAccionMitigacion"
        .Add "IDMitigacion"
        .Add "CodAccion"
        .Add "Accion"
        .Add "ResponsableAccion"
        .Add "FechaInicio"
        .Add "FechaFinPrevista"
        .Add "FechaFinReal"
        .Add "Estado"
        .Add "NombreIcono"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PMAccion.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get ColCamposParaCambios() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColCamposParaCambios Is Nothing Then
        Set ColCamposParaCambios = m_objColCamposParaCambios
        Exit Property
    End If
    Set m_objColCamposParaCambios = New Collection
    With m_objColCamposParaCambios
        .Add "Accion"
        .Add "ResponsableAccion"
        .Add "FechaInicio"
        .Add "FechaFinPrevista"
        .Add "FechaFinReal"
        .Add "Estado"
    End With
    Set ColCamposParaCambios = m_objColCamposParaCambios
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PMAccion.ColCamposParaCambios ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDAccionMitigacion" Then
        getPropiedad = Me.IDAccionMitigacion
        Exit Function
    End If
    If m_NombrePropiedad = "IDMitigacion" Then
        getPropiedad = Me.IDMitigacion
        Exit Function
    End If
    If m_NombrePropiedad = "CodAccion" Then
        getPropiedad = Me.CodAccion
        Exit Function
    End If
    If m_NombrePropiedad = "Accion" Then
        getPropiedad = Me.Accion
        Exit Function
    End If
    If m_NombrePropiedad = "ResponsableAccion" Then
        getPropiedad = Me.ResponsableAccion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicio" Then
        getPropiedad = Me.FechaInicio
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinPrevista" Then
        getPropiedad = Me.FechaFinPrevista
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinReal" Then
        getPropiedad = Me.FechaFinReal
        Exit Function
    End If
    If m_NombrePropiedad = "EsUltimaAccion" Then
        getPropiedad = Me.EsUltimaAccion
        Exit Function
    End If
    If m_NombrePropiedad = "Estado" Then
        getPropiedad = Me.Estado
        Exit Function
    End If
    If m_NombrePropiedad = "NombreIcono" Then
        getPropiedad = Me.NombreIcono
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en PMAccion.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PMAccion.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDAccionMitigacion" Then
        Me.IDAccionMitigacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDMitigacion" Then
        Me.IDMitigacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodAccion" Then
        Me.CodAccion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Accion" Then
        Me.Accion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ResponsableAccion" Then
        Me.ResponsableAccion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicio" Then
        Me.FechaInicio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinPrevista" Then
        Me.FechaFinPrevista = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinReal" Then
        Me.FechaFinReal = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsUltimaAccion" Then
        Me.EsUltimaAccion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Estado" Then
        Me.Estado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreIcono" Then
        Me.NombreIcono = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en PMAccion.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PMAccion.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function







