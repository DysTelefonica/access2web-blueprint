VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PCAccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDAccionContingencia As String
Public IDContingencia As String
Public CodAccion As String
Public Accion As String
Public ResponsableAccion As String
Public FechaInicio As String
Public FechaFinPrevista As String
Public FechaFinReal As String
Public Estado As String
Public NombreIcono As String

Private m_sNombreIconoCalculado As String
Private m_sIDAccionContingenciaCalculada As String
Private m_sCodigoAccionNueva As String

Private m_ObjContingencia As PC

Private m_objPCAccionEdicionAnterior As PCAccion
Private m_objPCAccionEdicionSiguiente As PCAccion
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
    m_eBorrable = Me.Contingencia.Borrable
    m_Error = Me.Contingencia.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Borrable = m_eBorrable
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PCAccion.Borrable ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get NecesitaReplanificacion() As EnumSiNo
    
    Dim m_FechaRef As Date
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
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
        m_Error = "El método PCAccion.NecesitaReplanificacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get IDAccionContingenciaCalculada() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDAccionContingenciaCalculada = DameID("TbRiesgosPlanContingenciaDetalle", "IDAccionContingencia", , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDAccionContingenciaCalculada = m_sIDAccionContingenciaCalculada
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PCAccion.IDContingenciaCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get CodigoAccionNueva() As String
    
    Dim m_IDAccionContingencia As Variant
    Dim m_ObjPCAccion As PCAccion
    Dim m_CodPCAccion As String
    Dim m_intOrdinalMaxCod As Integer
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.Contingencia Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se sabe a qué plan de contingencia Pertenece"
        Err.Raise 1000
    End If
    If Not Me.Contingencia.colAcciones Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        For Each m_IDAccionContingencia In Me.Contingencia.colAcciones.keys
            Set m_ObjPCAccion = Me.Contingencia.colAcciones(m_IDAccionContingencia)
            m_CodPCAccion = m_ObjPCAccion.CodAccion
            If IsNumeric(Right(m_CodPCAccion, Len(m_CodPCAccion) - 3)) Then
                If CInt(Right(m_CodPCAccion, Len(m_CodPCAccion) - 3)) > m_intOrdinalMaxCod Then
                    m_intOrdinalMaxCod = CInt(Right(m_CodPCAccion, Len(m_CodPCAccion) - 3))
                End If
            End If
            Set m_ObjPCAccion = Nothing
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





Public Property Get Contingencia() As PC

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjContingencia Is Nothing Then
        Set Contingencia = m_ObjContingencia
        Exit Property
    End If
    If Me.IDContingencia = "" Then
        m_Error = "La acción de Contingencia no se sabe a qué plan pertence"
        Err.Raise 1000
    End If
    Set m_ObjContingencia = Constructor.getPC(Me.IDContingencia, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Contingencia = m_ObjContingencia
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PCAccion.Contingencia ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set Contingencia(ByVal sNewValue As PC)

    Set m_ObjContingencia = sNewValue

End Property




Public Property Get PCAccionEdicionAnterior() As PCAccion
    
    Dim m_ObjPCAnterior As PC
    Dim m_IDAcion As Variant
    Dim m_ObjPCAccion As PCAccion
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objPCAccionEdicionAnterior Is Nothing Then
        Set PCAccionEdicionAnterior = m_objPCAccionEdicionAnterior
        Exit Property
    End If
    If Me.Contingencia Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se conoce a qué plan de contingencia pertenece la acción"
        Err.Raise 1000
    End If
    Set m_ObjPCAnterior = Me.Contingencia.PCEdicionAnterior
    m_Error = Me.Contingencia.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjPCAnterior Is Nothing Then
        Exit Property
    End If
    If m_ObjPCAnterior.colAcciones Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_IDAcion In m_ObjPCAnterior.colAcciones.keys
        Set m_ObjPCAccion = m_ObjPCAnterior.colAcciones(m_IDAcion)
        If m_ObjPCAccion.CodAccion = Me.CodAccion Then
            Set m_objPCAccionEdicionAnterior = m_ObjPCAccion
            Exit For
        End If
        Set m_ObjPCAccion = Nothing
    Next
    Set PCAccionEdicionAnterior = m_objPCAccionEdicionAnterior
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PCAccion.PCAccionEdicionAnterior ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PCAccionEdicionSiguiente() As PCAccion
    
    Dim m_ObjPCSiguiente As PC
    Dim m_IDAcion As Variant
    Dim m_ObjPCAccion As PCAccion
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objPCAccionEdicionSiguiente Is Nothing Then
        Set PCAccionEdicionSiguiente = m_objPCAccionEdicionSiguiente
        Exit Property
    End If
    If Me.Contingencia Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se conoce a qué riesgo plan de contingencia pertenece la acción"
        Err.Raise 1000
    End If
    Set m_ObjPCSiguiente = Me.Contingencia.PCEdicionSiguiente
    m_Error = Me.Contingencia.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjPCSiguiente Is Nothing Then
        Exit Property
    End If
    If m_ObjPCSiguiente.colAcciones Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_IDAcion In m_ObjPCSiguiente.colAcciones.keys
        Set m_ObjPCAccion = m_ObjPCSiguiente.colAcciones(m_IDAcion)
        If m_ObjPCAccion.CodAccion = Me.CodAccion Then
            Set m_objPCAccionEdicionSiguiente = m_ObjPCAccion
            Exit For
        End If
        Set m_ObjPCAccion = Nothing
    Next
    Set PCAccionEdicionSiguiente = m_objPCAccionEdicionSiguiente
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PCAccion.PCAccionEdicionSiguiente ha devuelto el error: " & vbNewLine & Err.Description
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
        m_Error = "El método PCAccion.UsuarioConectadoAutorizado ha devuelto el error: " & vbNewLine & Err.Description
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
    Set m_objEdicion = Me.Contingencia.riesgo.Edicion
    
    m_eEsActivo = m_objEdicion.EsActivo
    EsActivo = m_eEsActivo
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PCAccion.EsActivo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Function MotivoNoOK( _
                                Optional ByRef p_ObjPCAccionAlInicio As PCAccion, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_CampoRepetido As EnumSiNo
    Dim m_UltimaAccionRegistrada As EnumSiNo
    Dim m_Riesgo As riesgo
    Dim blnComprobar As Boolean
    On Error GoTo errores
        
    With Me
        
        
        If .Contingencia Is Nothing Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "No se conoce la Contingencia"
            Exit Function
        End If
        Set m_Riesgo = .Contingencia.riesgo
        If m_Riesgo Is Nothing Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "No se conoce el Riesgo"
            Exit Function
        End If
        If m_Riesgo.Estado = "" Then
            m_Riesgo.Estado = m_Riesgo.ESTADOCalculadoTexto
            p_Error = m_Riesgo.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_Riesgo.FichaRiesgoCompleta = EnumSiNo.No Then
            MotivoNoOK = "No se puede editar ni dar de alta una acción de Contingencia de un Riesgo al que le falta " & _
                        "completar los datos"
            Exit Function
        End If
        
        If .EsActivo <> EnumSiNo.Sí Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "El plan de contingencia al que pertenece esta acción no está activo"
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
        If p_ObjPCAccionAlInicio Is Nothing Then
            
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
        
        
        blnComprobar = False
        If IsDate(.FechaInicio) Then
            If p_ObjPCAccionAlInicio Is Nothing Then
                blnComprobar = True
            Else
                If Not IsDate(p_ObjPCAccionAlInicio.FechaInicio) Then
                    blnComprobar = True
                Else
                    If p_ObjPCAccionAlInicio.FechaInicio <> .FechaInicio Then
                        blnComprobar = True
                    End If
                End If
            End If
        End If
        If blnComprobar = True Then
            If IsDate(m_Riesgo.FechaMaterializado) Then
            
            
                If CDate(.FechaInicio) < CDate(m_Riesgo.FechaMaterializado) Then
                    MotivoNoOK = "La fecha de inicio de la acción no puede ser anterior a la de materialización (" & m_Riesgo.FechaMaterializado & ")"
                    Exit Function
                End If
            End If
        End If
        blnComprobar = False
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
        If IsDate(.FechaInicio) Then
            If Not IsDate(.FechaFinPrevista) Then
                MotivoNoOK = "Si se introduce la fecha inicial de la actividad se ha de estimar la fecha fin prevista"
                Exit Function
            Else
                If CDate(.FechaInicio) > CDate(.FechaFinPrevista) Then
                    MotivoNoOK = "Si la fecha fin prevista no puede ser anterior a la de inicio"
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
                            Optional ByRef p_ObjPCAccionAlInicio As PCAccion, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_MotivoNoOK As String
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    
    m_MotivoNoOK = MotivoNoOK(p_ObjPCAccionAlInicio, p_Error)
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
    If p_ObjPCAccionAlInicio Is Nothing Then
        m_SQL = "TbRiesgosPlanContingenciaDetalle"
        Me.IDAccionContingencia = Me.IDAccionContingenciaCalculada
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        m_SQL = "SELECT * FROM TbRiesgosPlanContingenciaDetalle " & _
            "WHERE IDAccionContingencia=" & Me.IDAccionContingencia & ";"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not p_ObjPCAccionAlInicio Is Nothing Then
            If .EOF Then
                p_Error = "No se ha encontrado el registro"
                Err.Raise 1000
            End If
        End If
        If p_ObjPCAccionAlInicio Is Nothing Then
            .AddNew
                .Fields("IDAccionContingencia") = Me.IDAccionContingencia
                .Fields("IDContingencia") = Me.IDContingencia
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
                
                .Fields("Estado") = Me.Estado
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
                
                .Fields("Estado") = Me.Estado
            .Update
        
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Set Me.Contingencia.colAcciones = Nothing
    Set Me.Contingencia.riesgo.ColPCs = Nothing
    RegistrarActualizacionRiesgo p_Riesgo:=Me.Contingencia.riesgo, p_Error:=p_Error
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
        p_Error = "El método PCAccion.ValidacionEliminar ha producido el error " & vbNewLine & Err.Description
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
    m_SQL = "DELETE * FROM TbRiesgosPlanContingenciaDetalle " & _
            "WHERE IDAccionContingencia=" & Me.IDAccionContingencia & ";"
    getdb().Execute m_SQL
    
    Set Me.Contingencia.colAcciones = Nothing
    RegistrarActualizacionRiesgo p_Riesgo:=Me.Contingencia.riesgo, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarActualizacionEdicion p_Edicion:=Me.Contingencia.riesgo.Edicion, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PCAccion.Borrar ha devuelto el error: " & Err.Description
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
    
    m_SQL = "UPDATE TbRiesgosPlanContingenciaDetalle SET NombreIcono = '" & Me.NombreIcono & "' " & _
            "WHERE IDAccionContingencia=" & Me.IDAccionContingencia & ";"
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PCAccion.GrabarNombreIcono ha devuelto el error: " & Err.Description
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
        .Add "IDAccionContingencia"
        .Add "IDContingencia"
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
        m_Error = "El método PCAccion.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
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
        m_Error = "El método PCAccion.ColCamposParaCambios ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDAccionContingencia" Then
        getPropiedad = Me.IDAccionContingencia
        Exit Function
    End If
    If m_NombrePropiedad = "IDContingencia" Then
        getPropiedad = Me.IDContingencia
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
    
    If m_NombrePropiedad = "Estado" Then
        getPropiedad = Me.Estado
        Exit Function
    End If
    If m_NombrePropiedad = "NombreIcono" Then
        getPropiedad = Me.NombreIcono
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en PCAccion.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PCAccion.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDAccionContingencia" Then
        Me.IDAccionContingencia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDContingencia" Then
        Me.IDContingencia = m_ValorPropiedad
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
    
    If m_NombrePropiedad = "Estado" Then
        Me.Estado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreIcono" Then
        Me.NombreIcono = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en PCAccion.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PCAccion.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function








