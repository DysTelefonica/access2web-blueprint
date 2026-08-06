VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public idUsuario As String
Public tipoHps As String
Public F_Concesion As String
Public F_Caducidad As String
Public Grado As String
Public F_Baja As String
Public Renovacion As String
Public F_Solicitud As String
Public Especialidad As String
Public ParaAlta As EnumSiNo
Private m_ObjUsuario As UsuarioHPS
Private m_sMESES_PARA_RENOVAR As String
Private m_eIrregular As EnumSiNo
Private m_eActivo As EnumSiNo
Private m_eBaja As EnumSiNo
Private m_eMarcadoNoRenovacion As EnumSiNo
Private m_eApuntoDeCaducar As EnumSiNo
Private m_eCaducado As EnumSiNo
Private m_eSolicitado As EnumSiNo
Private m_ePendienteRenovacion As EnumSiNo
Private m_sEstado As String
Private m_sMOTIVO_IRREGULAR As String
Private m_ObjColDatosHPS As Scripting.Dictionary
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String




Public Property Get Irregular() As EnumSiNo
    
    Dim m_Irregular As String
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    Set m_ObjColDatosHPS = Nothing
    Set m_ObjColDatosHPS = getColDatosHPS(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_Irregular = DatosHPSCalculados("Irregular", Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Irregular = "Sí" Then
        m_eIrregular = EnumSiNo.Sí
    ElseIf m_Irregular = "No" Then
        m_eIrregular = EnumSiNo.No
    End If
    If m_eIrregular = EnumSiNo.Sí Or m_eIrregular = EnumSiNo.No Then
        Irregular = m_eIrregular
    End If
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HPS.Irregular ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property
Public Property Let Irregular(ByVal eNewValue As EnumSiNo)

    m_eIrregular = eNewValue

End Property
Public Property Get Activo() As EnumSiNo
    
    Dim m_Activo As String
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    Set m_ObjColDatosHPS = Nothing
    Set m_ObjColDatosHPS = getColDatosHPS(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_Activo = DatosHPSCalculados("Activo", Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Activo = "Sí" Then
        m_eActivo = EnumSiNo.Sí
    ElseIf m_Activo = "No" Then
        m_eActivo = EnumSiNo.No
    End If
    If m_eActivo = EnumSiNo.Sí Or m_eActivo = EnumSiNo.No Then
        Activo = m_eActivo
    End If
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HPS.Activo ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property
Public Property Let Activo(ByVal eNewValue As EnumSiNo)

    m_eActivo = eNewValue

End Property

Public Property Get Baja() As EnumSiNo
    
    Dim m_Baja As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_ObjColDatosHPS = Nothing
    Set m_ObjColDatosHPS = getColDatosHPS(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_Baja = DatosHPSCalculados("Baja", Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Baja = "Sí" Then
        m_eBaja = EnumSiNo.Sí
    ElseIf m_Baja = "No" Then
        m_eBaja = EnumSiNo.No
    End If
    If m_eBaja = EnumSiNo.Sí Or m_eBaja = EnumSiNo.No Then
        Baja = m_eBaja
    End If
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HPS.Baja ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Let Baja(ByVal eNewValue As EnumSiNo)

    m_eBaja = eNewValue

End Property

Public Property Get MarcadoNoRenovacion() As EnumSiNo
    
    Dim m_MarcadoNoRenovacion As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_ObjColDatosHPS = Nothing
    Set m_ObjColDatosHPS = getColDatosHPS(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_MarcadoNoRenovacion = DatosHPSCalculados("MarcadoNoRenovacion", Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_MarcadoNoRenovacion = "Sí" Then
        m_eMarcadoNoRenovacion = EnumSiNo.Sí
    ElseIf m_MarcadoNoRenovacion = "No" Then
        m_eMarcadoNoRenovacion = EnumSiNo.No
    End If
    If m_eMarcadoNoRenovacion = EnumSiNo.Sí Or m_eMarcadoNoRenovacion = EnumSiNo.No Then
        MarcadoNoRenovacion = m_eMarcadoNoRenovacion
    End If
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HPS.MarcadoNoRenovacion ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property
Public Property Let MarcadoNoRenovacion(ByVal eNewValue As EnumSiNo)

    m_eMarcadoNoRenovacion = eNewValue

End Property
Public Property Get APuntoDeCaducar() As EnumSiNo
    
    Dim m_APuntoDeCaducar As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_ObjColDatosHPS = Nothing
    Set m_ObjColDatosHPS = getColDatosHPS(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_APuntoDeCaducar = DatosHPSCalculados("ApuntoDeCaducar", Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_APuntoDeCaducar = "Sí" Then
        m_eApuntoDeCaducar = EnumSiNo.Sí
    ElseIf m_APuntoDeCaducar = "No" Then
        m_eApuntoDeCaducar = EnumSiNo.No
    End If
    If m_eApuntoDeCaducar = EnumSiNo.Sí Or m_eApuntoDeCaducar = EnumSiNo.No Then
        APuntoDeCaducar = m_eApuntoDeCaducar
    End If
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HPS.ApuntoDeCaducar ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property
Public Property Let APuntoDeCaducar(ByVal eNewValue As EnumSiNo)

    m_eApuntoDeCaducar = eNewValue

End Property
Public Property Get Caducado() As EnumSiNo
    
    Dim m_Caducado As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_ObjColDatosHPS = Nothing
    Set m_ObjColDatosHPS = getColDatosHPS(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_Caducado = DatosHPSCalculados("Caducado", Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Caducado = "Sí" Then
        m_eCaducado = EnumSiNo.Sí
    ElseIf m_Caducado = "No" Then
        m_eCaducado = EnumSiNo.No
    End If
    If m_eCaducado = EnumSiNo.Sí Or m_eCaducado = EnumSiNo.No Then
        Caducado = m_eCaducado
    End If
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HPS.Caducado ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property
Public Property Let Caducado(ByVal eNewValue As EnumSiNo)

    m_eCaducado = eNewValue

End Property
Public Property Get Solicitado() As EnumSiNo
    
    Dim m_Solicitado As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_ObjColDatosHPS = Nothing
    Set m_ObjColDatosHPS = getColDatosHPS(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_Solicitado = DatosHPSCalculados("Solicitada", Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Solicitado = "Sí" Then
        m_eSolicitado = EnumSiNo.Sí
    ElseIf m_Solicitado = "No" Then
        m_eSolicitado = EnumSiNo.No
    End If
    If m_eSolicitado = EnumSiNo.Sí Or m_eSolicitado = EnumSiNo.No Then
        Solicitado = m_eSolicitado
    End If
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HPS.Solicitado ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property
Public Property Let Solicitado(ByVal eNewValue As EnumSiNo)

    m_eSolicitado = eNewValue

End Property
Public Property Get PendienteRenovacion() As EnumSiNo
    
    Dim m_PendienteRenovacion As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_ObjColDatosHPS = Nothing
    Set m_ObjColDatosHPS = getColDatosHPS(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_PendienteRenovacion = DatosHPSCalculados("PendienteRenovacion", Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_PendienteRenovacion = "Sí" Then
        m_ePendienteRenovacion = EnumSiNo.Sí
    ElseIf m_PendienteRenovacion = "No" Then
        m_ePendienteRenovacion = EnumSiNo.No
    End If
    If m_ePendienteRenovacion = EnumSiNo.Sí Or m_ePendienteRenovacion = EnumSiNo.No Then
        PendienteRenovacion = m_ePendienteRenovacion
    End If
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HPS.PendienteRenovacion  ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property
Public Property Let PendienteRenovacion(ByVal eNewValue As EnumSiNo)

    m_ePendienteRenovacion = eNewValue

End Property

Public Property Get Estado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_ObjColDatosHPS = Nothing
    Set m_ObjColDatosHPS = getColDatosHPS(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sEstado = DatosHPSCalculados("Estado", Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Estado = m_sEstado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HPS.Estado ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Let Estado(ByVal eNewValue As String)

    m_sEstado = eNewValue

End Property
Public Property Get MOTIVO_IRREGULAR() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_ObjColDatosHPS = Nothing
    Set m_ObjColDatosHPS = getColDatosHPS(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sMOTIVO_IRREGULAR = DatosHPSCalculados("MOTIVO_IRREGULAR", Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    MOTIVO_IRREGULAR = m_sMOTIVO_IRREGULAR
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HPS.MOTIVO_IRREGULAR ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Let MOTIVO_IRREGULAR(ByVal eNewValue As String)

    m_sMOTIVO_IRREGULAR = eNewValue

End Property
Public Property Get Usuario() As UsuarioHPS
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjUsuario Is Nothing Then
        Set Usuario = m_ObjUsuario
        Exit Property
    End If
    If Me.idUsuario = "" Then
        Exit Property
    End If
    Set m_ObjUsuario = getUsuarioHPS(Me.idUsuario, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Usuario = m_ObjUsuario
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HPS.Usuario ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property



Public Property Get MESES_PARA_RENOVAR() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    Set m_ObjColDatosHPS = Nothing
    Set m_ObjColDatosHPS = getColDatosHPS(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sMESES_PARA_RENOVAR = DatosHPSCalculados("MESES_PARA_RENOVAR", Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If IsNumeric(m_sMESES_PARA_RENOVAR) Then
        MESES_PARA_RENOVAR = m_sMESES_PARA_RENOVAR
    End If
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HPS.MESES_PARA_RENOVAR ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property


Public Property Let MESES_PARA_RENOVAR(ByVal eNewValue As String)

    m_sMESES_PARA_RENOVAR = eNewValue

End Property

Public Function MotivoNoOK( _
                                Optional ByRef p_Error As String _
                                ) As String
    Dim m_MotivoIrregular As String
    On Error GoTo errores
        
    With Me
        If .ParaAlta = Empty Then
            MotivoNoOK = "No se sabe si es alta o edición"
            Exit Function
            
        End If
        
        If .Usuario Is Nothing Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "No se ha podido conocer el usuario del HPS"
            Exit Function
        Else
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If .tipoHps = "" Then
            MotivoNoOK = "El TipoHPS es obligatorio"
            Exit Function
        End If
        If .F_Concesion = "" And .F_Caducidad = "" And _
            .Grado = "" And .F_Solicitud = "" And _
            .Renovacion = "" And .F_Baja = "" And _
            .tipoHps = "" Then
            MotivoNoOK = "No hay datos"
            Exit Function
        End If
        m_MotivoIrregular = .MOTIVO_IRREGULAR
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_MotivoIrregular <> "" Then
            MotivoNoOK = "Los datos son irregulares " & vbNewLine & m_MotivoIrregular
            Exit Function
        End If
        
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function



Public Function Registrar( _
                            Optional ByRef p_Error As String _
                            ) As String
                                
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim m_TipoHPS As String
    Dim m_FechaBajaCalculada As String
    On Error GoTo errores
     Me.Error = ""
    p_Error = ""
    
    If Me.ParaAlta = 0 Then
        p_Error = "No se sabe si es alta o edición"
        Err.Raise 1000
    End If
    m_MotivoNoOK = MotivoNoOK(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.Renovacion = "X" Then
        Me.Renovacion = "No"
    End If
    If Me.ParaAlta = EnumSiNo.Sí Then
        m_SQL = "TbHPS"
    Else
        m_SQL = "SELECT * " & _
                " FROM TbHPS " & _
                "WHERE IDUsuario=" & Me.idUsuario & _
                " AND TipoHPS='" & Me.tipoHps & "';"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Me.ParaAlta = EnumSiNo.No Then
            If .EOF Then
                p_Error = "No se ha encontrado el registro"
                Err.Raise 1000
            End If
        End If
        If Me.ParaAlta = EnumSiNo.Sí Then
            .AddNew
                .Fields("IDUsuario") = Me.idUsuario
                .Fields("TipoHPS") = Me.tipoHps
                If IsDate(Me.F_Concesion) Then
                    .Fields("F_Concesion") = Me.F_Concesion
                End If
                If IsDate(Me.F_Caducidad) Then
                    .Fields("F_Caducidad") = Me.F_Caducidad
                End If
                .Fields("Grado") = Me.Grado
                If IsDate(Me.F_Baja) Then
                    .Fields("F_Baja") = Me.F_Baja
                End If
                If Me.Renovacion <> "" Then
                    .Fields("Renovacion") = Me.Renovacion
                End If
                If IsDate(Me.F_Solicitud) Then
                    .Fields("F_Solicitud") = Me.F_Solicitud
                End If
                If Me.Especialidad <> "" Then
                    .Fields("Especialidad") = Me.Especialidad
                End If
            .Update
        Else
            .Edit
                If IsDate(Me.F_Concesion) Then
                    .Fields("F_Concesion") = Me.F_Concesion
                Else
                    .Fields("F_Concesion") = Null
                End If
                If IsDate(Me.F_Caducidad) Then
                    .Fields("F_Caducidad") = Me.F_Caducidad
                Else
                    .Fields("F_Caducidad") = Null
                End If
                .Fields("Grado") = Me.Grado
                If IsDate(Me.F_Baja) Then
                    .Fields("F_Baja") = Me.F_Baja
                Else
                    .Fields("F_Baja") = Null
                End If
                If Renovacion <> "" Then
                    .Fields("Renovacion") = Me.Renovacion
                Else
                    .Fields("Renovacion") = Null
                End If
                If IsDate(Me.F_Solicitud) Then
                    .Fields("F_Solicitud") = Me.F_Solicitud
                Else
                    .Fields("F_Solicitud") = Null
                End If
                If Me.Especialidad <> "" Then
                    .Fields("Especialidad") = Me.Especialidad
                Else
                    .Fields("Especialidad") = Null
                End If
            .Update
        
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_FechaBajaCalculada = Me.Usuario.F_BajaCalculada
    If Me.Baja = EnumSiNo.Sí Then
        If IsDate(m_FechaBajaCalculada) Then
            m_SQL = "UPDATE TbUsuarios SET TbUsuarios.F_Baja = #" & Format(m_FechaBajaCalculada, "mm/dd/yyyy") & "# " & _
                    "WHERE (((TbUsuarios.ID)=" & Me.Usuario.ID & "));"
            getdb().Execute m_SQL
            
        End If
        
    End If
    Me.Usuario.FechaHPSConcesionMinimaRegistrar
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HPS.Registrar ha devuelto el error: " & Err.Description
    End If
End Function
Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Application.TempVars("SoloLectura") = "Sí" Then
        p_Error = "El Sistema está en modo sólo lectura"
        Err.Raise 1000
    End If
    If Me.idUsuario = "" Or Me.tipoHps = "" Then
        p_Error = "No tenemos datos de IDUsuario ni de tipo de hps"
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HPS.ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
                                 
    Dim m_SQL As String
    Dim m_IDUsuario As String
    On Error GoTo errores
    p_Error = ""
    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_IDUsuario = Me.idUsuario
    m_SQL = "DELETE * " & _
            "FROM TbHPS " & _
            "WHERE IDUsuario=" & Me.idUsuario & " AND TipoHPS='" & Me.tipoHps & "';"
    getdb().Execute m_SQL
    RefreshHpsUserCacheAfterMutation p_IDUsuario:=m_IDUsuario, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioActivo = Constructor.getUsuarioHPS(m_IDUsuario, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Borrar ha devuelto el error: " & Err.Description
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
        .Add "IDUsuario"
        .Add "TipoHPS"
        .Add "F_Concesion"
        .Add "F_Caducidad"
        .Add "Grado"
        .Add "F_Baja"
        .Add "Renovacion"
        .Add "F_Solicitud"
        .Add "Especialidad"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HPS.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDUsuario" Then
        getPropiedad = Me.idUsuario
        Exit Function
    End If
    If m_NombrePropiedad = "TipoHPS" Then
        getPropiedad = Me.tipoHps
        Exit Function
    End If
    If m_NombrePropiedad = "F_Concesion" Then
        getPropiedad = Me.F_Concesion
        Exit Function
    End If
    If m_NombrePropiedad = "F_Caducidad" Then
        getPropiedad = Me.F_Caducidad
        Exit Function
    End If
    If m_NombrePropiedad = "Grado" Then
        getPropiedad = Me.Grado
        Exit Function
    End If
    If m_NombrePropiedad = "F_Baja" Then
        getPropiedad = Me.F_Baja
        Exit Function
    End If
    If m_NombrePropiedad = "Renovacion" Then
        getPropiedad = Me.Renovacion
        Exit Function
    End If
    If m_NombrePropiedad = "F_Solicitud" Then
        getPropiedad = Me.F_Solicitud
        Exit Function
    End If
    If m_NombrePropiedad = "Especialidad" Then
        getPropiedad = Me.Especialidad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en HPS.getPropiedad"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HPS.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDUsuario" Then
        Me.idUsuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TipoHPS" Then
        Me.tipoHps = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "F_Concesion" Then
        Me.F_Concesion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "F_Caducidad" Then
        Me.F_Caducidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Grado" Then
        Me.Grado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "F_Baja" Then
        Me.F_Baja = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Renovacion" Then
        Me.Renovacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "F_Solicitud" Then
        Me.F_Solicitud = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Especialidad" Then
        Me.Especialidad = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en HPS.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HPS.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function








