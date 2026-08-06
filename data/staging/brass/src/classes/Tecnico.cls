VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tecnico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Alias As String
Public Tipo As String
Public Nombre As String
Public Telefono As String
Public Email As String
Public FechaAlta As String
Public FechaBaja As String
Public DNI As String

Public ParaAlta As EnumSino
Private m_eConPrecioActivo As EnumSino
Private m_ObjTipoTecnicoPrecio As TipoTecnicoPrecio
Private m_objColLibranzas As Scripting.Dictionary
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get TipoTecnicoPrecio() As TipoTecnicoPrecio
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If Not m_ObjTipoTecnicoPrecio Is Nothing Then
        Set TipoTecnicoPrecio = m_ObjTipoTecnicoPrecio
        Exit Property
    End If
    
    If Me.Alias = "" Then
        Exit Property
    End If
    Set m_ObjTipoTecnicoPrecio = Constructor.getTipoTecnicoPrecio(, Me.Alias, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set TipoTecnicoPrecio = m_ObjTipoTecnicoPrecio
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tecnico.TipoTecnicoPrecio ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get ConPrecioActivo() As EnumSino
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_eConPrecioActivo <> 0 Then
        ConPrecioActivo = m_eConPrecioActivo
        Exit Property
    End If
    If Me.TipoTecnicoPrecio Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eConPrecioActivo = EnumSino.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eConPrecioActivo = EnumSino.Sí
    End If
    ConPrecioActivo = m_eConPrecioActivo
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tecnico.ConPrecioActivo ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get ColLibranzas() As Scripting.Dictionary

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If Not m_objColLibranzas Is Nothing Then
        Set ColLibranzas = m_objColLibranzas
        Exit Property
    End If
    
    If Me.Alias = "" Then
        Exit Property
    End If
    Set m_objColLibranzas = Constructor.getLibranzas(Me.Alias, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColLibranzas = m_objColLibranzas
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tecnico.ColLibranzas ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Set ColLibranzas(ByVal objNewValue As Variant)

    Set m_objColLibranzas = objNewValue

End Property
Public Function MotivoNoOK( _
                                Optional ByRef p_ObjTecnicoAlInicio As Tecnico, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_ValorExistente As EnumSino
    On Error GoTo errores
        
    With Me
        If .ParaAlta = 0 Then
            MotivoNoOK = "No se sabe si es alta o edición"
            Exit Function
            
        End If
        
        If .Alias = "" Then
            MotivoNoOK = "Obligatorio el Alias"
            Exit Function
        End If
        If .Tipo = "" Then
            MotivoNoOK = "Obligatorio el Tipo"
            Exit Function
        End If
        If .Nombre = "" Then
            MotivoNoOK = "Obligatorio el Nombre"
            Exit Function
        End If
        If .FechaAlta = "" Then
            MotivoNoOK = "Obligatorio la fecha de alta"
            Exit Function
        End If
        If Not p_ObjTecnicoAlInicio Is Nothing Then
            If p_ObjTecnicoAlInicio.Alias <> Me.Alias Then
                m_ValorExistente = ValorExistente("TbTecnicos", "Alias", Me.Alias, EnumSino.No, getdb, p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If m_ValorExistente = EnumSino.Sí Then
                    MotivoNoOK = "El alias ya ha sido usado para otro técnico"
                    Exit Function
                End If
            End If
            If p_ObjTecnicoAlInicio.FechaBaja = "" And Me.FechaBaja <> "" Then
                If CDate(Me.FechaBaja) < CDate(Me.FechaAlta) Then
                    MotivoNoOK = "La fecha de baja ha de ser posterior a la de alta"
                    Exit Function
                End If
            End If
        Else
            m_ValorExistente = ValorExistente("TbTecnicos", "Alias", Me.Alias, EnumSino.No, getdb, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ValorExistente = EnumSino.Sí Then
                MotivoNoOK = "El alias ya ha sido usado para otro técnico"
                Exit Function
            End If
            If IsDate(.FechaBaja) Then
                If CDate(Me.FechaBaja) < CDate(Me.FechaAlta) Then
                    MotivoNoOK = "La fecha de baja ha de ser posterior a la de alta"
                    Exit Function
                End If
            End If
        End If
        m_ValorExistente = ValorExistente("TbTipoTecnico", "Tipo", Me.Tipo, EnumSino.No, getdb, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ValorExistente = EnumSino.No Then
            MotivoNoOK = "El Tipo de técnico no ha sido registrado previamente"
            Exit Function
        End If
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método Tecnico.MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function



Public Function Registrar( _
                            Optional ByRef p_ObjTecnicoAlInicio As Tecnico, _
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
    If Me.ParaAlta = EnumSino.No Then
        If p_ObjTecnicoAlInicio Is Nothing Then
            Set p_ObjTecnicoAlInicio = Constructor.getTecnico(Me.Alias, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    m_MotivoNoOK = MotivoNoOK(p_ObjTecnicoAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.ParaAlta = EnumSino.Sí Then
        m_SQL = "TbTecnicos"
        
    Else
        m_SQL = "SELECT * FROM TbTecnicos " & _
            "WHERE Alias='" & p_ObjTecnicoAlInicio.Alias & "';"
        
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
                .Fields("ALIAS") = Me.Alias
                .Fields("TIPO") = Me.Tipo
                .Fields("NOMBRE") = Me.Nombre
                If Me.Telefono <> "" Then
                    .Fields("Telefono") = Me.Telefono
                End If
                
                If Me.Email <> "" Then
                    .Fields("Email") = Me.Email
                End If
                If Me.FechaAlta <> "" Then
                    .Fields("FechaAlta") = Me.FechaAlta
                End If
                If Me.FechaBaja <> "" Then
                    .Fields("FechaBaja") = Me.FechaBaja
                End If
                If Me.DNI <> "" Then
                    .Fields("DNI") = Me.DNI
                End If
            .Update
        Else
            .Edit
                .Fields("TIPO") = Me.Tipo
                .Fields("NOMBRE") = Me.Nombre
                If Me.Telefono <> "" Then
                    .Fields("Telefono") = Me.Telefono
                End If
                
                If Me.Email <> "" Then
                    .Fields("Email") = Me.Email
                End If
                If Me.FechaAlta <> "" Then
                    .Fields("FechaAlta") = Me.FechaAlta
                End If
                If Me.FechaBaja <> "" Then
                    .Fields("FechaBaja") = Me.FechaBaja
                End If
                If Me.DNI <> "" Then
                    .Fields("DNI") = Me.DNI
                End If
            .Update
        
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Tecnico.Registrar ha devuelto el error: " & Err.Description
    End If
End Function


Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
    '       1)- Sólo puede eliminar si no ha pertenecido a ninguna actividad, evento material
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_SQL = "SELECT Count(TbEventos.ALIASTECNICO) AS Cuenta " & _
            "FROM TbEventos " & _
            "WHERE (((TbEventos.ALIASTECNICO)='" & Me.Alias & "'));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If IsNumeric(rcdDatos.Fields("Cuenta")) Then
        If rcdDatos.Fields("Cuenta") > 0 Then
            p_Error = "Existe al menos un registro en la tabla Eventos en la que está el técnico"
            Err.Raise 1000
        End If
    End If
    
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "SELECT Count(TbActividades.ALIASTECNICO) AS Cuenta " & _
            "FROM TbActividades " & _
            "WHERE (((TbActividades.ALIASTECNICO)='" & Me.Alias & "'));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If IsNumeric(rcdDatos.Fields("Cuenta")) Then
        If rcdDatos.Fields("Cuenta") > 0 Then
            p_Error = "Existe al menos un registro en la tabla Actividades en la que está el técnico"
            Err.Raise 1000
        End If
    End If
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "SELECT Count(TbGastos.ALIAS) AS Cuenta " & _
            "FROM TbGastos " & _
            "WHERE (((TbGastos.ALIAS)='" & Me.Alias & "'));"
   
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If IsNumeric(rcdDatos.Fields("Cuenta")) Then
        If rcdDatos.Fields("Cuenta") > 0 Then
            p_Error = "Existe al menos un registro en la tabla Gastos en la que está el técnico"
            Err.Raise 1000
        End If
    End If
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Tecnico.ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
                                
    On Error GoTo errores
    
    Me.Error = ""
    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "DELETE * FROM TbTecnicos " & _
            "WHERE ALIAS='" & Me.Alias & "';"
    getdb().Execute m_SQL
   
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Tecnico.Borrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function AltaAusencias( _
                                p_ObjAusencia As LIbranza, _
                                p_ObjAusenciaAlInicio As LIbranza, _
                                p_FechaInicial As String, _
                                p_FechaFinal As String, _
                                Optional ByRef p_Error As String _
                                ) As String
        
   
    Dim m_Fecha As String
    Dim Dias As Long
    Dim i As Long
    Dim m_EsDiaLaborableSinLibranza As EnumSino
    Dim m_ID As String
    Dim int_DiasRegistrados As Long
    Dim m_ObjAusencia As LIbranza
    
    On Error GoTo errores
    If Me.Alias = "" Then
        p_Error = "No se ha indicado el alias"
        Err.Raise 1000
    End If
    If Not IsDate(p_FechaInicial) Then
        p_Error = "Se ha de indicar la fecha inicial"
        Err.Raise 1000
    End If
    p_FechaInicial = Format(p_FechaInicial, "dd/mm/yyyy")
    If Not IsDate(p_FechaFinal) Then
        p_FechaFinal = p_FechaInicial
    End If
    p_FechaFinal = Format(p_FechaFinal, "dd/mm/yyyy")
    If CDate(p_FechaFinal) < CDate(p_FechaInicial) Then
        p_Error = "La fecha final ha de ser posterior a la inicial"
        Err.Raise 1000
    End If
    If p_ObjAusencia.MOTIVO = "" Then
        p_Error = "El motivo es obligatorio"
        Err.Raise 1000
    End If
    Dias = DateDiff("d", CDate(p_FechaFinal), CDate(p_FechaInicial))
    If Dias < 1 Then
        Dias = Dias * -1
    End If
    For i = 0 To Dias
        Set m_ObjAusencia = New LIbranza
        If Not p_ObjAusenciaAlInicio Is Nothing Then
            m_ObjAusencia.ParaAlta = EnumSino.No
        Else
            m_ObjAusencia.ParaAlta = EnumSino.Sí
        End If
        m_ObjAusencia.MOTIVO = p_ObjAusencia.MOTIVO
        m_ObjAusencia.Observaciones = p_ObjAusencia.Observaciones
        
        m_ObjAusencia.Alias = p_ObjAusencia.Alias
        m_Fecha = DateAdd("d", CDbl(i), CDate(p_FechaInicial))
        m_EsDiaLaborableSinLibranza = EsDiaLaborable(m_Fecha, EnumSino.No, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_EsDiaLaborableSinLibranza = EnumSino.No Then
            GoTo siguiente
        End If
        m_ObjAusencia.FechaLibranza = m_Fecha
        m_ObjAusencia.Registrar p_ObjAusenciaAlInicio, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        int_DiasRegistrados = int_DiasRegistrados + 1
        
        
siguiente:
    Set m_ObjAusencia = Nothing
    Next
    
    If int_DiasRegistrados = 1 Then
        AltaAusencias = "Ausencia registrada correctamente "
    ElseIf int_DiasRegistrados = 0 Then
        AltaAusencias = "Ausencia no registrada al no ser un día laborable o ya tener esa ausencia registrada"
    Else
        AltaAusencias = "De los " & Dias + 1 & _
                        " días en el periodo se ha registrado un total de " & int_DiasRegistrados & _
                        " el resto no eran días laborables o ya había una ausencia registrada"
    End If
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AltaAusencias ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Function EsDiaLaborable( _
                                m_Fecha As String, _
                                Optional p_ConsiderandoLibranza As EnumSino = EnumSino.Sí, _
                                Optional ByRef p_Error As String _
                                ) As EnumSino
    Dim rcdDatos As DAO.Recordset
    Dim Dia As Long
    
   
    
    On Error GoTo errores
    If Me.Alias = "" Then
        p_Error = "Falta el alias"
        Err.Raise 1000
    End If
    If Not IsDate(m_Fecha) Then
        p_Error = "Falta la fecha"
        Err.Raise 1000
    End If
    Dia = VBA.Weekday(m_Fecha, vbMonday)
    If Dia = 6 Or Dia = 7 Then
        EsDiaLaborable = EnumSino.No
        Exit Function
    End If
    If p_ConsiderandoLibranza = EnumSino.Sí Then
        m_SQL = "SELECT TbTecnicosAusencias.FechaLibranza " & _
                    "FROM TbTecnicosAusencias " & _
                    "WHERE (((TbTecnicosAusencias.Alias)='" & Me.Alias & "') " & _
                    "AND ((TbTecnicosAusencias.FechaLibranza)=#" & Format(m_Fecha, "mm/dd/yyyy") & "#));"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        With rcdDatos
            If Not .EOF Then
                rcdDatos.Close
                Set rcdDatos = Nothing
                EsDiaLaborable = EnumSino.No
                Exit Function
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    
    
    m_SQL = "SELECT TbTecnicosFiestas.FechaFiesta " & _
                "FROM TbTecnicosFiestas " & _
                "WHERE (((TbTecnicosFiestas.FechaFiesta)=#" & Format(m_Fecha, "mm/dd/yyyy") & "#));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            EsDiaLaborable = EnumSino.No
            Exit Function
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    EsDiaLaborable = EnumSino.Sí
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EsDiaLaborable ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
        .Add "ALIAS"
        .Add "TIPO"
        .Add "NOMBRE"
        .Add "TELEFONO"
        .Add "EMAIL"
        .Add "FECHAALTA"
        .Add "FECHABAJA"
        .Add "DNI"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tecnico.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "ALIAS" Then
        getPropiedad = Me.Alias
        Exit Function
    End If
    If m_NombrePropiedad = "TIPO" Then
        getPropiedad = Me.Tipo
        Exit Function
    End If
    If m_NombrePropiedad = "NOMBRE" Then
        getPropiedad = Me.Nombre
        Exit Function
    End If
    If m_NombrePropiedad = "TELEFONO" Then
        getPropiedad = Me.Telefono
        Exit Function
    End If
    If m_NombrePropiedad = "EMAIL" Then
        getPropiedad = Me.Email
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAALTA" Then
        getPropiedad = Me.FechaAlta
        Exit Function
    End If
    If m_NombrePropiedad = "FECHABAJA" Then
        getPropiedad = Me.FechaBaja
        Exit Function
    End If
    If m_NombrePropiedad = "DNI" Then
        getPropiedad = Me.DNI
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Tecnico.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Tecnico.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "ALIAS" Then
        Me.Alias = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TIPO" Then
        Me.Tipo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NOMBRE" Then
        Me.Nombre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TELEFONO" Then
        Me.Telefono = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EMAIL" Then
        Me.Email = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAALTA" Then
        Me.FechaAlta = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHABAJA" Then
        Me.FechaBaja = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DNI" Then
        Me.DNI = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Tecnico.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Tecnico.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


