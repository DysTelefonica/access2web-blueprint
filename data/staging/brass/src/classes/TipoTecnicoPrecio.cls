VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TipoTecnicoPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDTipoTecnicoFecha As String
Public TIPOTECNICO As String
Public FechaInicial As String
Public FechaFinal As String
Public PrecioHoraLab As String
Public PrecioHoraExtra As String
Public ParaAlta As EnumSino
Private m_ObjTipoTecnicoPrecioAnterior As TipoTecnicoPrecio
Private m_objTipoTecnicoPrecioSiguiente As TipoTecnicoPrecio
Private m_ObjTipoTecnicoPrecioUltimo As TipoTecnicoPrecio
Private m_sIDTipoTecnicoFechaCalculada As String
Private m_eTipoActivo As EnumSino
Private m_eEsUltima As EnumSino
Private m_objTIPOTECNICOObj As Database.TIPOTECNICO
Private m_objColTecnicos As Scripting.Dictionary
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get TipoActivo() As EnumSino
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_eTipoActivo <> 0 Then
        TipoActivo = m_eTipoActivo
        Exit Property
    End If
    If Not Me.TipoTecnicoPrecioUltimo Is Nothing Then
        If Me.TipoTecnicoPrecioUltimo.IDTipoTecnicoFecha = Me.IDTipoTecnicoFecha Then
            m_eTipoActivo = EnumSino.Sí
        Else
            m_eTipoActivo = EnumSino.No
        End If
    Else
        m_eTipoActivo = EnumSino.No
    End If
    
    
    TipoActivo = m_eTipoActivo
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnicoPrecio.TipoActivo ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get TIPOTECNICOObj() As Database.TIPOTECNICO
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objTIPOTECNICOObj Is Nothing Then
        Set TIPOTECNICOObj = m_objTIPOTECNICOObj
        Exit Property
    End If
    If Me.TIPOTECNICO = "" Then
        Exit Property
    End If
    Set m_objTIPOTECNICOObj = Constructor.getTipoTecnico(Me.TIPOTECNICO, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set TIPOTECNICOObj = m_objTIPOTECNICOObj
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnicoPrecio.TIPOTECNICOObj ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set TIPOTECNICOObj(ByVal objNewValue As Variant)

    Set m_objTIPOTECNICOObj = objNewValue

End Property

Public Property Get IDTipoTecnicoFechaCalculada() As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    m_sIDTipoTecnicoFechaCalculada = DameID("TbTipoTecnicoPrecios", "IDTipoTecnicoFecha")
    
    IDTipoTecnicoFechaCalculada = m_sIDTipoTecnicoFechaCalculada
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnicoPrecio.IDTipoTecnicoFechaCalculada ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get TipoTecnicoPrecioAnterior() As TipoTecnicoPrecio

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_ObjTipoTecnicoPrecioAnterior Is Nothing Then
        Set TipoTecnicoPrecioAnterior = m_ObjTipoTecnicoPrecioAnterior
        Exit Property
    End If
    If Me.IDTipoTecnicoFecha = "" Then
        Exit Property
    End If
    If Me.TIPOTECNICO = "" Then
        Exit Property
    End If
    Set m_ObjTipoTecnicoPrecioAnterior = Constructor.getTipoTecnicoPrecioAntPost(Me.IDTipoTecnicoFecha, Me.TIPOTECNICO, EnumSino.Sí, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set TipoTecnicoPrecioAnterior = m_ObjTipoTecnicoPrecioAnterior
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnicoPrecio.TipoTecnicoPrecioAnterior ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property
Public Property Set TipoTecnicoPrecioAnterior(p_Nuevo As TipoTecnicoPrecio)

    Set m_ObjTipoTecnicoPrecioAnterior = p_Nuevo
End Property
Public Property Get TipoTecnicoPrecioSiguiente() As TipoTecnicoPrecio

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objTipoTecnicoPrecioSiguiente Is Nothing Then
        Set TipoTecnicoPrecioSiguiente = m_objTipoTecnicoPrecioSiguiente
        Exit Property
    End If
    If Me.IDTipoTecnicoFecha = "" Then
        Exit Property
    End If
    If Me.TIPOTECNICO = "" Then
        Exit Property
    End If
    Set m_objTipoTecnicoPrecioSiguiente = Constructor.getTipoTecnicoPrecioAntPost(Me.IDTipoTecnicoFecha, _
                Me.TIPOTECNICO, EnumSino.No, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set TipoTecnicoPrecioSiguiente = m_objTipoTecnicoPrecioSiguiente
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnicoPrecio.TipoTecnicoPrecioSiguiente ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property
Public Property Set TipoTecnicoPrecioSiguiente(p_Nuevo As TipoTecnicoPrecio)

    Set m_objTipoTecnicoPrecioSiguiente = p_Nuevo
End Property
Public Property Get TipoTecnicoPrecioUltimo() As TipoTecnicoPrecio

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_ObjTipoTecnicoPrecioUltimo Is Nothing Then
        Set TipoTecnicoPrecioUltimo = m_ObjTipoTecnicoPrecioUltimo
        Exit Property
    End If
    If Me.IDTipoTecnicoFecha = "" Then
        Exit Property
    End If
    If Me.TIPOTECNICO = "" Then
        Exit Property
    End If
    Set m_ObjTipoTecnicoPrecioUltimo = Constructor.getTipoTecnicoPrecioUltimo(, _
                Me.TIPOTECNICO, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set TipoTecnicoPrecioUltimo = m_ObjTipoTecnicoPrecioUltimo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnicoPrecio.TipoTecnicoPrecioUltimo ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property
Public Property Set TipoTecnicoPrecioUltimo(p_Nuevo As TipoTecnicoPrecio)

    Set m_ObjTipoTecnicoPrecioUltimo = p_Nuevo
End Property
Public Property Get EsUltima() As EnumSino
    
    Dim m_ObjTipoTecnicoPrecioUltimo As TipoTecnicoPrecio
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_eEsUltima <> 0 Then
        EsUltima = m_eEsUltima
        Exit Property
    End If
    Set m_ObjTipoTecnicoPrecioUltimo = Constructor.getTipoTecnicoPrecioUltimo(, Me.TIPOTECNICO, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjTipoTecnicoPrecioUltimo Is Nothing Then
        m_Error = "No se encuentra el últipo precio de este tipo"
        Err.Raise 1000
    End If
    If m_ObjTipoTecnicoPrecioUltimo.IDTipoTecnicoFecha = Me.IDTipoTecnicoFecha Then
        m_eEsUltima = EnumSino.Sí
    Else
        m_eEsUltima = EnumSino.No
    End If
    EsUltima = m_eEsUltima
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnicoPrecio.EsUltima ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get ColTecnicos() As Scripting.Dictionary
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If Not m_objColTecnicos Is Nothing Then
        Set ColTecnicos = m_objColTecnicos
        Exit Property
    End If
    Set m_objColTecnicos = Constructor.getTecnicosPorTipo(Me.TIPOTECNICO, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColTecnicos = m_objColTecnicos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnicoPrecio.ColTecnicos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function MotivoNoOK( _
                                Optional ByRef p_ObjTipoTecnicoPrecioAlInicio As TipoTecnicoPrecio, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_ObjTipoPrecioActivo As TipoTecnicoPrecio
    Dim m_TipoEnAlgunTecnicoPorFacturar As EnumSino
    On Error GoTo errores
        
    With Me
        If .ParaAlta = 0 Then
            MotivoNoOK = "No se sabe si es alta o edición"
            Exit Function
        End If
        If .TIPOTECNICOObj Is Nothing Then
            MotivoNoOK = "Obligatorio el Tipo Técnico"
            Exit Function
        End If
        If Not p_ObjTipoTecnicoPrecioAlInicio Is Nothing Then
            If Me.TipoActivo = EnumSino.No Then
                MotivoNoOK = "No se puede cambiar los datos de un tipo no activo"
                Exit Function
            End If
        End If
        If Not IsNumeric(.PrecioHoraLab) Then
            MotivoNoOK = "Se ha de indicar un precio de hora laboral"
            Exit Function
        End If
        If Not IsNumeric(.PrecioHoraExtra) Then
            MotivoNoOK = "Se ha de indicar un precio de hora Extra"
            Exit Function
        End If
        If .ParaAlta = EnumSino.Sí Then
            If IsDate(.FechaFinal) Then
                MotivoNoOK = "En el alta no se pude poner fecha final"
                Exit Function
            End If
            
        Else
            If p_ObjTipoTecnicoPrecioAlInicio.FechaFinal = "" And Me.FechaFinal <> "" Then
                m_TipoEnAlgunTecnicoPorFacturar = TipoEnAlgunTecnicoPorFacturar(Me.TIPOTECNICO, p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If m_TipoEnAlgunTecnicoPorFacturar = EnumSino.Sí Then
                    MotivoNoOK = "Si cierra la fecha de este tipo , no se van a poder facturar algunos elementos que tienen este perfil"
                    Exit Function
                End If
            End If
        End If
        
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método TipoTecnicoPrecio.MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function

Private Function TipoEnAlgunTecnicoPorFacturar( _
                                                p_Tipo As String, _
                                                Optional ByRef p_Error As String _
                                                ) As EnumSino
            
    Dim rcdDatos As DAO.Recordset
    
    On Error GoTo errores
        
    m_SQL = "SELECT Count(TbTecnicos.ALIAS) AS CuentaDeALIAS " & _
            "FROM (TbTecnicos INNER JOIN TbActividades ON TbTecnicos.ALIAS = TbActividades.ALIASTECNICO) " & _
            "LEFT JOIN TbFacturaActividadesInvolucradas ON " & _
            "TbActividades.IDActividad = TbFacturaActividadesInvolucradas.IDActividad " & _
            "WHERE (((TbTecnicos.TIPO)='" & p_Tipo & "') AND ((TbFacturaActividadesInvolucradas.IDActividad) Is Null));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If Not rcdDatos Is Nothing Then
        If IsNumeric(Nz(rcdDatos.Fields("CuentaDeALIAS"), "")) Then
            If CLng(rcdDatos.Fields("CuentaDeALIAS")) > 0 Then
                TipoEnAlgunTecnicoPorFacturar = EnumSino.Sí
                rcdDatos.Close
                Set rcdDatos = Nothing
            End If
        End If
    End If
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    m_SQL = "SELECT Count(TbTecnicos.ALIAS) AS CuentaDeALIAS " & _
            "FROM TbTecnicos INNER JOIN (TbEventos LEFT JOIN TbFacturaEventosInvolucrados ON " & _
            "TbEventos.IDEvento = TbFacturaEventosInvolucrados.IDEvento) ON TbTecnicos.ALIAS = TbEventos.ALIASTECNICO " & _
            "WHERE (((TbTecnicos.TIPO)='" & p_Tipo & "') AND ((TbFacturaEventosInvolucrados.IDEvento) Is Null));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If Not rcdDatos Is Nothing Then
        If IsNumeric(Nz(rcdDatos.Fields("CuentaDeALIAS"), "")) Then
            If CLng(rcdDatos.Fields("CuentaDeALIAS")) > 0 Then
                TipoEnAlgunTecnicoPorFacturar = EnumSino.Sí
                rcdDatos.Close
                Set rcdDatos = Nothing
            End If
        End If
    End If
    rcdDatos.Close
    Set rcdDatos = Nothing
    TipoEnAlgunTecnicoPorFacturar = EnumSino.No
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método TipoTecnicoPrecio.TipoEnAlgunTecnicoPorFacturar ha devuelto el error: " & Err.Description
     End If
                                
End Function


Public Function Registrar( _
                            Optional ByRef p_ObjTipoTecnicoPrecioAlInicio As TipoTecnicoPrecio, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim m_ObjTipoPrecioActivo As TipoTecnicoPrecio
    Dim blnCerradoPrevio As Boolean
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If Me.ParaAlta = 0 Then
        p_Error = "No se sabe si es alta o edición"
        Err.Raise 1000
    End If
    
    m_MotivoNoOK = MotivoNoOK(p_ObjTipoTecnicoPrecioAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjTipoPrecioActivo = Constructor.getTipoTecnicoActivo(Me.TIPOTECNICO, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjTipoPrecioActivo Is Nothing Then
        m_SQL = "UPDATE TbTipoTecnicoPrecios SET TbTipoTecnicoPrecios.FechaFinal =#" & Format(Date, "mm/dd/yyyy") & _
                "# WHERE IDTipoTecnicoFecha=" & m_ObjTipoPrecioActivo.IDTipoTecnicoFecha & ";"
        getdb().Execute m_SQL
        blnCerradoPrevio = True
    End If
    
    
    If Me.ParaAlta = EnumSino.Sí Then
        m_SQL = "TbTipoTecnicoPrecios"
        If Me.IDTipoTecnicoFecha = "" Then
            Me.IDTipoTecnicoFecha = Me.IDTipoTecnicoFechaCalculada
            m_Error = Me.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            If Me.IDTipoTecnicoFecha = "" Then
                m_Error = "No se ha podido obtener el ID"
                Err.Raise 1000
            End If
        End If
    Else
        
        m_SQL = "SELECT * FROM TbTipoTecnicoPrecios " & _
            "WHERE IDTipoTecnicoFecha=" & p_ObjTipoTecnicoPrecioAlInicio.IDTipoTecnicoFecha & ";"
        
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
                .Fields("IDTipoTecnicoFecha") = Me.IDTipoTecnicoFecha
                .Fields("TIPOTECNICO") = Me.TIPOTECNICO
                .Fields("FechaInicial") = Me.FechaInicial
                If Me.FechaFinal <> "" Then
                    .Fields("FechaFinal") = Me.FechaFinal
                End If
                
                .Fields("PrecioHoraLab") = Me.PrecioHoraLab
                
                
                    .Fields("PrecioHoraExtra") = Me.PrecioHoraExtra
                
                
            .Update
        Else
            .Edit
                .Fields("TIPOTECNICO") = Me.TIPOTECNICO
                .Fields("FechaInicial") = Me.FechaInicial
                If Me.FechaFinal <> "" Then
                    .Fields("FechaFinal") = Me.FechaFinal
                Else
                    .Fields("FechaFinal") = Null
                End If
                
                    .Fields("PrecioHoraLab") = Me.PrecioHoraLab
                
                
                    .Fields("PrecioHoraExtra") = Me.PrecioHoraExtra
                
            .Update
        
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If blnCerradoPrevio Then
        m_SQL = "UPDATE TbTipoTecnicoPrecios SET TbTipoTecnicoPrecios.FechaFinal =Null " & _
                "WHERE IDTipoTecnicoFecha=" & m_ObjTipoPrecioActivo.IDTipoTecnicoFecha & ";"
        getdb().Execute m_SQL
    End If
    If Err.Number <> 1000 Then
        p_Error = "El método TipoTecnicoPrecio.Registrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Cerrar( _
                            Optional ByRef p_FechaCierre As String, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    
    If Not IsDate(p_FechaCierre) Then
        p_FechaCierre = Date
    End If
    m_SQL = "UPDATE TbTipoTecnicoPrecios SET TbTipoTecnicoPrecios.FechaFinal = #" & Format(p_FechaCierre, "mm/dd/yyyy") & "# " & _
            "WHERE IDTipoTecnicoFecha='" & Me.IDTipoTecnicoFecha & "';"
    getdb().Execute m_SQL
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TipoTecnicoPrecio.Cerrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Abrir( _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    
    
    m_SQL = "UPDATE TbTipoTecnicoPrecios SET TbTipoTecnicoPrecios.FechaFinal = Null " & _
            "WHERE IDTipoTecnicoFecha=" & Me.IDTipoTecnicoFecha & ";"
    getdb().Execute m_SQL
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TipoTecnicoPrecio.Abrir ha devuelto el error: " & Err.Description
    End If
End Function
Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
    
    'solo se puede eliminar la ultima
    
    Dim m_ObjTipoTecnicoPrecioAnterior As TipoTecnicoPrecio
    Dim m_EsUltima As EnumSino
    Dim m_TipoTecnicoPrecioEnActividades As EnumSino
    Dim m_TipoTecnicoPrecioEnEventos As EnumSino
    
    On Error GoTo errores
    
    Me.Error = ""
    
    
    m_EsUltima = Me.EsUltima
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EsUltima <> EnumSino.Sí Then
        p_Error = "Sólo se puede eliminar la más reciente de cada tipo"
        Err.Raise 1000
    End If
    
    If Not Me.ColTecnicos Is Nothing Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "Al menos hay un técnico que tiene este perfil"
        Err.Raise 1000
    End If

    
    Set m_ObjTipoTecnicoPrecioAnterior = Me.TipoTecnicoPrecioAnterior
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjTipoTecnicoPrecioAnterior Is Nothing Then
        Exit Function
    End If
    m_TipoTecnicoPrecioEnEventos = TipoTecnicoPrecioEnEventos(, Me.TipoActivo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoTecnicoPrecioEnEventos = EnumSino.Sí Then
        p_Error = "No quedaría ningún precio de este perfil y hay al menos un evento sin facturar"
        Err.Raise 1000
    End If
    
    
    m_TipoTecnicoPrecioEnActividades = TipoTecnicoPrecioEnActividades(, Me.TipoActivo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoTecnicoPrecioEnActividades = EnumSino.Sí Then
        p_Error = "No quedaría ningún precio de este perfil y hay al menos una actividad sin facturar"
        Err.Raise 1000
    End If
    
        
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TipoTecnicoPrecio.ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
End Function


Private Function TipoTecnicoPrecioEnActividades( _
                                                Optional p_Alias As String, _
                                                Optional p_Tipo As String, _
                                                Optional ByRef p_Error As String _
                                                ) As EnumSino
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If p_Alias = "" And p_Tipo = "" Then
        Exit Function
    End If
    If p_Alias <> "" Then
         m_SQL = "SELECT Count(TbActividades.IDActividad) AS CuentaDeIDActividad " & _
                "FROM TbActividades LEFT JOIN TbFacturaActividadesInvolucradas ON TbActividades.IDActividad = TbFacturaActividadesInvolucradas.IDActividad " & _
                "WHERE ALIASTECNICO='" & p_Alias & "' AND TbFacturaActividadesInvolucradas.IDActividad Is Null;"
    Else
         m_SQL = "SELECT Count(TbActividades.IDActividad) AS CuentaDeIDActividad " & _
                "FROM (TbActividades LEFT JOIN TbFacturaActividadesInvolucradas ON TbActividades.IDActividad = TbFacturaActividadesInvolucradas.IDActividad) INNER JOIN TbTecnicos ON TbActividades.ALIASTECNICO = TbTecnicos.ALIAS " & _
                "WHERE TIPO='" & p_Tipo & "' AND TbFacturaActividadesInvolucradas.IDActividad Is Null;"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            TipoTecnicoPrecioEnActividades = EnumSino.No
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        If IsNumeric(Nz(.Fields("CuentaDeIDActividad"), "")) Then
            If CLng(.Fields("CuentaDeIDActividad")) > 0 Then
                TipoTecnicoPrecioEnActividades = EnumSino.Sí
            Else
                TipoTecnicoPrecioEnActividades = EnumSino.No
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método TipoTecnicoPrecioEnActividades ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Function TipoTecnicoPrecioEnEventos( _
                                                Optional p_Alias As String, _
                                                Optional p_Tipo As String, _
                                                Optional ByRef p_Error As String _
                                                ) As EnumSino
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If p_Alias = "" And p_Tipo = "" Then
        Exit Function
    End If
    If p_Tipo <> "" Then
        m_SQL = "SELECT Count(TbEventos.IDEvento) AS CuentaDeIDEvento " & _
                "FROM (TbEventos LEFT JOIN TbFacturaEventosInvolucrados ON " & _
                "TbEventos.IDEvento = TbFacturaEventosInvolucrados.IDEvento) INNER JOIN TbTecnicos ON " & _
                "TbEventos.ALIASTECNICO = TbTecnicos.ALIAS " & _
                "WHERE TIPO='" & p_Tipo & "' AND TbFacturaEventosInvolucrados.IDEvento Is Null;"
    
    Else
        m_SQL = "SELECT Count(TbEventos.IDEvento) AS CuentaDeIDEvento " & _
                "FROM TbEventos LEFT JOIN TbFacturaEventosInvolucrados ON TbEventos.IDEvento = TbFacturaEventosInvolucrados.IDEvento " & _
                "WHERE ALIASTECNICO='" & p_Alias & "' AND TbFacturaEventosInvolucrados.IDEvento Is Null;"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            TipoTecnicoPrecioEnEventos = EnumSino.No
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        If IsNumeric(Nz(.Fields("CuentaDeIDEvento"), "")) Then
            If CLng(.Fields("CuentaDeIDEvento")) > 0 Then
                TipoTecnicoPrecioEnEventos = EnumSino.Sí
            Else
                TipoTecnicoPrecioEnEventos = EnumSino.No
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método TipoTecnicoPrecioEnEventos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
                                
    Dim m_ObjTipoPrecioAnterior As TipoTecnicoPrecio
    On Error GoTo errores
    
    Me.Error = ""
    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjTipoPrecioAnterior = Nothing
    Set m_ObjTipoPrecioAnterior = Me.TipoTecnicoPrecioAnterior
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "DELETE * FROM TbTipoTecnicoPrecios " & _
            "WHERE IDTipoTecnicoFecha=" & Me.IDTipoTecnicoFecha & ";"
    getdb().Execute m_SQL
    If Not m_ObjTipoPrecioAnterior Is Nothing Then
        m_ObjTipoPrecioAnterior.Abrir p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Me.TIPOTECNICOObj.Borrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Tecnico.Borrar ha devuelto el error: " & Err.Description
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
        .Add "IDTipoTecnicoFecha"
        .Add "TIPOTECNICO"
        .Add "FechaInicial"
        .Add "FechaFinal"
        .Add "PrecioHoraLab"
        .Add "PrecioHoraExtra"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnicoPrecio.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDTipoTecnicoFecha" Then
        getPropiedad = Me.IDTipoTecnicoFecha
        Exit Function
    End If
    If m_NombrePropiedad = "TIPOTECNICO" Then
        getPropiedad = Me.TIPOTECNICO
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicial" Then
        getPropiedad = Me.FechaInicial
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinal" Then
        getPropiedad = Me.FechaFinal
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioHoraLab" Then
        getPropiedad = Me.PrecioHoraLab
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioHoraExtra" Then
        getPropiedad = Me.PrecioHoraExtra
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en TipoTecnicoPrecio.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TipoTecnicoPrecio.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDTipoTecnicoFecha" Then
        Me.IDTipoTecnicoFecha = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TIPOTECNICO" Then
        Me.TIPOTECNICO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicial" Then
        Me.FechaInicial = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinal" Then
        Me.FechaFinal = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioHoraLab" Then
        Me.PrecioHoraLab = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PrecioHoraExtra" Then
        Me.PrecioHoraExtra = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en TipoTecnicoPrecio.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TipoTecnicoPrecio.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


