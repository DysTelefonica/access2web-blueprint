VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDNoConformidad As String
Public Juridica As String
Public CodigoNoConformidad As String
Public EsNoConformidad As Boolean
Public Expediente As String
Public Proyecto As String
Public Vehiculo As String
Public Descripcion As String
Public Causa As String
Public EntidadResponsable As String
Public ResponsableTelefonica As String
Public FechaApertura As String
Public FechaCierre As String
Public FPREVCIERRE As String
Public Tipo As String
Public Notas As String
Public Borrado As Boolean
Public RequiereACR As Boolean
Public ACR As String
Public MotivoBorrado As String
Public RequiereControlEficacia As String
Public ControlEficacia As String
Public FechaControlEficacia As String
Public FechaPrevistaControlEficacia As String
Public ResponsableCalidad As String

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Private m_eEsNoConformidadCalculado As EnumSiNo
Private m_eBorradoCalculado As EnumSiNo
Private m_eRequiereACRCalculado As EnumSiNo
Private m_eRequiereControlEficaciaCalculado As EnumSiNo
Private m_sCodigoNoConformidadCalculado As String
Private m_sIDNoConformidadCalculado As String
Private m_sProyectoCalculado As String


Public Property Get EsNoConformidadCalculado() As EnumSiNo

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Me.EsNoConformidad = True Then
        m_eEsNoConformidadCalculado = EnumSiNo.Sí
    Else
        m_eEsNoConformidadCalculado = EnumSiNo.No
    End If
    EsNoConformidadCalculado = m_eEsNoConformidadCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método NC.EsNoConformidadCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let EsNoConformidadCalculado(ByVal eNewValue As EnumSiNo)

    m_eEsNoConformidadCalculado = eNewValue

End Property

Public Property Get BorradoCalculado() As EnumSiNo

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.Borrado = True Then
        m_eBorradoCalculado = EnumSiNo.Sí
    Else
        m_eBorradoCalculado = EnumSiNo.No
    End If
    BorradoCalculado = m_eBorradoCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método NC.BorradoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let BorradoCalculado(ByVal eNewValue As EnumSiNo)

    m_eBorradoCalculado = eNewValue

End Property

Public Property Get RequiereACRCalculado() As EnumSiNo

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.RequiereACR = True Then
        m_eRequiereACRCalculado = EnumSiNo.Sí
    Else
        m_eRequiereACRCalculado = EnumSiNo.No
    End If
    RequiereACRCalculado = m_eRequiereACRCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método NC.RequiereACRCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let RequiereACRCalculado(ByVal eNewValue As EnumSiNo)

    m_eRequiereACRCalculado = eNewValue

End Property

Public Property Get RequiereControlEficaciaCalculado() As EnumSiNo

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Me.RequiereControlEficacia = "Sí" Then
        m_eRequiereControlEficaciaCalculado = EnumSiNo.Sí
    ElseIf Me.RequiereControlEficacia = "No" Then
        m_eRequiereControlEficaciaCalculado = EnumSiNo.No
    End If
    RequiereControlEficaciaCalculado = m_eRequiereControlEficaciaCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método NC.RequiereControlEficaciaCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let RequiereControlEficaciaCalculado(ByVal eNewValue As EnumSiNo)

    m_eRequiereControlEficaciaCalculado = eNewValue

End Property


Public Property Get CodigoNoConformidadCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sCodigoNoConformidadCalculado = getCodigo(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CodigoNoConformidadCalculado = m_sCodigoNoConformidadCalculado
    
    RequiereControlEficaciaCalculado = m_eRequiereControlEficaciaCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método NC.CodigoNoConformidadCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get IDNoConformidadCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sIDNoConformidadCalculado = DameID("TbNoConformidades", "IDNoConformidad", getdb, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDNoConformidadCalculado = m_sIDNoConformidadCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método NC.IDNoConformidadCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get ProyectoCalculado() As String
    
    Dim dato
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.Proyecto = "" Then
        Exit Property
    End If
    If Me.Expediente = "" Then
        m_sProyectoCalculado = Me.Proyecto
        ProyectoCalculado = m_sProyectoCalculado
        Exit Property
    End If
    If InStr(1, Me.Proyecto, ".") = 0 Then
        m_sProyectoCalculado = Me.Expediente & "." & Me.Proyecto
    Else
        dato = Split(Me.Proyecto, ".")
        If Me.Expediente <> dato(0) Then
            m_sProyectoCalculado = Me.Expediente & "." & dato(1)
        Else
            m_sProyectoCalculado = Me.Proyecto
        End If
    End If
    ProyectoCalculado = m_sProyectoCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método NC.ProyectoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Function Registrar( _
                            Optional ByRef p_ObjNCAlInicio As NC, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim rcdDatos As DAO.Recordset
     Dim m_SQL As String
     Dim m_FechaRef As Date
     
    On Error GoTo errores
    
    Me.Error = ""
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    '-----------------------------------
    ' Comprobación de los datos
    '----------------------------------
    If Me.EsNoConformidadCalculado = 0 Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "No se sabe si es NC u Observación"
        Err.Raise 1000
    End If
    
    If Me.Juridica <> "TdE" And Me.Juridica <> "TSOL" Then
        p_Error = "La Jurídica ha de ser TdE o TSOL"
        Err.Raise 1000
    End If
    
    If Me.RequiereACRCalculado = 0 Then
        p_Error = "Se ha de indicar si requiere Análisis de la causa raíz"
        Err.Raise 1000
    End If
    
    If Me.RequiereControlEficaciaCalculado = 0 Then
        p_Error = "Se ha de indicar si requiere control de eficacia"
        Err.Raise 1000
    End If
    If Me.RequiereControlEficaciaCalculado = EnumSiNo.Sí Then
        
        If Not IsDate(Me.FechaPrevistaControlEficacia) Then
            p_Error = "Al requerir control de eficacia, se ha de indicar una fecha prevista del mismo"
            Err.Raise 1000
        Else
            If CDate(Me.FechaPrevistaControlEficacia) < m_FechaRef Then
                p_Error = "La fecha prevista del control de eficacia no puede ser anterior al día de hoy"
                Err.Raise 1000
            End If
        End If
    End If
    
    If Me.Descripcion = "" Then
        p_Error = "El campo Descripción es obligatorio"
        Err.Raise 1000
    End If
    If Me.Causa = "" Then
        p_Error = "El campo Causa raíz es obligatorio"
        Err.Raise 1000
    End If
  
    If Me.ACR = "" Then
        p_Error = "El campo Análisis Causa raíz es obligatorio"
        Err.Raise 1000
    End If
    If Me.Expediente = "" Then
        p_Error = "El campo Expediente es obligatorio"
        Err.Raise 1000
    End If
    If Me.Tipo = "" Then
        p_Error = "El campo Tipo es obligatorio"
        Err.Raise 1000
    End If
    If Me.EntidadResponsable = "" Then
        p_Error = "El campo Entidad responsable es obligatorio"
        Err.Raise 1000
    End If
    If Not IsDate(Me.FechaApertura) Then
        Me.FechaApertura = Now()
    End If
    If Not IsDate(Me.FPREVCIERRE) Then
        p_Error = "El campo Fecha Prevista de Cierre es obligatorio"
        Err.Raise 1000
    End If
    If CDate(Me.FPREVCIERRE) < CDate(Me.FechaApertura) Then
        p_Error = "La fecha prevista de cierre no puede ser anterior a la de apertura"
        Err.Raise 1000
    End If
    Me.CodigoNoConformidad = Me.CodigoNoConformidadCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_ObjNCAlInicio Is Nothing Then
        Me.IDNoConformidad = Me.IDNoConformidadCalculado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_SQL = "TbNoConformidades"
    Else
        If Not IsNumeric(Me.IDNoConformidad) Then
            p_Error = "No se conoce el IDNC"
            Err.Raise 1000
        End If
        m_SQL = "SELECT TbNoConformidades.* " & _
            "FROM TbNoConformidades " & _
            "WHERE (((TbNoConformidades.IDNoConformidad)=" & Me.IDNoConformidad & "));"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If p_ObjNCAlInicio Is Nothing Then
            .AddNew
                 .Fields("IDNoConformidad") = Me.IDNoConformidad
                 .Fields("CodigoNoConformidad") = Me.CodigoNoConformidad
                 .Fields("Juridica") = Me.Juridica
                 .Fields("EXPEDIENTE") = Me.Expediente
                 If Me.Proyecto <> "" Then
                     Me.Proyecto = Me.ProyectoCalculado
                     p_Error = Me.Error
                     If p_Error <> "" Then
                         Err.Raise 1000
                     End If
                     .Fields("PROYECTO") = Me.ProyectoCalculado
                 End If
                 If Me.Vehiculo <> "" Then
                     .Fields("VEHICULO") = Me.Vehiculo
                 End If
                 .Fields("DESCRIPCION") = Me.Descripcion
                 .Fields("CAUSA") = Me.Causa
                 .Fields("ACR") = Me.ACR
                 If Me.RequiereACRCalculado = EnumSiNo.Sí Then
                     .Fields("RequiereACR") = True
                 Else
                     .Fields("RequiereACR") = False
                 End If
                 If Me.RequiereControlEficaciaCalculado = EnumSiNo.Sí Then
                     .Fields("RequiereControlEficacia") = "Sí"
                 Else
                     .Fields("RequiereControlEficacia") = "No"
                 End If
                 .Fields("ENTIDADRESPONSABLE") = Me.EntidadResponsable
                 If Me.ResponsableTelefonica <> "" Then
                     .Fields("RESPONSABLETELEFONICA") = Me.ResponsableTelefonica
                 End If
                 .Fields("FECHAAPERTURA") = Me.FechaApertura
                
                 If IsDate(Me.FPREVCIERRE) Then
                     .Fields("FPREVCIERRE") = Me.FPREVCIERRE
                 End If
                 .Fields("TIPO") = Me.Tipo
                 If Me.Notas <> "" Then
                     .Fields("NOTAS") = Me.Notas
                 End If
                 .Fields("Borrado") = False
                 If IsDate(Me.FechaPrevistaControlEficacia) Then
                     .Fields("FechaPrevistaControlEficacia") = Me.FechaPrevistaControlEficacia
                 End If
                 .Fields("RESPONSABLECALIDAD") = m_ObjUsuarioConectado.Nombre
            .Update
        Else
            If .EOF Then
                p_Error = "No aparece la No Conformidad como registrada"
                Err.Raise 1000
            End If
            .Edit
                .Fields("Juridica") = Me.Juridica
                .Fields("EXPEDIENTE") = Me.Expediente
                If Me.Proyecto <> "" Then
                    Me.Proyecto = Me.ProyectoCalculado
                    p_Error = Me.Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    .Fields("PROYECTO") = Me.ProyectoCalculado
                Else
                    .Fields("PROYECTO") = Nothing
                End If
                If Me.Vehiculo <> "" Then
                    .Fields("VEHICULO") = Me.Vehiculo
                Else
                    .Fields("VEHICULO") = Nothing
                End If
                .Fields("DESCRIPCION") = Me.Descripcion
                .Fields("CAUSA") = Me.Causa
                .Fields("ACR") = Me.ACR
                If Me.RequiereACRCalculado = EnumSiNo.Sí Then
                    .Fields("RequiereACR") = True
                Else
                    .Fields("RequiereACR") = False
                End If
                If Me.RequiereControlEficaciaCalculado = EnumSiNo.Sí Then
                    .Fields("RequiereControlEficacia") = "Sí"
                Else
                    .Fields("RequiereControlEficacia") = "No"
                End If
                .Fields("ENTIDADRESPONSABLE") = Me.EntidadResponsable
                If Me.ResponsableTelefonica <> "" Then
                    .Fields("RESPONSABLETELEFONICA") = Me.ResponsableTelefonica
                End If
                .Fields("FECHAAPERTURA") = Me.FechaApertura
                
                If IsDate(Me.FPREVCIERRE) Then
                    .Fields("FPREVCIERRE") = Me.FPREVCIERRE
                Else
                    .Fields("FPREVCIERRE") = Null
                End If
                .Fields("TIPO") = Me.Tipo
                If Me.Notas <> "" Then
                    .Fields("NOTAS") = Me.Notas
                Else
                    .Fields("NOTAS") = Null
                End If
                .Fields("Borrado") = False
                If IsDate(Me.FechaPrevistaControlEficacia) Then
                    .Fields("FechaPrevistaControlEficacia") = Me.FechaPrevistaControlEficacia
                Else
                    .Fields("FechaPrevistaControlEficacia") = Null
                End If
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing

    ' --- Invalida SearchComboCache para que combos de búsqueda vean el nuevo
    '     EXPEDIENTE/PROYECTO/VEHICULO en el siguiente refresh.
    '     Crítico (issue #35): antes de este hook, SearchComboCache nunca se
    '     invalidaba desde producción, así que los combos mostraban valores
    '     stale hasta cerrar Access. Reset best-effort: si falla, no abortamos
    '     la mutación ya persistida. ---
    On Error Resume Next
    SearchComboCache_Reset ""
    On Error GoTo errores

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método NoConformidades.Registrar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If

End Function


Private Function getCodigo(ByRef p_Error As String) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_OrdinalUltimoActual As String
    Dim m_PrimeraParteCodigo As String
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    If Me.EsNoConformidadCalculado = 0 Then
        p_Error = "No se puede determinar si es una No conformidad o una observación"
        Err.Raise 1000
    End If
    If Me.EsNoConformidadCalculado = EnumSiNo.Sí Then
        m_PrimeraParteCodigo = "NC"
    Else
        m_PrimeraParteCodigo = "OB"
    End If
    
    m_SQL = "SELECT Max(CInt(Right([CodigoNoConformidad],4))) AS Ordinal " & _
                "FROM TbNoConformidades " & _
                "WHERE (((TbNoConformidades.CodigoNoConformidad) Like '" & m_PrimeraParteCodigo & _
                "*'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(.Fields("Ordinal")) Then
                m_OrdinalUltimoActual = .Fields("Ordinal")
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    getCodigo = m_PrimeraParteCodigo & Format(m_OrdinalUltimoActual + 1, "0000")
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método NoConformidades.getCodigo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function


Public Function Eliminar( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_SQL As String
    On Error GoTo errores
    
    m_SQL = "DELETE TbNoConformidades.IDNoConformidad " & _
            "FROM TbNoConformidades " & _
            "WHERE (((TbNoConformidades.IDNoConformidad)=" & Me.IDNoConformidad & "));"
    getdb.Execute m_SQL

    ' --- Invalida SearchComboCache: un DELETE puede quitar valores que
    '     estaban en combos (EXPEDIENTE/PROYECTO/VEHICULO). Mismo hook que
    '     en Registrar (issue #35). Reset best-effort: si falla, no abortamos. ---
    On Error Resume Next
    SearchComboCache_Reset ""
    On Error GoTo errores

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método NoConformidades.Eliminar ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
        .Add "IDNoConformidad"
        .Add "Juridica"
        .Add "CodigoNoConformidad"
        .Add "EsNoConformidad"
        .Add "EXPEDIENTE"
        .Add "PROYECTO"
        .Add "VEHICULO"
        .Add "DESCRIPCION"
        .Add "CAUSA"
        .Add "ENTIDADRESPONSABLE"
        .Add "RESPONSABLETELEFONICA"
        .Add "FECHAAPERTURA"
        .Add "FECHACIERRE"
        .Add "FPREVCIERRE"
        .Add "TIPO"
        .Add "NOTAS"
        .Add "Borrado"
        .Add "RequiereACR"
        .Add "ACR"
        .Add "MotivoBorrado"
        .Add "RequiereControlEficacia"
        .Add "ControlEficacia"
        .Add "FechaControlEficacia"
        .Add "FechaPrevistaControlEficacia"
        .Add "RESPONSABLECALIDAD"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método NC.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDNoConformidad" Then
        getPropiedad = Me.IDNoConformidad
        Exit Function
    End If
    If m_NombrePropiedad = "Juridica" Then
        getPropiedad = Me.Juridica
        Exit Function
    End If
    If m_NombrePropiedad = "CodigoNoConformidad" Then
        getPropiedad = Me.CodigoNoConformidad
        Exit Function
    End If
    If m_NombrePropiedad = "EsNoConformidad" Then
        getPropiedad = Me.EsNoConformidad
        Exit Function
    End If
    If m_NombrePropiedad = "EXPEDIENTE" Then
        getPropiedad = Me.Expediente
        Exit Function
    End If
    If m_NombrePropiedad = "PROYECTO" Then
        getPropiedad = Me.Proyecto
        Exit Function
    End If
    If m_NombrePropiedad = "VEHICULO" Then
        getPropiedad = Me.Vehiculo
        Exit Function
    End If
    If m_NombrePropiedad = "DESCRIPCION" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "CAUSA" Then
        getPropiedad = Me.Causa
        Exit Function
    End If
    If m_NombrePropiedad = "ENTIDADRESPONSABLE" Then
        getPropiedad = Me.EntidadResponsable
        Exit Function
    End If
    If m_NombrePropiedad = "RESPONSABLETELEFONICA" Then
        getPropiedad = Me.ResponsableTelefonica
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAAPERTURA" Then
        getPropiedad = Me.FechaApertura
        Exit Function
    End If
    If m_NombrePropiedad = "FECHACIERRE" Then
        getPropiedad = Me.FechaCierre
        Exit Function
    End If
    If m_NombrePropiedad = "FPREVCIERRE" Then
        getPropiedad = Me.FPREVCIERRE
        Exit Function
    End If
    If m_NombrePropiedad = "TIPO" Then
        getPropiedad = Me.Tipo
        Exit Function
    End If
    If m_NombrePropiedad = "NOTAS" Then
        getPropiedad = Me.Notas
        Exit Function
    End If
    If m_NombrePropiedad = "Borrado" Then
        getPropiedad = Me.Borrado
        Exit Function
    End If
    If m_NombrePropiedad = "RequiereACR" Then
        getPropiedad = Me.RequiereACR
        Exit Function
    End If
    If m_NombrePropiedad = "ACR" Then
        getPropiedad = Me.ACR
        Exit Function
    End If
    If m_NombrePropiedad = "MotivoBorrado" Then
        getPropiedad = Me.MotivoBorrado
        Exit Function
    End If
    If m_NombrePropiedad = "RequiereControlEficacia" Then
        getPropiedad = Me.RequiereControlEficacia
        Exit Function
    End If
    If m_NombrePropiedad = "ControlEficacia" Then
        getPropiedad = Me.ControlEficacia
        Exit Function
    End If
    If m_NombrePropiedad = "FechaControlEficacia" Then
        getPropiedad = Me.FechaControlEficacia
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaControlEficacia" Then
        getPropiedad = Me.FechaPrevistaControlEficacia
        Exit Function
    End If
    If m_NombrePropiedad = "RESPONSABLECALIDAD" Then
        getPropiedad = Me.ResponsableCalidad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en NC.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método NC.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function



Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDNoConformidad" Then
        Me.IDNoConformidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Juridica" Then
        Me.Juridica = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodigoNoConformidad" Then
        Me.CodigoNoConformidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsNoConformidad" Then
        Me.EsNoConformidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EXPEDIENTE" Then
        Me.Expediente = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PROYECTO" Then
        Me.Proyecto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "VEHICULO" Then
        Me.Vehiculo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DESCRIPCION" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CAUSA" Then
        Me.Causa = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ENTIDADRESPONSABLE" Then
        Me.EntidadResponsable = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RESPONSABLETELEFONICA" Then
        Me.ResponsableTelefonica = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAAPERTURA" Then
        Me.FechaApertura = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHACIERRE" Then
        Me.FechaCierre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FPREVCIERRE" Then
        Me.FPREVCIERRE = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TIPO" Then
        Me.Tipo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NOTAS" Then
        Me.Notas = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Borrado" Then
        Me.Borrado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RequiereACR" Then
        Me.RequiereACR = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ACR" Then
        Me.ACR = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "MotivoBorrado" Then
        Me.MotivoBorrado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RequiereControlEficacia" Then
        Me.RequiereControlEficacia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ControlEficacia" Then
        Me.ControlEficacia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaControlEficacia" Then
        Me.FechaControlEficacia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaControlEficacia" Then
        Me.FechaPrevistaControlEficacia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RESPONSABLECALIDAD" Then
        Me.ResponsableCalidad = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en NC.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método NC.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function



