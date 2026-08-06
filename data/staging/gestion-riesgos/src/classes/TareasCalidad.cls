VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TareasCalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_objColEdicionesPreparadasParaPublicar As Scripting.Dictionary
Private m_objColEdicionesAPuntoDeCaducar As Scripting.Dictionary
Private m_objColEdicionesCaducadas As Scripting.Dictionary
Private m_objColRiesgosPorReTipificar As Scripting.Dictionary

Private m_objColRiesgosAceptadosSinVisar As Scripting.Dictionary
Private m_objColRiesgosAceptadosRechazados As Scripting.Dictionary

Private m_objColRiesgosRetiradosSinVisar As Scripting.Dictionary
Private m_objColRiesgosRetiradosRechazados As Scripting.Dictionary

Private m_objColRiesgosMaterializadosPorDedicir As Scripting.Dictionary






Private m_sNEdicionesPreparadasParaPublicar As String


Private m_sNEdicionesAPuntoDeCaducar As String
Private m_sNEdicionesCaducadas As String

Private m_sNRiesgosPendientesDeTipificar As String
Private m_sNRiesgosRetipificados As String
Private m_sNRiesgosAceptadosSinVisar As String
Private m_sNRiesgosAceptadosRechazados As String


Private m_sNRiesgosRetiradosSinVisar As String
Private m_sNRiesgosRetiradosRechazados As String


Private m_sNRiesgosMaterializadosPorDedicir As String


Private m_eHayTareasPendientes As EnumSiNo
Private m_sTareasTotales As String

Private m_sTareasTotalesCalidad As String
Private m_sTareasTotalesAdministrador As String
Private m_sTareasTotalesTecnico As String
Public TareasObtenidas As EnumSiNo

Public Error As String
Private m_Error As String

Public Function RefrescarPendientesCalidad(Optional ByRef p_Error As String) As String
    On Error GoTo errores

    p_Error = ""
    Me.Error = ""
    m_Error = ""

    LimpiarCachePendientesCalidad
    m_sTareasTotalesCalidad = Me.TareasTotalesCalidad
    If Me.Error <> "" Then
        p_Error = Me.Error
        Err.Raise 1000
    End If

    RefrescarPendientesCalidad = m_sTareasTotalesCalidad
    Exit Function
errores:
    If p_Error = "" Then
        p_Error = "El método Tareas.RefrescarPendientesCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = p_Error
End Function

Private Sub LimpiarCachePendientesCalidad()
    Set m_objColEdicionesPreparadasParaPublicar = Nothing
    Set m_objColEdicionesAPuntoDeCaducar = Nothing
    Set m_objColEdicionesCaducadas = Nothing
    Set m_objColRiesgosPorReTipificar = Nothing
    Set m_objColRiesgosAceptadosSinVisar = Nothing
    Set m_objColRiesgosAceptadosRechazados = Nothing
    Set m_objColRiesgosRetiradosSinVisar = Nothing
    Set m_objColRiesgosRetiradosRechazados = Nothing
    Set m_objColRiesgosMaterializadosPorDedicir = Nothing

    m_sNEdicionesPreparadasParaPublicar = ""
    m_sNEdicionesAPuntoDeCaducar = ""
    m_sNEdicionesCaducadas = ""
    m_sNRiesgosPendientesDeTipificar = ""
    m_sNRiesgosRetipificados = ""
    m_sNRiesgosAceptadosSinVisar = ""
    m_sNRiesgosAceptadosRechazados = ""
    m_sNRiesgosRetiradosSinVisar = ""
    m_sNRiesgosRetiradosRechazados = ""
    m_sNRiesgosMaterializadosPorDedicir = ""
    m_sTareasTotales = ""
    m_sTareasTotalesCalidad = ""
    m_sTareasTotalesAdministrador = ""
    m_sTareasTotalesTecnico = ""
    TareasObtenidas = EnumSiNo.No
End Sub

Public Property Get NEdicionesPreparadasParaPublicar() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.ColEdicionesPreparadasParaPublicar Is Nothing Then
        m_sNEdicionesPreparadasParaPublicar = "0"
    Else
        m_sNEdicionesPreparadasParaPublicar = CStr(Me.ColEdicionesPreparadasParaPublicar.Count)
    End If
    NEdicionesPreparadasParaPublicar = m_sNEdicionesPreparadasParaPublicar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.NEdicionesPreparadasParaPublicar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get NRiesgosPendientesDeTipificar() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.ColRiesgosPorReTipificar Is Nothing Then
        m_sNRiesgosPendientesDeTipificar = "0"
    Else
        m_sNRiesgosPendientesDeTipificar = CStr(Me.ColRiesgosPorReTipificar.Count)
    End If
    NRiesgosPendientesDeTipificar = m_sNRiesgosPendientesDeTipificar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.NRiesgosPendientesDeTipificar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get NRiesgosAceptadosSinVisar() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.ColRiesgosAceptadosSinVisar Is Nothing Then
        m_sNRiesgosAceptadosSinVisar = "0"
    Else
        m_sNRiesgosAceptadosSinVisar = CStr(Me.ColRiesgosAceptadosSinVisar.Count)
    End If
    NRiesgosAceptadosSinVisar = m_sNRiesgosAceptadosSinVisar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.NRiesgosAceptadosSinVisar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get NRiesgosAceptadosRechazados() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.ColRiesgosAceptadosRechazados Is Nothing Then
        m_sNRiesgosAceptadosRechazados = "0"
    Else
        m_sNRiesgosAceptadosRechazados = CStr(Me.ColRiesgosAceptadosRechazados.Count)
    End If
    NRiesgosAceptadosRechazados = m_sNRiesgosAceptadosRechazados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.NRiesgosAceptadosRechazados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get NRiesgosRetiradosSinVisar() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.ColRiesgosRetiradosSinVisar Is Nothing Then
        
        m_sNRiesgosRetiradosSinVisar = "0"
    Else
        m_sNRiesgosRetiradosSinVisar = CStr(Me.ColRiesgosRetiradosSinVisar.Count)
    End If
    NRiesgosRetiradosSinVisar = m_sNRiesgosRetiradosSinVisar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.NRiesgosRetiradosSinVisar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get NRiesgosRetiradosRechazados() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.ColRiesgosRetiradosRechazados Is Nothing Then
        m_sNRiesgosRetiradosRechazados = "0"
    Else
        m_sNRiesgosRetiradosRechazados = CStr(Me.ColRiesgosRetiradosRechazados.Count)
    End If
    NRiesgosRetiradosRechazados = m_sNRiesgosRetiradosRechazados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.NRiesgosRetiradosRechazados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get NRiesgosMaterializadosPorDedicir() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.ColRiesgosMaterializadosPorDedicir Is Nothing Then
        
        m_sNRiesgosMaterializadosPorDedicir = "0"
    Else
        m_sNRiesgosMaterializadosPorDedicir = CStr(Me.ColRiesgosMaterializadosPorDedicir.Count)
    End If
    
    NRiesgosMaterializadosPorDedicir = m_sNRiesgosMaterializadosPorDedicir
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.NRiesgosMaterializadosPorDedicir ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get NEdicionesCaducadas() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.ColEdicionesCaducadas Is Nothing Then
        m_sNEdicionesCaducadas = "0"
    Else
        m_sNEdicionesCaducadas = CStr(Me.ColEdicionesCaducadas.Count)
    End If
    NEdicionesCaducadas = m_sNEdicionesCaducadas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.NEdicionesCaducadas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get NEdicionesAPuntoDeCaducar() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.ColEdicionesApuntoDeCaducar Is Nothing Then
        m_sNEdicionesAPuntoDeCaducar = "0"
    Else
        m_sNEdicionesAPuntoDeCaducar = CStr(Me.ColEdicionesApuntoDeCaducar.Count)
    End If
    NEdicionesAPuntoDeCaducar = m_sNEdicionesAPuntoDeCaducar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.NEdicionesAPuntoDeCaducar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get TareasTotalesCalidad() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sTareasTotalesCalidad = CStr(CInt(Me.NRiesgosPendientesDeTipificar) + CInt(Me.NRiesgosAceptadosSinVisar) + _
                                CInt(Me.NRiesgosRetiradosSinVisar) + CInt(Me.NRiesgosMaterializadosPorDedicir) + _
                                CInt(Me.NEdicionesPreparadasParaPublicar) + CInt(Me.NEdicionesCaducadas))
   
    TareasTotalesCalidad = m_sTareasTotalesCalidad
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.TareasTotalesCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let TareasTotalesCalidad(ByVal sNewValue As String)

    m_sTareasTotalesCalidad = sNewValue

End Property

Public Property Get TareasTotalesAdministrador() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sTareasTotalesAdministrador = CStr(CInt(Me.NEdicionesAPuntoDeCaducar) + CInt(Me.NEdicionesCaducadas) + _
                                CInt(Me.NRiesgosPendientesDeTipificar) + CInt(Me.NRiesgosAceptadosSinVisar) + _
                                CInt(Me.NRiesgosRetiradosSinVisar) + CInt(Me.NRiesgosMaterializadosPorDedicir) + _
                                CInt(Me.NEdicionesPreparadasParaPublicar))
   
    TareasTotalesAdministrador = m_sTareasTotalesAdministrador
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.TareasTotalesAdministrador ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let TareasTotalesAdministrador(ByVal sNewValue As String)
    
    m_sTareasTotalesAdministrador = sNewValue
    
End Property

Public Property Get TareasTotalesTecnico() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sTareasTotalesTecnico = CStr(CInt(Me.NEdicionesAPuntoDeCaducar) + CInt(Me.NEdicionesCaducadas) + _
                                CInt(Me.NRiesgosAceptadosRechazados) + CInt(Me.NRiesgosAceptadosSinVisar) + _
                                CInt(Me.NRiesgosRetiradosSinVisar) + CInt(Me.NRiesgosRetiradosRechazados) + _
                                CInt(Me.NRiesgosMaterializadosPorDedicir) + CInt(Me.NEdicionesPreparadasParaPublicar))
   
   
    TareasTotalesTecnico = m_sTareasTotalesTecnico
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.TareasTotalesTecnico ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let TareasTotalesTecnico(ByVal sNewValue As String)

    m_sTareasTotalesTecnico = sNewValue

End Property

Public Property Get HayTareasPendientes() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not IsNumeric(Me.TareasTotales) Then
        Exit Property
    End If
    If CInt(Me.TareasTotales) = 0 Then
        m_eHayTareasPendientes = EnumSiNo.No
    Else
        m_eHayTareasPendientes = EnumSiNo.Sí
    End If
    HayTareasPendientes = m_eHayTareasPendientes
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.HayTareasPendientes ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Let HayTareasPendientes(valor As EnumSiNo)

    m_eHayTareasPendientes = valor
End Property

Public Property Get TareasTotales() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If EsAdministrador = EnumSiNo.Sí Then
        m_sTareasTotales = Me.TareasTotalesAdministrador
    ElseIf EsCalidad = EnumSiNo.Sí Then
        m_sTareasTotales = Me.TareasTotalesCalidad
    Else
        m_sTareasTotales = Me.TareasTotalesTecnico
    End If
   
    TareasTotales = m_sTareasTotales
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.TareasTotales ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get ColEdicionesPreparadasParaPublicar() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColEdicionesPreparadasParaPublicar Is Nothing Then
        Set ColEdicionesPreparadasParaPublicar = m_objColEdicionesPreparadasParaPublicar
        Exit Property
    End If
    Set m_objColEdicionesPreparadasParaPublicar = getEdicionesPreparadasParaPublicar(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColEdicionesPreparadasParaPublicar = m_objColEdicionesPreparadasParaPublicar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.ColEdicionesPreparadasParaPublicar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColEdicionesPreparadasParaPublicar(ByVal objNewValue As Variant)

    Set m_objColEdicionesPreparadasParaPublicar = objNewValue

End Property

Public Property Get ColEdicionesApuntoDeCaducar() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColEdicionesAPuntoDeCaducar Is Nothing Then
        Set ColEdicionesApuntoDeCaducar = m_objColEdicionesAPuntoDeCaducar
        Exit Property
    End If
    Set m_objColEdicionesAPuntoDeCaducar = getEdicionesAPuntoDeCaducarSinPropuesta(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColEdicionesApuntoDeCaducar = m_objColEdicionesAPuntoDeCaducar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.ColEdicionesAPuntoDeCaducar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColEdicionesApuntoDeCaducar(ByVal objNewValue As Variant)

    Set m_objColEdicionesAPuntoDeCaducar = objNewValue

End Property
Public Property Get ColEdicionesCaducadas() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColEdicionesCaducadas Is Nothing Then
        Set ColEdicionesCaducadas = m_objColEdicionesCaducadas
        Exit Property
    End If
    Set m_objColEdicionesCaducadas = getEdicionesCaducadasSinPropuesta(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColEdicionesCaducadas = m_objColEdicionesCaducadas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.ColEdicionesCaducadas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColEdicionesCaducadas(ByVal objNewValue As Variant)

    Set m_objColEdicionesCaducadas = objNewValue

End Property




Public Property Get ColRiesgosPorReTipificar() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosPorReTipificar Is Nothing Then
        Set ColRiesgosPorReTipificar = m_objColRiesgosPorReTipificar
        Exit Property
    End If
    Set m_objColRiesgosPorReTipificar = getRiesgosPorRetipificar(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColRiesgosPorReTipificar = m_objColRiesgosPorReTipificar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.ColRiesgosPorReTipificar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosPorReTipificar(ByVal objNewValue As Variant)
    
    
    
    Set m_objColRiesgosPorReTipificar = objNewValue

End Property

Public Property Get ColRiesgosAceptadosSinVisar() As Scripting.Dictionary
    
        
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosAceptadosSinVisar Is Nothing Then
        Set ColRiesgosAceptadosSinVisar = m_objColRiesgosAceptadosSinVisar
        Exit Property
    End If
    Set m_objColRiesgosAceptadosSinVisar = getRiesgosAceptadosPorVisar(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColRiesgosAceptadosSinVisar = m_objColRiesgosAceptadosSinVisar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.ColRiesgosAceptadosSinVisar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosAceptadosSinVisar(ByVal objNewValue As Variant)

    Set m_objColRiesgosAceptadosSinVisar = objNewValue

End Property

Public Property Get ColRiesgosAceptadosRechazados() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosAceptadosRechazados Is Nothing Then
        Set ColRiesgosAceptadosRechazados = m_objColRiesgosAceptadosRechazados
        Exit Property
    End If
    Set m_objColRiesgosAceptadosRechazados = getRiesgosAceptadosRechazados(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColRiesgosAceptadosRechazados = m_objColRiesgosAceptadosRechazados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.ColRiesgosAceptadosRechazados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosAceptadosRechazados(ByVal objNewValue As Variant)

    Set m_objColRiesgosAceptadosRechazados = objNewValue

End Property

Public Property Get ColRiesgosRetiradosSinVisar() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosRetiradosSinVisar Is Nothing Then
        Set ColRiesgosRetiradosSinVisar = m_objColRiesgosRetiradosSinVisar
        Exit Property
    End If
    Set m_objColRiesgosRetiradosSinVisar = getRiesgosRetiradosPorVisar(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColRiesgosRetiradosSinVisar = m_objColRiesgosRetiradosSinVisar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.ColRiesgosRetiradosSinVisar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosRetiradosSinVisar(ByVal objNewValue As Variant)

    Set m_objColRiesgosRetiradosSinVisar = objNewValue

End Property

Public Property Get ColRiesgosRetiradosRechazados() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosRetiradosRechazados Is Nothing Then
        Set ColRiesgosRetiradosRechazados = m_objColRiesgosRetiradosRechazados
        Exit Property
    End If
    Set m_objColRiesgosRetiradosRechazados = getRiesgosRetiradosRechazados(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColRiesgosRetiradosRechazados = m_objColRiesgosRetiradosRechazados
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.ColRiesgosRetiradosRechazados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosRetiradosRechazados(ByVal objNewValue As Variant)
    
    Set m_objColRiesgosRetiradosRechazados = objNewValue

End Property

Public Property Get ColRiesgosMaterializadosPorDedicir() As Scripting.Dictionary
    
   
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosMaterializadosPorDedicir Is Nothing Then
        Set ColRiesgosMaterializadosPorDedicir = m_objColRiesgosMaterializadosPorDedicir
        Exit Property
    End If
    Set m_objColRiesgosMaterializadosPorDedicir = getRiesgosMaterializadosPorDecidir(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColRiesgosMaterializadosPorDedicir = m_objColRiesgosMaterializadosPorDedicir
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Tareas.ColRiesgosMaterializadosPorDedicir ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosMaterializadosPorDedicir(ByVal objNewValue As Variant)

    Set m_objColRiesgosMaterializadosPorDedicir = objNewValue

End Property




Public Function getEdicionesAPuntoDeCaducarSinPropuesta( _
                                                            Optional ByRef p_Error As String _
                                                            ) As Scripting.Dictionary
        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Edicion As Edicion
        
    On Error GoTo errores
    
    If EsTecnico = EnumSiNo.Sí Then
        m_SQL = "SELECT TbProyectosEdiciones.* " & _
                "FROM TbProyectos INNER JOIN TbProyectosEdiciones ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
                "WHERE (((TbProyectos.FechaCierre) Is Null) AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                "AND ((TbProyectosEdiciones.PropuestaRechazadaPorCalidadFecha) Is Null) " & _
                "AND ((DateDiff('d',Now(),[TbProyectos].[FechaMaxProximaPublicacion])) Between 1 And 15) " & _
                "AND ((TbProyectos.CadenaNombreAutorizados) Like '*" & m_ObjUsuarioConectado.Nombre & "*'));"
    Else
        If EsCalidad = EnumSiNo.Sí Or EsAdministrador = EnumSiNo.Sí Then
            If m_ObjUsuarioParaTareas Is Nothing Then
                m_SQL = "SELECT TbProyectosEdiciones.* " & _
                        "FROM TbProyectos INNER JOIN TbProyectosEdiciones ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
                        "WHERE (((TbProyectos.FechaCierre) Is Null) AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                        "AND ((TbProyectosEdiciones.PropuestaRechazadaPorCalidadFecha) Is Null) " & _
                        "AND ((DateDiff('d',Now(),[TbProyectos].[FechaMaxProximaPublicacion])) Between 1 And 15));"
            Else
                m_SQL = "SELECT TbProyectosEdiciones.* " & _
                        "FROM TbProyectos INNER JOIN TbProyectosEdiciones ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
                        "WHERE (((TbProyectos.FechaCierre) Is Null) AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                        "AND ((TbProyectosEdiciones.PropuestaRechazadaPorCalidadFecha) Is Null) " & _
                        "AND ((DateDiff('d',Now(),[TbProyectos].[FechaMaxProximaPublicacion])) Between 1 And 15) " & _
                        "AND ((TbProyectos.NombreUsuarioCalidad)='" & m_ObjUsuarioParaTareas.Nombre & "'));"
            End If
                   
        End If
        
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Edicion = New Edicion
            For Each m_Campo In m_Edicion.ColCampos
                m_Edicion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getEdicionesAPuntoDeCaducarSinPropuesta Is Nothing Then
                Set getEdicionesAPuntoDeCaducarSinPropuesta = New Scripting.Dictionary
                getEdicionesAPuntoDeCaducarSinPropuesta.CompareMode = TextCompare
            End If
            If Not getEdicionesAPuntoDeCaducarSinPropuesta.Exists(CStr(m_Edicion.IDEdicion)) Then
                getEdicionesAPuntoDeCaducarSinPropuesta.Add CStr(m_Edicion.IDEdicion), m_Edicion
            End If
            Set m_Edicion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getEdicionesAPuntoDeCaducarSinPropuesta ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getEdicionesCaducadasSinPropuesta( _
                                                            Optional ByRef p_Error As String _
                                                            ) As Scripting.Dictionary
        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Edicion As Edicion
    
    On Error GoTo errores
    m_SQL = getSQLEdicionesCaducadasSinPropuesta(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            If getEdicionesCaducadasSinPropuesta Is Nothing Then
                Set getEdicionesCaducadasSinPropuesta = New Scripting.Dictionary
                getEdicionesCaducadasSinPropuesta.CompareMode = TextCompare
            End If
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Edicion = New Edicion
            For Each m_Campo In m_Edicion.ColCampos
                m_Edicion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getEdicionesCaducadasSinPropuesta Is Nothing Then
                Set getEdicionesCaducadasSinPropuesta = New Scripting.Dictionary
                getEdicionesCaducadasSinPropuesta.CompareMode = TextCompare
            End If
            If Not getEdicionesCaducadasSinPropuesta.Exists(CStr(m_Edicion.IDEdicion)) Then
                getEdicionesCaducadasSinPropuesta.Add CStr(m_Edicion.IDEdicion), m_Edicion
            End If
            Set m_Edicion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getEdicionesCaducadasSinPropuesta ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getSQLEdicionesCaducadasSinPropuesta(Optional ByRef p_Error As String) As String

    On Error GoTo errores

    p_Error = ""

    If EsTecnico = EnumSiNo.Sí Then
        getSQLEdicionesCaducadasSinPropuesta = "SELECT TbProyectosEdiciones.* " & _
                                                "FROM TbProyectos INNER JOIN TbProyectosEdiciones ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
                                                "WHERE (((TbProyectos.FechaCierre) Is Null) AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                                                "AND ((TbProyectosEdiciones.PropuestaRechazadaPorCalidadFecha) Is Null) " & _
                                                "AND ((DateDiff('d',Now(),[TbProyectos].[FechaMaxProximaPublicacion]))<1) " & _
                                                "AND ((TbProyectos.CadenaNombreAutorizados) Like '*" & m_ObjUsuarioConectado.Nombre & "*'));"
        Exit Function
    End If

    If EsCalidad = EnumSiNo.Sí Or EsAdministrador = EnumSiNo.Sí Then
        If m_ObjUsuarioParaTareas Is Nothing Then
            getSQLEdicionesCaducadasSinPropuesta = "SELECT TbProyectosEdiciones.* " & _
                                                    "FROM TbProyectos INNER JOIN TbProyectosEdiciones ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
                                                    "WHERE (((TbProyectos.FechaCierre) Is Null) AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                                                    "AND ((TbProyectosEdiciones.PropuestaRechazadaPorCalidadFecha) Is Null) " & _
                                                    "AND ((DateDiff('d',Now(),[TbProyectos].[FechaMaxProximaPublicacion]))<1));"
        Else
            getSQLEdicionesCaducadasSinPropuesta = "SELECT TbProyectosEdiciones.* " & _
                                                    "FROM TbProyectos INNER JOIN TbProyectosEdiciones ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
                                                    "WHERE (((TbProyectos.FechaCierre) Is Null) AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                                                    "AND ((TbProyectosEdiciones.PropuestaRechazadaPorCalidadFecha) Is Null) " & _
                                                    "AND ((DateDiff('d',Now(),[TbProyectos].[FechaMaxProximaPublicacion]))<1) " & _
                                                    "AND ((TbProyectos.NombreUsuarioCalidad)='" & m_ObjUsuarioParaTareas.Nombre & "'));"
        End If
    End If

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getSQLEdicionesCaducadasSinPropuesta ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getEdicionesPreparadasParaPublicar( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Edicion As Edicion
        
    On Error GoTo errores
    
    If EsTecnico = EnumSiNo.Sí Then
        m_SQL = "SELECT TbProyectosEdiciones.* " & _
                "FROM TbProyectos INNER JOIN TbProyectosEdiciones ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
                "AND ((TbProyectos.FechaCierre) Is Null) " & _
                "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                "AND ((TbProyectosEdiciones.PropuestaRechazadaPorCalidadFecha) Is Null) " & _
                "AND ((TbProyectos.CadenaNombreAutorizados) Like'*" & m_ObjUsuarioConectado.Nombre & "*'));"
    Else
        If EsCalidad = EnumSiNo.Sí Or EsAdministrador = EnumSiNo.Sí Then
            If m_ObjUsuarioParaTareas Is Nothing Then
                m_SQL = "SELECT TbProyectosEdiciones.* " & _
                        "FROM TbProyectos INNER JOIN TbProyectosEdiciones " & _
                        "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
                        "WHERE ((Not (TbProyectosEdiciones.FechaPreparadaParaPublicar) Is Null) " & _
                        "AND ((TbProyectos.FechaCierre) Is Null) " & _
                        "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                        "AND ((TbProyectosEdiciones.PropuestaRechazadaPorCalidadFecha) Is Null));"
            Else
                m_SQL = "SELECT TbProyectosEdiciones.* " & _
                        "FROM TbProyectos INNER JOIN TbProyectosEdiciones ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
                        "WHERE ((Not (TbProyectosEdiciones.FechaPreparadaParaPublicar) Is Null) " & _
                        "AND ((TbProyectos.FechaCierre) Is Null) " & _
                        "AND ((TbProyectos.NombreUsuarioCalidad)='" & m_ObjUsuarioParaTareas.Nombre & "') " & _
                        "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                        "AND ((TbProyectosEdiciones.PropuestaRechazadaPorCalidadFecha) Is Null));"
            End If
                   
        End If
        
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Edicion = New Edicion
            For Each m_Campo In m_Edicion.ColCampos
                m_Edicion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getEdicionesPreparadasParaPublicar Is Nothing Then
                Set getEdicionesPreparadasParaPublicar = New Scripting.Dictionary
                getEdicionesPreparadasParaPublicar.CompareMode = TextCompare
            End If
            If Not getEdicionesPreparadasParaPublicar.Exists(CStr(m_Edicion.IDEdicion)) Then
                getEdicionesPreparadasParaPublicar.Add CStr(m_Edicion.IDEdicion), m_Edicion
            End If
            Set m_Edicion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getEdicionesPreparadasParaPublicar ha devuelto el error: " & Err.Description
    End If
End Function


Private Function getRiesgosAceptadosPorVisar( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Riesgo As riesgo
        
    On Error GoTo errores
    
    If EsTecnico = EnumSiNo.Sí Then
        m_SQL = "SELECT TbRiesgos.* " & _
                "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                "WHERE ((Not (TbRiesgos.FechaJustificacionAceptacionRiesgo) Is Null) " & _
                "AND ((TbProyectos.FechaCierre) Is Null) " & _
                "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) AND ((TbRiesgos.Mitigacion)='Aceptar') " & _
                "AND ((TbRiesgos.FechaRechazoAceptacionPorCalidad) Is Null) " & _
                "AND ((TbRiesgos.FechaAprobacionAceptacionPorCalidad) Is Null) " & _
                "AND ((TbProyectos.CadenaNombreAutorizados) Like'*" & m_ObjUsuarioConectado.Nombre & "*'));"
    Else
        If EsCalidad = EnumSiNo.Sí Or EsAdministrador = EnumSiNo.Sí Then
            If m_ObjUsuarioParaTareas Is Nothing Then
                m_SQL = "SELECT TbRiesgos.* " & _
                        "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                        "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                        "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                        "WHERE ((Not (TbRiesgos.FechaJustificacionAceptacionRiesgo) Is Null) " & _
                        "AND ((TbProyectos.FechaCierre) Is Null) " & _
                        "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                        "AND ((TbRiesgos.Mitigacion)='Aceptar') " & _
                        "AND ((TbRiesgos.FechaRechazoAceptacionPorCalidad) Is Null) " & _
                        "AND ((TbRiesgos.FechaAprobacionAceptacionPorCalidad) Is Null));"
            Else
                m_SQL = "SELECT TbRiesgos.* " & _
                        "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                        "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                        "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                        "WHERE ((Not (TbRiesgos.FechaJustificacionAceptacionRiesgo) Is Null) " & _
                        "AND ((TbProyectos.FechaCierre) Is Null) " & _
                        "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                        "AND ((TbRiesgos.Mitigacion)='Aceptar') " & _
                        "AND ((TbProyectos.NombreUsuarioCalidad)='" & m_ObjUsuarioParaTareas.Nombre & "') " & _
                        "AND ((TbRiesgos.FechaRechazoAceptacionPorCalidad) Is Null) " & _
                        "AND ((TbRiesgos.FechaAprobacionAceptacionPorCalidad) Is Null));"
            End If
                   
        End If
        
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Riesgo = New riesgo
            For Each m_Campo In m_Riesgo.ColCampos
                m_Riesgo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getRiesgosAceptadosPorVisar Is Nothing Then
                Set getRiesgosAceptadosPorVisar = New Scripting.Dictionary
                getRiesgosAceptadosPorVisar.CompareMode = TextCompare
            End If
            If Not getRiesgosAceptadosPorVisar.Exists(CStr(m_Riesgo.IDRiesgo)) Then
                getRiesgosAceptadosPorVisar.Add CStr(m_Riesgo.IDRiesgo), m_Riesgo
            End If
            Set m_Riesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosAceptadosPorVisar ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getRiesgosAceptadosRechazados( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Riesgo As riesgo
        
    On Error GoTo errores
    
    If EsTecnico = EnumSiNo.Sí Then
        m_SQL = "SELECT TbRiesgos.* " & _
                "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                "WHERE ((Not (TbRiesgos.FechaJustificacionAceptacionRiesgo) Is Null) " & _
                "AND ((TbProyectos.FechaCierre) Is Null) " & _
                "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) AND ((TbRiesgos.Mitigacion)='Aceptar') " & _
                "AND (Not(TbRiesgos.FechaRechazoAceptacionPorCalidad) Is Null) " & _
                "AND ((TbProyectos.CadenaNombreAutorizados) Like'*" & m_ObjUsuarioConectado.Nombre & "*'));"
    Else
        If EsCalidad = EnumSiNo.Sí Or EsAdministrador = EnumSiNo.Sí Then
            If m_ObjUsuarioParaTareas Is Nothing Then
                m_SQL = "SELECT TbRiesgos.* " & _
                        "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                        "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                        "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                        "WHERE ((Not (TbRiesgos.FechaJustificacionAceptacionRiesgo) Is Null) " & _
                        "AND ((TbProyectos.FechaCierre) Is Null) " & _
                        "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                        "AND ((TbRiesgos.Mitigacion)='Aceptar') " & _
                        "AND (Not(TbRiesgos.FechaRechazoAceptacionPorCalidad) Is Null));"
            Else
                m_SQL = "SELECT TbRiesgos.* " & _
                        "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                        "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                        "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                        "WHERE ((Not (TbRiesgos.FechaJustificacionAceptacionRiesgo) Is Null) " & _
                        "AND ((TbProyectos.FechaCierre) Is Null) " & _
                        "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                        "AND ((TbRiesgos.Mitigacion)='Aceptar') " & _
                        "AND ((TbProyectos.NombreUsuarioCalidad)='" & m_ObjUsuarioParaTareas.Nombre & "') " & _
                        "AND (Not(TbRiesgos.FechaRechazoAceptacionPorCalidad) Is Null));"
            End If
                   
        End If
        
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Riesgo = New riesgo
            For Each m_Campo In m_Riesgo.ColCampos
                m_Riesgo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getRiesgosAceptadosRechazados Is Nothing Then
                Set getRiesgosAceptadosRechazados = New Scripting.Dictionary
                getRiesgosAceptadosRechazados.CompareMode = TextCompare
            End If
            If Not getRiesgosAceptadosRechazados.Exists(CStr(m_Riesgo.IDRiesgo)) Then
                getRiesgosAceptadosRechazados.Add CStr(m_Riesgo.IDRiesgo), m_Riesgo
            End If
            Set m_Riesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosAceptadosRechazados ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getRiesgosRetiradosPorVisar( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Riesgo As riesgo
        
    On Error GoTo errores
    
    If EsTecnico = EnumSiNo.Sí Then
        m_SQL = "SELECT TbRiesgos.* " & _
                "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                "WHERE ((Not (TbRiesgos.FechaJustificacionRetiroRiesgo) Is Null) " & _
                "AND ((TbProyectos.FechaCierre) Is Null) " & _
                "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                "AND ((TbRiesgos.FechaRechazoRetiroPorCalidad) Is Null) " & _
                "AND ((TbRiesgos.FechaAprobacionRetiroPorCalidad) Is Null) " & _
                "AND ((TbProyectos.CadenaNombreAutorizados) Like'*" & m_ObjUsuarioConectado.Nombre & "*'));"
    Else
        If EsCalidad = EnumSiNo.Sí Or EsAdministrador = EnumSiNo.Sí Then
            If m_ObjUsuarioParaTareas Is Nothing Then
                m_SQL = "SELECT TbRiesgos.* " & _
                        "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                        "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                        "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                        "WHERE ((Not (TbRiesgos.FechaJustificacionRetiroRiesgo) Is Null) " & _
                        "AND ((TbProyectos.FechaCierre) Is Null) " & _
                        "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                        "AND ((TbRiesgos.FechaRechazoRetiroPorCalidad) Is Null) " & _
                        "AND ((TbRiesgos.FechaAprobacionRetiroPorCalidad) Is Null));"
            Else
                m_SQL = "SELECT TbRiesgos.* " & _
                        "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                        "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                        "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                        "WHERE ((Not (TbRiesgos.FechaJustificacionRetiroRiesgo) Is Null) " & _
                        "AND ((TbProyectos.FechaCierre) Is Null) " & _
                        "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                        "AND ((TbProyectos.NombreUsuarioCalidad)='" & m_ObjUsuarioParaTareas.Nombre & "') " & _
                        "AND ((TbRiesgos.FechaRechazoRetiroPorCalidad) Is Null) " & _
                        "AND ((TbRiesgos.FechaAprobacionRetiroPorCalidad) Is Null));"
            End If
                   
        End If
        
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Riesgo = New riesgo
            For Each m_Campo In m_Riesgo.ColCampos
                m_Riesgo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getRiesgosRetiradosPorVisar Is Nothing Then
                Set getRiesgosRetiradosPorVisar = New Scripting.Dictionary
                getRiesgosRetiradosPorVisar.CompareMode = TextCompare
            End If
            If Not getRiesgosRetiradosPorVisar.Exists(CStr(m_Riesgo.IDRiesgo)) Then
                getRiesgosRetiradosPorVisar.Add CStr(m_Riesgo.IDRiesgo), m_Riesgo
            End If
            Set m_Riesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosRetiradosPorVisar ha devuelto el error: " & Err.Description
    End If
End Function
Private Function getRiesgosRetiradosRechazados( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Riesgo As riesgo
        
    On Error GoTo errores
    
    If EsTecnico = EnumSiNo.Sí Then
        m_SQL = "SELECT TbRiesgos.* " & _
                "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                "WHERE ((Not (TbRiesgos.FechaJustificacionRetiroRiesgo) Is Null) " & _
                "AND ((TbProyectos.FechaCierre) Is Null) " & _
                "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                "AND (Not(TbRiesgos.FechaRechazoRetiroPorCalidad) Is Null) " & _
                 "AND ((TbProyectos.CadenaNombreAutorizados) Like'*" & m_ObjUsuarioConectado.Nombre & "*'));"
    Else
        If EsCalidad = EnumSiNo.Sí Or EsAdministrador = EnumSiNo.Sí Then
            If m_ObjUsuarioParaTareas Is Nothing Then
                m_SQL = "SELECT TbRiesgos.* " & _
                       "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                        "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                        "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                        "WHERE ((Not (TbRiesgos.FechaJustificacionRetiroRiesgo) Is Null) " & _
                        "AND ((TbProyectos.FechaCierre) Is Null) " & _
                        "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                        "AND (Not(TbRiesgos.FechaRechazoRetiroPorCalidad) Is Null));"
            Else
                m_SQL = "SELECT TbRiesgos.* " & _
                        "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                        "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                        "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                        "AND ((TbProyectos.FechaCierre) Is Null) " & _
                        "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                        "AND ((TbProyectos.NombreUsuarioCalidad)='" & m_ObjUsuarioParaTareas.Nombre & "') " & _
                        "AND (Not(TbRiesgos.FechaRechazoRetiroPorCalidad) Is Null));"
            End If
                   
        End If
        
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Riesgo = New riesgo
            For Each m_Campo In m_Riesgo.ColCampos
                m_Riesgo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getRiesgosRetiradosRechazados Is Nothing Then
                Set getRiesgosRetiradosRechazados = New Scripting.Dictionary
                getRiesgosRetiradosRechazados.CompareMode = TextCompare
            End If
            If Not getRiesgosRetiradosRechazados.Exists(CStr(m_Riesgo.IDRiesgo)) Then
                getRiesgosRetiradosRechazados.Add CStr(m_Riesgo.IDRiesgo), m_Riesgo
            End If
            Set m_Riesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosRetiradosRechazados ha devuelto el error: " & Err.Description
    End If
End Function
Private Function getRiesgosPorRetipificar( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Riesgo As riesgo
        
    On Error GoTo errores
    
    If EsTecnico = EnumSiNo.Sí Then
        m_SQL = "SELECT TbRiesgos.* " & _
                "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                "WHERE (((TbRiesgos.CodRiesgoBiblioteca)='" & m_ObjEntorno.CodigoBibliotecaParaRetipificar & "') " & _
                "AND ((TbProyectos.CadenaNombreAutorizados)Like'*" & m_ObjUsuarioConectado.Nombre & "*') " & _
                "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null));"
    ElseIf EsCalidad = EnumSiNo.Sí Then
        m_SQL = "SELECT TbRiesgos.* " & _
                "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                "WHERE (((TbRiesgos.CodRiesgoBiblioteca)='" & m_ObjEntorno.CodigoBibliotecaParaRetipificar & "') " & _
                "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                "AND ((TbProyectos.NombreUsuarioCalidad)='" & m_ObjUsuarioParaTareas.Nombre & "'));"
    Else
        m_SQL = "SELECT TbRiesgos.* " & _
                "FROM TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                "WHERE (((TbRiesgos.CodRiesgoBiblioteca)='" & m_ObjEntorno.CodigoBibliotecaParaRetipificar & "') " & _
                "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null));"
    End If
        
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Riesgo = New riesgo
            For Each m_Campo In m_Riesgo.ColCampos
                m_Riesgo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getRiesgosPorRetipificar Is Nothing Then
                Set getRiesgosPorRetipificar = New Scripting.Dictionary
                getRiesgosPorRetipificar.CompareMode = TextCompare
            End If
            If Not getRiesgosPorRetipificar.Exists(CStr(m_Riesgo.IDRiesgo)) Then
                getRiesgosPorRetipificar.Add CStr(m_Riesgo.IDRiesgo), m_Riesgo
            End If
            Set m_Riesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosPorRetipificar ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getRiesgosMaterializadosPorDecidir( _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary
        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_RiesgoMaterializado As RiesgoMaterializacion
   
    On Error GoTo errores
    
    If EsTecnico = EnumSiNo.Sí Then
        m_SQL = "SELECT TbRiesgosMaterializaciones.* " & _
                "FROM ((TbProyectos INNER JOIN TbRiesgosMaterializaciones " & _
                "ON TbProyectos.IDProyecto = TbRiesgosMaterializaciones.IDProyecto) " & _
                "INNER JOIN TbRiesgos " & _
                "ON (TbRiesgosMaterializaciones.IDEdicion = TbRiesgos.IDEdicion) " & _
                "AND (TbRiesgosMaterializaciones.CodigoRiesgo = TbRiesgos.CodigoRiesgo)) " & _
                "INNER JOIN TbProyectosEdiciones " & _
                "ON TbRiesgos.IDEdicion = TbProyectosEdiciones.IDEdicion " & _
                "WHERE (((TbProyectos.FechaCierre) Is Null) " & _
                "AND ((TbProyectos.CadenaNombreAutorizados) Like '*" & m_ObjUsuarioConectado.Nombre & "*') " & _
                "AND ((TbRiesgosMaterializaciones.IDNC) Is Null) " & _
                "AND ((TbRiesgosMaterializaciones.ParaNC) Is Null) " & _
                "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                "AND ((TbRiesgos.FechaAprobacionAceptacionPorCalidad) Is Null) " & _
                "AND ((TbRiesgos.FechaAprobacionRetiroPorCalidad) Is Null)) " & _
                "ORDER BY TbRiesgosMaterializaciones.ID;"
    Else
        
        If m_ObjUsuarioParaTareas Is Nothing Then
            m_SQL = "SELECT TbRiesgosMaterializaciones.* " & _
                    "FROM ((TbProyectos INNER JOIN TbRiesgosMaterializaciones " & _
                    "ON TbProyectos.IDProyecto = TbRiesgosMaterializaciones.IDProyecto) " & _
                    "INNER JOIN TbRiesgos " & _
                    "ON (TbRiesgosMaterializaciones.IDEdicion = TbRiesgos.IDEdicion) " & _
                    "AND (TbRiesgosMaterializaciones.CodigoRiesgo = TbRiesgos.CodigoRiesgo)) " & _
                    "INNER JOIN TbProyectosEdiciones " & _
                    "ON TbRiesgos.IDEdicion = TbProyectosEdiciones.IDEdicion " & _
                    "WHERE (((TbProyectos.FechaCierre) Is Null) " & _
                    "AND ((TbRiesgosMaterializaciones.IDNC) Is Null) " & _
                    "AND ((TbRiesgosMaterializaciones.ParaNC) Is Null) " & _
                    "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                    "AND ((TbRiesgos.FechaAprobacionAceptacionPorCalidad) Is Null) " & _
                    "AND ((TbRiesgos.FechaAprobacionRetiroPorCalidad) Is Null)) " & _
                    "ORDER BY TbRiesgosMaterializaciones.ID;"
        Else
            m_SQL = "SELECT TbRiesgosMaterializaciones.* " & _
                    "FROM ((TbProyectos INNER JOIN TbRiesgosMaterializaciones " & _
                    "ON TbProyectos.IDProyecto = TbRiesgosMaterializaciones.IDProyecto) " & _
                    "INNER JOIN TbRiesgos " & _
                    "ON (TbRiesgosMaterializaciones.IDEdicion = TbRiesgos.IDEdicion) " & _
                    "AND (TbRiesgosMaterializaciones.CodigoRiesgo = TbRiesgos.CodigoRiesgo)) " & _
                    "INNER JOIN TbProyectosEdiciones " & _
                    "ON TbRiesgos.IDEdicion = TbProyectosEdiciones.IDEdicion " & _
                    "WHERE (((TbProyectos.FechaCierre) Is Null) " & _
                    "AND ((TbRiesgosMaterializaciones.IDNC) Is Null) " & _
                    "AND ((TbRiesgosMaterializaciones.ParaNC) Is Null) " & _
                    "AND ((TbProyectos.NombreUsuarioCalidad)='" & m_ObjUsuarioParaTareas.Nombre & "') " & _
                    "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                    "AND ((TbRiesgos.FechaAprobacionAceptacionPorCalidad) Is Null) " & _
                    "AND ((TbRiesgos.FechaAprobacionRetiroPorCalidad) Is Null)) " & _
                    "ORDER BY TbRiesgosMaterializaciones.ID;"
           
        End If
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_RiesgoMaterializado = New RiesgoMaterializacion
            
            For Each m_Campo In m_RiesgoMaterializado.ColCampos
                m_RiesgoMaterializado.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getRiesgosMaterializadosPorDecidir Is Nothing Then
                Set getRiesgosMaterializadosPorDecidir = New Scripting.Dictionary
                getRiesgosMaterializadosPorDecidir.CompareMode = TextCompare
            End If
            If Not getRiesgosMaterializadosPorDecidir.Exists(CStr(m_RiesgoMaterializado.ID)) Then
                getRiesgosMaterializadosPorDecidir.Add CStr(m_RiesgoMaterializado.ID), m_RiesgoMaterializado
            End If

            
            Set m_RiesgoMaterializado = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosMaterializadosPorDecidir ha devuelto el error: " & Err.Description
    End If
End Function
























