VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArbolTareasTecnico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_ObjObjetoSeleccionado As Object


Public Error As String

Private m_ObjArbol As TreeView

Public Property Get ObjetoSeleccionado() As Object

    Set ObjetoSeleccionado = m_ObjObjetoSeleccionado

End Property

Public Property Set ObjetoSeleccionado(ByVal objNewValue As Object)

    Set m_ObjObjetoSeleccionado = objNewValue

End Property


Public Function CargarArbolTareasTecnico( _
                                            ByRef p_Arbol As TreeView, _
                                            ByRef p_ListImages As Object, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    
    Dim m_Key As String
    Dim m_KeyPadre As String
    Dim m_NumeroElementos As Integer
    Dim m_NombreIcono As String
    Dim m_NOmbreIconoExpandido As String
    Dim m_TituloNodo As String
    Dim m_Nodo As MSComctlLib.Node
    Dim m_NodoCategoria As MSComctlLib.Node
    Dim m_ID As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_objEdicion As Edicion
    
    
    On Error GoTo errores
'   EDICIONES QUE REQUIEREN PROPUESTA DE PUBLICACIÓN            EDICIONAPUNTODECADUCARSINPROPUESTA
'   EDICIONES PROPUESTAS RECHAZADAS POR CALIDAD                 EDICIONESPROPUESTASRECHAZADAS
'   RIESGOS ACEPTADOS RECHAZADOS                                RIESGOSACEPTADOSRECHAZADOS
'   RIESGOS RETIRADOS RECHAZADOS                                RIESGOSRETIRADOSRECHAZADOS
    If p_Arbol Is Nothing Then
        Exit Function
    End If
    Set m_ObjArbol = p_Arbol
    If p_ListImages Is Nothing Then
        Exit Function
    End If
    RellenarListaImagenes p_ListImages, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    'Set p_Arbol = New TreeView
    With p_Arbol
        .Style = 7 'tvwTreelinesPlusMinusText
        '.Style = 6 'tvwTreelinesPlusMinusText
        .MousePointer = ccArrow
        .LineStyle = 1 'tvwRootLines
        .LabelEdit = tvwManual
        .ImageList = p_ListImages
        '.BorderStyle = ccFixedSingle
        .BorderStyle = 0
        .Appearance = ccFlat
        .OLEDragMode = ccOLEDragManual
        .OLEDropMode = ccOLEDropNone
        .Indentation = 300
        .PathSeparator = "\"
        .HideSelection = False
        .sorted = False
        .FullRowSelect = True
        .Enabled = True
        .CheckBoxes = False
        .SingleSel = True
        .Scroll = True
        .HotTracking = False
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Refresh
        .Nodes.Clear
        
    End With
    If m_ObjTareasTecnico Is Nothing Then
        EstablecerTareasTecnico EnumSiNo.Sí, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Arbol.Refresh
        CargarArbolTareasTecnico = "OK"
        Exit Function
        
    End If
    
       
    'EDICIONESAPUNTODECADUCARSINPROPUESTA
        
        m_Key = "EDICIONESAPUNTODECADUCARSINPROPUESTA"
        m_KeyPadre = m_Key
        m_NumeroElementos = CInt(m_ObjTareasTecnico.NEdicionesAPuntoDeCaducarSinPropuesta)
        m_TituloNodo = "(" & m_NumeroElementos & ") " & "EDICIONES QUE REQUIEREN PROPUESTA DE PUBLICACIÓN"
        If CLng(m_NumeroElementos) > 0 Then
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
        Else
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaAbierta32)
        End If
        Set m_NodoCategoria = p_Arbol.Nodes.Add(, , m_Key, m_TituloNodo, m_NombreIcono)
        m_NodoCategoria.ExpandedImage = m_NOmbreIconoExpandido
        If CLng(m_NumeroElementos) > 0 Then
            For Each m_ID In m_ObjTareasTecnico.ColEdicionesAPuntoDeCaducarSinPropuesta
                Set m_objEdicion = m_ObjTareasTecnico.ColEdicionesAPuntoDeCaducarSinPropuesta(m_ID)
                
                    m_Key = "EDICIONAPUNTODECADUCARSINPROPUESTA|" & m_objEdicion.IDEdicion
                    m_TituloNodo = m_objEdicion.Proyecto.NombreParaNodo
                   
                    m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
                    Set m_Nodo = p_Arbol.Nodes.Add(m_KeyPadre, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
                Set m_objEdicion = Nothing
            Next
            m_NodoCategoria.Expanded = True
        End If
    'EDICIONESPROPUESTASRECHAZADAS
        
        m_Key = "EDICIONESPROPUESTASRECHAZADAS"
        m_KeyPadre = m_Key
        m_NumeroElementos = CInt(m_ObjTareasTecnico.NEdicionesPropuestasRechazadas)
        m_TituloNodo = "(" & m_NumeroElementos & ") " & "EDICIONES PROPUESTAS RECHAZADAS POR CALIDAD"
        If CLng(m_NumeroElementos) > 0 Then
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
        Else
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaAbierta32)
        End If
        Set m_NodoCategoria = p_Arbol.Nodes.Add(, , m_Key, m_TituloNodo, m_NombreIcono)
        m_NodoCategoria.ExpandedImage = m_NOmbreIconoExpandido
        If CLng(m_NumeroElementos) > 0 Then
            For Each m_ID In m_ObjTareasTecnico.ColEdicionesPropuestasRechazadas
                Set m_objEdicion = m_ObjTareasTecnico.ColEdicionesPropuestasRechazadas(m_ID)
                
                    m_Key = "EDICIONPROPUESTARECHAZADA|" & m_objEdicion.IDEdicion
                    m_TituloNodo = m_objEdicion.Proyecto.NombreParaNodo
                   
                    m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
                    Set m_Nodo = p_Arbol.Nodes.Add(m_KeyPadre, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
                Set m_objEdicion = Nothing
            Next
            m_NodoCategoria.Expanded = True
        End If
    
   
    
    'RIESGOSACEPTADOSRECHAZADOS
        m_Key = "RIESGOSACEPTADOSRECHAZADOS"
        m_KeyPadre = m_Key
        m_NumeroElementos = CInt(m_ObjTareasTecnico.NRiesgosAceptadosRechazados)
        m_TituloNodo = "(" & m_NumeroElementos & ") " & "RIESGOS ACEPTADOS RECHAZADOS"
        If CLng(m_NumeroElementos) > 0 Then
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
        Else
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaAbierta32)
        End If
        Set m_NodoCategoria = p_Arbol.Nodes.Add(, , m_Key, m_TituloNodo, m_NombreIcono)
        m_NodoCategoria.ExpandedImage = m_NOmbreIconoExpandido
        If CLng(m_NumeroElementos) > 0 Then
            For Each m_ID In m_ObjTareasTecnico.ColRiesgosAceptadosRechazados
                Set m_ObjRiesgo = m_ObjTareasTecnico.ColRiesgosAceptadosRechazados(m_ID)
                    m_Key = "RIESGOACEPTADORECHAZADO|" & m_ObjRiesgo.IDRiesgo
                    m_TituloNodo = m_ObjRiesgo.NombreParaNodo
                    m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
                    Set m_Nodo = p_Arbol.Nodes.Add(m_KeyPadre, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
                Set m_ObjRiesgo = Nothing
            Next
            m_NodoCategoria.Expanded = True
        End If
    
    'RIESGOSRETIRADOSRECHAZADOS
        m_Key = "RIESGOSRETIRADOSRECHAZADOS"
        m_KeyPadre = m_Key
        m_NumeroElementos = CInt(m_ObjTareasTecnico.NRiesgosRetiradosRechazados)
        m_TituloNodo = "(" & m_NumeroElementos & ") " & "RIESGOS RETIRADOS RECHAZADOS"
        If CLng(m_NumeroElementos) > 0 Then
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
        Else
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)
        End If
        Set m_NodoCategoria = p_Arbol.Nodes.Add(, , m_Key, m_TituloNodo, m_NombreIcono)
        
        If CLng(m_NumeroElementos) > 0 Then
            For Each m_ID In m_ObjTareasTecnico.ColRiesgosRetiradosRechazados
                Set m_ObjRiesgo = m_ObjTareasTecnico.ColRiesgosRetiradosRechazados(m_ID)
                    m_Key = "RIESGORETIRADORECHAZADO|" & m_ObjRiesgo.IDRiesgo
                    m_TituloNodo = m_ObjRiesgo.NombreParaNodo
                    m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
                    Set m_Nodo = p_Arbol.Nodes.Add(m_KeyPadre, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
                Set m_ObjRiesgo = Nothing
            Next
            m_NodoCategoria.Expanded = True
        End If
    
    
    p_Arbol.Refresh
    CargarArbolTareasTecnico = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarArbolTareasTecnico ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function TipoNodoTareaSeleccionado( _
                                            ByVal p_Key As String, _
                                            Optional ByRef p_Error As String _
                                            ) As EnumTipoNodoTareasTecnico
    Dim m_valor As String
    Dim m_ID As String
    Dim dato
    On Error GoTo errores

    
    If p_Key = "" Then
        Exit Function
    End If
    If InStr(1, p_Key, "EDICIONESAPUNTODECADUCARSINPROPUESTA") <> 0 Then
        
        Set Me.ObjetoSeleccionado = Nothing
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasTecnico.EDICIONESAPUNTODECADUCARSINPROPUESTA
        
        Exit Function
    End If
    If InStr(1, p_Key, "EDICIONAPUNTODECADUCARSINPROPUESTA") <> 0 Then
        
        dato = Split(p_Key, "|")
        m_ID = dato(1)
        Set Me.ObjetoSeleccionado = Constructor.getEdicion(m_ID, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasTecnico.EDICIONAPUNTODECADUCARSINPROPUESTA
        
        Exit Function
    End If
    If InStr(1, p_Key, "EDICIONESPROPUESTASRECHAZADAS") <> 0 Then
        
        Set Me.ObjetoSeleccionado = Nothing
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasTecnico.EDICIONESPROPUESTASRECHAZADAS
        
        Exit Function
    End If
    If InStr(1, p_Key, "EDICIONPROPUESTARECHAZADA") <> 0 Then
        
        dato = Split(p_Key, "|")
        m_ID = dato(1)
        Set Me.ObjetoSeleccionado = Constructor.getEdicion(m_ID, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasTecnico.EDICIONPROPUESTARECHAZADA
        
        Exit Function
    End If
    
    If InStr(1, p_Key, "RIESGOSACEPTADOSRECHAZADOS") <> 0 Then
        
        Set Me.ObjetoSeleccionado = Nothing
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasTecnico.RIESGOSACEPTADOSRECHAZADOS
        
        Exit Function
    End If
    If InStr(1, p_Key, "RIESGOACEPTADORECHAZADO") <> 0 Then
        
        dato = Split(p_Key, "|")
        m_ID = dato(1)
        Set Me.ObjetoSeleccionado = Constructor.getRiesgo(m_ID, , , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasTecnico.RIESGOACEPTADORECHAZADO
        
        Exit Function
    End If
    If InStr(1, p_Key, "RIESGOSRETIRADOSRECHAZADOS") <> 0 Then
        
        Set Me.ObjetoSeleccionado = Nothing
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasTecnico.RIESGOSRETIRADOSRECHAZADOS
        
        Exit Function
    End If
    If InStr(1, p_Key, "RIESGORETIRADORECHAZADO") <> 0 Then
        
        dato = Split(p_Key, "|")
        m_ID = dato(1)
        Set Me.ObjetoSeleccionado = Constructor.getRiesgo(m_ID, , , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasTecnico.RIESGORETIRADORECHAZADO
        
        Exit Function
    End If
    
    
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Arbol.TipoNodoTareaSeleccionado ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


