VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArbolTareasCalidad"
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


Public Function CargarArbolTareasCalidad( _
                                            ByRef p_Arbol As MSComctlLib.TreeView, _
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
    Dim m_ObjRiesgoMaterializado As RiesgoMaterializacion
    Dim m_ColEdCercanasACaducarResultante As Scripting.Dictionary
    Dim m_ColEdCaducadasResultante As Scripting.Dictionary
    
    Dim m_objEdicion As Edicion
    
    
    On Error GoTo errores
'   EDICIONES PREPARADAS PARA PUBLICAR  EDICIONESPREPARADASPARAPUBLICAR
'   EDICIONES CADUCADAS                             EDICIONESCADUCADAS
'   EDICIONES PRÓXIMAS A CADUCAR                    EDICIONESAPUNTODECADUCAR
'   RIESGOS POR RETIPIFICAR                         RIESGOSPENDIENTESRETIPIFICAR
'   RIESGOS ACEPTADOS POR VISAR                     RIESGOSACEPTADOSPORVISAR
'   RIESGOS RETIRADOS POR VISAR                     RIESGOSRETIRADOSPORVISAR
'   RIESGOS MATERIALIZADOS POR DECIDIR              RIESGOSMATERIALIZADOSPORDECIDIR


    p_Error = ""
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
    On Error Resume Next
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
        .Font.Size = 9
        .Refresh
        .Nodes.Clear
        
    End With
    If Err.Number Then Err.Clear
    On Error GoTo errores
    If m_ObjTareasCalidad Is Nothing Then
        EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.Sí, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Arbol.Refresh
        CargarArbolTareasCalidad = "OK"
        Exit Function
        
    End If
    Set m_ColEdCercanasACaducarResultante = getColCercanasACaducarResultante( _
                                        m_ObjTareasCalidad.ColEdicionesApuntoDeCaducar, _
                                        m_ObjTareasCalidad.ColEdicionesPreparadasParaPublicar, _
                                        p_Error)
    Set m_ColEdCaducadasResultante = getColCercanasACaducarResultante( _
                                        m_ObjTareasCalidad.ColEdicionesCaducadas, _
                                        m_ObjTareasCalidad.ColEdicionesPreparadasParaPublicar, _
                                        p_Error)

        
    'EDICIONESPREPARADASPARAPUBLICAR
        
        m_Key = "EDICIONESPREPARADASPARAPUBLICAR"
        m_KeyPadre = m_Key
        m_NumeroElementos = CInt(m_ObjTareasCalidad.NEdicionesPreparadasParaPublicar)
        m_TituloNodo = "(" & m_NumeroElementos & ") " & "EDICIONES PREPARADAS PARA PUBLICAR"
        If CLng(m_NumeroElementos) > 0 Then
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
        Else
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaAbierta32)
        End If
        Set m_NodoCategoria = p_Arbol.Nodes.Add(, , m_Key, m_TituloNodo, m_NombreIcono)
        m_NodoCategoria.ExpandedImage = m_NOmbreIconoExpandido
        If CLng(m_NumeroElementos) > 0 Then
            For Each m_ID In m_ObjTareasCalidad.ColEdicionesPreparadasParaPublicar
                Set m_objEdicion = m_ObjTareasCalidad.ColEdicionesPreparadasParaPublicar(m_ID)
                    m_Key = "EDICIONPREPARADAPARAPUBLICAR|" & m_objEdicion.IDEdicion
                    m_TituloNodo = m_objEdicion.NombreParaNodo1
                    m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
                    Set m_Nodo = p_Arbol.Nodes.Add(m_KeyPadre, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
                Set m_objEdicion = Nothing
            Next
            m_NodoCategoria.Expanded = True
        End If
    
    'EDICIONESCADUCADAS
        
        m_Key = "EDICIONESCADUCADAS"
        m_KeyPadre = m_Key
        If m_ColEdCaducadasResultante Is Nothing Then
            m_NumeroElementos = 0
        Else
            m_NumeroElementos = CInt(m_ColEdCaducadasResultante.Count)
        End If
        
        
        m_TituloNodo = "(" & m_NumeroElementos & ") " & "EDICIONES CADUCADAS"
        If CLng(m_NumeroElementos) > 0 Then
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
        Else
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaAbierta32)
        End If
        Set m_NodoCategoria = p_Arbol.Nodes.Add(, , m_Key, m_TituloNodo, m_NombreIcono)
        m_NodoCategoria.ExpandedImage = m_NOmbreIconoExpandido
        
        If CLng(m_NumeroElementos) > 0 Then
            For Each m_ID In m_ColEdCaducadasResultante
                Set m_objEdicion = m_ColEdCaducadasResultante(m_ID)
                
                    m_Key = "EDICIONCADUCADA|" & m_objEdicion.IDEdicion
                    m_TituloNodo = m_objEdicion.NombreParaNodo1
                   
                    m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
                    Set m_Nodo = p_Arbol.Nodes.Add(m_KeyPadre, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
                Set m_objEdicion = Nothing
            Next
            m_NodoCategoria.Expanded = True
        End If
    'EDICIONESAPUNTODECADUCAR
        
        m_Key = "EDICIONESAPUNTODECADUCAR"
        m_KeyPadre = m_Key
        If m_ColEdCercanasACaducarResultante Is Nothing Then
            m_NumeroElementos = 0
        Else
            m_NumeroElementos = CInt(m_ColEdCercanasACaducarResultante.Count)
        End If
        
        m_TituloNodo = "(" & m_NumeroElementos & ") " & "EDICIONES PRÓXIMAS A CADUCAR"
        If CLng(m_NumeroElementos) > 0 Then
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
        Else
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaAbierta32)
        End If
        Set m_NodoCategoria = p_Arbol.Nodes.Add(, , m_Key, m_TituloNodo, m_NombreIcono)
        m_NodoCategoria.ExpandedImage = m_NOmbreIconoExpandido
        
        If CLng(m_NumeroElementos) > 0 Then
            For Each m_ID In m_ColEdCercanasACaducarResultante
                Set m_objEdicion = m_ColEdCercanasACaducarResultante(m_ID)
                
                    m_Key = "EDICIONAPUNTODECADUCAR|" & m_objEdicion.IDEdicion
                    m_TituloNodo = m_objEdicion.NombreParaNodo1
                   
                    m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
                    Set m_Nodo = p_Arbol.Nodes.Add(m_KeyPadre, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
                Set m_objEdicion = Nothing
            Next
            m_NodoCategoria.Expanded = True
        End If
    
   
    'RIESGOSPENDIENTESRETIPIFICAR
        
        m_Key = "RIESGOSPENDIENTESRETIPIFICAR"
        m_KeyPadre = m_Key
        m_NumeroElementos = CInt(m_ObjTareasCalidad.NRiesgosPendientesDeTipificar)
        m_TituloNodo = "(" & m_NumeroElementos & ") " & "RIESGOS POR RETIPIFICAR"
        If CLng(m_NumeroElementos) > 0 Then
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
        Else
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaAbierta32)
        End If
        Set m_NodoCategoria = p_Arbol.Nodes.Add(, , m_Key, m_TituloNodo, m_NombreIcono)
        m_NodoCategoria.ExpandedImage = m_NOmbreIconoExpandido
        
        If CLng(m_NumeroElementos) > 0 Then
            For Each m_ID In m_ObjTareasCalidad.ColRiesgosPorReTipificar
                Set m_ObjRiesgo = m_ObjTareasCalidad.ColRiesgosPorReTipificar(m_ID)
                    m_Key = "RIESGOPENDIENTERETIPIFICAR|" & m_ObjRiesgo.IDRiesgo
                    m_TituloNodo = m_ObjRiesgo.NombreParaNodo & " (" & UCase(m_ObjRiesgo.Edicion.Proyecto.NombreUsuarioCalidadParaNodo) & ")"
                    
                    m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
                    Set m_Nodo = p_Arbol.Nodes.Add(m_KeyPadre, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
                Set m_ObjRiesgo = Nothing
            Next
            m_NodoCategoria.Expanded = True
        End If
    'RIESGOSACEPTADOSPORVISAR
        m_Key = "RIESGOSACEPTADOSPORVISAR"
        m_KeyPadre = m_Key
        m_NumeroElementos = CInt(m_ObjTareasCalidad.NRiesgosAceptadosSinVisar)
        m_TituloNodo = "(" & m_NumeroElementos & ") " & "RIESGOS ACEPTADOS POR VISAR"
        If CLng(m_NumeroElementos) > 0 Then
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
        Else
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaAbierta32)
        End If
        Set m_NodoCategoria = p_Arbol.Nodes.Add(, , m_Key, m_TituloNodo, m_NombreIcono)
        m_NodoCategoria.ExpandedImage = m_NOmbreIconoExpandido
        
        If CLng(m_NumeroElementos) > 0 Then
            For Each m_ID In m_ObjTareasCalidad.ColRiesgosAceptadosSinVisar
                Set m_ObjRiesgo = m_ObjTareasCalidad.ColRiesgosAceptadosSinVisar(m_ID)
                    m_Key = "RIESGOACEPTADOPORVISAR|" & m_ObjRiesgo.IDRiesgo
                    m_TituloNodo = m_ObjRiesgo.NombreParaNodo & " (" & UCase(m_ObjRiesgo.Edicion.Proyecto.NombreUsuarioCalidadParaNodo) & ")"
                    m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
                    Set m_Nodo = p_Arbol.Nodes.Add(m_KeyPadre, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
                Set m_ObjRiesgo = Nothing
            Next
            m_NodoCategoria.Expanded = True
        End If
    
    'RIESGOSRETIRADOSPORVISAR
        m_Key = "RIESGOSRETIRADOSPORVISAR"
        m_KeyPadre = m_Key
        m_NumeroElementos = CInt(m_ObjTareasCalidad.NRiesgosRetiradosSinVisar)
        m_TituloNodo = "(" & m_NumeroElementos & ") " & "RIESGOS RETIRADOS POR VISAR  "
        If CLng(m_NumeroElementos) > 0 Then
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
        Else
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaAbierta32)
        End If
        Set m_NodoCategoria = p_Arbol.Nodes.Add(, , m_Key, m_TituloNodo, m_NombreIcono)
        m_NodoCategoria.ExpandedImage = m_NOmbreIconoExpandido
        
        If CLng(m_NumeroElementos) > 0 Then
            For Each m_ID In m_ObjTareasCalidad.ColRiesgosRetiradosSinVisar
                Set m_ObjRiesgo = m_ObjTareasCalidad.ColRiesgosRetiradosSinVisar(m_ID)
                    m_Key = "RIESGORETIRADOPORVISAR|" & m_ObjRiesgo.IDRiesgo
                    m_TituloNodo = m_ObjRiesgo.NombreParaNodo & " (" & UCase(m_ObjRiesgo.Edicion.Proyecto.NombreUsuarioCalidadParaNodo) & ")"
                    m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
                    Set m_Nodo = p_Arbol.Nodes.Add(m_KeyPadre, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
                Set m_ObjRiesgo = Nothing
            Next
            m_NodoCategoria.Expanded = True
        End If
    
    'RIESGOSMATERIALIZADOSPORDECIDIR
        m_Key = "RIESGOSMATERIALIZADOSPORDECIDIR"
        m_KeyPadre = m_Key
        m_NumeroElementos = CInt(m_ObjTareasCalidad.NRiesgosMaterializadosPorDedicir)
        m_TituloNodo = "(" & m_NumeroElementos & ") " & "RIESGOS MATERIALIZADOS POR DECIDIR"
        If CLng(m_NumeroElementos) > 0 Then
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaIncompletaAbierta32)
        Else
            m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)
            m_NOmbreIconoExpandido = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaAbierta32)
        End If
        Set m_NodoCategoria = p_Arbol.Nodes.Add(, , m_Key, m_TituloNodo, m_NombreIcono)
        m_NodoCategoria.ExpandedImage = m_NOmbreIconoExpandido
        If CLng(m_NumeroElementos) > 0 Then
            For Each m_ID In m_ObjTareasCalidad.ColRiesgosMaterializadosPorDedicir
                Set m_ObjRiesgoMaterializado = m_ObjTareasCalidad.ColRiesgosMaterializadosPorDedicir(m_ID)
                m_Key = "RIESGOMATERIALIZADOPORDECIDIR|" & m_ObjRiesgoMaterializado.ID
                If Not m_ObjRiesgoMaterializado.riesgo Is Nothing Then
                    m_TituloNodo = m_ObjRiesgoMaterializado.riesgo.CodigoRiesgo & " ( " & m_ObjRiesgoMaterializado.Proyecto.NombreParaNodo & " )" & " ( Edición " & m_ObjRiesgoMaterializado.Edicion.Edicion & " )" & " (" & UCase(m_ObjRiesgoMaterializado.Proyecto.NombreUsuarioCalidadParaNodo) & ")"
                    m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoItemInCompleto32)
                    Set m_Nodo = p_Arbol.Nodes.Add(m_KeyPadre, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
                Else
                    Stop
                End If
                
                Set m_ObjRiesgoMaterializado = Nothing
            Next
            m_NodoCategoria.Expanded = True
        End If
    p_Arbol.Refresh
    CargarArbolTareasCalidad = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarArbolTareasCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function





Public Function TipoNodoTareaSeleccionado( _
                                            ByVal p_Key As String, _
                                            Optional ByRef p_Error As String _
                                            ) As EnumTipoNodoTareasCalidad
    Dim m_valor As String
    Dim m_ID As String
    Dim dato
    On Error GoTo errores

    
    If p_Key = "" Then
        Exit Function
    End If
    If InStr(1, p_Key, "EDICIONCADUCADA") <> 0 Then
        
        dato = Split(p_Key, "|")
        m_ID = dato(1)
        Set Me.ObjetoSeleccionado = Constructor.getEdicion(m_ID, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.EDICIONCADUCADA
        
        Exit Function
    End If
    If InStr(1, p_Key, "EDICIONESCADUCADAS") <> 0 Then
        Set Me.ObjetoSeleccionado = Nothing
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.EDICIONESCADUCADAS
        
        Exit Function
    End If
    If InStr(1, p_Key, "EDICIONAPUNTODECADUCAR") <> 0 Then
        
        dato = Split(p_Key, "|")
        m_ID = dato(1)
        Set Me.ObjetoSeleccionado = Constructor.getEdicion(m_ID, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.EDICIONAPUNTODECADUCAR
        
        Exit Function
    End If
    If InStr(1, p_Key, "EDICIONESAPUNTODECADUCAR") <> 0 Then
        Set Me.ObjetoSeleccionado = Nothing
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.EDICIONESAPUNTODECADUCAR
        
        Exit Function
    End If
    If InStr(1, p_Key, "EDICIONAPUNTODECADUCAR") <> 0 Then
        
        dato = Split(p_Key, "|")
        m_ID = dato(1)
        Set Me.ObjetoSeleccionado = Constructor.getEdicion(m_ID, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.EDICIONAPUNTODECADUCAR
        
        Exit Function
    End If
    If InStr(1, p_Key, "EDICIONESPREPARADASPARAPUBLICAR") <> 0 Then
        
        Set Me.ObjetoSeleccionado = Nothing
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.EDICIONESPREPARADASPARAPUBLICAR
        Exit Function
    End If
    If InStr(1, p_Key, "EDICIONPREPARADAPARAPUBLICAR") <> 0 Then
        
        dato = Split(p_Key, "|")
        m_ID = dato(1)
        Set Me.ObjetoSeleccionado = Constructor.getEdicion(m_ID, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.EDICIONPREPARADAPARAPUBLICAR
        
        Exit Function
    End If
    If InStr(1, p_Key, "RIESGOSPENDIENTESRETIPIFICAR") <> 0 Then
        
        Set Me.ObjetoSeleccionado = Nothing
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.RIESGOSPENDIENTESRETIPIFICAR
        
        Exit Function
    End If
    If InStr(1, p_Key, "RIESGOPENDIENTERETIPIFICAR") <> 0 Then
        
        dato = Split(p_Key, "|")
        m_ID = dato(1)
        Set Me.ObjetoSeleccionado = Constructor.getRiesgo(m_ID, , , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.RIESGOPENDIENTERETIPIFICAR
        
        Exit Function
    End If
    If InStr(1, p_Key, "RIESGOSACEPTADOSPORVISAR") <> 0 Then
        
        Set Me.ObjetoSeleccionado = Nothing
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.RIESGOSACEPTADOSPORVISAR
        
        Exit Function
    End If
    If InStr(1, p_Key, "RIESGOACEPTADOPORVISAR") <> 0 Then
        
        dato = Split(p_Key, "|")
        m_ID = dato(1)
        Set Me.ObjetoSeleccionado = Constructor.getRiesgo(m_ID, , , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.RIESGOACEPTADOPORVISAR
        
        Exit Function
    End If
    If InStr(1, p_Key, "RIESGOSRETIRADOSPORVISAR") <> 0 Then
        
        Set Me.ObjetoSeleccionado = Nothing
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.RIESGOSRETIRADOSPORVISAR
        
        Exit Function
    End If
    If InStr(1, p_Key, "RIESGORETIRADOPORVISAR") <> 0 Then
        
        dato = Split(p_Key, "|")
        m_ID = dato(1)
        Set Me.ObjetoSeleccionado = Constructor.getRiesgo(m_ID, , , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.RIESGORETIRADOPORVISAR
        
        Exit Function
    End If
    
    
    If InStr(1, p_Key, "RIESGOSMATERIALIZADOSPORDECIDIR") <> 0 Then
        
        Set Me.ObjetoSeleccionado = Nothing
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.RIESGOSMATERIALIZADOSPORDECIDIR
        
        Exit Function
    End If
    If InStr(1, p_Key, "RIESGOMATERIALIZADOPORDECIDIR") <> 0 Then
        
        dato = Split(p_Key, "|")
        m_ID = dato(1)
        Set Me.ObjetoSeleccionado = Constructor.getRiesgoMaterializado(m_ID, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        TipoNodoTareaSeleccionado = EnumTipoNodoTareasCalidad.RIESGOMATERIALIZADOPORDECIDIR
        
        Exit Function
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Arbol.TipoNodoTareaSeleccionado ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


