

Option Compare Database
Option Explicit

' ==============================================================================
' FORMULARIO: Gestion de Suministradores (Arbol)
' CARACTERISTICAS:
'   - Soporta Drag & Drop desde catalogo (derecha) y reordenacion interna.
'   - Modo Hibrido: Persistencia en BD (Edicion) o en Memoria DTO (Alta).
'   - Integracion con Servicio Transaccional.
'   - PRUEBA-003/REFAC-4b2b (closes #60): thinned RefrescarAArbol delegates
'     tree state to Helper_ExpedienteSuministradoresArbol (pure model).
' ==============================================================================

' --- Constantes de Logica del Arbol ---
Private Const K_ROOT_ORGANO As String = "ROOT_ORGANO"

' --- Constantes de Imagenes (Deben coincidir con los ficheros .ico en /Iconos) ---
Private Const IMG_KEY_CONTRATISTA As String = "ICON_ORGANO_CONTROL"
Private Const IMG_KEY_EMPRESA As String = "ICON_EMPRESA_NO_GRUPO"
Private Const IMG_KEY_EMPRESA_GRUPO As String = "ICON_EMPRESA_GRUPO"
Private Const IMG_KEY_EMPRESA_LISTA As String = "ICON_EMPRESA_LISTA"

' --- Constantes para Drag & Drop y Menus ---
Private Const PREFIX_SRC As String = "SRC_" ' Identifica origen desde lista derecha
Private Const TV_SELECTED_VAR As String = "SelectedIDSuministrador" ' TempVar para dialogos

' --- Estado ---
Private m_IDExpediente As String

Private m_Error As String
Private m_Servicio As ExpedienteSuministradorServicio

' PRUEBA-003/REFAC-4b2b: pure-model helper that owns the supplier-chain
' tree state. Initialized lazily on first Form_Open, re-initialized by
' RefrescarAArbol. The form is now a thin wiring layer over this helper.
Private m_Helper As Helper_ExpedienteSuministradoresArbol

Private WithEvents m_FormAltaSuministrador As Form_FormSuministrador

' PRUEBA-003/REFAC-4b2b: test seam for atoms 9 wiring. Allows the test
' atom to inject m_IDExpediente without opening the form (the form's
' normal initialization path is Form_Open, which requires UI state).
Public Property Let IDExpediente_ForTest(ByVal p_Value As String)
    m_IDExpediente = p_Value
End Property

' PRUEBA-003/REFAC-4b2b: strategy (c) read-back seam. Returns the
' canonical JSON envelope from m_Helper.Serializar(m_Error) so tests can
' observe the helper state without exposing the helper instance itself.
Public Property Get NodosSerializados_() As String
    Dim p_Error As String
    p_Error = ""
    If m_Helper Is Nothing Then
        NodosSerializados_ = "{""ok"":false,""value"":null,""payload"":null,""error"":""helper-not-initialized"",""logs"":[]}"
        Exit Property
    End If
    NodosSerializados_ = m_Helper.Serializar(p_Error)
End Property

Private Function ParsePayload(ByVal p_Json As String) As Object
    On Error GoTo EH
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(p_Json)
    If parsed Is Nothing Then Set ParsePayload = Nothing: Exit Function
    If parsed.Exists("payload") Then Set ParsePayload = parsed("payload") Else Set ParsePayload = Nothing
    Exit Function
EH:
    Set ParsePayload = Nothing
End Function



' ==============================================================================
' 1. INICIALIZACION Y CIERRE
' ==============================================================================

Private Sub Form_Open(Cancel As Integer)
    Dim ctl As Control
    On Error GoTo errores
    m_Error = ""
    DoCmd.Hourglass True
    
    ' Validar Contexto DTO
    If m_ObjExpedienteDTOActivo Is Nothing Then
        m_Error = "No hay contexto de expediente activo."
        Err.Raise 1000
    End If
    Me.Parent.Controls("ComandoRegistrar").Visible = False
    ' Inicializar Estaado
    m_IDExpediente = m_ObjExpedienteDTOActivo.Expediente.IDExpediente

    Set m_Servicio = New ExpedienteSuministradorServicio

    ' PRUEBA-003/REFAC-4b2b: instantiate the pure-model helper so RefrescarAArbol
    ' can delegate tree-state construction to it.
    Set m_Helper = New Helper_ExpedienteSuministradoresArbol
    
    ' Configurar Imagenes (UI)
    RellenarListaImagenes Me.ListaImagenes, m_Error
    If m_Error <> "" Then Err.Raise 1000
    Dim imgList As Object
    Set imgList = Me.ListaImagenes.Object
    On Error Resume Next
    imgList.ListImages.Remove IMG_KEY_CONTRATISTA
    imgList.ListImages.Remove IMG_KEY_EMPRESA
    imgList.ListImages.Remove IMG_KEY_EMPRESA_GRUPO
    imgList.ListImages.Remove IMG_KEY_EMPRESA_LISTA
    On Error GoTo errores
    imgList.ListImages.Add , IMG_KEY_CONTRATISTA, LoadPicture(m_ObjEntorno.URLIconoCadenaContratacionOrganoControl)
    imgList.ListImages.Add , IMG_KEY_EMPRESA, LoadPicture(m_ObjEntorno.URLIconoCadenaContratacionEmpresaNoGrupo)
    imgList.ListImages.Add , IMG_KEY_EMPRESA_GRUPO, LoadPicture(m_ObjEntorno.URLIconoCadenaContratacionEmpresaGrupo)
    imgList.ListImages.Add , IMG_KEY_EMPRESA_LISTA, LoadPicture(m_ObjEntorno.URLIconoEmpresaLista)
    
    ' Vincular ImageList
    Me.tvCadena.ImageList = Me.ListaImagenes.Object
    Me.lvFuente.SmallIcons = Me.ListaImagenes.Object
    
    ' Preparar Controles
    PrepararTreeView
    ConfigurarListView
    
    ' Carga Inicial
    RefrescarAArbol
    CargarListaFuente "", False
    
    ' Gestionar Permisos (En Alta siempre es editable la memoria)
   
    Me.AllowEdits = False
    If AbiertoParaEditar = True Then
        If EsAdministrador = EnumSiNo.Sí Then
            Me.AllowEdits = True
        End If
    End If
   For Each ctl In Me.Controls
        If Nz(ctl.Tag, "") = "EJECUTIVO" Then
            If AbiertoParaEditar = True Then
                If EsAdministrador = EnumSiNo.Sí Then
                    ctl.Enabled = True
                Else
                    ctl.Enabled = False
                End If
            Else
                ctl.Enabled = False
            End If
        End If
    Next
    DoCmd.Hourglass False
   
    ' Estaado inicial botones detalle
    GestionarBotones Nothing
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Error al abrir subformulario Suministradores: " & m_Error, vbCritical, "Error"
    Cancel = True
End Sub

Private Sub Form_Close()
    Set m_Servicio = Nothing
End Sub

' ==============================================================================
' 2. GESTI_N DEL CAT_LOGO (DERECHA - LISTVIEW)
' ==============================================================================

Private Sub ConfigurarListView()
    On Error Resume Next
    Dim anchoColumna As Long
    
    With Me.lvFuente
        .View = 3 ' lvwReport
        .HideSelection = False
        .OLEDragMode = 1 ' ccOLEDragAutomatic
        .ColumnHeaders.Clear
        ' Ajuste dinamico de ancho
        anchoColumna = .Width - 300
        If anchoColumna < 1000 Then anchoColumna = 2000
        .ColumnHeaders.Add , , "Empresa / Suministrador", anchoColumna
    End With
End Sub

Private Sub lvFuente_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
    If Button = 2 Then
        Dim hitItem As Object
        Set hitItem = Me.lvFuente.HitTest(x, y)
        If Not hitItem Is Nothing Then
            Set Me.lvFuente.SelectedItem = hitItem
            MostrarMenuContextual "LISTA", hitItem
        End If
    End If
End Sub


Private Sub tvCadena_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
    If Button = 2 Then
        Dim hitNode As Object
        Set hitNode = Me.tvCadena.HitTest(x, y)
        If Not hitNode Is Nothing Then
            Set Me.tvCadena.SelectedItem = hitNode
            GestionarBotones hitNode
            MostrarMenuContextual "ARBOL", hitNode
        End If
    End If
End Sub

Private Sub txtFiltro_Change()
    ' Filtro al escribir
    CargarListaFuente Me.txtFiltro.Text, Nz(Me.chkNuestras.value, False)
End Sub

Private Sub chkNuestras_Click()
    ' Filtro checkbox
    Dim sTexto As String
    If Me.ActiveControl.Name = "txtFiltro" Then sTexto = Me.txtFiltro.Text Else sTexto = Nz(Me.txtFiltro.value, "")
    CargarListaFuente sTexto, Nz(Me.chkNuestras.value, False)
End Sub

Private Sub CargarListaFuente(sTexto As String, bSoloPropias As Boolean)
    Dim dict As Scripting.Dictionary
    Dim sSum As Suministrador
    Dim k As Variant
    Dim li As Object
    Dim sKey As String
    
    On Error Resume Next
    Me.lvFuente.ListItems.Clear
    
    Set dict = constructor.getSuministradoresFiltrados(sTexto, bSoloPropias, m_Error)
    If dict Is Nothing Then Exit Sub
    
    For Each k In dict.Keys
        Set sSum = dict(k)
        sKey = PREFIX_SRC & sSum.IDSuministrador
        
        Dim sNombreDisplay As String
        Dim sImgKey As String
        If sSum.Nemotecnico <> "" Then
            sNombreDisplay = sSum.Nemotecnico
        Else
            sNombreDisplay = sSum.Nombre
        End If
        sImgKey = IMG_KEY_EMPRESA_LISTA
        
        Set li = Me.lvFuente.ListItems.Add(, sKey, sNombreDisplay, , sImgKey)
        li.Tag = sSum.IDSuministrador ' ID Real
    Next
End Sub

' ==============================================================================
' 3. GESTI_N DEL ARBOL (IZQUIERDA - TREEVIEW) - LOGICA HIBRIDA
' ==============================================================================

Private Sub PrepararTreeView()
    On Error Resume Next
    With Me.tvCadena
        .Nodes.Clear
        .LineStyle = 1
        .Indentation = 300
        .HideSelection = False
        .LabelEdit = 1
        .OLEDragMode = 1
        .OLEDropMode = 1
    End With
End Sub

Public Sub RefrescarAArbol()
    Dim p_Error As String
    Dim colDatos As Collection
    Dim itemsDict As Object
    Dim sNombreOrgano As String

    On Error GoTo errores

    ' Resolve organ name (best-effort; m_ObjExpedienteDTOActivo may be Nothing in tests)
    sNombreOrgano = ResolveOrganoNombre()

    ' 1. DAO adapter (testable: p_Db injectable in fixtures)
    Set colDatos = ExpedienteSuministradorRepositorio.GetDatosArbolTestable( _
        CLng(Nz(m_IDExpediente, 0)), Nothing, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    ' 2. Build Dictionary that m_Helper.Cargar expects
    Set itemsDict = BuildHelperItemsFromCollection(colDatos, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    ' 3. Wire the pure-model helper (Init + Cargar)
    If m_Helper Is Nothing Then Set m_Helper = New Helper_ExpedienteSuministradoresArbol
    m_Helper.Init sNombreOrgano, p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    m_Helper.Cargar itemsDict, p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    ' 4. Paint UI from the helper (best-effort: Me.tvCadena requires OpenForm)
    On Error Resume Next
    PintarNodosDesdeHelper
    ActualizarEtiquetaSituacion
    On Error GoTo errores
    Exit Sub

errores:
    MsgBox "Error RefrescarAArbol: " & Err.Description, vbCritical, "Error"
End Sub

' Resolves the organ name from m_ObjExpedienteDTOActivo.Expediente.OrganoContratacion
' (a nested property: Expediente.OrganoContratacion returns OrganoContratacion
' object; that object's .OrganoContratacion is the display string). Falls
' back to "DESCONOCIDO" when the DTO or its OrganoContratacion is unavailable.
Private Function ResolveOrganoNombre() As String
    Dim oOrgano As OrganoContratacion
    On Error Resume Next
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        If Not m_ObjExpedienteDTOActivo.Expediente Is Nothing Then
            Set oOrgano = m_ObjExpedienteDTOActivo.Expediente.OrganoContratacion
            If Not oOrgano Is Nothing Then
                ResolveOrganoNombre = Nz(oOrgano.OrganoContratacion, "DESCONOCIDO")
                If ResolveOrganoNombre = "" Then ResolveOrganoNombre = "DESCONOCIDO"
            Else
                ResolveOrganoNombre = "DESCONOCIDO"
            End If
        End If
    End If
    On Error GoTo 0
End Function

' Translates Collection<ExpedienteSuministrador> from GetDatosArbolTestable into
' the Dictionary shape that Helper_ExpedienteSuministradoresArbol.Cargar expects:
'   keys = arbitrary (we use "k1", "k2", ...)
'   values = Dictionary(IDExpedienteSuministrador, IdPadre, Text, Tag, EsEmpresaPropia)
'   Tag = "RELID=<id>;IDS=<sumID>" — preserves drag&drop semantics from #58.
Private Function BuildHelperItemsFromCollection( _
    ByVal p_Col As Collection, _
    ByRef p_Error As String _
) As Object
    Dim dict As Object
    Dim item As ExpedienteSuministrador
    Dim dictItem As Object
    Dim vIdPadre As Variant
    Dim sIdPadreJson As String
    Dim i As Long

    On Error GoTo errores
    p_Error = ""

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = TextCompare

    If p_Col Is Nothing Then
        Set BuildHelperItemsFromCollection = dict
        Exit Function
    End If

    For i = 1 To p_Col.Count
        Set item = p_Col(i)
        If item Is Nothing Then
            p_Error = "BuildHelperItemsFromCollection: collection element at index " & i & " is Nothing"
            Err.Raise 1000
        End If

        Set dictItem = CreateObject("Scripting.Dictionary")
        dictItem("IDExpedienteSuministrador") = CLng(item.getPropiedad("IDExpedienteSuministrador", p_Error))
        If p_Error <> "" Then Err.Raise 1000

        vIdPadre = item.getPropiedad("IdPadre", p_Error)
        If p_Error <> "" Then Err.Raise 1000
        If IsNull(vIdPadre) Or vIdPadre = "" Or vIdPadre = 0 Then
            dictItem("IdPadre") = 0
        Else
            dictItem("IdPadre") = CLng(vIdPadre)
        End If

        dictItem("Text") = CStr(Nz(item.getPropiedad("Descripcon", p_Error), ""))
        If p_Error <> "" Then Err.Raise 1000
        dictItem("Tag") = "RELID=" & CStr(dictItem("IDExpedienteSuministrador")) & ";IDS=" & CStr(item.getPropiedad("IDSuministrador", p_Error))
        If p_Error <> "" Then Err.Raise 1000

        ' EsEmpresaPropia: derive from m_Servicio when available (cheap lookup),
        ' else false. The DAO doesn't carry this flag.
        On Error Resume Next
        dictItem("EsEmpresaPropia") = m_Servicio.EsEmpresaPropia(CStr(item.getPropiedad("IDSuministrador", p_Error)))
        If Err.Number <> 0 Or IsEmpty(dictItem("EsEmpresaPropia")) Then
            dictItem("EsEmpresaPropia") = False
        End If
        On Error GoTo errores

        dict.Add "k" & CStr(i), dictItem
    Next i

    Set BuildHelperItemsFromCollection = dict
    Exit Function

errores:
    If Err.Number <> 1000 Then p_Error = "BuildHelperItemsFromCollection: " & Err.Description
    Set BuildHelperItemsFromCollection = New Scripting.Dictionary
End Function

' Paints the helper's serialized tree into Me.tvCadena. Best-effort: when the
' form is not OpenForm'd, all Me.* references raise and the OLE painting is
' skipped silently. This lets tests exercise the data wire-up without UI.
Private Sub PintarNodosDesdeHelper()
    Dim p_Error As String
    Dim json As String
    Dim parsed As Object
    Dim nodes As Object
    Dim nodRaiz As Object
    Dim sNombreOrgano As String
    Dim i As Long
    Dim vNode As Variant
    Dim sKey As String, sParentKey As String, sText As String, sTag As String
    Dim bEsRaiz As Boolean

    On Error Resume Next

    json = m_Helper.Serializar(p_Error)
    If p_Error <> "" Then Exit Sub

    Set parsed = JsonConverter.ParseJson(json)
    If parsed Is Nothing Then Exit Sub
    If Not parsed.Exists("payload") Then Exit Sub
    If Not parsed("payload").Exists("nodes") Then Exit Sub
    Set nodes = parsed("payload")("nodes")
    If nodes Is Nothing Then Exit Sub

    Me.tvCadena.Nodes.Clear

    ' Root node (always first per helper serialization contract)
    If nodes.Count >= 1 Then
        sNombreOrgano = CStr(Nz(parsed("payload")("OrganoNombre"), ""))
        Set nodRaiz = Me.tvCadena.Nodes.Add(, , K_ROOT_ORGANO, _
            "Órgano de Contratación (" & sNombreOrgano & ")", IMG_KEY_CONTRATISTA)
        nodRaiz.Bold = True
        nodRaiz.ForeColor = RGB(0, 51, 102)
    End If

    ' Children
    For i = 1 To nodes.Count
        Set vNode = nodes(i)
        If SafeStr(vNode("key")) = K_ROOT_ORGANO Then GoTo NextNode

        sKey = "K" & SafeStr(vNode("key"))
        If IsNull(vNode("parentKey")) Then
            sParentKey = K_ROOT_ORGANO
        Else
            sParentKey = "K" & SafeStr(vNode("parentKey"))
        End If

        Set nodRaiz = Me.tvCadena.Nodes.Add(sParentKey, 4, sKey, SafeStr(vNode("text")))
        nodRaiz.Tag = SafeStr(vNode("tag"))
        If CBool(vNode("esEmpresaPropia")) Then
            nodRaiz.Image = IMG_KEY_EMPRESA_GRUPO
            nodRaiz.Bold = True
        Else
            nodRaiz.Image = IMG_KEY_EMPRESA
            nodRaiz.Bold = False
        End If
        nodRaiz.Expanded = True
NextNode:
    Next i

    Me.tvCadena.Nodes(K_ROOT_ORGANO).Expanded = True
End Sub

' Small helper to coerce Scripting.Dictionary values to safe strings
Private Function SafeStr(ByVal p_Value As Variant) As String
    If IsNull(p_Value) Then
        SafeStr = ""
    ElseIf IsEmpty(p_Value) Then
        SafeStr = ""
    Else
        SafeStr = CStr(p_Value)
    End If
End Function

Private Function ObtenerTextoSocios(colTodos As Collection, sNombreActual As String) As String
    ObtenerTextoSocios = modExpedienteSuministradoresHelper.ExpedienteSuministradores_ObtenerTextoSocios(colTodos, sNombreActual, m_Error)
End Function


' ==============================================================================
' 4. DRAG & DROP (LOGICA CORE)
' ==============================================================================

Private Sub lvFuente_OLEStartDrag(Data As Object, AllowedEffects As Long)
    If Not Me.AllowEdits Then AllowedEffects = 0: Exit Sub
    If Me.lvFuente.SelectedItem Is Nothing Then Exit Sub
    
    AllowedEffects = 1 ' Copy
    Data.setData Me.lvFuente.SelectedItem.key, 1
End Sub

Private Sub tvCadena_OLEStartDrag(Data As Object, AllowedEffects As Long)
    If Not Me.AllowEdits Then AllowedEffects = 0: Exit Sub
    If Me.tvCadena.SelectedItem Is Nothing Then Exit Sub
    
    ' No mover raices
    If Left(Me.tvCadena.SelectedItem.key, 4) = "ROOT" Then AllowedEffects = 0: Exit Sub
    
    AllowedEffects = 2 ' Move
    Data.setData Me.tvCadena.SelectedItem.key, 1
End Sub

Private Sub tvCadena_OLEDragOver(Data As Object, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, state As Integer)
    Dim targetNode As Object
    Set targetNode = Me.tvCadena.HitTest(x, y)
    
    If Not Me.AllowEdits Then Effect = 0: Exit Sub
    
    If targetNode Is Nothing Then
        ' Permitir soltar en vacio solo si es Alta Externa (Copy) para crear raices
        If Data.GetFormat(1) Then
             If Left(Data.GetData(1), 4) = PREFIX_SRC Then Effect = 1 Else Effect = 0
        End If
        Set Me.tvCadena.DropHighlight = Nothing
    Else
        Set Me.tvCadena.DropHighlight = targetNode
        Dim sData As String
        On Error Resume Next
        sData = Data.GetData(1)
        
        If Left(sData, 4) = PREFIX_SRC Then
            Effect = 1 ' Copy
        Else
            If targetNode.key = sData Then Effect = 0 Else Effect = 2 ' Move
        End If
    End If
End Sub

Private Sub tvCadena_OLEDragDrop(Data As Object, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sourceNode As Object
    Dim targetNode As Object
    Dim sData As String
    Dim sTipoNodoDestino As String
    Dim vRelIDPadreDestino As Variant
    Dim sKeyDestino As String
    Dim bEsAltaExterna As Boolean
    Dim sRelIDMover As String
    
    On Error GoTo errores
    m_Error = ""
    
    ' Limpiar el resaltado visual
    Set Me.tvCadena.DropHighlight = Nothing
    
    ' 1. OBTENER DATOS Y CONTEXTO
    If Data.GetFormat(1) Then sData = Data.GetData(1)
    Set targetNode = Me.tvCadena.HitTest(x, y)
    
    If Left(sData, 4) = PREFIX_SRC Then
        bEsAltaExterna = True
    Else
        bEsAltaExterna = False
        Set sourceNode = Me.tvCadena.SelectedItem
        If sourceNode Is Nothing Then Exit Sub
    End If
    
    ' ---------------------------------------------------------
    ' LOGICA DE DESTINO
    ' ---------------------------------------------------------
    
    ' CASO A: ALTA EXTERNA
    If bEsAltaExterna Then
        Dim sIDSuministrador As String
        sIDSuministrador = Mid(sData, 5)
        
        ' Si soltamos en vacio o en ROOT_ORGANO -> Nivel 1
        Dim esRaiz As Boolean
        esRaiz = False
        If Not targetNode Is Nothing Then
            If targetNode.key = K_ROOT_ORGANO Then esRaiz = True
        Else
            esRaiz = True ' Vacio se asume Raiz
        End If
        
        If esRaiz Then
            sTipoNodoDestino = "ROOT_CONTR"
            vRelIDPadreDestino = Null
            sKeyDestino = K_ROOT_ORGANO
        Else
            sKeyDestino = targetNode.key
            If Left(sKeyDestino, 1) = "K" Then
                vRelIDPadreDestino = ExtraerTag(targetNode.Tag, "RELID")
                sTipoNodoDestino = "CHILD"
            End If
        End If
        
        EjecutarAlta sIDSuministrador, sTipoNodoDestino, vRelIDPadreDestino, sKeyDestino
        
    ' CASO B: MOVER INTERNO
    Else
        If targetNode Is Nothing Then Exit Sub
        If targetNode.key = sourceNode.key Then Exit Sub
        
        sRelIDMover = ExtraerTag(sourceNode.Tag, "RELID")
        sKeyDestino = targetNode.key
        
        If sKeyDestino = K_ROOT_ORGANO Then
            sTipoNodoDestino = "ROOT_CONTR"
            vRelIDPadreDestino = Null
        Else
            vRelIDPadreDestino = ExtraerTag(targetNode.Tag, "RELID")
            sTipoNodoDestino = "CHILD"
            If Not targetNode.Parent Is Nothing Then
                If targetNode.Parent.key = sourceNode.key Then Exit Sub
            End If
        End If
        
        EjecutarMovimiento sRelIDMover, sTipoNodoDestino, vRelIDPadreDestino, sKeyDestino
    End If
    
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Error DragDrop Nuevo: " & Err.Description, vbCritical
    RefrescarAArbol
End Sub

' ==============================================================================
' 6. BOTONES MANUALES Y MENÚ CONTEXTUAL
' ==============================================================================

Private Sub Alta()
    Dim sTipoNodo As String, vParentRelID As Variant, sNodeKey As String
    Dim iRespuesta As Integer
    Dim sIDSuministradorElegido As String
    
    If Me.tvCadena.SelectedItem Is Nothing Then
        ' Nada seleccionado -> Alta en Raiz (Organo)
        sTipoNodo = "ROOT_CONTR": vParentRelID = Null: sNodeKey = K_ROOT_ORGANO
    Else
        sNodeKey = Me.tvCadena.SelectedItem.key
        If sNodeKey = K_ROOT_ORGANO Then
            ' Seleccionado Organo -> Alta en Raiz (Organo)
            sTipoNodo = "ROOT_CONTR": vParentRelID = Null
        Else
            vParentRelID = ExtraerTag(Me.tvCadena.SelectedItem.Tag, "RELID")
            sTipoNodo = "CHILD"
        End If
    End If
    
    If Not IsNull(TempVars(TV_SELECTED_VAR)) Then TempVars.Remove TV_SELECTED_VAR
    DoCmd.OpenForm "FormSuministradoresGestion", WindowMode:=acDialog
    sIDSuministradorElegido = Nz(TempVars(TV_SELECTED_VAR), "")
    TempVars.Remove TV_SELECTED_VAR
    If sIDSuministradorElegido = "" Then Exit Sub
    
    EjecutarAlta sIDSuministradorElegido, sTipoNodo, vParentRelID, sNodeKey
End Sub

Private Sub cmdEliminar_Click()
    If Not Me.AllowEdits Then Exit Sub
    If Me.tvCadena.SelectedItem Is Nothing Then Exit Sub
    If Left(Me.tvCadena.SelectedItem.key, 4) = "ROOT" Then
        MsgBox "No se puede eliminar la raíz.", vbExclamation, "Error"
        Exit Sub
    End If
    
    If MsgBox("¿Eliminar seleccionado?", vbYesNo + vbQuestion, "Eliminar Nodo") = vbYes Then
        Dim sRelID As String
        sRelID = ExtraerTag(Me.tvCadena.SelectedItem.Tag, "RELID")
        EjecutarEliminacion sRelID
    End If
End Sub

Private Sub cmdVerDetalle_Click()
    Dim sIDSuministrador As String
    If Me.tvCadena.SelectedItem Is Nothing Then Exit Sub
    If Left(Me.tvCadena.SelectedItem.key, 4) = "ROOT" Then Exit Sub
    
    sIDSuministrador = ExtraerTag(Me.tvCadena.SelectedItem.Tag, "IDS")
    If sIDSuministrador <> "" Then
        Set m_ObjSuministradorActivo = constructor.getSuministrador(p_IDSuministrador:=sIDSuministrador)
        DoCmd.OpenForm "FormSuministrador", WindowMode:=acDialog
    End If
End Sub


Private Sub tvCadena_NodeClick(ByVal Node As Object)
    GestionarBotones Node
End Sub

Private Sub GestionarBotones(ByVal Node As Object)
    Dim haySel As Boolean
    Dim nodeKey As String
    haySel = Not (Node Is Nothing)
    If haySel Then nodeKey = Node.key Else nodeKey = ""
    Dim payload As Object
    Set payload = ParsePayload(modExpedienteSuministradoresHelper.ExpedienteSuministradores_GestionarBotonesState(haySel, nodeKey, Me.AllowEdits, m_Error))
    If payload Is Nothing Then Exit Sub
    Me.cmdEliminar.Enabled = CBool(payload("eliminarEnabled"))
    Me.cmdVerDetalle.Enabled = CBool(payload("verDetalleEnabled"))
End Sub

Private Sub MostrarMenuContextual(TIpo As String, ObjetoSeleccionado As Object)
    ' Este sub delega en MenuContextual.bas via EjecutarAccionMenu
    ' Copia la versión con Dispatcher que te di anteriormente
    Dim cBar As Object, btn As Object, sBarName As String
    On Error Resume Next
    sBarName = "CtxMenuExp"
    Application.CommandBars(sBarName).Delete
    On Error GoTo errores
    Set cBar = Application.CommandBars.Add(sBarName, 5, False, True)
    
    If TIpo = "ARBOL" Then
        Dim esRaiz As Boolean: esRaiz = (Left(ObjetoSeleccionado.key, 4) = "ROOT")
        If Not esRaiz Then
             Set btn = cBar.Controls.Add(1): btn.Caption = "Ver Detalle": btn.FaceId = 995: btn.OnAction = "=EjecutarAccionMenu('Public_VerDetalleArbol')"
        End If
        If Me.AllowEdits Then
             Set btn = cBar.Controls.Add(1): btn.Caption = "Añadir Aquí...": btn.FaceId = 21: btn.OnAction = "=EjecutarAccionMenu('Public_Alta')"
             If Not esRaiz Then
                 Set btn = cBar.Controls.Add(1): btn.Caption = "Eliminar": btn.FaceId = 478: btn.OnAction = "=EjecutarAccionMenu('Public_EliminarDelArbol')"
             End If
        End If
    ElseIf TIpo = "LISTA" Then
        Set btn = cBar.Controls.Add(1): btn.Caption = "Ver Ficha": btn.FaceId = 995: btn.OnAction = "=EjecutarAccionMenu('Public_VerDetalleLista')"
        If Me.AllowEdits Then
            Set btn = cBar.Controls.Add(1): btn.Caption = "Añadir al Expediente...": btn.FaceId = 265: btn.OnAction = "=EjecutarAccionMenu('Public_AnadirDesdeLista')"
        End If
    End If
    If cBar.Controls.Count > 0 Then cBar.ShowPopup
    cBar.Delete
    Exit Sub
errores:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

' ==============================================================================
' 7. FUNCIONES PÚBLICAS (DISPATCHER)
' ==============================================================================

Public Function Public_Alta()
    Alta
End Function
Public Function Public_EliminarDelArbol()
    cmdEliminar_Click
End Function
Public Function Public_VerDetalleArbol()
    cmdVerDetalle_Click
End Function
Public Function Public_VerDetalleLista()
    ' Ver detalle desde lista
    Dim sKey As String, sID As String
    If Me.lvFuente.SelectedItem Is Nothing Then Exit Function
    sKey = Me.lvFuente.SelectedItem.key
    If Left(sKey, 4) = PREFIX_SRC Then
        sID = Mid(sKey, 5)
        Set m_ObjSuministradorActivo = constructor.getSuministrador(p_IDSuministrador:=sID)
        DoCmd.OpenForm "FormSuministrador", WindowMode:=acDialog
    End If
End Function
Public Function Public_AnadirComoAdjudicatario()
    If Me.lvFuente.SelectedItem Is Nothing Then Exit Function
    EjecutarAlta Mid(Me.lvFuente.SelectedItem.key, 5), "ROOT_CONTR", Null, K_ROOT_ORGANO
End Function
Public Function Public_AnadirDesdeLista()
    Dim sID As String
    If Me.lvFuente.SelectedItem Is Nothing Then Exit Function
    
    sID = Mid(Me.lvFuente.SelectedItem.key, 5)
    EjecutarAlta sID, "ROOT_CONTR", Null, K_ROOT_ORGANO
End Function

' ==============================================================================
' 8. HELPERS INTERNOS
' ==============================================================================

Private Function ObtenerNombreEmpresa(p_ID As String) As String
    Dim oSum As Suministrador
    Set oSum = constructor.getSuministrador(p_IDSuministrador:=p_ID)
    If Not oSum Is Nothing Then
        If oSum.Nemotecnico <> "" Then ObtenerNombreEmpresa = oSum.Nemotecnico Else ObtenerNombreEmpresa = oSum.Nombre
    Else
        ObtenerNombreEmpresa = "ID:" & p_ID
    End If
End Function

Private Function ExtraerTag(ByVal tagText As String, ByVal keyName As String) As String
    ExtraerTag = modExpedienteSuministradoresHelper.ExpedienteSuministradores_ExtraerTag(tagText, keyName, m_Error)
End Function
Private Sub cmdNuevoSuministrador_Click()
    On Error GoTo errores
    
    ' 1. Reseteamos la variable global para indicar al formulario destino que es un ALTA NUEVA
    Set m_ObjSuministradorActivo = Nothing
    
    ' 2. Abrimos el formulario en modo normal (NO DIÁLOGO) para poder capturar eventos
    ' Si lo abres en modo acDialog, el código se detiene y no podemos setear la variable m_FormAltaSuministrador
    If FormularioAbierto("FormSuministrador") Then
        DoCmd.Close acForm, "FormSuministrador", acSaveNo
    End If
    
    DoCmd.OpenForm "FormSuministrador"
    
    ' 3. Vinculamos la instancia abierta a nuestra variable con eventos
    Set m_FormAltaSuministrador = Forms("FormSuministrador")
    
    Exit Sub
errores:
    MsgBox "Error al abrir alta de suministrador: " & Err.Description, vbCritical, "Error"
End Sub
Private Sub m_FormAltaSuministrador_Alta(p_ID As String)
    ' Este evento salta cuando el usuario guarda correctamente en FormSuministrador
    Dim oSum As Suministrador
    Dim sNombreFiltro As String
    
    On Error GoTo errores
    
    ' 1. Recuperamos los datos del suministrador recién creado usando el ID que nos devuelve el evento
    Set oSum = constructor.getSuministrador(p_IDSuministrador:=p_ID)
    
    If Not oSum Is Nothing Then
        ' Preferimos el Nombre para el filtro visual
        sNombreFiltro = oSum.Nombre
        
        ' 2. Actualizamos los controles de filtro según tus requisitos
        Me.chkNuestras.value = False      ' Desmarcar "Nuestras empresas"
        Me.txtFiltro.value = sNombreFiltro ' Poner el nombre en el cuadro de texto
        Me.txtFiltro.SetFocus             ' Foco al filtro (opcional, por usabilidad)
        
        ' 3. Refrescamos la lista de la derecha explícitamente
        ' (Asignar valor por código a un textbox no dispara el evento _Change, por eso llamamos a la funcion manual)
        CargarListaFuente sNombreFiltro, False
    End If
    
    ' 4. Limpiamos la referencia
    Set m_FormAltaSuministrador = Nothing
    
    Exit Sub
errores:
    MsgBox "Error al procesar el alta del suministrador: " & Err.Description, vbExclamation, "Error"
End Sub
Private Sub ActualizarEtiquetaSituacion()
    Dim bSomosPrincipal As Boolean
    
    On Error Resume Next
    

    ' --- MODO BD ---
    ' Delegamos en el servicio/repositorio que ya tiene esta consulta optimizada
    bSomosPrincipal = m_Servicio.YaSomosAdjudicatarios(m_IDExpediente)
   
    
   
    ' Sincronizar el objeto expediente en memoria para que otras partes de la app lo sepan
    If Not m_ObjExpedienteDTOActivo.Expediente Is Nothing Then
        m_ObjExpedienteDTOActivo.Expediente.ContratistaPrincipal = IIf(bSomosPrincipal, "Sí", "No")
    End If
End Sub




' ==============================================================================
' 5. LOGICA UNIFICADA DE PERSISTENCIA (BD vs MEMORIA)
' ==============================================================================

Private Sub EjecutarAlta(sIDS As String, sTipo As String, vPadre As Variant, sKeyDest As String)
    Dim exito As Boolean
    VBA.DoEvents
    DoCmd.Hourglass True
    
    
        ' Base de Datos
    exito = m_Servicio.AgregarSuministrador(m_ObjExpedienteDTOActivo.Expediente, sIDS, sTipo, vPadre, m_Error)
    
    
    If exito Then
        RefrescarAArbol
        On Error Resume Next
        If sKeyDest <> "" Then Me.tvCadena.Nodes(sKeyDest).Expanded = True
    Else
        MsgBox "Error Alta: " & m_Error, vbCritical, "Error"
    End If
    DoCmd.Hourglass False
End Sub

Private Sub EjecutarMovimiento(sRelID As String, sTipo As String, vPadre As Variant, sKeyDest As String)
    Dim exito As Boolean
    VBA.DoEvents
    DoCmd.Hourglass True
    
    
    ' Base de Datos
    exito = m_Servicio.MoverRelacion(m_ObjExpedienteDTOActivo.Expediente, sRelID, sTipo, vPadre, m_Error)
    
    
    If exito Then
        RefrescarAArbol
        On Error Resume Next
        Me.tvCadena.Nodes(sKeyDest).Expanded = True
    Else
        MsgBox "Error Mover: " & m_Error, vbCritical, "Error"
    End If
    DoCmd.Hourglass False
End Sub

Private Sub EjecutarEliminacion(sRelID As String)
    Dim exito As Boolean
    VBA.DoEvents
    DoCmd.Hourglass True
    
   
    If m_Servicio.TieneHijos(sRelID) Then
        m_Error = "Tiene dependencias."
        exito = False
    Else
        exito = m_Servicio.EliminarRelacion(m_ObjExpedienteDTOActivo.Expediente, sRelID, m_Error)
    End If
   
    
    If exito Then
        ' Eliminación visual rápida
        On Error Resume Next
        Me.tvCadena.Nodes.Remove Me.tvCadena.SelectedItem.key
        GestionarBotones Nothing
    Else
        MsgBox "Error Eliminar: " & m_Error, vbCritical, "Error"
    End If
    DoCmd.Hourglass False
End Sub

' --- FUNCIONES AUXILIARES PARA MEMORIA (DTO) ---

Private Function AgregarSuministradorMemoria(sIDS As String, sTipo As String, vPadre As Variant) As Boolean
    Dim newItem As ExpedienteSuministrador
    Dim sTempID As String
    
    On Error GoTo errMem
    ' Generar ID temporal único
    sTempID = "TMP_" & Format(Now, "hhmmss") & "_" & Int(Rnd * 10000)
    
    Set newItem = New ExpedienteSuministrador
    With newItem
        .IDExpedienteSuministrador = sTempID
        .IDSuministrador = sIDS
        .IdPadre = vPadre
        .Tag = sTempID ' Marca
        Select Case sTipo
            Case "ROOT_CONTR": .ContratistaPrincipal = "Sí": .SubContratista = "No": .Descripcon = "Contratista Ppal"
            Case "CHILD": .ContratistaPrincipal = "No": .SubContratista = "Sí": .Descripcon = "Subcontrata"
            Case Else: Err.Raise 1000
        End Select
    End With
    
    m_ObjExpedienteDTOActivo.ColArbolSuministradores.Add sTempID, newItem
    AgregarSuministradorMemoria = True
    Exit Function
errMem:
    m_Error = Err.Description
End Function

Private Function MoverRelacionMemoria(sRelID As String, sTipo As String, vPadre As Variant) As Boolean
    Dim item As ExpedienteSuministrador
    On Error GoTo errMem
    
    If m_ObjExpedienteDTOActivo.ColArbolSuministradores.Exists(sRelID) Then
        Set item = m_ObjExpedienteDTOActivo.ColArbolSuministradores(sRelID)
        item.IdPadre = vPadre
        Select Case sTipo
            Case "ROOT_CONTR": item.ContratistaPrincipal = "Sí": item.SubContratista = "No"
            Case "CHILD": item.ContratistaPrincipal = "No": item.SubContratista = "Sí"
            Case Else: Err.Raise 1000
        End Select
        MoverRelacionMemoria = True
    Else
        m_Error = "Elemento no encontrado en memoria."
        MoverRelacionMemoria = False
    End If
    Exit Function
errMem:
    m_Error = Err.Description
    MoverRelacionMemoria = False
End Function

Private Function EliminarRelacionMemoria(sRelID As String) As Boolean
    ' Verificar hijos en memoria
    Dim k As Variant
    Dim TieneHijos As Boolean
    TieneHijos = False
    For Each k In m_ObjExpedienteDTOActivo.ColArbolSuministradores.Keys
        If m_ObjExpedienteDTOActivo.ColArbolSuministradores(k).IdPadre = sRelID Then
            TieneHijos = True
            Exit For
        End If
    Next
    
    If TieneHijos Then
        EliminarRelacionMemoria = False
    Else
        If m_ObjExpedienteDTOActivo.ColArbolSuministradores.Exists(sRelID) Then
            m_ObjExpedienteDTOActivo.ColArbolSuministradores.Remove sRelID
            EliminarRelacionMemoria = True
        End If
    End If
End Function


