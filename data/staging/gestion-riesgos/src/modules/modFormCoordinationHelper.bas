Attribute VB_Name = "modFormCoordinationHelper"
Option Compare Database
Option Explicit

' ----------------------------------------------------------------------------
' Helper de coordinacion entre formularios.
'   Skill ref: access-vba-e2e-methodology Hard rule 3 + Hard rule 4.
'
'   Que resuelve: hasta ahora, los forms Materializado / Mitigacion /
'   Retirado y FormTecnicoTareaRiesgosAceptadosRetirados llamaban
'   directamente a Form_FormRiesgosGestion.m_NodoSeleccionado,
'   Forms("FormRiesgosGestion").Controls("comboVerDescripcion"),
'   Form_FormRiesgosGestionRiesgo.EstablecerDatos,
'   Form_FormRiesgo.EstablecerDatos y
'   Form_FormTecnicoTareas.m_RiesgoSeleccionado desde su propio .cls.
'   Eso son 11 violaciones de Hard rule 3 (un form .cls accediendo a
'   otro form .cls / controls).
'
'   Que hace este helper: centraliza esos accesos detras de funciones
'   Public con prefijo Coord_<Verbo>. Las formas externas llaman al
'   helper y este helper es el UNICO modulo autorizado a tocar
'   Forms("X") o los Public fields de los forms origen.
'
'   Que NO hace: logica de negocio, validacion, persistencia, DAO.
'   Solo enruta lecturas / llamadas que ya existian. Sin MsgBox
'   (Hard rule 7); errores via Optional ByRef p_Error.
'
'   Naming: Coord_<Verbo> por convencion (prefijo corto para no
'   chocar con Publics legacy).
'
'   Limites explicitos:
'     - Coord_GetSelectedNode / Coord_GetComboVerDescripcion leen
'       campos Public del form FormRiesgosGestion. Esto es una
'       excepcion deliberada al Hard rule 4 (Public business method
'       in a form), porque el estado del TreeView + combo es
'       UI-adapter (no logica de negocio) y moverlo a un servicio
'       obligaria a cambiar el modelo de eventos del TreeView. La
'       deuda se documenta en el cap-doc §7 confidence ledger como
'       Verified-static hasta que el slice de extraccion de
'       ArbolEstado lo elimine.
'     - Coord_RefreshRiesgoGestionRiesgo / Coord_RefreshRiesgo /
'       Coord_RefreshRiesgosGestion llaman al helper
'       modRiesgoEstablecerDatosHelper (proxy entries _Riesgo /
'       _RiesgoGestionRiesgo / _RiesgosGestion), que a su vez
'       delega al Public Sub Refrescar del form (UI adapter tiny,
'       Hard rule 4 OK). El caller queda limpio (Hard rule 3 OK).
'       Hard rule 4 sobre los Public Function EstablecerDatos de
'       los forms queda resuelto por el helper
'       modRiesgoEstablecerDatosHelper (slice HR4 / HR4b).
' ----------------------------------------------------------------------------

' ----------------------------------------------------------------------------
' Estado del nodo seleccionado en el arbol de FormRiesgosGestion.
'   Patron A (function-based) - no usamos clase con ciclo de vida
'   porque el TreeView ya mantiene el SelectedItem; el helper solo
'   lo expone.
' ----------------------------------------------------------------------------
Public Function Coord_GetSelectedNode(Optional ByRef p_Error As String) As Object
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        ' Estado normal (form no abierto): Nothing sin raise. El caller
        ' decide si aborta o usa el valor por defecto de su propio campo.
        Exit Function
    End If

    Set Coord_GetSelectedNode = Forms("FormRiesgosGestion").m_NodoSeleccionado
    Exit Function

errores:
    p_Error = "Coord_GetSelectedNode: " & Err.Description
    ' Defensivo: devolver Nothing en error sin propagar raise al caller.
End Function

' ----------------------------------------------------------------------------
' Valor del combo de seleccion de descripcion en FormRiesgosGestion.
' ----------------------------------------------------------------------------
Public Function Coord_GetComboVerDescripcion(Optional ByRef p_Error As String) As String
    On Error GoTo errores
    p_Error = ""
    Coord_GetComboVerDescripcion = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        Exit Function
    End If

    Coord_GetComboVerDescripcion = Nz(Forms("FormRiesgosGestion").Controls("comboVerDescripcion").Value, "")
    Exit Function

errores:
    p_Error = "Coord_GetComboVerDescripcion: " & Err.Description
End Function

' ----------------------------------------------------------------------------
' Riesgo seleccionado en FormTecnicoTareas.
'   Patron A: lectura directa del Public field del form origen.
' ----------------------------------------------------------------------------
Public Function Coord_GetRiesgoSeleccionadoDeTecnicoTareas(Optional ByRef p_Error As String) As Object
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormTecnicoTareas") Then
        Exit Function
    End If

    Set Coord_GetRiesgoSeleccionadoDeTecnicoTareas = Forms("FormTecnicoTareas").m_RiesgoSeleccionado
    Exit Function

errores:
    p_Error = "Coord_GetRiesgoSeleccionadoDeTecnicoTareas: " & Err.Description
End Function

' ----------------------------------------------------------------------------
' Refresco de FormRiesgosGestionRiesgo (calidad).
'   HR4 fix (slice 2e08ce0+): proxy -> modRiesgoEstablecerDatosHelper.
'   Coord NO llama form.EstablecerDatos (Public business method). Llama al
'   helper, que delega al Public Sub Refrescar del form (UI adapter tiny,
'   Hard rule 4 OK).
' ----------------------------------------------------------------------------
Public Sub Coord_RefreshRiesgoGestionRiesgo(Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestionRiesgo") Then
        p_Error = "Coord_RefreshRiesgoGestionRiesgo: FormRiesgosGestionRiesgo no esta abierto (no-op)"
        Exit Sub
    End If

    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_RiesgoGestionRiesgo p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_RefreshRiesgoGestionRiesgo: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' Refresco de FormRiesgo (padre de gemelos).
'   HR4 fix: proxy -> modRiesgoEstablecerDatosHelper (no Public business
'   method del form).
' ----------------------------------------------------------------------------
Public Sub Coord_RefreshRiesgo(Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgo") Then
        p_Error = "Coord_RefreshRiesgo: FormRiesgo no esta abierto (no-op)"
        Exit Sub
    End If

    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Riesgo p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_RefreshRiesgo: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' Refresco de FormRiesgosGestion (parent LIST form).
'   Slice HR4b: misma mecánica que los dos anteriores — proxy ->
'   modRiesgoEstablecerDatosHelper (no Public business method del form).
'   Coord NO llama form.EstablecerDatos (Public business method). Llama al
'   helper, que delega al Public Sub Refrescar del form (UI adapter tiny,
'   Hard rule 4 OK).
' ----------------------------------------------------------------------------
Public Sub Coord_RefreshRiesgosGestion(Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        p_Error = "Coord_RefreshRiesgosGestion: FormRiesgosGestion no esta abierto (no-op)"
        Exit Sub
    End If

    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_RiesgosGestion p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_RefreshRiesgosGestion: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' CargarArbolPM - puente al Public Function CargarArbolPM de
'   Form_FormRiesgosGestion. El caller (Form_FormRiesgosGestionRiesgo en
'   su handler m_FormPlan_PlanNuevo) ya no llama al form
'   directamente; pasa por el helper para evitar violar Hard rule 3.
'
'   Hard rule 2: inyecto solo lo que el destino necesita. P_NodoRiesgo,
'   p_Riesgo y p_PM llegan del caller, p_borrarNodoSeleccionado y
'   p_Refrescando son flags del algoritmo de refresco (no requieren db).
'   Sin db: el arbol del form no toca DAO, solo manipula m_Arbol.Nodes.
'
'   Hard rule 7: sin MsgBox. Errores via Optional ByRef p_Error.
' ----------------------------------------------------------------------------
Public Sub Coord_CargarArbolPM( _
                                ByRef P_NodoRiesgo As Object, _
                                ByRef p_Riesgo As Object, _
                                ByRef p_PM As Object, _
                                Optional ByRef p_borrarNodoSeleccionado As EnumSiNo = EnumSiNo.Sí, _
                                Optional ByRef p_Refrescando As EnumSiNo = EnumSiNo.No, _
                                Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        p_Error = "Coord_CargarArbolPM: FormRiesgosGestion no esta abierto (no-op)"
        Exit Sub
    End If

    Form_FormRiesgosGestion.CargarArbolPM _
        P_NodoRiesgo:=P_NodoRiesgo, _
        p_Riesgo:=p_Riesgo, _
        p_PM:=p_PM, _
        p_borrarNodoSeleccionado:=p_borrarNodoSeleccionado, _
        p_Refrescando:=p_Refrescando, _
        p_Error:=p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_CargarArbolPM: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' CargarArbolPC - puente al Public Function CargarArbolPC de
'   Form_FormRiesgosGestion. Mismo patron que Coord_CargarArbolPM,
'   hermano para planes de contingencia. Ver notas en Coord_CargarArbolPM
'   sobre Hard rule 2, 3, 4, 7.
' ----------------------------------------------------------------------------
Public Sub Coord_CargarArbolPC( _
                                ByRef P_NodoRiesgo As Object, _
                                ByRef p_Riesgo As Object, _
                                ByRef p_PC As Object, _
                                Optional ByRef p_borrarNodoSeleccionado As EnumSiNo = EnumSiNo.Sí, _
                                Optional ByRef p_Refrescando As EnumSiNo = EnumSiNo.No, _
                                Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        p_Error = "Coord_CargarArbolPC: FormRiesgosGestion no esta abierto (no-op)"
        Exit Sub
    End If

    Form_FormRiesgosGestion.CargarArbolPC _
        P_NodoRiesgo:=P_NodoRiesgo, _
        p_Riesgo:=p_Riesgo, _
        p_PC:=p_PC, _
        p_borrarNodoSeleccionado:=p_borrarNodoSeleccionado, _
        p_Refrescando:=p_Refrescando, _
        p_Error:=p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_CargarArbolPC: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' HR3c slice: 11 nuevos Public entries que cierran las violaciones
'   cross-form restantes en Form_FormPlanAcciones / Form_FormRiesgosGestion
'   PlanAcciones / Form_FormRiesgosGestionPlanPrincipal / Form_FormRiesgos
'   GestionRiesgo. Patron consistente con entries previos (slice HR3b):
'     - Reads: no-op defensivo cuando el form no esta abierto (devuelve
'       default: Empty/False/Nothing, p_Error queda vacio).
'     - Sub proxies: no-op + p_Error legible cuando el form no esta
'       abierto, llamado al form si esta abierto.
'   Naming: Coord_<Verbo> por convencion.
' ----------------------------------------------------------------------------

' ----------------------------------------------------------------------------
' RefrescarArbolRiesgosScope - puente al Public Function de
'   Form_FormRiesgosGestion. Reemplaza Form_FormRiesgosGestion.
'   RefrescarArbolRiesgosScope p_Scope:="..." en Form_FormPlanAcciones
'   (L80, L121), Form_FormRiesgosGestionPlanAcciones (L82, L167) y
'   Form_FormRiesgosGestionPlanPrincipal (L172). Patron A: thin proxy.
'
'   Hard rule 2: p_Scope es el unico parametro que el destino necesita
'   (string literal: "plan" / "action" / "action-delete"). Sin db.
'   Hard rule 7: sin MsgBox; errores via Optional ByRef p_Error.
' ----------------------------------------------------------------------------
Public Sub Coord_RefrescarArbolRiesgosScope( _
                                            ByVal p_Scope As String, _
                                            Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        p_Error = "Coord_RefrescarArbolRiesgosScope: FormRiesgosGestion no esta abierto (no-op)"
        Exit Sub
    End If

    Form_FormRiesgosGestion.RefrescarArbolRiesgosScope _
        p_Scope:=p_Scope, _
        p_Error:=p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_RefrescarArbolRiesgosScope: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' m_EsMitigacion (Public EnumSiNo en FormRiesgosGestion) - lectura
'   defensiva. Reemplaza Form_FormRiesgosGestion.m_EsMitigacion en
'   Form_FormRiesgosGestionPlanAcciones (L210, L214) y Form_FormRiesgos
'   GestionPlanPrincipal (L266, L270). Hard rule 7 sin MsgBox.
' ----------------------------------------------------------------------------
Public Function Coord_GetEsMitigacion(Optional ByRef p_Error As String) As String
    On Error GoTo errores
    p_Error = ""
    Coord_GetEsMitigacion = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        Exit Function
    End If

    Coord_GetEsMitigacion = CStr(Forms("FormRiesgosGestion").m_EsMitigacion)
    Exit Function

errores:
    p_Error = "Coord_GetEsMitigacion: " & Err.Description
End Function

' ----------------------------------------------------------------------------
' m_ObjSeleccionado (Public Object en FormRiesgosGestion) - lectura
'   defensiva. Reemplaza Form_FormRiesgosGestion.m_ObjSeleccionado en
'   Form_FormRiesgosGestionPlanAcciones (L320-325) para TypeOf checks.
'   Devuelve Nothing si el form no esta abierto. Hard rule 7 sin MsgBox.
' ----------------------------------------------------------------------------
Public Function Coord_GetObjSeleccionado(Optional ByRef p_Error As String) As Object
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        Exit Function
    End If

    Set Coord_GetObjSeleccionado = Forms("FormRiesgosGestion").m_ObjSeleccionado
    Exit Function

errores:
    p_Error = "Coord_GetObjSeleccionado: " & Err.Description
End Function

' ----------------------------------------------------------------------------
' blnPermitidoEditar (Public Boolean en FormRiesgosGestion) - lectura
'   defensiva. Reemplaza Form_FormRiesgosGestion.blnPermitidoEditar en
'   Form_FormRiesgosGestionPlanPrincipal (L80, L142, L299). Hard rule 7.
' ----------------------------------------------------------------------------
Public Function Coord_GetPermitidoEditar(Optional ByRef p_Error As String) As Boolean
    On Error GoTo errores
    p_Error = ""
    Coord_GetPermitidoEditar = False

    If Not FormularioAbierto("FormRiesgosGestion") Then
        Exit Function
    End If

    Coord_GetPermitidoEditar = Forms("FormRiesgosGestion").blnPermitidoEditar
    Exit Function

errores:
    p_Error = "Coord_GetPermitidoEditar: " & Err.Description
End Function

' ----------------------------------------------------------------------------
' Arbol_NodeClick - puente al Public Sub de Form_FormRiesgosGestion.
'   Reemplaza Form_FormRiesgosGestion.Arbol_NodeClick en Form_Form
'   RiesgosGestionPlanPrincipal (L113). Patron A: thin proxy.
' ----------------------------------------------------------------------------
Public Sub Coord_ArbolNodeClick( _
                                ByVal p_Node As Object, _
                                Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        p_Error = "Coord_ArbolNodeClick: FormRiesgosGestion no esta abierto (no-op)"
        Exit Sub
    End If

    Form_FormRiesgosGestion.Arbol_NodeClick p_Node
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_ArbolNodeClick: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' ArbolEliminarNodoConRefresco - patron post-delete del FormRiesgos
'   GestionPlanPrincipal.ComandoEliminar_Click: selecciona el padre
'   del nodo, dispara Arbol_NodeClick sobre el padre, y elimina el
'   nodo del arbol. Reemplaza 3 violaciones (L112, L113, L117) por
'   una sola llamada. El caller pasa el Key del nodo a eliminar.
' ----------------------------------------------------------------------------
Public Sub Coord_ArbolEliminarNodoConRefresco( _
                                                ByVal p_Key As String, _
                                                Optional ByRef p_Error As String)
    Dim m_NodoAEliminar As Object
    Dim m_NodoPadre As Object
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        p_Error = "Coord_ArbolEliminarNodoConRefresco: FormRiesgosGestion no esta abierto (no-op)"
        Exit Sub
    End If

    Set m_NodoAEliminar = Form_FormRiesgosGestion.m_Arbol.Nodes(p_Key)
    Set m_NodoPadre = m_NodoAEliminar.Parent

    ' Seleccionar el padre y disparar NodeClick (mismo orden que el
    ' call site original L112, L113 para preservar UX).
    m_NodoPadre.Selected = True
    Form_FormRiesgosGestion.Arbol_NodeClick m_NodoPadre
    m_NodoPadre.Selected = True
    Form_FormRiesgosGestion.m_Arbol.Nodes.Remove p_Key
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_ArbolEliminarNodoConRefresco: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' m_Arbol.SelectedItem (TreeView Public field en FormRiesgosGestion) -
'   lectura defensiva. Reemplaza Form_FormRiesgosGestion.m_Arbol
'   .SelectedItem en Form_FormRiesgosGestionPlanPrincipal (L84) y
'   Form_FormRiesgosGestionRiesgo.ComandoEliminar_Click (L123). Hard rule 7.
' ----------------------------------------------------------------------------
Public Function Coord_ArbolSelectedItem(Optional ByRef p_Error As String) As Object
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        Exit Function
    End If

    Set Coord_ArbolSelectedItem = Form_FormRiesgosGestion.m_Arbol.SelectedItem
    Exit Function

errores:
    p_Error = "Coord_ArbolSelectedItem: " & Err.Description
End Function

' ----------------------------------------------------------------------------
' CargarArbol - puente al Public Function de Form_FormRiesgosGestion.
'   Reemplaza Form_FormRiesgosGestion.CargarArbol en Form_FormRiesgos
'   GestionRiesgo.ComandoEliminar_Click (L129). Patron A: thin proxy.
' ----------------------------------------------------------------------------
Public Sub Coord_CargarArbol(Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        p_Error = "Coord_CargarArbol: FormRiesgosGestion no esta abierto (no-op)"
        Exit Sub
    End If

    Form_FormRiesgosGestion.CargarArbol _
        p_Refrescando:=EnumSiNo.No, _
        p_Error:=p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_CargarArbol: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' SeleccionarNodo - puente al Public Function de Form_FormRiesgosGestion.
'   Reemplaza Form_FormRiesgosGestion.SeleccionarNodo en Form_Form
'   RiesgosGestionRiesgo.ComandoEliminar_Click (L130) y en
'   Form_FormRiesgosGestionRiesgo.m_FormPlan_PlanNuevo (L444, fuera de
'   scope - ya marcado en HR3b). Patron A: thin proxy.
' ----------------------------------------------------------------------------
Public Sub Coord_SeleccionarNodo( _
                                ByRef p_Objeto As Object, _
                                Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        p_Error = "Coord_SeleccionarNodo: FormRiesgosGestion no esta abierto (no-op)"
        Exit Sub
    End If

    Form_FormRiesgosGestion.SeleccionarNodo p_Objeto, p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_SeleccionarNodo: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' CargarArbolAccion - puente al Public Function de Form_FormRiesgos
'   Gestion. Reemplaza Form_FormRiesgosGestion.CargarArbolAccion en
'   Form_FormRiesgosGestionPlanPrincipal.m_FormAccion_AccionNueva (L398).
'   Patron A: thin proxy.
' ----------------------------------------------------------------------------
Public Sub Coord_CargarArbolAccion( _
                                    ByRef p_NodoPlan As Object, _
                                    ByRef p_Accion As Object, _
                                    Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        p_Error = "Coord_CargarArbolAccion: FormRiesgosGestion no esta abierto (no-op)"
        Exit Sub
    End If

    Form_FormRiesgosGestion.CargarArbolAccion _
        P_NodoPlan:=p_NodoPlan, _
        p_Accion:=p_Accion, _
        p_Error:=p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_CargarArbolAccion: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' ColRiesgosAplicados RemoveIfExists - borra un riesgo del Dictionary
'   Public m_ColRiesgosAplicados de FormRiesgosGestion solo si existe
'   y solo si el Dictionary esta instanciado. Reemplaza 3 violaciones
'   en Form_FormRiesgosGestionRiesgo.ComandoEliminar_Click (L124, L125,
'   L126) por una sola llamada. Patron A: defensivo encapsulado.
'
'   Hard rule 2: solo inyecta p_Key (lo unico que necesita el destino).
' ----------------------------------------------------------------------------
Public Sub Coord_ColRiesgosAplicados_RemoveIfExists( _
                                                    ByVal p_Key As String, _
                                                    Optional ByRef p_Error As String)
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        p_Error = "Coord_ColRiesgosAplicados_RemoveIfExists: FormRiesgosGestion no esta abierto (no-op)"
        Exit Sub
    End If

    Set m_Col = Forms("FormRiesgosGestion").m_ColRiesgosAplicados
    If m_Col Is Nothing Then
        Exit Sub
    End If
    If m_Col.Exists(p_Key) Then
        m_Col.Remove p_Key
    End If
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_ColRiesgosAplicados_RemoveIfExists: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' HR3d slice (third HR3 closure) - 3 nuevas entradas que cierran las
'   violaciones cross-form en Form_FormRiesgosEstablecerPrioridades,
'   Form_FormGestionRiesgosDatosGenerales y Form_FormPublicacionCalidadPublicar.
'   Naming consistente con entries previos (Coord_<Verbo>): prefijo corto
'   para no chocar con Publics legacy.
' ----------------------------------------------------------------------------

' ----------------------------------------------------------------------------
' ExpedientesBusqueda Filtrar - thin proxy al Public Function Filtrar de
'   Form_FormExpedientesBusqueda. Reemplaza Form_FormExpedientesBusqueda.
'   Filtrar en Form_FormGestionRiesgosDatosGenerales
'   .ComandoIrABusquedaExpedientes_Click (L79) y .ComandoBuscarExpediente
'   _Click (L316). Patron A: thin proxy.
'
'   Hard rule 2: p_PalabraClave es el unico parametro que el destino
'   necesita (string). El destino es un Public Sub sin return, sin db.
'   Hard rule 7: sin MsgBox; errores via Optional ByRef p_Error.
' ----------------------------------------------------------------------------
Public Sub Coord_ExpedientesBusqueda_Filtrar( _
                                                Optional ByVal p_PalabraClave As String = "", _
                                                Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormExpedientesBusqueda") Then
        p_Error = "Coord_ExpedientesBusqueda_Filtrar: FormExpedientesBusqueda no esta abierto (no-op)"
        Exit Sub
    End If

    If Len(Trim$(p_PalabraClave)) > 0 Then
        Forms("FormExpedientesBusqueda").PalabraClave = p_PalabraClave
    End If
    Form_FormExpedientesBusqueda.Filtrar p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_ExpedientesBusqueda_Filtrar: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' ComandoActualizarContadorRiesgosGestion - thin proxy al Public Sub
'   ComandoActualizarContador_Click de Form_FormRiesgosGestion. Reemplaza
'   Form_FormRiesgosGestion.ComandoActualizarContador_Click en
'   Form_FormRiesgosEstablecerPrioridades.EstablecerPriorizaciones
'   (L575). Patron A: thin proxy.
'
'   Hard rule 2: el destino es un Public Sub sin params, sin return, sin
'   db. La entrada coord no inyecta nada.
'   Hard rule 7: sin MsgBox; errores via Optional ByRef p_Error.
' ----------------------------------------------------------------------------
Public Sub Coord_ComandoActualizarContadorRiesgosGestion(Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        p_Error = "Coord_ComandoActualizarContadorRiesgosGestion: FormRiesgosGestion no esta abierto (no-op)"
        Exit Sub
    End If

    Form_FormRiesgosGestion.ComandoActualizarContador_Click
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_ComandoActualizarContadorRiesgosGestion: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' EstablecerLblRechazadoEnDetalleRiesgos - encapsula la navegacion cross-
'   form Forms("FormRiesgosGestion").FormDetalle.SourceObject = "...
'   Edicion" y Set mForm = mForm.FormDetalle.Form + llamada a
'   EstablecerlblRechazado (que vive en Funciones Generales.bas). Reemplaza
'   la violacion cross-form en Form_FormPublicacionCalidadPublicar
'   .m_FormMotivos_Motivado (L407-415). Patron A: defensivo encapsulado.
'
'   Hard rule 2: el destino solo necesita saber "donde esta el subform
'   FormRiesgosGestionEdicion dentro del FormRiesgosGestion abierto". El
'   helper hace la navegacion; el caller no toca Forms("X") ni
'   mForm.FormDetalle.
'   Hard rule 7: sin MsgBox; errores via Optional ByRef p_Error. El
'   helper interno EstablecerlblRechazado tampoco usa MsgBox.
' ----------------------------------------------------------------------------
Public Sub Coord_EstablecerLblRechazadoEnDetalleRiesgos(Optional ByRef p_Error As String)
    Dim mForm As Form
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        p_Error = "Coord_EstablecerLblRechazadoEnDetalleRiesgos: FormRiesgosGestion no esta abierto (no-op)"
        Exit Sub
    End If

    Set mForm = Forms("FormRiesgosGestion")
    If mForm.FormDetalle.SourceObject <> "FormRiesgosGestionEdicion" Then
        ' Subform no es la edicion: no hay label que actualizar. No-op
        ' silencioso (mismo comportamiento que el call site original:
        ' solo llamaba a EstablecerlblRechazado cuando el SourceObject
        ' era FormRiesgosGestionEdicion).
        Exit Sub
    End If

    Set mForm = mForm.FormDetalle.Form
    EstablecerlblRechazado mForm, p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_EstablecerLblRechazadoEnDetalleRiesgos: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' SetComandoNoExisteRiesgoVisible - puente al control .Visible de
'   Form_FormRiesgosBibliotecaGestion. Reemplaza el cross-form
'   Form_FormRiesgosBibliotecaGestion.ComandoNoExisteRiesgo.Visible
'   = False en Form_Form0BDOpciones.ComandoBibliotecaRiesgos_Click
'   (L228). Patron A: thin proxy.
' ----------------------------------------------------------------------------
Public Sub Coord_SetComandoNoExisteRiesgoVisible( _
                                                ByVal p_Visible As Boolean, _
                                                Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosBibliotecaGestion") Then
        p_Error = "Coord_SetComandoNoExisteRiesgoVisible: FormRiesgosBibliotecaGestion no esta abierto (no-op)"
        Exit Sub
    End If

    Forms("FormRiesgosBibliotecaGestion").ComandoNoExisteRiesgo.Visible = p_Visible
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_SetComandoNoExisteRiesgoVisible: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' FiltrarRiesgosBibliotecaGestion - puente al Public Function
'   Filtrar de Form_FormRiesgosBibliotecaGestion. Reemplaza el
'   cross-form Form_FormRiesgosBibliotecaGestion.Filtrar en
'   Form_FormGestionRiesgosRiesgosOferta.ComandoSeleccionar
'   RiesgoDeBiblioteca_Click (L302). Patron A: thin proxy.
' ----------------------------------------------------------------------------
Public Sub Coord_FiltrarRiesgosBibliotecaGestion(Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosBibliotecaGestion") Then
        p_Error = "Coord_FiltrarRiesgosBibliotecaGestion: FormRiesgosBibliotecaGestion no esta abierto (no-op)"
        Exit Sub
    End If

    Form_FormRiesgosBibliotecaGestion.Filtrar p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_FiltrarRiesgosBibliotecaGestion: " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------------------
' RefrescarSubformDetalleRiesgo - dispatcher para refrescar el
'   subform de detalle del riesgo (Form_FormRiesgoX.EstablecerDatos)
'   segun el subform activo en Form_FormRiesgo. Reemplaza el
'   cross-form Select Case en Form_FormRiesgo.ComandoActualizar
'   _Click (L31-43) que llamaba Form_FormRiesgoDefinicion /
'   Form_FormRiesgoPlazoCosteCalidad / Form_FormRiesgoVulnerabilidad /
'   Form_FormRiesgoMitigacion / Form_FormRiesgoRetirado /
'   Form_FormRiesgoMaterializado.EstablecerDatos (6 forms).
'   Patron A (thin proxy) + dispatch table.
' ----------------------------------------------------------------------------
Public Sub Coord_RefrescarSubformDetalleRiesgo( _
                                                ByVal p_SubformName As String, _
                                                Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Len(Trim$(p_SubformName)) = 0 Then
        p_Error = "Coord_RefrescarSubformDetalleRiesgo: nombre de subform vacio"
        Exit Sub
    End If

    If Not FormularioAbierto("FormRiesgo") Then
        p_Error = "Coord_RefrescarSubformDetalleRiesgo: FormRiesgo no esta abierto (no-op)"
        Exit Sub
    End If

    ' Dispatch table: subform name -> form .cls -> form.EstablecerDatos
    Select Case p_SubformName
        Case "FormRiesgoDefinicion"
            Form_FormRiesgoDefinicion.EstablecerDatos
        Case "FormRiesgoPlazoCosteCalidad"
            Form_FormRiesgoPlazoCosteCalidad.EstablecerDatos
        Case "FormRiesgoVulnerabilidad"
            Form_FormRiesgoVulnerabilidad.EstablecerDatos
        Case "FormRiesgoMitigacion"
            Form_FormRiesgoMitigacion.EstablecerDatos
        Case "FormRiesgoRetirado"
            Form_FormRiesgoRetirado.EstablecerDatos
        Case "FormRiesgoMaterializado"
            Form_FormRiesgoMaterializado.EstablecerDatos
        Case Else
            p_Error = "Coord_RefrescarSubformDetalleRiesgo: subform no reconocido '" & p_SubformName & "'"
            Exit Sub
    End Select
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "Coord_RefrescarSubformDetalleRiesgo: " & Err.Description
    End If
End Sub
