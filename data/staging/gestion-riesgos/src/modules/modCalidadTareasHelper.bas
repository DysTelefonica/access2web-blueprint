Attribute VB_Name = "modCalidadTareasHelper"
' ============================================================
' modCalidadTareasHelper.bas
'
' Helper que extrae la logica de negocio de:
'   - Form_FormCalidadTareas.SeleccionarNodoRiesgo (line 429)
'   - Form_FormCalidadTareas.CargarArbol          (line 355)
'
' Slice HR4-sister (HR3d follow-up #9, #4 - 2026-07-01).
'
' NOTA: CargarArbol tiene 2 callers externos (Form_FormTecnicoTareas.cls:305
' y Form_FormRiesgosGestion.cls:1696 - 4 callers). El form-side Public
' Function CargarArbol se mantiene Public como thin adapter que delega
' a este helper (los callers externos llaman al form, el form delega
' al helper).
'
' Arbol_NodeClick es un event handler y se hace Private sin extraccion
' (per Hard rule 4 exception: "event wiring").
'
' Hard rule 1: cero business logic en event handlers o Public methods en forms.
' Hard rule 2: honest signature — solo inyecta lo que la funcion necesita.
' Hard rule 4: Public business methods en forms son anti-patron.
' Hard rule 5: nombres Public globalmente unicos con ModuleName_ prefix.
' Hard rule 7: NO MsgBox/InputBox en helpers. Errores via ByRef p_Error.
' ============================================================
Option Compare Database
Option Explicit

' =============================================================================
' Public Sub CalidadTareas_SeleccionarNodoRiesgo
'   Valida input + construye el key del nodo segun tipo + estado del riesgo.
'   El form-side hace el lookup en el TreeView (UI-specific).
'   Raises: NO. Errores via p_Error ByRef.
' =============================================================================
Public Sub CalidadTareas_SeleccionarNodoRiesgo( _
    ByVal p_ObjRiesgo As riesgo, _
    ByVal p_TipoRiesgoTarea As EnumTipoRiesgoTarea, _
    ByRef p_OutKey As String, _
    Optional ByRef p_Error As String)

    On Error GoTo errores
    p_Error = ""
    p_OutKey = ""

    ' Edge case: Nothing -> set p_Error (matching original behavior)
    If p_ObjRiesgo Is Nothing Then
        p_Error = "Se ha de indicar el Riesgo"
        Exit Sub
    End If

    Dim m_Estado As EnumRiesgoEstado
    m_Estado = p_ObjRiesgo.EstadoEnum

    If p_TipoRiesgoTarea = EnumTipoRiesgoTarea.Aceptados Then
        If m_Estado = EnumRiesgoEstado.AceptadoSinVisar Then
            p_OutKey = "RIESGOACEPTADOPORVISAR|" & p_ObjRiesgo.IDRiesgo
        End If
    ElseIf p_TipoRiesgoTarea = EnumTipoRiesgoTarea.Retirados Then
        If m_Estado = EnumRiesgoEstado.RetiradoSinVisar Then
            p_OutKey = "RIESGOSRETIRADOSPORVISAR|" & p_ObjRiesgo.IDRiesgo
        End If
    ElseIf p_TipoRiesgoTarea = EnumTipoRiesgoTarea.Retipificados Then
        p_OutKey = "RIESGOPENDIENTERETIPIFICAR|" & p_ObjRiesgo.IDRiesgo
    Else
        p_Error = "Se ha de seleccionar el tipo de riesgo para tarea"
        Exit Sub
    End If

    ' Edge case: estado no mappable -> set p_Error
    If p_OutKey = "" Then
        p_Error = "Nodo no encontrado"
        Exit Sub
    End If
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        p_Error = "CalidadTareas_SeleccionarNodoRiesgo: " & Err.Description
    End If
End Sub

' =============================================================================
' Public Sub CalidadTareas_CargarArbol
'   Valida que el objeto ArbolTareasCalidad este disponible.
'   Raises: NO. Errores via p_Error ByRef.
' =============================================================================
Public Sub CalidadTareas_CargarArbol( _
    ByVal p_ObjArbolCalidad As ArbolTareasCalidad, _
    Optional ByRef p_Error As String)

    On Error GoTo errores
    p_Error = ""

    If p_ObjArbolCalidad Is Nothing Then
        p_Error = "No se ha inicializado el arbol de tareas de calidad"
        Exit Sub
    End If
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        p_Error = "CalidadTareas_CargarArbol: " & Err.Description
    End If
End Sub
