Attribute VB_Name = "modProyectosGestionHelper"
' ============================================================
' modProyectosGestionHelper.bas
'
' Helper que extrae la logica de negocio de:
'   - Form_FormProyectosGestion.ActualizarProyecto (line 919)
'   - Form_FormProyectosGestion.ListaFiltrados_Click (line 669) [cross-form surface]
'
' Slice HR4-sister-b (HR3d follow-up #9, #3, #6 - 2026-07-01).
'
' El metodo form-side ListaFiltrados_Click queda Private (UI event handler).
' Su superficie cross-form-callable se encapsula en
' ProyectosGestion_ListaFiltrados_Click_Run para que callers externos
' (sibling forms, class modules) usen el helper, NO el form.
'
' El metodo form-side ComandoFiltrar_Click tambien es Private (event handler
' de Access) sin extraccion a helper (per Hard rule 4 exception:
' "event wiring, lifecycle, tiny UI adapters").
'
' Hard rule 1: cero business logic en event handlers o Public methods en forms.
' Hard rule 2: honest signature — solo inyecta lo que la funcion necesita.
' Hard rule 3: form .cls MUST NOT call other form .cls — callers must use helper.
' Hard rule 4: Public business methods en forms son anti-patron.
' Hard rule 5: nombres Public globalmente unicos con ModuleName_ prefix.
' Hard rule 7: NO MsgBox/InputBox en helpers. Errores via ByRef p_Error.
' Hard rule 11: variables declaradas al inicio del modulo / sub.
' ============================================================
Option Compare Database
Option Explicit

' --- Module-level constants (Hard rule 11: declarations at top) ---
' ListaFiltrados header columns in Form_FormProyectosGestion. The column index
' for IDProyecto is 0 (first column). Used by both helper entries.
Private Const LISTA_FILTRADOS_COL_IDPROYECTO As Long = 0

' =============================================================================
' Public Sub ProyectosGestion_ActualizarProyecto_ActualizarLinea
'   Logica de ActualizarProyecto: si p_ObjProyecto es Nothing, no-op.
'   Si p_Form es Nothing, no-op (form cerrado). Si ambos Nothing, no-op.
'   El helper encapsula la logica de "actualizar la fila de ListaFiltrados
'   correspondiente al proyecto" para que el form pueda ser un thin adapter.
'
'   Hard rule 4 OK: el form-side Public Function se hace Private y
'   este Sub queda como el seam publico.
' =============================================================================
Public Sub ProyectosGestion_ActualizarProyecto_ActualizarLinea( _
    ByVal p_Form As Object, _
    ByVal p_ObjProyecto As Proyecto, _
    Optional ByRef p_Error As String)

    On Error GoTo errores
    p_Error = ""

    ' --- Local declarations at top of Sub (Hard rule 11) ---
    Dim m_ObjProyecto As Proyecto
    Dim lst As ListBox
    Dim i As Integer
    Dim m_LineaActual As String
    Dim m_LineaFinal As String

    ' Edge case: ambos Nothing -> no-op sin error
    If p_ObjProyecto Is Nothing Then
        Exit Sub
    End If

    ' Edge case: form cerrado -> no-op sin error
    If p_Form Is Nothing Then
        Exit Sub
    End If

    ' --- Recargar el proyecto (logica original: Constructor.getProyecto) ---
    Set m_ObjProyecto = Constructor.getProyecto(p_ObjProyecto.IDProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjProyecto Is Nothing Then
        Exit Sub
    End If

    ' --- Localizar la fila en ListaFiltrados ---
    Set lst = p_Form.ListaFiltrados
    For i = 1 To lst.ListCount - 1
        If Nz(lst.Column(LISTA_FILTRADOS_COL_IDPROYECTO, i), "") = m_ObjProyecto.IDProyecto Then
            m_LineaActual = Nz(lst.Column(0, i), "") & ";" & Nz(lst.Column(1, i), "") & ";" & Nz(lst.Column(2, i), "") & ";" & _
                        Nz(lst.Column(3, i), "") & ";" & Nz(lst.Column(4, i), "") & ";" & _
                        Nz(lst.Column(5, i), "")
            lst.Selected(i) = True
        End If
    Next

    With m_ObjProyecto
        m_LineaFinal = .IDProyecto & ";" & .Proyecto & ";" & .NombreProyecto & ";" & .Juridica & ";" & _
                            .FechaCierre & ";" & .FechaPrevistaCierreCalculada
    End With

    If m_LineaActual <> m_LineaFinal Then
        lst.RowSource = Replace(lst.RowSource, m_LineaActual, m_LineaFinal)
        lst.Requery
    End If
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        p_Error = "ProyectosGestion_ActualizarProyecto_ActualizarLinea: " & Err.Description
    End If
End Sub

' =============================================================================
' Public Sub ProyectosGestion_ListaFiltrados_Click_Run
'   Replica cross-form-callable de Form_FormProyectosGestion.ListaFiltrados_Click.
'   El form-side ListaFiltrados_Click (Private event handler) hace:
'     - Reset visual labels (lblUltimoArchivoPublicado, ImagenDocumentoUltimPublicacion)
'     - Validar seleccion (ListCount==2 -> selecciona fila 1)
'     - Cargar m_ObjProyectoSeleccionado desde cache o Constructor
'     - Llamar EstablecerBotonera (Private) para habilitar botones
'     - Leer .EdicionUltimaPublicada.URLArchivoInforme y mostrar/ocultar labels
'
'   Esta version NO toca los miembros Private del form (m_ObjProyectoSeleccionado,
'   m_URLInformeUltimaEdicionProyectoSeleccionado, EstablecerBotonera). Solo opera
'   sobre controles publicos (ListaFiltrados, lblUltimoArchivoPublicado,
'   ImagenDocumentoUltimPublicacion) y re-deriva el URLArchivoInforme desde el
'   proyecto mismo.
'
'   Hard rule 3: callers cross-form deben usar este helper, no
'   Form_FormProyectosGestion.ListaFiltrados_Click directamente.
'   Hard rule 7: no MsgBox. Errores via ByRef p_Error.
' =============================================================================
Public Sub ProyectosGestion_ListaFiltrados_Click_Run( _
    ByVal p_Form As Object, _
    Optional ByRef p_Error As String)

    On Error GoTo errores
    p_Error = ""

    ' --- Local declarations at top of Sub (Hard rule 11) ---
    Dim lst As ListBox
    Dim m_IDProyectoSeleccionado As String
    Dim m_ObjProyectoSeleccionado As Proyecto
    Dim m_URLInforme As String

    ' Edge case: form cerrado o Nothing -> no-op sin error
    If p_Form Is Nothing Then
        Exit Sub
    End If

    ' --- Reset visual labels (public controls only) ---
    p_Form.lblUltimoArchivoPublicado.Visible = False
    p_Form.ImagenDocumentoUltimPublicacion.Visible = False

    Set lst = p_Form.ListaFiltrados
    If lst Is Nothing Then
        Exit Sub
    End If

    ' Edge case: listbox vacio -> no-op
    If lst.ListCount < 1 Then
        Exit Sub
    End If

    ' Re-seleccionar fila 1 si la lista tiene 2 filas (header + 1 row)
    ' (form-side ListaFiltrados_Click hace lo mismo)
    If lst.ListCount = 2 Then
        lst.Selected(1) = True
    End If

    ' --- Identificar proyecto seleccionado ---
    m_IDProyectoSeleccionado = Nz(lst.Column(LISTA_FILTRADOS_COL_IDPROYECTO), "")
    If m_IDProyectoSeleccionado = "" Then
        Exit Sub
    End If

    ' --- Cargar el proyecto (re-derive el state que el form cachea en private) ---
    Set m_ObjProyectoSeleccionado = GetCachedProyecto(m_IDProyectoSeleccionado, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjProyectoSeleccionado Is Nothing Then
        Exit Sub
    End If

    ' --- Re-derivar URLArchivoInforme (lo que el form hace via EdicionUltimaPublicada) ---
    m_URLInforme = ""
    If Not m_ObjProyectoSeleccionado.EdicionUltimaPublicada Is Nothing Then
        m_URLInforme = m_ObjProyectoSeleccionado.EdicionUltimaPublicada.URLArchivoInforme
    End If

    ' --- Actualizar labels segun el URL ---
    If m_URLInforme <> "" Then
        p_Form.lblUltimoArchivoPublicado.Visible = True
        p_Form.ImagenDocumentoUltimPublicacion.Visible = True
    End If

    Exit Sub

errores:
    If Err.Number <> 1000 Then
        p_Error = "ProyectosGestion_ListaFiltrados_Click_Run: " & Err.Description
    End If
End Sub