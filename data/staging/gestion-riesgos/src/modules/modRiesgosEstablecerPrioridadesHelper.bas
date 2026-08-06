Attribute VB_Name = "modRiesgosEstablecerPrioridadesHelper"
' ============================================================
' modRiesgosEstablecerPrioridadesHelper.bas
'
' Helper que extrae la logica de negocio de:
'   - Form_FormRiesgosEstablecerPrioridades.VaciarTbAux   (line 411)
'   - Form_FormRiesgosEstablecerPrioridades.RellenarTbAux (line 425)
'
' Slice HR4-sister (HR3d follow-up #9, #2 - 2026-07-01).
'
' Hard rule 1: cero business logic en event handlers o Public methods en forms.
' Hard rule 2: honest signature — solo inyecta lo que la funcion necesita.
' Hard rule 4: Public business methods en forms son anti-patron (VaciarTbAux /
'               RellenarTbAux deben ser Private despues de extraer).
' Hard rule 5: nombres Public globalmente unicos con ModuleName_ prefix.
' Hard rule 7: NO MsgBox/InputBox en helpers. Errores via ByRef p_Error.
'
' Capa:
'   - RiesgoPrioridad_VaciarTbAux(p_Error)           — thin wrapper (5-line body)
'   - RiesgoPrioridad_RellenarTbAux(...)             — pure compute (50-line body)
' ============================================================
Option Compare Database
Option Explicit

' =============================================================================
' Public Sub RiesgoPrioridad_VaciarTbAux
'   Thin wrapper sobre Priorizacion_ClearTemp. Extraido del form para
'   que el caller form-side pueda ser Private.
'   Hard rule 4 OK: el form-side Public Function se hace Private, y
'   este Sub queda como el seam publico para que otros modulos
'   (test, coord) puedan limpiar la TbAuxPriorizacion.
' =============================================================================
Public Sub RiesgoPrioridad_VaciarTbAux(Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not Priorizacion_ClearTemp(p_Error) Then
        Err.Raise 1000
    End If
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        p_Error = "RiesgoPrioridad_VaciarTbAux: " & Err.Description
    End If
End Sub

' =============================================================================
' Public Sub RiesgoPrioridad_RellenarTbAux
'   Pure compute: itera p_ColRiesgosActivos, vacia primero, luego inserta
'   snapshot rows via Priorizacion_InsertSnapshotRow.
'   No toca Forms() (cumple Hard rule 3).
'   Raises: NO. Errores via p_Error ByRef.
' =============================================================================
Public Sub RiesgoPrioridad_RellenarTbAux( _
    ByVal p_ColRiesgosActivos As Scripting.Dictionary, _
    ByVal p_ProyConRiesgosBiblioteca As EnumSiNo, _
    Optional ByRef p_Error As String)

    On Error GoTo errores
    p_Error = ""

    RiesgoPrioridad_VaciarTbAux p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    If p_ColRiesgosActivos Is Nothing Then
        Exit Sub
    End If

    Dim m_ID As Variant
    Dim m_Riesgo As riesgo
    Dim m_Pri As String
    Dim m_PriEdAnterior As String
    Dim m_Descripcion As String
    Dim m_CausaRaiz As String

    For Each m_ID In p_ColRiesgosActivos
        Set m_Riesgo = p_ColRiesgosActivos(m_ID)
        m_Descripcion = m_Riesgo.Descripcion

        m_Pri = m_Riesgo.Priorizacion
        If Not IsNumeric(m_Pri) Then m_Pri = "-"
        m_PriEdAnterior = m_Riesgo.PriorizacionEdAnterior
        If p_ProyConRiesgosBiblioteca = EnumSiNo.Sí Then
            m_CausaRaiz = m_Riesgo.CausaRaiz
        Else
            m_CausaRaiz = vbNullString
        End If

        If Not Priorizacion_InsertSnapshotRow(CLng(m_ID), m_Pri, m_PriEdAnterior, _
            m_Riesgo.CodigoRiesgo, m_Descripcion, m_CausaRaiz, p_Error) Then
            Err.Raise 1000
        End If

        Set m_Riesgo = Nothing
    Next
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        p_Error = "RiesgoPrioridad_RellenarTbAux: " & Err.Description
    End If
End Sub
