Attribute VB_Name = "modNotificacionPorCorreoHelper"
Option Compare Database
Option Explicit

' =============================================================================
' modNotificacionPorCorreoHelper.bas
'
' Helper para componer el cuerpo del correo de riesgo materializado.
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
' SDD:     e2e-form-by-form-2026-06-22 - Bloque 2 - REQ-CAL-11
'
' Helper p\u00fablico (1):
'   NotificarPorCorreo - compone el cuerpo HTML (no env\u00eda).
'
' DECISI\u00d3N 2026-06-22 (sdd-apply):
'   * p_ObjRiesgo se tipa As Object (no As riesgo) para que los \u00e1tomos TDD
'     puedan pasar stubs ligeros sin acoplar al ciclo de vida completo de
'     Constructor (un riesgo real exige fila en TbRiesgos + IDEdicion
'     resuelto v\u00eda Constructor.getEdicion).
'   * El par\u00e1metro db se incluye por consistencia con el patr\u00f3n del proyecto
'     (Test_Fixtures.GetTestDb en los \u00e1tomos), aunque el cuerpo es funci\u00f3n pura
'     del objeto y NO consulta la BD.
'   * p_PromptResult se incluye para futura migraci\u00f3n del MsgBox del form
'     Form_FormRiesgoMaterializado.ComandoFechaMaterializado_Click (no se usa
'     en este helper: la composici\u00f3n del cuerpo es no interactiva).
' =============================================================================

' -----------------------------------------------------------------------------
' NotificarPorCorreo
'
' Compone el cuerpo HTML del correo de riesgo materializado. NO env\u00eda el correo.
' Valida que el objeto tenga la cadena .Edicion.Proyecto navegable y delega
' en GetBodyCorreoRiesgoMaterializado (Funciones Generales.bas:270).
'
' Par\u00e1metros:
'   p_ObjRiesgo        - objeto riesgo (o stub con .Edicion.Proyecto.Proyecto
'                        y .Edicion.Proyecto.CodigoDocumento)
'   db                 - DAO.Database (no usado por el cuerpo; presente para
'                        consistencia con el patr\u00f3n del proyecto)
'   p_PromptResult     - reservado para migraci\u00f3n del MsgBox del form
'   p_Error (out)      - descripci\u00f3n del error si la validaci\u00f3n o composici\u00f3n
'                        falla; cadena vac\u00eda si OK
'
' Retorna:
'   Cuerpo HTML (cadena) si OK; "" si p_Error est\u00e1 poblado.
' -----------------------------------------------------------------------------
Public Function NotificarPorCorreo( _
    ByRef p_ObjRiesgo As Object, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_PromptResult As Long, _
    Optional ByRef p_Error As String) As String

    Dim m_Body As String

    On Error GoTo errores
    p_Error = ""
    m_Body = ""

    ' 1. p_ObjRiesgo no debe ser Nothing
    If p_ObjRiesgo Is Nothing Then
        p_Error = "NotificarPorCorreo: p_ObjRiesgo es Nothing"
        NotificarPorCorreo = ""
        Exit Function
    End If

    ' 2. p_ObjRiesgo.Edicion no debe ser Nothing
    If p_ObjRiesgo.Edicion Is Nothing Then
        p_Error = "NotificarPorCorreo: p_ObjRiesgo.Edicion es Nothing"
        NotificarPorCorreo = ""
        Exit Function
    End If

    ' 3. p_ObjRiesgo.Edicion.Proyecto no debe ser Nothing
    If p_ObjRiesgo.Edicion.Proyecto Is Nothing Then
        p_Error = "NotificarPorCorreo: p_ObjRiesgo.Edicion.Proyecto es Nothing"
        NotificarPorCorreo = ""
        Exit Function
    End If

    ' 4. Delegar en GetBodyCorreoRiesgoMaterializado (Public en Funciones Generales.bas)
    m_Body = GetBodyCorreoRiesgoMaterializado(p_ObjRiesgo, p_Error)
    If p_Error <> "" Then
        NotificarPorCorreo = ""
        Exit Function
    End If

    NotificarPorCorreo = m_Body
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "NotificarPorCorreo: " & Err.Number & " - " & Err.description
    End If
    NotificarPorCorreo = ""
End Function

