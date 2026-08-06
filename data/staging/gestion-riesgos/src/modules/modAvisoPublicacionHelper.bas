Attribute VB_Name = "modAvisoPublicacionHelper"
Option Compare Database
Option Explicit

' =============================================================================
' modAvisoPublicacionHelper.bas
'
' Helper para registrar el aviso de publicación de una edición.
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
' SDD:     e2e-form-by-form-2026-06-22 - Bloque 2 - REQ-CAL-16A
'
' Helper público (1):
'   AvisarPublicacion - inserta una fila en TbCorreosEnviados notificando
'                       la publicación de la edición (NO envía el correo).
'
' DECISIÓN 2026-06-22 (sdd-apply, PR-2 sobre PR-1):
'   * p_ObjEdicion se tipa As Object (no As Edicion) para que los átomos TDD
'     puedan pasar stubs ligeros sin acoplar al ciclo de vida completo de
'     Constructor (un Edicion real exige fila en TbProyectosEdiciones).
'   * IDEdicion en TbCorreosEnviados es Long post-issue-70 (id-70 migration).
'     El objeto llega con .IDEdicion como String (convención del proyecto
'     p_ID* As String — ver Edicion.cls línea 18); el helper hace CLng()
'     internamente para evitar overflow cuando IDEdicion > 32,767 (Integer
'     max). Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax
'     ya cubre la migración del schema; este helper cubre la ruta VBA ? schema.
'   * p_PromptResult se incluye para futura migración del MsgBox del form
'     Form_FormPublicacionCalidadPublicarEjecutar.ComandoPublicar_Click
'     (no se usa en este helper: la inserción es no interactiva).
'   * m_IDAplicacion / m_ObjEntorno NO se tocan (DB-only write thin helper).
' =============================================================================

' -----------------------------------------------------------------------------
' AvisarPublicacion
'
' Inserta una fila en TbCorreosEnviados con la marca de publicación de la
' edición. NO envía correo; el envío queda fuera del scope (responsabilidad
' del caller, que recibe el conteo de filas escrito y decide qué hacer).
'
' Parámetros:
'   p_ObjEdicion      - objeto con propiedad .IDEdicion (String por convención
'                       del proyecto, ver Edicion.cls línea 18). El helper
'                       hace CLng() internamente para escribir como Long y
'                       aceptar valores > 32,767 sin overflow.
'   db                - DAO.Database inyectado por el átomo TDD. Si Nothing,
'                       usa CurrentDb (patrón del proyecto).
'   p_PromptResult    - Long, ByRef. Reservado para futura migración del
'                       MsgBox del form. No se usa en este helper.
'   p_Error (out)     - String, ByRef. Descripción del error si la validación
'                       o INSERT falla; cadena vacía si OK.
'
' Retorna:
'   1L  si la fila se insertó en TbCorreosEnviados.
'   0L  si p_Error está poblado (validación o INSERT falló).
'
' Pre-condiciones:
'   - p_ObjEdicion NO debe ser Nothing.
'   - p_ObjEdicion.IDEdicion debe ser parseable como Long (no vacío, numérico).
'
' Post-condiciones (éxito):
'   - Una fila en TbCorreosEnviados con:
'       IDEdicion       = CLng(p_ObjEdicion.IDEdicion)  [Long, NO String]
'       Aplicacion      = "Gestion_Riesgos"
'       Asunto          = "Publicación de edición " & IDEdicion
'       FechaEnvio      = Now()
'       FechaGrabacion  = Now()
'   - Otros campos quedan en su default del schema (NULL o autonum).
' -----------------------------------------------------------------------------
Public Function AvisarPublicacion( _
    ByRef p_ObjEdicion As Object, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_PromptResult As Long, _
    Optional ByRef p_Error As String) As Long

    Dim m_Db As DAO.Database
    Dim m_IDEdicion As Long
    Dim m_Asunto As String

    On Error GoTo errores
    p_Error = ""
    AvisarPublicacion = 0

    ' --- 1. Validación: p_ObjEdicion no debe ser Nothing ---
    If p_ObjEdicion Is Nothing Then
        p_Error = "AvisarPublicacion: p_ObjEdicion es Nothing"
        Exit Function
    End If

    ' --- 2. Validación: .IDEdicion debe existir y ser parseable como Long ---
    '     Convención del proyecto: .IDEdicion es String (ver Edicion.cls línea 18),
    '     pero la columna TbCorreosEnviados.IDEdicion es Long post-issue-70.
    '     Por eso hacemos CLng() explícito: si .IDEdicion > 32,767 y la columna
    '     sigue siendo Integer, la inserción falla con DAO error 3035 (overflow).
    '     Si la columna ya es Long (post-migration), CLng preserva el valor.
    Dim m_IDEdicionStr As String
    m_IDEdicionStr = Nz(CStr(p_ObjEdicion.IDEdicion), "")
    If Len(m_IDEdicionStr) = 0 Then
        p_Error = "AvisarPublicacion: p_ObjEdicion.IDEdicion está vacío"
        Exit Function
    End If
    If Not IsNumeric(m_IDEdicionStr) Then
        p_Error = "AvisarPublicacion: p_ObjEdicion.IDEdicion no es numérico (""" & m_IDEdicionStr & """)"
        Exit Function
    End If
    m_IDEdicion = CLng(m_IDEdicionStr)

    ' --- 3. Resolver db (inyectado por el átomo, o CurrentDb en producción) ---
    If db Is Nothing Then
        Set m_Db = CurrentDb
    Else
        Set m_Db = db
    End If

    ' --- 4. INSERT en TbCorreosEnviados ---
    '    IDEdicion como Long literal (sin comillas) — Long, no String.
    '    Asunto lleva el IDEdicion concatenado para identificación rápida en UI.
    '    FechaEnvio y FechaGrabacion = Now() para trazabilidad inmediata.
    m_Asunto = "Publicación de edición " & CStr(m_IDEdicion)

    m_Db.Execute "INSERT INTO TbCorreosEnviados " & _
                 "(IDEdicion, Aplicacion, Asunto, FechaEnvio, FechaGrabacion) " & _
                 "VALUES (" & m_IDEdicion & ", 'Gestion_Riesgos', " & _
                 """" & m_Asunto & """, " & _
                 "#" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#, " & _
                 "#" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    AvisarPublicacion = 1
    Exit Function

errores:
    ' El error 1000 lo usan otros helpers del proyecto como sentinel de
    ' "error funcional" — si viene de fuera (DAO, COM), preservamos el código.
    If Err.Number <> 1000 Then
        p_Error = "AvisarPublicacion: " & Err.Number & " - " & Err.description
    Else
        p_Error = "AvisarPublicacion: " & Err.description
    End If
    AvisarPublicacion = 0
End Function

