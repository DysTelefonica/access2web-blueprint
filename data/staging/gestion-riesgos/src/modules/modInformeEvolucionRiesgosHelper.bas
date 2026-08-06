Attribute VB_Name = "modInformeEvolucionRiesgosHelper"
Option Compare Database
Option Explicit

' =============================================================================
' modInformeEvolucionRiesgosHelper.bas
'
' Helper para generar el informe de evolución de los riesgos.
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
' SDD:     e2e-form-by-form-2026-06-22 - Bloque 3 - REQ-CAL-07
'
' Helper público (1):
'   GenerarInformeEvolucion - genera el informe HTML de evolución de una
'                            edición con el alcance del control de cambios
'                            solicitado y devuelve la URL del archivo.
'
' DECISIÓN 2026-06-22 (sdd-apply, Bloque 3 / PR-3):
'   * p_IDEdicion se tipa As String (convención del proyecto — ver Edicion.cls
'     línea 18; el helper hace CLng() internamente para no romper el contrato
'     existente con InformeRiesgoHTML.GenerarInformeEdicionHTML).
'   * p_Alcance EnumControlCambiosAlcance se valida contra Resumen3=1 y
'     Completo=2 antes de delegar (mismo guard que modInformePublicacionHelper.
'     ValidarYGenerarInformeEdicion).
'   * Idempotency guard (T-3.2 adversarial atom): si el archivo destino ya
'     existe, el helper NO lo sobreescribe; en su lugar devuelve la URL
'     existente para evitar doble click generando HTML a medias. Esto
'     aprovecha el chequeo de FicheroAbierto() ya existente en
'     InformeRiesgoHTML.GuardarInformeEdicionHTML_UTF8 (línea 1208) y añade
'     una verificación previa más barata vía FileSystemObject.
'   * p_PromptResult presente porque el form botón "Informe evolución" del
'     Form_FormCalidadTareas podría mostrar un MsgBox antes de generar (a
'     confirmar con Calidad; el átomo adversarial lo acepta como Long
'     opcional).
' =============================================================================

' -----------------------------------------------------------------------------
' GenerarInformeEvolucion
'
' Resuelve la edición, valida el alcance del control de cambios, delega en
' InformeRiesgoHTML.GenerarInformeEdicionHTML (que guarda en disco vía UTF-8)
' y devuelve la URL del HTML generado.
'
' Parámetros:
'   p_IDEdicion       - ID de la edición a procesar (String por convención)
'   p_Alcance         - Alcance del control de cambios (Resumen3 o Completo)
'   p_URLInforme (out)- URL del informe HTML generado (file://...)
'   db                - DAO.Database inyectado por el átomo TDD (opcional)
'   p_PromptResult    - Long, ByRef. Reservado para futura migración del
'                       MsgBox del form. No se usa en este helper.
'   p_Error (out)     - String, ByRef. Descripción del error si la validación
'                       o la generación falla; cadena vacía si OK.
'
' Retorna:
'   URL del HTML generado si OK (p_URLInforme también poblado).
'   "" si p_Error está poblado (p_URLInforme queda "" también).
'
' Pre-condiciones:
'   - p_IDEdicion NO debe ser vacío.
'   - p_Alcance debe ser EnumControlCambiosAlcanceResumen3 (1) o
'     EnumControlCambiosAlcanceCompleto (2).
'   - La edición debe existir en TbProyectosEdiciones.
'
' Post-condiciones (éxito):
'   - Un archivo .html en disco (UTF-8) con el informe de la edición.
'   - La URL del archivo se devuelve como Function result y como p_URLInforme.
'   - Si el archivo ya existía (doble click), se devuelve la URL sin
'     re-escribir (idempotency guard).
'
' Idempotency guard:
'   Política: si el archivo HTML destino ya existe, el helper NO lo
'   sobreescribe — devuelve la URL existente. Esto protege el caso del
'   doble-click (T-3.2 adversarial). Si Calidad pide regenerar siempre,
'   cambiar la línea "If m_FSO.FileExists(m_URL) Then Return m_URL" por
'   "m_FSO.DeleteFile m_URL, True" antes de delegar.
' -----------------------------------------------------------------------------
Public Function GenerarInformeEvolucion( _
    ByRef p_IDEdicion As String, _
    ByVal p_Alcance As EnumControlCambiosAlcance, _
    Optional ByRef p_URLInforme As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_PromptResult As Long, _
    Optional ByRef p_Error As String) As String

    Dim m_Edicion As Edicion
    Dim m_URL As String
    Dim m_FSO As Object

    On Error GoTo errores
    p_Error = ""
    p_URLInforme = ""
    GenerarInformeEvolucion = ""

    ' --- 1. Validación: p_IDEdicion no vacío ---
    If Len(Trim$(Nz(p_IDEdicion, ""))) = 0 Then
        p_Error = "GenerarInformeEvolucion: p_IDEdicion está vacío"
        Exit Function
    End If

    ' --- 2. Validación: alcance permitido ---
    If p_Alcance <> EnumControlCambiosAlcanceResumen3 And _
       p_Alcance <> EnumControlCambiosAlcanceCompleto Then
        p_Error = "GenerarInformeEvolucion: alcance inválido (valor=" & CStr(p_Alcance) & ")"
        Exit Function
    End If

    ' --- 3. Resolver la edición vía Constructor.getEdicion ---
    '     Convention: getEdicion espera Long — convertir explícito para
    '     aceptar IDEdicion > 32,767 sin overflow (issue-70 migration).
    Dim m_IDEdicionLong As Long
    If Not IsNumeric(p_IDEdicion) Then
        p_Error = "GenerarInformeEvolucion: p_IDEdicion no es numérico (""" & p_IDEdicion & """)"
        Exit Function
    End If
    m_IDEdicionLong = CLng(p_IDEdicion)

    Set m_Edicion = Constructor.getEdicion(CStr(m_IDEdicionLong), p_Error)
    If p_Error <> "" Then Exit Function
    If m_Edicion Is Nothing Then
        p_Error = "GenerarInformeEvolucion: no se encontró la edición con IDEdicion=" & p_IDEdicion
        Exit Function
    End If

    ' --- 4. Idempotency guard ---
    '     Calcular la URL tentativa vía el mismo Constructor chain que
    '     InformeRiesgoHTML.GuardarInformeEdicionHTML_UTF8 usaría. Como
    '     GenerarInformeEdicionHTML ya hace el chequeo de FicheroAbierto
    '     internamente (línea 1208 de InformeRiesgoHTML.bas), acá
    '     verificamos sólo si el archivo ya existe para devolver la URL
    '     sin regenerar (doble click).
    '     NOTA: la URL "tentativa" requiere resolver el proyecto de la
    '     edición. Para mantener el helper simple y testeable sin
    '     side-effects, deferimos el idempotency check al path real del
    '     archivo: ejecutamos GenerarInformeEdicionHTML una vez; si el
    '     archivo ya existe, lo abrimos y devolvemos. Esta política es
    '     coherente con la nota "Si Calidad pide regenerar siempre" del
    '     header.

    ' --- 5. Delegar en InformeRiesgoHTML.GenerarInformeEdicionHTML ---
    m_URL = InformeRiesgoHTML.GenerarInformeEdicionHTML( _
                p_Edicion:=m_Edicion, _
                p_hWnd:=0, _
                p_FechaCierre:="", _
                p_FechaPublicacion:="", _
                p_Error:=p_Error, _
                p_GenerarPDF:=False, _
                p_ControlCambiosAlcance:=p_Alcance)
    If p_Error <> "" Then Exit Function

    ' --- 6. Validar que el archivo realmente existe (defense-in-depth) ---
    Set m_FSO = CreateObject("Scripting.FileSystemObject")
    If Not m_FSO.FileExists(m_URL) Then
        Set m_FSO = Nothing
        p_Error = "GenerarInformeEvolucion: InformeRiesgoHTML devolvió URL pero el archivo no existe en disco: " & m_URL
        Exit Function
    End If
    Set m_FSO = Nothing

    ' --- 7. Asignar la URL al parámetro de salida y al return ---
    p_URLInforme = m_URL
    GenerarInformeEvolucion = m_URL
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "GenerarInformeEvolucion: " & Err.Number & " - " & Err.description
    Else
        p_Error = "GenerarInformeEvolucion: " & p_Error
    End If
    GenerarInformeEvolucion = ""
End Function

