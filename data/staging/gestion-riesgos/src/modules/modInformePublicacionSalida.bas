Attribute VB_Name = "modInformePublicacionSalida"
Option Compare Database
Option Explicit

' ============================================================
' modInformePublicacionSalida.bas
'
' Helper para validar el tipo de salida del informe de publicacion.
' Project: gestion_riesgos
' SDD:     e2e-form-by-form-2026-06-22 - Bloque 4 - T-R.2
'
' Helper publico (1):
'   ValidarTipoSalidaInformePublicacion - bloquea Excel (eliminado del
'                                         proyecto); permite html y PDF.
'
' DECISION 2026-06-23 (issue-78):
'   * Excel se elimino de los tipos de salida permitidos. La generacion
'     de Excel para publicacion ya no existe (ver tambien el modulo
'     ExcelInforme.bas que fue removido).
'   * PDF se permite cuando el contexto lo requiera (ver tambien
'     InformePublicacionDebeGenerarPDF en Funciones Generales.bas).
'   * El atomo Test_InformePublicacionSalida_BloqueaExcel exige que el
'     mensaje de error contenga la palabra "Excel" (case-insensitive).
' ============================================================

' -----------------------------------------------------------------------------
' ValidarTipoSalidaInformePublicacion
'
' Valida que el tipo de salida solicitado para el informe de publicacion sea
' uno de los permitidos. Bloquea Excel; permite html.
'
' Parametros:
'   p_TipoInforme      - EnumTipoInformePublicacion: tipo solicitado.
'   p_Error (out)      - String, ByRef. Mensaje de error si se bloquea;
'                        cadena vacia si el tipo es valido.
'
' Retorna:
'   True   si el tipo es valido (html).
'   False  si el tipo es invalido (Excel u otro no reconocido).
'
' Pre-condiciones: ninguna.
'
' Post-condiciones (exito):
'   - p_Error queda en "".
'   - Funcion retorna True.
'
' Post-condiciones (bloqueo):
'   - p_Error contiene la palabra "Excel" cuando se bloquea ese tipo.
'   - Funcion retorna False.
' -----------------------------------------------------------------------------
Public Function ValidarTipoSalidaInformePublicacion( _
    ByVal p_TipoInforme As EnumTipoInformePublicacion, _
    Optional ByRef p_Error As String) As Boolean

    Dim m_Mensaje As String

    On Error GoTo errores
    p_Error = ""
    m_Mensaje = ""

    Select Case p_TipoInforme
        Case EnumTipoInformePublicacion.html
            ValidarTipoSalidaInformePublicacion = True

        Case EnumTipoInformePublicacion.Excel
            m_Mensaje = "Excel no esta permitido como tipo de salida del informe de publicacion. Use html o PDF."
            ValidarTipoSalidaInformePublicacion = False

        Case Else
            m_Mensaje = "Tipo de informe de publicacion no reconocido (valor=" & CStr(p_TipoInforme) & "). Permitidos: html, PDF."
            ValidarTipoSalidaInformePublicacion = False
    End Select

    p_Error = m_Mensaje
    Exit Function

errores:
    p_Error = "ValidarTipoSalidaInformePublicacion: " & Err.Number & " - " & Err.Description
    ValidarTipoSalidaInformePublicacion = False
End Function
