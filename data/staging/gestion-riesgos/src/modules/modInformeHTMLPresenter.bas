Attribute VB_Name = "modInformeHTMLPresenter"
Option Compare Database
Option Explicit

' =============================================================================
' modInformeHTMLPresenter.bas
'
' Helper module "presenter" para resolver la etiqueta del encabezado de fecha y
' el prefijo per-row/per-cell de fecha que los informes HTML y Excel muestran
' en el cuadro de control y en la tabla principal de riesgos.
'
' Project: gestion_riesgos
' Branch:  feat/metodologia-e2e-riesgos-2026-06-19
' SDD:     e2e-form-by-form-2026-06-22 - Bloque 4 - Refactor inline - T-R.1
'
' Helpers públicos (2):
'   1. ObtenerEncabezadoFecha   - devuelve el texto del encabezado <th> de la
'                                columna de fecha, en función del tipo de
'                                informe (tabla histórica vs tabla principal
'                                de riesgos).
'   2. ObtenerPrefijoFechaCelda - devuelve el prefijo por celda ("Pub. " /
'                                "Creac. " / "") que antecede a la fecha
'                                formateada en cada celda de la tabla.
'
' Decisión 2026-06-22 (sdd-apply Bloque 4 / Refactor inline / T-R.1):
'   * Helper PURE - NO DAO, NO CurrentDb, NO recordsets. Las firmas reflejan
'     honestamente lo que el helper realmente depende (per skill
'     access-vba-e2e-methodology §hard rule 2): solo strings de entrada y
'     strings de salida. NO se incluye `Optional ByRef p_Error As String`
'     porque no hay modo de fallo relevante (un `p_TipoInforme` desconocido
'     cae al default explícito, no es error).
'   * Esto cumple el contrato de la skill access-vba-e2e-methodology §hard
'     rule 1: forms thin. Los informes (InformeRiesgoHTML.bas) ya no
'     contienen los literales "Fecha pub. / creación", "Pub. ", "Creac. "
'     inline; llaman a este helper.
'
'     (Históricamente también ExcelInforme.bas los consumía. Ese módulo
'     se eliminó: la generación de Excel para publicación ya no existe.)
'   * Decisión de TIPO DE INFORME: el task sólo conoce dos tipos
'     ("EdicionHistorico" y "RiesgoEstado"). Si llega un tipo desconocido,
'     el helper retorna "Fecha" (genérico) sin error — defensa explícita
'     contra typos en los call sites.
'   * Per skill access-vba-tdd §1.6: módulo `modInformeHTMLPresenter` ?
'     funciones públicas `ObtenerEncabezadoFecha` y
'     `ObtenerPrefijoFechaCelda` — el módulo usa el prefijo `mod`, las
'     funciones no. Esto evita el error de compilación VBA "Se esperaba
'     una variable o un procedimiento, no un módulo".
'   * p_TipoInforme es case-insensitive (UCase$ + Trim$). p_FechaPublicacion
'     y p_FechaEdicion se inspeccionan con IsDate (defensa: si vienen como
'     Null o vacío o string no-fecha, IsDate retorna False).
'
' Reglas (acordadas con Calidad pre-PR-7):
'   - ObtenerEncabezadoFecha("EdicionHistorico") ? "Fecha pub. / creación"
'       (header del Cuadro de Control del informe HTML / Excel: la celda
'        puede contener FechaPublicacion o FechaEdicion según la edición).
'   - ObtenerEncabezadoFecha("RiesgoEstado")     ? "Fecha"
'       (header de la tabla principal de riesgos: celda estándar, sin
'        ambigüedad entre publicación y creación).
'   - ObtenerEncabezadoFecha(<otro>)            ? "Fecha"
'       (default defensivo; no es error).
'   - ObtenerPrefijoFechaCelda(...)
'       * Si IsDate(p_FechaPublicacion)           ? "Pub. "
'       * ElseIf IsDate(p_FechaEdicion)           ? "Creac. "
'       * Else                                    ? ""
' =============================================================================

' -----------------------------------------------------------------------------
' ObtenerEncabezadoFecha
'
' Resuelve el texto del encabezado <th> de la columna de fecha del informe.
' PURE — depende solo de p_TipoInforme (string).
'
' Parámetros:
'   p_TipoInforme (String)   - Identificador del tipo de informe:
'                                "EdicionHistorico" ? "Fecha pub. / creación"
'                                "RiesgoEstado"     ? "Fecha"
'                                otro / vacío       ? "Fecha" (default defensivo)
'
' Retorna:
'   El texto del encabezado para el tipo de informe solicitado. Nunca vacío.
' -----------------------------------------------------------------------------
Public Function ObtenerEncabezadoFecha(ByVal p_TipoInforme As String) As String
    Dim m_Tipo As String
    m_Tipo = UCase$(Trim$(Nz(p_TipoInforme, "")))

    Select Case m_Tipo
        Case "EDICIONHISTORICO"
            ObtenerEncabezadoFecha = "Fecha pub. / creación"
        Case "RIESGOESTADO"
            ObtenerEncabezadoFecha = "Fecha"
        Case Else
            ' Default defensivo: tipo desconocido o vacío ? "Fecha".
            ' No es error; es contrato explícito del helper.
            ObtenerEncabezadoFecha = "Fecha"
    End Select
End Function

' -----------------------------------------------------------------------------
' ObtenerPrefijoFechaCelda
'
' Resuelve el prefijo por celda que antecede a la fecha formateada en cada
' celda de la tabla del cuadro de control (HTML y Excel). PURE — depende
' solo de los strings de fecha pasados.
'
' Parámetros:
'   p_TipoInforme      (String)  - Identificador del tipo de informe.
'                                   Aceptado por consistencia con
'                                   ObtenerEncabezadoFecha (mismo call site),
'                                   pero el contrato actual es independiente
'                                   del tipo: el prefijo se resuelve solo por
'                                   presencia de fecha.
'   p_FechaPublicacion (String)  - Valor de FechaPublicacion (puede ser Null,
'                                   cadena vacía, o fecha en formato
'                                   reconocible por IsDate).
'   p_FechaEdicion     (String)  - Valor de FechaEdicion (idem).
'
' Retorna:
'   "Pub. "    si IsDate(p_FechaPublicacion).
'   "Creac. "  si IsDate(p_FechaEdicion).
'   ""         si ninguna de las dos es fecha (caso de edición vacía o
'              fila sin información temporal).
' -----------------------------------------------------------------------------
Public Function ObtenerPrefijoFechaCelda( _
    ByVal p_TipoInforme As String, _
    ByVal p_FechaPublicacion As String, _
    ByVal p_FechaEdicion As String) As String

    ' VBA no hace short-circuit de Or/And (vba-access §1.6.1); separar checks.
    If IsDate(Nz(p_FechaPublicacion, "")) Then
        ObtenerPrefijoFechaCelda = "Pub. "
    ElseIf IsDate(Nz(p_FechaEdicion, "")) Then
        ObtenerPrefijoFechaCelda = "Creac. "
    Else
        ObtenerPrefijoFechaCelda = ""
    End If
End Function

