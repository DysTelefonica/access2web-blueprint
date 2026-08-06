Attribute VB_Name = "modResponsableFiltroHelper"
Option Compare Database
Option Explicit

' -----------------------------------------------------------------------------
' modResponsableFiltroHelper
'
' Helpers PUROS (sin UI, sin SQL) que envuelven la traduccion entre el valor
' que expone un combo responsable (`Calidad` / `Seguridad`) y el filtro
' aplicado a la busqueda de expedientes.  Toda la logica de mapeo vive
' aqui para que sea unit-testeable sin abrir un Form.
'
' Casos cubiertos (per #71 AC10 / access-vba-e2e-methodology):
'   * BC=1 vs BC=2: el combo puede exponer el ID como valor (BoundColumn=1,
'     columna 1 = "id"), o exponer el "id;nombre" como valor (BoundColumn=2,
'     columna 1 = "id;nombre", columna 2 = "id").  El helper acepta ambas
'     formas y siempre extrae el ID canonico (string).
'   * "Todos": sentinel para "sin filtro" -> "" (string vacio).
'   * Empty: combo sin seleccionar -> "" (string vacio).
'   * "0" / "N/A": el sentinel UX de Seguridad -> "" (string vacio).
'   * No-match: ID que no esta en el Dictionary -> "" (string vacio) y la
'     firma permite al caller decidir si renderiza un row extra o lo omite.
'
' Forma de uso esperada:
'   ' En Form_FormExpedientesGestion.Filtrar() (donde se aplicaba el antiguo
'   ' MarcoCalidad 1..7 hardcodeado):
'   Dim m_IDCalidad As String
'   m_IDCalidad = modResponsableFiltroHelper.ResponsableFiltro_ValorAId( _
'                       Me.IDResponsableCalidad.value, p_Error)
'   .ResponsableCalidad = m_IDCalidad
'
'   ' Y a la inversa, en setExpBusqueda (donde se queria mostrar el ID):
'   Dim m_Row As String
'   m_Row = modResponsableFiltroHelper.ResponsableFiltro_IdACmbRow( _
'               .ResponsableCalidad, m_DictCalidad, p_ConSentinel:=False)
'   Me.IDResponsableCalidad = m_Row
'
' Reglas de diseno:
'   1. Sin SQL.  Sin DAO.Database.  Sin referencia a Forms / Screen.  Sin
'      Set / Get Forms(...) -- esto lo hace el caller.  Pura traduccion
'      string <-> string.
'   2. Err.Raise 1000 solo si se incumple la precondicion del parametro
'      (no por valores esperados -- los "edge cases" como "Todos", "" o
'      "0;N/A" son entradas validas y devuelven "" sin error).
'   3. Convencion de error canonico: `Optional ByRef p_Error As String` es
'      el ultimo parametro y solo se setea si hay error real.
' -----------------------------------------------------------------------------

' UX sentinel de Seguridad ("sin responsable").  Mantener sincronizado con
' `modResponsablePorRolService.SENTINEL_SEG_ID` / `SENTINEL_SEG_LABEL`.
Public Const RESPONSABLE_FILTRO_TODOS As String = "Todos"
Public Const RESPONSABLE_FILTRO_NA_ID As String = "0"
Public Const RESPONSABLE_FILTRO_NA_LABEL As String = "N/A"

Private m_Dummy As String


' -----------------------------------------------------------------------------
' API publica
' -----------------------------------------------------------------------------

' Devuelve True si `p_Valor` representa el sentinel "Todos" (sin filtro).
' Acepta como entrada:
'   * Empty string ""
'   * Null
'   * La cadena "Todos" (case-insensitive)
'   * La cadena "0;N/A" (sentinel UX de Seguridad, cuando el combo expone
'     "id;nombre" y el sentinel esta como primer AddItem)
'   * El literal "0" (cuando el combo expone solo el ID y BC=1)
Public Function ResponsableFiltro_EsTodos(ByVal p_Valor As Variant) As Boolean
    Dim s As String
    On Error GoTo errores

    s = Trim$(CStr(Nz(p_Valor, "")))
    If Len(s) = 0 Then
        ResponsableFiltro_EsTodos = True
        Exit Function
    End If
    If StrComp(s, RESPONSABLE_FILTRO_TODOS, vbTextCompare) = 0 Then
        ResponsableFiltro_EsTodos = True
        Exit Function
    End If
    If StrComp(s, RESPONSABLE_FILTRO_NA_ID & ";" & RESPONSABLE_FILTRO_NA_LABEL, _
                vbTextCompare) = 0 Then
        ResponsableFiltro_EsTodos = True
        Exit Function
    End If
    If StrComp(s, RESPONSABLE_FILTRO_NA_ID, vbTextCompare) = 0 Then
        ResponsableFiltro_EsTodos = True
        Exit Function
    End If
    ResponsableFiltro_EsTodos = False
    Exit Function

errores:
    ResponsableFiltro_EsTodos = True   ' Tratamos cualquier error como "sin filtro"
End Function


' Mapea el `p_Valor` expuesto por un combo responsable a un ID canonico
' (string).  Si el valor representa "sin filtro", devuelve "".
'
' Formatos aceptados para `p_Valor`:
'   * "" / Null / "Todos" / "0;N/A" / "0"   -> ""
'   * "<ID>"        (BC=1; columna ID)       -> "<ID>"
'   * "<ID>;<nombre>" (BC=2; col id;nombre)  -> "<ID>"
'
'   p_Error    : parametro de error estandar del proyecto.  Solo se setea
'                si `p_Valor` viola el contrato (no string convertible).
Public Function ResponsableFiltro_ValorAId( _
                                            ByVal p_Valor As Variant, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim s As String
    Dim m_Pos As Long
    On Error GoTo errores

    p_Error = ""

    If IsNull(p_Valor) Then
        ResponsableFiltro_ValorAId = ""
        Exit Function
    End If

    s = Trim$(CStr(p_Valor))
    If Len(s) = 0 Then
        ResponsableFiltro_ValorAId = ""
        Exit Function
    End If

    ' Sentinel "Todos" -> ""
    If StrComp(s, RESPONSABLE_FILTRO_TODOS, vbTextCompare) = 0 Then
        ResponsableFiltro_ValorAId = ""
        Exit Function
    End If

    ' Sentinel "0;N/A" (BC=2) o "0" (BC=1) -> ""
    If StrComp(s, RESPONSABLE_FILTRO_NA_ID & ";" & RESPONSABLE_FILTRO_NA_LABEL, _
                vbTextCompare) = 0 Then
        ResponsableFiltro_ValorAId = ""
        Exit Function
    End If
    If StrComp(s, RESPONSABLE_FILTRO_NA_ID, vbTextCompare) = 0 Then
        ResponsableFiltro_ValorAId = ""
        Exit Function
    End If

    ' Formato "<ID>;<nombre>" -> "<ID>"
    m_Pos = InStr(1, s, ";", vbBinaryCompare)
    If m_Pos > 0 Then
        ResponsableFiltro_ValorAId = Trim$(Left$(s, m_Pos - 1))
        Exit Function
    End If

    ' Formato "<ID>" puro.
    ResponsableFiltro_ValorAId = s
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ResponsableFiltro_ValorAId: " & Err.Number & _
                    " - " & Err.Description
    End If
    ResponsableFiltro_ValorAId = ""
End Function


' Construye la fila "id;nombre" (o "" si no matchea) para asignar al
' combo responsable.  Pensado para `setExpBusqueda` cuando queremos
' reflejar el ID guardado en el expediente (incluyendo IDs huerfanos).
'
'   p_ID            : ID del responsable guardado en el expediente.
'                     "" -> "" (no match).
'   p_Dict          : Scripting.Dictionary (key=CStr(Id), value=Usuario)
'                     cargado por el helper.  Nothing -> "" (no match).
'   p_ConSentinel   : True para Seguridad (prependeria "0;N/A" si el ID
'                     esta vacio y asi el sentinel queda visible en el
'                     combo).  Calidad usa False.
'
' Salidas:
'   * Si p_ID vacio y p_ConSentinel=True -> "0;N/A" (sentinel UX).
'   * Si p_ID vacio y p_ConSentinel=False -> "" (Calidad no admite "ninguno").
'   * Si p_Dict contiene CStr(p_ID) -> "<ID>;<Nombre>".
'   * Si p_Dict NO contiene CStr(p_ID) -> "" (no match; caller puede
'     tratar el ID como huerfano via el helper).
Public Function ResponsableFiltro_IdACmbRow( _
                                            ByVal p_ID As String, _
                                            ByVal p_Dict As Scripting.Dictionary, _
                                            ByVal p_ConSentinel As Boolean _
                                            ) As String
    Dim m_IDKey As String
    Dim m_Usuario As Usuario
    On Error GoTo errores

    m_IDKey = Trim$(p_ID)

    If Len(m_IDKey) = 0 Then
        If p_ConSentinel Then
            ResponsableFiltro_IdACmbRow = RESPONSABLE_FILTRO_NA_ID & ";" & _
                                            RESPONSABLE_FILTRO_NA_LABEL
        Else
            ResponsableFiltro_IdACmbRow = ""
        End If
        Exit Function
    End If

    If p_Dict Is Nothing Then
        ResponsableFiltro_IdACmbRow = ""
        Exit Function
    End If

    If Not p_Dict.Exists(m_IDKey) Then
        ResponsableFiltro_IdACmbRow = ""
        Exit Function
    End If

    Set m_Usuario = p_Dict(m_IDKey)
    If m_Usuario Is Nothing Then
        ResponsableFiltro_IdACmbRow = ""
        Exit Function
    End If
    ResponsableFiltro_IdACmbRow = CStr(m_Usuario.ID) & ";" & CStr(m_Usuario.Nombre)
    Set m_Usuario = Nothing
    Exit Function

errores:
    ResponsableFiltro_IdACmbRow = ""
End Function


' Compara el `p_Valor` del combo con un ID canonico de filtro.  Util para
' `HaHabidoCambiosBusqueda` y para detectar si el usuario toco el combo.
' Devuelve True si el valor del combo (en cualquier formato BC=1 / BC=2 /
' "Todos" / "0;N/A") representa `p_ID` canonico.
'
'   p_Valor : valor del combo (Variant -- .Value / .Column(0) / .Column(1)).
'   p_ID    : ID canonico ("119", "127", etc.) o "" (sin filtro).
Public Function ResponsableFiltro_ValorEsId( _
                                            ByVal p_Valor As Variant, _
                                            ByVal p_ID As String _
                                            ) As Boolean
    Dim m_IDExtraido As String
    Dim m_IDFiltro As String

    m_IDExtraido = ResponsableFiltro_ValorAId(p_Valor, m_Dummy)
    m_IDFiltro = Trim$(CStr(Nz(p_ID, "")))

    If Len(m_IDExtraido) = 0 And Len(m_IDFiltro) = 0 Then
        ResponsableFiltro_ValorEsId = True
        Exit Function
    End If
    ResponsableFiltro_ValorEsId = (StrComp(m_IDExtraido, m_IDFiltro, _
                                            vbTextCompare) = 0)
End Function