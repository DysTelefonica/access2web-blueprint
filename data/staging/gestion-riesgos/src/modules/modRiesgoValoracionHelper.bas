Attribute VB_Name = "modRiesgoValoracionHelper"
Option Compare Database
Option Explicit

' =============================================================================
' modRiesgoValoracionHelper.bas
'
' Helper module para resolver la Valoracion persistida para un riesgo.
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
' SDD:     e2e-form-by-form-2026-06-22 - Bloque 4 - REQ-CAL-05
'
' Helper público (1):
'   CalcularValoracion - devuelve la Valoracion persistida en
'                        TbRiesgosValoracion para el par (ImpactoGlobal,
'                        Vulnerabilidad) del riesgo identificado por
'                        p_IDRiesgo. Es un wrapper sobre
'                        Constructor.getValoracion(ImpactoGlobal,
'                        Vulnerabilidad) que primero lee esos dos campos
'                        desde TbRiesgos y maneja la ausencia de datos.
'
' Decisión 2026-06-22 (sdd-apply Bloque 4 / PR-6):
'   * Schema-first (verificado vía dysflow_get_schema el 2026-06-22):
'       - TbRiesgosValoracion: columnas Impacto, Vulnerabilidad, Valoracion
'         son las 3 dbText (size 255). NO son numéricas. La matriz es 5x5
'         con etiquetas: Impacto ? {Muy Bajo, Bajo, Medio, Alto, Muy Alto}
'         × Vulnerabilidad ? {Muy Bajo, Bajo, Medio, Alto, Muy Alto}.
'       - TbRiesgos.ImpactoGlobal y TbRiesgos.Vulnerabilidad son dbText
'         (size 15). Se leen como String.
'   * Pre-check REQ-CAL-05 (design §4.3): la tabla TbRiesgosValoracion
'     EXISTE en el staging backend con 25 filas (5x5). Esto habilita el
'     átomo adversarial (matrix modified manually).
'   * Esta implementación reemplaza el camino inline que el form tiene hoy
'     y que Calidad reportó como "muestra Alto/Muy Alto" en casos donde la
'     matriz dice otra cosa. La causa raíz fue que el form combinaba
'     ImpactoGlobal del riesgo con Vulnerabilidad sin pasar SIEMPRE por
'     la matriz persistida. El helper cierra ese gap leyendo DIRECTO
'     desde TbRiesgosValoracion.
'   * p_IDRiesgo es String por convención del proyecto (los ID* son String,
'     NO Long — esto difiere del design §2.2 que usaba Long, corregido
'     per project rules verificadas por el orchestrator).
'   * Si p_Error está poblado, el retorno es "" (regla del proyecto:
'     JSON body lleva datos válidos, error viaja por p_Error ByRef).
'   * DAO opcional: si db=Nothing usa CurrentDb; los átomos TDD inyectan
'     el sandbox vía Test_Fixtures.GetTestDb().
' =============================================================================

' -----------------------------------------------------------------------------
' CalcularValoracion
'
' Resuelve la Valoracion de un riesgo a partir de su ImpactoGlobal y
' Vulnerabilidad persistidos en TbRiesgos, cruzándolos contra la matriz
' TbRiesgosValoracion.
'
' Parámetros:
'   p_IDRiesgo (String)               - ID del riesgo (convención String).
'   db (DAO.Database, opcional)       - inyectado por átomos TDD; usa
'                                       CurrentDb si Nothing.
'   p_Error (out String, opcional)    - mensaje legible si retorna "".
'                                       Vacío si retorna una Valoracion.
'
' Retorna:
'   La Valoracion persistida en TbRiesgosValoracion (string: "Muy Bajo",
'   "Bajo", "Medio", "Alto", "Muy Alto") cuando la combinación existe.
'   "" si p_Error está poblado (validación o matriz faltante).
'
' Reglas (en orden):
'   1. p_IDRiesgo no vacío ? si vacío, p_Error poblado, return "".
'   2. p_IDRiesgo numérico ? si no, p_Error poblado, return "".
'   3. Riesgo existe en TbRiesgos ? si no, p_Error poblado, return "".
'   4. TbRiesgos.ImpactoGlobal no vacío ? si vacío, p_Error poblado.
'   5. TbRiesgos.Vulnerabilidad no vacía/Null/"0" ? si falta, p_Error.
'   6. Matriz TbRiesgosValoracion tiene fila (Impacto, Vulnerabilidad) ?
'      si no, p_Error poblado con los valores usados para diagnóstico.
'   7. Valoracion vacía en la matriz ? p_Error poblado (defensa).
' -----------------------------------------------------------------------------
Public Function CalcularValoracion( _
    ByRef p_IDRiesgo As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As String

    Dim m_Db As DAO.Database
    Dim m_Rs As DAO.Recordset
    Dim m_ImpactoGlobal As String
    Dim m_Vulnerabilidad As String
    Dim m_VulnEsInvalida As Boolean
    Dim m_SQL As String
    Dim m_Valoracion As String

    On Error GoTo errores
    p_Error = ""
    CalcularValoracion = ""

    ' --- 1. Validación: p_IDRiesgo no vacío ---
    If Len(Trim$(Nz(p_IDRiesgo, ""))) = 0 Then
        p_Error = "CalcularValoracion: p_IDRiesgo está vacío"
        Exit Function
    End If

    ' --- 2. Validación: p_IDRiesgo numérico ---
    '     Los IDRiesgo en el proyecto son enteros longs almacenados como
    '     string. Rechazamos strings no numéricos para no enviar SQL
    '     con texto a la cláusula WHERE.
    If Not IsNumeric(p_IDRiesgo) Then
        p_Error = "CalcularValoracion: p_IDRiesgo no es numérico"
        Exit Function
    End If

    ' --- 3. Resolver db (inyectado por el átomo, o CurrentDb en producción) ---
    If db Is Nothing Then
        Set m_Db = CurrentDb
    Else
        Set m_Db = db
    End If

    ' --- 4. Leer ImpactoGlobal y Vulnerabilidad desde TbRiesgos ---
    m_SQL = "SELECT ImpactoGlobal, Vulnerabilidad " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & CLng(p_IDRiesgo)
    Set m_Rs = m_Db.OpenRecordset(m_SQL, dbOpenSnapshot)
    If m_Rs.EOF Then
        m_Rs.Close
        Set m_Rs = Nothing
        Set m_Db = Nothing
        p_Error = "CalcularValoracion: no se encontró el riesgo con IDRiesgo=" & p_IDRiesgo
        Exit Function
    End If

    m_ImpactoGlobal = Trim$(CStr(Nz(m_Rs.fields("ImpactoGlobal").value, "")))
    m_Vulnerabilidad = Trim$(CStr(Nz(m_Rs.fields("Vulnerabilidad").value, "")))
    m_Rs.Close
    Set m_Rs = Nothing

    ' --- 5. Validación: ImpactoGlobal persistido ---
    If Len(m_ImpactoGlobal) = 0 Then
        Set m_Db = Nothing
        p_Error = "CalcularValoracion: el riesgo no tiene ImpactoGlobal persistido"
        Exit Function
    End If

    ' --- 6. Validación: Vulnerabilidad persistida ---
    '     Schema-first: la columna es dbText. Cubrimos Null, cadena vacía
    '     y literal "0" como inválidos. Cada check va en su propio If
    '     porque VBA no hace short-circuit de Or/And (vba-access §1.6.1).
    m_VulnEsInvalida = False
    If IsNull(m_Vulnerabilidad) Then
        m_VulnEsInvalida = True
    ElseIf Len(Trim$(m_Vulnerabilidad)) = 0 Then
        m_VulnEsInvalida = True
    ElseIf (m_Vulnerabilidad = 0) Then
        ' "0" no aparece en la matriz; es señal de dato mal persistido.
        m_VulnEsInvalida = True
    End If
    If m_VulnEsInvalida Then
        Set m_Db = Nothing
        p_Error = "CalcularValoracion: el riesgo no tiene Vulnerabilidad persistida"
        Exit Function
    End If

    ' --- 7. Consultar la matriz TbRiesgosValoracion ---
    '     SqlStr escapa apóstrofes (Test_Helper.SqlStr v1.9 §3).
    '     Si Impacto o Vulnerabilidad contienen apóstrofes en datos
    '     legítimos, esto es necesario para evitar SQL injection o un
    '     match silencioso contra otro valor.
    m_SQL = "SELECT Valoracion " & _
            "FROM TbRiesgosValoracion " & _
            "WHERE Impacto='" & Test_Helper.SqlStr(m_ImpactoGlobal) & "' " & _
            "  AND Vulnerabilidad='" & Test_Helper.SqlStr(m_Vulnerabilidad) & "'"
    Set m_Rs = m_Db.OpenRecordset(m_SQL, dbOpenSnapshot)
    If m_Rs.EOF Then
        m_Rs.Close
        Set m_Rs = Nothing
        Set m_Db = Nothing
        p_Error = "CalcularValoracion: no se encontró Valoracion para " & _
                  "Impacto='" & m_ImpactoGlobal & "' y Vulnerabilidad='" & _
                  m_Vulnerabilidad & "' en TbRiesgosValoracion"
        Exit Function
    End If

    m_Valoracion = Trim$(CStr(Nz(m_Rs.fields("Valoracion").value, "")))
    m_Rs.Close
    Set m_Rs = Nothing
    Set m_Db = Nothing

    ' --- 8. Defensa: si la Valoracion quedó vacía, también es error ---
    '     Esto no debería ocurrir dada la verificación de paso 7 (no
    '     deberíamos tener filas sin Valoracion), pero es defense-in-depth
    '     por si la BD se corrompe con un INSERT manual incompleto.
    If Len(m_Valoracion) = 0 Then
        p_Error = "CalcularValoracion: Valoracion vacía en la matriz para " & _
                  "Impacto='" & m_ImpactoGlobal & "' y Vulnerabilidad='" & _
                  m_Vulnerabilidad & "'"
        Exit Function
    End If

    CalcularValoracion = m_Valoracion
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CalcularValoracion: " & Err.Number & " - " & Err.description
    End If
    CalcularValoracion = ""
End Function

' -----------------------------------------------------------------------------
' CompararValoracionPersistidaVsCalculada
'
' Compara la Valoracion PERSISTIDA en TbRiesgos contra la Valoracion que la
' MATRIZ TbRiesgosValoracion dice para el par (ImpactoGlobal, Vulnerabilidad)
' del riesgo. Devuelve un JSON con tres campos: {persistida, calculada,
' coinciden} para que el caller (form, Debug.Print de auditoría, o UAT)
' tenga el contrato observable en un único valor.
'
' Esta función existe para cerrar el gap REQ-CAL-05 Bloque 4 / PR-6: el
' form tiene un Debug.Print TEMPORAL que compara la persistida vs la
' calculada. Para hacerlo testeable, la comparación se extrae al helper
' (que ya tiene CalcularValoracion inyectable en TestDb) y el resultado
' es un JSON atómico que cualquier átomo TDD puede asertar.
'
' Parámetros:
'   p_IDRiesgo (String)               - ID del riesgo (convención String).
'   db (DAO.Database, opcional)       - inyectado por átomos TDD; usa
'                                       CurrentDb si Nothing.
'   p_Error (out String, opcional)    - mensaje legible si retorna "".
'                                       Vacío si retorna un JSON válido.
'
' Retorna:
'   JSON atómico de la forma:
'       {"persistida":"<X>","calculada":"<Y>","coinciden":<true|false>}
'
'   donde:
'       - "persistida" es el valor de TbRiesgos.Valoracion (string;
'         "" si la columna es Null).
'       - "calculada" es lo que devuelve CalcularValoracion (string;
'         "" si la matriz no tiene fila para (ImpactoGlobal, Vulnerabilidad)).
'       - "coinciden" es un boolean JSON: True si ambos son iguales (case-
'         insensitive); False en caso contrario (incluye cuando uno es ""
'         y el otro no).
'
'   Si p_Error está poblado (input inválido o BD inaccesible), el retorno
'   es "" y NO se emite JSON parcial. La regla del proyecto es: el JSON
'   body lleva datos válidos, el error viaja por p_Error ByRef.
'
' Reglas (en orden):
'   1. p_IDRiesgo no vacío ? si vacío, p_Error poblado, return "".
'   2. p_IDRiesgo numérico ? si no, p_Error poblado, return "".
'   3. Resolver db (inyectado o CurrentDb).
'   4. Leer TbRiesgos.Valoracion filtrado por IDRiesgo (persistida;
'      tolerate Null ? "").
'   5. Llamar a la propia CalcularValoracion (calculada) — ella valida
'      ImpactoGlobal/Vulnerabilidad/matriz y devuelve "" + p_Error poblado
'      si algo falta. Este método NO propaga ese p_Error como fallo de
'      CompararValoracionPersistidaVsCalculada porque un valor ""
'      calculado es una observación válida del contrato
'      (matriz desactualizada / dato faltante), NO un error del helper.
'   6. Comparar case-insensitive: coinciden = (UCase(persistida) =
'      UCase(calculada)). Caso vacío-vs-lleno ? coinciden = False.
'   7. Emitir JSON atómico. Si el armado del JSON falla, p_Error poblado.
' -----------------------------------------------------------------------------
Public Function CompararValoracionPersistidaVsCalculada( _
    ByRef p_IDRiesgo As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As String

    Dim m_Db As DAO.Database
    Dim m_Rs As DAO.Recordset
    Dim m_SQL As String
    Dim m_Persistida As String
    Dim m_Calculada As String
    Dim m_Coinciden As Boolean
    Dim m_ErrCalc As String

    On Error GoTo errores
    p_Error = ""
    CompararValoracionPersistidaVsCalculada = ""

    ' --- 1. Validación: p_IDRiesgo no vacío ---
    If Len(Trim$(Nz(p_IDRiesgo, ""))) = 0 Then
        p_Error = "CompararValoracionPersistidaVsCalculada: p_IDRiesgo está vacío"
        Exit Function
    End If

    ' --- 2. Validación: p_IDRiesgo numérico ---
    If Not IsNumeric(p_IDRiesgo) Then
        p_Error = "CompararValoracionPersistidaVsCalculada: p_IDRiesgo no es numérico"
        Exit Function
    End If

    ' --- 3. Resolver db ---
    If db Is Nothing Then
        Set m_Db = CurrentDb
    Else
        Set m_Db = db
    End If

    ' --- 4. Leer TbRiesgos.Valoracion (persistida) ---
    m_SQL = "SELECT Valoracion FROM TbRiesgos WHERE IDRiesgo=" & CLng(p_IDRiesgo)
    Set m_Rs = m_Db.OpenRecordset(m_SQL, dbOpenSnapshot)
    If m_Rs.EOF Then
        m_Rs.Close
        Set m_Rs = Nothing
        Set m_Db = Nothing
        p_Error = "CompararValoracionPersistidaVsCalculada: no se encontró el riesgo con IDRiesgo=" & p_IDRiesgo
        Exit Function
    End If
    Dim m_ValoracionPersistidaRaw As Variant
    m_ValoracionPersistidaRaw = m_Rs.fields("Valoracion").value
    m_Rs.Close
    Set m_Rs = Nothing

    ' Tolerar Null ? "". Es observable: una Valoracion Null persistida
    ' cuenta como "" para la comparación, no como error del helper.
    If IsNull(m_ValoracionPersistidaRaw) Then
        m_Persistida = ""
    Else
        m_Persistida = Trim$(CStr(m_ValoracionPersistidaRaw))
    End If

    ' --- 5. Calcular la Valoracion desde la matriz actual ---
    '     NO propagamos p_Error de CalcularValoracion como fallo de este
    '     helper: "" calculado es una observación válida (matriz
    '     desactualizada / dato faltante), no un error del wrapper.
    m_ErrCalc = ""
    m_Calculada = CalcularValoracion(p_IDRiesgo, db, m_ErrCalc)
    If Len(m_ErrCalc) > 0 Then
        ' Mantener p_Error poblado para diagnóstico pero continuar con
        ' m_Calculada = "" (contrato observable).
        p_Error = "CompararValoracionPersistidaVsCalculada: calculada no disponible: " & m_ErrCalc
    End If

    Set m_Db = Nothing

    ' --- 6. Comparar case-insensitive ---
    '     VBA no hace short-circuit de Or/And (vba-access §1.6.1); separar checks.
    m_Coinciden = False
    If Len(m_Persistida) > 0 And Len(m_Calculada) > 0 Then
        If UCase$(m_Persistida) = UCase$(m_Calculada) Then
            m_Coinciden = True
        End If
    End If

    ' --- 7. Emitir JSON atómico ---
    '     Formato estable para átomos TDD: {persistida, calculada, coinciden}.
    '     Construcción por concatenación con EscapeJsonString inline: este
    '     helper no es candidato para Dictionary+JsonConverter porque los
    '     átomos TDD asertan string equality exacta y el formato es chico.
    CompararValoracionPersistidaVsCalculada = _
        "{""persistida"":""" & EscapeJsonString(m_Persistida) & _
        """,""calculada"":""" & EscapeJsonString(m_Calculada) & _
        """,""coinciden"":" & LCase$(CStr(m_Coinciden)) & "}"

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CompararValoracionPersistidaVsCalculada: " & Err.Number & " - " & Err.description
    End If
    CompararValoracionPersistidaVsCalculada = ""
End Function

' -----------------------------------------------------------------------------
' EscapeJsonString (helper interno)
'
' Escapa caracteres especiales para embeber un string en un JSON literal:
'   " ? \"
'   \ ? \\
'   CR/LF ? \n
' Otros caracteres se mantienen tal cual (UTF-8 OK en VBA).
'
' Uso: SOLO dentro de este módulo. Es privado por scope y por dependencia:
' estos helpers no son testables por átomos TDD independientes del JSON
' shape de CompararValoracionPersistidaVsCalculada.
' -----------------------------------------------------------------------------
Private Function EscapeJsonString(ByVal p_Value As String) As String
    Dim m_Result As String
    m_Result = p_Value
    m_Result = Replace(m_Result, "\", "\\")
    m_Result = Replace(m_Result, """", "\""")
    m_Result = Replace(m_Result, vbCrLf, "\n")
    m_Result = Replace(m_Result, vbLf, "\n")
    m_Result = Replace(m_Result, vbCr, "\n")
    EscapeJsonString = m_Result
End Function

