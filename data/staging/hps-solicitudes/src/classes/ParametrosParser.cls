VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParametrosParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'*******************************************************************************
' Clase: ParametrosParser
' Propósito: Parsear cadenas de parámetros clave-valor de manera robusta
' Autor: Sistema HPS
' Fecha: 2024
'*******************************************************************************



Private m_Error As String

' Propiedad para obtener errores
Public Property Get Error() As String
    Error = m_Error
End Property

Private Property Let Error(ByVal valor As String)
    m_Error = valor
End Property

'*******************************************************************************
' Función: ParsearCadenaRecursos
' Propósito: Parsea una cadena de recursos y devuelve un diccionario
' Parámetros:
'   - cadenaRecursos: Cadena a parsear
' Retorna: Scripting.Dictionary con las claves-valor parseadas
' Formatos soportados:
'   - ||Campo||:Valor||| (formato actual esperado)
'   - ||Campo||:Valor| (formato actual generado)
'   - Campo:Valor| (formato simplificado)
'*******************************************************************************
Public Function ParsearCadenaRecursos(ByVal cadenaRecursos As String) As Scripting.Dictionary
    
    Dim diccionario As New Scripting.Dictionary
    Dim elementos() As String
    Dim elemento As String
    Dim partes() As String
    Dim clave As String
    Dim valor As String
    Dim i As Integer
    
    On Error GoTo errores
    
    m_Error = ""
    
    ' Validar entrada
    If Trim(cadenaRecursos) = "" Then
        Set ParsearCadenaRecursos = diccionario
        Exit Function
    End If
    
    ' Normalizar separadores - convertir ||| a | para unificar formatos
    cadenaRecursos = Replace(cadenaRecursos, "|||", "|")
    
    ' Dividir por el separador principal
    elementos = Split(cadenaRecursos, "|")
    
    For i = 0 To UBound(elementos)
        elemento = Trim(elementos(i))
        
        ' Saltar elementos vacíos
        If elemento <> "" Then
            ' Buscar el separador clave:valor
            If InStr(elemento, ":") > 0 Then
                partes = Split(elemento, ":", 2) ' Limitar a 2 partes por si el valor contiene ":"
                
                If UBound(partes) >= 1 Then
                    clave = Trim(partes(0))
                    valor = Trim(partes(1))
                    
                    ' Limpiar las barras verticales de la clave si existen
                    clave = Replace(clave, "||", "")
                    
                    ' Agregar al diccionario (sobrescribir si ya existe)
                    If clave <> "" Then
                        diccionario(clave) = valor
                    End If
                End If
            End If
        End If
    Next i
    
    Set ParsearCadenaRecursos = diccionario
    Exit Function
    
errores:
    m_Error = "Error en ParsearCadenaRecursos: " & Err.Description
    Set ParsearCadenaRecursos = diccionario
End Function

'*******************************************************************************
' Función: GenerarCadenaRecursos
' Propósito: Genera una cadena de recursos desde un diccionario
' Parámetros:
'   - diccionario: Scripting.Dictionary con las claves-valor
'   - formato: Formato de salida ("ACTUAL" o "NUEVO")
' Retorna: String con la cadena generada
'*******************************************************************************
'Public Function GenerarCadenaRecursos(ByVal diccionario As Scripting.Dictionary, _
'                                     Optional ByVal formato As String = "ACTUAL") As String
'
'    Dim resultado As String
'    Dim clave As Variant
'    Dim valor As String
'    Dim separadorElementos As String
'    Dim formatoClave As String
'
'    On Error GoTo errores
'
'    m_Error = ""
'
'    ' Configurar formato según parámetro
'    Select Case UCase(formato)
'        Case "ACTUAL"
'            separadorElementos = "|||"
'            formatoClave = "||{CLAVE}||"
'        Case "NUEVO"
'            separadorElementos = "|"
'            formatoClave = "{CLAVE}"
'        Case Else
'            separadorElementos = "|||"
'            formatoClave = "||{CLAVE}||"
'    End Select
'
'    ' Generar cadena
'    For Each clave In diccionario.Keys
'        valor = diccionario(clave)
'
'        If resultado <> "" Then
'            resultado = resultado & separadorElementos
'        End If
'
'        resultado = resultado & Replace(formatoClave, "{CLAVE}", CStr(clave)) & ":" & valor
'    Next clave
'
'    GenerarCadenaRecursos = resultado
'    Exit Function
'
'errores:
'    m_Error = "Error en GenerarCadenaRecursos: " & Err.Description
'    GenerarCadenaRecursos = ""
'End Function
Public Function GenerarCadenaRecursos(ByVal diccionario As Scripting.Dictionary) As String
    
    Dim resultado As String
    Dim clave As Variant
    Dim valor As String
    
    On Error GoTo errores
    
    m_Error = ""
    resultado = ""
    
    ' Recorremos el diccionario
    For Each clave In diccionario.Keys
        valor = diccionario(clave)
        
        ' Construimos el formato: {||Clave||:Valor}
        ' NOTA: En Solicitud.cls las claves ya se añaden con las barras (ej: .Add "||Gestor||", ...)
        ' por lo que aquí solo concatenamos las llaves y los dos puntos.
        
        resultado = resultado & "{" & clave & ":" & valor & "}"
        
    Next clave
    
    GenerarCadenaRecursos = resultado
    Exit Function
    
errores:
    m_Error = "Error en GenerarCadenaRecursos: " & Err.Description
    GenerarCadenaRecursos = ""
End Function

'*******************************************************************************
' Función: ValidarCadenaRecursos
' Propósito: Valida si una cadena de recursos tiene formato correcto
' Parámetros:
'   - cadenaRecursos: Cadena a validar
' Retorna: Boolean indicando si es válida
'*******************************************************************************
Public Function ValidarCadenaRecursos(ByVal cadenaRecursos As String) As Boolean
    
    Dim diccionario As Scripting.Dictionary
    
    On Error GoTo errores
    
    m_Error = ""
    
    ' Intentar parsear
    Set diccionario = Me.ParsearCadenaRecursos(cadenaRecursos)
    
    ' Si no hay error y el diccionario tiene elementos, es válida
    ValidarCadenaRecursos = (m_Error = "" And diccionario.Count > 0)
    
    Exit Function
    
errores:
    m_Error = "Error en ValidarCadenaRecursos: " & Err.Description
    ValidarCadenaRecursos = False
End Function

