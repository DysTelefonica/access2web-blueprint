Attribute VB_Name = "TestParametrosParser"
Option Compare Database
'*******************************************************************************
' Módulo: TestParametrosParser
' Propósito: Pruebas unitarias para la clase ParametrosParser
' Autor: Sistema HPS
' Fecha: 2024
'*******************************************************************************

Option Explicit

'*******************************************************************************
' Subrutina: EjecutarTodasLasPruebas
' Propósito: Ejecuta todas las pruebas de la clase ParametrosParser
'*******************************************************************************
Public Sub EjecutarTodasLasPruebas()
    
    Debug.Print "=== INICIANDO PRUEBAS DE ParametrosParser ==="
    Debug.Print ""
    
    ' Pruebas de parseo
    Call PruebaFormatoActual
    Call PruebaFormatoSimplificado
    Call PruebaFormatoMixto
    Call PruebaCadenaVacia
    Call PruebaCadenaInvalida
    Call PruebaValoresConDosPuntos
    Call PruebaEspaciosEnBlanco
    
    ' Pruebas de generación
    Call PruebaGenerarCadenaFormatoActual
    Call PruebaGenerarCadenaFormatoNuevo
    
    ' Pruebas de validación
    Call PruebaValidacion
    
    Debug.Print ""
    Debug.Print "=== PRUEBAS COMPLETADAS ==="
    
End Sub

'*******************************************************************************
' Subrutina: PruebaFormatoActual
' Propósito: Prueba el parseo del formato actual ||Campo||:Valor|||
'*******************************************************************************
Private Sub PruebaFormatoActual()
    
    Debug.Print "--- Prueba: Formato Actual ---"
    
    Dim parser As New ParametrosParser
    Dim cadena As String
    Dim resultado As Scripting.Dictionary
    
    ' Formato esperado actual
    cadena = "||BuzonSeguridad||:te_seguridad_industrial@telefonica.com|||mailto||:martina.torralbarodriguez@telefonica.com|||Gestor||:Juan Pérez"
    
    Set resultado = parser.ParsearCadenaRecursos(cadena)
    
    Debug.Print "Cadena: " & cadena
    Debug.Print "Elementos parseados: " & resultado.Count
    Debug.Print "BuzonSeguridad: " & resultado("BuzonSeguridad")
    Debug.Print "mailto: " & resultado("mailto")
    Debug.Print "Gestor: " & resultado("Gestor")
    Debug.Print "Error: " & parser.Error
    Debug.Print ""
    
End Sub

'*******************************************************************************
' Subrutina: PruebaFormatoSimplificado
' Propósito: Prueba el parseo del formato simplificado Campo:Valor|
'*******************************************************************************
Private Sub PruebaFormatoSimplificado()
    
    Debug.Print "--- Prueba: Formato Simplificado ---"
    
    Dim parser As New ParametrosParser
    Dim cadena As String
    Dim resultado As Scripting.Dictionary
    
    ' Formato simplificado
    cadena = "BuzonSeguridad:te_seguridad_industrial@telefonica.com|mailto:martina.torralbarodriguez@telefonica.com|Gestor:Juan Pérez"
    
    Set resultado = parser.ParsearCadenaRecursos(cadena)
    
    Debug.Print "Cadena: " & cadena
    Debug.Print "Elementos parseados: " & resultado.Count
    If resultado.Exists("BuzonSeguridad") Then Debug.Print "BuzonSeguridad: " & resultado("BuzonSeguridad")
    If resultado.Exists("mailto") Then Debug.Print "mailto: " & resultado("mailto")
    If resultado.Exists("Gestor") Then Debug.Print "Gestor: " & resultado("Gestor")
    Debug.Print "Error: " & parser.Error
    Debug.Print ""
    
End Sub

'*******************************************************************************
' Subrutina: PruebaFormatoMixto
' Propósito: Prueba el parseo de formatos mixtos
'*******************************************************************************
Private Sub PruebaFormatoMixto()
    
    Debug.Print "--- Prueba: Formato Mixto ---"
    
    Dim parser As New ParametrosParser
    Dim cadena As String
    Dim resultado As Scripting.Dictionary
    
    ' Formato mixto (actual generado por Solicitud.cls)
    cadena = "BuzonSeguridad:te_seguridad_industrial@telefonica.com|mailto:martina.torralbarodriguez@telefonica.com"
    
    Set resultado = parser.ParsearCadenaRecursos(cadena)
    
    Debug.Print "Cadena: " & cadena
    Debug.Print "Elementos parseados: " & resultado.Count
    If resultado.Exists("BuzonSeguridad") Then Debug.Print "BuzonSeguridad: " & resultado("BuzonSeguridad")
    If resultado.Exists("mailto") Then Debug.Print "mailto: " & resultado("mailto")
    Debug.Print "Error: " & parser.Error
    Debug.Print ""
    
End Sub

'*******************************************************************************
' Subrutina: PruebaCadenaVacia
' Propósito: Prueba el manejo de cadenas vacías
'*******************************************************************************
Private Sub PruebaCadenaVacia()
    
    Debug.Print "--- Prueba: Cadena Vacía ---"
    
    Dim parser As New ParametrosParser
    Dim resultado As Scripting.Dictionary
    
    Set resultado = parser.ParsearCadenaRecursos("")
    
    Debug.Print "Cadena: (vacía)"
    Debug.Print "Elementos parseados: " & resultado.Count
    Debug.Print "Error: " & parser.Error
    Debug.Print ""
    
End Sub

'*******************************************************************************
' Subrutina: PruebaCadenaInvalida
' Propósito: Prueba el manejo de cadenas con formato inválido
'*******************************************************************************
Private Sub PruebaCadenaInvalida()
    
    Debug.Print "--- Prueba: Cadena Inválida ---"
    
    Dim parser As New ParametrosParser
    Dim cadena As String
    Dim resultado As Scripting.Dictionary
    
    cadena = "esto no tiene formato válido"
    
    Set resultado = parser.ParsearCadenaRecursos(cadena)
    
    Debug.Print "Cadena: " & cadena
    Debug.Print "Elementos parseados: " & resultado.Count
    Debug.Print "Error: " & parser.Error
    Debug.Print ""
    
End Sub

'*******************************************************************************
' Subrutina: PruebaValoresConDosPuntos
' Propósito: Prueba valores que contienen dos puntos
'*******************************************************************************
Private Sub PruebaValoresConDosPuntos()
    
    Debug.Print "--- Prueba: Valores con Dos Puntos ---"
    
    Dim parser As New ParametrosParser
    Dim cadena As String
    Dim resultado As Scripting.Dictionary
    
    cadena = "URL:http://www.ejemplo.com:8080|Hora:14:30:00"
    
    Set resultado = parser.ParsearCadenaRecursos(cadena)
    
    Debug.Print "Cadena: " & cadena
    Debug.Print "Elementos parseados: " & resultado.Count
    If resultado.Exists("URL") Then Debug.Print "URL: " & resultado("URL")
    If resultado.Exists("Hora") Then Debug.Print "Hora: " & resultado("Hora")
    Debug.Print "Error: " & parser.Error
    Debug.Print ""
    
End Sub

'*******************************************************************************
' Subrutina: PruebaEspaciosEnBlanco
' Propósito: Prueba el manejo de espacios en blanco
'*******************************************************************************
Private Sub PruebaEspaciosEnBlanco()
    
    Debug.Print "--- Prueba: Espacios en Blanco ---"
    
    Dim parser As New ParametrosParser
    Dim cadena As String
    Dim resultado As Scripting.Dictionary
    
    cadena = " Campo1 : Valor con espacios | Campo2: Otro valor |  "
    
    Set resultado = parser.ParsearCadenaRecursos(cadena)
    
    Debug.Print "Cadena: " & cadena
    Debug.Print "Elementos parseados: " & resultado.Count
    If resultado.Exists("Campo1") Then Debug.Print "Campo1: [" & resultado("Campo1") & "]"
    If resultado.Exists("Campo2") Then Debug.Print "Campo2: [" & resultado("Campo2") & "]"
    Debug.Print "Error: " & parser.Error
    Debug.Print ""
    
End Sub

'*******************************************************************************
' Subrutina: PruebaGenerarCadenaFormatoActual
' Propósito: Prueba la generación de cadenas en formato actual
'*******************************************************************************
Private Sub PruebaGenerarCadenaFormatoActual()
    
    Debug.Print "--- Prueba: Generar Formato Actual ---"
    
    Dim parser As New ParametrosParser
    Dim diccionario As New Scripting.Dictionary
    Dim resultado As String
    
    diccionario.Add "BuzonSeguridad", "te_seguridad_industrial@telefonica.com"
    diccionario.Add "mailto", "martina.torralbarodriguez@telefonica.com"
    diccionario.Add "Gestor", "Juan Pérez"
    
    resultado = parser.GenerarCadenaRecursos(diccionario)
    
    Debug.Print "Diccionario con " & diccionario.Count & " elementos"
    Debug.Print "Resultado: " & resultado
    Debug.Print "Error: " & parser.Error
    Debug.Print ""
    
End Sub

'*******************************************************************************
' Subrutina: PruebaGenerarCadenaFormatoNuevo
' Propósito: Prueba la generación de cadenas en formato nuevo
'*******************************************************************************
Private Sub PruebaGenerarCadenaFormatoNuevo()
    
    Debug.Print "--- Prueba: Generar Formato Nuevo ---"
    
    Dim parser As New ParametrosParser
    Dim diccionario As New Scripting.Dictionary
    Dim resultado As String
    
    diccionario.Add "BuzonSeguridad", "te_seguridad_industrial@telefonica.com"
    diccionario.Add "mailto", "martina.torralbarodriguez@telefonica.com"
    diccionario.Add "Gestor", "Juan Pérez"
    
    resultado = parser.GenerarCadenaRecursos(diccionario)
    
    Debug.Print "Diccionario con " & diccionario.Count & " elementos"
    Debug.Print "Resultado: " & resultado
    Debug.Print "Error: " & parser.Error
    Debug.Print ""
    
End Sub

'*******************************************************************************
' Subrutina: PruebaValidacion
' Propósito: Prueba la función de validación
'*******************************************************************************
Private Sub PruebaValidacion()
    
    Debug.Print "--- Prueba: Validación ---"
    
    Dim parser As New ParametrosParser
    
    Debug.Print "Cadena válida: " & parser.ValidarCadenaRecursos("Campo1:Valor1|Campo2:Valor2")
    Debug.Print "Cadena inválida: " & parser.ValidarCadenaRecursos("formato incorrecto")
    Debug.Print "Cadena vacía: " & parser.ValidarCadenaRecursos("")
    Debug.Print ""
    
End Sub

