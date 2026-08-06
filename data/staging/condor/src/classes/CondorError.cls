VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CondorError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO DE CLASE: CondorError
' RESPONSABILIDAD: Encapsular toda la información de un error para una
'                  propagación y registro robustos.
' ==========================================================================

Private m_Number As Long
Private m_Description As String
Private m_Source As String
Private m_CallStack As Collection

' Propiedad para el número de error original (ej: 3022)
Public Property Get Number() As Long
    Number = m_Number
End Property

' Propiedad para la descripción original del error
Public Property Get description() As String
    description = m_Description
End Property

' Propiedad para el origen del error (ej: "WorkflowServicio.EjecutarTransicion")
Public Property Get source() As String
    source = m_Source
End Property

' Colección que almacena la ruta que ha seguido el error
Public Property Get CallStack() As Collection
    Set CallStack = m_CallStack
End Property

' Método para inicializar el objeto cuando ocurre un error
Public Sub Create(ByVal errNumber As Long, ByVal errDescription As String, ByVal errSource As String)
    m_Number = errNumber
    m_Description = errDescription
    m_Source = errSource
    Set m_CallStack = New Collection
    Me.AddToCallStack errSource ' El primer elemento de la pila es el origen
End Sub

' Método para añadir un nuevo nivel a la pila de llamadas
Public Sub AddToCallStack(ByVal sourceProcedure As String)
    If m_CallStack Is Nothing Then Set m_CallStack = New Collection
    m_CallStack.Add sourceProcedure
End Sub

Public Sub Raise()
    ' 1. Guardamos el objeto de error completo en una variable global.
    Set g_objLastError = Me
    
    ' 2. Auto-log a tbLogErrores (wrapped — no puede fallar)
    Call ErrorLogger.LogToTable(Me)
    
    ' 3. Re-lanzamos el error ORIGINAL con toda la información.
    '    Preserve m_Number para que callers puedan distinguir errores (ej: 513 = validación).
    '    Description incluye la pila completa para debugging.
    Err.Raise m_Number, Me.source, Me.FullDescription
End Sub

' Convierte la pila de llamadas en string para persistencia
Public Function CallStackToString() As String
    Dim i As Long
    Dim result As String
    
    If Not m_CallStack Is Nothing Then
        For i = m_CallStack.count To 1 Step -1
            result = result & vbNewLine & "   en " & m_CallStack(i)
        Next i
    End If
    
    CallStackToString = Trim(result)
End Function
' Propiedad que devuelve la descripción completa, incluyendo la pila de llamadas
Public Property Get FullDescription() As String
    Dim i As Long
    Dim stack As String
    
    If Not m_CallStack Is Nothing Then
        For i = m_CallStack.count To 1 Step -1
            stack = stack & vbNewLine & "   en " & m_CallStack(i)
        Next i
    End If
    
    FullDescription = "Error #" & m_Number & ": " & m_Description & vbNewLine & "Pila de Llamadas:" & stack
End Property

