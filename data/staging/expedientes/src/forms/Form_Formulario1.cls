Option Compare Database
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
#Else
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
#End If



' Constantes para mostrar/ocultar Access
Const SW_HIDE = 0
Const SW_NORMAL = 1
Const SW_MINIMIZED = 2
Const SW_MAXIMIZED = 3

Private Sub Form_Open(Cancel As Integer)
    ' Inicializar Firebase al abrir el formulario
    Call InitializeFirebase
End Sub
Private Sub btnSend_Click()
    ' Enviar datos al nodo "notificaesto" en Firebase
    Dim message As String
    Dim m_Error As String
    message = Me.txtInput.value
    
    ' Ejecutar JavaScript en el WebBrowser para enviar datos
    Dim script As String
    script = modFormulario1Helper.Formulario1_BuildFirebaseSetScript(message, m_Error)
    If m_Error <> "" Then Err.Raise 1000, "btnSend_Click", m_Error
    
    Me.WebBrowser0.Object.Document.parentWindow.execScript script, "JavaScript"
End Sub

Private Sub InitializeFirebase()
    ' Crear una página HTML con la configuración de Firebase
    Dim m_Error As String
    Dim tempFilePath As String
    tempFilePath = modFormulario1Helper.Formulario1_BuildTempFilePath(Environ("TEMP"), m_Error)
    If m_Error <> "" Then Err.Raise 1000, "InitializeFirebase", m_Error
    
    Dim htmlContent As String
    htmlContent = modFormulario1Helper.Formulario1_BuildFirebaseHtml(m_Error)
    If m_Error <> "" Then Err.Raise 1000, "InitializeFirebase", m_Error
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open tempFilePath For Output As #fileNum
    Print #fileNum, htmlContent
    Close #fileNum
    
    ' Cargar la página HTML en el control WebBrowser
    Me.WebBrowser0.Navigate tempFilePath
End Sub

Public Sub UpdateRealtime(value As String)
    ' Actualizar el cuadro de texto con los datos recibidos
    Me.txtRealtime.value = value
End Sub

Private Sub txtInput_Change()
    ' Enviar datos al nodo "notificaesto" en Firebase
    Dim message As String
    Dim m_Error As String
    message = Me.txtInput.Text ' Usar .Text para obtener el contenido del cuadro de texto
    
    ' Ejecutar JavaScript en el WebBrowser para enviar datos
    Dim script As String
    script = modFormulario1Helper.Formulario1_BuildFirebaseSetScript(message, m_Error)
    If m_Error <> "" Then Err.Raise 1000, "txtInput_Change", m_Error
    
    ' Ejecutar el script en el WebBrowser
    Me.WebBrowser0.Object.Document.parentWindow.execScript script, "JavaScript"
End Sub
