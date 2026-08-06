
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_frmSplash.cls (CORREGIDO - Reseteo del contador de progreso)
' ==========================================================================

Public ProgresoActual As Long

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Errores
    ' --- INICIO DE LA LÓGICA DE OCULTACIÓN DEL RIBBON ---
    ' Ocultamos la cinta de opciones aquí, que es el primer punto seguro
    ' donde sabemos que el entorno VBA está completamente cargado.
    ' Comprobamos que no estemos en modo DEV (bypass con Shift)
    If InStr(1, Nz(Command(), ""), "/dev", vbTextCompare) = 0 Then
        Call GestionarRibbon(False)
    End If
    ' --- FIN DE LA LÓGICA DE OCULTACIÓN ---
    Ajustar Me
    
    ' --- INICIO DE LA CORRECCIÓN ---
    ' Llamamos a la nueva función para reiniciar el contador estático.
    ' Esto asegura que la barra de progreso siempre comience desde 0.
    Call ResetAvanceCounter
    ' --- FIN DE LA CORRECCIÓN ---
    
    ' Aseguramos que la barra empieza visualmente vacía.
    Me.lblProgresoBarra.Width = 0
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("Form_Open", Me)
    Resume LimpiarYSalir
End Sub

Private Sub Form_Timer()
    Const MIN_DISPLAY_TIME As Single = 2 ' Reducimos un poco para agilizar el arranque
    Dim t1 As Single, t2 As Single, elapsed As Single
    
    On Error GoTo Errores
    
    Me.TimerInterval = 0
    t1 = Timer
    
    Call EVE
    
    t2 = Timer
    elapsed = t2 - t1
    
    If elapsed < MIN_DISPLAY_TIME Then
        Sleep (MIN_DISPLAY_TIME - elapsed) * 1000
    End If
    
    Select Case rolUsuario
        Case rol.Administrador, rol.Calidad
            DoCmd.OpenForm "frm0Ppal"
        Case rol.Tecnico
            DoCmd.OpenForm "frm0PpalTecnico"
        Case Else
            Err.Raise 513, , "Rol de usuario desconocido. La aplicación se cerrará."
    End Select
    
    DoCmd.Close acForm, Me.name
    Exit Sub
    
Errores:
    Me.Visible = False
    ' Utilizamos nuestro manejador estandarizado incluso para errores de arranque.
    ' Creamos un objeto CondorError para pasar la información completa.
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "frmSplash.Form_Timer (Arranque Crítico)"
    MsgBox "Se ha producido un error crítico durante el arranque:" & vbNewLine & errObj.FullDescription, vbCritical, "Error de Inicialización"
    Call CerrarConexionesGlobales ' <-- AÑADIR AQUÍ
   ' Application.Quit
End Sub