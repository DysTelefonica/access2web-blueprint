Attribute VB_Name = "modDevTools"
' ==========================================================================
' MÓDULO: modDevTools.bas (VERSIÓN CORREGIDA CON ZORDER)
' RESPONSABILIDAD: Contiene funciones de ayuda para el desarrollador.
' ==========================================================================
Option Compare Database
Option Explicit

Public Sub AjustarPanelesBloqueo(ByVal traerAlFrente As Boolean)
    ' RESPONSABILIDAD: Recorre todos los formularios del proyecto. Si encuentra
    '                  un control llamado "lblBloqueo", lo trae al frente o lo
    '                  envía al fondo EN MODO DISEÑO, y guarda el formulario.
    
    Dim db As DAO.Database
    Dim doc As DAO.Document
    Dim frm As Form
    Dim ctl As Control
    Dim modo As String
    Dim contador As Long
    
    On Error GoTo Errores
    
    If traerAlFrente Then
        modo = "TRAYENDO AL FRENTE"
    Else
        modo = "ENVIANDO AL FONDO"
    End If
    
    Debug.Print "======================================================"
    Debug.Print "INICIANDO PROCESO: " & modo & " PANELES DE BLOQUEO"
    Debug.Print "======================================================"
    
    Application.Echo False ' Desactiva el refresco de pantalla para acelerar el proceso
    
    Set db = CurrentDb()
    contador = 0
    
    ' Iteramos sobre todos los formularios del proyecto
    For Each doc In db.Containers("Forms").Documents
        On Error GoTo Errores
        
        ' Abrimos el formulario en modo diseño, pero oculto para que no moleste
        DoCmd.OpenForm doc.name, acDesign, , , , acHidden
        Set frm = Forms(doc.name)
        
        On Error Resume Next
        ' Buscamos el control específico
        Set ctl = frm.Controls("lblBloqueo")
        
        If Err.Number = 0 Then
            On Error GoTo Errores
            
            ' --- INICIO DE LA LÓGICA CORREGIDA ---
            ' Si se encontró el control, manipulamos su orden Z
            If traerAlFrente Then
                ctl.SetFocus
                'ctl.InSelection = True
                On Error Resume Next
                DoCmd.RunCommand acCmdBringToFront
                If Err.Number <> 2046 Then
                    Err.Raise Err.Number, Err.source, Err.description
                Else
                    Err.Clear
                    On Error GoTo Errores
                End If
            Else
                ctl.InSelection = True
                On Error Resume Next
                DoCmd.RunCommand acCmdSendToBack
                If Err.Number <> 2046 Then
                    Err.Raise Err.Number, Err.source, Err.description
                Else
                    Err.Clear
                    On Error GoTo Errores
                End If
            End If
            ' --- FIN DE LA LÓGICA CORREGIDA ---
            
            Debug.Print "  -> Procesado '" & doc.name & "'... OK"
            contador = contador + 1
            
            ' Guardamos y cerramos el formulario
            DoCmd.Close acForm, doc.name, acSaveYes
        Else
            ' Si el control no se encontró, limpiamos el error y continuamos
            Err.Clear
            Debug.Print "  -> Omitiendo '" & doc.name & "' (no tiene lblBloqueo)."
            ' Cerramos sin guardar por si acaso, aunque no debería haber cambios
            DoCmd.Close acForm, doc.name, acSaveNo
        End If
        
        Set ctl = Nothing
        Set frm = Nothing
        On Error GoTo Errores ' Reactivamos el manejador de errores principal
    Next doc
    
    Debug.Print "------------------------------------------------------"
    Debug.Print "PROCESO FINALIZADO. Se han ajustado " & contador & " paneles de bloqueo."
    Debug.Print "======================================================"
    
LimpiarYSalir:
    On Error Resume Next
    Application.Echo True ' Reactivamos el refresco de pantalla
    Set doc = Nothing
    Set db = Nothing
    MsgBox "Proceso finalizado. Se han ajustado " & contador & " paneles de bloqueo.", vbInformation, "Herramienta de Desarrollador"
    Exit Sub
    
Errores:
    MsgBox "Se ha producido un error inesperado durante el proceso:" & vbCrLf & vbCrLf & _
           "Error #" & Err.Number & ": " & Err.description, vbCritical, "Error en DevTool"
    Resume LimpiarYSalir
End Sub


' --- FUNCIONES DE ACCESO RÁPIDO (SIN CAMBIOS) ---

Public Sub Helper_PanelesBloqueo_AlFrente()
    ' Para preparar la aplicación para una demo o entrega al cliente
    Call AjustarPanelesBloqueo(True)
End Sub

Public Sub Helper_PanelesBloqueo_AlFondo()
    ' Para preparar el entorno para desarrollo y diseño de formularios
    Call AjustarPanelesBloqueo(False)
End Sub

