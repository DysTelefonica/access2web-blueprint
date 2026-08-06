' MÓDULO: Form_frmSplash.cls

Option Compare Database
Option Explicit

Public ProgresoActual As Long

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo errores
    
    ' Ocultamos Ribbon (bypass con /dev)
    If InStr(1, Nz(Command(), ""), "/dev", vbTextCompare) = 0 Then
        Call GestionarRibbon(False)
    End If
    
    Ajustar Me
    Call ResetAvanceCounter
    Me.lblProgresoBarra.Width = 0
    
    Exit Sub
errores:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Timer()
    Dim m_Error As String
    On Error GoTo errores
    
    ' Detenemos el timer para que no se ejecute dos veces
    Me.TimerInterval = 0
   
    ' 1. Carga de Configuración General (mueve la barra azul)
    Me.Caption = "Iniciando HPS SOLICITUDES ..."
    Me.lblEstado.Caption = "Cargando configuración..."
    DoEvents
    
    EVE m_Error
    If m_Error <> "" Then Err.Raise 1000, , m_Error
    
    ' 2. Carga Pesada de Datos (Calculado en Entorno)
    ' Como la barra de progreso ya habrá llegado casi al final con EVE,
    ' cambiamos el texto para que el usuario sepa que está trabajando.
    Me.lblEstado.Caption = "Analizando y calculando tareas pendientes..."
    DoEvents
    
    ' Esta es la llamada clave que reemplaza tu lógica anterior
    If Not m_ObjEntorno Is Nothing Then
        m_ObjEntorno.CalcularTotalesPendientes m_ObjUsuarioConectado, EsTecnico
    End If
    
    ' 3. Transición
    DoCmd.OpenForm "Form0BDOpciones"
    DoCmd.Close acForm, Me.Name
    
    Exit Sub
    
errores:
    MsgBox "Ha habido un error en el inicio:" & vbNewLine & _
           Err.Description & vbNewLine & m_Error, vbCritical, "Error Crítico"
    ' Intentamos abrir el principal de todas formas por seguridad
    DoCmd.OpenForm "Form0BDOpciones"
    DoCmd.Close acForm, Me.Name
End Sub
