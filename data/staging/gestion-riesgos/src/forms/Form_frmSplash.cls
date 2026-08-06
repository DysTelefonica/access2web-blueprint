' MÓDULO: Form_frmSplash.cls
Attribute VB_Name = "Form_frmSplash"
Option Compare Database
Option Explicit

Public ProgresoActual As Long
'-------------------------------------------
' Nombre: Form_Load (inicialización del modelo jerárquico)
' Propósito: Establecer la bandera de entorno por defecto para la cadena jerárquica
' Parámetros: ninguno
' Retorno: ninguno
'-------------------------------------------
Private Sub Form_Load()
    On Error Resume Next
End Sub
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo errores
    
    ' Ocultamos Ribbon (bypass con /dev)
    If InStr(1, Nz(Command(), ""), "/dev", vbTextCompare) = 0 Then
        Call GestionarRibbon(False)
    End If
    
    AjustarTamaño Me
    Call ResetAvanceCounter
    Me.lblProgresoBarra.Width = 0
    
    Exit Sub
errores:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Timer()
    Dim m_Error As String
    Dim m_frmNombre As String
    On Error GoTo errores
    
    ' Detenemos el timer para que no se ejecute dos veces
    Me.TimerInterval = 0
   
    ' 1. Carga de Configuración General (mueve la barra azul)
    Me.Caption = "GESTIÓN DE RIESGOS - Iniciando..."
    Me.lblEstado.Caption = "Cargando configuración..."
    DoEvents
    m_frmNombre = Me.Name
    EVE m_Error
    If m_Error <> "" Then Err.Raise 1000, , m_Error
    
    ' 2. Carga Pesada de Datos (Calculado en Entorno)
    ' Como la barra de progreso ya habrá llegado casi al final con EVE,
    ' cambiamos el texto para que el usuario sepa que está trabajando.
    Me.lblEstado.Caption = "Analizando y calculando tareas pendientes..."
    DoEvents
    Set m_ObjUltimoProyecto = Constructor.getUltimoProyecto(, m_ObjUsuarioConectado.UsuarioRed, m_Error)
    If m_Error <> "" Then Err.Raise 1000, , m_Error
    If EsTecnico = EnumSiNo.Sí Then
        EstablecerTareasTecnico EnumSiNo.Sí, m_Error
        If m_Error <> "" Then Err.Raise 1000, , m_Error
        DoCmd.OpenForm "Form0BDOpcionesTecnico"
        If m_ObjTareasTecnico.TareasTotales <> "0" Then
            DoCmd.OpenForm "FormTecnicoTareas"
        End If
    Else
        m_ListaUsuarios = m_ObjEntorno.ListaUsuarios
        Set m_ObjUsuarioParaTareas = m_ObjUsuarioConectado
        EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.Sí, m_Error
        If m_Error <> "" Then Err.Raise 1000
        DoCmd.OpenForm "Form0BDOpciones"
        If m_ObjTareasCalidad.HayTareasPendientes = EnumSiNo.Sí Then
            DoCmd.OpenForm "FormCalidadTareas"
        End If
    End If
    
    
    ' 3. Transición
    
    DoCmd.Close acForm, m_frmNombre
    
    Exit Sub
    
errores:
    MsgBox "Ha habido un error en el inicio:" & vbNewLine & _
           Err.Description & vbNewLine & m_Error, vbCritical, "Error Crítico"
    ' Intentamos abrir el principal de todas formas por seguridad
    
    DoCmd.Close acForm, m_frmNombre
End Sub
