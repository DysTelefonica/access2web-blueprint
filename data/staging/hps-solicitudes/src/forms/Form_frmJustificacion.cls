' Formulario: frmJustificacion
Option Compare Database
Option Explicit

' --- EVENTOS ---
' Este evento notifica al formulario buscador que hemos terminado
Public Event JustificacionGuardada(id As Long, Titulo As String)

Private m_Servicio As New JustificacionServicio
Private m_Entidad As Justificacion
Private m_IdActual As Long

Private Sub cmdSalir_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Load()
    ' Si pasamos un ID en OpenArgs, es edición. Si no, es alta.
    If Not IsNull(Me.OpenArgs) Then
        m_IdActual = CLng(Me.OpenArgs)
    Else
        m_IdActual = 0
    End If
    
    CargarDatos
    AplicarReglasNegocio
End Sub

Private Sub CargarDatos()
    If m_IdActual = 0 Then
        ' --- NUEVA ---
        Set m_Entidad = Factoria.CreateEntity("Justificacion")
        m_Entidad.activa = True
        Me.lblTitulo.Caption = "Nueva Justificación"
    Else
        ' --- EDICIÓN ---
        Set m_Entidad = m_Servicio.ObtenerPorID(m_IdActual)
        Me.lblTitulo.Caption = "Editar Justificación"
    End If
    
    ' Mapeo Objeto -> Formulario
    Me.txtTitulo = m_Entidad.Titulo
    Me.txtDescripcion = m_Entidad.Descripcion
    Me.chkActiva = m_Entidad.activa
End Sub

Private Sub AplicarReglasNegocio()
    ' REGLA: "Las justificaciones no se pueden cambiar ni borrar si ya han sido usadas"
    ' Consultamos al servicio si está en uso
    If m_IdActual > 0 Then
        If m_Servicio.VerificarUso(m_IdActual) Then
            ' BLOQUEAMOS edición de datos clave
            Me.txtTitulo.Locked = True
            Me.txtTitulo.BackColor = RGB(240, 240, 240) ' Gris visual
            
            Me.txtDescripcion.Locked = True
            Me.txtDescripcion.BackColor = RGB(240, 240, 240)
            
            ' Mostramos aviso
            Me.lblAviso.Visible = True
            
            ' Solo permitimos cambiar si está activa o no
            Me.chkActiva.SetFocus
        Else
            Me.lblAviso.Visible = False
        End If
    Else
        Me.lblAviso.Visible = False
    End If
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo Manejador
    
    ' ---------------------------------------------------------
    ' 1. DETECCIÓN DE CAMBIOS (Optimización)
    ' ---------------------------------------------------------
    ' Comparamos los valores actuales de la pantalla contra el objeto en memoria.
    ' Usamos Nz() para evitar errores con nulos.
    
    If Nz(Me.txtTitulo.value, "") = m_Entidad.Titulo And _
       Nz(Me.txtDescripcion.value, "") = m_Entidad.Descripcion And _
       Me.chkActiva.value = m_Entidad.activa Then
        
        ' Si no hay diferencias, avisamos y cerramos sin tocar la BD
        MsgBox "No se han detectado cambios para guardar.", vbInformation, "Sin cambios"
        ' Opcional: Cerrar el formulario directamente si se considera "Cancelación"
        ' DoCmd.Close acForm, Me.Name
        Exit Sub
    End If
    
    ' ---------------------------------------------------------
    ' 2. MAPEO DE DATOS
    ' ---------------------------------------------------------
    ' Volcamos los datos de los controles al objeto Entidad
    m_Entidad.Titulo = Nz(Me.txtTitulo, "")
    m_Entidad.Descripcion = Nz(Me.txtDescripcion, "")
    m_Entidad.activa = Me.chkActiva
    
    ' ---------------------------------------------------------
    ' 3. GUARDADO (Llamada al Servicio)
    ' ---------------------------------------------------------
    ' El servicio validará y el repositorio ejecutará el SQL (INSERT o UPDATE)
    m_Servicio.GuardarJustificacion m_Entidad
    
    ' ---------------------------------------------------------
    ' 4. RECUPERACIÓN DE ID (Corrección DMax)
    ' ---------------------------------------------------------
    Dim idFinal As Long
    
    If m_IdActual = 0 Then
        ' CASO NUEVA:
        ' No podemos usar DMax porque la tabla no está vinculada en CurrentDb.
        ' Usamos el método del servicio que consulta al Back-End directamente.
        idFinal = m_Servicio.ObtenerUltimoIDCreado()
    Else
        ' CASO EDICIÓN:
        ' Mantenemos el ID que ya teníamos
        idFinal = m_IdActual
    End If
    
    ' ---------------------------------------------------------
    ' 5. FINALIZACIÓN
    ' ---------------------------------------------------------
    ' Notificamos al formulario padre (Buscador) para que se actualice
    RaiseEvent JustificacionGuardada(idFinal, m_Entidad.Titulo)
    
    ' Cerramos este formulario
    DoCmd.Close acForm, Me.Name
    
    Exit Sub

Manejador:
    MsgBox "Error al guardar: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdCancelar_Click()
    DoCmd.Close acForm, Me.Name
End Sub