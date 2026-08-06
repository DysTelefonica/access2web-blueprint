VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JustificacionServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------
' CLASE: JustificacionServicio.cls
' DESCRIPCIÓN: Contiene la lógica de negocio para la gestión de Justificaciones.
'              Orquesta las validaciones y llamadas al Repositorio.
' -------------------------------------------------------------------------
Option Compare Database
Option Explicit

' --- MÉTODOS DE LECTURA ---

Public Function ObtenerListaParaSeleccion() As Collection
    ' Llama al repositorio para obtener todas las justificaciones activas
    ' Esta función es usada principalmente por m_ObjEntorno para llenar la caché
    Set ObtenerListaParaSeleccion = JustificacionRepositorio.ObtenerTodas()
End Function

Public Function ObtenerPorID(id As Long) As Justificacion
    ' Obtiene una entidad completa por su ID para edición
    If id <= 0 Then Exit Function
    Set ObtenerPorID = JustificacionRepositorio.ObtenerPorID(id)
End Function

' --- MÉTODOS DE ESCRITURA (Transaccionales) ---

Public Sub GuardarJustificacion(obj As Justificacion)
    ' 1. VALIDACIONES DE NEGOCIO
    ' Validamos que el objeto no sea nulo
    If obj Is Nothing Then
        Err.Raise 51000, "JustificacionServicio", "No se ha recibido una justificación válida."
    End If
    
    ' Validamos que el título no esté vacío
    If Len(Trim(obj.Titulo)) = 0 Then
        Err.Raise 51001, "JustificacionServicio", "El campo 'Título' es obligatorio."
    End If
    
    ' Validamos la longitud máxima (ajustar según tu tabla, usualmente 255 para Texto Corto)
    If Len(obj.Titulo) > 255 Then
        Err.Raise 51002, "JustificacionServicio", "El título es demasiado largo (máx 255 caracteres)."
    End If

    ' 2. PERSISTENCIA
    ' Delegamos el guardado físico al repositorio
    JustificacionRepositorio.Guardar obj
    
    ' 3. GESTIÓN DE CACHÉ
    ' Es vital invalidar la caché global para que el próximo desplegable muestre los cambios
    If Not m_ObjEntorno Is Nothing Then
        m_ObjEntorno.RecargarJustificaciones
    End If
End Sub

Public Sub EliminarJustificacion(id As Long)
    ' 1. VALIDACIONES
    If id <= 0 Then
        Err.Raise 51003, "JustificacionServicio", "ID de justificación no válido para eliminar."
    End If
    
    ' (Opcional) Aquí podrías validar si la justificación está siendo usada en Solicitudes activas
    ' Si existiera esa regla, se consultaría a un SolicitudRepositorio.
    
    ' 2. PERSISTENCIA (Borrado Lógico según arquitectura)
    JustificacionRepositorio.Eliminar id
    
    ' 3. GESTIÓN DE CACHÉ
    ' Forzamos recarga para que desaparezca de la lista
    If Not m_ObjEntorno Is Nothing Then
        m_ObjEntorno.RecargarJustificaciones
    End If
End Sub
' Añadir esta función para que el formulario sepa si bloquear controles
Public Function VerificarUso(id As Long) As Boolean
    If id <= 0 Then
        VerificarUso = False
    Else
        VerificarUso = JustificacionRepositorio.EstaSiendoUsada(id)
    End If
End Function
' Clase: JustificacionServicio.cls

Public Function ObtenerUltimoIDCreado() As Long
    ObtenerUltimoIDCreado = JustificacionRepositorio.ObtenerUltimoID()
End Function
Public Function getJustificacionParaCampo(Optional id As Long, Optional p_Justificacion As Justificacion) As String
    Dim m_Justificacion As Justificacion
    Const M_MAX_LEN As Long = 100
    If p_Justificacion Is Nothing Then
        Set m_Justificacion = ObtenerPorID(id)
    Else
        Set m_Justificacion = p_Justificacion
    End If
    
    If m_Justificacion Is Nothing Then
        Exit Function
    End If
    With m_Justificacion
        If Len(.Descripcion) > M_MAX_LEN Then
            getJustificacionParaCampo = .Titulo & " (" & Left(.Descripcion, M_MAX_LEN) & "...)"
        Else
            getJustificacionParaCampo = .Titulo & " (" & .Descripcion & ")"
        End If
    End With
    Set m_Justificacion = Nothing
    
End Function
