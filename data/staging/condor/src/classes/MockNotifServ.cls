VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockNotifServ"
Option Compare Database
Option Explicit

' ==========================================================================
' CLASS: MockNotifServ
' PURPOSE: Test stub for NotificacionHelper. Implements the FIVE-argument
'          EnviarNotificacion(vm, asunto, destinatarioEmail, copiaEmail,
'          [nombreEstadoNotif]) seam that mirrors the real
'          NotificacionServicio.EnviarNotificacion signature so the helper
'          can be exercised without the real DAO + SolicitudViewModel
'          backend.
'
' Notes:
'   * Variant parameters absorb Null and any concrete type the helper
'     forwards, matching the helper's Variant-in policy.
'   * Internal flags let atoms verify the helper called the seam with the
'     expected arguments (m_Asunto / m_Destinatario / m_Copia /
'     m_NombreEstadoNotif / m_Count) and that failures are swallowed
'     (m_RaiseError drives the sad-path only when a test opts in).
'   * The mock NEVER raises on its own during happy-path tests; it only
'     raises when RaiseError = True.
'   * Property setters/getters stay short on purpose: this class is a test
'     fixture, not production code.
' ==========================================================================

Private m_RaiseError As Boolean
Private m_Called As Boolean
Private m_Count As Long
Private m_Asunto As String
Private m_Destinatario As String
Private m_Copia As String
Private m_NombreEstadoNotif As String

Public Property Let RaiseError(ByVal v As Boolean)
    m_RaiseError = v
End Property

Public Property Get RaiseError() As Boolean
    RaiseError = m_RaiseError
End Property

Public Property Get WasCalled() As Boolean
    WasCalled = m_Called
End Property

Public Property Get CallCount() As Long
    CallCount = m_Count
End Property

Public Property Get Asunto() As String
    Asunto = m_Asunto
End Property

Public Property Get Destinatario() As String
    Destinatario = m_Destinatario
End Property

Public Property Get Copia() As String
    Copia = m_Copia
End Property

Public Property Get NombreEstadoNotif() As String
    NombreEstadoNotif = m_NombreEstadoNotif
End Property

Public Sub EnviarNotificacion( _
    ByRef vm As Variant, _
    ByVal asunto As Variant, _
    ByVal destinatarioEmail As Variant, _
    ByVal copiaEmail As Variant, _
    Optional ByVal nombreEstadoNotif As Variant = "")

    m_Called = True
    m_Count = m_Count + 1
    m_Asunto = CStr(Nz(asunto, ""))
    m_Destinatario = CStr(Nz(destinatarioEmail, ""))
    m_Copia = CStr(Nz(copiaEmail, ""))
    m_NombreEstadoNotif = CStr(Nz(nombreEstadoNotif, ""))
    If m_RaiseError Then Err.Raise 424
End Sub

Public Sub ResetMock()
    m_RaiseError = False
    m_Called = False
    m_Count = 0
    m_Asunto = ""
    m_Destinatario = ""
    m_Copia = ""
    m_NombreEstadoNotif = ""
End Sub
