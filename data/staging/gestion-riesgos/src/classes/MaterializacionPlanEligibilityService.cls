VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaterializacionPlanEligibilityService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================
' MaterializacionPlanEligibilityService (issue #40, slice 1)
' Pure decision logic for materialization plan eligibility and
' editability. Extracted from Form_FormRiesgoMaterializado.
'
' Slice 1 scope: combine the two orthogonal signals (edicion-active,
' user-authorized) into the single "is materialization editable"
' boolean. State lookups (Edicion.EsActivo, Proyecto.UsuarioAutorizado)
' stay in the form; the service is the pure combinator.
'
' Slice 2 (next): wrap the state lookups so the form passes just
' the Edicion and the service resolves everything.
' ============================================================

Public Function IsMaterializationEditable( _
    ByVal p_EdicionActiva As EnumSiNo, _
    ByVal p_UsuarioAutorizado As EnumSiNo _
) As Boolean
    IsMaterializationEditable = (p_EdicionActiva = EnumSiNo.Sí) And _
                                (p_UsuarioAutorizado = EnumSiNo.Sí)
End Function
