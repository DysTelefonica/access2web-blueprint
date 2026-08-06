Attribute VB_Name = "TestUnexpectedErrorDTO"
Option Compare Database
Option Explicit

Private Const CONTROLLED_ERROR_NUMBER As Long = vbObjectError + 2048
Private Const CONTROLLED_ERROR_SOURCE As String = "TestUnexpectedErrorDTO.Exists"
Private Const CONTROLLED_ERROR_DESCRIPTION As String = "controlled unexpected DTO Exists failure"

Public Function Exists(ByVal p_Key As Variant) As Boolean
    Err.Raise CONTROLLED_ERROR_NUMBER, CONTROLLED_ERROR_SOURCE, CONTROLLED_ERROR_DESCRIPTION
End Function
