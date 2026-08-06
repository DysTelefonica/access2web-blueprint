Attribute VB_Name = "HashHelper"

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: HashHelper.bas
' DESCRIPCIÓN: Utilidades para cálculo de Hash (SHA256)
' IMPLEMENTACIÓN: API de Windows (advapi32.dll) - Sin .NET
' COMPATIBILIDAD: 32-bit y 64-bit Access
' ==========================================================================

Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" _
                            (ByRef phProv As LongPtr, ByVal pszContainer As String, ByVal pszProvider As String, _
                             ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" _
                            (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" _
                            (ByVal hProv As LongPtr, ByVal algId As Long, ByVal hKey As LongPtr, ByVal dwFlags As Long, _
                             ByRef phHash As LongPtr) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" _
                            (ByVal hHash As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" _
                            (ByVal hHash As LongPtr, pbData As Any, ByVal cbData As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" _
                            (ByVal hHash As LongPtr, ByVal dwParam As Long, pbData As Any, ByRef pcbData As Long, _
                             ByVal dwFlags As Long) As Long

Private Const PROV_RSA_AES    As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000

Private Const ALG_TYPE_ANY    As Long = 0
Private Const ALG_CLASS_HASH  As Long = 32768
Private Const ALG_SID_SHA_256 As Long = 12
Private Const CALG_SHA_256    As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_256)

Private Const HP_HASHVAL      As Long = 2

Public Function CalcularSHA256(ByVal texto As String) As String
    Dim hProv As LongPtr
    Dim hHash As LongPtr
    Dim abytHash(0 To 31) As Byte
    Dim lngLength As Long
    Dim lngResult As Long
    Dim strHash As String
    Dim i As Long
    Dim abytData() As Byte
    
    If Len(texto) = 0 Then
        CalcularSHA256 = ""
        Exit Function
    End If
    
    abytData = StrConv(texto, vbFromUnicode)
    
    strHash = ""
    
    If CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) <> 0 Then
        If CryptCreateHash(hProv, CALG_SHA_256, 0&, 0&, hHash) <> 0 Then
            lngLength = UBound(abytData) - LBound(abytData) + 1
            If lngLength > 0 Then
                lngResult = CryptHashData(hHash, abytData(LBound(abytData)), lngLength, 0&)
            Else
                lngResult = CryptHashData(hHash, ByVal 0&, 0&, 0&)
            End If
            If lngResult <> 0 Then
                lngLength = UBound(abytHash) - LBound(abytHash) + 1
                If CryptGetHashParam(hHash, HP_HASHVAL, abytHash(LBound(abytHash)), lngLength, 0) <> 0 Then
                    For i = 0 To lngLength - 1
                        strHash = strHash & Right$("0" & Hex$(abytHash(LBound(abytHash) + i)), 2)
                    Next
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0&
    End If
    
    CalcularSHA256 = LCase$(strHash)
End Function


