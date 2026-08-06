Attribute VB_Name = "Critografia"
Option Compare Database
Option Explicit

'Reference: http://khoiriyyah.blogspot.com/2012/06/vb6-hash-class-md5-sha-1-sha-256-sha.html
'The above article is excellent but that code works for 32-bit Access/Excel only.
'This works on both 32-bit and 64-bit Access/Excel.
'Requires no dll References in the VB Editor.  It uses advapi32.dll, which needs no Reference
'  as these are Windows API calls only.
'I have not seen documented anywhere else on the Internet how to make these calls in 64-bit apps.

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

Public Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "BCrypt.dll" (ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
Public Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "BCrypt.dll" (ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
Public Declare PtrSafe Function BCryptGetProperty Lib "BCrypt.dll" (ByVal hObject As LongPtr, ByVal pszProperty As LongPtr, ByRef pbOutput As Any, ByVal cbOutput As Long, ByRef pcbResult As Long, ByVal dfFlags As Long) As Long
Public Declare PtrSafe Function BCryptSetProperty Lib "BCrypt.dll" (ByVal hObject As LongPtr, ByVal pszProperty As LongPtr, ByRef pbInput As Any, ByVal cbInput As Long, ByVal dfFlags As Long) As Long
Public Declare PtrSafe Function BCryptCreateHash Lib "BCrypt.dll" (ByVal hAlgorithm As LongPtr, ByRef phHash As LongPtr, pbHashObject As Any, ByVal cbHashObject As Long, ByVal pbSecret As LongPtr, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Public Declare PtrSafe Function BCryptHashData Lib "BCrypt.dll" (ByVal hHash As LongPtr, pbInput As Any, ByVal cbInput As Long, Optional ByVal dwFlags As Long = 0) As Long
Public Declare PtrSafe Function BCryptFinishHash Lib "BCrypt.dll" (ByVal hHash As LongPtr, pbOutput As Any, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
Public Declare PtrSafe Function BCryptDestroyHash Lib "BCrypt.dll" (ByVal hHash As LongPtr) As Long
Public Declare PtrSafe Function BCryptGenRandom Lib "BCrypt.dll" (ByVal hAlgorithm As LongPtr, pbBuffer As Any, ByVal cbBuffer As Long, ByVal dwFlags As Long) As Long
Public Declare PtrSafe Function BCryptGenerateSymmetricKey Lib "BCrypt.dll" (ByVal hAlgorithm As LongPtr, ByRef hKey As LongPtr, pbKeyObject As Any, ByVal cbKeyObject As Long, pbSecret As Any, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Public Declare PtrSafe Function BCryptEncrypt Lib "BCrypt.dll" (ByVal hKey As LongPtr, pbInput As Any, ByVal cbInput As Long, pPaddingInfo As Any, pbIV As Any, ByVal cbIV As Long, pbOutput As Any, ByVal cbOutput As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
Public Declare PtrSafe Function BCryptDecrypt Lib "BCrypt.dll" (ByVal hKey As LongPtr, pbInput As Any, ByVal cbInput As Long, pPaddingInfo As Any, pbIV As Any, ByVal cbIV As Long, pbOutput As Any, ByVal cbOutput As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
Public Declare PtrSafe Function BCryptDestroyKey Lib "BCrypt.dll" (ByVal hKey As LongPtr) As Long

Public Declare PtrSafe Sub RtlMoveMemory Lib "Kernel32.dll" (Destination As Any, Source As Any, ByVal Length As LongPtr)


Private Const PROV_RSA_FULL   As Long = 1
Private Const PROV_RSA_AES    As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000

Private Const HP_HASHVAL      As Long = 2
Private Const HP_HASHSIZE     As Long = 4

Private Const ALG_TYPE_ANY    As Long = 0
Private Const ALG_CLASS_HASH  As Long = 32768

Private Const ALG_SID_MD2     As Long = 1
Private Const ALG_SID_MD4     As Long = 2
Private Const ALG_SID_MD5     As Long = 3
Private Const ALG_SID_SHA     As Long = 4
Private Const ALG_SID_SHA_256 As Long = 12
Private Const ALG_SID_SHA_384 As Long = 13
Private Const ALG_SID_SHA_512 As Long = 14

Private Const CALG_MD2        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2)
Private Const CALG_MD4        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4)
Private Const CALG_MD5        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)
Private Const CALG_SHA        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)
Private Const CALG_SHA_256    As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_256)
Private Const CALG_SHA_384    As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_384)
Private Const CALG_SHA_512    As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_512)

Private Const clOneMask = 16515072          '000000 111111 111111 111111
Private Const clTwoMask = 258048            '111111 000000 111111 111111
Private Const clThreeMask = 4032            '111111 111111 000000 111111
Private Const clFourMask = 63               '111111 111111 111111 000000

Private Const clHighMask = 16711680         '11111111 00000000 00000000
Private Const clMidMask = 65280             '00000000 11111111 00000000
Private Const clLowMask = 255               '00000000 00000000 11111111

Private Const cl2Exp18 = 262144             '2 to the 18th power
Private Const cl2Exp12 = 4096               '2 to the 12th
Private Const cl2Exp6 = 64                  '2 to the 6th
Private Const cl2Exp8 = 256                 '2 to the 8th
Private Const cl2Exp16 = 65536              '2 to the 16th

Const BCRYPT_BLOCK_PADDING  As Long = &H1

Public Type QuadSextet
    s1 As Byte
    s2 As Byte
    s3 As Byte
    s4 As Byte
End Type

' Create Hash
Private Function CreateHash(abytData() As Byte, ByVal lngAlgID As Long) As String
    Dim hProv As LongPtr
    Dim hHash As LongPtr
    Dim abytHash(0 To 63) As Byte
    Dim lngLength As Long
    Dim lngResult As Long
    Dim strHash As String
    Dim i As Long
    Dim z As Long
    strHash = ""
    
    z = 0
    If CryptAcquireContext(hProv, vbNullString, vbNullString, _
                           IIf(lngAlgID >= CALG_SHA_256, PROV_RSA_AES, PROV_RSA_FULL), _
                           CRYPT_VERIFYCONTEXT) <> 0& Then
        If CryptCreateHash(hProv, lngAlgID, 0&, 0&, hHash) <> 0& Then
            lngLength = UBound(abytData()) - LBound(abytData()) + 1
            If lngLength > 0 Then lngResult = CryptHashData(hHash, abytData(LBound(abytData())), lngLength, 0&) _
                             Else lngResult = CryptHashData(hHash, ByVal 0&, 0&, 0&)
            If lngResult <> 0& Then
                lngLength = UBound(abytHash()) - LBound(abytHash()) + 1
                If CryptGetHashParam(hHash, HP_HASHVAL, abytHash(LBound(abytHash())), lngLength, z) <> 0& Then
                    For i = 0 To lngLength - 1
                        strHash = strHash & Right$("0" & Hex$(abytHash(LBound(abytHash()) + i)), 2)
                    Next
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0&
    End If
    CreateHash = LCase$(strHash)
End Function

' Create Hash From String(Shift_JIS)
Private Function CreateHashString(ByVal strData As String, ByVal lngAlgID As Long) As String
    CreateHashString = CreateHash(StrConv(strData, vbFromUnicode), lngAlgID)
End Function

' Create Hash From File
Private Function CreateHashFile(ByVal strFileName As String, ByVal lngAlgID As Long) As String
    Dim abytData() As Byte
    Dim intFile As Integer
    Dim lngError As Long
    On Error Resume Next
        If Len(Dir(strFileName)) > 0 Then
            intFile = FreeFile
            Open strFileName For Binary Access Read Shared As #intFile
            abytData() = InputB(LOF(intFile), #intFile)
            Close #intFile
        End If
        lngError = Err.Number
    On Error GoTo 0
    If lngError = 0 Then CreateHashFile = CreateHash(abytData(), lngAlgID) _
                    Else CreateHashFile = ""
End Function

' MD5
Public Function CreateMD5Hash(abytData() As Byte) As String
    CreateMD5Hash = CreateHash(abytData(), CALG_MD5)
End Function

Public Function CreateMD5HashString(ByVal strData As String) As String
    CreateMD5HashString = CreateHashString(strData, CALG_MD5)
End Function

Public Function CreateMD5HashFile(ByVal strFileName As String) As String
    CreateMD5HashFile = CreateHashFile(strFileName, CALG_MD5)
End Function

' SHA-1
Public Function CreateSHA1Hash(abytData() As Byte) As String
    CreateSHA1Hash = CreateHash(abytData(), CALG_SHA)
End Function

Public Function CreateSHA1HashString(ByVal strData As String) As String
    CreateSHA1HashString = CreateHashString(strData, CALG_SHA)
End Function

Public Function CreateSHA1HashFile(ByVal strFileName As String) As String
    CreateSHA1HashFile = CreateHashFile(strFileName, CALG_SHA)
End Function

' SHA-256
Public Function CreateSHA256Hash(abytData() As Byte) As String
    CreateSHA256Hash = CreateHash(abytData(), CALG_SHA_256)
End Function

Public Function SHA256(ByVal strData As String) As String
    SHA256 = CreateHashString(strData, CALG_SHA_256)
End Function

Public Function CreateSHA256HashFile(ByVal strFileName As String) As String
    CreateSHA256HashFile = CreateHashFile(strFileName, CALG_SHA_256)
End Function

' SHA-384
Public Function CreateSHA384Hash(abytData() As Byte) As String
    CreateSHA384Hash = CreateHash(abytData(), CALG_SHA_384)
End Function

Public Function SHA384(ByVal strData As String) As String
    SHA384 = CreateHashString(strData, CALG_SHA_384)
End Function

Public Function CreateSHA384HashFile(ByVal strFileName As String) As String
    CreateSHA384HashFile = CreateHashFile(strFileName, CALG_SHA_384)
End Function

' SHA-512
Public Function CreateSHA512Hash(abytData() As Byte) As String
    CreateSHA512Hash = CreateHash(abytData(), CALG_SHA_512)
End Function

Public Function SHA512(ByVal strData As String) As String
    SHA512 = CreateHashString(strData, CALG_SHA_512)
End Function

Public Function CreateSHA512HashFile(ByVal strFileName As String) As String
    CreateSHA512HashFile = CreateHashFile(strFileName, CALG_SHA_512)
End Function

Public Function ToBase64(b() As Byte) As String
    Const Base64Table = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim l As Long
    Dim output As String
    Dim UBoundOut As Long
    UBoundOut = UBound(b) + 1
    If UBoundOut Mod 3 <> 0 Then
        UBoundOut = UBoundOut + (3 - UBoundOut Mod 3)
    End If
    UBoundOut = (UBoundOut \ 3) * 4
    output = String(UBoundOut, vbNullChar)
    Dim qs As QuadSextet
    For l = 0 To (UBound(b) - 2) \ 3
        qs = BytesToQuadSextet(b(l * 3), b(l * 3 + 1), b(l * 3 + 2))
        Mid(output, (l * 4) + 1, 1) = Mid(Base64Table, qs.s1 + 1, 1)
        Mid(output, (l * 4) + 2, 1) = Mid(Base64Table, qs.s2 + 1, 1)
        Mid(output, (l * 4) + 3, 1) = Mid(Base64Table, qs.s3 + 1, 1)
        Mid(output, (l * 4) + 4, 1) = Mid(Base64Table, qs.s4 + 1, 1)
    Next
    If UBound(b) + 1 - (l * 3) = 2 Then
        qs = BytesToQuadSextet(b(l * 3), b(l * 3 + 1))
        Mid(output, (l * 4) + 1, 1) = Mid(Base64Table, qs.s1 + 1, 1)
        Mid(output, (l * 4) + 2, 1) = Mid(Base64Table, qs.s2 + 1, 1)
        Mid(output, (l * 4) + 3, 1) = Mid(Base64Table, qs.s3 + 1, 1)
        Mid(output, (l * 4) + 4, 1) = "="
    ElseIf UBound(b) + 1 - (l * 3) = 1 Then
        qs = BytesToQuadSextet(b(l * 3))
        Mid(output, (l * 4) + 1, 1) = Mid(Base64Table, qs.s1 + 1, 1)
        Mid(output, (l * 4) + 2, 1) = Mid(Base64Table, qs.s2 + 1, 1)
        Mid(output, (l * 4) + 3, 2) = "=="
    End If
    ToBase64 = output
End Function

Public Function Base64ToBytes(strBase64 As String) As Byte()
    Dim outBytes() As Byte
    Dim lenBytes As Long
    lenBytes = Len(strBase64) * 3 \ 4
    If Right(strBase64, 1) = "=" Then lenBytes = lenBytes - 1
    If Right(strBase64, 2) = "==" Then lenBytes = lenBytes - 1
    ReDim outBytes(0 To lenBytes - 1)
    Dim l As Long
    Dim qs As QuadSextet
    For l = 0 To lenBytes - 1
        Select Case l Mod 3
            Case 0
                qs = Base64ToQuadSextet(Mid(strBase64, (l \ 3) * 4 + 1, 4))
                outBytes(l) = qs.s1 * 2 ^ 2 + (qs.s2 \ 2 ^ 4)
            Case 1
                outBytes(l) = (qs.s2 * 2 ^ 4 And 255) + qs.s3 \ 2 ^ 2
            Case 2
                outBytes(l) = (qs.s3 * 2 ^ 6 And 255) + qs.s4
        End Select
    Next
    Base64ToBytes = outBytes
End Function

Public Function BytesToQuadSextet(b1 As Byte, Optional b2 As Byte, Optional b3 As Byte) As QuadSextet
    BytesToQuadSextet.s1 = b1 \ 4
    BytesToQuadSextet.s2 = (((b1 * 2 ^ 6) And 255) \ 4) + b2 \ (2 ^ 4)
    BytesToQuadSextet.s3 = (((b2 * 2 ^ 4) And 255) \ 4) + b3 \ (2 ^ 6)
    BytesToQuadSextet.s4 = (((b3 * 2 ^ 2) And 255) \ 4)
End Function

Public Function Base64ToQuadSextet(strBase64 As String) As QuadSextet
    Const Base64Table = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    Base64ToQuadSextet.s1 = InStr(Base64Table, Mid(strBase64, 1, 1)) - 1
    Base64ToQuadSextet.s2 = InStr(Base64Table, Mid(strBase64, 2, 1)) - 1
    Base64ToQuadSextet.s3 = InStr(Base64Table, Mid(strBase64, 3, 1)) - 1
    Base64ToQuadSextet.s4 = InStr(Base64Table, Mid(strBase64, 4, 1)) - 1
End Function

Public Function StringToBase64(str As String) As String
    StringToBase64 = ToBase64(StrConv(str, vbFromUnicode))
End Function

Public Function HashBytes(Data() As Byte, Optional HashingAlgorithm As String = "SHA1") As Byte()
    HashBytes = NGHash(VarPtr(Data(LBound(Data))), UBound(Data) - LBound(Data) + 1, HashingAlgorithm)
End Function

Public Function NGHash(pData As LongPtr, lenData As Long, Optional HashingAlgorithm As String = "SHA1") As Byte()
    'Erik A, 2019
    'Hash data by using the Next Generation Cryptography API
    'Loosely based on https://learn.microsoft.com/en-us/windows/desktop/SecCNG/creating-a-hash-with-cng
    'Allowed algorithms:  https://learn.microsoft.com/en-us/windows/desktop/SecCNG/cng-algorithm-identifiers. Note: only hash algorithms, check OS support
    'Error handling not implemented!
    On Error GoTo VBErrHandler
    Dim errorMessage As String

    Dim hAlg As LongPtr
    Dim algId As String

    'Open crypto provider
    algId = HashingAlgorithm & vbNullChar
    If BCryptOpenAlgorithmProvider(hAlg, StrPtr(algId), 0, 0) Then GoTo ErrHandler
    
    'Determine hash object size, allocate memory
    Dim bHashObject() As Byte
    Dim cmd As String
    cmd = "ObjectLength" & vbNullString
    Dim Length As Long
    If BCryptGetProperty(hAlg, StrPtr(cmd), Length, LenB(Length), 0, 0) <> 0 Then GoTo ErrHandler
    ReDim bHashObject(0 To Length - 1)
    
    'Determine digest size, allocate memory
    Dim hashLength As Long
    cmd = "HashDigestLength" & vbNullChar
    If BCryptGetProperty(hAlg, StrPtr(cmd), hashLength, LenB(hashLength), 0, 0) <> 0 Then GoTo ErrHandler
    Dim bHash() As Byte
    ReDim bHash(0 To hashLength - 1)
    
    'Create hash object
    Dim hHash As LongPtr
    If BCryptCreateHash(hAlg, hHash, bHashObject(0), Length, 0, 0, 0) <> 0 Then GoTo ErrHandler
    
    'Hash data
    If BCryptHashData(hHash, ByVal pData, lenData) <> 0 Then GoTo ErrHandler
    If BCryptFinishHash(hHash, bHash(0), hashLength, 0) <> 0 Then GoTo ErrHandler
    
    'Return result
    NGHash = bHash
ExitHandler:
    'Cleanup
    If hAlg <> 0 Then BCryptCloseAlgorithmProvider hAlg, 0
    If hHash <> 0 Then BCryptDestroyHash hHash
    Exit Function
VBErrHandler:
    errorMessage = "VB Error " & Err.Number & ": " & Err.Description
ErrHandler:
    If errorMessage <> "" Then MsgBox errorMessage
    Resume ExitHandler
End Function

Public Sub NGRandom(pData As LongPtr, lenData As Long, Optional Algorithm As String = "RNG")
    'Erik A, 2019
    'Fills data at pointer with random bytes
    'Error handling not implemented!
    
    Dim hAlg As LongPtr
    Dim algId As String
    
    'Open crypto provider
    algId = Algorithm & vbNullChar
    BCryptOpenAlgorithmProvider hAlg, StrPtr(algId), 0, 0
    
    'Fill bytearray with random data
    BCryptGenRandom hAlg, ByVal pData, lenData, 0
    
    'Cleanup
    BCryptCloseAlgorithmProvider hAlg, 0
End Sub

Public Sub NGRandomW(Data() As Byte, Optional Algorithm As String = "RNG")
    If LBound(Data) = -1 Then Exit Sub
    NGRandom VarPtr(Data(LBound(Data))), UBound(Data) - LBound(Data) + 1, Algorithm
End Sub

Public Function NGEncrypt(pData As LongPtr, lenData As Long, inpIV As LongPtr, inpIVLength As Long, inpSecret As LongPtr, inpSecretLength As Long) As Byte()
    'Encrypt pData using AES encryption, inpIV and inpSecret
    'Input: pData -> mempointer to data. lenData: amount of bytes to encrypt. inpIV: mempointer to IV. inpSecret: mempointer to 128-bits secret.
    'Output: Bytearray containing encrypted data
    Dim errorMessage As String
    On Error GoTo VBErrHandler
    
    Dim hAlg As LongPtr
    Dim algId As String

    'Open algorithm provider
    algId = "AES" & vbNullChar
    BCryptOpenAlgorithmProvider hAlg, StrPtr(algId), 0, 0

    'Allocate memory to hold the KeyObject
    Dim cmd As String
    Dim keyObjectLength As Long
    cmd = "ObjectLength" & vbNullString
    BCryptGetProperty hAlg, StrPtr(cmd), keyObjectLength, LenB(keyObjectLength), 0, 0
    Dim bKeyObject() As Byte
    ReDim bKeyObject(0 To keyObjectLength - 1)

    'Check block length = 128 bits, copy IV
    Dim ivLength As Long
    Dim bIV() As Byte
    cmd = "BlockLength" & vbNullChar
    BCryptGetProperty hAlg, StrPtr(cmd), ivLength, LenB(ivLength), 0, 0
    If ivLength > inpIVLength Then
        Debug.Print
    End If
    ReDim bIV(0 To ivLength - 1)
    RtlMoveMemory bIV(0), ByVal inpIV, ivLength

    'Set chaining mode
    cmd = "ChainingMode" & vbNullString
    Dim val As String
    val = "ChainingModeCBC" & vbNullString
    BCryptSetProperty hAlg, StrPtr(cmd), ByVal StrPtr(val), LenB(val), 0

    'Create KeyObject using secret
    Dim hKey As LongPtr
    BCryptGenerateSymmetricKey hAlg, hKey, bKeyObject(0), keyObjectLength, ByVal inpSecret, inpSecretLength, 0

    'Calculate output buffer size, allocate output buffer
    Dim cipherTextLength As Long
    BCryptEncrypt hKey, ByVal pData, lenData, ByVal 0, bIV(0), ivLength, ByVal 0, 0, cipherTextLength, BCRYPT_BLOCK_PADDING
    Dim bCipherText() As Byte
    ReDim bCipherText(0 To cipherTextLength - 1)

    'Encrypt the data
    Dim dataLength As Long
    BCryptEncrypt hKey, ByVal pData, lenData, ByVal 0, bIV(0), ivLength, bCipherText(0), cipherTextLength, dataLength, BCRYPT_BLOCK_PADDING
    
    'Output the encrypted data
    NGEncrypt = bCipherText
    
ExitHandler:
    'Destroy the key
    If hKey <> 0 Then BCryptDestroyKey hKey
    If hAlg <> 0 Then BCryptCloseAlgorithmProvider hAlg, 0
    Exit Function
VBErrHandler:
    errorMessage = "VB Error " & Err.Number & ": " & Err.Description
ErrHandler:
    If errorMessage <> "" Then MsgBox errorMessage
    Resume ExitHandler
End Function

Public Function NGEncryptW(pData() As Byte, pIV() As Byte, pSecret() As Byte) As Byte()
    NGEncryptW = NGEncrypt(VarPtr(pData(LBound(pData))), UBound(pData) - LBound(pData) + 1, VarPtr(pIV(LBound(pIV))), UBound(pIV) - LBound(pIV) + 1, VarPtr(pSecret(LBound(pSecret))), UBound(pSecret) - LBound(pSecret) + 1)
End Function


Public Function NGDecrypt(pData As LongPtr, lenData As Long, pIV As LongPtr, lenIV As Long, pSecret As LongPtr, lenSecret As Long) As Byte()
    Dim errorMessage As String
    On Error GoTo VBErrHandler
    Dim hAlg As LongPtr
    Dim algId As String

    'Open algorithm provider
    algId = "AES" & vbNullChar
    If BCryptOpenAlgorithmProvider(hAlg, StrPtr(algId), 0, 0) <> 0 Then GoTo ErrHandler

    'Allocate memory to hold the KeyObject
    Dim cmd As String
    Dim keyObjectLength As Long
    cmd = "ObjectLength" & vbNullString
    If BCryptGetProperty(hAlg, StrPtr(cmd), keyObjectLength, LenB(keyObjectLength), 0, 0) <> 0 Then GoTo ErrHandler
    Dim bKeyObject() As Byte
    ReDim bKeyObject(0 To keyObjectLength - 1)

    'Calculate the block length for the IV, resize the IV
    Dim ivLength As Long
    Dim bIV() As Byte
    cmd = "BlockLength" & vbNullChar
    If BCryptGetProperty(hAlg, StrPtr(cmd), ivLength, LenB(ivLength), 0, 0) <> 0 Then GoTo ErrHandler
    ReDim bIV(0 To ivLength - 1)
    RtlMoveMemory bIV(0), ByVal pIV, ivLength

    'Set chaining mode
    cmd = "ChainingMode" & vbNullString
    Dim val As String
    val = "ChainingModeCBC" & vbNullString
    If BCryptSetProperty(hAlg, StrPtr(cmd), ByVal StrPtr(val), LenB(val), 0) <> 0 Then GoTo ErrHandler

    'Create KeyObject using secret
    Dim hKey As LongPtr
    If BCryptGenerateSymmetricKey(hAlg, hKey, bKeyObject(1), keyObjectLength, ByVal pSecret, lenSecret, 0) <> 0 Then GoTo ErrHandler

    'Calculate output buffer size, allocate output buffer
    Dim OutputSize As Long
    If BCryptDecrypt(hKey, ByVal pData, lenData, ByVal 0, bIV(0), ivLength, ByVal 0, 0, OutputSize, BCRYPT_BLOCK_PADDING) <> 0 Then GoTo ErrHandler
    Dim bDecrypted() As Byte
    ReDim bDecrypted(0 To OutputSize - 1)

    'Decrypt the data
    Dim dataLength As Long
    If BCryptDecrypt(hKey, ByVal pData, lenData, ByVal 0, bIV(0), ivLength, bDecrypted(0), OutputSize, dataLength, BCRYPT_BLOCK_PADDING) <> 0 Then GoTo ErrHandler

    NGDecrypt = bDecrypted
    
    'Cleanup
ExitHandler:
    BCryptDestroyKey hKey
    BCryptCloseAlgorithmProvider hAlg, 0
    Exit Function
VBErrHandler:
    errorMessage = "VB Error " & Err.Number & ": " & Err.Description
ErrHandler:
    If errorMessage <> "" Then MsgBox errorMessage
    GoTo ExitHandler
End Function

Public Function NGDecryptW(pData() As Byte, pIV() As Byte, pSecret() As Byte) As Byte()
    NGDecryptW = NGDecrypt(VarPtr(pData(LBound(pData))), UBound(pData) - LBound(pData) + 1, VarPtr(pIV(LBound(pIV))), UBound(pIV) - LBound(pIV) + 1, VarPtr(pSecret(LBound(pSecret))), UBound(pSecret) - LBound(pSecret) + 1)
End Function

Public Function EncryptData(inpData() As Byte, inpKey() As Byte) As Byte()
    'SHA1 the key and data
    Dim keyHash() As Byte
    keyHash = HashBytes(inpKey, "SHA1")
    Dim dataHash() As Byte
    dataHash = HashBytes(inpData, "SHA1")
    Dim dataLength As Long
    dataLength = UBound(inpData) - LBound(inpData) + 1
    Dim toEncrypt() As Byte
    'To encrypt = Long (4 bytes) + dataLength + SHA1 (20 bytes)
    ReDim toEncrypt(0 To dataLength + 23)
    'Append length (in bytes) to start of array
    RtlMoveMemory toEncrypt(0), dataLength, 4
    'Then data
    RtlMoveMemory toEncrypt(4), inpData(LBound(inpData)), dataLength
    'Then hash of data
    RtlMoveMemory toEncrypt(dataLength + 4), dataHash(0), 20
    
    'Generate IV
    Dim IV(0 To 15) As Byte
    NGRandomW IV
    'Encrypt data
    Dim encryptedData() As Byte
    encryptedData = NGEncrypt(VarPtr(toEncrypt(0)), dataLength + 24, VarPtr(IV(0)), 16, VarPtr(keyHash(0)), 16)
    'Deallocate copy made to encrypt
    Erase toEncrypt
    'Extend encryptedData to append IV
    ReDim Preserve encryptedData(LBound(encryptedData) To UBound(encryptedData) + 16)
    'Append IV
    RtlMoveMemory encryptedData(UBound(encryptedData) - 15), IV(0), 16
    'Return result
    EncryptData = encryptedData
End Function

Public Function DecryptData(inpData() As Byte, inpKey() As Byte, outDecrypted() As Byte) As Boolean
    If LBound(inpData) <> 0 Then Exit Function 'Array must start at 0
    Dim arrLength As Long
    arrLength = UBound(inpData) + 1
    'IV = 16 bytes, length = 4 bytes
    If arrLength < 20 Then Exit Function
    'SHA1 the key
    Dim keyHash() As Byte
    keyHash = HashBytes(inpKey, "SHA1")
    'Get the pointer to the IV
    Dim pIV As LongPtr
    pIV = VarPtr(inpData(UBound(inpData) - 15)) 'Last 16 bytes = IV
    'Decrypt the data
    Dim decryptedData() As Byte
    decryptedData = NGDecrypt(VarPtr(inpData(0)), UBound(inpData) - LBound(inpData) - 15, pIV, 16, VarPtr(keyHash(0)), 16)
    'Check we got some data
    If StrPtr(decryptedData) = 0 Then Exit Function ' Weirdly, this checks for uninitialized byte arrays
    If UBound(decryptedData) < 3 Then Exit Function
    'Get the data length
    Dim dataLength As Long
    RtlMoveMemory dataLength, decryptedData(0), 4
    'Check if length is valid, with invalid key length = random data
    If dataLength > (UBound(decryptedData) - 3) Or dataLength < 0 Then Exit Function
    'Hash the decrypted data
    Dim hashResult() As Byte
    hashResult = NGHash(VarPtr(decryptedData(4)), dataLength, "SHA1")
    'Verify the hash
    Dim l As Byte
    For l = 0 To 19
        If hashResult(l) <> decryptedData(l + 4 + dataLength) Then
            'Stored hash not equal to hash with decrypted data, key incorrect or encrypted data tampered with
            'Don't touch output, return false by default
            Exit Function
        End If
    Next

    'Initialize output array
    ReDim outDecrypted(0 To dataLength - 1)
    'Copy data to output array
    RtlMoveMemory outDecrypted(0), decryptedData(4), dataLength
    DecryptData = True
End Function

Public Function Encriptar(inpString As String, inpKey As String) As String
    Dim Data() As Byte
    Data = inpString
    Dim key() As Byte
    key = inpKey
    Encriptar = ToBase64(EncryptData(Data, key))
End Function

Public Function DesEncriptar(inpEncryptedString As String, inpKey As String) As String
    Dim Data() As Byte
    Data = Base64ToBytes(inpEncryptedString)
    Dim key() As Byte
    key = inpKey
    Dim out() As Byte
    DecryptData Data, key, out
    DesEncriptar = out
End Function






