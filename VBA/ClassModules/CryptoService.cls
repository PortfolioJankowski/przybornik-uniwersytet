VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CryptoService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' === Windows Crypto API ===
#If VBA7 Then
    Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" _
        Alias "CryptAcquireContextA" ( _
        ByRef hProv As LongPtr, _
        ByVal pszContainer As String, _
        ByVal pszProvider As String, _
        ByVal dwProvType As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" ( _
        ByVal hProv As LongPtr, _
        ByVal Algid As Long, _
        ByVal hKey As LongPtr, _
        ByVal dwFlags As Long, _
        ByRef phHash As LongPtr) As Long

    Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr, _
        ByRef pbData As Any, _
        ByVal dwDataLen As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr, _
        ByVal dwParam As Long, _
        ByRef pbData As Any, _
        ByRef pdwDataLen As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr) As Long

    Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" ( _
        ByVal hProv As LongPtr, _
        ByVal dwFlags As Long) As Long
#End If

' === Consts ===
Private Const PROV_RSA_AES As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CALG_SHA_256 As Long = &H800C
Private Const HP_HASHVAL As Long = &H2

'------------------------------------------------------------
' HashPassword
' Tworzy hash SHA-256 z losowym SALTem
' Zwraca: "SHA256:salt:hash"
'------------------------------------------------------------
Public Function HashPassword(ByVal password As String) As String
    Dim salt As String
    salt = CreateSalt(16)

    Dim hash As String
    hash = Sha256Hash(password & salt)

    HashPassword = "SHA256:" & salt & ":" & hash
End Function


Private Function Sha256Hash(ByVal text As String) As String
    Dim hProv As LongPtr, hHash As LongPtr
    Dim bytes() As Byte
    Dim hashLen As Long
    Dim hashBytes() As Byte
    Dim i As Long

    bytes = StrConv(text, vbFromUnicode)

    If CryptAcquireContext(hProv, vbNullString, vbNullString, _
        PROV_RSA_AES, CRYPT_VERIFYCONTEXT) = 0 Then Exit Function

    If CryptCreateHash(hProv, CALG_SHA_256, 0, 0, hHash) = 0 Then GoTo Cleanup

    CryptHashData hHash, bytes(0), UBound(bytes) + 1, 0

    CryptGetHashParam hHash, HP_HASHVAL, ByVal 0&, hashLen, 0
    ReDim hashBytes(hashLen - 1)

    CryptGetHashParam hHash, HP_HASHVAL, hashBytes(0), hashLen, 0

    For i = 0 To UBound(hashBytes)
        Sha256Hash = Sha256Hash & LCase(Right("0" & Hex(hashBytes(i)), 2))
    Next

Cleanup:
    If hHash <> 0 Then CryptDestroyHash hHash
    If hProv <> 0 Then CryptReleaseContext hProv, 0
End Function

Private Function CreateSalt(ByVal length As Long) As String
    Const chars As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    Dim i As Long

    Randomize
    For i = 1 To length
        CreateSalt = CreateSalt & mId(chars, 1 + Int(Rnd * Len(chars)), 1)
    Next
End Function

Public Function VerifyPassword(ByVal inputPassword As String, _
                               ByVal storedValue As String) As Boolean

    Dim parts() As String
    parts = Split(storedValue, ":")

    If UBound(parts) <> 2 Then Exit Function
    If parts(0) <> "SHA256" Then Exit Function

    Dim salt As String
    salt = parts(1)

    Dim expectedHash As String
    expectedHash = parts(2)

    VerifyPassword = _
        (Sha256Hash(inputPassword & salt) = expectedHash)
End Function
