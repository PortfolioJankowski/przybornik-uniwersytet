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


Public Function Sha256Hash(ByVal text As String) As String
    Dim enc As Object
    Dim bytes() As Byte
    Dim hashBytes() As Byte
    Dim i As Long
    
    ' Tworzymy obiekt Microsoft.XMLDOM do obliczeñ
    Set enc = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    ' Konwersja tekstu na bajty
    bytes = StrConv(text, vbFromUnicode)
    
    ' Obliczamy hash
    hashBytes = enc.ComputeHash_2(bytes)
    
    ' Zamiana na HEX
    Sha256Hash = ""
    For i = LBound(hashBytes) To UBound(hashBytes)
        Sha256Hash = Sha256Hash & LCase(Right("0" & Hex(hashBytes(i)), 2))
    Next
End Function

Private Function CreateSalt(ByVal length As Long) As String
    Const chars As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    Dim i As Long

    Randomize
    For i = 1 To length
        CreateSalt = CreateSalt & mId(chars, 1 + Int(Rnd * Len(chars)), 1)
    Next
End Function

Public Function VerifyPassword(ByVal inputPassword As String, ByVal storedValue As String) As Boolean
    Dim parts() As String
    parts = Split(storedValue, ":")

    ' Sprawdzenie poprawnego formatu
    If UBound(parts) <> 2 Then
        VerifyPassword = False
        Exit Function
    End If

    If parts(0) <> "SHA256" Then
        VerifyPassword = False
        Exit Function
    End If

    Dim salt As String
    salt = parts(1)

    Dim expectedHash As String
    expectedHash = parts(2)

    ' Porównanie hashy w tej samej formie (ma³e litery)
    VerifyPassword = (LCase(Sha256Hash(inputPassword & salt)) = LCase(expectedHash))
End Function
