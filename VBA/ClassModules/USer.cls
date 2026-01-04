VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "USer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' === Fields ===
Private mId As Long
Private mNazwaUzytkownika As String
Private mHashHaslo As String
Private mImie As String
Private mNazwisko As String
Private mAdresEmail As String
Private mTytul As String
Private mKomunikat As String

' === Identity ===
Public Haslo As String
Public PowtorzHaslo As String
Public Property Get Id() As Long
    Id = mId
End Property
Public Property Let Id(ByVal value As Long)
    mId = value
End Property

' === Nazwa u¿ytkownika ===
Public Property Get NazwaUzytkownika() As String
    NazwaUzytkownika = mNazwaUzytkownika
End Property
Public Property Let NazwaUzytkownika(ByVal value As String)
    mNazwaUzytkownika = value
End Property

' === Has³o (hash) ===
Public Property Get HashHaslo() As String
    HashHaslo = mHashHaslo
End Property
Public Property Let HashHaslo(ByVal value As String)
    mHashHaslo = value
End Property

' === Imiê ===
Public Property Get Imie() As String
    Imie = mImie
End Property
Public Property Let Imie(ByVal value As String)
    mImie = value
End Property

' === Nazwisko ===
Public Property Get Nazwisko() As String
    Nazwisko = mNazwisko
End Property
Public Property Let Nazwisko(ByVal value As String)
    mNazwisko = value
End Property

' === Email ===
Public Property Get AdresEmail() As String
    AdresEmail = mAdresEmail
End Property
Public Property Let AdresEmail(ByVal value As String)
    mAdresEmail = value
End Property

' === Tytu³ ===
Public Property Get Tytul() As String
    Tytul = mTytul
End Property
Public Property Let Tytul(ByVal value As String)
    mTytul = value
End Property

Public Property Get Komunikat() As String
    Komunikat = mKomunikat
End Property

Public Property Let Komunikat(ByVal value As String)
    mKomunikat = value
End Property


' === Convenience ===
Public Property Get PelnaNazwa() As String
    PelnaNazwa = Trim(mImie & " " & mNazwisko)
End Property

Public Function IsValidUser() As Boolean
    ' Domyœlnie OK
    IsValidUser = False
    mKomunikat = ""

    If IsNullOrWhiteSpace(Me.NazwaUzytkownika) Then
        mKomunikat = "Uzupe³nij nazwê u¿ytkownika"
        Exit Function
    End If

    If IsNullOrWhiteSpace(Me.Imie) Then
        mKomunikat = "Uzupe³nij imiê"
        Exit Function
    End If

    If IsNullOrWhiteSpace(Me.Nazwisko) Then
        mKomunikat = "Uzupe³nij nazwisko"
        Exit Function
    End If

    If IsNullOrWhiteSpace(Me.AdresEmail) Then
        mKomunikat = "Uzupe³nij adres email"
        Exit Function
    End If

    If InStr(1, Me.AdresEmail, "@") = 0 Then
        mKomunikat = "Niepoprawny adres email"
        Exit Function
    End If

    If IsNullOrWhiteSpace(Me.Haslo) Then
        mKomunikat = "Uzupe³nij has³o"
        Exit Function
    End If

    If Len(Me.Haslo) < 8 Then
        mKomunikat = "Has³o musi mieæ co najmniej 8 znaków"
        Exit Function
    End If

    If Me.Haslo <> Me.PowtorzHaslo Then
        mKomunikat = "Has³a nie s¹ takie same"
        Exit Function
    End If


    IsValidUser = True
End Function
