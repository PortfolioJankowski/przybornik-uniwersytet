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

' === Identity ===
Public Property Get Id() As Long
    Id = mId
End Property
Friend Property Let Id(ByVal value As Long)
    mId = value
End Property

' === Nazwa u¿ytkownika ===
Public Property Get NazwaUzytkownika() As String
    NazwaUzytkownika = mNazwaUzytkownika
End Property
Friend Property Let NazwaUzytkownika(ByVal value As String)
    mNazwaUzytkownika = value
End Property

' === Has³o (hash) ===
Public Property Get HashHaslo() As String
    HashHaslo = mHashHaslo
End Property
Friend Property Let HashHaslo(ByVal value As String)
    mHashHaslo = value
End Property

' === Imiê ===
Public Property Get Imie() As String
    Imie = mImie
End Property
Friend Property Let Imie(ByVal value As String)
    mImie = value
End Property

' === Nazwisko ===
Public Property Get Nazwisko() As String
    Nazwisko = mNazwisko
End Property
Friend Property Let Nazwisko(ByVal value As String)
    mNazwisko = value
End Property

' === Email ===
Public Property Get AdresEmail() As String
    AdresEmail = mAdresEmail
End Property
Friend Property Let AdresEmail(ByVal value As String)
    mAdresEmail = value
End Property

' === Tytu³ ===
Public Property Get Tytul() As String
    Tytul = mTytul
End Property
Friend Property Let Tytul(ByVal value As String)
    mTytul = value
End Property

' === Convenience ===
Public Property Get PelnaNazwa() As String
    PelnaNazwa = Trim(mImie & " " & mNazwisko)
End Property

