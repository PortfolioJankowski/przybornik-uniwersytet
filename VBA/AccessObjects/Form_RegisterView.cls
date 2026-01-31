VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_RegisterView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnRegister_Click()
    Dim newUser As user: Set newUser = New user
    Set newUser = CreateUser()
    If Not newUser.IsValidUser Then
        MsgBox newUser.Komunikat, vbCritical, "B³¹d walidacji"
        Exit Sub
    End If
    
    If Not IsLoginAvailable() Then
        MsgBox "Login ju¿ istnieje", vbCritical, "B³¹d walidacji"
        Exit Sub
    End If
    
    Call HashUserPassword(newUser)
    Call InsertUser(newUser)
    Call ChangeConfiguration
    
    MsgBox "U¿ytkownik utworzony", vbOKOnly, "Sukces"
    
    DoCmd.Close acForm, Me.name, acSaveNo
    DoCmd.OpenForm "LoginView"
End Sub

Private Function CreateUser() As user
    Dim u As New user

    With u
        .NazwaUzytkownika = Trim(NzSafe(Me.txtLogin.value, ""))
        .Imie = Trim(NzSafe(Me.txtName.value, ""))
        .Nazwisko = Trim(NzSafe(Me.txtSurname.value, ""))
        .AdresEmail = Trim(NzSafe(Me.txtEmail.value, ""))
        .Tytul = Trim(NzSafe(Me.txtTitle.value, ""))
        .Haslo = NzSafe(Me.txtPassword.value, "")
        .PowtorzHaslo = NzSafe(Me.txtRepeatPassword.value, "")
    End With

    Set CreateUser = u
End Function

Private Sub HashUserPassword(ByRef user As user)
    Dim cryptoService As cryptoService
    Set cryptoService = New cryptoService
    user.HashHaslo = cryptoService.HashPassword(user.Haslo)
End Sub

Private Sub InsertUser(ByRef user As user)
    Dim ORM As New UserMapper
    Dim query As String
    
    query = ORM.ToInsertSql(user)
    
    Dim db As DAO.database
    Set db = CurrentDb
    
    db.Execute query, dbFailOnError
    
    Set db = Nothing
End Sub

Private Function IsLoginAvailable() As Boolean
    Dim db As DAO.database
    Dim rs As DAO.Recordset
    Dim loginToCheck As String
    
    loginToCheck = Trim(Me.txtLogin.value)
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Uzytkownicy", dbOpenDynaset)
    
    rs.Filter = "NazwaUzytkownika = '" & Replace(loginToCheck, "'", "''") & "'"
    
    If rs.RecordCount > 0 Then
        IsLoginAvailable = False
    Else
        IsLoginAvailable = True
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

Private Sub ChangeConfiguration()
     Dim rs As Recordset
            Set rs = CurrentDb.OpenRecordset("SELECT * FROM Konfiguracja WHERE Wlasciwosc = '" & Consts.CONFIG_FIRST_TIME_LOGIN & "'", dbOpenDynaset)
                rs.Edit
                rs!Wartosc = "False"
                rs.Update
End Sub
