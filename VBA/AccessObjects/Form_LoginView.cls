VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LoginView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnLogin_Click()
    Dim cryptoService As New cryptoService
    Me.txtError.Visible = False
    
    If IsNullOrWhiteSpace(Me.txtLogin) Then
        Me.txtError = "Nieprawid這wy login"
        Me.txtError.Visible = True
        Exit Sub
    ElseIf IsNullOrWhiteSpace(Me.txtPassword) Then
        Me.txtError = "Uzupe軟ij has這"
        Me.txtError.Visible = True
        Exit Sub
    ElseIf Len(Me.txtPassword.value) < 8 Then
        Me.txtError = "Podano za kr鏒kie has這"
        Me.txtError.Visible = True
        Exit Sub
    
    End If
    
    
    Dim storedHash As String
    Dim rs As Recordset
    Dim db As Database
    Set db = CurrentDb
    
    Dim query As String
    query = "SELECT * FROM Uzytkownicy " & _
    "WHERE NazwaUzytkownika = '" & Replace(Me.txtLogin.value, "'", "''") & "'"
    Set rs = db.OpenRecordset(query, dbOpenSnapshot)
    
    If rs.RecordCount = 0 Then
        Me.txtError = "Nie znaleziono podanego loginu"
        Me.txtError.Visible = True
        Exit Sub
    End If
    
    If Not cryptoService.VerifyPassword(txtPassword, rs!HashHaslo) Then
        Me.txtError = "Nieprawid這we has這"
        Me.txtError.Visible = True
        Exit Sub
    End If

    Set db = Nothing
    Set rs = Nothing
    
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    DoCmd.OpenForm "LandingView"
End Sub
