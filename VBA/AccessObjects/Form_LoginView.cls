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
Private mUserRecordset As DAO.Recordset
Private Sub btnLogin_Click()
    On Error GoTo EH

    ClearError

    If Not ValidateInput() Then Exit Sub
    If Not TryAuthenticate() Then Exit Sub

    InitializeSession
    PublishLoginEvent
    NavigateToLanding
    DoCmd.OpenForm "EventDispatcher", , , , , acHidden

    Exit Sub
    
    
EH:
    ShowError "Wyst¹pi³ nieoczekiwany b³¹d"
End Sub

Private Function ValidateInput() As Boolean
    If IsNullOrWhiteSpace(Me.txtLogin) Then
        ShowError "Nieprawid³owy login"
    ElseIf IsNullOrWhiteSpace(Me.txtPassword) Then
        ShowError "Uzupe³nij has³o"
    ElseIf Len(Me.txtPassword.value) < 8 Then
        ShowError "Podano za krótkie has³o"
    Else
        ValidateInput = True
        Exit Function
    End If

    ValidateInput = False
End Function

Private Function TryAuthenticate() As Boolean
    Dim rs As DAO.Recordset
    Set rs = GetUserRecordset(Me.txtLogin.value)

    If rs.EOF Then
        ShowError "Nie znaleziono podanego loginu"
        Exit Function
    End If

    If Not VerifyPassword(rs) Then
        ShowError "Nieprawid³owe has³o"
        Exit Function
    End If

    Set mUserRecordset = rs
    TryAuthenticate = True
End Function

Private Function GetUserRecordset(login As String) As DAO.Recordset
    Dim sql As String
    sql = "SELECT * FROM Uzytkownicy WHERE NazwaUzytkownika = '" & _
          Replace(login, "'", "''") & "'"

    Set GetUserRecordset = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
End Function

Private Function VerifyPassword(rs As DAO.Recordset) As Boolean
    Dim crypto As New cryptoService
    VerifyPassword = crypto.VerifyPassword(Me.txtPassword.value, rs!HashHaslo)
End Function

Private Sub InitializeSession()
    Dim ORM As New UserMapper

    Set App = New AppContext
    Set App.CurrentUser = ORM.CreateUserFromRecord(mUserRecordset)
    Set App.db = CurrentDb
    Set App.Logger = New Logger
End Sub

Private Sub PublishLoginEvent()
    Dim e As Zdarzenie
    e.Name = EventMdl.UserLoggedInEvent
    e.Message = "Pomyœlne logowanie"

    EventGateway e
End Sub

Private Sub ShowError(msg As String)
    Me.txtError = msg
    Me.txtError.Visible = True
End Sub

Private Sub ClearError()
    Me.txtError.Visible = False
End Sub

Private Sub NavigateToLanding()
    DoCmd.Close acForm, Me.Name, acSaveNo
    DoCmd.OpenForm "LandingView"
End Sub
