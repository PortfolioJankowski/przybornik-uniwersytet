Attribute VB_Name = "AutoExec"
Option Compare Database
Option Explicit


Public Sub AutoExec()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM Konfiguracja WHERE Wlasciwosc = '" & Consts.CONFIG_FIRST_TIME_LOGIN & "'", dbOpenSnapshot)
    
    If Not rs.EOF Then
        If rs!Wartosc = "True" Then
            DoCmd.OpenForm "RegisterView", acNormal
        Else
            DoCmd.OpenForm "LoginView", acNormal
        End If
    Else
        MsgBox "Nie znaleziono konfiguracji!", vbExclamation
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub


