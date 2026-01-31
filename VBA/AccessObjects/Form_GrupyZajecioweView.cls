VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GrupyZajecioweView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnAddGroup_Click()
    DoCmd.OpenForm Consts.DodajGrupeView
End Sub

Private Sub btnAddStudent_Click()
    On Error GoTo catch
    If Me.Identyfikator = 0 Then Exit Sub
    Dim NrAlbumu As String: NrAlbumu = InputBox("Podaj nr albumu")
    
    If IsNullOrWhiteSpace(NrAlbumu) Then
        MsgBox "Podaj prawid³owy numer"
    End If
        
    Dim idStudenta As Variant
    idStudenta = SqlStringService.GetStudentIdByNrAlbumu(NrAlbumu)
    
    If idStudenta = Null Then
        MsgBox "Student o podanym numerze albumu nie widnieje w bazie"
        Exit Sub
    End If
    
    Dim query As String
    query = SqlStringService.InsertStudentToGroup(CStr(Me.Identyfikator), CStr(idStudenta))
    CurrentDb.Execute (query)
        
    Dim zdarzenie As zdarzenie
    zdarzenie = EventFactory("Dodawanie studenta o id " & CStr(idStudenta) & " do grupy o id " & Me.Identyfikator, _
    EventMdl.StudentAddedToGroup, True)
    EventGateway zdarzenie
        
    MsgBox "Pomyœlnie dodano studenta do grupy"
    Exit Sub
catch:
    MsgBox "Wyst¹pi³ nieoczekiwany b³¹d"
    Dim log As wpis
    
    With log
        .CallStac = err.HelpContext
        .Description = err.Description
        .ErrorNumber = err.Number
    End With
    
    App.Logger.Add log
End Sub

Private Sub btnExit_Click()
    DoCmd.Close acForm, Consts.GrupyZajecioweView
End Sub
