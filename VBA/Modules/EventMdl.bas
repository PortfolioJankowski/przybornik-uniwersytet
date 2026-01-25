Attribute VB_Name = "EventMdl"
Option Compare Database
Option Explicit

Public Const UserLoggedInEvent = "Zalogowano-uzytkownika"
Public Const UserRegisterdedEvent = "Zarejestrowano-uzytkownika"
Public Const SubjectAddedEvent = "Dodano-przedmiot"
Public Const CycleAddedEvent = "Dodano-cykl-dydaktyczny"
Public Const PresentationDownloadedEvent = "Pobrano-prezentacje"
Public Const SubjectDeletedEvent = "Usuniêto-przedmiot"

Public Sub EventGateway(toSend As Zdarzenie)
    Dim rs As Recordset
    
    Set rs = CurrentDb.OpenRecordset("Zdarzenia", dbOpenDynaset)
    With rs
        .AddNew
            rs!Data = Now
            rs!Nazwa = toSend.Name
            rs!Wiadomosc = toSend.Message
            rs!CzyPrzetworzono = toSend.IsProcessed
        .Update
    End With
End Sub


Public Sub ProcessPendingEvents()
    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT * FROM Zdarzenia " & _
          "WHERE CzyPrzetworzono = False " & _
          "ORDER BY Identyfikator"

    Set rs = CurrentDb.OpenRecordset(sql, dbOpenDynaset)

    Do While Not rs.EOF
        DispatchEvent rs
        rs.MoveNext
    Loop

    rs.Close
End Sub

Private Sub DispatchEvent(eventRs As Recordset)
    On Error GoTo EH

    Select Case eventRs!Nazwa

        Case UserRegisterdedEvent
        Case UserLoggedInEvent
            Debug.Print "Siema user"
        

    End Select

    MarkEventAsProcessed eventRs
    Exit Sub

EH:
    Dim wpis As wpis
    wpis.ErrorNumber = err.Number
    wpis.Description = err.Description
    wpis.CallStac = "Dispatch event while dispatching " & eventRs!Name
    App.Logger.Add wpis
End Sub

Private Sub MarkEventAsProcessed(rs As DAO.Recordset)
    rs.Edit
        rs!Processed = True
        rs!ProcessedAt = Now
    rs.Update
End Sub
