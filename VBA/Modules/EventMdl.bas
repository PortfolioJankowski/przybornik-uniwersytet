Attribute VB_Name = "EventMdl"
Option Compare Database
Option Explicit

Public Const UserLoggedInEvent = "Zalogowano-uzytkownika"
Public Const UserRegisterdedEvent = "Zarejestrowano-uzytkownika"
Public Const SubjectAddedEvent = "Dodano-przedmiot"
Public Const CycleAddedEvent = "Dodano-cykl-dydaktyczny"
Public Const PresentationDownloadedEvent = "Pobrano-prezentacje"
Public Const SubjectDeletedEvent = "Usuniêto-przedmiot"
Public Const SubjectEditedEvent = "Edytowano-przedmiot"
Public Const ClassAddedEvent = "Dodano-zajecia"
Public Const ClassEditedEvent = "Edytowano-zajecia"
Public Const ClassDeletedEvent = "Usuniêto-zajêcia"
Public Const NotesDownloadedEvent = "Pobrano-notatki"
Public Const StudentsImportedEvent = "Zaimportowano-studentów"
Public Const GroupAddedEvent = "Dodano-grupe-zajeciowa"
Public Const StudentAddedToGroup = "Dodano-studenta-do-grupy"
Public Const FinishedClassSaved = "Zrealizowano-zajecia-zapisano"

Public Sub EventGateway(toSend As zdarzenie)
    Dim rs As Recordset
    
    Set rs = CurrentDb.OpenRecordset("Zdarzenia", dbOpenDynaset)
    With rs
        .AddNew
            rs!Data = Now
            rs!Nazwa = toSend.name
            rs!Wiadomosc = toSend.Message
            rs!CzyPrzetworzono = toSend.isProcessed
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
    wpis.CallStac = "Dispatch event while dispatching " & eventRs!name
    Logger.Add wpis
End Sub

Private Sub MarkEventAsProcessed(rs As DAO.Recordset)
    rs.Edit
        rs!Processed = True
        rs!ProcessedAt = Now
    rs.Update
End Sub

Public Function EventFactory(Opis As String, name As String, isProcessed As Boolean) As zdarzenie
    Dim e As zdarzenie
    e.Message = Opis
    e.name = name
    e.ProcessingDate = Now
    e.isProcessed = isProcessed
    EventFactory = e
End Function
