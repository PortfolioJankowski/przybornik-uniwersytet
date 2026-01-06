Attribute VB_Name = "EventMdl"
Option Compare Database
Option Explicit

Public Const UserLoggedInEvent = "Zalogowano-uzytkownika"
Public Const UserRegisterdedEvent = "Zarejestrowano-uzytkownika"

Public Sub EventGateway(toSend As Zdarzenie)
    Dim rs As Recordset
    
    Set rs = App.db.OpenRecordset("Zdarzenia", dbOpenDynaset)
    With rs
        .AddNew
            rs!Data = toSend.EventDate
            rs!Nazwa = toSend.Name
            rs!Wiadomosc = toSend.Message
            rs!CzyPrzetworzono = False
        .Update
    End With
End Sub

