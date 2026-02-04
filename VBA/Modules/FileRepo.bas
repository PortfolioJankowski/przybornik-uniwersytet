Attribute VB_Name = "FileRepo"
Option Compare Database
Option Explicit

Public Sub DownloadNotes(idClass As Long)
    If (idClass = 0) Then Exit Sub
    
    Dim fso As FileDialog
    Set fso = Application.FileDialog(msoFileDialogFolderPicker)
    On Error GoTo catch
    
    Dim saveAsPath As String
    With fso
        .Title = "Wybierz lokalizacje zapisu"
        .InitialFileName = "Notatki"
        .Show
        
        saveAsPath = .SelectedItems(1)
    End With
    
    If saveAsPath = "" Then Exit Sub
    Dim FileRepo As New FileRepository
    FileRepo.Init "Zajecia"
    FileRepo.PobierzPlikZBazy idClass, Notatki, saveAsPath
    
    Dim e As zdarzenie
    e.isProcessed = False
    e.name = EventMdl.NotesDownloadedEvent
    e.Message = "Pobrano notatki dla zajêæ " & idClass
    EventGateway e
    
    MsgBox "Pobrano"
    Exit Sub
    
catch:
    If (err.Number = 3839) Then
        MsgBox "W wybranej lokalizacji znajduje siê ju¿ pobierany plik!"
        Dim fileExists As wpis
        fileExists.Description = "Pobierany plik ju¿ istnieje w wybranej lokalizacji"
        fileExists.LogDate = Now
        App.Logger.Add fileExists
        Exit Sub
    End If
    
    Dim log As wpis
    log.CallStac = "Id zajecia =" & idClass & " sciezka zapisu=" & saveAsPath
    log.ErrorNumber = err.Number
    log.LogDate = Now
    App.Logger.Add log
    
    MsgBox "Wyst¹pi³ nieoczekiwany b³¹d"
End Sub

Public Sub DownloadPres(idClass As Long)
    If (idClass = 0) Then Exit Sub
    
    Dim fso As FileDialog
    Set fso = Application.FileDialog(msoFileDialogFolderPicker)
    On Error GoTo catch
    
    Dim saveAsPath As String
    With fso
        .Title = "Wybierz lokalizacje zapisu"
        .InitialFileName = "Prezentacja"
        .Show
        
        saveAsPath = .SelectedItems(1)
    End With
    
    If saveAsPath = "" Then Exit Sub
    Dim FileRepo As New FileRepository
    FileRepo.Init "Zajecia"
    FileRepo.PobierzPlikZBazy idClass, Prezentacja, saveAsPath
    
    Dim e As zdarzenie
    e.isProcessed = False
    e.name = EventMdl.PresentationDownloadedEvent
    e.Message = "Pobrano prezentacje dla zajêæ " & idClass
    EventGateway e
    
    MsgBox "Pobrano"
    Exit Sub
    
catch:
    If (err.Number = 3839) Then
        MsgBox "W wybranej lokalizacji znajduje siê ju¿ pobierany plik!"
        Dim fileExists As wpis
        fileExists.Description = "Pobierany plik ju¿ istnieje w wybranej lokalizacji"
        fileExists.LogDate = Now
        App.Logger.Add fileExists
        Exit Sub
    End If
    
    Dim log As wpis
    log.CallStac = "Id zajecia =" & idClass & " sciezka zapisu=" & saveAsPath
    log.ErrorNumber = err.Number
    log.LogDate = Now
    App.Logger.Add log
    
    MsgBox "Wyst¹pi³ nieoczekiwany b³¹d"
End Sub
