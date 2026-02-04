VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PrzedmiotyView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim idSubject As Long
Dim idSylabus As Long
Dim idCycle As Long
Dim idClass As Long

Private Enum SubjectCols
    id = 0
    name = 1
    desc = 2
    SylabusId = 3
    CycleId = 4
End Enum

Private Sub HandleSubjectSelection()
    If IsNull(Me.lbxSubjects.value) Then Exit Sub

    idSubject = CLng(Me.lbxSubjects.value)
    FillEditTab
    HandleGotoActivation
    HandleSubformsVisiblity
    ChangeClassesVisibility
    LoadClasses
End Sub

Private Sub ChangeClassesVisibility()
    Dim visibility As Boolean
    visibility = (idCycle <> 0)
    With Me
        .lbxClasses.Visible = visibility
        .tabClasses.Visible = visibility
        .btnDeleteClass.Visible = visibility
        .lblCycle.Visible = visibility
        .lblCycle2.Visible = visibility
        .imgCycle.Visible = visibility
        .btnGoToClasses.Visible = visibility

    End With
End Sub
Private Sub LoadClasses()
    If (idCycle = 0) Then
        lbxClasses.RowSource = ""
        lbxClasses.AddItem ";"
        Me.Refresh
        Exit Sub
    End If
    
    lbxClasses.RowSourceType = "Value List"
    lbxClasses.ColumnCount = 2
    lbxClasses.BoundColumn = 1
    lbxClasses.ColumnWidths = "0cm;5cm"
    
    lbxClasses.RowSource = ""
    
    Dim rsClasses As Recordset
    Set rsClasses = CurrentDb.OpenRecordset(SqlStringService.GetClassesByCycleId(idCycle), dbOpenSnapshot)
    
    If rsClasses.EOF Or rsClasses.BOF Then
        Exit Sub
    End If
    
    Dim header As String
    While Not rsClasses.EOF
        header = rsClasses!Kolejnosc & ". " & rsClasses!Tytul
        
       
        lbxClasses.AddItem rsClasses!Identyfikator & ";" & header
        
        rsClasses.MoveNext
    Wend
    
    rsClasses.Close
    Set rsClasses = Nothing
End Sub


Private Sub FillEditTab()
    tabCtrl.value = 1
    txtEditDesc.value = Me.lbxSubjects.Column(2)
    txtEditName.value = Me.lbxSubjects.Column(1)
    cmbEditSylabus.value = Me.lbxSubjects.Column(3)
    idSylabus = CLng(Me.lbxSubjects.Column(3))
    idCycle = CLng(Me.lbxSubjects.Column(4))
End Sub

Private Sub ClearEditForm()
    With Me
        .txtEditDesc.value = ""
        .txtEditName.value = ""
        .cmbEditSylabus.value = ""
    End With
End Sub

Private Sub ClearAddForm()
     With Me
        .txtAddName.value = ""
        .txtAddDesc.value = ""
        .cmbAddSylabus.value = ""
    End With
End Sub

Private Sub HandleGotoActivation()
    btnGoToSylabus.enabled = (idSylabus <> 0)
    btnGoToClasses.Visible = (idClass <> 0)
End Sub

Private Sub HandleSubformsVisiblity()
    Dim cycleVis As Boolean: cycleVis = (idCycle <> 0)
    Dim sylabusVis As Boolean: sylabusVis = (idSylabus <> 0)
    
    srSylabusy.Visible = sylabusVis

    
    lblCycle.Visible = cycleVis
    lblCycle2.Visible = cycleVis
    lbxClasses.Visible = cycleVis
    tabClasses.Visible = cycleVis
    imgCycle.Visible = cycleVis
    btnDeleteClass.Visible = cycleVis
    btnGoToSylabus.Visible = sylabusVis
End Sub

Private Sub btnAdd_Click()
    On Error GoTo catch
    Dim db As DAO.database
    Dim rsCykl As DAO.Recordset
    Dim rsPrzedmiot As DAO.Recordset
    Dim lngCyklID As Long

    ' Walidacja danych
    If IsNull(Me.txtAddDesc) Or Me.txtAddDesc = "" Then
        MsgBox "Podaj opis przedmiotu.", vbExclamation
        Exit Sub
    End If

    If IsNull(Me.txtAddName) Or Me.txtAddName = "" Then
        MsgBox "Podaj nazwê przedmiotu.", vbExclamation
        Exit Sub
    End If
    
    If IsNull(Me.cmbAddSylabus) Or Me.cmbAddSylabus = "" Then
        MsgBox "Podaj sylabus przedmiotu.", vbExclamation
        Exit Sub
    End If

    Set db = CurrentDb
    Set rsCykl = db.OpenRecordset(Consts.CyklDydaktycznyTable, dbOpenDynaset)

    rsCykl.AddNew
        rsCykl!Nazwa = Me.txtAddName
    rsCykl.Update

    Dim addCycleEvent As zdarzenie
    With addCycleEvent
        .isProcessed = True
        .Message = "Dodano cykl dydaktyczny o nazwie " & Me.txtAddName.value
        .ProcessingDate = DateTime.Now
    EventGateway addCycleEvent
    End With
    
    ' Pobranie ID nowo dodanego cyklu
    rsCykl.Bookmark = rsCykl.LastModified
    lngCyklID = rsCykl!id

    rsCykl.Close
    Set rsCykl = Nothing
    Set rsPrzedmiot = db.OpenRecordset(Consts.PrzedmiotyTable, dbOpenDynaset)

    rsPrzedmiot.AddNew
        rsPrzedmiot!Nazwa = Me.txtAddName
        rsPrzedmiot!Opis = Nz(Me.txtAddDesc, "")
        rsPrzedmiot!SylabusId = CInt(Me.cmbAddSylabus.value)
        rsPrzedmiot!CyklDydaktycznyId = lngCyklID
    rsPrzedmiot.Update

    'Refresh form
    DoEvents
    Me.Requery
    DoEvents
    
    DoEvents
    Me.lbxSubjects.Requery
    DoEvents
    
    Dim lngPrzedmiotID As Long
    rsPrzedmiot.Bookmark = rsPrzedmiot.LastModified
    lngPrzedmiotID = rsPrzedmiot!Identyfikator
    
    ' Check Lbx
    Dim i As Integer
    For i = 0 To Me.lbxSubjects.ListCount - 1
        Me.lbxSubjects.Selected(i) = False
        If Me.lbxSubjects.Column(0, i) = CStr(lngPrzedmiotID) Then ' kolumna 0 to ID
            Me.lbxSubjects.Selected(i) = True
            HandleSubjectSelection
            ClearAddForm
            Exit For
        End If
    Next i
    
    Dim addSubjectEvent As zdarzenie
    With addSubjectEvent
        .isProcessed = True
        .Message = "Dodano przedmiot " & txtAddName
        .name = EventMdl.CycleAddedEvent
        .ProcessingDate = DateTime.Now
    End With
    EventGateway addSubjectEvent
    
    rsPrzedmiot.Close
    Set rsPrzedmiot = Nothing
    Set db = Nothing
    
    MsgBox "Przedmiot zosta³ dodany." & vbNewLine & "Mo¿esz przejsæ do tworzenia cyklu dydaktycznego", vbInformation

catch:
    Dim wpis As wpis
    wpis.CallStac = err.Source
    wpis.Description = "B³¹d podczas dodawania przedmiotu " & Me.txtAddName
    wpis.ErrorNumber = err.Number
    wpis.LogDate = DateTime.Now
    App.Logger.Add wpis
End Sub

Private Sub ClearEditClassForm()
    With Me
        .cmbEditOrder.value = ""
        .txtEditClassDesc.value = ""
        .txtEditClassTitle.value = ""
    End With
End Sub
Private Sub HandleDeleteClassBtnActivation()
    Dim enabled As Boolean
    enabled = idClass <> 0
    Me.btnDeleteClass.enabled = enabled
End Sub






Private Sub btnAddClass_Click()
On Error GoTo catch
    If idCycle = 0 Then Exit Sub
    If idSubject = 0 Then Exit Sub
    
    If IsNullOrWhiteSpace(txtAddTitle.value) Then
        err.Raise 2, "B³¹d walidacji", "Tytu³ zajêæ nie mo¿e byæ pusty"
    End If
    
    If IsNullOrWhiteSpace(txtAddDescription.value) Then
        err.Raise 2, "B³¹d walidacji", "Opis zajêæ nie mo¿e byæ pusty"
    End If
    
    If IsNullOrWhiteSpace(cmbAddOrder.value) Then
        err.Raise 2, "B³¹d walidacji", "Nale¿y ustaliæ kolejnoœæ zajêæ"
    End If
    
    Dim FileRepo As FileRepository
    Set FileRepo = New FileRepository
    FileRepo.Init Consts.ZajeciaTable
    
    If (IsNullOrWhiteSpace(txtAddFileName.value) Or IsNullOrWhiteSpace(txtAddNotes.value)) Then
        Dim odpowiedz As Byte
        odpowiedz = MsgBox("Dodajesz zajecia bez prezentacji lub notatek. Czy chcesz kontynuowaæ?", vbYesNo)
        If odpowiedz = vbNo Then
            Exit Sub
        End If
    End If
    
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("Zajecia", dbOpenDynaset)

    rs.AddNew
    rs!CyklDydaktycznyId = idCycle
    rs!Tytul = txtAddTitle.value
    rs!Opis = txtAddDescription.value
    rs!Kolejnosc = cmbAddOrder.value
    rs.Update

    rs.Bookmark = rs.LastModified
    Dim idZajecia As Long
    idZajecia = rs!Identyfikator
    rs.Close
    
    If Not IsNullOrWhiteSpace(txtAddFileName.value) Then
        FileRepo.WgrajPlikDoBazy idZajecia, Prezentacja, txtAddFileName.value
    End If
    
     If Not IsNullOrWhiteSpace(txtAddNotes.value) Then
        FileRepo.WgrajPlikDoBazy idZajecia, Prezentacja, txtAddFileName.value
    End If
    
    Dim e As zdarzenie
    e.isProcessed = True
    e.Message = "Dodano zajêcia do cyklu " & CStr(idCycle) & _
    " o tytule: " & txtAddDescription.value
    e.name = EventMdl.ClassAddedEvent
    
    EventGateway e

    HandleSubjectSelection
    
    txtAddDescription.value = ""
    txtAddFileName.value = ""
    txtAddNotes.value = ""
    txtAddTitle.value = ""
    cmbAddOrder.value = ""
    
  

    MsgBox "Zajecia zosta³y dodane!"
    Exit Sub
catch:

    If err.Number = 2 Then
        MsgBox err.Description
    End If
    
    Dim log As wpis
    log.CallStac = "Dodawanie zajêcia dla cyklu " & idCycle
End Sub

'============== Event Handlers =========================
'Subject Removing
Private Sub btnDelete_Click()
    If idSubject = 0 Then Exit Sub

    On Error GoTo catch

    Dim query As String
    query = DeleteSubjectById(idSubject)

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        SqlStringService.GetSubjectById(idSubject), _
        dbOpenSnapshot)

    If rs.EOF Then
        err.Raise 3021, , "Nie znaleziono przedmiotu"
    End If

    Dim subjectName As String
    subjectName = rs!Nazwa

    rs.Close
    Set rs = Nothing

    CurrentDb.Execute query, dbFailOnError

    Dim e As zdarzenie
    e.name = EventMdl.SubjectDeletedEvent
    e.Message = "Usuniêto przedmiot " & subjectName
    e.isProcessed = True
    e.ProcessingDate = Now

    EventGateway e
    
    If Me.lbxSubjects.ListCount > 0 Then
        Me.lbxSubjects.Selected(0) = True
        HandleSubjectSelection
    End If

    MsgBox "Pomyœlnie usuniêto przedmiot"
    DoCmd.Close acForm, Consts.SubjectView
    Exit Sub

catch:
    Dim log As wpis
    log.Description = err.Description
    log.ErrorNumber = err.Number
    App.Logger.Add log

    MsgBox "Nie uda³o siê usun¹æ przedmiotu"
End Sub


'Class Removing
Private Sub btnDeleteClass_Click()
    If (idClass = 0) Then Exit Sub
    
    Dim confirmation As String
    confirmation = MsgBox("Czy na pewno chcesz usun¹æ wybrane zajêcia?", vbYesNo)
    If Not (confirmation = vbYes) Then Exit Sub
    
    Dim query As String: query = SqlStringService.DeleteClassById(idClass)
    On Error GoTo errorhandler
    CurrentDb.Execute query
    HandleSubjectSelection
    tabClasses.value = 0
    
    Exit Sub
errorhandler:
    MsgBox "Wystapi³ nieoczekiwany b³¹d", vbCritical
End Sub

Private Sub btnDownloadNotes_Click()
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

Private Sub btnDownloadPres_Click()
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

'Subject Editing
Private Sub btnEdit_Click()
    On Error GoTo catch

    If idSubject = 0 Then Exit Sub

    ' Walidacja danych
    If IsNull(Me.txtEditDesc) Or Me.txtEditDesc = "" Then
        MsgBox "Podaj opis przedmiotu.", vbExclamation
        Exit Sub
    End If

    If IsNull(Me.txtEditName) Or Me.txtEditName = "" Then
        MsgBox "Podaj nazwê przedmiotu.", vbExclamation
        Exit Sub
    End If
    
    If IsNull(Me.cmbEditSylabus) Or Me.cmbEditSylabus = "" Then
        MsgBox "Podaj sylabus przedmiotu.", vbExclamation
        Exit Sub
    End If
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset(SqlStringService.GetSubjectById(idSubject), dbOpenDynaset)
    
    If rs.EOF Then err.Raise 2, , "Brak przedmiotu"
    rs.Edit
    rs!Nazwa = Me.txtEditName
    rs!Opis = Me.txtEditDesc
    rs!SylabusId = Me.cmbEditSylabus
    rs.Update
    
    rs.Close
    Set rs = Nothing
    
    Dim e As zdarzenie
    With e
        .name = EventMdl.SubjectEditedEvent
        .Message = "Edytowano przedmiot " & Me.txtEditName.value
        .ProcessingDate = Now
        .isProcessed = True
    End With
    
    EventGateway e
    
    MsgBox "Pomyœlnie edytowano przedmiot"
    Me.lbxSubjects.Requery
    
    Exit Sub

catch:
    Dim log As wpis
    log.ErrorNumber = err.Number
    log.Description = err.Description
    App.Logger.Add log
End Sub


Private Sub btnEditClassSAve_Click()
    If idSubject = 0 Then Exit Sub
    If idCycle = 0 Then Exit Sub
    If idClass = 0 Then Exit Sub
    On Error GoTo catch
    
    Dim sourceText As String: sourceText = "B³¹d walidacji"
    
    If IsNullOrWhiteSpace(txtEditClassDesc.value) Then
        err.Raise 2, sourceText, "Opis zajêæ nie mo¿e byæ pusty"
    End If
    
    If IsNullOrWhiteSpace(txtEditClassTitle.value) Then
        err.Raise 2, sourceText, "Tytu³ zajêæ nie mo¿e byæ pusty"
    End If
        
    If IsNullOrWhiteSpace(cmbEditOrder.value) Then
        err.Raise 2, sourceText, "Wska¿ kolejnoœæ zajêæ"
    End If
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset(SqlStringService.GetClassById(idClass), dbOpenDynaset)
    
    rs.Edit
    rs!Tytul = txtEditClassTitle.value
    rs!Opis = txtEditClassDesc.value
    rs!Kolejnosc = cmbEditOrder.value
    rs.Update
    
    
    Dim e As zdarzenie
    e.name = EventMdl.ClassEditedEvent
    e.Message = "Eytowano zajêcia o identyfikatorze: " & idClass
    e.isProcessed = True
    e.ProcessingDate = Now
    
    EventGateway e
    
    MsgBox "Edytowano zajêcia"
    Exit Sub
catch:
    If err.Number = 2 Then
        MsgBox err.Description
    End If
End Sub

Private Sub btnExit_Click()
    DoCmd.Close acForm, Consts.SubjectView
End Sub

Private Sub btnGoToClasses_Click()
 If idClass = 0 Then
        DoCmd.OpenForm Consts.ClassView, , , , acFormAdd
    Else
        DoCmd.OpenForm Consts.ClassView, , , _
            "Identyfikator=" & idClass
    End If
End Sub
Private Function GetSelectedFileName(Title As String) As String
 Dim fs As FileDialog
    Set fs = FileDialog(msoFileDialogOpen)
    
    With fs
        .Title = Title
        .Show
        
        Dim path As String
        path = .SelectedItems(1)

    End With
    
    Set fs = Nothing
    
    GetSelectedFileName = path
End Function
Private Sub btnUploadNotes_Click()
    Me.txtAddNotes.value = GetSelectedFileName("Wybierz notatki")
End Sub

Private Sub btnUploadPres_Click()
 Me.txtAddFileName.value = GetSelectedFileName("Wybierz prezentacje")
End Sub

'Tab changed
Private Sub tabClasses_Change()
    With tabClasses
        Dim addIndex As Integer: addIndex = 0
        If tabClasses.value = addIndex Then
            'remove selection
            idClass = 0
            Dim i As Integer
            For i = 0 To Me.lbxClasses.ListCount - 1
                Me.lbxClasses.Selected(i) = False
            Next i
            'clear edit class
            ClearEditClassForm
            HandleDeleteClassBtnActivation
        End If
    End With
End Sub

Private Sub lbxClasses_Click()
    With Me.lbxClasses
        If IsNull(.value) Then Exit Sub
        idClass = CLng(.value)
        HandleDeleteClassBtnActivation
        FillEditForm
    End With
End Sub

Private Sub FillEditForm()
    tabClasses.value = 1
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset(SqlStringService.GetClassById(idClass), dbOpenSnapshot)
    
    With rs
        Me.txtEditClassDesc = rs!Opis
        Me.txtEditClassTitle = rs!Tytul
        Me.cmbEditOrder = rs!Kolejnosc
    End With
    
End Sub

Private Sub btnGoToSylabus_Click()
     If idSylabus = 0 Then
        DoCmd.OpenForm Consts.SylabusView, , , , acFormAdd
    Else
        DoCmd.OpenForm Consts.SylabusView, , , _
            "Identyfikator=" & idSylabus
    End If
End Sub

Private Sub Form_Load()
    HandleGotoActivation
    HandleSubformsVisiblity
End Sub

Private Sub lbxSubjects_AfterUpdate()
    HandleSubjectSelection
End Sub

Private Sub tabCtrl_Change()
    If tabCtrl.value = Consts.AddSubjectTabIndex Then
             Dim i As Long
        For i = 0 To Me.lbxSubjects.ListCount - 1
            Me.lbxSubjects.Selected(i) = False
        Next i
        idSylabus = 0
        idSubject = 0
        idCycle = 0
        HandleGotoActivation
        HandleSubformsVisiblity
        ClearEditForm
        ChangeClassesVisibility
    End If
End Sub


