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
    Id = 0
    Name = 1
    desc = 2
    SylabusId = 3
    cycleId = 4
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
    txtEditDesc.value = Me.lbxSubjects.Column(SubjectCols.desc)
    txtEditName.value = Me.lbxSubjects.Column(SubjectCols.Name)
    cmbEditSylabus.value = Me.lbxSubjects.Column(SubjectCols.SylabusId)
    idSylabus = CLng(Me.lbxSubjects.Column(SubjectCols.SylabusId))
    idCycle = CLng(Me.lbxSubjects.Column(SubjectCols.cycleId))
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

    Dim addCycleEvent As Zdarzenie
    With addCycleEvent
        .IsProcessed = True
        .Message = "Dodano cykl dydaktyczny o nazwie " & Me.txtAddName.value
        .ProcessingDate = DateTime.Now
    EventGateway addCycleEvent
    End With
    
    ' Pobranie ID nowo dodanego cyklu
    rsCykl.Bookmark = rsCykl.LastModified
    lngCyklID = rsCykl!Id

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
    
    Dim addSubjectEvent As Zdarzenie
    With addSubjectEvent
        .IsProcessed = True
        .Message = "Dodano przedmiot " & txtAddName
        .Name = EventMdl.CycleAddedEvent
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






'============== Event Handlers =========================
'Subject Removing
Private Sub btnDelete_Click()

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
    
    Exit Sub
errorhandler:
    MsgBox "Wystapi³ nieoczekiwany b³¹d", vbCritical
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
    
    
    
    Dim fileRepo As New FileRepository
    fileRepo.Init "Zajecia"
    fileRepo.PobierzPlikZBazy idClass, Prezentacja, saveAsPath
    
    Dim e As Zdarzenie
    e.IsProcessed = False
    e.Name = EventMdl.PresentationDownloadedEvent
    e.Message = "Pobrano prezentacje dla zajêæ " & idClass
    EventGateway e
    
    MsgBox "Pobrano"
    Exit Sub
    
catch:
    If (err.Number = 3839) Then
        MsgBox "W wybranej lokalizacji znajduje siê ju¿ pobierany plik!"
        Dim fileExists As wpis
        fileExists.Description "Pobierany plik ju¿ istnieje w wybranej lokalizacji"
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
    'TODO
End Sub

Private Sub btnGoToClasses_Click()
 If idClass = 0 Then
        DoCmd.OpenForm Consts.ClassView, , , , acFormAdd
    Else
        DoCmd.OpenForm Consts.ClassView, , , _
            "Identyfikator=" & idClass
    End If
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


