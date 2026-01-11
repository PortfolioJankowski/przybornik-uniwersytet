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

Private Enum SubjectCols
    Id = 0
    Name = 1
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
End Sub

Private Sub FillEditTab()
    tabCtrl.value = 1
    txtEditDesc.value = Me.lbxSubjects.Column(SubjectCols.desc)
    txtEditName.value = Me.lbxSubjects.Column(SubjectCols.Name)
    cmbEditSylabus.value = Me.lbxSubjects.Column(SubjectCols.SylabusId)
    idSylabus = CLng(Me.lbxSubjects.Column(SubjectCols.SylabusId))
    idCycle = CLng(Me.lbxSubjects.Column(SubjectCols.CycleId))
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
    btnGoToSylabus.Enabled = (idSylabus <> 0)
End Sub

Private Sub HandleSubformsVisiblity()
    If IsNull(idCycle) Then Exit Sub
    

    
    
End Sub

Private Sub btnAdd_Click()
    
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

    ' =========================
    ' Dodaj cykl dydaktyczny
    ' =========================
    Set rsCykl = db.OpenRecordset(Consts.CyklDydaktycznyTable, dbOpenDynaset)

    rsCykl.AddNew
        rsCykl!Nazwa = Me.txtAddName
    rsCykl.Update

    ' Pobranie ID nowo dodanego cyklu
    rsCykl.Bookmark = rsCykl.LastModified
    lngCyklID = rsCykl!Id

    rsCykl.Close
    Set rsCykl = Nothing

    ' =========================
    ' Dodaj przedmiot
    ' =========================
    Set rsPrzedmiot = db.OpenRecordset(Consts.PrzedmiotyTable, dbOpenDynaset)

    rsPrzedmiot.AddNew
        rsPrzedmiot!Nazwa = Me.txtAddName
        rsPrzedmiot!Opis = Nz(Me.txtAddDesc, "")
        rsPrzedmiot!SylabusId = CInt(Me.cmbAddSylabus.value)
        rsPrzedmiot!CyklDydaktycznyID = lngCyklID
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
    
    
    rsPrzedmiot.Close
    Set rsPrzedmiot = Nothing
    Set db = Nothing
    
    MsgBox "Przedmiot zosta³ dodany." & vbNewLine & "Mo¿esz przejsæ do tworzenia cyklu dydaktycznego", vbInformation

End Sub

'============== Event Handlers =========================

Private Sub btnEdit_Click()
    'TODO
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
    End If
End Sub


