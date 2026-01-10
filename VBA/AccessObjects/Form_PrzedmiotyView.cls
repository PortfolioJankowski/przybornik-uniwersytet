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

Private Enum SubjectCols
    Id = 0
    Name = 1
    desc = 2
    SylabusId = 3
End Enum


Private Sub btnEdit_Click()
    Dim desc As String
    desc = txtEditDesc.value
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
End Sub

Private Sub lbxSubjects_AfterUpdate()
    idSubject = CLng(Me.lbxSubjects.value)
    FillEditTab
    HandleGotoActivation
End Sub

Private Sub FillEditTab()
    tabCtrl.value = 1
    txtEditDesc.value = Me.lbxSubjects.Column(SubjectCols.desc)
    txtEditName.value = Me.lbxSubjects.Column(SubjectCols.Name)
    cmbEditSylabus.value = Me.lbxSubjects.Column(SubjectCols.SylabusId)
    idSylabus = CLng(Me.lbxSubjects.Column(SubjectCols.SylabusId))
End Sub


Public Sub HandleGotoActivation()
    btnGoToSylabus.Enabled = (idSylabus <> 0)
End Sub
