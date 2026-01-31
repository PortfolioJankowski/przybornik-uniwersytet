VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_RozpocznijZajeciaView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public groupId As String
Public classId As String

Private Sub btnExit_Click()
    DoCmd.Close acForm, Consts.RozpocznijZajeciaView
End Sub

Public Sub SetUp(id As String)
    groupId = id
End Sub

Private Sub Form_Load()
    groupId = Me.OpenArgs
End Sub

Public Sub Initialize()

    Dim lastClassQuery As String
    Dim classInfoQuery As String
    Dim studentsQuery As String

    Dim lastClassId As Variant
    Dim actualClassNo As Long

    ' --- 1. OSTATNIE ZREALIZOWANE ZAJÊCIA ---
    lastClassQuery = _
        "SELECT TOP 1 ZajeciaId " & _
        "FROM ZrealizowaneZajecia " & _
        "WHERE GrupyZajecioweId = " & groupId & " " & _
        "ORDER BY ZajeciaId DESC"

    lastClassId = DLookup("ZajeciaId", _
                          "ZrealizowaneZajecia", _
                          "GrupyZajecioweId = " & groupId)

    If IsNull(lastClassId) Then
        actualClassNo = 1
    Else
        actualClassNo = CLng(lastClassId) + 1
    End If

    classId = actualClassNo

    ' --- 2. INFO O AKTUALNYCH ZAJÊCIACH ---
    classInfoQuery = _
        "SELECT * FROM Zajecia " & _
        "WHERE Identyfikator = " & classId

    ' (tu mo¿esz OpenRecordset / DLookup i przypisaæ pola do kontrolek)

    ' --- 3. LISTA STUDENTÓW Z GRUPY ---
    studentsQuery = _
        "SELECT * FROM Studenci " & _
        "WHERE Identyfikator IN (" & _
            "SELECT StudentId " & _
            "FROM StudenciWGrupach " & _
            "WHERE GrupaId = " & groupId & _
        ")"

    ' przyk³ad: przypisanie do listboxa
    Me.lbxPresent.RowSource = studentsQuery
    Me.lbxPresent.Requery

End Sub

Private Sub Form_Open(Cancel As Integer)

End Sub
