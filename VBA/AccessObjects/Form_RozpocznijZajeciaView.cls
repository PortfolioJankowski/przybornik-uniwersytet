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
Public classNo As String
Dim isFinished As Boolean

Private Sub btnDownloadNotes_Click()
    FileRepo.DownloadNotes (CLng(classNo))
End Sub

'CLASS ID DO ZNALEZIENIA
Private Sub btnDownloadPres_Click()
    FileRepo.DownloadPres (CLng(classNo))
    
End Sub

Private Sub btnExit_Click()
    DoCmd.Close acForm, Consts.RozpocznijZajeciaView
End Sub


Public Sub Initialize()
    groupId = txtGroupId.value
    isFinished = False 'zajecia jeszcze nie zrealizowane
    
    Dim lastClassQuery As String
    Dim classInfoQuery As String
    Dim studentsQuery As String

    Dim lastclassNo As Variant
    Dim actualClassNo As Long

    ' --- 1. OSTATNIE ZREALIZOWANE ZAJÊCIA ---
    lastClassQuery = _
        "SELECT TOP 1 ZajeciaId " & _
        "FROM ZrealizowaneZajecia " & _
        "WHERE GrupyZajecioweId = " & groupId & " " & " CzyZrealizowano=False" & _
        "ORDER BY ZajeciaId DESC"

    lastclassNo = DLookup("ZajeciaId", _
                          "ZrealizowaneZajecia", _
                          "GrupyZajecioweId = " & groupId)

    If IsNull(lastclassNo) Then
        actualClassNo = 1
    Else
        actualClassNo = CLng(lastclassNo) + 1
    End If

    classNo = actualClassNo

    ' --- 2. INFO O AKTUALNYCH ZAJÊCIACH ---

    Dim zajeciaQuery As String
    classInfoQuery = "SELECT * FROM Zajecia WHERE CyklDydaktycznyId in (" & _
    "SELECT CyklDydaktycznyId FROM Przedmioty Where Identyfikator in (" & _
    "SELECT PrzedmiotId From GrupyZajeciowe where Identyfikator = " & groupId & "))" & _
    " AND Kolejnosc = " & actualClassNo
    
     Dim rsZajecia As Recordset
    Set rsZajecia = CurrentDb.OpenRecordset(classInfoQuery, dbOpenSnapshot)
    
    If rsZajecia.EOF And rsZajecia.BOF Then
        err.Raise 2, , "Nie odnaleziono zajêæ"
    End If
    txtDesc = rsZajecia!Opis
    txtTitle = rsZajecia!Tytul
    txtNumber = rsZajecia!Kolejnosc
    txtNumber.enabled = False
    txtNotes = rsZajecia!Notatki

    ' --- 3. LISTA STUDENTÓW Z GRUPY ---
    studentsQuery = _
        "SELECT Identyfikator, NrAlbumu,  Imie, Nazwisko FROM Studenci " & _
        "WHERE Identyfikator IN (" & _
            "SELECT StudentId " & _
            "FROM StudenciWGrupach " & _
            "WHERE GrupaId = " & groupId & _
        ")"

    Me.lbxPresent.ColumnCount = 4
    Me.lbxPresent.BoundColumn = 1
    
    ' przyk³ad: przypisanie do listboxa
    Me.lbxPresent.RowSource = studentsQuery
    Me.lbxPresent.Requery
    
    Dim rsNewClass As DAO.Recordset
    Set rsNewClass = CurrentDb.OpenRecordset(Consts.ZrealizowaneZajeciaTable, dbOpenDynaset)
    
    rsNewClass.AddNew
    rsNewClass!GrupyZajecioweId = groupId
    rsNewClass!ZajeciaId = classNo
    rsNewClass!CzyZrealizowano = False
    rsNewClass!Notatki = ""
    rsNewClass.Update
    
    rsNewClass.Bookmark = rsNewClass.LastModified
    txtIdZrealizowaneZajecia = rsNewClass!id
    
    rsNewClass.Close
    Set rsNewClass = Nothing
    rsZajecia.Close
    Set rsZajecia = Nothing
End Sub

Private Sub btnMove_Click()
    Dim i As Variant
    
    If Me.lbxPresent.ItemsSelected.Count = 0 Then Exit Sub
    
    For Each i In Me.lbxPresent.ItemsSelected
        
        Me.lbxAbsent.AddItem _
            Me.lbxPresent.Column(0, i) & ";" & _
            Me.lbxPresent.Column(1, i) & ";" & _
            Me.lbxPresent.Column(2, i) & ";" & _
            Me.lbxPresent.Column(3, i)
            
    Next i
    

End Sub

Private Sub btnRemoveAbsent_Click()
    Dim i As Long
    
    If Me.lbxAbsent.ListCount = 0 Then Exit Sub
    
    For i = Me.lbxAbsent.ListCount - 1 To 0 Step -1
        If Me.lbxAbsent.Selected(i) Then
            Me.lbxAbsent.RemoveItem i
        End If
    Next i
End Sub

