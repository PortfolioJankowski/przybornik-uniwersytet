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
Dim isFinished As Boolean

Private Sub btnExit_Click()
    DoCmd.Close acForm, Consts.RozpocznijZajeciaView
End Sub


Public Sub Initialize()
    groupId = txtGroupId.value
    isFinished = False 'zajecia jeszcze nie zrealizowane
    
    Dim lastClassQuery As String
    Dim classInfoQuery As String
    Dim studentsQuery As String

    Dim lastClassId As Variant
    Dim actualClassNo As Long

    ' --- 1. OSTATNIE ZREALIZOWANE ZAJÊCIA ---
    lastClassQuery = _
        "SELECT TOP 1 ZajeciaId " & _
        "FROM ZrealizowaneZajecia " & _
        "WHERE GrupyZajecioweId = " & groupId & " " & " CzyZrealizowano=False" & _
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
    rsNewClass!ZajeciaId = classId
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

