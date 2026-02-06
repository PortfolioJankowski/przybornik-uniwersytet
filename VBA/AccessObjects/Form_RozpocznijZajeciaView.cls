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
Public actualClassId As String
Public classId As String
Public cycleId As String
Dim isFinished As Boolean

Private Sub btnDownloadNotes_Click()
    FileRepo.DownloadNotes (CLng(classId))
End Sub

'CLASS ID DO ZNALEZIENIA
Private Sub btnDownloadPres_Click()
    FileRepo.DownloadPres (CLng(classId))
End Sub

Private Sub btnExit_Click()
    DoCmd.Close acForm, Consts.RozpocznijZajeciaView
End Sub


Public Sub Initialize()
    groupId = txtGroupId.value
    
    Dim rsZajecia As DAO.Recordset
    Dim rsRealized As DAO.Recordset
    Dim lastCompletedID As Variant
    Dim currentClassNo As Long
    Dim studentsQuery As String

    Set rsRealized = CurrentDb.OpenRecordset( _
        "SELECT id, ZajeciaId, CzyZrealizowano, Notatki FROM ZrealizowaneZajecia " & _
        "WHERE GrupyZajecioweId = " & groupId & " AND CzyZrealizowano = False", dbOpenDynaset)

    If Not (rsRealized.EOF And rsRealized.BOF) Then
        actualClassId = CStr(rsRealized!id)
        currentClassNo = rsRealized!ZajeciaId
        isFinished = False
    Else
        lastCompletedID = DMax("ZajeciaId", "ZrealizowaneZajecia", _
                               "GrupyZajecioweId = " & groupId & " AND CzyZrealizowano = True")
        
        If IsNull(lastCompletedID) Then
            currentClassNo = 1
        Else
            currentClassNo = CLng(lastCompletedID) + 1
        End If

        rsRealized.AddNew
        rsRealized!GrupyZajecioweId = groupId
        rsRealized!ZajeciaId = currentClassNo
        rsRealized!CzyZrealizowano = False
        rsRealized!Notatki = ""
        rsRealized.Update
        
        rsRealized.Bookmark = rsRealized.LastModified
        actualClassId = CStr(rsRealized!id)
        isFinished = False
    End If
    
    classNo = CStr(currentClassNo)
    txtIdZrealizowaneZajecia = actualClassId
    Dim classInfoQuery As String
    classInfoQuery = "SELECT * FROM Zajecia WHERE CyklDydaktycznyId IN (" & _
        "SELECT CyklDydaktycznyId FROM Przedmioty WHERE Identyfikator IN (" & _
        "SELECT PrzedmiotId FROM GrupyZajeciowe WHERE Identyfikator = " & groupId & "))" & _
        " AND Kolejnosc = " & currentClassNo
    
    Set rsZajecia = CurrentDb.OpenRecordset(classInfoQuery, dbOpenSnapshot)
    
    If rsZajecia.EOF And rsZajecia.BOF Then
        MsgBox "B³¹d: Nie odnaleziono definicji zajêæ nr " & currentClassNo & " dla tej grupy.", vbCritical
        Exit Sub
    End If

    txtDesc = rsZajecia!Opis
    txtTitle = rsZajecia!Tytul
    txtNumber = rsZajecia!Kolejnosc
    txtNumber.enabled = False
    
    If Not IsNull(rsRealized) Then
         txtNotes.value = rsRealized!Notatki
    End If
    
    classId = CStr(rsZajecia!Identyfikator)
    cycleId = CStr(rsZajecia!CyklDydaktycznyId)

    studentsQuery = "SELECT Identyfikator, NrAlbumu, Imie, Nazwisko FROM Studenci " & _
                    "WHERE Identyfikator IN (SELECT StudentId FROM StudenciWGrupach " & _
                    "WHERE GrupaId = " & groupId & ")"

    With Me.lbxPresent
        .ColumnCount = 4
        .BoundColumn = 1
        .RowSource = studentsQuery
        .Requery
    End With

    rsRealized.Close
    rsZajecia.Close
    Set rsRealized = Nothing
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

