VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_StudenciView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnExit_Click()
    DoCmd.Close acForm, Consts.StudentsView
End Sub

Private Sub btnMassImport_Click()
    Dim path As String
    Dim fso As FileDialog
    MsgBox "Aby zaimportowaæ studentów wska¿ plik Excela z kolumnami: Imie, Nazwisko, NrAlbumu"
    On Error GoTo catch
    
    Set fso = FileDialog(msoFileDialogFilePicker)
    
    With fso
        .Title = "Wska¿ listê studentów"
        .Show
        
        ImportujZExcela .SelectedItems(1)
    End With

    Dim e As Zdarzenie
        e.IsProcessed = True
        e.Message = "Pomyœlnie zaimportowano studentów " & Now
        e.Name = EventMdl.StudentsImportedEvent
    EventGateway e
    
    MsgBox "Zaimportowano studentów pomyœlnie"
    DoCmd.Close acForm, Consts.StudentsView
    Exit Sub
catch:
    Dim wpis As wpis
    wpis.CallStac = "Importowanie studentów z excela"
    wpis.Description = err.Description
    wpis.ErrorNumber = err.Number
    App.Logger.Add wpis
End Sub


Private Sub ImportujZExcela(sciezkaDoExcela As String)

    Dim db As DAO.database
    Dim rs As DAO.Recordset
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWS As Object
    Dim ostatniWiersz As Long
    Dim i As Long
    Dim NrAlbumu As Variant
    Dim Imie As String
    Dim Nazwisko As String

    Set db = CurrentDb
    Set rs = db.OpenRecordset("Studenci", dbOpenDynaset)


    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Open(sciezkaDoExcela)
    Set xlWS = xlWB.Worksheets(1)


    ostatniWiersz = xlWS.Cells(xlWS.Rows.Count, 1).End(-4162).Row '


    For i = 2 To ostatniWiersz
        NrAlbumu = xlWS.Cells(i, 3).value
        Imie = xlWS.Cells(i, 1).value
        Nazwisko = xlWS.Cells(i, 2).value

        If DCount("*", "Studenci", "NrAlbumu = '" & NrAlbumu & "'") = 0 Then
            rs.AddNew
            rs!NrAlbumu = NrAlbumu
            rs!Imie = Imie
            rs!Nazwisko = Nazwisko
            rs.Update
        End If

    Next i


    rs.Close
    xlWB.Close False
    xlApp.Quit

    Set rs = Nothing
    Set db = Nothing
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing



End Sub
