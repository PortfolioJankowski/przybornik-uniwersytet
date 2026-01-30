VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LandingView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnDirections_Click()
    DoCmd.OpenForm "KierunkiView"
End Sub

Private Sub btnGroups_Click()
    DoCmd.OpenForm Consts.GrupyZajecioweView
End Sub

Private Sub btnPrzedmioty_Click()
    DoCmd.OpenForm "PrzedmiotyView"
End Sub

Private Sub Command14_Click()
    DoCmd.OpenForm Consts.SylabusView
End Sub

Private Sub Form_Load()
If App Is Nothing Then
    Exit Sub
End If
    With Me
        .txtName.value = App.CurrentUser.Imie
        .txtDate.value = "Dzisiaj jest " & CStr(Date)
    End With
    
   ZaladujSchedule
End Sub

Private Sub btnGoToStudents_Click()
    DoCmd.OpenForm Consts.StudentsView
End Sub

Private Sub ZaladujSchedule()

    Dim sql As String

    sql = "SELECT " & _
      "G.Identyfikator AS GrupaId, " & _
      "P.Nazwa & ' - ' & K.Nazwa & ' - ' & G.Rok & " & _
      "' - Grupa: ' & G.Numer & ' : ' & " & _
      "Switch(" & _
          "G.DzienTygodnia=1,'Poniedzia³ek'," & _
          "G.DzienTygodnia=2,'Wtorek'," & _
          "G.DzienTygodnia=3,'Œroda'," & _
          "G.DzienTygodnia=4,'Czwartek'," & _
          "G.DzienTygodnia=5,'Pi¹tek'," & _
          "G.DzienTygodnia=6,'Sobota'," & _
          "G.DzienTygodnia=7,'Niedziela'" & _
      ") & ' ' & G.Godzina AS Opis " & _
      "FROM (GrupyZajeciowe AS G " & _
      "INNER JOIN Przedmioty AS P ON G.PrzedmiotId = P.Identyfikator) " & _
      "INNER JOIN Kierunki AS K ON G.KierunekId = K.Identyfikator " & _
      "WHERE G.CzyAktywna = True " & _
      "ORDER BY G.DzienTygodnia, G.Godzina;"

    With Me.lbxSchedule
        .RowSourceType = "Table/Query"
        .RowSource = sql
        .ColumnCount = 2
        .ColumnWidths = "0cm;15cm"
        .BoundColumn = 1
    End With

End Sub
