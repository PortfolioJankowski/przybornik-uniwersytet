VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DodajGrupeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnExit_Click()
    DoCmd.Close acForm, Consts.DodajGrupeView
End Sub

Private Sub btnSave_Click()
On Error GoTo catch
    If IsNullOrWhiteSpace(cmbSubject.value) Then
        err.Raise 2, , "Wybierz przedmiot"
    End If
    
    If IsNullOrWhiteSpace(cmbDay.value) Then
        err.Raise 2, , "Wybierz dzieñ tygodnia"
    End If
    
    If IsNullOrWhiteSpace(cmbDirection.value) Then
        err.Raise 2, , "Wybierz kierunek"
    End If
    
    If IsNullOrWhiteSpace(txtHour.value) Then
        err.Raise 2, , "Wpisz godzinê"
    End If
    
    If IsNullOrWhiteSpace(txtStudents.value) Then
        err.Raise 2, , "Wpisz numery albumów"
    End If
    
     If IsNullOrWhiteSpace(txtNumber.value) Then
        err.Raise 2, , "Wpisz numer grupy"
    End If
    
     If IsNullOrWhiteSpace(txtYear.value) Then
        err.Raise 2, , "Na którym roku jest grupa?"
    End If
    
    
    DodajGrupeZeStudentami cmbDirection.value, cmbSubject.value, txtNumber.value, txtYear.value, cmbDay.value, txtHour, txtStudents.value
    
    Dim e As zdarzenie
    With e
        .name = EventMdl.GroupAddedEvent
        .isProcessed = True
        .Message = "Pomyœlnie dodano grupê dla kierunku " & _
        cmbDirection.value & _
        " i przedmiotu " & _
        cmbSubject.value
    End With
    
    EventGateway e
    
    MsgBox "Grupa i studenci dodani poprawnie.", vbInformation
    DoCmd.Close acForm, Consts.DodajGrupeView
    Exit Sub
catch:
    If err.Number = 2 Then
        MsgBox err.Description
    End If
    
    Dim log As wpis
    With log
        .CallStac = "Dodawanie przedmiotu"
        .Description = err.Description
        .ErrorNumber = err.Number
        .LogDate = Now
    End With
    
    App.Logger.Add log
End Sub

Private Function SlownikDniTygodnia() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    dict.Add "poniedzia³ek", 1
    dict.Add "wtorek", 2
    dict.Add "œroda", 3
    dict.Add "czwartek", 4
    dict.Add "pi¹tek", 5
    dict.Add "sobota", 6
    dict.Add "niedziela", 7

    Set SlownikDniTygodnia = dict
End Function

Public Sub DodajGrupeZeStudentami( _
    KierunekId As Long, _
    PrzedmiotId As Long, _
    NumerGrupy As Integer, _
    Rok As Integer, _
    DzienTygodniaSlownie As String, _
    GodzinaZajec As String, _
    ListaNumerowAlbumow As String _
)

    Dim db As DAO.database
    Dim rsGrupa As DAO.Recordset
    Dim rsStudenci As DAO.Recordset
    Dim dictDni As Object
    Dim DzienNr As Integer
    Dim GrupaId As Long
    Dim arrAlbumy() As String
    Dim i As Long
    Dim NrAlbumu As String
    Dim studentId As Long

    Set db = CurrentDb
    Set dictDni = SlownikDniTygodnia

    ' --- dzieñ tygodnia ---
    DzienTygodniaSlownie = LCase(Trim(DzienTygodniaSlownie))

    If Not dictDni.Exists(DzienTygodniaSlownie) Then
        MsgBox "Niepoprawny dzieñ tygodnia!", vbCritical
        Exit Sub
    End If

    DzienNr = dictDni(DzienTygodniaSlownie)

    ' --- normalizacja godziny ---
    GodzinaZajec = GodzinaZajec

    ' =========================================================
    ' 1. DODANIE GRUPY
    ' =========================================================
    Set rsGrupa = db.OpenRecordset("GrupyZajeciowe", dbOpenDynaset)

    rsGrupa.AddNew
    rsGrupa!KierunekId = KierunekId
    rsGrupa!PrzedmiotId = PrzedmiotId
    rsGrupa!Numer = NumerGrupy
    rsGrupa!CzyAktywna = True
    rsGrupa!Rok = Rok
    rsGrupa!DzienTygodnia = DzienNr
    rsGrupa!Godzina = GodzinaZajec
    rsGrupa.Update

    rsGrupa.Bookmark = rsGrupa.LastModified
    GrupaId = rsGrupa!Identyfikator
    rsGrupa.Close

    ' =========================================================
    ' 2. DODANIE STUDENTÓW DO GRUPY
    ' =========================================================
    arrAlbumy = Split(ListaNumerowAlbumow, ",")

    For i = LBound(arrAlbumy) To UBound(arrAlbumy)

        NrAlbumu = Trim(arrAlbumy(i))

        ' pobranie StudentId
        studentId = Nz(DLookup("Identyfikator", "Studenci", _
                    "NrAlbumu = '" & NrAlbumu & "'"), 0)

        If studentId > 0 Then

            ' sprawdzenie czy ju¿ jest w grupie
            If DCount("*", "StudenciWGrupach", _
                "GrupaId = " & GrupaId & " AND StudentId = " & studentId) = 0 Then

                db.Execute _
                    "INSERT INTO StudenciWGrupach (GrupaId, StudentId) " & _
                    "VALUES (" & GrupaId & ", " & studentId & ")", _
                    dbFailOnError
            End If

        End If
    Next i

End Sub
