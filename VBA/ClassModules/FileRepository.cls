VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private tableName As String

Public Sub Init(tblName As String)
    If Len(Trim(tblName)) = 0 Then
        err.Raise 67, "FileRepository.Init", "Nazwa tabeli nie mo¿e byæ pusta!"
    End If
    tableName = tblName
End Sub

Private Function NazwaKolumny(typPliku As PrzechowywanyPlik) As String
    Select Case typPliku
        Case PrzechowywanyPlik.Prezentacja
            NazwaKolumny = "Prezentacja"
        Case PrzechowywanyPlik.Notatki
            NazwaKolumny = "Notatki"
        Case Else
            err.Raise 5, "FileRepository", "Nieznany typ pliku"
    End Select
End Function


Public Sub PobierzPlikZBazy( _
        ByVal IDRekordu As Long, _
        typPliku As PrzechowywanyPlik, _
        SciezkaFolderu As String)

    Dim db As DAO.database
    Dim rs As DAO.Recordset
    Dim rsAttach As DAO.Recordset2
    Dim kolumna As String
    Dim pelnaSciezka As String
    Dim nazwaPliku As String

    kolumna = NazwaKolumny(typPliku)
    Set db = CurrentDb

    Set rs = db.OpenRecordset( _
        "SELECT [" & kolumna & "] FROM [" & tableName & "] WHERE Identyfikator = " & IDRekordu, _
        dbOpenDynaset)

    If rs.EOF Then err.Raise 3021, , "Nie znaleziono rekordu"

    Set rsAttach = rs.Fields(kolumna).value

    If rsAttach.RecordCount = 0 Then
        err.Raise 5, , "Brak za³¹cznika"
    End If

    rsAttach.MoveFirst

    nazwaPliku = rsAttach.Fields("FileName").value

    If Right(SciezkaFolderu, 1) <> "\" Then
        SciezkaFolderu = SciezkaFolderu & "\"
    End If

    pelnaSciezka = SciezkaFolderu & nazwaPliku

    rsAttach.Fields("FileData").SaveToFile pelnaSciezka

    rsAttach.Close
    rs.Close

End Sub

Public Sub WgrajPlikDoBazy( _
    IDRekordu As Long, _
    typPliku As PrzechowywanyPlik, _
    SciezkaPliku As String)

    Dim db As DAO.database
    Dim rs As DAO.Recordset
    Dim rsAttach As DAO.Recordset2
    Dim kolumna As String

    If Dir(SciezkaPliku) = "" Then
        err.Raise 53, , "Plik nie istnieje"
    End If

    kolumna = NazwaKolumny(typPliku)
    Set db = CurrentDb

    Set rs = db.OpenRecordset( _
        "SELECT * FROM [" & tableName & "] WHERE Identyfikator=" & IDRekordu, _
        dbOpenDynaset)

    If rs.EOF Then err.Raise 3021, , "Nie znaleziono rekordu"

    rs.Edit
    Set rsAttach = rs.Fields(kolumna).value

    If rsAttach.RecordCount > 0 Then
        rsAttach.MoveFirst
        Do Until rsAttach.EOF
            rsAttach.Delete
            If Not rsAttach.EOF Then rsAttach.MoveNext
        Loop
    End If

    rsAttach.AddNew
    rsAttach.Fields("FileData").LoadFromFile SciezkaPliku
    rsAttach.Update

    rs.Update
 
    rs.Close
End Sub

