VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Implements IEntityMapper

Public Function Map(rs As DAO.Recordset) As user
    Set Map = IEntityMapper_Map(rs)
End Function

Public Function ToInsertSql(u As user) As String
    ToInsertSql = BuildInsertSql(u)
End Function

Public Function ToUpdateSql(u As user) As String
    ToUpdateSql = BuildUpdateSql(u)
End Function


Private Function IEntityMapper_Map(rs As DAO.Recordset) As Object
    Dim u As New user

    u.id = rs!id
    u.NazwaUzytkownika = rs!NazwaUzytkownika
    u.HashHaslo = rs!HashHaslo
    u.Imie = rs!Imie
    u.Nazwisko = rs!Nazwisko
    u.AdresEmail = rs!AdresEmail
    u.Tytul = rs!Tytul

    Set IEntityMapper_Map = u
End Function

Private Function IEntityMapper_ToSqlInsert(entity As Object) As String
    Dim u As user
    Set u = entity

    IEntityMapper_ToSqlInsert = BuildInsertSql(u)
End Function

Private Function IEntityMapper_ToSqlUpdate(entity As Object) As String
    Dim u As user
    Set u = entity

    IEntityMapper_ToSqlUpdate = BuildUpdateSql(u)
End Function

Private Function BuildInsertSql(u As user) As String
    BuildInsertSql = _
        "INSERT INTO Uzytkownicy (" & _
        "NazwaUzytkownika, HashHaslo, Imie, Nazwisko, AdresEmail, Tytul" & _
        ") VALUES (" & _
        SqlStringService.GetSqlString(u.NazwaUzytkownika) & "," & _
        SqlStringService.GetSqlString(u.HashHaslo) & "," & _
        SqlStringService.GetSqlString(u.Imie) & "," & _
        SqlStringService.GetSqlString(u.Nazwisko) & "," & _
        SqlStringService.GetSqlString(u.AdresEmail) & "," & _
        SqlStringService.GetSqlString(u.Tytul) & ")"
End Function

Private Function BuildUpdateSql(u As user) As String
    BuildUpdateSql = _
        "UPDATE tblUsers SET " & _
        "NazwaUzytkownika=" & SqlStringService.GetSqlString(u.NazwaUzytkownika) & ", " & _
        "HashHaslo=" & SqlStringService.GetSqlString(u.HashHaslo) & ", " & _
        "Imie=" & SqlStringService.GetSqlString(u.Imie) & ", " & _
        "Nazwisko=" & SqlStringService.GetSqlString(u.Nazwisko) & ", " & _
        "AdresEmail=" & SqlStringService.GetSqlString(u.AdresEmail) & ", " & _
        "Tytul=" & SqlStringService.GetSqlString(u.Tytul) & _
        " WHERE Id=" & u.id
End Function

Private Function Q(v As Variant) As String
    If IsNull(v) Or v = "" Then
        Q = "NULL"
    Else
        Q = "'" & Replace(v, "'", "''") & "'"
    End If
End Function

Public Function CreateUserFromRecord(rs As DAO.Recordset) As user
    Dim u As user
    Set u = New user
    If rs.EOF Then
        Set CreateUserFromRecord = Nothing
        Exit Function
    End If
    
    With u
        .NazwaUzytkownika = Trim(NzSafe(rs!NazwaUzytkownika, ""))
        .Imie = Trim(NzSafe(rs!Imie, ""))
        .Nazwisko = Trim(NzSafe(rs!Nazwisko, ""))
        .AdresEmail = Trim(NzSafe(rs!AdresEmail, ""))
        .Tytul = Trim(NzSafe(rs!Tytul, ""))
    End With
    
    Set CreateUserFromRecord = u
End Function
