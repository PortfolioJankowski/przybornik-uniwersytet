Attribute VB_Name = "SqlStringService"
Option Compare Database
Option Explicit

Public Function GetSqlString(v As Variant) As String
    If IsNull(v) Or v = "" Then
        GetSqlString = "NULL"
    Else
        GetSqlString = "'" & Replace(v, "'", "''") & "'"
    End If
End Function

Public Function GetClassesByCycleId(CycleId As Long) As String
    GetClassesByCycleId = "SELECT * FROM Zajecia WHERE CyklDydaktycznyId = " & CStr(CycleId)
End Function

Public Function DeleteClassById(classId As Long) As String
    DeleteClassById = "DELETE FROM Zajecia WHERE Identyfikator =" & CStr(classId)
End Function

Public Function GetClassById(classId As Long) As String
    GetClassById = "SELECT * FROM Zajecia WHERE Identyfikator=" & CStr(classId)
End Function

Public Function DeleteSubjectById(subjectId As Long) As String
    DeleteSubjectById = "DELETE FROM Przedmioty WHERE Identyfikator=" & CStr(subjectId)
End Function

Public Function GetSubjectById(subjectId As Long) As String
    GetSubjectById = "Select * FROM Przedmioty where Identyfikator =" & CStr(subjectId)
End Function

Public Function InsertIntoClasses(Class As Class) As String
    InsertIntoClasses = "INSERT INTO Zajecia (Tytul, Kolejnosc, Opis, CyklDydaktycznyId) Values ('" _
    & Class.Title & "'," & CStr(Order) & ",'" & Class.Description & "'," & Class.CycleId & ")"
End Function
