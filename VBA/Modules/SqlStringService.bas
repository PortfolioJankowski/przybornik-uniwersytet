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

Public Function GetClassesByCycleId(cycleId As Long) As String
    GetClassesByCycleId = "SELECT * FROM Zajecia WHERE CyklDydaktycznyId = " & CStr(cycleId)
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

Public Function GetClassIdByClassNoAndCycleID(classNo As String, groupId As String) As String
    GetClassIdByClassNoAndCycleID = "SELECT Identyfikator from Zajecia WHERE CyklDydaktycznyId =" & groupId & "And Kolejnosc=" & classNo
End Function


Public Function GetStudentIdByNrAlbumu(NrAlbumu As String) As Variant
    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT Identyfikator FROM Studenci WHERE NrAlbumu = '" & NrAlbumu & "'"

    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)

    If rs.EOF Then
        GetStudentIdByNrAlbumu = Null
    Else
        GetStudentIdByNrAlbumu = rs!Identyfikator
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function InsertStudentToGroup(groupId As Long, studentId As Long) As String
   InsertStudentToGroup = _
        "INSERT INTO StudenciWGrupach (GrupaId, StudentId) " & _
        "VALUES (" & groupId & ", " & studentId & ")"

End Function
