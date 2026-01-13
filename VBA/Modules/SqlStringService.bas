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
