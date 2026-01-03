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
