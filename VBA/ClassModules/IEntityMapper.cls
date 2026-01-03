VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEntityMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function Map(rs As DAO.Recordset) As Object
End Function

Public Function ToSqlInsert(entity As Object) As String
End Function

Public Function ToSqlUpdate(entity As Object) As String
End Function
