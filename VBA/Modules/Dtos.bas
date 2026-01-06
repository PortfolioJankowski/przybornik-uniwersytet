Attribute VB_Name = "Dtos"
Option Compare Database
Option Explicit

Public Type Zdarzenie
     EventDate As Date
     Name As String
     Message As String
     IsProcessed As Boolean
     ProcessingDate As Date
End Type

Public Type LoginValidationDto
    UserRecordset As Recordset
    DoesUserExist As Boolean
End Type
