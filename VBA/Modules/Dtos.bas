Attribute VB_Name = "Dtos"
Option Compare Database
Option Explicit

Public Type zdarzenie
     name As String
     Message As String
     isProcessed As Boolean
     ProcessingDate As Date
End Type

Public Type LoginValidationDto
    UserRecordset As Recordset
    DoesUserExist As Boolean
End Type

Public Type wpis
    LogDate As Date
    ErrorNumber As Integer
    Description As String
    CallStac As String
End Type

Public Enum PrzechowywanyPlik
    Prezentacja = 0
    Notatki = 1
End Enum


Public Type Class
    Order As Integer
    Title As String
    Description As String
    cycleId As String
End Type
