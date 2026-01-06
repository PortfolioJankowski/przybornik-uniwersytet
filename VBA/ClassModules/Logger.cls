VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sub Add(wpis As wpis)
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("Logi", dbOpenDynaset)
    rs.AddNew
        rs!Data = Now
        rs!NumerBledu = wpis.ErrorNumber
        rs!Opis = wpis.Description
        rs!StosWywolan = wpis.CallStac
    rs.Update
    Set rs = Nothing
End Sub
