Attribute VB_Name = "Logger"
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

