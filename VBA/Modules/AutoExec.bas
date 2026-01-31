Attribute VB_Name = "AutoExec"
Option Compare Database
Option Explicit


Public Sub AutoExec()
    DoCmd.OpenForm "LoginView", acNormal
    
End Sub


