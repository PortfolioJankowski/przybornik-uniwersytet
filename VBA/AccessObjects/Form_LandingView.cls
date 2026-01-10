VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LandingView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnPrzedmioty_Click()
    DoCmd.OpenForm "PrzedmiotyView"
End Sub

Private Sub Form_Load()
If App Is Nothing Then
    Exit Sub
End If
    With Me
        .txtName.value = App.CurrentUser.Imie
        .txtDate.value = "Dzisiaj jest " & CStr(Date)
    End With
End Sub
