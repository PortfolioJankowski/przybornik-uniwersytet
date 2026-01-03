Attribute VB_Name = "GitHelperService"
Option Compare Database

Public Sub ExportAllVBA()
    EnsureBaseFolders
    ClearVbaFolders

    ExportModules
    ExportClassModules
    ExportAccessObjects

    MsgBox "Export VBA completed (clean)", vbInformation
End Sub

Public Sub ImportAllVBA()
    RemoveAllCode
    ImportFromFolder GetProjectPath & "Modules\"
    ImportFromFolder GetProjectPath & "ClassModules\"
    ImportFromFolder GetProjectPath & "AccessObjects\"

    MsgBox "Import VBA completed", vbInformation
End Sub

Private Function GetProjectPath() As String
    GetProjectPath = CurrentProject.path & "\VBA\"
End Function

Private Sub EnsureFolder(ByVal path As String)
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
End Sub

Private Sub EnsureBaseFolders()
    EnsureFolder GetProjectPath
    EnsureFolder GetProjectPath & "Modules\"
    EnsureFolder GetProjectPath & "ClassModules\"
    EnsureFolder GetProjectPath & "AccessObjects\"
End Sub

Public Sub ClearVbaFolders()
    ClearFolder GetProjectPath & "Modules\"
    ClearFolder GetProjectPath & "ClassModules\"
    ClearFolder GetProjectPath & "AccessObjects\"
End Sub

Private Sub ClearFolder(ByVal folderPath As String)
    Dim file As String

    If Dir(folderPath, vbDirectory) = "" Then Exit Sub

    file = Dir(folderPath & "*.*")
    Do While file <> ""
        Kill folderPath & file
        file = Dir
    Loop
End Sub

Private Sub ExportModules()
    Dim cmp As Object
    
    For Each cmp In Application.VBE.VBProjects(1).VBComponents
        If CStr(cmp.Type) = "1" Then
            cmp.Export GetProjectPath & "Modules\" & cmp.Name & ".bas"
        End If
    Next
End Sub

Private Sub ExportClassModules()
    Dim cmp As Object
    For Each cmp In Application.VBE.VBProjects(1).VBComponents
        If CStr(cmp.Type) = "2" Then
            cmp.Export GetProjectPath & "ClassModules\" & cmp.Name & ".cls"
        End If
    Next
End Sub

Private Sub ExportAccessObjects()
    Dim cmp As Object
    For Each cmp In Application.VBE.VBProjects(1).VBComponents
        If CStr(cmp.Type) = "100" Then
            cmp.Export GetProjectPath & "AccessObjects\" & cmp.Name & ".cls"
        End If
    Next
End Sub

Private Sub RemoveAllCode()
    Dim cmp As Object

    For Each cmp In Application.VBE.VBProjects(1).VBComponents
        Select Case cmp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                Application.VBE.VBProjects(1).VBComponents.Remove cmp
        End Select
    Next
End Sub

Private Sub ImportFromFolder(ByVal folderPath As String)
    Dim file As String
    file = Dir(folderPath & "*.*")

    Do While file <> ""
        Application.VBE.VBProjects(1).VBComponents.Import folderPath & file
        file = Dir
    Loop
End Sub
