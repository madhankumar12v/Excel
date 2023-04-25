Rename multiple files of a folder in Excel with VBA code

1. Hold down the ALT + F11 keys, and it opens the Microsoft Visual Basic for Applications Window.

2. Click Insert > Module, and paste the following macro in the Module window.

VBA code: Rename multiple files in a folder

More Details:https://www.extendoffice.com/documents/excel/2339-excel-rename-files-in-a-folder.html

#Just and Paste

```
Sub RenameFiles()
'Updateby20141124
Dim xDir As String
Dim xFile As String
Dim xRow As Long
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
If .Show = -1 Then
    xDir = .SelectedItems(1)
    xFile = Dir(xDir & Application.PathSeparator & "*")
    Do Until xFile = ""
        xRow = 0
        On Error Resume Next
        xRow = Application.Match(xFile, Range("A:A"), 0)
        If xRow > 0 Then
            Name xDir & Application.PathSeparator & xFile As _
            xDir & Application.PathSeparator & Cells(xRow, "B").Value
        End If
        xFile = Dir
    Loop
End If
End With
End Sub

```
