Import multiple file names into worksheet cells with VBA code

The following VBA code can help you import the file names, file extensions and folder name into the worksheet cells, please do with following steps:

1. Launch a new worksheet that you want to import the file names.

2. Hold down the ALT + F11 keys to open the Microsoft Visual Basic for Applications window.

3. Click Insert > Module, and paste the following code in the Module Window.

VBA code: Import multiple file names into cells of worksheet



```
Sub GetFileList()
'updateby Extendoffice
    Dim xFSO As Object
    Dim xFolder As Object
    Dim xFile As Object
    Dim xFiDialog As FileDialog
    Dim xPath As String
    Dim i As Integer
    Set xFiDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If xFiDialog.Show = -1 Then
        xPath = xFiDialog.SelectedItems(1)
    End If
    Set xFiDialog = Nothing
    If xPath = "" Then Exit Sub
    Set xFSO = CreateObject("Scripting.FileSystemObject")
    Set xFolder = xFSO.GetFolder(xPath)
    ActiveSheet.Cells(1, 1) = "Folder name"
    ActiveSheet.Cells(1, 2) = "File name"
    ActiveSheet.Cells(1, 3) = "File extension"
    i = 1
    For Each xFile In xFolder.Files
        i = i + 1
        ActiveSheet.Cells(i, 1) = xPath
        ActiveSheet.Cells(i, 2) = Left(xFile.Name, InStrRev(xFile.Name, ".") - 1)
        ActiveSheet.Cells(i, 3) = Mid(xFile.Name, InStrRev(xFile.Name, ".") + 1)
    Next
End Sub
```
