*To create multiple Excel files each containing a single worksheet using VBA, you can use the following code:

#Copy code

```
Sub CreateFiles()

Dim wb As Workbook
Dim ws As Worksheet
Dim FilePath As String

'Loop through each worksheet in the current workbook
For Each ws In ThisWorkbook.Worksheets

    'Create a new workbook
    Set wb = Workbooks.Add

    'Copy the current worksheet to the new workbook
    ws.Copy Before:=wb.Sheets(1)

    'Get the file path and name
    FilePath = ThisWorkbook.Path & "\" & ws.Name & ".xlsx"

    'Save the new workbook with the worksheet name as the file name
    wb.SaveAs FilePath
    wb.Close False

Next ws

End Sub
```
*This code will loop through each worksheet in the current workbook and create a new workbook for each one. It will then copy the current worksheet to the new workbook, save the new workbook with the worksheet name as the file name, and close the workbook.

Note that the code assumes that you want to save the new files in the same folder as the current workbook. If you want to save the files in a different folder, you will need to modify the FilePath variable accordingly.
