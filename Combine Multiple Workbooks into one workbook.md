**Combine Multiple Workbooks into one workbook(convert multiple sheets)**

#Line 10 here, you need to change the value for the path variable with the folder location 
where you have all the Excel files that you want to combine

```Sub CopyRange()

Dim Path As String
Dim Filename As String
Dim Sheet As Worksheet

'Change the Path
Path = "C:\Users\Dell\Desktop\sample-data\"

Filename = Dir(Path & "*.xlsx")
Do While Filename <> ""

Workbooks.Open Filename:=Path & Filename, ReadOnly:=True

    For Each Sheet In ActiveWorkbook.Sheets
        Sheet.Copy After:=ThisWorkbook.Sheets(1)
    Next Sheet

Workbooks(Filename).Close

Filename = Dir()
Loop

End Sub
```
