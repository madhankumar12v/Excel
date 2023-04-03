**Merging all sheets of active workbook into one sheet with VBA**

In this section, I provide a VBA code which will create a new sheet to collect all sheets of the active workbook while you running it.

1. Activate the workbook you want to combine its all sheets, then press Alt + F11 keys to open Microsoft Visual Basic for Applications window.

2. In popping window, click Insert > Module to create a new Module script.

3. Copy below code and paste them to the script.

##Just Copy and Paste
```
Sub Combine()
'UpdatebyExtendoffice
Dim J As Integer
On Error Resume Next
Sheets(1).Select
Worksheets.Add
Sheets(1).Name = "Combined"
Sheets(2).Activate
Range("A1").EntireRow.Select
Selection.Copy Destination:=Sheets(1).Range("A1")
For J = 2 To Sheets.Count
Sheets(J).Activate
Range("A1").Select
Selection.CurrentRegion.Select
Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
Selection.Copy Destination:=Sheets(1).Range("A65536").End(xlUp)(2)
Next
End Sub

```
4. Press F5 key, then all data across sheets have been merged in to a new sheet named Combined which is placed in the front of all sheets.

More:https://www.extendoffice.com/documents/excel/1184-excel-merge-multiple-worksheets-into-one.html
