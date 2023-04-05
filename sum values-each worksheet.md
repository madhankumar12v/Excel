**The following program sums column B of each worksheet. It also writes 'Total' at last row of column A. 
It assumes your data starts from column A**

```
Sub sumLoop()
Dim WS As Worksheet
For Each WS In ThisWorkbook.Worksheets
WS.Activate
Range("A" & Rows.Count).End(xlUp).Select
last = Selection.Row
totRow = last + 1
WS.Range("A" & totRow) = "Total"
'if you want to add more columns. Just line No:14 copy and paste in Next Line (Replace the column "Require Column" instead of "B" )
WS.Range("B" & totRow) = Application.WorksheetFunction.Sum(Columns("B:B"))
Next WS
End Sub
```
