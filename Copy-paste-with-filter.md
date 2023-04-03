Suppose you are asked to apply filter on a column and paste result of a filter into a new worksheet or workbook and 
same process goes until all the unique values of the column are covered. In other words, 
this needs to be done for each unique values in a column in which we have applied filter. 
It is a very time consuming process if you do it manually. For example, you have a column in which there are 50 unique values. 
You have to do it 50 times which is a tedious and error-prone task. It can be easily done with Excel VBA programming.

The sample data is shown below :

Filtering and Copying Data

How to Use
Open an Excel Workbook
Press Alt+F11 to open VBA Editor
Go to Insert Menu >> Module
In the module, paste the below program
Save the file as Macro Enabled Workbook (xlsm) or Excel 97-2003 Workbook (xls)

In the following excel macro, it is assumed a filter is applied on column F (Rank) and data starts from cell A1.

Excel Macro : Filter and Paste Unique Values to New Sheets



This macro would filter a column and paste distinct values to the sheets with their respective names. In this case, 
it creates four worksheets - 1 , 2, 3, 4 as these are unique values in column Rank (column F).

```Sub filter()
Application.ScreenUpdating = False
Dim x As Range
Dim rng As Range
Dim last As Long
Dim sht As String

'specify sheet name in which the data is stored
sht = "DATA Sheet"

'change filter column in the following code (F Replace instead of  filter Column name)
last = Sheets(sht).Cells(Rows.Count, "F").End(xlUp).Row

'set the Range(only add Last Column name example:F)
Set rng = Sheets(sht).Range("A1:F" & last)

'" F " Replace instead of  filter Column name
Sheets(sht).Range("F1:F" & last).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("AA1"), Unique:=True

For Each x In Range([AA2], Cells(Rows.Count, "AA").End(xlUp))

With rng
.AutoFilter
''" 6 " Replace instead of  filter Column index number
.AutoFilter Field:=6, Criteria1:=x.Value
.SpecialCells(xlCellTypeVisible).Copy

Sheets.Add(After:=Sheets(Sheets.Count)).Name = x.Value
ActiveSheet.Paste
End With
Next x

' Turn off filter
Sheets(sht).AutoFilterMode = False

With Application
.CutCopyMode = False
.ScreenUpdating = True
End With

End Sub
```

How to Filter and Paste Values to New Workbook

How to Customize the above program

1. Specify name of the sheet in which data is stored. Change the below line of code in the program.
sht = "DATA Sheet"

2. Change filter column (column F) and starting cell of range (A1) in the code.
last = Sheets(sht).Cells(Rows.Count, "F").End(xlUp).Row
Set rng = Sheets(sht).Range("A1:F" & last)

3. Starting cell of filter column - F1. Unique values of  column F are stored in column AA.
Sheets(sht).Range("F1:F" & last).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("AA1"), Unique:=True
For Each x In Range([AA2], Cells(Rows.Count, "AA").End(xlUp))

4. Change the value in this part of the code. In this case, 6 refers to column index number (i.e. Column F is 6th column).
.AutoFilter Field:=6, Criteria1:=x.Value

More details:
https://www.listendata.com/2015/04/excel-vba-filtering-and-copy-pasting-to.html
