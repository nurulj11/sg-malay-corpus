Attribute VB_Name = "Module2"

Sub A4_TransferAllToFirstColumn()

'Deletes blank cells in between rows in each column, followed by 
'Compiling all columns into a single column on the left i.e. Column A

Dim ws As Worksheet
Dim i As Integer
Dim lastcol As Long

For Each ws In Worksheets

ws.Activate

'Deletes blank cells in each columns

    ws.Columns("A:AZ").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete shift:=xlUp
 
'Insert blank column to the left, i.e. Column A

    ws.Columns("A:A").EntireColumn.Select
    Selection.Insert shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
    
'Compile all columns into column A

lastcol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    
    For i = 2 To lastcol
    
        Range(Cells(1, i), Cells(Rows.Count, i).End(xlUp)).Select
        Application.CutCopyMode = False
        Selection.Copy
        ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial
        
    Next i

'Next ws

End Sub


Sub A5_CopyColumnsAtoNewSheet()

'Copy column A from each sheet to sheet named "All"
'Words from all episodes in each sheet will be compiled into single sheet "All"

Sheets.Add(Before:=Sheets(1)).Name = "All"

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Select
    
    Range(Cells(2, 1), Cells(Rows.Count, 1).End(xlUp)).Select
    Selection.Copy
    ActiveSheet.Paste Destination:=Worksheets("All").Cells(1, Columns.Count).End(xlToLeft).Offset(0, 1)
    
Next ws

End Sub
