Attribute VB_Name = "Module2"
Sub A7_TransferAllToFirstColumn()

Dim ws As Worksheet
Dim i As Integer
Dim lastcol As Long

For Each ws In Worksheets

ws.Activate

'deleting blank cells in all columns

    ws.Columns("A:Z").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete shift:=xlUp
 
'inserting empty column to the left

    ws.Columns("A:A").EntireColumn.Select
    Selection.Insert shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
    
'arrange all columns into single column A

lastcol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    
    For i = 2 To lastcol
    
        Range(Cells(1, i), Cells(Rows.Count, i).End(xlUp)).Select
        Application.CutCopyMode = False
        Selection.Copy
        ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial
        
    Next i

'Next ws

End Sub

Sub A8_CopyColumnsAtoNewSheet()

Sheets.Add(Before:=Sheets(1)).Name = "all"

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Select
    
    Range(Cells(2, 1), Cells(Rows.Count, 1).End(xlUp)).Select
    Selection.Copy
    ActiveSheet.Paste Destination:=Worksheets("all").Cells(1, Columns.Count).End(xlToLeft).Offset(0, 1)
    
Next ws

End Sub
