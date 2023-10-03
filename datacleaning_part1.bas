Sub A1_Del_TimeStampsNameErrors()

' Corrects #NAME? error by replacing those that begin with =-
' Deletes time markers and also numbers

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Activate
    
'Corrects for the #NAME error formula due to the addition of=- at the start of the sentence.
'Removing =-will reveal the sentence instead of #NAME error in the cell
    ActiveSheet.Columns("A:A").Replace What:="=-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
'Filters column A for cells with integers greater or equal than 1 or cells with--, which represent
'cells with timestamps
    ActiveSheet.Range("A:A").AutoFilter Field:=1, Criteria1:=">=1", _
        Operator:=xlOr, Criteria2:="=*-->*"
    
'Deletes cells that fulfil the criteria above
    ActiveSheet.Range("A:A").SpecialCells(xlCellTypeVisible).Select
    Selection.EntireRow.Delete
    
'Deletes cells with blanks
    ActiveSheet.Range("A:A").SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete shift:=xlUp
    
Next ws

End Sub


Sub A2_RemovePunctuations()

'Looks for punctuation marks in each row i.e. sentence and removes them

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Activate
    
    ActiveSheet.Columns("A:A").Select
    
    Selection.Replace What:="~?", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:=".", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="!", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:=",", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="@", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
      
            
Next ws

End Sub


Sub A3_TextToColumns()

'Converts sentence text to one word per column, with space as the delimiter 

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Activate
    
    ActiveSheet.Range("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1)), _
        TrailingMinusNumbers:=True
    
Next ws
    
End Sub



