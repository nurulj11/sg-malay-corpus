Attribute VB_Name = "Module1"
Sub A1_ImportFiles()

' importing multiple files -- compatible for Windows system only

        Dim sheet As Worksheet
        Dim total As Integer
        Dim intChoice As Integer
        Dim strPath As String
        Dim i As Integer
        Dim wbNew As Workbook
        Dim wbSource As Workbook
        Set wbNew = Workbooks.Add


        'allow the user to select multiple files
        Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
        'make the file dialog visible to the user
        intChoice = Application.FileDialog(msoFileDialogOpen).Show

        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

        'determine what choice the user made
        If intChoice <> 0 Then
            'get the file path selected by the user
            For i = 1 To Application.FileDialog(msoFileDialogOpen).SelectedItems.Count
                strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(i)

                Set wbSource = Workbooks.Open(strPath)

                For Each sheet In wbSource.Worksheets
                    total = wbNew.Worksheets.Count
                    wbSource.Worksheets(sheet.Name).Copy _
                    after:=wbNew.Worksheets(total)
                Next sheet

                wbSource.Close
            Next i
        End If

wbNew.Activate

Sheets("Sheet1").Select
ActiveWindow.SelectedSheets.Delete

    End Sub

Sub A2_Del_timemarkers()

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

Sub A3_RemovePunctuationMarks()

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
    Selection.Replace What:="Ã¢â‚¬Å", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="Ã¢â‚¬Â", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="Ã¢â‚¬Ëœ", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
      
            
Next ws

End Sub

Sub A4_Correct_dasherror()
'
' words that start with - to remove
'
Dim ws As Worksheet

For Each ws In Worksheets

    ws.Select

    ActiveSheet.Columns("A:A").EntireColumn.Select
    Selection.AutoFilter
    ActiveSheet.Range(Cells(1, 1), Cells(Rows.Count, 1).End(xlUp)).AutoFilter Field:=1, Criteria1:="=-*", _
        Operator:=xlAnd
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveSheet.Range(Cells(1, 1), Cells(Rows.Count, 1).End(xlUp)).AutoFilter Field:=1

Next ws

End Sub

Sub A5_CorrectforNAMEerror()

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Activate
    
    ActiveSheet.Columns("A:A").EntireColumn.Select
    
    Selection.Replace What:="=-", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
Next ws

End Sub


Sub A6_TextToColumns()

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



