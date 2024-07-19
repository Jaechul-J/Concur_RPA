Sub SortAndMergeCenter()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Main")  ' Replace "Main" with your actual sheet name
    
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' Find the last row and last column with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Sort by column A in descending order and expand to other columns
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Sort Key1:=ws.Cells(1, 1), _
        Order1:=xlDescending, Header:=xlYes
End Sub
