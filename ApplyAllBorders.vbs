Sub ApplyAllBorders()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    
    For Each ws In ThisWorkbook.Worksheets
        With ws
            ' Find the last used row and column
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
            
            ' Apply all borders to the used range
            .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Borders.LineStyle = xlContinuous
        End With
    Next ws
End Sub