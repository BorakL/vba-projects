Sub AutoFitToA4LandscapeEnhanced()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long, lastCol As Long
    Dim cellWidth As Double, cellHeight As Double
    Dim fontSize As Integer
    Dim pageHeightPoints As Double, pageWidthPoints As Double
    
    ' Set the worksheet (modify "Sheet1" if needed)
    Set ws = Worksheets("Sheet1")
    
    ' Determine the used range of the worksheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' Set page layout to A4 landscape
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape
        .Zoom = False   ' Ensure fit-to-page
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    
    ' Get A4 dimensions in points (landscape)
    pageWidthPoints = 29.7 * 28.346   ' Width of A4 in points
    pageHeightPoints = 21 * 28.346    ' Height of A4 in points

    ' Adjust column width and row height to maximize space
    cellWidth = pageWidthPoints / lastCol       ' Calculate optimal width per column
    cellHeight = pageHeightPoints / lastRow     ' Calculate optimal height per row

    With rng
        .Columns.ColumnWidth = cellWidth / 5.8 ' Convert points to Excel units for width
        .Rows.RowHeight = cellHeight / 1.1    ' Increase row height for better fit
    End With
    
    ' Adjust font size dynamically based on calculated dimensions
    fontSize = WorksheetFunction.Min(cellWidth / 3, cellHeight / 1.5)
    If fontSize < 8 Then fontSize = 8 ' Set a minimum font size for readability
    rng.Font.Size = fontSize
    
    ' Center-align text and wrap it
    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    MsgBox "Worksheet optimized to fill A4 landscape page, using all available space!"
End Sub