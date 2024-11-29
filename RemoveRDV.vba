Sub removeRDV()

    Dim regex As Object
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim cellText As String
    Dim pattern As String
    
    ' Define the regex pattern
    pattern = "\(\d+\-(\d+)?[DRV]\)" ' Replace this with your regex pattern
    
    ' Create regex object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.Global = True ' Allow global matches (multiple matches in a cell)
    
    ' Define the worksheet
    Set ws = ThisWorkbook.Sheets(2) ' Change to your worksheet name or index
    
    ' Find the last row and column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Loop through all cells
    For i = 1 To lastRow
        For j = 1 To lastCol
            cellText = ws.Cells(i, j).Value ' Get cell value
            If regex.Test(cellText) Then
                ' Remove regex matches
                ws.Cells(i, j).Value = regex.Replace(cellText, "")
            End If
        Next j
    Next i
    
    MsgBox "Regex matches cleared from the worksheet!", vbInformation

End Sub