Sub grouping()
' removeRDV
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
    
    ' Define the active worksheet
    Set ws = ActiveSheet
    
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
                
' groupingOfShippingDocuments
Dim lastRowG As Long
Dim lastRowH As Long
Dim isRowEmpty As Boolean

' Copy data from column A to column G
Columns("A:A").Copy
Columns("G:G").PasteSpecial xlPasteValues
Application.CutCopyMode = False

' Find the last filled row and last column
lastRowG = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

' Loop through all rows and delete rows where all cells are empty
For i = lastRowG To 1 Step -1
    isRowEmpty = True
    For j = 1 To lastCol
        If Trim(ws.Cells(i, j).Value) <> "" Then
            isRowEmpty = False
            Exit For
        End If
    Next j
    If isRowEmpty Then
        ws.Rows(i).Delete
    End If
Next i

' Update the last filled row in column G after deleting empty rows
lastRowG = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

' Remove duplicates dynamically based on the last filled row in column G
ws.Range("G1:G" & lastRowG).RemoveDuplicates Columns:=1, Header:=xlNo

' Write the SUMIF formula in H1
ws.Range("H1").FormulaR1C1 = "=SUMIF(C[-7],RC[-1],C[-3])"

' Find the last filled row in column H (after duplicates are removed)
lastRowH = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

' AutoFill the formula in H1 down to the last filled row in column G
ws.Range("H1").AutoFill Destination:=ws.Range("H1:H" & lastRowG), Type:=xlFillDefault
End Sub
