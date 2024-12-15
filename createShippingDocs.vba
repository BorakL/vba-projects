Sub createShippingDocs()

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
 

' createTable Macro

    Columns("B:B").Select
    ActiveWindow.Zoom = 85
    ActiveWindow.Zoom = 100
    ActiveWindow.Zoom = 115
    Range("B:B,C:C,D:D").Select
    Range("D1").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("D:D,E:E,F:F").Select
    Range("F1").Activate
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.ScrollColumn = 2
    Range("F:F,G:G,H:H").Select
    Range("H1").Activate
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.ScrollColumn = 1
    Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Font.Size = 14
    Columns("B:B").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Size = 16
    Columns("D:D").Select
    Selection.Font.Size = 16
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveWindow.ScrollColumn = 2
    Columns("F:F").Select
    Selection.Font.Size = 16
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveWindow.ScrollColumn = 1
    Cells.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Range("C13").Select 


    Dim regexMlekaCaj As Object, regexVanRFZO As Object
    Dim patternMlekaCaj As String, patternVanRFZO As String
    Dim foundMlekaCaj As Boolean, foundVanRFZO As Boolean
    
    ' Define the patterns
    patternMlekaCaj = "\(M\-|\(\u010C\-|\(C\-"
    patternVanRFZO = "VAN RFZO"
    
    ' Create regex objects
    Set regexMlekaCaj = CreateObject("VBScript.RegExp")
    regexMlekaCaj.Pattern = patternMlekaCaj
    regexMlekaCaj.Global = False ' Match the first occurrence only
    
    Set regexVanRFZO = CreateObject("VBScript.RegExp")
    regexVanRFZO.Pattern = patternVanRFZO
    regexVanRFZO.Global = False ' Match the first occurrence only
    
    ' Define the active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row and column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Initialize flags
    foundMlekaCaj = False
    foundVanRFZO = False
    
    ' Loop through all cells
    For i = 1 To lastRow
        For j = 1 To lastCol
            cellText = ws.Cells(i, j).Value ' Get cell value
            
            ' Check for mleko/čaj patterns
            If regexMlekaCaj.Test(cellText) Then
                foundMlekaCaj = True
            End If
            
            ' Check for VAN RFZO pattern
            If regexVanRFZO.Test(cellText) Then
                foundVanRFZO = True
            End If
            
            ' Exit loops early if both are found
            If foundMlekaCaj And foundVanRFZO Then Exit For
        Next j
        If foundMlekaCaj And foundVanRFZO Then Exit For
    Next i
    
    ' Show messages based on findings
    If foundMlekaCaj Then
        MsgBox "Ima mleko/čaj", vbInformation
    End If
    
    If foundVanRFZO Then
        MsgBox "Ima van RFZO-a", vbInformation
    End If
End Sub