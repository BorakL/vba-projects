Sub a()
 Dim ws As Worksheet
    Dim colCount As Integer
    Dim rowCount As Integer
    Dim totalHeightUnits As Double
    Dim rowHeightUnits As Double
    Dim i As Integer
    Dim effectiveHeight As Double
    Dim rng As Range
    Dim cell As Range
    Dim maxFontSize As Double
    Dim tempFontSize As Double
    Dim fit As Boolean
    
    ' Set the active worksheet
    Set ws = ActiveSheet

    ' Get the number of columns and rows in the used range
    colCount = ws.UsedRange.Columns.Count
    rowCount = ws.UsedRange.Rows.Count

    ' Efektivna visina A4 stranice u landscape (u Excel jedinicama)
    ' Prosečna efektivna visina A4 u landscape je oko 80 Excel jedinica
    effectiveHeight = 800 ' Prilagođeno eksperimentalno za ceo prikaz

    ' Raspodela visine među redovima
    rowHeightUnits = effectiveHeight / rowCount

    ' Postavljanje visine za svaki red
    For i = 1 To rowCount
        ws.Rows(i).RowHeight = rowHeightUnits
    Next i

    ' Dodatni deo koda za podešavanje širina kolona (već imaš, ali ponavljam za celovitost)
    Dim firstColWidthUnits As Double
    Dim otherColsWidthUnits As Double
    Dim effectiveWidth As Double

    ' A4 landscape effective width in Excel units
    effectiveWidth = 200 ' Eksperimentalno za širinu cele stranice

    ' Determine width for the first column based on rules
    Select Case colCount
        Case Is < 4
            firstColWidthUnits = 0.66 * effectiveWidth
        Case Is > 4
            firstColWidthUnits = 0.33 * effectiveWidth
        Case 4
            firstColWidthUnits = 0.5 * effectiveWidth
    End Select

    ' Distribute remaining width among other columns
    If colCount > 1 Then
        otherColsWidthUnits = (effectiveWidth - firstColWidthUnits) / (colCount - 1)
    End If

    ' Set column widths
    ws.Columns(1).ColumnWidth = firstColWidthUnits
    For i = 2 To colCount
        ws.Columns(i).ColumnWidth = otherColsWidthUnits
    Next i

    ' Apply borders to the entire used range
    With ws.UsedRange
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ' Set print settings
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
    End With
End Sub
