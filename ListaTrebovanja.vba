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
    Dim numChars As Integer
    Dim colWidth As Double
    Dim fontSize As Double
    Dim constant As Double
    
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

    ' Konstant za računanje font size-a
    constant = 12 ' Možeš promeniti kako bi dobio optimalan rezultat

    ' Maksimalni font size
    maxFontSize = Rows(2).RowHeight * 0.8 ' Postavi maksimalnu veličinu fonta

    ' Prvi red - postavljanje font size
    For Each cell In ws.Rows(2).Cells
        If cell.Value <> "" Then
            numChars = Len(cell.Value)
            colWidth = ws.Columns(cell.Column).ColumnWidth

            ' Izračunaj font size na osnovu broja karaktera i širine kolone
            fontSize = (colWidth / numChars) * constant

            ' Ograniči maksimalnu veličinu fonta
            If fontSize > maxFontSize Then fontSize = Round(maxFontSize)

            ' Postavi font size
            cell.Font.Size = fontSize
        End If
    Next cell

    ' Prva kolona - postavljanje font size
    For Each cell In ws.Columns(1).Cells
        If cell.Value <> "" Then
            numChars = Len(cell.Value)
            colWidth = ws.Columns(1).ColumnWidth

            ' Izračunaj font size na osnovu broja karaktera i širine kolone
            fontSize = (colWidth / numChars) * constant

            ' Ograniči maksimalnu veličinu fonta
            If fontSize > maxFontSize Then fontSize = Round(maxFontSize)

            ' Postavi font size
            cell.Font.Size = fontSize
        End If
    Next cell

  ' Nađi poslednji popunjeni red i poslednju popunjenu kolonu
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row ' Poslednji red na osnovu druge kolone
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column ' Poslednja kolona na osnovu drugog reda

    ' Definiši opseg za promenu fonta (od B2 do poslednje ćelije)
    Set tableRange = ws.Range(ws.Cells(3, 2), ws.Cells(lastRow, lastCol))

    ' Promeni font za sve ćelije u tabeli
    With tableRange
        .Font.Name = "Calibri"         ' Postavi font
        .Font.Size = Rows(2).RowHeight ' Prilagođena veličina fonta prema visini reda
        .Font.Bold = True              ' Postavi bold
        .HorizontalAlignment = xlCenter ' Horizontalno centriranje
        .VerticalAlignment = xlCenter   ' Vertikalno centriranje
    End With

End Sub
