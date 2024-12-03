Sub a()
    Dim ws As Worksheet
    Dim rg As Range
    Dim secondRow As Range
    Dim foundCell As Range
    Dim valuesToFind As Variant
    Dim value As Variant
    Dim secondRowValue As String
    Dim searchText As String
    Dim columnsToDelete As Range

    ' Postavi aktivni radni list
    Set ws = ActiveSheet
    
    ' Definiši opseg tabele
    Set rg = ws.UsedRange ' Ili konkretan opseg ako je poznat, npr. ws.Range("A1:G10")

    ' Kreiraj opseg za drugi red
    Set firstRow = Intersect(rg, ws.Rows(1))
    Set secondRow = Intersect(rg, ws.Rows(2))

    ' Proveri da li je secondRow validan
    If Not secondRow Is Nothing Then
        ' Niz vrednosti koje treba pronaći
        valuesToFind = Array("HIRURGIJA 2", "BLOK A", "BLOK B", "INFEKTIVNE I TROPSKE BOLESTI", "ENDOKRINOLOGIJA")
        
        ' Prođi kroz niz i pronađi vrednost u drugom redu
        For Each value In valuesToFind
            Set foundCell = secondRow.Find(What:=value, LookIn:=xlValues, LookAt:=xlWhole)
            If Not foundCell Is Nothing Then
                secondRowValue = foundCell.value ' Pronađena vrednost se čuva u secondRowValue
                Exit For ' Prekidamo petlju jer smo našli traženu vrednost
            End If
        Next value
    End If
    
    ' Obrada prema vrednosti pronađenoj u drugom redu, ili obradi kao podrazumevani slučaj
    Select Case secondRowValue
        Case "HIRURGIJA 2"
            ' Kombinuj kolone A i E za brisanje
            Set columnsToDelete = Union(ws.Columns("A"), ws.Columns("E"))
            columnsToDelete.Delete
            ws.Cells(1, 1).value = secondRowValue
        Case "BLOK A", "BLOK B", "INFEKTIVNE I TROPSKE BOLESTI"
            ' Kombinuj kolone A i C za brisanje
            Set columnsToDelete = Union(ws.Columns("A"), ws.Columns("C"))
            columnsToDelete.Delete
            ws.Cells(1, 1).value = secondRowValue
        Case "ENDOKRINOLOGIJA"
            ' Samo obriši kolonu A
            ws.Columns("A").Delete
            ws.Cells(1, 1).value = "INTERNA B"
        Case Else
            ' Kombinuj kolone A i D za brisanje, obriši prvi red
            Set columnsToDelete = Union(ws.Columns("A"), ws.Columns("D"))
            columnsToDelete.Delete
            ws.Cells(1,1).value = ws.Cells(2,2).value
    End Select

    ' Pronađi ćeliju koja sadrži tekst "ukupno obroka" i zameni sa "UKUPNO"
    searchText = "ukupno obroka"
    Set foundCell = ws.Cells.Find(What:=searchText, LookIn:=xlValues, LookAt:=xlPart)
    If Not foundCell Is Nothing Then
        foundCell.value = "UKUPNO"
    End If
    
    
    
    
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
    Dim maxAllowedFontSize As Double
    Dim calculatedFontSize As Double
    
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

    
    ' Nađi poslednji popunjeni red i poslednju popunjenu kolonu
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row              ' Poslednji red na osnovu druge kolone
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column    ' Poslednja kolona na osnovu drugog reda

    Dim regex As Object 
    Dim cellText As String
    Dim pattern As String
    
    ' Define the regex pattern
    pattern = "\(\d+\-(\d+)?[DRV]\)" ' Replace this with your regex pattern
    
    ' Create regex object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.Global = True ' Allow global matches (multiple matches in a cell)
    
    ' Loop through all cells
    For i = 1 To lastRow
        cellText = ws.Cells(i, 1).Value ' Get cell value
        If regex.Test(cellText) Then
            ' Remove regex matches
            ws.Cells(i, 1).Value = regex.Replace(cellText, "")
        End If
    Next i



    ' Konstant za računanje font size-a
    constant = 12 ' Možeš promeniti kako bi dobio optimalan rezultat

    ' Maksimalni font size
    maxFontSize = ws.Rows(2).RowHeight * 0.8 ' Postavi maksimalnu veličinu fonta

    ' Postavi maksimalnu dozvoljenu veličinu fonta
    maxAllowedFontSize = 72

    ' Prvi red - postavljanje font size
    For Each cell In ws.Rows(2).Cells
        If cell.value <> "" Then
            numChars = Len(cell.value)
            colWidth = ws.Columns(cell.Column).ColumnWidth

            ' Izračunaj font size na osnovu broja karaktera i širine kolone
            fontSize = (colWidth / numChars) * constant

            ' Ograniči maksimalnu veličinu fonta
            If fontSize > maxFontSize Then fontSize = Round(maxFontSize)
            If lastRow <6 And numChars <11 Then fontSize = maxAllowedFontSize

            ' Postavi font size
            cell.Font.Size = fontSize
        End If
    Next cell

    ' Prva kolona - postavljanje font size
    For Each cell In ws.Columns(1).Cells
        If cell.value <> "" Then
            numChars = Len(cell.value)
            colWidth = ws.Columns(1).ColumnWidth

            ' Izračunaj font size na osnovu broja karaktera i širine kolone
            fontSize = (colWidth / numChars) * constant

            ' Ograniči maksimalnu veličinu fonta
            If fontSize > maxFontSize Then fontSize = Round(maxFontSize)
            If lastRow <6 And numChars <11 Then fontSize = maxAllowedFontSize

            ' Postavi font size
            cell.Font.Size = fontSize
        End If
    Next cell

    ' Definiši opseg za promenu fonta (od B2 do poslednje ćelije)
    Set tableRange = ws.Range(ws.Cells(3, 2), ws.Cells(lastRow, lastCol))

    ' Promeni font za sve ćelije u tabeli
    If ws.Rows(2).RowHeight > maxAllowedFontSize Then
            calculatedFontSize = maxAllowedFontSize
        Else
            calculatedFontSize = ws.Rows(2).RowHeight
        End If

    With tableRange
        .Font.Name = "Calibri"              ' Postavi font
        .Font.Size = calculatedFontSize     ' Prilagođena veličina fonta prema visini reda
        .Font.Bold = True                   ' Postavi bold
        .HorizontalAlignment = xlCenter     ' Horizontalno centriranje
        .VerticalAlignment = xlCenter       ' Vertikalno centriranje
    End With
    
    
    Dim shadingTableRange As Range
    Dim rowIndex As Long
    Dim headerColor As Long
    Dim shadeColor As Long
    Dim startRow As Long
    
    ' Odredi opseg tabele (počevši od drugog reda)
    startRow = 2
    Set shadingTableRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count))
    
    ' Definiši boje
    headerColor = RGB(200, 200, 200) ' Tamno siva za zaglavlje
    shadeColor = RGB(230, 230, 230)  ' Svetlo siva za šrafirane redove
    
    ' Formatiraj zaglavlje (drugi red)
    With ws.UsedRange.Rows(startRow)
        .Interior.Color = headerColor
    End With
    
    ' Alternativno šrafiranje ostatka tabele
    For rowIndex = startRow + 1 To shadingTableRange.Rows.Count + startRow - 1
        If rowIndex Mod 2 = 0 Then
            ws.UsedRange.Rows(rowIndex).Interior.Color = shadeColor
        Else
            ws.UsedRange.Rows(rowIndex).Interior.Color = xlNone ' Bela pozadina
        End If
    Next rowIndex

    ' Ako je poslednja kolona "UKUPNO", primeni željeni font i veličinu fonta
    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
    If ws.Cells(3, lastCol).Value = "UKUPNO" Then
        ws.UsedRange.Columns(lastCol).Interior.Color = RGB(200,200,200)
        With ws.UsedRange.Columns(lastCol)
            .Font.Size = calculatedFontSize   ' Postavi željenu veličinu fonta
            .HorizontalAlignment = xlCenter     ' Centriraj tekst
            .VerticalAlignment = xlCenter       ' Centriraj tekst vertikalno
            .Font.Bold = True                   ' Ako treba, postavi bold
        End With
    End If
    
    ' Podesi treći red
    Set thirdRow = ws.Rows(3)
    With thirdRow
        .Font.Size = 30                     ' Postavi veličinu fonta
        .RowHeight = 35                     ' Postavi visinu reda
        .HorizontalAlignment = xlCenter     ' Horizontala centriranja
        .VerticalAlignment = xlCenter       ' Vertikalna centriranja
        .Font.Bold = False                  ' Ukloni bold
    End With
    ' Podesi header
    With firstRow
        .Font.Bold = True                    ' Boldiranje
        .HorizontalAlignment = xlCenter      ' Horizontala centriranja
        .VerticalAlignment = xlCenter        ' Vertikalna centriranja
    End With
    If(lastRow < 6) Then firstRow.Font.Size = maxAllowedFontSize

End Sub