Sub ProveriIZbirajDrugojTabeli()
    Dim doc As Document
    Dim tbl As Table
    Dim brojRedova As Long
    Dim brojKolona As Long
    Dim zadnjaKolona As Long
    Dim suma As Long
    Dim i As Long
    Dim poslednjaCelijaPrvaKolona As String
    Dim prethodnaSuma As Long
    Dim novaSuma As Long
    Dim tekstNakonPrveTabele As String
    Dim rangeIspodTabele As Range
    
    ' Provera da li je otvoren dokument
    If Documents.Count = 0 Then
        MsgBox "Nema otvorenih Word dokumenata.", vbExclamation, "Greška"
        Exit Sub
    End If
    
    ' Rad sa aktivnim dokumentom
    Set doc = ActiveDocument
    
    ' Provera broja tabela u dokumentu
    If doc.Tables.Count < 3 Then
        MsgBox "Dokument nema tri tabele. Proverite format dokumenta.", vbExclamation, "Neispravan format"
        Exit Sub
    End If
    
    ' Fokusiranje na prvu tabelu
    Set tbl = doc.Tables(1)
    
    ' Dobijanje teksta odmah ispod prve tabele
    Set rangeIspodTabele = tbl.Range
    rangeIspodTabele.Collapse Direction:=wdCollapseEnd ' Postavljamo opseg na kraj tabele
    tekstNakonPrveTabele = Trim(rangeIspodTabele.Text)
    tekstNakonPrveTabele = Replace(tekstNakonPrveTabele, Chr(13) & Chr(7), "") ' Uklanja znakove za kraj reda
    
    ' Provera da li tekst ispod prve tabele sadrži reč "OTPREMNICA"
    If InStr(1, rangeIspodTabele.Next(Unit:=wdParagraph, Count:=1).Text, "OTPREMNICA", vbTextCompare) = 0 Then
        MsgBox "Ovaj dokument ne sadrži rec OTPREMNICA!", vbExclamation, "Neispravan dokument"
        Exit Sub
    End If
    
    ' Fokusiranje na drugu tabelu
    Set tbl = doc.Tables(2)
    
    ' Dobijanje broja redova i kolona
    brojRedova = tbl.Rows.Count
    brojKolona = tbl.Columns.Count
    zadnjaKolona = brojKolona
    
    ' Provera poslednje ćelije u prvoj koloni (SUMA)
    poslednjaCelijaPrvaKolona = Trim(tbl.Cell(brojRedova, 1).Range.Text)
    poslednjaCelijaPrvaKolona = Replace(poslednjaCelijaPrvaKolona, Chr(13) & Chr(7), "") ' Uklanja znakove za kraj reda
    
    If Not LCase(poslednjaCelijaPrvaKolona) Like "suma*" Then
        MsgBox "Poslednji red prve kolone ne sadrži tekst 'SUMA'. Proverite format dokumenta.", vbExclamation, "Neispravan format"
        Exit Sub
    End If
    
    ' Dohvatanje prethodne vrednosti SUMA iz zadnje kolone
    prethodnaSuma = Val(tbl.Cell(brojRedova, zadnjaKolona).Range.Text)
    
    ' Provera i sabiranje brojeva u zadnjoj koloni
    suma = 0
    For i = 1 To brojRedova - 1 ' Prolazimo kroz sve redove osim poslednjeg
        On Error Resume Next ' Ignorisanje greške ako ćelija nije broj
        suma = suma + Val(tbl.Cell(i, zadnjaKolona).Range.Text)
        On Error GoTo 0
    Next i
    
    ' Ako se nova suma razlikuje od prethodne, ažuriraj SUMA i obavesti
    If suma <> prethodnaSuma Then
        tbl.Cell(brojRedova, zadnjaKolona).Range.Text = suma
        MsgBox "Suma obroka je ažurirana sa " & prethodnaSuma & " na " & suma & ".", vbInformation, "Ažuriranje SUMA"
    End If
End Sub