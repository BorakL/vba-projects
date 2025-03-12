Public Const prviRed As Integer = 11
Public Const zadnjiRedKriterijum As String = "UKUPNO:"
Public Const prvaKolona As Integer = 1
Public Const zadnjaKolona As Integer = 3
Public Const prvaKolonaOznaka As String = "A"
Public Const zadnjaKolonaOznaka As String = "C"

Function PronadjiPoslednjiRed() As Long
    Dim ws As Worksheet
    Dim poslednjiRed As Long
    
    ' Postavljanje reference na aktivni radni list
    Set ws = ActiveSheet ' Možeš promeniti na konkretan list ako je potrebno
    
    ' Početni red je fiksiran na 11
    poslednjiRed = prviRed
    
    ' Pronalazak poslednjeg reda tabele (gde u koloni A piše "UKUPNO:")
    Do While ws.Cells(poslednjiRed, prvaKolona).value <> zadnjiRedKriterijum And Not IsEmpty(ws.Cells(poslednjiRed, prvaKolona).value)
        poslednjiRed = poslednjiRed + 1
    Loop
    
    ' Vraćanje poslednjeg reda
    PronadjiPoslednjiRed = poslednjiRed
End Function

Sub ObojiRedovePoKriterijumima(kriterijumi As Variant)
    Dim ws As Worksheet
    Dim poslednjiRed As Long
    Dim i As Long, j As Long
    Dim pronadjeno As Boolean
    pronadjeno = False ' Postavljamo na False dok ne nađemo bar jedan kriterijum
    
    ' Postavljanje reference na aktivni radni list
    Set ws = ActiveSheet
    
    ' Pronalazi poslednji popunjeni red
    poslednjiRed = PronadjiPoslednjiRed()
    
    ' Iteracija kroz kolonu A (A11:A(poslednjiRed))
    For i = prviRed To poslednjiRed
        ' Provera za svaki kriterijum u nizu
        For j = LBound(kriterijumi) To UBound(kriterijumi)
            ' Proverava da li tekst u celiji sadrži kriterijum (ne mora da bude tačno)
            If InStr(1, ws.Cells(i, prvaKolona).value, kriterijumi(j), vbTextCompare) > 0 Then
                ' Ako pronađe kriterijum, boji ceo red (A do C) u svetložutu boju
                ws.Range(prvaKolonaOznaka & i & ":" & zadnjaKolonaOznaka & i).Interior.Color = RGB(255, 255, 153)
                pronadjeno = True ' Obeležavamo da je bar jedan kriterijum pronađen
                Exit For ' Prekidamo unutrašnju petlju jer smo već našli podudaranje
            End If
        Next j
    Next i
    
    ' Ako ništa nije pronađeno, prikaži poruku
    If Not pronadjeno Then
        MsgBox "Ni jedan od navedenih kriterijuma nije pronađen.", vbInformation, "Obaveštenje"
    End If
End Sub

Sub ProveriPodatke()
    Dim ws As Worksheet
    Dim i As Long
    Dim j As Integer
    Dim cellValue As String
    Dim foundItems As Object
    Dim key As Variant
    Dim message As String 
    Dim poslednjiRed As Long
    Dim opseg As Range
    Dim dataObj As Object

    ' Postavljanje reference na radni list
    Set ws = ActiveSheet ' Može se promeniti u konkretan list, npr. ThisWorkbook.Sheets("Sheet1")
    
    poslednjiRed = PronadjiPoslednjiRed()
    
    ' Kreiranje rečnika sa ključnim rečima i porukama
    Set foundItems = CreateObject("Scripting.Dictionary")
    foundItems.Add "BS", "Ima bistra supa"
    foundItems.Add "DB", "Ima dnevna bolnica"
    foundItems.Add "VAN RFZO", "Ima van RFZO"
    foundItems.Add "DNEVNA", "Ima dnevna bolnica"
    foundItems.Add "M-D", "Ima mleko"
    foundItems.Add "HD", "Ima HD. Izdvoji ako je KOŽNO!"
    foundItems.Add ChrW(268) & "-D", "Ima čaj" ' ASCII karakter za Č

    ' Kreiranje skupa za pronađene stavke
    Dim results As Object
    Set results = CreateObject("Scripting.Dictionary")

    ' Iteracija kroz kolone A do C u opsegu od prvog do poslednjeg reda
    For i = prviRed To poslednjiRed - 1
        For j = 1 To 3 ' Kolone A (1), B (2), C (3)
            cellValue = ws.Cells(i, j).value

            ' Ako je vrednost "HEMODIJALIZA SENDVIČI", promeni u "DNEVNA BOLNICA"
            If InStr(1, cellValue, "HEMODIJALIZA SENDVI" & ChrW(268) & "I", vbTextCompare) > 0 Then
                cellValue = Replace(cellValue, "HEMODIJALIZA SENDVI" & ChrW(268) & "I", "DNEVNA BOLNICA", 1, -1, vbTextCompare)
                ws.Cells(i, j).value = cellValue
                MsgBox "HEMODIJALIZA SENDVICI je prepravljen u DNEVNA BOLNICA! Sačuvaj fajl."
            End If
            
            ' Provera svih ključnih reči
            For Each key In foundItems.Keys
                If InStr(1, cellValue, key, vbTextCompare) > 0 Then
                    If Not results.Exists(key) Then
                        results.Add key, foundItems(key)
                    End If
                End If
            Next key
        Next j
    Next i

    ' Formiranje i prikaz poruke ako su pronađeni rezultati
    If results.Count > 0 Then
        message = "Otpremnica sadrži:" & vbCrLf
        For Each key In results.Keys
            message = message & "- " & results(key) & vbCrLf
        Next key
        MsgBox message, vbInformation, "Rezultat provere"
    Else
        MsgBox "Nema pronađenih stavki u otpremnici.", vbInformation, "Rezultat provere"
    End If
End Sub

Sub IzdvojSpecijalneObroke()
    Dim kriterijumi As Variant
    kriterijumi = Array("BS", "M-D", "HD", ChrW(268) & "-D") ' Kriterijumi za specijalne obroke
    Call ObojiRedovePoKriterijumima(kriterijumi)
End Sub

Sub IzdvojVRFZO()
    Dim kriterijumi As Variant
    kriterijumi = Array("VAN RFZO") ' Kriterijum za bistru supu
    Call ObojiRedovePoKriterijumima(kriterijumi)
End Sub

Sub IzdvojDnevnuBolnicu()
    Dim kriterijumi As Variant
    kriterijumi = Array("DB", "DNEVNA") ' Kriterijum za bistru supu
    Call ObojiRedovePoKriterijumima(kriterijumi)
End Sub


Sub AzurirajSumu()
    Dim ws As Worksheet
    Dim i As Long
    Dim poslednjiRed As Long
    Dim suma As Double
    Dim staraVrednost As Double
    Dim sumaCelija As Range
    
    ' Postavljanje reference na aktivni radni list
    Set ws = ActiveSheet

    ' Pronalazi poslednji red
    poslednjiRed = PronadjiPoslednjiRed()

    ' Postavljanje reference na ćeliju gde je suma (pretpostavljam da je poslednji red, kolona C)
    Set sumaCelija = ws.Cells(poslednjiRed, zadnjaKolona)

    ' Čuvanje stare vrednosti sume
    staraVrednost = sumaCelija.value

    ' Inicijalizacija sume
    suma = 0

    ' Iteracija kroz redove od 11 do poslednjiRed - 1
    For i = prviRed + 1 To poslednjiRed - 1 ' Ne uključuje poslednji red (gde je UKUPNO)
        ws.Range(prvaKolonaOznaka & i & ":" & zadnjaKolonaOznaka & i).Interior.ColorIndex = xlNone   ' Brišemo boju pozadine svih ćelija u redu (A do C)
        suma = suma + ws.Cells(i, zadnjaKolona).value  ' Sabira vrednosti iz kolone C (3. kolona)
    Next i

    ' Ažuriranje sume u tabeli
    sumaCelija.value = suma

    ' Prikazivanje poruke sa starom i novom vrednošću
    MsgBox "Vrednost sume je promenjena sa " & staraVrednost & " na " & suma, vbInformation, "Ukupna Suma"
End Sub


Sub StampajOtpremnicu()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Možeš promeniti na konkretan sheet ako je potrebno
    
    With ws.PageSetup
        .Zoom = False ' Isključuje ručno skaliranje
        .FitToPagesWide = 1 ' Smanjuje širinu na 1 stranicu
        .FitToPagesTall = 1 ' Smanjuje visinu na 1 stranicu
    End With
    
    ' Štampa sve na jednu stranicu, dva primerka
    ws.PrintOut Copies:=2
End Sub


Sub removeRDV()
    Dim regex As Object
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim cellText As String
    Dim pattern As String
    
    ' Definiši regex pattern
    pattern = "\(\d+\-\d*[DRV]\)" ' Pokriva sve slučajeve
    
    ' Kreiraj regex objekat
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.Global = True ' Omogućava zamene u celoj ćeliji
    
    ' Koristi ActiveSheet umesto fiksnog sheet-a
    Set ws = ActiveSheet
    
    ' Pronađi poslednji red i kolonu
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Iteriraj kroz sve ćelije
    For i = 1 To lastRow
        For j = 1 To lastCol
            If Not IsEmpty(ws.Cells(i, j).value) And Not IsError(ws.Cells(i, j).value) Then
                cellText = CStr(ws.Cells(i, j).value) ' Konvertuj u string
                
                ' Debugging
                Debug.Print "Original: " & cellText
                
                If regex.Test(cellText) Then
                    ws.Cells(i, j).value = regex.Replace(cellText, "")
                    Debug.Print "Izmenjeno: " & ws.Cells(i, j).value
                End If
            End If
        Next j
    Next i
    
    ' Oslobodi objekat
    Set regex = Nothing

End Sub