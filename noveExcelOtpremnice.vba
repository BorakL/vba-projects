Function PronadjiPoslednjiRed() As Long
    Dim ws As Worksheet
    Dim poslednjiRed As Long
    
    ' Postavljanje reference na aktivni radni list
    Set ws = ActiveSheet ' Možeš promeniti na konkretan list ako je potrebno
    
    ' Početni red je fiksiran na 11
    poslednjiRed = 11
    
    ' Pronalazak poslednjeg reda tabele (gde u koloni A piše "UKUPNO:")
    Do While ws.Cells(poslednjiRed, 1).value <> "UKUPNO:" And Not IsEmpty(ws.Cells(poslednjiRed, 1).value)
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
    For i = 11 To poslednjiRed
        ' Provera za svaki kriterijum u nizu
        For j = LBound(kriterijumi) To UBound(kriterijumi)
            ' Proverava da li tekst u celiji sadrži kriterijum (ne mora da bude tačno)
            If InStr(1, ws.Cells(i, 1).value, kriterijumi(j), vbTextCompare) > 0 Then
                ' Ako pronađe kriterijum, boji ceo red (A do C) u svetložutu boju
                ws.Range("A" & i & ":C" & i).Interior.Color = RGB(255, 255, 153)
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
    Dim prviRed As Long
    Dim poslednjiRed As Long
    Dim opseg As Range
    Dim dataObj As Object

    ' Postavljanje reference na radni list
    Set ws = ActiveSheet ' Može se promeniti u konkretan list, npr. ThisWorkbook.Sheets("Sheet1")
    
    prviRed = 11
    poslednjiRed = PronadjiPoslednjiRed()
    
    ' Kreiranje rečnika sa ključnim rečima i porukama
    Set foundItems = CreateObject("Scripting.Dictionary")
    foundItems.Add "BS", "Ima bistra supa"
    foundItems.Add "DB", "Ima dnevna bolnica"
    foundItems.Add "VAN RFZO", "Ima van RFZO"
    foundItems.Add "DNEVNA", "Ima dnevna usluga"
    foundItems.Add "M-D", "Ima mleko"
    foundItems.Add ChrW(268) & "-D", "Ima čaj" ' ASCII karakter za Č
    foundItems.Add "HEMODIJALIZA SENDVI" & ChrW(268) & "I", "Ima hemodijaliza sendviči. Ako je Punkt 1 prepravi u DNEVNA BOLNICA"

    ' Kreiranje skupa za pronađene stavke
    Dim results As Object
    Set results = CreateObject("Scripting.Dictionary")

    ' Iteracija kroz kolone A do C u opsegu od prvog do poslednjeg reda
    For i = prviRed To poslednjiRed - 1
        For j = 1 To 3 ' Kolone A (1), B (2), C (3)
            cellValue = ws.Cells(i, j).value
            
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
    kriterijumi = Array("BS", "M-D", ChrW(268) & "-D") ' Kriterijumi za specijalne obroke
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

Sub IzdvojHemodijalizaSendvici()
    Dim kriterijumi As Variant
    kriterijumi = Array("HEMODIJALIZA SENDVI" & ChrW(268) & "I")
    Call ObojiRedovePoKriterijumima(kriterijumi)
End Sub

Sub AzurirajSumu()
    Dim ws As Worksheet
    Dim i As Long
    Dim poslednjiRed As Long
    Dim suma As Double

    ' Postavljanje reference na aktivni radni list
    Set ws = ActiveSheet

    ' Pronalazi poslednji red
    poslednjiRed = PronadjiPoslednjiRed()

    ' Inicijalizacija sume
    suma = 0

    ' Iteracija kroz redove od 11 do poslednjiRed - 1
    For i = 11 + 1 To poslednjiRed - 1 ' Ne uključuje poslednji red (gde je UKUPNO)
        ws.Range("A" & i & ":C" & i).Interior.ColorIndex = xlNone   ' Brišemo boju pozadine svih ćelija u redu (A do C)
        suma = suma + ws.Cells(i, 3).value  ' Sabira vrednosti iz kolone C (3. kolona)
    Next i

    ' Prikazivanje nove vrednosti sume u poruci
    MsgBox "Nova suma iz kolone C je: " & suma, vbInformation, "Ukupna Suma"
End Sub


Sub StampajOtpremnicu()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Možeš promeniti na konkretan sheet ako je potrebno

    ' Štampa samo prvu stranicu, dva primerka
    ws.PrintOut From:=1, To:=1, Copies:=2
End Sub
