

Sub IzdvojVRFZO()
    Dim doc As Document
    Dim tbl As Table
    Dim brojRedova As Integer
    Dim tekstPrveKolone As String
    Dim i As Integer
    Dim foundVRFZO As Boolean
    
    ' Referenca na aktivni dokument
    Set doc = ActiveDocument
    
    ' Provera da li postoje tri tabele
    If doc.Tables.Count < 3 Then
        MsgBox "Dokument mora sadržati najmanje tri tabele. Neispravan format dokumenta.", vbExclamation, "Greška"
        Exit Sub
    End If
    
    ' Referenca na drugu tabelu
    Set tbl = doc.Tables(2)
    brojRedova = tbl.Rows.Count
    foundVRFZO = False
    
    ' Otkrij sve redove u tabeli
    For i = 1 To brojRedova
        tbl.Rows(i).Range.Font.Hidden = False
    Next i
    
    ' Provera da li postoji "VAN RFZO" u prvoj koloni
    For i = 1 To brojRedova - 1
        tekstPrveKolone = Trim(tbl.cell(i, 1).Range.Text)
        tekstPrveKolone = Replace(tekstPrveKolone, Chr(13) & Chr(7), "") ' Uklanja znakove za kraj reda
        
        If InStr(1, tekstPrveKolone, "VAN RFZO", vbTextCompare) > 0 Then
            foundVRFZO = True
            Exit For ' Nije potrebno dalje pretraživati
        End If
    Next i
    
    ' Ako nema "VAN RFZO", prijavi grešku i prekini
    If Not foundVRFZO Then
        MsgBox "Ni jedan obrok nije van RFZO-a!", vbExclamation, "Greška"
        Exit Sub
    End If
    
    ' Iteracija kroz redove i sakrivanje nepotrebnih
    For i = 1 To brojRedova - 1
        tekstPrveKolone = Trim(tbl.cell(i, 1).Range.Text)
        tekstPrveKolone = Replace(tekstPrveKolone, Chr(13) & Chr(7), "") ' Uklanja znakove za kraj reda
        
        ' Sakrij redove koji ne sadrže "VAN RFZO", osim zadnjeg reda
        If InStr(1, tekstPrveKolone, "VAN RFZO", vbTextCompare) = 0 And i <> brojRedova Then
            tbl.Rows(i).Range.Font.Hidden = True
        End If
    Next i
    
    ' Ažuriranje vrednosti Suma
    Call UpdateSum(tbl)
End Sub







Sub IzdvojBsDbCM()
    Dim doc As Document
    Dim tbl As Table
    Dim brojRedova As Integer
    Dim tekstPrveKolone As String
    Dim i As Integer
    Dim keywords As Variant
    Dim foundKeyword As Boolean
    
    ' Ključne reči koje tražimo u tabeli
    keywords = Array("BS", "M-D", ChrW(268) & "-D", "DNEVNA")
    
    ' Referenca na aktivni dokument
    Set doc = ActiveDocument
    
    ' Provera da li postoje tri tabele
    If doc.Tables.Count < 3 Then
        MsgBox "Dokument mora sadržati najmanje tri tabele. Neispravan format dokumenta.", vbExclamation, "Greška"
        Exit Sub
    End If
    
    ' Referenca na drugu tabelu
    Set tbl = doc.Tables(2)
    brojRedova = tbl.Rows.Count
    
    ' Prvo otkrivanje (unhide) svih redova u tabeli
    For i = 1 To brojRedova
        tbl.Rows(i).Range.Font.Hidden = False
    Next i
    
    ' Provera da li postoji red sa ključnim rečima
    Dim hasValidRow As Boolean
    hasValidRow = False
    
    For i = 1 To brojRedova - 1
        tekstPrveKolone = Trim(tbl.cell(i, 1).Range.Text)
        tekstPrveKolone = Replace(tekstPrveKolone, Chr(13) & Chr(7), "") ' Uklanja znakove za kraj reda
        
        ' Provera da li red sadrži neku od ključnih reči
        For Each keyword In keywords
            If InStr(1, tekstPrveKolone, keyword, vbTextCompare) > 0 Then
                hasValidRow = True
                Exit For
            End If
        Next keyword
        If hasValidRow Then Exit For
    Next i
    
    ' Ako nijedan red ne sadrži ključne reči, prijavi grešku i prekini
    If Not hasValidRow Then
        MsgBox "Ni jedan obrok ne odgovara traženim kriterijumima!", vbExclamation, "Greška"
        Exit Sub
    End If
    
    ' Iteracija kroz redove i sakrivanje nepotrebnih
    For i = 1 To brojRedova - 1
        tekstPrveKolone = Trim(tbl.cell(i, 1).Range.Text)
        tekstPrveKolone = Replace(tekstPrveKolone, Chr(13) & Chr(7), "") ' Uklanja znakove za kraj reda
        
        foundKeyword = False
        For Each keyword In keywords
            If InStr(1, tekstPrveKolone, keyword, vbTextCompare) > 0 Then
                foundKeyword = True
                Exit For
            End If
        Next keyword
        
        ' Sakrij red ako ne sadrži ključne reči i nije poslednji red
        If Not foundKeyword And i <> brojRedova Then
            tbl.Rows(i).Range.Font.Hidden = True
        End If
    Next i
    
    ' Ažuriranje vrednosti Suma
    Call UpdateSum(tbl)
     
End Sub






Sub IzbaciBsDbVrfzo()
    Dim doc As Document
    Dim tbl As Table
    Dim brojRedova As Integer
    Dim tekstPrveKolone As String
    Dim i As Integer
    Dim keywords As Variant
    Dim suma As Long
    
    ' Ključne reči koje treba sakriti
    keywords = Array("VAN RFZO", "BS", "M-D", ChrW(268) & "-D", "DNEVNA")
    
    ' Referenca na aktivni dokument
    Set doc = ActiveDocument
    
    ' Provera da li postoje tri tabele
    If doc.Tables.Count < 3 Then
        MsgBox "Dokument mora sadržati najmanje tri tabele. Neispravan format dokumenta.", vbExclamation, "Greška"
        Exit Sub
    End If
    
    ' Referenca na drugu tabelu
    Set tbl = doc.Tables(2)
    brojRedova = tbl.Rows.Count
    
    ' Prvo otkrivanje (unhide) svih redova u tabeli
    For i = 1 To brojRedova
        tbl.Rows(i).Range.Font.Hidden = False
    Next i
    
    ' Iteracija kroz redove i sakrivanje onih sa ključnim rečima
    For i = 1 To brojRedova - 1
        tekstPrveKolone = Trim(tbl.cell(i, 1).Range.Text)
        tekstPrveKolone = Replace(tekstPrveKolone, Chr(13) & Chr(7), "") ' Uklanja znakove za kraj reda
        
        Dim foundKeyword As Boolean
        foundKeyword = False
        
        ' Provera da li red sadrži neku od ključnih reči
        For Each keyword In keywords
            If InStr(1, tekstPrveKolone, keyword, vbTextCompare) > 0 Then
                foundKeyword = True
                Exit For
            End If
        Next keyword
        
        ' Sakrij red ako sadrži ključnu reč
        If foundKeyword Then
            tbl.Rows(i).Range.Font.Hidden = True
        End If
    Next i
    
    ' Ažuriranje sume koristeći funkciju UpdateSum
    Call UpdateSum(tbl)
     
End Sub




Sub ProveriOtpremnicu()
    Dim doc As Document
    Dim tbl1 As Table, tbl2 As Table
    Dim rng As Range, searchRange As Range
    Dim kriterijumi As Object
    Dim kljuc As Variant
    Dim regex As Object
    Dim foundDrDragi As Boolean
    Dim brojRedova As Integer
     
    ' Kreiranje RegEx objekta
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = False
    regex.Pattern = "DR\s+DRAGI" ' RegEx za pretragu "DR DRAGI"
    
    ' Referenca na aktivni dokument
    Set doc = ActiveDocument
    
    ' Provera da li postoje bar dve tabele
    If doc.Tables.Count < 2 Then
        MsgBox "Dokument mora imati bar dve tabele.", vbExclamation
        Exit Sub
    End If
    
    ' Referenca na prvu i drugu tabelu
    Set tbl1 = doc.Tables(1)
    Set tbl2 = doc.Tables(2)
    
    brojRedova = tbl2.Rows.Count
    ' Prvo otkrivanje (unhide) svih redova u tabeli
    For i = 1 To brojRedova
        tbl2.Rows(i).Range.Font.Hidden = False
    Next i
    
    ' Podešavanje pretrage između prve i druge tabele
    Set searchRange = doc.Range(tbl1.Range.End, tbl2.Range.Start)
    
    ' Provera da li "DR DRAGI" postoji u tekstu između tabela
    foundDrDragi = regex.Test(searchRange.Text)
    
    If foundDrDragi Then
        ' Kreiranje kriterijuma za zamenu
        Set kriterijumi = CreateObject("Scripting.Dictionary")
        kriterijumi.Add "KLINIKA B", "INTERNA B"

        ' Iteracija kroz kriterijume za zamenu
        For Each kljuc In kriterijumi.Keys
            Set rng = searchRange.Duplicate ' Kopira opseg za pretragu
            With rng.Find
                .Text = kljuc
                .Replacement.Text = kriterijumi(kljuc)
                .MatchCase = False ' Ignoriše velika/mala slova
                .Wrap = wdFindStop ' Pretraga se zaustavlja unutar definisanog opsega
                .Execute Replace:=wdReplaceAll ' Zamenjuje sve instance pronađenog teksta
            End With
        Next kljuc
    End If
    
    Dim row As row
    Dim cell As cell
    Dim foundItems As Collection
    Dim message As String
    
    ' Kreiranje RegEx objekta
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = False ' Traži prvi pogodak u tekstu
    
    ' Kreiranje kriterijuma za pretragu i poruka
    Set kriterijumi = CreateObject("Scripting.Dictionary")
    kriterijumi.Add "BS", "BISTRA SUPA"
    kriterijumi.Add "VAN RFZO", "VAN RFZO"
    kriterijumi.Add "DNEVNA", "DNEVNA BOLNICA"
    kriterijumi.Add ChrW(268) & "-D", "CAJ"
    kriterijumi.Add "M-D", "MLEKO"
    
    ' Kolekcija za pronalaženje pogodaka
    Set foundItems = New Collection
    
    ' Referenca na aktivni dokument
    Set doc = ActiveDocument
    
    ' Provera da li dokument ima bar jednu tabelu
    If doc.Tables.Count = 0 Then
        MsgBox "Dokument ne sadrži tabele.", vbExclamation
        Exit Sub
    End If
    
    ' Referenca na prvu tabelu
    Set tbl = doc.Tables(2)
    
    ' Iteracija kroz redove u tabeli
    For Each row In tbl.Rows
        For Each cell In row.Cells
            Dim cellText As String
            cellText = Trim(cell.Range.Text)
            cellText = Replace(cellText, Chr(13) & Chr(7), "") ' Uklanja specijalne znakove
            
            ' Provera svakog kriterijuma
            For Each kljuc In kriterijumi.Keys
                regex.Pattern = kljuc
                If regex.Test(cellText) Then
                    On Error Resume Next ' Sprečava duplikate u kolekciji
                    foundItems.Add kriterijumi(kljuc), CStr(kriterijumi(kljuc))
                    On Error GoTo 0
                End If
            Next kljuc
        Next cell
    Next row
    
    ' Formiranje poruke
    If foundItems.Count > 0 Then
        message = "Otpremnica sadrži:" & vbCrLf
        For Each kljuc In foundItems
            message = message & "- " & kljuc & vbCrLf
        Next kljuc
        MsgBox message, vbInformation
    End If
    
    ' Ažuriranje sume koristeći funkciju UpdateSum
    Call UpdateSum(tbl2)
End Sub
 




Function UpdateSum(tbl As Table)
    Dim suma As Long
    Dim i As Integer
    Dim broj As String
    
    suma = 0 ' Resetovanje sume na početku
    
    ' Iteracija kroz redove i sabiranje vrednosti iz poslednje kolone vidljivih redova
    For i = 1 To tbl.Rows.Count - 1
        If tbl.Rows(i).Range.Font.Hidden = False Then
            broj = Trim(tbl.cell(i, tbl.Columns.Count).Range.Text)
            broj = Replace(broj, Chr(13) & Chr(7), "") ' Uklanja znakove za kraj reda
            
            If IsNumeric(broj) Then
                suma = suma + CLng(broj)
            End If
        End If
    Next i
    
    ' Ažuriranje vrednosti u poslednjem redu (SUMA)
    tbl.cell(tbl.Rows.Count, tbl.Columns.Count).Range.Text = suma
End Function





Sub StampajOtpremnicu()
'
' StampajOtpremnicu Macro
'
    Dim currentPage As Long
    Dim totalPages As Long
    Dim rng As Range
    Dim printRange As String

    ' Dobijanje reference na aktivni dokument
    Dim doc As Document
    Set doc = ActiveDocument

    ' Provera da li je dokument otvoren
    If doc Is Nothing Then
        MsgBox "Nema otvorenog dokumenta.", vbExclamation
        Exit Sub
    End If

    ' Dobijanje trenutne stranice
    Set rng = Selection.Range
    currentPage = rng.Information(wdActiveEndPageNumber)

    ' Dobijanje ukupnog broja stranica
    totalPages = doc.BuiltInDocumentProperties(wdPropertyPages)

    ' Provera da li je trenutna stranica validna
    If currentPage < 1 Or currentPage > totalPages Then
        MsgBox "Ne postoji trenutna stranica za štampu.", vbExclamation
        Exit Sub
    End If

    ' Formatiranje opsega za štampu (trenutna stranica)
    printRange = CStr(currentPage)

    ' Štampanje trenutne stranice u 2 primerka
    Application.PrintOut Range:=wdPrintRangeOfPages, Pages:=printRange, Copies:=2

    MsgBox "Trenutna stranica (" & printRange & ") je odštampana u 2 primerka.", vbInformation
End Sub





Sub SacuvajDokument()
    Dim doc As Document
    Dim tbl1 As Table, tbl2 As Table
    Dim searchRange As Range
    Dim lines As Variant, line As String
    Dim keyValue As Variant
    Dim key As String, value As String
    Dim klinika As String, dan As String, obrok As String, dodatak As String
    Dim fileName As String, folderPath As String
    Dim regex As Object
    Dim lineIndex As Long

    ' Kreiranje RegEx objekta
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True

    ' Referenca na aktivni dokument
    Set doc = ActiveDocument

    ' Provera da li postoje bar dve tabele
    If doc.Tables.Count < 2 Then
        MsgBox "Dokument mora sadržati bar dve tabele.", vbExclamation
        Exit Sub
    End If

    ' Referenca na prve dve tabele
    Set tbl1 = doc.Tables(1)
    Set tbl2 = doc.Tables(2)

    ' Tekst između prve i druge tabele
    Set searchRange = doc.Range(tbl1.Range.End, tbl2.Range.Start)
    Debug.Print searchRange.Text
    Dim cleanText As String
    cleanText = searchRange.Text
    
    ' Zameni sve vrste novih redova sa "_đ"
    cleanText = Replace(cleanText, vbCrLf, "_") ' Windows stil (\r\n)
    cleanText = Replace(cleanText, vbCr, "_")   ' Stari Mac stil (\r)
    cleanText = Replace(cleanText, vbLf, "_")   ' Unix/Linux stil (\n)
    cleanText = Replace(cleanText, Chr(11), "_") ' Vertical Tab ako postoji
    
    lines = Split(cleanText, "_")
    ' Inicijalizacija promenljivih
    klinika = ""
    dan = ""
    obrok = ""
    dodatak = ""
 
    ' Mapiranje kraćih naziva klinika
    Dim klinikaMap As Object
    Set klinikaMap = CreateObject("Scripting.Dictionary")
    klinikaMap.Add "KLINIKA ZA KARDIOHIRURGIJU", "KARDIOHIRURGIJA"
    klinikaMap.Add "KLINIKA ZA VASKULARNU HIRURGIJU", "VASKULARNA"
    klinikaMap.Add "KLINIKA ZA PULMOLOGIJU", "PULMOLOGIJA"
    klinikaMap.Add "INSTITUT ZA ORTOPEDIJU BANJICA", "BANJICA"
    klinikaMap.Add "KLINIKA ZA ORTOPEDSKU HIRURGIJU I TRAUMATOLOGIJU (A)", "ORTOPEDIJA A"
    klinikaMap.Add "KLINIKA ZA ENDOKRINOLOGIJU DIJABETES I BOLESTI METABOLIZMA", "ENDOKRINOLOGIJA"
    klinikaMap.Add "KLINIKA ZA KARDIOLOGIJU KLINI" & ChrW(268) & "KO ODELJENJE 3", "KARDIOLOGIJA"
    klinikaMap.Add "KLINIKA ZA NEUROLOGIJU", "NEUROLOGIJA"
    klinikaMap.Add "KLINIKA ZA ORL I MFH", "ORL I MFH"
    klinikaMap.Add "KLINIKA ZA DERMATOVENEROLOGIJU", "DERMATOVENEROLOGIJA"
    klinikaMap.Add "KLINIKA ZA O" & ChrW(268) & "NE BOLESTI", "O" & ChrW(268) & "NO"
    klinikaMap.Add "KLINIKA ZA PSIHIJATRIJU", "PSIHIJATRIJA"
    klinikaMap.Add "KBC BEŽANIJSKA KOSA", "BEŽANIJA"
    klinikaMap.Add "OPŠTA BOLNICA VALJEVO", "VALJEVO"
    klinikaMap.Add "INSTITUT ZA NEONATOLOGIJU", "NEONATOLOGIJA"
    klinikaMap.Add "KLINIKA ZA NEFROLOGIJU", "NEFROLOGIJA"
    klinikaMap.Add "INSTITUT ZA REUMATOLOGIJU", "REUMATOLOGIJA"
    klinikaMap.Add "KLINIKA ZA UROLOGIJU - PASTEROVA 2", "UROLOGIJA 2"
    klinikaMap.Add "KLINIKA ZA NEUROHIRURGIJU - PUNKT 1", "NEUROHIRURGIJA P1"
    klinikaMap.Add "KLINIKA ZA NEUROHIRURGIJU - PUNKT 2", "NEUROHIRURGIJA P2"
    klinikaMap.Add "KLINIKA ZA UROLOGIJU", "UROLOGIJA UKC"
    klinikaMap.Add "KLINIKA ZA OPEKOTINE, PLASTI" & ChrW(268) & "NU I REKONSTRUKTIVNU HIRURGIJU", "PLASTI" & ChrW(268) & "NA H"
    klinikaMap.Add "KLINIKA ZA ORTOPEDSKU HIRURGIJU I TRAUMATOLOGIJU (B)", "ORTOPEDIJA B"
    klinikaMap.Add "KLINIKA ZA GINEKOLOGIJU I AKUŠERSTVO", "GAK"
    klinikaMap.Add "INSTITUT ZA MENTALNO ZDRAVLJE", "PALMOTI" & Chr(196) & "EVA"
    klinikaMap.Add "INSTITUT ZA ONKOLOGIJU I RADIOLOGIJU SRBIJE", "ONKOLOGIJA"
    klinikaMap.Add "INTERNA KLINIKA", "A BLOK - INTERNA"
    klinikaMap.Add "HIRURGIJA", "A BLOK - HIRURGIJA"
    klinikaMap.Add "ORL", "A BLOK - ORL"
    klinikaMap.Add "PSIHIJATRIJA", "A BLOK - PSIHIJATRIJA"
    klinikaMap.Add "DE" & ChrW(268) & "IJA", "B BLOK - DE" & ChrW(268) & "IJA"
    klinikaMap.Add "UROLOGIJA", "B BLOK - UROLOGIJA"
    klinikaMap.Add "GINEKOLOGIJA", "B BLOK - GINEKOLOGIJA"
    klinikaMap.Add "KLINIKA B", "B BLOK - GERIJATRIJA"
    klinikaMap.Add "INTERNA B", "B BLOK - GERIJATRIJA"
    klinikaMap.Add "OPŠTA BOLNICA:" & Chr(34) & "STEFAN VISOKI" & Chr(34), "S PALANKA"
    ' Dodajte ostale unose prema vašoj tabeli
    
    ' Parsiranje linija
    For lineIndex = LBound(lines) To UBound(lines)
        line = Trim(lines(lineIndex))
    
        ' Proveri da li linija sadrži key:value par
        If InStr(line, ":") > 0 Then
            keyValue = Split(line, ":")
            If UBound(keyValue) >= 1 Then ' Proveri da li postoje dva dela
                key = Trim(keyValue(0))
                value = Trim(keyValue(1))
    
                Select Case key
                    Case "Bolnica"
                        ' Bolnica se ne koristi u obradi
                    Case "Klinika"
                        If klinikaMap.Exists(value) Then
                            klinika = klinikaMap(value)
                        Else
                            klinika = value ' Ako nema mapiranja, koristi originalni naziv
                        End If
                    Case "Dan"
                        dan = Replace(value, "-", ".") ' Promena formata datuma
                End Select
            End If
        ElseIf LCase(line) = "doru" & ChrW(269) & "ak" Or LCase(line) = "ru" & ChrW(269) & "ak" Or LCase(line) = "ve" & ChrW(269) & "era" Then
            obrok = line
        End If
    Next lineIndex

    ' Provera druge tabele za dodatak
    Dim row As row
    Dim cell As cell
    Dim cellText As String

    For Each row In tbl2.Rows
        For Each cell In row.Cells
            cellText = Replace(cell.Range.Text, Chr(13) & Chr(7), "")
            cellText = Trim(cellText)

            If InStr(cellText, "VAN RFZO") > 0 Then
                dodatak = " VRFZO"
            ElseIf InStr(cellText, "BS") > 0 Then
                dodatak = " BS"
            ElseIf InStr(cellText, "M-D") > 0 Then
                dodatak = "M"
            ElseIf InStr(cellText, ChrW(268) & "-D") > 0 Then
                dodatak = " Caj"
            ElseIf InStr(cellText, "DNEVNA") > 0 Then
                dodatak = " DB"
            End If
        Next cell
    Next row

    ' Kreiranje naziva fajla i putanje
    folderPath = "C:\Users\Luka\Desktop\Otpremnice" & obrok & "\"
    fileName = klinika & dodatak & " " & dan
 
    ' Kreiraj punu putanju
    Dim fullPath As String
    
    fullPath = folderPath & fileName
    CopyToClipboardShell fileName
    
   
End Sub

Sub CopyToClipboardShell(ByVal txt As String)
'
' CopyToClipboard Macro
'
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    
    ' Kreiraj komandnu liniju za kopiranje teksta u clipboard
    objShell.Run "cmd /c echo " & txt & " | clip", 0, True
End Sub

