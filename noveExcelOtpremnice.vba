Function PronadjiTabelu() As Range
    Dim ws As Worksheet
    Dim prviRed As Long
    Dim poslednjiRed As Long
    Dim tbl As Range
    
    ' Postavljanje reference na radni list
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Promeni ime lista ako je potrebno
    
    ' Početni red tabele je fiksiran na 11
    prviRed = 11
    
    ' Pronalazak poslednjeg reda tabele (gde u koloni A piše "UKUPNO:")
    poslednjiRed = prviRed
    
    ' Iteracija kroz redove dok ne pronađemo "UKUPNO:"
    Do While ws.Cells(poslednjiRed, 1).Value <> "UKUPNO:"
        poslednjiRed = poslednjiRed + 1
    Loop
    
    ' Postavljanje reference na tabelu od A do C
    Set tbl = ws.Range(ws.Cells(prviRed, 1), ws.Cells(poslednjiRed, 3))
    
    ' Vraćanje reference na tabelu
    Set PronadjiTabelu = tbl
End Function




Function UpdateSum(tbl As Range)
    Dim suma As Double ' Koristi Double za bolje rukovanje decimalnim vrednostima
    Dim i As Integer
    Dim broj As String
    Dim lastColumn As Integer
    
    suma = 0 ' Resetovanje sume na početku
    
    ' Pronalazak poslednje kolone u tabeli koja sadrži podatke
    lastColumn = tbl.Columns.Count
    
    ' Iteracija kroz redove i sabiranje vrednosti iz poslednje kolone vidljivih redova
    For i = 1 To tbl.Rows.Count - 1
        If Not tbl.Rows(i).Hidden Then ' Proveravamo da li je red vidljiv
            broj = Trim(tbl.Cells(i, lastColumn).Value) ' Uzimamo vrednost iz poslednje kolone
            broj = Replace(broj, Chr(13) & Chr(7), "") ' Uklanjamo znakove za kraj reda
            
            If IsNumeric(broj) Then
                suma = suma + CDbl(broj) ' Koristimo CDbl za rad sa decimalnim brojevima
            End If
        End If
    Next i
    
    ' Ažuriranje vrednosti u poslednjem redu (SUMA)
    tbl.Cells(tbl.Rows.Count, lastColumn).Value = suma ' Postavljanje sume u poslednju ćeliju
End Function





Sub IzdvojVRFZO()
    Dim ws As Worksheet
    Dim tbl As Range
    Dim brojRedova As Integer
    Dim tekstPrveKolone As String
    Dim i As Integer
    Dim foundVRFZO As Boolean
    Dim stvarniRed As Long
    
    ' Postavljanje reference na aktivni radni list
    Set ws = ActiveSheet
    
    ' Referenca na tabelu
    Set tbl = PronadjiTabelu()
    If tbl Is Nothing Then
        MsgBox "Tabela nije pronađena!", vbExclamation
        Exit Sub
    End If
    
    brojRedova = tbl.Rows.Count
    foundVRFZO = False
    
    ' Otkrij sve redove u tabeli
    tbl.EntireRow.Hidden = False
    
    ' Provera da li postoji "VAN RFZO" u prvoj koloni
    For i = 1 To brojRedova
        tekstPrveKolone = Trim(tbl.Cells(i, 1).Value)
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
    For i = 1 To brojRedova
        tekstPrveKolone = Trim(tbl.Cells(i, 1).Value)
        tekstPrveKolone = Replace(tekstPrveKolone, Chr(13) & Chr(7), "") ' Uklanja znakove za kraj reda
        
        ' Stvarni red u radnom listu (potrebno ako tabela ne počinje od reda 1)
        stvarniRed = tbl.Cells(i, 1).Row
        
        ' Sakrij redove koji ne sadrže "VAN RFZO"
        If InStr(1, tekstPrveKolone, "VAN RFZO", vbTextCompare) = 0 Then
            ws.Rows(stvarniRed).Hidden = True
        End If
    Next i
    
    ' Ažuriranje vrednosti Suma
    Call UpdateSum(tbl)
End Sub






Sub IzdvojBsDbCM()
    Dim ws As Worksheet
    Dim tbl As Range
    Dim brojRedova As Integer
    Dim tekstPrveKolone As String
    Dim i As Integer
    Dim keywords As Variant
    Dim foundKeyword As Boolean
    Dim stvarniRed As Long
    
    ' Ključne reči koje tražimo u tabeli
    keywords = Array("BS", "M-D", ChrW(268) & "-D", "DNEVNA")
    
    ' Postavljanje reference na aktivni radni list
    Set ws = ActiveSheet
    
    ' Referenca na tabelu
    Set tbl = PronadjiTabelu()
    If tbl Is Nothing Then
        MsgBox "Tabela nije pronađena!", vbExclamation
        Exit Sub
    End If
    
    brojRedova = tbl.Rows.Count
    
    ' Otkrivanje svih redova u tabeli (jednim pozivom)
    tbl.EntireRow.Hidden = False
    
    ' Provera da li postoji red sa ključnim rečima
    Dim hasValidRow As Boolean
    hasValidRow = False
    
    For i = 1 To brojRedova
        tekstPrveKolone = Trim(tbl.Cells(i, 1).Value)
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
    For i = 1 To brojRedova
        tekstPrveKolone = Trim(tbl.Cells(i, 1).Value)
        tekstPrveKolone = Replace(tekstPrveKolone, Chr(13) & Chr(7), "") ' Uklanja znakove za kraj reda
        
        foundKeyword = False
        For Each keyword In keywords
            If InStr(1, tekstPrveKolone, keyword, vbTextCompare) > 0 Then
                foundKeyword = True
                Exit For
            End If
        Next keyword
        
        ' Stvarni red u radnom listu
        stvarniRed = tbl.Cells(i, 1).Row
        
        ' Sakrij red ako ne sadrži ključne reči
        If Not foundKeyword Then
            ws.Rows(stvarniRed).Hidden = True
        End If
    Next i
    
    ' Ažuriranje vrednosti Suma
    Call UpdateSum(tbl)
End Sub





 
Sub IzbaciBsDbVrfzo()
    Dim ws As Worksheet
    Dim tbl As Range
    Dim brojRedova As Integer
    Dim tekstPrveKolone As String
    Dim i As Integer
    Dim keywords As Variant
    Dim stvarniRed As Long
    
    ' Ključne reči koje treba sakriti
    keywords = Array("VAN RFZO", "BS", "M-D", ChrW(268) & "-D", "DNEVNA")
    
    ' Postavljanje reference na aktivni radni list
    Set ws = ActiveSheet
    
    ' Referenca na tabelu
    Set tbl = PronadjiTabelu()
    If tbl Is Nothing Then
        MsgBox "Tabela nije pronađena!", vbExclamation
        Exit Sub
    End If
    
    brojRedova = tbl.Rows.Count
    
    ' Otkrivanje svih redova odjednom
    tbl.EntireRow.Hidden = False
    
    ' Iteracija kroz redove i sakrivanje onih sa ključnim rečima
    For i = 1 To brojRedova
        tekstPrveKolone = Trim(tbl.Cells(i, 1).Value)
        tekstPrveKolone = Replace(tekstPrveKolone, Chr(13) & Chr(7), "") ' Uklanja znakove za kraj reda
        
        ' Provera da li red sadrži neku od ključnih reči
        For Each keyword In keywords
            If InStr(1, tekstPrveKolone, keyword, vbTextCompare) > 0 Then
                ' Sakrij red ako sadrži ključnu reč
                stvarniRed = tbl.Cells(i, 1).Row
                ws.Rows(stvarniRed).Hidden = True
                Exit For ' Nema potrebe da tražimo dalje u tom redu
            End If
        Next keyword
    Next i
    
    ' Ažuriranje sume koristeći funkciju UpdateSum
    Call UpdateSum(tbl)
End Sub






Sub proveriOtpremnicu()
    Dim doc As Worksheet
    Dim tbl2 As Range
    Dim row As Range
    Dim cell As Range
    Dim kriterijumi As Object
    Dim kljuc As Variant
    Dim regex As Object
    Dim foundItems As Collection
    Dim message As String
    Dim i As Integer
    
    ' Referenca na aktivni radni list
    Set doc = ActiveSheet
    
    ' Referenca na tabelu
    Set tbl2 = PronadjiTabelu()
    
    ' Provera da li tabela postoji
    If tbl2 Is Nothing Then
        MsgBox "Tabela nije pronađena!", vbExclamation
        Exit Sub
    End If
    
    ' Otkrivanje svih redova u tabeli
    tbl2.EntireRow.Font.Hidden = False
    
    ' Kreiranje RegEx objekta
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = False ' Traži samo prvi pogodak u tekstu
    
    ' Kreiranje kriterijuma za pretragu
    Set kriterijumi = CreateObject("Scripting.Dictionary")
    kriterijumi.Add "BS", "BISTRA SUPA"
    kriterijumi.Add "VAN RFZO", "VAN RFZO"
    kriterijumi.Add "DNEVNA", "DNEVNA BOLNICA"
    kriterijumi.Add "DB", "DNEVNA BOLNICA"
    kriterijumi.Add "HEMODIJALIZA SENDVI" & ChrW(268) & "I", "DNEVNA BOLNICA"
    kriterijumi.Add ChrW(268) & "-D", "CAJ"
    kriterijumi.Add "M-D", "MLEKO"
    
    ' Kolekcija za pronalaženje pogodaka
    Set foundItems = New Collection
    
    ' Iteracija kroz redove u tabeli
    For Each row In tbl2.Rows
        For Each cell In row.Cells
            Dim cellText As String
            cellText = Trim(cell.Value)
            cellText = Replace(cellText, Chr(13) & Chr(7), "") ' Uklanja specijalne znakove
            
            ' Provera svakog kriterijuma
            For Each kljuc In kriterijumi.Keys
                regex.Pattern = "\b" & kljuc & "\b" ' Traženje cele reči
                If regex.Test(cellText) Then
                    On Error Resume Next ' Sprečava duplikate u kolekciji
                    foundItems.Add kriterijumi(kljuc), CStr(kriterijumi(kljuc))
                    On Error GoTo 0
                    Exit For ' Ako nađe podudaranje, izlazi iz unutrašnje petlje
                End If
            Next kljuc
        Next cell
    Next row
    
    ' Formiranje poruke ako su pronađeni rezultati
    If foundItems.Count > 0 Then
        message = "Otpremnica sadrži:" & vbCrLf
        For Each kljuc In foundItems
            message = message & "- " & kljuc & vbCrLf
        Next kljuc
        MsgBox message, vbInformation, "Rezultat provere"
    Else
        MsgBox "Nema pronađenih stavki u otpremnici.", vbInformation, "Rezultat provere"
    End If
    
    ' Ažuriranje sume koristeći funkciju UpdateSum
    Call UpdateSum(tbl2)
End Sub

 
 


Sub StampajOtpremnicu()
'
' StampajOtpremnicu Macro
'
    ' Definišemo promenljive
    Dim ws As Worksheet
    
    ' Postavljanje reference na aktivni radni list
    Set ws = ActiveSheet
    
    ' Štampanje prvog primerka
    ws.PrintOut Copies:=1
    
    ' Štampanje drugog primerka
    ws.PrintOut Copies:=1
End Sub
 