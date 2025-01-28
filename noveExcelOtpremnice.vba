Function PronadjiTabelu() As Range
    Dim ws As Worksheet
    Dim prviRed As Long
    Dim poslednjiRed As Long
    Dim tbl As Range
    
    ' Postavljanje reference na radni list
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Promeni ime lista ako je potrebno
    
    ' Početni red tabele je fiksiran na 11
    prviRed = 11
    
    ' Pronalazak poslednjeg reda tabele (gde u koloni A piše "UKUPNO")
    poslednjiRed = prviRed ' Početno postavljamo poslednji red na početni red
    
    ' Prolazimo kroz redove dok ne pronađemo "UKUPNO:" ili praznu ćeliju
    Do While ws.Cells(poslednjiRed, 1).Value <> ""
        If ws.Cells(poslednjiRed, 1).Value = "UKUPNO:" Then
            Exit Do
        End If
        poslednjiRed = poslednjiRed + 1
    Loop
    
    ' Ako smo pronašli "UKUPNO", postavljamo referencu na tabelu
    If ws.Cells(poslednjiRed, 1).Value = "UKUPNO:" Then
        Set tbl = ws.Range(ws.Cells(prviRed, 1), ws.Cells(poslednjiRed - 1, 3)) ' Ispravljena referenca na poslednji red
    Else
        Set tbl = Nothing ' Ako tabela nije pronađena
    End If
    
    Set PronadjiTabelu = tbl
End Function



Function UpdateSum(tbl As Range)
    Dim suma As Long
    Dim i As Integer
    Dim broj As String
    
    suma = 0 ' Resetovanje sume na početku
    
    ' Iteracija kroz redove i sabiranje vrednosti iz poslednje kolone vidljivih redova
    For i = 1 To tbl.Rows.Count - 1
        If tbl.Rows(i).Range.Font.Hidden = False Then
            broj = Trim(tbl.Cells(i, tbl.Columns.Count).Value) ' Ispravljena referenca na Cells
            broj = Replace(broj, Chr(13) & Chr(7), "") ' Uklanja znakove za kraj reda
            
            If IsNumeric(broj) Then
                suma = suma + CLng(broj)
            End If
        End If
    Next i
    
    ' Ažuriranje vrednosti u poslednjem redu (SUMA)
    tbl.Cells(tbl.Rows.Count, tbl.Columns.Count).Value = suma ' Ispravljena referenca na Cells
End Function




Sub IzdvojVRFZO()
    Dim ws As Worksheet
    Dim tbl As Range
    Dim brojRedova As Integer
    Dim tekstPrveKolone As String
    Dim i As Integer
    Dim foundVRFZO As Boolean
    
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
    For i = 1 To brojRedova
        tbl.Rows(i).EntireRow.Hidden = False
    Next i
    
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
        
        ' Sakrij redove koji ne sadrže "VAN RFZO", osim zadnjeg reda
        If InStr(1, tekstPrveKolone, "VAN RFZO", vbTextCompare) = 0 And i <> brojRedova Then
            tbl.Rows(i).EntireRow.Hidden = True
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
    
    ' Prvo otkrivanje (unhide) svih redova u tabeli
    For i = 1 To brojRedova
        tbl.Rows(i).EntireRow.Hidden = False
    Next i
    
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
        
        ' Sakrij red ako ne sadrži ključne reči i nije poslednji red
        If Not foundKeyword And i <> brojRedova Then
            tbl.Rows(i).EntireRow.Hidden = True
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
    Dim suma As Long
    
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
    
    ' Prvo otkrivanje (unhide) svih redova u tabeli
    For i = 1 To brojRedova
        tbl.Rows(i).EntireRow.Hidden = False
    Next i
    
    ' Iteracija kroz redove i sakrivanje onih sa ključnim rečima
    For i = 1 To brojRedova
        tekstPrveKolone = Trim(tbl.Cells(i, 1).Value)
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
            tbl.Rows(i).EntireRow.Hidden = True
        End If
    Next i
    
    ' Ažuriranje sume koristeći funkciju UpdateSum
    Call UpdateSum(tbl)
End Sub





Sub proveriOtpremnicu()
    Dim doc As Worksheet
    Dim tbl2 As Range
    Dim row As Range ' Ispravljeno tipiranje
    Dim cell As Range ' Ispravljeno tipiranje
    Dim kriterijumi As Object
    Dim kljuc As Variant
    Dim regex As Object
    Dim foundItems As Collection
    Dim message As String
    
    ' Referenca na aktivni radni list
    Set doc = ActiveSheet
    
    ' Referenca na tabelu
    Set tbl2 = PronadjiTabelu()
    
    ' Prolazak kroz redove i otkrivanje (unhide) svih redova u tabeli
    For i = 1 To tbl2.Rows.Count
        tbl2.Rows(i).Range.Font.Hidden = False
    Next i
    
    ' Kreiranje RegEx objekta
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = False ' Traži prvi pogodak u tekstu
    
    ' Kreiranje kriterijuma za pretragu
    Set kriterijumi = CreateObject("Scripting.Dictionary")
    kriterijumi.Add "BS", "BISTRA SUPA"
    kriterijumi.Add "VAN RFZO", "VAN RFZO"
    kriterijumi.Add "DNEVNA", "DNEVNA BOLNICA"
    kriterijumi.Add "DB", "DNEVNA BOLNICA"
    kriterijumi.Add "HEMODIJALIZA SENDVI"&ChrW(268)&"I", "DNEVNA BOLNICA"
    kriterijumi.Add ChrW(268) & "-D", "CAJ"
    kriterijumi.Add "M-D", "MLEKO"
    
    ' Kolekcija za pronalaženje pogodaka
    Set foundItems = New Collection
    
    ' Provera da li dokument ima bar jednu tabelu
    If doc.Tables.Count = 0 Then
        MsgBox "Dokument ne sadrži tabele.", vbExclamation
        Exit Sub
    End If
    
    ' Iteracija kroz redove u tabeli
    For Each row In tbl2.Rows
        For Each cell In row.Cells
            Dim cellText As String
            cellText = Trim(cell.Value)
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
 