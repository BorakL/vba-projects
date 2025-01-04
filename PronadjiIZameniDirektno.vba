Sub PronadjiIZameniDirektno()
    Dim doc As Document
    Dim rng As Range
    Dim kriterijumi As Object
    Dim kljuc As Variant
    Dim pronadjenoBS As Boolean
    Dim pronadjenoVANRFZO As Boolean
    Dim pronadjenoCD As Boolean
    
    ' Inicijalizacija
    Set doc = ActiveDocument
    pronadjenoBS = False
    pronadjenoVANRFZO = False
    pronadjenoCD = False
    
    ' Definišemo kriterijume: tekst koji tražimo i odgovarajuće zamene
    Set kriterijumi = CreateObject("Scripting.Dictionary")
    kriterijumi.Add "PULMOLOGIJA", "KLINIKA ZA PULMOLOGIJU"
    kriterijumi.Add "ORL I MFH", "KLINIKA ZA ORL I MFH"
    kriterijumi.Add "NEUROLOGIJA", "KLINIKA ZA NEUROLOGIJU"
    kriterijumi.Add "INFEKTIVNE I TROPSKE BOLESTI", "KLINIKA ZA INFEKTIVNE I TROPSKE BOLESTI"
    kriterijumi.Add "GAK", "KLINIKA ZA GINEKOLOGIJU I AKU" & ChrW(352) & "ERSTVO"
    kriterijumi.Add "PLASTIKA", "KLINIKA ZA OPEKOTINE, PLASTI" & ChrW(268) & "NU I REKONSTRUKTIVNU HIRURGIJU"
    kriterijumi.Add "UROLOGIJA UKC", "KLINIKA ZA UROLOGIJU - Resavska 51"
    kriterijumi.Add "PUNKT1", "KLINIKA ZA NEUROHIRURGIJU - Punkt 2"
    kriterijumi.Add "PUNKT2", "KLINIKA ZA NEUROHIRURGIJU - Punkt 1"
    kriterijumi.Add "UROLOGIJA 2", "KLINIKA ZA UROLOGIJU - Pasterova 2"
    kriterijumi.Add "NEFROLOGIJA", "KLINIKA ZA NEFROLOGIJU"
    kriterijumi.Add "ENDOKRINOLOGIJA", "KLINIKA ZA ENDOKRINOLOGIJU, DIJABETES I BOLESTI METABOLIZMA"
    kriterijumi.Add "KARDIOLOGIJA", "KLINIKA ZA KARDIOLOGIJU KO 3"
    kriterijumi.Add "O" & ChrW(268) & "NO", "KLINIKA ZA O" & ChrW(268) & "NE BOLESTI"
    kriterijumi.Add "KOŽNO", "KLINIKA ZA DERMATOVENEROLOGIJU"
    ' Provera za "BS"
    Set rng = doc.Content
    With rng.Find
        .Text = "BS"
        .MatchCase = False ' Ignoriše velika/mala slova
        .Wrap = wdFindStop
        If .Execute Then
            pronadjenoBS = True
        End If
    End With

    ' Provera za "VAN RFZO"
    Set rng = doc.Content
    With rng.Find
        .Text = "VAN RFZO"
        .MatchCase = False ' Ignoriše velika/mala slova
        .Wrap = wdFindStop
        If .Execute Then
            pronadjenoVANRFZO = True
        End If
    End With
    
    ' Provera za "Č-D"
    Set rng = doc.Content
    With rng.Find
        .Text = ChrW(268) & "-D" ' Koristi Unicode za "Č" (268)
        .MatchCase = False ' Ignoriše velika/mala slova
        .Wrap = wdFindStop
        If .Execute Then
            pronadjenoCD = True
        End If
    End With
    
    ' Iteracija kroz kriterijume za zamenu
    For Each kljuc In kriterijumi.Keys
        Set rng = doc.Content
        With rng.Find
            .Text = kljuc
            .Replacement.Text = kriterijumi(kljuc)
            .MatchCase = False ' Ignoriše velika/mala slova
            .Wrap = wdFindContinue ' Nastavlja pretragu kroz ceo dokument
            .Execute Replace:=wdReplaceAll ' Zamenjuje sve instance pronađenog teksta
        End With
    Next kljuc
    
    ' Obaveštavanje korisnika o rezultatima
    If pronadjenoBS Then
        MsgBox "Ima bistra supa", vbInformation, "Obaveštenje"
    End If
    
    If pronadjenoVANRFZO Then
        MsgBox "Ima van RFZO", vbInformation, "Obaveštenje"
    End If
    
    If pronadjenoCD Then
        MsgBox "Ima caj", vbInformation, "Obaveštenje"
    End If
    
End Sub