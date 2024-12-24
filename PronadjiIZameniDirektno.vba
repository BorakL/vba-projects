Sub PronadjiIZameniDirektno()
    Dim doc As Document
    Dim rng As Range
    Dim kriterijumi As Object
    Dim kljuc As Variant
    Dim pronadjenoBS As Boolean
    
    ' Inicijalizacija
    Set doc = ActiveDocument
    Set rng = doc.Content
    pronadjenoBS = False
    
    ' Definišemo kriterijume: tekst koji tražimo i odgovarajuće zamene
    Set kriterijumi = CreateObject("Scripting.Dictionary")
    kriterijumi.Add "GAK", "KLINIKA ZA GINEKOLOGIJU I AKUŠERSTVO"
    kriterijumi.Add "PLASTIKA", "KLINIKA ZA OPEKOTINE, PLASTIČNU I REKONSTRUKTIVNU HIRURGIJU"
    kriterijumi.Add "UROLOGIJA UKC", "KLINIKA ZA UROLOGIJU - Resavska 51"
    kriterijumi.Add "PUNKT1", "KLINIKA ZA NEUROHIRURGIJU - Punkt 2"
    kriterijumi.Add "PUNKT2", "KLINIKA ZA NEUROHIRURGIJU - Punkt 1"
    kriterijumi.Add "UROLOGIJA 2", "KLINIKA ZA UROLOGIJU - Pasterova 2"
    kriterijumi.Add "NEFROLOGIJA", "KLINIKA ZA NEFROLOGIJU"
    
    ' Proveravamo da li postoji "BS"
    Set rng = doc.Content
    With rng.Find
        .Text = "BS"
        .MatchCase = False ' Ignoriše velika/mala slova
        .Wrap = wdFindStop
        If .Execute Then
            pronadjenoBS = True
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
End Sub
