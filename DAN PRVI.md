# DAN PRVI

## UVODNA PRIČA

### PREDSTAVLJANJE

### CILJEVI

_Milivoje_

Cilj: Da za ova tri dana napiše iz glave velike kod i razume 99% komandi koje mogu biti otkucane. 

Šta zapravo Milivoje treba da savlada?
Da shvati gde bi ovo pogao da primeni?

1) Da li imate neke rutiske poslove, koje svakodnevno obavlajate i razhtevaju vrijeme a mogle bi se automatizovati?
2) Da li možda radite sa većim brojem workbook-va, velikim projem podataka koje kombinujete i niste nikada sigurni da li ste sve dobro kopirali?
3) Da možda sam generiše izveštaje 

### VBA

VBA je zapravo Excelov programski jezik kao i drugih Office programa.
Zapravo, služi nam da izvšimo skoro sve zamisli, kako one koje su već predviđene tako i nešto što nema implementaciju. Dakle da sami možemo zamišljeno programirati.
No, da bismo bili u stanju da izvedemo veliku većinu željenog neophodno je da savladamo osnove.
Takođe, na ovaj način možemo povezivati radnje excela , sa drugim office progrmaima.

* Podešavanje Developer Tub-a
1) Otvorimo File/Oprions/Ribbon/Add Developer
2) Dobili smo prezentaciju

Kada govorimo o kodiranju i VBA, prva stvar koju ćemo pored VB prozora videti jeste Macro.
Makro je zapravo dio koda koji je iskucan u VB porgramskom jeziku i koji pokretanjem je u stanju da izvši neke rutinske zadatke. 




### Osnovna struktura rada
krenucemo sa osnovnim pa ćemo sutra preći na sveske u rad sa većim brojem svesaka
kao i pravljenjem tabel



## UVOD i PRVI MACRO
Pojasniti podatke o berzama eketriče energije.

* Recording Macro (Prices)

* Run Macro (Pojasniti da imamo više načina runovanja macroa)

* Pojasniti delove prozor koje imamo, Module i slično 

* Preimenovati module na Dan1 

* Pojasniti naziv, komponente Sub/End Sub su početak i kraj Macroa (pripremi se za private, public za svaki slucaj)

* Obrisati ovaj sadržaj ručno pa pokrenuti macro na oba načina i preko skraćenice i kroz run, a pokazati da ima i run dugme u VBA prozoru

* Ispraviti Macro, promeniti pozicije, dodati naslovni red i promeniti boju slova i centrirati



### Selektovanje 
* Vid pristupa celiju, skupu celija, pozicioniranje u sonovi

* Mi dok radimo imamo razlicite vrste objekata i postoje funkcionalnosti koje se mogu primeniti na određene objekte samo

* Cilj danas je da shvatimo sta zapravo moze da se radi na ćeliji, na skupu ćelija i kako da im pristupimo, čitamo vrednosti iz njih i menjamo

* Sutra ćemo prelaziti na list i radnu svesku, no nužno je da se dobro upoznamo sa svim sto jedno polje nosi

```
Sub m_record()

' m_record Macro
' Prikaz cena na poèetku, sredini i kraju dana.
'
' Keyboard Shortcut: Ctrl+m

    Range("A1:A29,B1:B29,M1:M29,Y1:Y29").Select
    Range("Y1").Activate
    Selection.Copy
    Range("A32").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=-15
    Application.CutCopyMode = False
    
End Sub
```
```
Sub m_selektovanje1()
'Selektovanje polja A1
    Range("A1").Select
End Sub
```
```
Sub m_selektovanje2()
'Selektovanje opsega polja - Range
    Range("A1:B29").Select

End Sub
```
```
Sub m_selektovanje3()
'Selektovanje na drugom radnom listu
    
    'Aktiviranje drugog radnog lista
    Sheets("Podaci").Activate
    'Selektovanje
    Range("A2").Select

End Sub
```
```
Sub m_selektovanje4()
'Selektovanje više redova (sve kolone)
    Range("2:9").Select
    Rows("2:9").Select
End Sub
```
```
Sub m_selektovanje5()
'Selekcija jednog reda - Rows
    'Range("2").Select
     Rows(2).Select
End Sub
```
```
Sub m_selektovanje6()
'Selekcija kolona
    Range("A:C").Select
    Columns("A:C").Select
End Sub
```
```
Sub m_selektovanje7()
'Selekcija jedne kolone - Columns
    'Range("D").Select
    Columns(4).Select
End Sub
```
```
Sub m_selektovanje8()
'Selekcija jedne æelije - Cells
    Range("D6").Select
    Cells(8, 1).Select
End Sub
```
```
Sub m_selektovanje9()
'Selekcija opsega æelija
    Range(Cells(1, 1), Cells(3, 3)).Select
End Sub
```
--------------------
### Brisanje sadrzaja
```
Sub m_brisanje()
    'Erase the contents of column A
    Rows("32:62").ClearContents
End Sub
```

### Modifikovanje ćelija i vrednosti u ćelijama 
Nakon što smo naučili da se pozicioniramo i krećemo po radnom listu, to je bas lijepo, ali naš cilj je da menjamo sdaržaj i modifikujemo vrednosti u ćelijama
Nakon što na Range dodamo tačku otvara se padajući meni, koji sadrži sve moguće funkcionalnosti koje bismo mi mogli sprovoditi nad ćelijom.
```
Sub m_modifikovanje1()
'Pristupiti vrednosti æelije
MsgBox Range("A1").Value
MsgBox "Završeno"
End Sub
```
```
Sub m_modifikovanje2()
'Promena vrednosti i formata æelije
    'Nova vrednost
    Range("A1").Value = "DAN/SAT"
    Range("A1") = "D/H"
    Cells(1, 2).Value = "1H"
    
    
    'Ureðivanje Text-a u æeliji
    Cells(1, 1).Font.Name = "Arial"
    Cells(1, 1).Font.Size = "24"
    Cells(1, 1).Font.ColorIndex = 10
    
    'Ureðvanje izgleda æelije
    Cells(1, 1).Interior.ColorIndex = 2 '.Color=RGB(174, 240, 194)
    Cells(1, 1).Border.Weight = 4
    
End Sub
```
```
Sub m_modifikovanje3()
'Promena željenog na aktivnoj æeliji
    ActiveCell.Value = Cells(1, 1).Value
    
    With ActiveCell
        .Borders.Weight = 3
        .Font.Bold = True
    End With
       
End Sub
```
```
Sub m_modifikovanje4_zadatak()
'Kopiran macro
    Range("A1:A29,B1:B29,M1:M29,Y1:Y29").Select
    Selection.Copy
    Range("A32").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    
    'Cela Tabela
    Range(Cells(32, 1), Cells(60, 4)).Font.ColorIndex = 25
    Range(Cells(32, 1), Cells(60, 4)).Borders.Weight = 3
    Range(Cells(32, 1), Cells(60, 4)).HorizontalAlignment = xlCenter
    
    'Naslov
    With Range(Cells(32, 1), Cells(32, 4))
        With .Font
            .Size = 20
            .ColorIndex = 25
            .Italic = True
            .Bold = True
        End With
        .Interior.ColorIndex = 6
    End With
    
    'Kako da dodamo  H
    Cells(32, 2).Value = Cells(32, 2).Value + "H"
    'Cells(32, 3).Value = Cells(32, 3).Value + "H"
    'Cells(32, 4).Value = Cells(32, 4).Value + "H"
        
End Sub
```

### Promenljive 

* Primer kada nece da sabere
* Razgovarajmo o tipovima podataka koje imamo 

```
Sub m_prom1()
'Inicijalizacija promenljivih
Dim promInt As Integer
Dim promStr As String
Dim promDate As Date


'Dim promInt As Integer, promStr As String, promDate As Date


    pomDate = Cells(33, 1).Value
    pomInt = Cells(33, 2).Value
    promStr = Cells(32, 3).Value
    
    'MsgBox pomInt + 5
    MsgBox promStr + "H"
    MsgBox TypeName(promStr)
    
     Cells(32, 3).Value = promStr + "H"
    'Cells(32, 3).Value = CStr(Cells(32, 3).Value )+ "H"
End Sub
```
```
Sub m_prom2()

Dim br_red As Integer
    br_red = 32
    MsgBox (Cells(br_red, 1).Value + "H")
    
End Sub
```
```
Sub m_pom_const()
Const Rate As Double = 118

Dim br_red As Integer
    br_red = 33
    
    Cells(br_red, 5).Value = Cells(br_red, 2).Value * Rate
    Cells(br_red, 6).Value = Cells(br_red, 3).Value * Rate
    Cells(br_red, 7).Value = Cells(br_red, 4).Value * Rate
    
End Sub 
```
```
Sub m_dodatno1()
    'Offset opcija
    Cells(br_red, 2).Offset(0, 6).Value = Cells(br_red, 2).Value * Rate
    'ActiveCell
    MsgBox ActiveCell.Offset(0, -3).Value
End Sub
```
Za posednji deo idi po radnim sveskama, ako je dvanaesti sat uzmi samo 12 sat i kopiraj iz svih mogućih 12 sat a onda naprav tabelu od toga i grafikon i tabel
