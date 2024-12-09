# DAN TREĆI 

## ZADATAK
Šta ako je vaš zadatak da na mjesecnom, sedmicnom ili cak dnevnom nivou posaljete gomilu ovakvih identicnih izvestaja? 
Koji sadrže tabelu sa podacima, sumiranim prosecim i recimo grafik? Vi svaki dam moreate, ma koliko se izvezbali
Zasto to ne biste radili na klik?

1) Napravimo malo dugme i dashboard za pokretanje macroa kreiranje izvestaja i macroa za slanje podataka
2) Pokupimo podatke iz većeg broja radnih sveski i kreiramo na novu da radnim listom koji ima podatke
3) Na novom listu kreiramo tabelu sa sumiranim rezultatima
4) Ispod kreiramo grafikon
5) I posaljemo izvestaj

### Priča o primerima
Kako bi se polako vratili na početak i na osnovo pitanje, kako sve ovo vama može da pomogne 
Voleo bih da spomenem par stvari, recimo ono sto bih vam mogao pokazati jeste nacin na koji sam ja sebi olaksao rad na prethodnom poslu.
Primer MAIN i MG (Navesti primere skidanja sa stranica, pravljenja izvestaja u Excelu, ali i pravljenja Word tipslih dokumenata, 

### Dodaj
```
If MsgBox("Text", vbYesNo, "Title") = vbYes Then

InputBox("Text ?", "Title", "Default value")
```

```
Sub m_napravi_izvestaj()
Set app = CreateObject("Excel.Application")
Dim w_report As Workbook, w_berza As Workbook
Dim p_red As Integer, p_kolona As Integer
Dim tabela As ListObject
Dim t_ime As String
Dim MyChart As ChartObject

'Napomenuti da mogu jednu po jednu ukoliko se nalaze na razlictim mestima podaci

'Dodajmo novi wb u koji æemo ispisati iyvestaj
Set w_main = ActiveWorkbook

Workbooks.Add.SaveAs Filename:="\\SNS06CFSH01\HomeFolderR\VBA BASIC\Izvestaj_Berze.xlsx"
Set w_report = ActiveWorkbook

'-----------------------Citanje
'Citam wb koji se nalaze u nekom folderu
'Putanja do foldera se uvijek zavrsava slesom
folder = "\\SNS06CFSH01\HomeFolderR\VBA BASIC\Cene\"


file = Dir(folder)


's brojac radnih svezaka
s = 1
While (Len(file) > 0)
    Set w_berza = Workbooks.Open(folder & file)

    'Aktiviraj radni list
    'w_berza.Sheets(1).Activate
    MsgBox (ActiveSheet.Name)

    'Pronaði poslednji red i kolonu u kojima imam nesto zapisano

    p_red = Cells(Rows.Count, "A").End(xlUp).Row
    p_kolona = Cells(1, Columns.Count).End(xlToLeft).Column

        For r = 2 To p_red

            'Izracunati prosecnu vrednost na cene struje za dan
            Cells(r, p_kolona + 1).Value = WorksheetFunction.Average(Range(Cells(r, 2), Cells(r, p_kolona)))

            'Upisati u izvešstaj
            w_report.Sheets(1).Cells(r, s + 1) = Cells(r, p_kolona + 1).Value
            'Datum u prvoj koloni
            w_report.Sheets(1).Cells(r, 1) = Cells(r, 1).Value
            'Naziv berze u prvom redu
            w_report.Sheets(1).Cells(1, s + 1) = Sheets(1).Name
        Next

    'Neæu da èuvam srednje vrednosti da tom listu
    w_berza.Close SaveChanges:=False

    'radna sveska nova
    s = s + 1

    'Preði na sledeæi file i prebaci brojaè
    file = Dir
Wend


'-------------------------------------Sacuvaj
w_report.Sheets(1).Name = "Prosek"



'
p_red_iz = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
p_kolona_iz = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column



'Tabela
t_ime = "ProsecnaCena"

Set tabela = ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(p_red_iz, p_kolona_iz)), , xlYes)
tabela.Name = t_ime
tabela.TableStyle = "TableStyleMedium19"


'Grafik

Set MyChart = w_report.Sheets("Prosek").ChartObjects.Add(Left:=230, Width:=500, Top:=10, Height:=200)

  With MyChart.Chart
   .SetSourceData ActiveSheet.Range(Cells(1, 1), Cells(p_red_iz, p_kolona_iz))
   .ChartType = xlLine
   'Dodavanje naslova hrafikonu
    .SetElement (msoElementChartTitleAboveChart)
    .ChartTitle.Text = "Februar"
  End With


'Sacuvati izvestaj

w_report.Close SaveChanges:=True

End Sub
```
## Slanje maila

Stari primer za liste prisutstva
