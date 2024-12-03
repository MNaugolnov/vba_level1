### DAN DRUGI 
Ponovimo šta smo radili dan prije toga u jednom kodu koji će biti priprema za naredni dan, na koji ćemo dodavati uslove

Dodati komentar da mi je sada jasno da se pitate i zasto nesto da radim na celiji, danas sve dobija smisao, jer cemo dosadasnja znanja primenjivati sa svhom tamo gde imamo sansu da pogresimo i ne mozemo u kratkom roku da sve vidimo

Mogu im reći za ovo 
```Application.ScreenUpdating = False```

## USLOVI 
* Prikazati na slide primer uslova i pojasniti
* Navesti primere kako imamo raylicite potrebe i tako moze da nas zanima da nam tabela kaze

-------------------
Provera vrednosti
```
Sub d2_uslov1()
'Tri opsega cena (0,20],(20,40],(40,+)
Dim cena_1h As Integer, zakljucak As String

cena_1h = Cells(2, 2).Value

    If cena_1h > 0 And cena_1h <= 20 Then
            
            'cena_1h.font.ColorIndex=6
            Cells(2, 2).Font.ColorIndex = 6
            zakljucak = "Rang 1"
            
    ElseIf cena_1h > 20 And cena_1h <= 40 Then
    
            Cells(2, 2).Font.ColorIndex = 10
            zakljucak = "Rang 2"
    
    Else
            zakljucak = "Rang 3"
        
         
    End If
    
Cells(33, 1).Value = zakljucak

End Sub
```

-----------
'Provera tipa podataka
```
Sub d2_ulov2()
'Konvertovanje vrednosti aktivne æelije
Const Rate As Double = 118
Dim cena_1h As Integer, cena_1h_din As Integer

If IsNumeric(ActiveCell) Then 'ActiveCell.Value
        cena_1h = ActiveCell.Value
        cena_1h_din = cena_1h * Rate
        MsgBox cena_1h_din
Else
    MsgBox "Selektovana vrednosti nije cena!"
    
End If
End Sub

```
#Prikazati sta sve može da stoji u uslovu i koja provera tipova podataka na prezentacji


--------------------------------------------------Petlje------------------------------------

#Ali mi ne zelimo da primenimo uslov samo na jednoj ćeliji
-----------FOR--------------
```
Sub d2_petlje_for1()

'Obojiti vrednosti u prema opsegu cena-cela kolona
Dim cena As Integer

For i = 2 To 29
    cena = Cells(i, 2).Value

     If cena > 0 And cena <= 20 Then
        
            Cells(i, 2).Font.ColorIndex = 6
            
     ElseIf cena > 20 And cena <= 40 Then
    
            Cells(i, 2).Font.ColorIndex = 10
     Else
           Cells(i, 2).Font.ColorIndex = 16
          
    End If

Next
End Sub
```

```
Sub d2_petlje_for2()

'Obijo vrednosti u prema opsegu cena - cela kolona
Dim cena As Integer

For Each cell In Range("B2:I29") 'Range("B2:B29") ili Range(Cells(1,1),Cells(29,8)
        cena = cell.Value
    
         If cena > 0 And cena <= 20 Then
            
                cell.Font.ColorIndex = 6
                
         ElseIf cena > 20 And cena <= 40 Then
        
                cell.Font.ColorIndex = 10
         Else
                cell.Font.ColorIndex = 16
              
        End If
    
Next

End Sub
```
-------------While----------
```
Sub d2_petlje_while1()
'Prvo pojavljivanje cene treæeg ranga
Dim r As Integer
    
    r = 2
    While Cells(r, 2) <= 40
        'Cells(r, 2).Interior.ColorIndex = 3
        r = r + 1
    Wend
    MsgBox ("Datum cene " + CStr(Cells(r, 1).Value))

End Sub
```
```
Sub d2_petlje_while2()
Dim r As Integer
    r = 2
    Do
      r = r + 1
    Loop While Cells(r, 2) <= 40
    MsgBox ("Datum cene " + CStr(Cells(r, 1).Value))
End Sub
```
-------------------------------------------------------------------------------------




--------------------Procedure i funkcije -------------------------
Preurediti na primeru kada vraca koji je prosek nekog sata za sve dane
```
Function konverzija(cena_e As Double) As Double
'Funkcije-izazna vrednost koja može postati deo neke druge procedure
    Const Rate As Double = 118
    
    konverzija = cena_e * Rate
    
End Function
```
```
Sub macro_test()
    Dim cena_din As Double
    cena_din = konverzija(ActiveCell.Value)
    MsgBox cena_din
End Sub
```



----------------------------WB i WS------------------------------------------------------
* Dosadadanji rad, rad na jednom listu, kao refetentni uvijek uzima aktivni sheet (predji na D1 i ponovo odradi proceduru neku)
* Ako želimo da radimo na nekom drugom listu, moramo posebno da se pozovemo
* Kako smo se opcijama Range,Cells pozivali na celije ili skup celija u okviru jednog lista, na isti način pristupamo i radnim listovima
Pojasniti takako do sada samo imali rad samo na jednom listu, i kod uvijek uzima aktivan list kao onaj na kome izvrsava sve operacije, ukoliko ne naglasimo drugacije

```
Sub d2_radni_list1()
'Aktivnost radnog lista i pristupanje podacima sa radnog lista
'   MsgBox ActiveSheet.Name
'   MsgBox ActiveSheet.Cells(2, 2).Value
'   MsgBox Sheets("D1").Cells(32, 1).Value
'   MsgBox Sheets("D1").Name
'   MsgBox Sheets(1).Name
'   MsgBox ActiveSheet.Name


'.ACTIVATE
'Aktivan list ostaje poæetni, neovisno od instrukcija, dok drugacije ne neglasimo
    Sheets("D1").Activate
    MsgBox ActiveSheet.Name
    
'Kada se doda novi list, tada on postaje aktivan, ako se ne naglasi suprotno
    Sheets.Add.Name = "Izvestaj1"
    MsgBox ActiveSheet.Name


    'MsgBox ActiveWorkbook.Name
'    MsgBox ActiveWorkbook.Sheets("D1").Cells(32, 1).Value
'
'    Sheets("Podaci").Visible = 0
'    Sheets("D1").Tab.ColorIndex = 6
   
End Sub
```
```
Sub d2_radnilist2()
'Inicijalizacija
Dim cell As Range
Dim wsh As Worksheet
Dim m_wb, c_wb As Workbook
Dim sh_num As Integer
'Otvoriti postojeæu radnu svesku i vrednosti prvog sata kopirati za sve berze


'Setovnje
Set m_wb = ActiveWorkbook
Set c_wb = Workbooks.Open("C:\Users\jovana.arsic\DS\VBA Course\Februar_Cene\Cene.xlsx") 'xlsm napomena


'Provera
MsgBox m_wb.Name
MsgBox c_wb.Name

'Dodati novi list u koji æemo upisivati željene podatke
m_wb.Sheets.Add.Name = "Cene02"


'Prolazak kroz listove i èuvanje postojeæeg lista
'Koliko ima listova
sh_num = c_wb.Sheets.Count

    For i = 1 To sh_num
    'Prolaz kroz redove
        For r = 1 To 29
                'Ostavljamo slobodan red za nazive berzi
                m_wb.Sheets("Cene02").Cells(r + 1, i + 1).Value = c_wb.Sheets(i).Cells(r, 2).Value
                'Datume procitamo sa jednog od listova
                m_wb.Sheets("Cene02").Cells(r + 1, 1).Value = c_wb.Sheets(1).Cells(r, 1).Value
        Next
        
        'Naziv kolone u izvestaju jedank nazivu berze koji se nalazi u nazivu lista
        m_wb.Sheets("Cene02").Cells(1, i + 1).Value = c_wb.Sheets(i).Name
              
    Next


'NAPOMENA: KADA RADIMO SA NOVO SVESKOM ili LISTOM to postaje nova aktivna wb ili sheet
    'MsgBox ActiveWorkbook.Name

'Zatvorimo radnu svesku bez promena
m_wb.Save
c_wb.Close

End Sub
```

## TABELA 
```
Sub d2_tabela1()
Dim c_sh As Worksheet
Dim c_tabela As ListObject


'Precizirajmo list
Set c_sh = ActiveWorkbook.Sheets("Cene02")

Set c_tabela = c_sh.ListObjects.Add(xlSrcRange, Range("A1:E29"), xlYes)

c_tabela.Name = "Cena_Februar"
c_tabela.TableStyle = "TableStyleMedium19"


End Sub
```


## GRAFIKON

```
Sub d2_grafikon1()
Dim graf As ChartObject
Dim c_sh As Worksheet


'Precizirajmo list
Set c_sh = ActiveWorkbook.Sheets("Cene02")

'Preciziramo podatke koje želimo predstaviti
Set Rng = c_sh.Range("A1:E29")
Set graf = c_sh.ChartObjects.Add(Left:=180, Width:=400, Top:=7, Height:=210)

With graf.Chart

    .SetSourceData Rng '
    .ChartType = xlLine
    .SetElement (msoElementChartTitleAboveChart)
    .ChartTitle.Text = "Februar"
    
    
End With

End Sub
```
