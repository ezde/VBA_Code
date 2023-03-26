Attribute VB_Name = "Evidence"
Sub vyhledej_materialy_na_sklade()
On Error GoTo konec
Set b = Workbooks("Evidence laboratorních materiálů").Sheets("Skladová evidence")
rb = b.UsedRange.Rows.Count

With uf_evidencelab
    .lb_zasoba.Clear
    .lb_zasoba.ColumnCount = 10

For i = 2 To rb
    If b.Cells(i, 1) = CDbl(.PLU) Then
        .lb_zasoba.AddItem b.Cells(i, 2)
        .lb_zasoba.List(.lb_zasoba.ListCount - 1, 1) = b.Cells(i, 4)
        .lb_zasoba.List(.lb_zasoba.ListCount - 1, 2) = b.Cells(i, 9)
        .lb_zasoba.List(.lb_zasoba.ListCount - 1, 3) = b.Cells(i, 3)
        .lb_zasoba.List(.lb_zasoba.ListCount - 1, 4) = CStr(b.Cells(i, 7))
        .lb_zasoba.List(.lb_zasoba.ListCount - 1, 5) = CStr(b.Cells(i, 11))
        .lb_zasoba.List(.lb_zasoba.ListCount - 1, 6) = b.Cells(i, 8)
        .lb_zasoba.List(.lb_zasoba.ListCount - 1, 7) = b.Cells(i, 13)
        .lb_zasoba.List(.lb_zasoba.ListCount - 1, 8) = b.Cells(i, 5)
        .lb_zasoba.List(.lb_zasoba.ListCount - 1, 9) = b.Cells(i, 6)
    End If
Next i

End With
konec:
End Sub
 
Sub objednavky_aktualizuj()
On Error GoTo konec
Set b = Workbooks("Evidence laboratorních materiálů").Sheets("Objednávky")
rb = b.UsedRange.Rows.Count

With uf_Objednavky
    .lb_seznamobj.Clear
    .lb_seznamobj.ColumnCount = 6

For i = 2 To rb
    If b.Cells(i, 3) = "�ek� na vy��zen�" Then
        .lb_seznamobj.AddItem b.Cells(i, 1)
        .lb_seznamobj.List(.lb_seznamobj.ListCount - 1, 1) = b.Cells(i, 2)
        .lb_seznamobj.List(.lb_seznamobj.ListCount - 1, 2) = b.Cells(i, 3)
        .lb_seznamobj.List(.lb_seznamobj.ListCount - 1, 3) = b.Cells(i, 4)
        .lb_seznamobj.List(.lb_seznamobj.ListCount - 1, 4) = b.Cells(i, 5)
        .lb_seznamobj.List(.lb_seznamobj.ListCount - 1, 5) = b.Cells(i, 7)
    End If
Next i

End With
konec:
End Sub

Sub Dodavky_aktualizuj()
On Error GoTo konec
Set b = Workbooks("Evidence laboratorn�ch materi�l�").Sheets("Objedn�vky")
rb = b.UsedRange.Rows.Count

With uf_Objednavky
    .lbox_Objednavky.Clear
    .lbox_Objednavky.ColumnCount = 3

For i = 2 To rb
    If b.Cells(i, 3) = "Objedn�no" Then
        .lbox_Objednavky.AddItem b.Cells(i, 1)
        .lbox_Objednavky.List(.lbox_Objednavky.ListCount - 1, 1) = b.Cells(i, 2)
        .lbox_Objednavky.List(.lbox_Objednavky.ListCount - 1, 2) = b.Cells(i, 7)
    End If
Next i

End With
konec:
End Sub

Sub TiskID()

'Pro přípda, že budu chtít tisknout ID štítky

On Error GoTo konec
sablona = "C:\Programov�n�\VBA\Manager chemik�li�\�ablony\St�tek_ID.xltm"
Set a = Workbooks("Evidence laboratorn�ch materi�l�").Sheets("Skladov� evidence")
ra = a.UsedRange.Rows.Count

With uf_ObjednavkyPrijem
    nazev = .nazev
    ID = CDbl(a.Cells(ra, 2))
    PLU = .PLU
    expirace = .tb_expirace

End With

Workbooks.Add template:=sablona
Set b = Workbooks(ActiveWorkbook.Name).Sheets("ID")
    b.Range("nazev") = nazev
    b.Range("ID") = ID
    b.Range("PLU") = PLU
    b.Range("Expirace") = expirace
    b.PrintOut
    
ActiveWorkbook.Close savechanges:=False
Exit Sub
konec:
End Sub


Sub TiskExpirace()

'Pro případ, že budu chtít tiskonout štítky s Expirací

On Error GoTo konec
sablona = "C:\Programování\VBA\Manager chemikálií\šablony\Stítek_Expirace.xltm"
Set a = Workbooks("Evidence laboratorních materiálů").Sheets("Skladová evidence")
ra = a.UsedRange.Rows.Count

With uf_HlaseniStavu
    nazev = .lb_nazev
    ID = CDbl(.lb_ID)
    For i = 2 To ra
    If ID = a.Cells(i, 2) Then
    expirace = CDate(a.Cells(i, 11))
    End If
    Next i
    Datum = Date
End With

Workbooks.Add template:=sablona
Set b = Workbooks(ActiveWorkbook.Name).Sheets("Expirace")
    b.Range("Název") = nazev
    b.Range("ID") = ID
    b.Range("Datum") = Datum
    b.Range("Expirace") = expirace
    b.PrintOut
    
ActiveWorkbook.Close savechanges:=False
Exit Sub
konec:
End Sub

Sub ZavedMaterial1()

With uf_ObjednavkyPrijem

If .ch_NA = True And .ch_PhEur = True And .ch_USP = True Then
 MsgBox "Pozor, nemůžeš mít označeny všechny volby (Ph. Eur., USP, N/A)! ", vbExclamation, "DATUM EXPIRACE"
 Exit Sub
End If

If .ch_NA = True And .ch_PhEur = False And .ch_USP = True Then
 MsgBox "Pozor, nemůžeš mít označeny dvě volby (N/A, USP)! ", vbExclamation, "DATUM EXPIRACE"
 Exit Sub
End If

If .ch_NA = False And .ch_PhEur = True And .ch_USP = True Then
 MsgBox "Pozor, nemůžeš mít označeny dvě volby (Ph. Eur., USP)! ", vbExclamation, "DATUM EXPIRACE"
 Exit Sub
End If

If .ch_NA = True And .ch_PhEur = True And .ch_USP = False Then
 MsgBox "Pozor, nemůžeš mít označeny dvě volby (Ph. Eur., N/A)! ", vbExclamation, "DATUM EXPIRACE"
 Exit Sub
End If

If .ch_PhEur = True Or .ch_NA = True Or .ch_USP = True Then
GoTo dal
If IsDate(.tb_expirace) = False Then
  MsgBox "Zadej datum ve formátu dd.mm.rrrr", vbExclamation, "DATUM EXPIRACE"
Exit Sub
End If
End If

dal:
If .tb_sarzeOMD = "" Then
 i = MsgBox("Nezapomněl jsi zadat číslo šarže?", vbYesNo, "POZOR")
 Select Case i
  Case vbNo
    GoTo dal1
  Case vbYes
   MsgBox "Tak ho teď doplň", vbInformation, "ŠARžE ONCOMED"
   Exit Sub
 End Select
End If
  
dal1:

If .tb_cistota = "" Then
 i = MsgBox("Nezapomněl jsi zadat údaj o čistotě?", vbYesNo, "POZOR")
 Select Case i
  Case vbNo
    GoTo dal2
  Case vbYes
   MsgBox "Tak ho teď doplň", vbInformation, "ČISTOTA"
   Exit Sub
 End Select
End If

If IsNumeric(.tb_cistota) = False Then
  MsgBox "Zadávej pouze čísla", vbExclamation, "ČISTOTA"
  Exit Sub
End If

dal2:
If IsNumeric(.tb_mnozstvi) = False Then
  MsgBox "Zadávej pouze číselné hodnoty", vbExclamation, "MNOŽSTVÍ"
  Exit Sub
End If

If .cmb_jednotka = "" Then
    MsgBox "Zadej jednotku", vbExclamation, "JEDNOTKA"
    Exit Sub
End If


'P�id� materi�l do evidence
pocet = .lb_davka.ListCount - 1
For j = 0 To pocet

Set a = Workbooks("Evidence laboratorních materiálů").Sheets("Skladová evidence")
ra = a.UsedRange.Rows.Count

    a.Cells(ra + 1, 1) = CDbl(.PLU)
    a.Cells(ra + 1, 2) = CDbl(a.Cells(ra, 2) + 1)
    a.Cells(ra + 1, 3) = "Nová"
 If IsNumeric(.tb_sarzeOMD) = True Then
    a.Cells(ra + 1, 4) = CDbl(.tb_sarzeOMD)
 Else
    a.Cells(ra + 1, 4) = .tb_sarzeOMD
 End If
    a.Cells(ra + 1, 5) = CDbl(.tb_mnozstvi)
    a.Cells(ra + 1, 6) = .cmb_jednotka
 If .ch_PhEur = True Then
    a.Cells(ra + 1, 7) = "Ph. Eur."
 End If
 If .ch_USP = True Then
    a.Cells(ra + 1, 7) = "USP"
 End If
 If .ch_NA = True Then
    a.Cells(ra + 1, 7) = "N/A"
 End If
 If .ch_PhEur = False And .ch_NA = False And .ch_USP = False Then
  If IsDate(.tb_expirace) = True Then
    a.Cells(ra + 1, 7) = CDate(.tb_expirace)
  Else
    a.Cells(ra + 1, 7) = .tb_expirace
  End If
 End If
 
 If IsNumeric(.tb_cistota) = True Then
    a.Cells(ra + 1, 8) = CDec(.tb_cistota)
 Else
    a.Cells(ra + 1, 8) = .tb_cistota
 End If
    a.Cells(ra + 1, 9) = .tb_sarzevyr
    a.Cells(ra + 1, 15) = .nazev
    a.Cells(ra + 1, 16) = .tb_expirace

Set b = Workbooks("Evidence laboratorních materiálů").Sheets("Objednávky")
    rb = b.UsedRange.Rows.Count
    For i = 3 To rb
    
    If CDbl(.PLU) = CDbl(b.Cells(i, 1)) And CDbl(.lb_davka.List(j, 1)) = CDbl(b.Cells(i, 7)) Then
    
    b.Cells(i, 3) = "Vyřazeno"
    b.Cells(i, 8) = Date
    End If
    Next i

Evidence.TiskID

Next j

.lb_davka.Clear
End With


End Sub

Sub ZavedMaterial2()

With uf_ObjednavkyPrijem

If .tb_expirace = "" Then
 i = MsgBox("Nezapomněl jsi zadat Expiraci?", vbYesNo, "POZOR")
 Select Case i
  Case vbNo
    GoTo dal
  Case vbYes
   MsgBox "Tak ji teď doplň", vbInformation, "EXPIRACE"
   Exit Sub
 End Select
End If


If IsDate(.tb_expirace) = False Then
  MsgBox "Zadej datum ve formátu dd.mm.rrrr", vbExclamation, "DATUM EXPIRACE"
Exit Sub
End If

dal:
If .tb_sarzeOMD = "" Then
 i = MsgBox("Nezapomněl jsi zadat číslo šarže?", vbYesNo, "POZOR")
 Select Case i
  Case vbNo
    GoTo dal1
  Case vbYes
   MsgBox "Tak ho teď doplň", vbInformation, "ŠARŽE ONCOMED"
   Exit Sub
 End Select
End If


dal1:
If IsNumeric(.tb_mnozstvi) = False Then
  MsgBox "Zadávej pouze číselné hodnoty", vbExclamation, "MNOŽSTVÍ"
  Exit Sub
End If

If .cmb_jednotka = "" Then
    MsgBox "Zadej jednotku", vbExclamation, "JEDNOTKA"
    Exit Sub
End If


'P�id� materi�l do evidence
pocet = .lb_davka.ListCount - 1
For j = 0 To pocet

Set a = Workbooks("Evidence laboratorních materiálů").Sheets("Skladová evidence_Spotřební")
ra = a.UsedRange.Rows.Count

    a.Cells(ra + 1, 1) = CDbl(.PLU)
    a.Cells(ra + 1, 2) = a.Cells(ra, 2) + 1
    a.Cells(ra + 1, 3) = "Nová"
 If IsNumeric(.tb_sarzeOMD) = True Then
    a.Cells(ra + 1, 4) = CDbl(.tb_sarzeOMD)
 Else
    a.Cells(ra + 1, 4) = .tb_sarzeOMD
 End If
    a.Cells(ra + 1, 5) = CDbl(.tb_mnozstvi)
    a.Cells(ra + 1, 6) = .cmb_jednotka
 If IsNumeric(.tb_expirace) = True Then
    a.Cells(ra + 1, 7) = CDate(.tb_expirace)
 Else
    a.Cells(ra + 1, 7) = .tb_expirace
 End If
 If IsNumeric(.tb_cistota) = True Then
    a.Cells(ra + 1, 8) = CDec(.tb_cistota)
 Else
    a.Cells(ra + 1, 8) = .tb_cistota
 End If
    a.Cells(ra + 1, 9) = .tb_sarzevyr
    a.Cells(ra + 1, 15) = .nazev


Set b = Workbooks("Evidence laboratorních materiálů").Sheets("Objednávky")
    rb = b.UsedRange.Rows.Count
    For i = 3 To rb
    
    If CDbl(.PLU) = CDbl(b.Cells(i, 1)) And CDbl(.lb_davka.List(j, 1)) = CDbl(b.Cells(i, 7)) Then
    
    b.Cells(i, 3) = "Vyřazeno"
    b.Cells(i, 8) = Date
    End If
    Next i

Next j

.lb_davka.Clear
End With

End Sub

Sub TiskStitku()
'Nefunk�n�
Set a = Workbooks("AMApplikace.xlsm").Sheets("Komponenty")
ra = a.UsedRange.Rows.Count
Set b = Workbooks("Evidence laboratorních materiálů").Sheets("Skladová evidence")
rb = b.UsedRange.Rows.Count
    
With uf_TiskStitku
.lb_vyber.Clear
.lb_vyber.ColumnCount = 3
    
    For i = 2 To rb
     For j = 2 To ra
  If b.Cells(i, 1) = a.Cells(j, 1) Then
        
        .lb_vyber.AddItem b.Cells(i, 1)
        .lb_vyber.List(.lb_vyber.ListCount - 1, 1) = b.Cells(i, 2)
        .lb_vyber.List(.lb_vyber.ListCount - 1, 2) = a.Cells(j, 2)
        
        End If
        Next j
    Next i
    End With



End Sub

Sub KdoJeKdo()
kdo = Application.UserName
Set a = Workbooks("LabFinder.xlsm").Sheets("Jm�na")
ra = a.UsedRange.Rows.Count

For i = 1 To ra
If a.Cells(i, 1) & " " & a.Cells(i, 2) = kdo Then
'a.Range("Role") = a.Cells(i, 3)
zde = a.Cells(i, 3)
End If
Next i

End Sub


