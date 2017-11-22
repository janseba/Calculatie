Attribute VB_Name = "menu_handelingen"
Option Explicit

Sub werkposten_zoeken(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    If cel_in_tabel_aanwezig(ActiveCell) Then
        FRM_werkposten.Show
    End If
End Sub

Sub artikelen_zoeken(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    FRM_artikelen_selecteren.Show
End Sub

Sub calculatie_tonen(welke)
    Dim cel
    Dim verwijzing

    verwijzing = ""
    With Sheets("voorblad")
        For Each cel In Range("b2:b6")
            If cel.Value = welke Then
                verwijzing = "begin_calculatie_" & cel.Offset(, 2).Value
                Exit For
            End If
        Next cel
        If verwijzing = "" Then
            For Each cel In Range("b2:b6")

            Next cel
        End If
        With ActiveWindow.Panes(2)
            .ScrollRow = Range(verwijzing).row
        End With
    End With
End Sub

Sub regels_met_kosten_tonen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim lo As ListObject
    Dim rij As ListRow
    Dim tonen As Boolean

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    beveiliging (False)

    tonen = Not Range("regels_zonder_kosten_tonen").Value

    Range("regels_zonder_kosten_tonen").Value = tonen

    For Each lo In ActiveSheet.ListObjects
        For Each rij In lo.ListRows

            If rij.Range.row > 20 Then
            
            If tonen = True Then
            Rows(rij.Range.row).RowHeight = 15
            Else
                If (Cells(rij.Range.row, 7).Value + Cells(rij.Range.row, 18).Value) = 0 Then
                    Rows(rij.Range.row).RowHeight = 0.1
                    Else
            Rows(rij.Range.row).RowHeight = 15
                End If

            End If
            End If

        Next rij
    Next lo

    beveiliging (True)

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationManual
    End With

End Sub

Sub groepen_met_kosten_tonen()
    Dim lo As ListObject
    Dim rij As ListRow
    Dim tonen As Boolean
    Dim hoogte
    Dim bereik
Dim naam
Dim verw

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    beveiliging (False)
    
    tonen = Not Range("groepen_zonder_kosten_verbergen").Value
    Range("groepen_zonder_kosten_verbergen").Value = tonen
    
    For Each lo In ActiveSheet.ListObjects
    
    hoogte = IIf(Cells(lo.DataBodyRange.row - 1, 19).Value = 0 And Not (tonen), 0.1, 15)
    If lo.name <> "template_tabel" Then
    Range(Cells(lo.DataBodyRange.row - 1, 1), Cells(lo.DataBodyRange.row + lo.DataBodyRange.Rows.Count + 1, 1)).EntireRow.RowHeight = hoogte
    End If
    Next lo

    For Each bereik In ActiveSheet.Names
        naam = bereik.name
        verw = bereik
        
        If InStr(1, naam, "_vast", vbTextCompare) > 0 Or InStr(1, naam, "_var", vbTextCompare) > 0 Then
        hoogte = IIf(Cells(Range(verw).row, 19) = 0 And Not (tonen), 0.1, 22)
        Range(verw).RowHeight = hoogte
        End If
        
        If naam = "einde_calculatie" Then Range(verw).RowHeight = 22

Next bereik

    beveiliging (True)

    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationManual
    End With

End Sub

Sub tarieven_tonen_verbergen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Application.ScreenUpdating = False
    beveiliging (False)
    Range("tarieven_tonen").Value = Not Range("tarieven_tonen").Value
    beveiliging (True)
    toon_tarieven
    Application.ScreenUpdating = True
End Sub

Sub uren_tonen_verbergen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Application.ScreenUpdating = False
    beveiliging (False)
    Range("totaal_uren_tonen").Value = Not Range("totaal_uren_tonen").Value
    beveiliging (True)
    toon_uren
    Application.ScreenUpdating = True
End Sub

Sub open_begroting_tonen_verbergen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    beveiliging (False)
    Range("open_begroting_tonen").Value = Not Range("open_begroting_tonen").Value
    beveiliging (True)
    open_begroting_tonen
End Sub
Sub open_begroting_tonen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim wel_niet As String
    wel_niet = "wel"
    If Range("open_begroting_tonen").Value Then wel_niet = "niet"
    MsgBox "Open begroting " & wel_niet & " tonen!", vbInformation
End Sub
Sub toon_tarieven(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim cel


    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    With ActiveSheet
        If InStr(1, .name, "calculatie_", vbTextCompare) > 0 Then
            beveiliging (False)
            For Each cel In .Range("uren")
                If InStr(1, LCase(cel.Value), "tarief", vbTextCompare) > 0 Then
                    cel.EntireColumn.Hidden = Not .Range("tarieven_tonen").Value
                End If
            Next cel
        End If
        beveiliging (True)
    End With

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With

End Sub

Sub toon_uren(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim cel

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    With ActiveSheet
        If InStr(1, .name, "calculatie_", vbTextCompare) > 0 Then

            beveiliging (False)
            For Each cel In .Range("uren")
                If InStr(1, LCase(cel.Value), "totaal  uren", vbTextCompare) > 0 Then
                    cel.EntireColumn.Hidden = Not .Range("totaal_uren_tonen").Value
                End If
            Next cel
        End If

        beveiliging (True)

    End With

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With

End Sub

Sub rijen_invoegen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim lrow As Long
    Dim olo As ListObject
    Dim ocell
    Dim r
    Dim rij

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    On Error GoTo fout

    Set olo = ActiveCell.ListObject
    Set ocell = ActiveCell

    beveiliging (False)
    If cel_in_tabel_aanwezig(ocell) Then
        r = InputBox("Hoeveel rijen invoegen?", "Aantal rijen?", 1)

        If IsNumeric(r) Then
            beveiliging (False)
            For rij = 1 To r
                lrow = ocell.row - olo.DataBodyRange.Cells(1, 1).row + 1
                Application.StatusBar = "Invoegen rij " & rij
                'ActiveCell.ListObject.ListRows.Add (lrow)
                
                Rows(olo.DataBodyRange.Cells(1, 1).row + 1).Insert
                
                
                
            Next rij
            beveiliging (True)
        End If
    End If 'If cel_in_tabel_aanwezig(cel)

verder:

    rijen_goed_zetten

    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .StatusBar = False
    End With

    Set olo = Nothing
    Set ocell = Nothing

    End

fout:
    MsgBox "Er is een fout opgetreden bij het invoegen van rijen", vbInformation

    GoTo verder

End Sub

Sub rijen_verwijderen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim cel
    Dim tabel
    Dim c
    Dim r

    Application.ScreenUpdating = False

    Set cel = ActiveCell

    'kijken in welke tabel de rijen verwijderd dienen te worden
    For Each tabel In ActiveSheet.ListObjects
        If Intersect(cel, tabel.Range) Is Nothing Then
        Else
            If Intersect(cel, tabel.DataBodyRange) Is Nothing Then
                Exit For
            Else
                If tabel.DataBodyRange.Rows.Count > 1 Then
                    beveiliging (False)
                    For r = Application.Selection.row + Selection.Rows.Count - 1 To Application.Selection.row Step -1
                        Debug.Print Cells(r, Selection.Column).Address
                        Rows(r).Delete Shift:=xlUp
                    Next r
                    beveiliging (True)
                End If
                Exit For
            End If
        End If
    Next tabel

    rijen_goed_zetten

    Set cel = Nothing

    Application.ScreenUpdating = True

End Sub


Function cel_in_tabel_aanwezig(cel)
    Dim tabel As ListObject
    For Each tabel In ActiveSheet.ListObjects
        If Intersect(cel, tabel.Range) Is Nothing Then
            'If Intersect(cel, tabel.Range) Is Nothing And (cel.Row < tabel.Range.Row - 1) Then
        Else
            cel_in_tabel_aanwezig = True
            Exit For
        End If ' If Not lo Is Nothing
    Next tabel
End Function
Sub subgroep_verwijderen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim r As Range
    Dim lo As ListObject
    Dim begin_rij
    Dim eind_rij

    Set r = ActiveCell
    Set lo = r.ListObject
    If Not lo Is Nothing Then

        Dim melding
        Dim groepsnaam

        '                If InStr(1, lo, "_vast", vbTextCompare) > 0 Then
        '                    MsgBox "Uit de vaste delen van de calculatie kunnen geen groepen worden verwijderd!", vbCritical + vbOK
        '                    End
        '                End If

        melding = "Zeker weten dat je de groep ""naam_groep"" wilt verwijderen"
        groepsnaam = "van regel " & lo.Range.row & " t/m " & lo.Range.row + lo.Range.Rows.Count - 1
        If Range(naam_kolom & lo.Range.row - 1) <> "" Then groepsnaam = Range(naam_kolom & lo.Range.row - 1)
        melding = Replace(melding, "naam_groep", groepsnaam)
        Select Case MsgBox(melding, vbYesNo)
            Case vbYes
            Application.ScreenUpdating = False

            begin_rij = lo.Range.row - 1
            eind_rij = begin_rij + lo.Range.Rows.Count + 1

            beveiliging (False)
            Rows(begin_rij & ":" & eind_rij).Delete Shift:=xlUp
            beveiliging (True)

            ActiveCell.Select

            rijen_goed_zetten
            Application.ScreenUpdating = True
        End Select
    End If
End Sub

Function subgroep_invoegen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim tbl As ListObject
    Dim bron_rij, doel_rij
    Dim bron_tabel
    Dim doel_tabel
    Dim aantal_adres_bron
    Dim aantal_adres_doel
    Dim laatste_tabel_rij
    Dim kolom
    Dim rij
    Dim tabel_naam
    Dim tabel
    Dim tabellen
    Dim adres
    Dim tabel_niet_aanwezig As Boolean
    Dim tbl_naam
    Dim melding
    Dim teller
    Dim rij_teller
    Dim eerste_tabel_rij
    Dim naam
    Dim test
    Dim bovengrens_rij
    Dim ondergrens_rij
    Dim groep_rij
    Dim max_rij_Tabel
    Dim verschil_naam
    Dim verschil
    Dim bron_tabell
    Dim eerste_tabel_rij_aanmaak
    Dim cel
    Dim lo
    Dim k
    Dim r
    Dim ri
    Dim lr
    Dim subgroep_rijen
    Dim lr_groepen
    Dim aanwezige_rijen
    Dim x
    Dim rji
    Dim begin_rij_groep
    Dim aantal_regels_subgroep
    Dim sh_sub_groep
    Dim titel

Dim naam_zonder_underscores

    test = False
    melding = ""

    beveiliging (False)

    instellingen_ophalen

    If ActiveCell.row > Range(Range(groepnaam & max_groep_rij).Value).row Then
        melding = "Er kan geen groep toegevoegd worden na het einde van de calculatie"
        GoTo einde
    End If

    If ActiveCell.row < Range(Range(groepnaam & 1).Value).row Then
        melding = "Er kan geen groep toegevoegd worden voor het begin van de calculatie"
        GoTo einde
    End If

    ActiveSheet.Range("aan_te_maken_subgroep").Value = ""

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .Cursor = xlWait

    End With

    'FRM_Melding.Show
    'ActiveSheet.Shapes("subgroepmelding_tonen").Visible = True

    With ActiveSheet
'    Stop
        Set bron_tabell = Sheets("calculatie_").ListObjects("template_tabel")
        'controle op aanwezigheid in tabel
        If cel_in_tabel_aanwezig(ActiveCell) = True Then
            melding = "In een subgroep kan geen andere groep toegevoegd worden, selecteer een andere cel"
            GoTo einde
        End If

        For rij = 1 To max_groep_rij
            If ActiveCell.row = Range(Range(groepnaam & rij).Value).row Then
                melding = "Op deze rij kan geen subgroep worden toegevoegd"
                GoTo einde
            End If
        Next rij  ' For rij = 1 To max_groep_rij

        If test Then Stop

        'kijken in welke deel van de calculatie de actieve cel zich bevindt aan de hand daarvan de voorloop van de naam bepalen
        For rij = 1 To max_groep_rij - 1
            adres = ("A" & Range(Range(groepnaam & rij).Value).row) & ":" & ("AA" & Range(Range(groepnaam & rij + 1).Value).row)
            'adres = Range(Range(groepnaam & rij).Value).Row & ":" & Range(Range(groepnaam & rij + 1).Value).Row
            If Intersect(ActiveCell, Range(adres)) Is Nothing Then
            
            Else
                tbl_naam = Range(groepnaam & rij).Value
                Exit For
            End If
        Next rij  ' For rij = 1 To 7

'controleren of er een speciale subgroep geselecteerd moet worden, bij de ef projecten hoeft dit niet

    If InStr(1, tbl_naam, "ef_") = 0 Then
    Select Case MsgBox("Wilt u een voor gedefinieerde subgroep toevoegen?", vbYesNo, "Subgroep toevoegen?")
        Case vbYes
        frmSubgroep.Show
        Case vbNo
        '´`ActiveSheet.Range("aan_te_maken_subgroep").Value = ""
    End Select
    End If

        'controleren of de tabel vlak boven een andere groep zit
        max_rij_Tabel = 9
        groep_rij = Range(Range(groepnaam & rij + 1).Value).row - 1
        For rij_teller = 0 To max_rij_Tabel
            If ActiveCell.row + rij_teller >= groep_rij Then
            beveiliging (False)
                Rows(ActiveCell.row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Rows(8).Copy Rows(ActiveCell.row)
            End If
        Next rij_teller

        If test Then Stop

        'kijken welk volgnummer de tabelnaam dient te krijgen door alle tabellen door te lopen
        'en te kijken of de tabel met het bewuste volgnummer aanwezig is.
        For Each tabel In ActiveSheet.ListObjects
            tabellen = tabellen & "|" & tabel.name & "|"
        Next tabel
        tabel_niet_aanwezig = False
        While tabel_niet_aanwezig = False
        teller = teller + 1
        tabel_naam = tbl_naam & "_" & teller & "_" & GetCodenameFromWorksheet(ActiveSheet.name) & IIf(ActiveSheet.Range("aan_te_maken_subgroep") <> "", "__" & ActiveSheet.Range("aan_te_maken_subgroep"), "")
        If InStr(1, tabellen, "|" & tabel_naam, vbTextCompare) = 0 Then
            tabel_niet_aanwezig = True
        End If
    Wend

    doel_tabel = tabel_naam

    If test Then Stop

    ' controleren of de subgroep wordt aangemaakt in een andere subgroep dat is namelijk niet toegestaan
    For Each lo In ActiveSheet.ListObjects
        If InStr(1, lo.name, tbl_naam, vbTextCompare) > 0 Then
            eerste_tabel_rij = (lo.Range.row - 1)
            If ActiveCell.row = eerste_tabel_rij Then
                melding = "In een subgroep kan geen andere subgroep toegevoegd worden, selecteer een andere cel"
                GoTo einde
            End If
        End If
    Next lo

    'controleren of de actieve cell gelijk onder een andere tabel staat
    ' indien dit het geval is, regel invoegen en deze ingevoegde regel selecteren
    If test Then Stop
    beveiliging (False)
    For Each lo In ActiveSheet.ListObjects
        If InStr(1, lo.name, tbl_naam, vbTextCompare) > 0 Then
            If test Then Stop
            laatste_tabel_rij = (lo.Range.row + lo.Range.Rows.Count)
            
'            Stop
            Dim rij_
            rij_ = ActiveCell.row
            
            If (ActiveCell.row - laatste_tabel_rij) < 3 And (ActiveCell.row - laatste_tabel_rij) >= 0 Then
            
            For x = 1 To 3
                Rows(ActiveCell.row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                               Rows(8).Copy Rows(ActiveCell.row)
                                             Cells(ActiveCell.row + 1, ActiveCell.Column).Select
                                Next x
                
              '  Cells(rij_, ActiveCell.Column).Select
                                Exit For
            
            End If
        End If
    Next lo

 If test Then Stop

    For Each lo In ActiveSheet.ListObjects
        If InStr(1, lo.name, tbl_naam, vbTextCompare) > 0 Then
            If test Then Stop
            eerste_tabel_rij = (lo.Range.row - 2)
            If ActiveCell.row = eerste_tabel_rij Then
                Rows(ActiveCell.row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                Rows(8).Copy Rows(ActiveCell.row + 1)
                Cells(ActiveCell.row + 1, ActiveCell.Column).Select
                Exit For
            End If
        End If
    Next lo

    verschil = 9999999

    For Each lo In ActiveSheet.ListObjects
        If InStr(1, lo.name, tbl_naam, vbTextCompare) > 0 Then
            'If lo.Range.Row - 4 - ActiveCell.Row < verschil And (lo.Range.Row - 2 - ActiveCell.Row) >= 0 Then
            If lo.Range.row - 3 - ActiveCell.row < verschil And (lo.Range.row - 3 - ActiveCell.row) >= 0 Then
                verschil_naam = lo.name
                verschil = lo.Range.row - 4 - ActiveCell.row
                'verschil = lo.Range.Row - 2 - ActiveCell.Row
            End If
        End If
    Next lo

    If verschil_naam <> "" Then
        Set lo = ActiveSheet.ListObjects(verschil_naam)
        If test Then Stop
        eerste_tabel_rij = (lo.Range.row - 2)
        If eerste_tabel_rij >= ActiveCell.row Then
            cel = ActiveCell.Address
            eerste_tabel_rij_aanmaak = ActiveCell.row
            beveiliging (False)
            While eerste_tabel_rij_aanmaak + max_rij_Tabel >= eerste_tabel_rij
            Rows(ActiveCell.row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(ActiveCell.row + 1, ActiveCell.Column).Select
            eerste_tabel_rij = (lo.Range.row - 2)
        Wend
        Range(cel).Select
    End If
End If

If test Then Stop

'bepalen hoever de actieve cel van de ondergrens af zit
bovengrens_rij = Range("wtb_var").row + 3
ondergrens_rij = Range("einde_calculatie").row - 2

' controleren of de geselecteerde cel ver genoeg zit van de ondergrens
If ActiveCell.row + 7 > ondergrens_rij Then
    For rij_teller = ActiveCell.row To ActiveCell.row + (ActiveCell.row + 6 - ondergrens_rij)
        Stop
        Rows(ActiveCell.row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next rij_teller
End If

' controleren of de geselecteerde cel ver genoeg zit van de bovengrens
If ActiveCell.row <= bovengrens_rij Then
    For rij_teller = ActiveCell.row To ActiveCell.row + (ActiveCell.row + 1 - bovengrens_rij)
        Rows(ActiveCell.row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Cells(ActiveCell.row + 1, ActiveCell.Column).Select
    Next rij_teller
End If

bron_rij = bron_tabell.DataBodyRange.Cells(1, 1).row - 1

bron_tabell.Range.Copy Destination:=Range(naam_kolom & ActiveCell.row + 1)
.Range(naam_kolom & ActiveCell.row + 1).ListObject.name = doel_tabel

Set tbl = .ListObjects(doel_tabel)
tbl.ShowTotals = True
doel_rij = tbl.DataBodyRange.Cells(1, 1).row - 1

'kop_rij kopieren
.Rows(bron_rij).EntireRow.Copy .Rows(doel_rij)

If ActiveSheet.Range("aan_te_maken_subgroep").Value <> "" Then

    lr_groepen = Sheets("subgroepen").Cells(Sheets("subgroepen").Rows.Count, Sheets("subgroepen").Range("a1").Column).End(xlUp).row

    For r = 1 To lr_groepen
        If Sheets("subgroepen").Cells(r, 1).Value = ActiveSheet.Range("aan_te_maken_subgroep") Then
            subgroep_rijen = subgroep_rijen + 1
        End If
    Next r

    With ActiveSheet.ListObjects(doel_tabel)

        rji = ActiveSheet.ListObjects(doel_tabel).Range.row + 1
        aanwezige_rijen = .Range.Rows.Count - 1

        'rijen toevoegen om de voorgedefinieerde regels toe te kunnen voegen
        For x = 1 To subgroep_rijen - aanwezige_rijen + 1
            Rows(rji).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next x

        rji = ActiveSheet.ListObjects(doel_tabel).Range.row

        'beginregel voor de subgroep bepalen
        With Sheets("subgroepen")
            For r = 1 To lr_groepen
                If .Cells(r, 1).Value = ActiveSheet.Range("aan_te_maken_subgroep").Value Then
                    begin_rij_groep = r
                    Exit For
                End If
            Next r
            aantal_regels_subgroep = Application.WorksheetFunction.CountIf(.Range("A:A"), ActiveSheet.Range("aan_te_maken_subgroep").Value)
        End With

        Set sh_sub_groep = Sheets("subgroepen")

        With ActiveSheet
            rji = .ListObjects(doel_tabel).Range.row
            teller = 0
            For r = rji To rji + aantal_regels_subgroep - 1
                .Cells(r, Range(naam_kolom & "1").Column).Formula = sh_sub_groep.Cells(begin_rij_groep + teller, 2).Formula
                With .Cells(r, Range(naam_kolom & "1").Column + 1)
                    .NumberFormat = "@"
                    .Value = sh_sub_groep.Cells(begin_rij_groep + teller, 3).Value
                End With
                'FORMULES TBV PRIJZEN OPZOEKEN TOEVOEGEN
                teller = teller + 1
            Next r

            .Cells(rji, Range(naam_kolom & "1").Column + 2).FormulaR1C1 = "=IFERROR(VLOOKUP([@Kolom2],'Prijslijst to be'!R2C1:R613C13,4,FALSE),0)"
            .Cells(rji, Range(naam_kolom & "1").Column + 5).FormulaR1C1 = "=IFERROR(VLOOKUP([@Kolom2],'Prijslijst to be'!R2C1:R613C13,11,FALSE),0)"

        End With

        'TITEL VAN DE SUBGROEP TOEVOEGEN
        titel = ActiveSheet.Range("aan_te_maken_subgroep").Value

        'Type subgroep in subgroepkop zetten
        Range(naam_kolom & rji - 1).Value = Mid(UCase(titel), 1, 1) & Mid(LCase(titel), 2)

    End With
End If

naam_zonder_underscores = Replace(tabel_naam, " ", "_")

'aanpassen van de totalisering van de nieuw toe te voegen tabel
Range("s" & tbl.DataBodyRange.row - 1).FormulaR1C1 = Replace(Range("s" & tbl.DataBodyRange.row - 1).FormulaR1C1, "template_tabel", naam_zonder_underscores)

For k = 1 To Range("kolom_kop").Columns.Count
    r = Range("kolom_kop").row
    'Aantal kolom aanpassen naar het aantal van de desbetreffende subgroep
    If InStr(1, LCase(.Cells(r, k).Value), "totaal") Then
        ri = tbl.DataBodyRange.row
        .Cells(ri, k).Formula = Replace(.Cells(ri, k).Formula, Range("aantal_indicatie").Address, "$E$" & doel_rij)
    End If ' If InStr(1, LCase(.Cells(r, k).Value), "totaal")

    'aanpassen van de totalisering naar de huidige tabel
    If InStr(1, LCase(.Cells(r, k).Value), "verkoopbedrag") Then .Cells(doel_rij, k).Formula = Replace(Cells(doel_rij, k).Formula, bron_tabell.name, doel_tabel)

Next k 'For k = 1 To Range("kolom_kop").Columns.Count

If ActiveSheet.Range("aan_te_maken_subgroep").Value <> "" Then

End If

'indien nodig aantal rijen in tabel verwijderen totdat het er max 5 zijn
If ActiveSheet.Range("aan_te_maken_subgroep").Value = "" Then
    With tbl
        While .DataBodyRange.Rows.Count > 3
        ActiveSheet.Rows(.Range.row + 1).EntireRow.Delete
    Wend
End With 'With tbl
End If

'FRM_Melding.Hide
ActiveSheet.Shapes("subgroepmelding_tonen").Visible = False

beveiliging (True)
rijen_goed_zetten

macro_einde:

With Application
    .ScreenUpdating = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
    .Cursor = Default
End With

End

einde:
If melding <> "" Then
    MsgBox melding, vbInformation
End If

GoTo macro_einde
End
End With

fout:
MsgBox "We hebben een foutmelding!", vbCritical

GoTo einde
End Function

Sub beveiliging(wat)

    ActiveSheet.Unprotect
    If wat Then
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Else
        ActiveSheet.Unprotect
    End If
End Sub

Sub rijen_goed_zetten()

    Dim rij
    Dim naam_rij
    Dim naam
    Dim adressen
    Dim adres

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    instellingen_ophalen

    With ActiveSheet
        adressen = ""

        For naam_rij = 1 To max_groep_rij
            Application.StatusBar = "Rijen verzamelen " & naam_rij
            naam = .Range(groepnaam & naam_rij)
            If naam_rij > 1 Then adressen = "|" & adressen & "|" & Cells(Range(Range(groepnaam & naam_rij)).row - 1, 1).Address & "|"
            If naam_rij < max_groep_rij Then
                adressen = adressen & Cells(Range(Range(groepnaam & naam_rij)).row + 1, 1).Address
            End If 'If naam_rij < max_groep_rij
        Next naam_rij 'For naam_rij = 1 To max_groep_rij

        beveiliging (False)
        For rij = Range("wtb_var").row To Range("einde_calculatie").row
            adres = Cells(rij, 1).Address
            Application.StatusBar = "Rijen tonen/verbergen: " & naam_rij
            If InStr(1, adressen, "|" & adres & "|") > 0 Or .Cells(rij, 1) = "Kolom1" Then
                If .Rows(rij).EntireRow.Hidden = False Then .Rows(rij).EntireRow.Hidden = True
            Else
                If .Rows(rij).EntireRow.Hidden = True Then .Rows(rij).EntireRow.Hidden = False
            End If
        Next rij
        beveiliging (True)
    End With

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .StatusBar = False
    End With

End Sub

Sub formules_goedzetten(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim rij
    Dim kolom
    Dim LastCol
    Dim letter
    Dim groep_rij_begin
    Dim groep_rij_eind
    Dim eind_tekst
    Dim begin_tekst
    Dim nieuwe_formule
    Dim formule
    Dim t

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    beveiliging (False)

    instellingen_ophalen

    With ActiveSheet
        For rij = 1 To max_groep_rij - 1
            groep_rij_begin = Range(Range(groepnaam & rij).Value).row
            groep_rij_eind = Range(Range(groepnaam & rij + 1).Value).row - 1
            LastCol = .Cells(groep_rij_begin, .Columns.Count).End(xlToLeft).Column
            For kolom = 1 To LastCol
                letter = LCase(Split(Cells(groep_rij_begin, kolom).Address, "$")(1))
                formule = LCase(.Cells(groep_rij_begin, kolom).Formula)
                If Left(formule, 1) = "=" Then
                    t = Split(formule, letter)
                    begin_tekst = Left(formule, InStr(1, formule, "(", vbTextCompare))
                    eind_tekst = Mid(formule, InStr(1, formule, ")", vbTextCompare))
                    nieuwe_formule = begin_tekst & letter & groep_rij_begin + 1 & ":" & letter & groep_rij_eind & eind_tekst
                    .Cells(groep_rij_begin, kolom).Formula = nieuwe_formule
                End If
            Next kolom
        Next rij
        beveiliging (True)
    End With

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With

End Sub

Sub groepen_samenvatten_tonen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim rij
    Dim adres
    Dim tbl_naam

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    instellingen_ophalen

    beveiliging (False)
    If Range("samenvatten") Then
        Range("samenvatten") = False
        For rij = 1 To max_groep_rij - 1
            adres = naam_kolom & Range(Range(groepnaam & rij).Value).row + 1 & ":A" & Range(Range(groepnaam & rij + 1).Value).row - 2
            Range(adres).RowHeight = 15
        Next rij  ' For rij = 1 To 7

        For rij = Range("begin_calculatie").row To Range("einde_calculatie").row
            If LCase(Cells(rij, 1).Value) = "kolom1" Then Cells(rij, 1).RowHeight = 0
        Next rij
    Else
        'kijken in welke deel van de calculatie de actieve cel zich bevindt aan de hand daarvan de voorloop van de naam bepalen
        For rij = 1 To max_groep_rij - 1
            adres = naam_kolom & Range(Range(groepnaam & rij).Value).row + 1 & ":" & naam_kolom & Range(Range(groepnaam & rij + 1).Value).row - 2
            Range(adres).RowHeight = 0.1
        Next rij  ' For rij = 1 To 7
        Range("samenvatten") = True
    End If
    beveiliging (True)

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With

End Sub

Sub subgroepen_samenvatten_tonen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim rij
    Dim adres
    Dim tbl_naam

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    beveiliging (False)

    If Range("samenvatten") Then

        Range("samenvatten") = False

        For rij = 1 To max_groep_rij - 1
            adres = naam_kolom & Range(Range(groepnaam & rij).Value).row + 1 & ":" & naam_kolom & Range(Range(groepnaam & rij + 1).Value).row - 2
            Range(adres).RowHeight = 15
        Next rij  ' For rij = 1 To 7

        For rij = Range("begin_calculatie").row To Range("einde_calculatie").row
            If LCase(Cells(rij, 1).Value) = "kolom1" Then Cells(rij, 1).RowHeight = 0
        Next rij

    Else

        'kijken in welke deel van de calculatie de actieve cel zich bevindt aan de hand daarvan de voorloop van de naam bepalen
        For rij = 1 To max_groep_rij - 1
            adres = naam_kolom & Range(Range(groepnaam & rij).Value).row + 1 & ":" & naam_kolom & Range(Range(groepnaam & rij + 1).Value).row - 2
            Range(adres).RowHeight = 0.1
        Next rij

        Range("samenvatten") = True

    End If 'If Range("samenvatten")

    beveiliging (True)

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With ' With Application

End Sub


Sub subgroep_kopieren(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)

    MsgBox "Gekopieerd!"

End Sub

Function vette_tekst_aan_uit(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    beveiliging (False)
    Selection.Font.Bold = Not (Selection.Font.Italic)
    beveiliging (True)
End Function

Function cursieve_tekst_aan_uit(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    beveiliging (False)
    Selection.Font.Italic = Not (Selection.Font.Italic)
    beveiliging (True)
End Function


'
'    Selection.Font.Bold = True
'    Selection.Font.Italic = True
'    Selection.Font.Underline = xlUnderlineStyleSingle
'    Selection.Font.Underline = xlUnderlineStyleNone
'    Selection.Font.Italic = False
'    Selection.Font.Bold = False


Sub calculatie_sheet_aanmaken()
Dim c
For Each c In ThisWorkbook.VBProject.VBComponents
If LCase(Left(c.name, 10)) = "calculatie" Then
End If

Debug.Print c.name
Next c
ThisWorkbook.VBProject.VBComponents(ActiveSheet.name).name = "calculatie_2"
End Sub

    Sub namen_aanpassen()
Dim naam As name
For Each naam In ActiveSheet.Names
If InStr(1, naam.name, "variabel", vbTextCompare) > 0 Then
Debug.Print naam.name
Debug.Print naam.RefersTo
ActiveSheet.Names.Add Replace(naam.name, "variabel", "var"), naam.RefersTo
End If
Next naam

Dim lo As ListObject

For Each lo In ActiveSheet.ListObjects

If InStr(1, lo.name, "variabel", vbTextCompare) > 0 Then
lo.name = Replace(lo.name, "variabel", "var")
End If

If InStr(1, lo.name, "calculatie", vbTextCompare) > 0 Then
lo.name = Replace(lo.name, "calculatie", "calc")
End If


Next lo

End Sub


