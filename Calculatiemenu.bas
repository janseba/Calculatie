Attribute VB_Name = "Calculatiemenu"
Option Explicit

Public Const Mname As String = "MyPopUpMenu"

Sub DeletePopUpMenu(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    'Delete PopUp menu if it exist
    On Error Resume Next
    Application.CommandBars(Mname).Delete
    On Error GoTo 0
End Sub

Sub CreateDisplayPopUpMenu(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Call DeletePopUpMenu
    Call Custom_PopUpMenu_1
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopup
    On Error GoTo 0
End Sub

Sub Custom_PopUpMenu_1(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim MenuItem As CommandBarPopup
    Dim laatste_groep_rij
    Dim rij
    Dim LastRow As Long
    Dim naam
    Dim lo
    Dim naam_rijen As String
    Dim b
    Dim k
    Dim x
    Dim LastCol As Integer
    Dim adres
    'Dim drop As msoControlComboBox
    Dim drop
    Dim li

    instellingen_ophalen

    With Application.CommandBars.Add(name:=Mname, Position:=msoBarPopup, _
        MenuBar:=False, Temporary:=True)
        Select Case LCase(ActiveSheet.name)

            Case "voorblad"

            Set MenuItem = .Controls.Add(Type:=msoControlPopup)
            With MenuItem
                .Caption = "Ga naar &calculatie"
                For k = 1 To 8 Step 6
                    For b = 2 To 6

                        If Sheets("voorblad").Cells(b, k).Value = True Then
                            With .Controls.Add(Type:=msoControlButton)
                                naam = IIf(Sheets("voorblad").Cells(b, k + 1).Value = "", "Calculatie " & Sheets("voorblad").Cells(b, k + 2).Value, Sheets("voorblad").Cells(b, k + 1).Value)
                                .Caption = naam
                                .OnAction = "'" & ThisWorkbook.name & "'!'calculatie_tonen """ & naam & """'"
                            End With
                        End If
                    Next b
                Next k
            End With

            Case Else

            With ActiveSheet
                laatste_groep_rij = max_groep_rij
            End With

            With ActiveSheet
                LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
            End With

            If cel_in_tabel_aanwezig(ActiveCell) Then

                If LCase(Cells(21, ActiveCell.Column).Value) = "artikelnr" Then
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "&Artikel zoeken"
                        .OnAction = "'" & ThisWorkbook.name & "'!" & "artikelen_zoeken"
                    End With
                End If

                '                If LCase(Cells(22, ActiveCell.Column).Value) = "werkpost" Then
                '                    With .Controls.Add(Type:=msoControlButton)
                '                        .Caption = "&Werkpost toevoegen"
                '                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "werkposten_zoeken"
                '                    End With
                '                End If

                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "Selectie &kopiëren"
                    .OnAction = "'" & ThisWorkbook.name & "'!" & "kopieren"
                    .BeginGroup = True
                End With

                '                    Set olo = ActiveCell.ListObject
                '    Set ocell = ActiveCell
                '
                '    If cel_in_tabel_aanwezig(ocell) Then
                '        r = InputBox("Hoeveel rijen invoegen?", "Aantal rijen?", 1)
                '
                '        If IsNumeric(r) Then
                '            beveiliging (False)
                '            For rij = 1 To r
                '                lrow = ocell.Row - olo.DataBodyRange.Cells(1, 1).Row + 1
                '                Application.StatusBar = "Invoegen rij " & rij
                '                ActiveCell.ListObject.ListRows.Add (lrow)
                '            Next rij
                '            beveiliging (True)
                '        End If
                '    End If 'If cel_in_tabel_aanwezig(cel)
                '
                'verder:
                '
                'Stop
                Dim olo
                Set olo = ActiveCell.ListObject
                If InStr(1, olo.name, "wtb_var", vbTextCompare) > 0 Then
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "Subgroep kopiëren"
                        .OnAction = "'" & ThisWorkbook.name & "'!" & "subgroep_kopieren"
                        .BeginGroup = True
                    End With
                End If
                If ActiveCell.Locked = False Then
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "&Plakken als waarde"
                        .OnAction = "'" & ThisWorkbook.name & "'!" & "plakken"
                    End With

                    '                Set MenuItem = .Controls.Add(Type:=msoControlPopup)
                    '               With MenuItem
                    '                  .Caption = "Subgroep rijen"
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "Rij(en) &toevoegen"
                        .OnAction = "'" & ThisWorkbook.name & "'!" & "rijen_invoegen"
                        .BeginGroup = True
                    End With 'With .Controls.Add(Type:=msoControlButton)

                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "Rij(en) &verwijderen"
                        .OnAction = "'" & ThisWorkbook.name & "'!" & "rijen_verwijderen"
                    End With 'With MenuItem
                End If 'If ActiveCell.Locked = False

                Dim antwoord, verwijderen

                Set lo = ActiveCell.ListObject
                If Not lo Is Nothing Then
                    If InStr(1, lo, "_vast", vbTextCompare) = 0 Then
                        With .Controls.Add(Type:=msoControlButton)
                            .OnAction = "'" & ThisWorkbook.name & "'!" & "subgroep_verwijderen"
                            .Caption = "Calculatie sub&groep verwijderen"
                            .BeginGroup = True
                        End With
                    End If
                End If

            Else ' If cel_in_tabel_aanwezig(ActiveCell)

                If ActiveCell.row < Range("wtb_vast").row - 1 And ActiveCell.row > Range("wtb_var").row - 1 Then
                    If (ActiveCell.row < Range(Range(groepnaam & max_groep_rij).Value).row) And (ActiveCell.row > Range(Range(groepnaam & 1).Value).row) Then
                        'regels verzamelen van de kopregels van de subgroepen waarin geen subgroep ingevoegd mag worden
                        For Each lo In ActiveSheet.ListObjects
                            naam_rijen = naam_rijen & "|" & lo.Range.row - 1 & "|"
                        Next lo
                        'regels verzamelen van regels waarin titels staan van de groepen, hierin mag geen subgroep worden ingevoegd
                        For rij = 1 To max_groep_rij
                            naam_rijen = naam_rijen & "|" & Range(Range(groepnaam & rij).Value).row & "|"
                        Next rij  ' For rij = 1 To max_groep_rij
                        If InStr(1, naam_rijen, "|" & ActiveCell.row & "|", vbTextCompare) = 0 Then
                            For Each lo In ActiveSheet.ListObjects
                                If lo.Range.row - 1 <> ActiveCell.row And lo.Range.row + lo.Range.Rows.Count - 1 <> ActiveCell.row Then
                                    With .Controls.Add(Type:=msoControlButton)
                                        .OnAction = "'" & ThisWorkbook.name & "'!" & "subgroep_invoegen"
                                        .Caption = "Calculatie subgroep invoegen"
                                        Exit For
                                    End With
                                End If
                            Next lo
                        End If
                    End If
                End If
            End If

            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "'" & ThisWorkbook.name & "'!" & "navigatie_scherm_starten"
                .BeginGroup = True
                .Caption = "Toon navigatiescherm"
            End With


            '            Set MenuItem = .Controls.Add(Type:=msoControlPopup)
            '            With MenuItem
            '                .Caption = "Ga naar..."
            '
            '                Set drop = .Controls.Add(Type:=msoControlComboBox)
            '                With drop
            '                    .Width = 30
            '                    .BeginGroup = True
            '                    .Caption = "Groep"
            '                    .Tag = "groep"
            '                    .AddItem Replace("Ga naar groep", "_", " ")
            '                    For rij = 1 To laatste_groep_rij
            '                        naam = Cells(rij, Range("groep_naam").Column).Value
            '                        .AddItem Replace(naam, "_", " ")
            '                        .OnAction = "'" & ThisWorkbook.Name & "'!'navigeer_naar_groep """ & naam & """'"
            '                    Next rij
            '                End With
            '
            '                Set drop = .Controls.Add(Type:=msoControlComboBox)
            '                With drop
            '                    .Width = 30
            '                    .BeginGroup = True
            '                    Dim li
            '                    .Caption = "subgroep"
            '                    .Tag = "subgroep"
            '                    For Each li In ActiveSheet.ListObjects
            '                        If li.Name <> "template_tabel" Then
            '                            naam = Range(naam_kolom & li.DataBodyRange.Row - 1).Value
            '                            If naam <> "" Then
            '                                .AddItem naam
            '                            End If
            '                        End If
            '                    Next li
            '                    .OnAction = "'" & ThisWorkbook.Name & "'!'navigeer_naar_groep'"
            '                End With
            '
            '            End With

            '            Set drop = .Controls.Add(Type:=msoControlComboBox)
            '            With drop
            '                .Width = 30
            '                .Tag = "subgroep"
            '                .BeginGroup = True
            '                For Each li In ActiveSheet.ListObjects
            '                    If li.Name <> "template_tabel" Then
            '                        naam = Range(naam_kolom & li.DataBodyRange.Row - 1).Value
            '                        If naam <> "" Then
            '                            .AddItem naam
            '
            '                        End If
            '                    End If
            '                Next li
            '                .OnAction = "'" & ThisWorkbook.Name & "'!'navigeer_naar_groep'"
            '            End With

            Set MenuItem = .Controls.Add(Type:=msoControlPopup)
            With MenuItem
                .Caption = "Tonen/verbergen"

                With .Controls.Add(Type:=msoControlButton)
                    .OnAction = "'" & ThisWorkbook.name & "'!" & "'tarieven_tonen_verbergen'"
                    .BeginGroup = True
                    If Range("tarieven_tonen") Then
                        .Caption = "Verberg &tarieven"
                    Else
                        .Caption = "Toon tarieven"
                    End If
                End With

                With .Controls.Add(Type:=msoControlButton)
                    .OnAction = "'" & ThisWorkbook.name & "'!" & "'uren_tonen_verbergen'"
                    If Range("totaal_uren_tonen") Then
                        .Caption = "Verberg totaal uren"
                    Else
                        .Caption = "Toon totaal uren"
                    End If
                End With

                With .Controls.Add(Type:=msoControlButton)
                    .OnAction = "'" & ThisWorkbook.name & "'!" & "'regels_met_kosten_tonen'"
                    If Range("regels_zonder_kosten_tonen") Then
                        .Caption = "Verberg subgroepregels zonder kosten"
                    Else
                        .Caption = "Toon subgroepregels zonder kosten"
                    End If
                End With

                With .Controls.Add(Type:=msoControlButton)
                    .OnAction = "'" & ThisWorkbook.name & "'!" & "'groepen_met_kosten_tonen'"
                    If Range("groepen_zonder_kosten_verbergen") Then
                        .Caption = "Verberg subgroepen zonder kosten"
                    Else
                        .Caption = "Toon subgroepregels zonder kosten"
                    End If
                End With
            End With


            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "'" & ThisWorkbook.name & "'!" & "groepen_samenvatten_tonen"
                .BeginGroup = True
                If Range("samenvatten") Then
                    .Caption = "Groepen tonen"
                Else
                    .Caption = "Groepen samenvatten"
                End If
            End With

            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "'" & ThisWorkbook.name & "'!" & "subgroepen_samenvatten_tonen"
                If Range("subsamenvatten") Then
                    .Caption = "Subgroepen tonen"
                Else
                    .Caption = "Subgroepen samenvatten"
                End If
            End With
        End Select

        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.name & "'!" & "afdrukken"
            .BeginGroup = True
            .Caption = "Calculatie afdrukken"
        End With

        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.name & "'!" & "werkposten_optelling_tonen"
            .BeginGroup = True
            .Caption = "Werkposten optelling tonen"
        End With

    End With
End Sub

Sub navigeer_naar_groep()
    Dim rij
    Dim li
    Dim octl As CommandBarControl
    Dim naam

    For Each octl In CommandBars("MyPopUpMenu").Controls
        If octl.Tag <> "" Then Debug.Print octl.Tag
        If octl.Tag = "subgroep" Then
            naam = octl.text '  Debug.Print octl.text '.ListIndex = 2
        End If
    Next octl

    If naam <> "" Then


        For Each li In ActiveSheet.ListObjects
            rij = li.DataBodyRange.row - 1
            If Range(naam_kolom & li.DataBodyRange.row - 1).Value = naam Then

                Exit For
            End If
        Next li

        With ActiveWindow.Panes(2)
            .ScrollRow = rij
        End With
    End If
End Sub
Sub navigeer_naar_subgroep(naam)
    Dim rij
    For rij = Range("begin_calculatie").row + 1 To Range("einde_calculatie").row
        If Cells(rij, 1).Value = naam Then
            With ActiveWindow.Panes(2)
                .ScrollRow = rij
                Exit For
            End With
        End If
    Next rij
End Sub

Sub kopieren(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Selection.Copy
End Sub

Sub plakken(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)

    Application.DisplayAlerts = False

    'save previous selection
    OldRange = NewRange

    'get current selection
    NewRange = Selection.Address

    'check if copy mode has been turned off

    If Application.CutCopyMode = 1 Then

        On Error GoTo melding
        ActiveCell.PasteSpecial xlPasteValues
    End If

    '     'if copy mode has been turned on, save Old Range
    '    If Application.CutCopyMode = 1 And ChangeRange = False Then
    '         'boolean to hold "SaveRange" address til next copy/paste operation
    '        ChangeRange = True
    '         'Save last clipboard contents range address
    '        SaveRange = OldRange
    '    End If
einde:
    Application.DisplayAlerts = True
    End
melding:
    MsgBox "De data kan hier niet worden geplakt, mogelijk is de ruimte te beperkt...", vbInformation
    Err.Clear
    GoTo einde

End Sub

Sub menu_aan(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)

    Range("menu").Value = True

End Sub

Sub menu_uit(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim oude_sheet
    Application.ScreenUpdating = False
    Range("menu").Value = False
    oude_sheet = ActiveSheet.name

    '    Sheets(1).Select
    Sheets(oude_sheet).Select
    ActiveSheet.Unprotect
    Application.ScreenUpdating = True
End Sub





