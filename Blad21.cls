VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Blad21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub worksheet_Change(ByVal Target As Range)
'    Dim naam As String
'    Dim rest_waarde
'    Dim cel As Range
'
'    '    On Error GoTo foutje
'
'    Application.EnableEvents = False
'
'    For Each cel In Target
'
'        Select Case cel.Address
'            Case [offerte_aanneemsom].Address
'
'                If Target.Value = 0 Then
'                    Target.Formula = "= frigotronic_aanneemsom_offerte"
'                Else
'
'                    rest_waarde = Target.Value - ([frigotronic_aanneemsom_offerte])
'                    rest_waarde = Replace(rest_waarde, ",", ".")
'                    Target.Formula = "=" & rest_waarde & " + frigotronic_aanneemsom_offerte"
'
'                End If
'
'            Case [opdracht_aanneemsom].Address
'                If cel.Value = 0 Then
'                    cel.Formula = "= + frigotronic_aanneemsom_opdracht"
'                Else
'                    cel.Formula = "=" & Replace(cel.Value - ([frigotronic_aanneemsom_opdracht]), ",", ".") & " + frigotronic_aanneemsom_opdracht"
'                End If
'
'                naam = ActiveSheet.Name
'                Application.ScreenUpdating = False
'                Call opdracht_vrijgeven(cel >= 100000 Or cel = 0)
'                Application.ScreenUpdating = True
'                Sheets(naam).Select
'                cel.Select
'
'            Case [offerte_werkderden].Address
'                cel.Formula = "=" & cel.Formula & " + frigotronic_werkderden_offerte + IF(begroting_bedoeld_voor=""offerte"", onderaannemer_totaal_kosten,0)"
'
'            Case [opdracht_werkderden].Address
'                cel.Formula = "=" & cel.Formula & " + frigotronic_werkderden_opdracht + IF(begroting_bedoeld_voor=""opdracht"", onderaannemer_totaal_kosten,0)"
'
'            Case [offerte_materiaal].Address
'                cel.Formula = "=" & cel.Formula & " + frigotronic_materiaal_offerte + IF(begroting_bedoeld_voor=""offerte"",materiaal_totaal,0)"
'
'            Case [opdracht_materiaal].Address
'                cel.Formula = "=" & cel.Formula & " + frigotronic_materiaal_opdracht + IF(begroting_bedoeld_voor=""opdracht"",materiaal_totaal,0)"
'
'            Case [offerte_reisuren].Address
'                cel.Formula = "=" & cel.Formula & " + IFERROR( frigotronic_reisuren_offerte*1,0) + IF(begroting_bedoeld_voor=""offerte"",IFERROR(reisuren_nieuwbouw_monteur*1,0) + IFERROR(reisuren_inbedrijfstel_monteur*1,0),0)"
'
'            Case [opdracht_reisuren].Address
'                cel.Formula = "=" & cel.Formula & " + IFERROR( frigotronic_reisuren_opdracht * 1,0) + IF(begroting_bedoeld_voor=""opdracht"",IFERROR(reisuren_nieuwbouw_monteur*1,0) + IFERROR(reisuren_inbedrijfstel_monteur*1,0),0)"
'
'            Case [offerte_montageuren].Address
'                cel.Formula = "=" & cel.Formula & " + frigotronic_montageuren_offerte+IF(begroting_bedoeld_voor=""offerte"",aantal_uur_nwb_locatie+aantal_uur_ibs+nb_monteurs_werkplaats_aantal,0)"
'
'            Case [opdracht_montageuren].Address
'                cel.Formula = "=" & cel.Formula & " + frigotronic_montageuren_opdracht + IF(begroting_bedoeld_voor=""opdracht"",aantal_uur_nwb_locatie+aantal_uur_ibs+nb_monteurs_werkplaats_aantal,0)"
'
'            Case [offerte_wvburen].Address
'                cel.Formula = "=" & cel.Formula & " + frigotronic_wvburen_offerte + frigotronic_wvbeturen_offerte +  IF(begroting_bedoeld_voor=""offerte"" , aantal_uur_wvb + aantal_uur_cad + aantal_uur_elektrotechniek,0)"
'
'            Case [opdracht_wvburen].Address
'                cel.Formula = "=" & cel.Formula & " + frigotronic_wvburen_opdracht+frigotronic_wvbeturen_opdracht + IF(begroting_bedoeld_voor=""opdracht"" , aantal_uur_wvb + aantal_uur_cad + aantal_uur_elektrotechniek,0)"
'
'            Case [offerte_overigekostendiversen].Address
'                cel.Formula = "=" & cel.Formula & " + frigotronic_projectkosten_diversen_offerte+ IF(begroting_bedoeld_voor=""offerte"",transport_bestelkosten_totaal+hijs_en_hef_totaal+diversen_transport+verblijfkosten_totaal+buitenland_toeslag_totaal+bankgarantie+certificering_kosten ,0)"
'
'            Case [opdracht_overigekostendiversen].Address
'                cel.Formula = "=" & cel.Formula & " + frigotronic_projectkosten_diversen_opdracht + IF(begroting_bedoeld_voor=""opdracht"",transport_bestelkosten_totaal+hijs_en_hef_totaal+diversen_transport+verblijfkosten_totaal+buitenland_toeslag_totaal+bankgarantie+certificering_kosten ,0)"
'
'            Case [offerte_overigekostenmonteurs].Address
'                If cel.Value = 0 Then
'                    '    cel.FormulaR1C1 = "=(offerte_montageuren*IF(offerte_aanvang="""",0,int_off_arbeidsmiddelen_toeslag))-(frigotronic_montageuren_offerte)*IF(offerte_aanvang="""",0,int_off_arbeidsmiddelen_toeslag)+frigotronic_projectkosten_monteurs_offerte + reiskosten_km_totaal + reiskosten_vliegen_overtocht_totaal"
'                    cel.FormulaR1C1 = "=+frigotronic_projectkosten_monteurs_offerte + IF(begroting_bedoeld_voor=""offerte"",reiskosten_km_totaal + reiskosten_vliegen_overtocht_totaal,0)"
'                Else
'                    'rest_waarde = cel.Value - ((offerte_montageuren] * [int_off_arbeidsmiddelen_toeslag]) - ([frigotronic_montageuren_offerte]) * IIf([offerte_aanvang] = """", 0, [int_off_arbeidsmiddelen_toeslag]) + [frigotronic_projectkosten_monteurs_offerte])
'                    rest_waarde = cel.Value - 0    '(([offerte_montageuren] * [int_off_arbeidsmiddelen_toeslag]) - ([frigotronic_montageuren_offerte]) * IIf([offerte_aanvang] = """", 0, [int_off_arbeidsmiddelen_toeslag]) + [frigotronic_projectkosten_monteurs_offerte])
'                    rest_waarde = Replace(rest_waarde, ",", ".")
'                    'cel.Formula = "=" & rest_waarde & " + (offerte_montageuren*int_off_arbeidsmiddelen_toeslag)-(frigotronic_montageuren_offerte)*IF(offerte_aanvang="""",0,int_off_arbeidsmiddelen_toeslag)+frigotronic_projectkosten_monteurs_offerte + reiskosten_km_totaal + reiskosten_vliegen_overtocht_totaal"
'                    cel.Formula = "=" & rest_waarde & " + frigotronic_projectkosten_monteurs_offerte + reiskosten_km_totaal + reiskosten_vliegen_overtocht_totaal"
'                End If
'
'            Case [opdracht_overigekostenmonteurs].Address
'                If cel.Value = 0 Then
'                    'cel.FormulaR1C1 = "=(opdracht_montageuren*IF(opdracht_aanvang="""",0,int_opdr_arbeidsmiddelen_toeslag))-(frigotronic_montageuren_opdracht)*IF(opdracht_aanvang="""",0,int_opdr_arbeidsmiddelen_toeslag)+frigotronic_projectkosten_monteurs_opdracht"
'                    cel.FormulaR1C1 = "=+frigotronic_projectkosten_monteurs_opdracht + IF(begroting_bedoeld_voor=""opdracht"",reiskosten_km_totaal + reiskosten_vliegen_overtocht_totaal,0)"
'                Else
'                    ' rest_waarde = cel.Value - (([opdracht_montageuren] * [int_opdr_arbeidsmiddelen_toeslag]) - ([frigotronic_montageuren_opdracht]) * IIf([opdracht_aanvang] = """", 0, [int_opdr_arbeidsmiddelen_toeslag]) + [frigotronic_projectkosten_monteurs_opdracht])
'                    rest_waarde = cel.Value - 0
'                    rest_waarde = Replace(rest_waarde, ",", ".")
'                    ' cel.Formula = "=" & rest_waarde & " + (opdracht_montageuren*int_opdr_arbeidsmiddelen_toeslag)-(frigotronic_montageuren_opdracht)*IF(opdracht_aanvang="""",0,int_opdr_arbeidsmiddelen_toeslag)+frigotronic_projectkosten_monteurs_opdracht"
'
'                    cel.Formula = "=" & rest_waarde & " + frigotronic_projectkosten_monteurs_opdracht + IF(begroting_bedoeld_voor=""opdracht"", reiskosten_km_totaal + reiskosten_vliegen_overtocht_totaal,0)"
'
'                End If
'
'            Case [offerte_pluren].Address
'                cel.Formula = "=" & cel.Formula & " + frigotronic_pluren_offerte + IF(begroting_bedoeld_voor=""offerte"",aantal_uur_pl,0)"
'
'            Case [opdracht_pluren].Address
'
'                cel.Formula = "=" & cel.Formula & " + frigotronic_pluren_opdracht + IF(begroting_bedoeld_voor=""opdracht"",aantal_uur_pl,0)"
'
'            Case [offerte_montagetarief].Address
'                If IsNumeric(cel.Value) = False Or cel.Value = 0 Then
'                    cel.Formula = "=int_off_montagetarief"
'                End If
'
'            Case [opdracht_montage_tarief].Address
'                If IsNumeric(cel.Value) = False Or cel.Value = 0 Then
'                    cel.Formula = "=int_opdr_montagetarief"
'                End If
'
'            Case Else
'                Call overigekosten_check(cel)
'        End Select
'
'    Next cel
'
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True
'
'    Exit Sub
'
'foutje:
'    Application.EnableEvents = True
'    MsgBox "Er is een fout opgetreden"
End Sub

'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
''controleren of er een cel is geselecteerd waarbij de dropdown moet worden getoond
'    Shapes("shp_marktsector").Visible = Not Application.Intersect([marktsector], Target) Is Nothing
'    Shapes("shp_contact").Visible = Not Application.Intersect([contact], Target) Is Nothing
'    If Target.Address = [offerte_werkderden].Address Then
'    End If
'End Sub
