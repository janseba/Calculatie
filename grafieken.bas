Attribute VB_Name = "grafieken"
Sub grafieken_uitlijnen()
'hiermee worden de grafieken op het voorblad netjes uitgelijnd
'is niet persé iets voor de gebruiker

If LCase(ActiveSheet.name) = "voorblad" Then
For Each co In ActiveSheet.ChartObjects

If InStr(1, co.name, "verdeling_eam") > 0 Then
nummer = Replace(Replace(co.name, "calc", ""), "_verdeling_eam", "")
If nummer = "totaal" Then
nummer = ""
Else
nummer = "_" & nummer
End If
co.Left = 475
co.Top = Cells(Range("begin_calculatie" & nummer).row + 1, Range("begin_calculatie" & nummer).Column).Top
co.Height = Range("a19: A54").Height / 2
End If

If InStr(1, co.name, "verdeling_KEEIW") > 0 Then
grafiek = Replace(co.name, "verdeling_KEEIW", "verdeling_eam")
co.Left = 475
co.Top = ActiveSheet.ChartObjects(grafiek).Top + ActiveSheet.ChartObjects(grafiek).Height
co.Height = Range("a19: A54").Height / 2
End If
co.Width = 775

Next co
End If
End Sub

Sub grafieken_tonen_verbergen()
'hiermee worden de grafieken getoond en verborgen op het voorblad
If LCase(ActiveSheet.name) = "voorblad" Then
For Each co In ActiveSheet.ChartObjects
co.Visible = Not co.Visible
Next co
End If 'If LCase(ActiveSheet.Name) = "voorblad"
End Sub

