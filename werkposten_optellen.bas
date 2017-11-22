Attribute VB_Name = "werkposten_optellen"
Sub werkposten_optelling_tonen()
Overzicht_werkposten.Show
End Sub

Sub werkposten_sommeren()
'Toon een popup met daarin de optelling van de verschillende werkposten
Dim bedrag As dictionary
Dim percentage As dictionary
Dim lo As ListObject
Dim imgname As String

With Application
.DisplayAlerts = False
.ScreenUpdating = False
.Calculation = xlCalculationManual
.EnableEvents = False
.Cursor = xlWait
End With

Set bedrag = New dictionary
Set percentage = New dictionary

totaal = 0

With ActiveSheet
For Each lo In .ListObjects
If .Cells(lo.Range.row - 1, 19).Value <> 0 Then

With .Cells(lo.Range.row - 1, 1)
wp = IIf(.Value = "", "Onbekend", .Value)
End With
bedrag(wp) = bedrag(wp) + .Cells(lo.Range.row - 1, 19).Value
totaal = totaal + .Cells(lo.Range.row - 1, 19).Value
End If
Next lo
End With

If IsUserFormLoaded("Overzicht_werkposten") Then Overzicht_werkposten.lb_Werkpost_tellen.Clear

On Error Resume Next

Sheets("werkpost_grafiek").Delete
On Error GoTo 0
Set werkpost_grafiek = Sheets.Add
werkpost_grafiek.name = "werkpost_grafiek"
werkpost_grafiek.Visible = False
Application.ScreenUpdating = True
Application.DisplayAlerts = True

rij = 0
p = 0
p = p + 1

With Overzicht_werkposten.lb_Werkpost_tellen
.AddItem
.List(p - 1, 0) = "Werkpost"
.List(p - 1, 1) = "Bedrag"
.List(p - 1, 2) = "Aandeel in totaal"
End With

For Each i In bedrag
rij = rij + 1
p = p + 1
werkpost_grafiek.Cells(rij, 1).Value = i
werkpost_grafiek.Cells(rij, 2).Value = bedrag(i)

With Overzicht_werkposten.lb_Werkpost_tellen
.AddItem
.List(p - 1, 0) = i
bed = bedrag(i)
If Int(bed) = bed Then
bed = bed & ",00"
End If

bed = "                " & bed
.List(p - 1, 1) = "€" & Right(bed, 10)

per = Round((bedrag(i) / totaal) * 100, 2)
per = "            " & per & " %"
.List(p - 1, 2) = Right(per, 9)
End With
Next i

Set ChartData = werkpost_grafiek.Range("b1:b20")
chartname = "Verdeling werkposten"
    Set grafiek = werkpost_grafiek.Shapes.AddChart(xlPie).Chart
    
With grafiek
.SeriesCollection(1).name = "=" & werkpost_grafiek.name & "!$A$1:$A$" & rij
.SeriesCollection(1).Values = "=" & werkpost_grafiek.name & "!$B$1:$B$" & rij
.SeriesCollection(1).XValues = "=" & werkpost_grafiek.name & "!$A$1:$A$" & rij
.ChartTitle.text = "Verdeling werkposten"
.ChartArea.Width = Overzicht_werkposten.lb_Werkpost_tellen.Width
.ChartArea.Height = Overzicht_werkposten.lb_Werkpost_tellen.Height
imgname = Environ("temp") & Application.PathSeparator & "tmpchart.gif"
.Export Filename:=imgname
End With

Overzicht_werkposten.Image1.Picture = LoadPicture(imgname)

With Application
.DisplayAlerts = True
.ScreenUpdating = True
.Calculation = xlCalculationAutomatic
.EnableEvents = True
.Cursor = xlDefault
End With

End Sub

Function IsUserFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object
     
    IsUserFormLoaded = False
    For Each UForm In VBA.UserForms
        If UForm.name = UFName Then
            IsUserFormLoaded = True
            Exit For
        End If
    Next
End Function 'IsUserFormLoaded
