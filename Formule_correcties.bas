Attribute VB_Name = "Formule_correcties"
Option Explicit
Private Sub formule_correcties_doorvoeren()

    'doorlopen van alle tabellen om formules aan te kunnen passen
    'kan nog wel eens handig zijn

    Dim li As ListObject
    Dim aantal_rij
Dim kolommen
Dim kolom

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    beveiliging (False)
    For Each li In ActiveSheet.ListObjects
        aantal_rij = li.Range.row - 1
        If li.name <> "template_tabel" And li.name <> "wtb_vast_reiskosten" Then

            Range(li.name & "[Kolom7]") = "=[@Kolom4]*[@Kolom5]*[@Kolom6]*$E$" & aantal_rij
            Range(li.name & "[Kolom9]") = "=[@Kolom8]*$E$" & aantal_rij
            'Range(li.Name & "[Kolom9]").FormulaR1C1 = "=IF([@Kolom8]<>0,R4C23,"""")"
            Range(li.name & "[Kolom10]").FormulaR1C1 = "=R4C23"
             Range(li.name & "[Kolom11]").FormulaR1C1 = "=IF([@Kolom8]<>0,[@Kolom9]*[@Kolom10],"""")"
            'Range(li.Name & "[Kolom13]") = "=[@Kolom12]*$E$" & aantal_rij
            Range(li.name & "[Kolom13]").FormulaR1C1 = "=IF([@Kolom12]<>0,[@Kolom12]*R26C5,"""")"
            'Range(li.Name & "[Kolom14]").FormulaR1C1 = "=R4C23"
            Range(li.name & "[Kolom14]").FormulaR1C1 = "=IF([@Kolom12]<>0,R4C23,"""")"
            'Range(li.Name & "[Kolom15]") = "=[@Kolom13]*[@Kolom14]"
            Range(li.name & "[Kolom15]").FormulaR1C1 = "=IF([@Kolom12]<>0,[@Kolom13]*[@Kolom14],"""")"
            Range(li.name & "[Kolom16]").FormulaR1C1 = "=IF([@Kolom9]<>"""",[@Kolom9],0)+IF([@Kolom13]<>"""",[@Kolom13],0)"
            Range(li.name & "[Kolom17]").FormulaR1C1 = "=IF([@Kolom11]<>"""",[@Kolom11],0)+IF([@Kolom15]<>"""",[@Kolom15],0)"
            Cells(li.DataBodyRange.row - 1, Range(li.name & "[Kolom18]").Column).FormulaR1C1 = Replace(Range("S" & li.DataBodyRange.row - 1).FormulaR1C1, "template_tabel", li.name)

kolommen = "15,7,9,10,13,14"
li.Range.FormatConditions.Delete
For Each kolom In Split(kolommen, ",")
' Range(li.Name & "[Kolom15]")

    With Range(li.name & "[kolom" & kolom & "]") ', li.Name & "kolom[9]")
.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
  '.FormatConditions.Add Type:=xlExpression, Formula1:="=0"
  .FormatConditions(.FormatConditions.Count).NumberFormat = ";;;" 'change for other color when ticked
End With
Next kolom
    
    
    
'    Range("P27").Select
 
        End If
    Next li
    beveiliging (True)
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub

Sub test()
Dim lo As ListObject
Dim col 'As Range
Set lo = ActiveSheet.ListObjects("wtb_variabel_1_calculatie_3__Lagedruk afscheider")
For col = 1 To lo.ListColumns.Count
Debug.Print lo.ListColumns(col).name
Debug.Print
Next col

''    Range("Q27").Select
'Range(li.Name & "[Kolom16]").FormulaR1C1 = "=[@Kolom11]+IF([@Kolom15]<>"""",[@Kolom15]+0)"
'    Range("Q27").Select
'    Range(li.Name & "[Kolom16]").FormulaR1C1 = "=[@Kolom11]+IF([@Kolom15]<>"""",[@Kolom15],0)"
'  '  Range("Q27").Select
'    Range(li.Name & "[Kolom16]").FormulaR1C1 = _
'        "=IF([@Kolom11]<>"""",[@Kolom11],0)+IF([@Kolom15]<>"""",[@Kolom15],0)"
' '   Range("Q27").Select
'    Range(li.Name & "[Kolom16]").FormulaR1C1 = _
'        "=IF([@Kolom9]<>"""",[@Kolom9],0)+IF([@Kolom13]<>"""",[@Kolom13],0)"
'' Range("Q27").Select
'    Range(li.Name & "[Kolom16]").FormulaR1C1 = _
'        "=IF([@Kolom9]<>"""",[@Kolom9],0)+IF([@Kolom13]<>"""",[@Kolom13],0)"
''    Range("R27").Select
'    Range(li.Name & "[Kolom17]").FormulaR1C1 = _
'        "=IF([@Kolom11]<>"""",[@Kolom11],0)+IF([@Kolom15]<>"""",[@Kolom15],0)"

With lo



End With

End Sub
