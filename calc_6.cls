VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "calc_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ToggleButton1_Click()
toon_tarieven
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub koudemiddel_Change()


For Each tabel In ActiveSheet.ListObjects

voor_gedef = InStr(1, tabel, "__", vbTextCompare)

If InStr(1, tabel, "wtb_var", vbTextCompare) > 0 And voor_gedef > 0 Then

template = Mid(tabel, voor_gedef + 2)




'wtb_variabel_1_calculatie_3_Lagedruk_afscheider



'wtb_variabel_7_calculatie_3



End If

Next tabel

End Sub
Private Sub Worksheet_Deactivate()
    Dim namedRange As name, tblKengetallenCalculatie As ListObject, calculatie As Integer, i As Integer
    Dim row As ListRow
    calculatie = 1
    Set tblKengetallenCalculatie = wksRekenblad.ListObjects("tblKenGetallenCalculatie")
    
    Application.AutoCorrect.AutoFillFormulasInLists = False
    With tblKengetallenCalculatie
        For Each row In .ListRows
            row.Range(3).Formula = "=" & Me.name & "!" & Me.Names(Me.name & "!calc.mat." & row.Range(2).Value).RefersToRange.Address
            row.Range(4).Formula = "=" & Me.name & "!" & Me.Names(Me.name & "!calc.uren." & row.Range(2).Value).RefersToRange.Address
            row.Range(5).Formula = "=" & Me.name & "!" & Me.Names(Me.name & "!calc.kosten." & row.Range(2).Value).RefersToRange.Address
        Next row
    End With
    Application.AutoCorrect.AutoFillFormulasInLists = True
    
End Sub
