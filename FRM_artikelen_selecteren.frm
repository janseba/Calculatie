VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_artikelen_selecteren 
   Caption         =   "Artikelen selecteren"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14250
   OleObjectBlob   =   "FRM_artikelen_selecteren.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_artikelen_selecteren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LB_gevonden_artikelen_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Dim i As Long
Dim Msg As String
 With LB_gevonden_artikelen
    If .ListIndex <> -1 Then
        artikelomschrijving = .Column(0)
        artikelnummer = .Column(1)
        prijs = .Column(2)
        werkpost = .Column(3)
    End If
End With
    
    With ActiveSheet
If InStr(1, LCase(.Name), "calculatie", 0) Then

.Cells(ActiveCell.Row, 3).Value = artikelnummer
.Cells(ActiveCell.Row, 4).Value = artikelomschrijving
.Cells(ActiveCell.Row, 6).Value = prijs
'If werkpost <> "" Then
'.Cells(ActiveCell.Row, 8).Value = werkpost
'End If
End If
End With

'End If
End Sub

Private Sub TXT_omschrijving_Change()


If Len(TXT_omschrijving) > 2 Then

LB_gevonden_artikelen.Clear

With Sheets("prijslijst to be")
laatste_rij = .Cells(.Rows.Count, "D").End(xlUp).Row
For rij = 3 To laatste_rij
celinhoud = .Cells(rij, Range("D1").Column).Value

tekst = Replace(TXT_omschrijving, " ", "|")
If Right(tekst, 1) = "|" Then tekst = Mid(tekst, 1, Len(tekst) - 1)
zoek = Split(tekst, "|")

aanwezig = 0
For t = 0 To UBound(zoek)
If zoek(t) <> "" Then
If Len(zoek(t)) >= 2 Then
If InStr(1, LCase(celinhoud), LCase(zoek(t)), vbTextCompare) > 0 Then

aanwezig = aanwezig + 1
End If
End If
End If
Next t

If aanwezig = UBound(zoek) + 1 Then
With LB_gevonden_artikelen
    .AddItem celinhoud
    kolom = 1
    .List(.ListCount - 1, kolom) = Sheets("prijslijst to be").Cells(rij, Sheets("prijslijst to be").Range("a1").Column).Value: kolom = kolom + 1
    .List(.ListCount - 1, kolom) = Sheets("prijslijst to be").Cells(rij, Sheets("prijslijst to be").Range("K1").Column).Value: kolom = kolom + 1
    .List(.ListCount - 1, kolom) = Sheets("prijslijst to be").Cells(rij, Sheets("prijslijst to be").Range("L1").Column).Value: kolom = kolom + 1
End With
End If
Next rij

End With
Else
LB_gevonden_artikelen.Clear
End If
End Sub

Private Sub UserForm_Initialize()
With LB_gevonden_artikelen
    .ColumnCount = 4
    .ColumnWidths = "225;100;100;50"
End With

End Sub
