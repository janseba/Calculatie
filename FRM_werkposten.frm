VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_werkposten 
   Caption         =   "Werkpost informatie"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4380
   OleObjectBlob   =   "FRM_werkposten.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_werkposten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LB_werkposten_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Dim i As Long
Dim Msg As String
 With LB_werkposten
    If .ListIndex <> -1 Then
        werkpost = .Column(0)
    End If
End With
    
'If LCase(ActiveSheet.Name) = "calculatie" Then
If InStr(1, ActiveSheet.name, "calculatie", vbTextCompare) > 0 Then
Cells(ActiveCell.row, 7).Value = werkpost
'End With
End If
Unload Me
End Sub

Private Sub TextBox1_Change()
If Len(TextBox1.Value) >= 2 Then
LB_werkposten.Clear
Set Bs = Sheets("basisinformatie")
With Bs
For rij = 2 To 89

celinhoud = .Cells(rij, Range("D1").Column).Value

With LB_werkposten
 If InStr(1, LCase(celinhoud), LCase(TextBox1.Value), False) > 0 Then
    .AddItem Bs.Cells(rij, Bs.Range("A1").Column).Value
    kolom = 1
    .List(.ListCount - 1, kolom) = Bs.Cells(rij, Bs.Range("B1").Column).Value: kolom = kolom + 1
End If
End With

Next rij
End With ' With sheets("basisinformatie")
End If ' if len(TextBox1.Value)>2 THen
End Sub

Private Sub UserForm_Initialize()
With LB_werkposten
    .ColumnCount = 2
    .ColumnWidths = "30;100"
End With
End Sub
