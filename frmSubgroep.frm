VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSubgroep 
   Caption         =   "Subgroep toevoegen"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4755
   OleObjectBlob   =   "frmSubgroep.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSubgroep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAnnuleren_Click()
With Application
.ScreenUpdating = True
.EnableEvents = True
.Calculation = xlCalculationAutomatic
End With
End
End Sub

Private Sub btnToevoegen_Click()
beveiliging (False)
ActiveSheet.Range("aan_te_maken_subgroep").Value = frmSubgroep.lb_subgroep.Value
Unload frmSubgroep
beveiliging (True)
End Sub


Private Sub lb_subgroep_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
beveiliging (False)
ActiveSheet.Range("aan_te_maken_subgroep").Value = frmSubgroep.lb_subgroep.Value
beveiliging (True)
Unload frmSubgroep
End Sub

Private Sub txt_Zoek_subgroep_Change()
If Len(txt_Zoek_subgroep) >= 3 Then


End If
End Sub

Private Sub UserForm_Initialize()
Dim lr
Dim subgroepen As String

frmSubgroep.lb_subgroep.Clear

With Sheets("subgroepen")
        lr = .Cells(.Rows.Count, Range("A1").Column).End(xlUp).Row
subgroepen = ""
For r = 2 To lr
If InStr(1, subgroepen, .Cells(r, 1).Value, vbTextCompare) = 0 Then subgroepen = subgroepen & "|" & .Cells(r, 1).Value & "|"
Next r

subgroepen = Mid(subgroepen, 2, Len(subgroepen) - 2)

For Each subgroep In Split(subgroepen, "||")
frmSubgroep.lb_subgroep.AddItem subgroep
Next subgroep

End With

End Sub

