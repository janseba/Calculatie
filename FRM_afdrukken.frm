VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_afdrukken 
   Caption         =   "Afdrukken calculatie"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11895
   OleObjectBlob   =   "FRM_afdrukken.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_afdrukken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Afdrukken_Click()
    MsgBox "printerdeprint", vbInformation
End Sub

Private Sub btn_alles_deselecteren_Click()
    selecteren (False)
End Sub

Private Sub btn_alles_selecteren_Click()
    selecteren (True)
End Sub

Private Sub btn_Annuleren_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    With Sheets("Voorblad")
        For teller = 1 To 10
            If .Range("B" & teller + 1).Value <> "" Then
                tekst = .Range("B" & teller + 1).Value
            Else
                tekst = "Calculatie " & teller
            End If
            FRM_afdrukken.Controls("calculatie_" & teller).Caption = tekst
        Next teller
    End With
End Sub

Sub selecteren(wat)
    For Each ctl In Me.Controls
        If InStr(1, ctl.Name, "calculatie", vbTextCompare) Then
            ctl.Value = wat
        End If
    Next ctl
End Sub
