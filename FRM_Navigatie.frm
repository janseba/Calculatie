VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_Navigatie 
   Caption         =   "Subgroep navigatie"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5100
   OleObjectBlob   =   "FRM_Navigatie.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_Navigatie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CHK_bedrag_Click()
toon_zoekresultaten
End Sub

Private Sub LB_navigatie_Click()
For Each lb In ActiveSheet.ListObjects
naam = Cells(lb.DataBodyRange.row - 1, lb.DataBodyRange.Column)
If LB_navigatie.Value = naam Then
      With ActiveWindow.Panes(2)
            .ScrollRow = lb.DataBodyRange.row - 3
        End With
        Cells(lb.DataBodyRange.row, lb.DataBodyRange.Column).Select
Exit For
End If
Next lb
End Sub

Private Sub LB_groepen_Click()
For x = 1 To max_groep_rij - 1
If ActiveSheet.Range(groepnaam & x).Offset(0, 1) = lb_groepen.Value Then
With ActiveWindow.Panes(2)
y = ActiveSheet.Range(groepnaam & x)
     With ActiveWindow.Panes(2)
            .ScrollRow = ActiveSheet.Names(y).RefersToRange.row
        End With
End With
Exit For
End If
Next x
End Sub

Private Sub TXT_zoeken_Change()
toon_zoekresultaten
End Sub

Private Sub UserForm_Initialize()
'    HookFormScroll Me
Dim lb As ListObject
With lb_groepen
For rij = 1 To max_groep_rij - 1
.AddItem ActiveSheet.Range(groepnaam & rij).Offset(0, 1)
Next rij
End With
End Sub

Private Sub LB_navigatie_MouseMove( _
             ByVal Button As Integer, ByVal Shift As Integer, _
             ByVal x As Single, ByVal y As Single)
' start tthe hook
     HookListBoxScroll
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     UnhookListBoxScroll
End Sub


Sub SortListBox(oLb As MSForms.ListBox, sCol As Integer, sType As Integer, sDir As Integer)
    Dim vaItems As Variant
    Dim i As Long, j As Long
    Dim c As Integer
    Dim vTemp As Variant
     
     'Put the items in a variant array
    vaItems = oLb.List
     
     'Sort the Array Alphabetically(1)
    If sType = 1 Then
        For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
            For j = i + 1 To UBound(vaItems, 1)
                 'Sort Ascending (1)
                If sDir = 1 Then
                    If vaItems(i, sCol) > vaItems(j, sCol) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                     
                     'Sort Descending (2)
                ElseIf sDir = 2 Then
                    If vaItems(i, sCol) < vaItems(j, sCol) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                End If
                 
            Next j
        Next i
         'Sort the Array Numerically(2)
         '(Substitute CInt with another conversion type (CLng, CDec, etc.) depending on type of numbers in the column)
    ElseIf sType = 2 Then
        For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
            For j = i + 1 To UBound(vaItems, 1)
                 'Sort Ascending (1)
                If sDir = 1 Then
                    If CInt(vaItems(i, sCol)) > CInt(vaItems(j, sCol)) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                     
                     'Sort Descending (2)
                ElseIf sDir = 2 Then
                    If CInt(vaItems(i, sCol)) < CInt(vaItems(j, sCol)) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                End If
                 
            Next j
        Next i
    End If
     
     'Set the list to the array
    oLb.List = vaItems
End Sub

Sub toon_zoekresultaten()

LB_navigatie.Clear

For Each lb In ActiveSheet.ListObjects

tabel_naam = Cells(lb.DataBodyRange.row - 1, lb.DataBodyRange.Column)
bedrag = Cells(lb.DataBodyRange.row - 1, 19)

If Replace(tabel_naam, " ", "") = "" Then tabel_naam = lb.name

If InStr(1, LCase(tabel_naam), LCase(TXT_zoeken.Value), vbTextCompare) > 0 Then
If CHK_bedrag Then
If bedrag > 0 Then
FRM_Navigatie.LB_navigatie.AddItem tabel_naam
End If
Else
FRM_Navigatie.LB_navigatie.AddItem tabel_naam
End If
End If
Next lb

If LB_navigatie.ListCount > 0 Then SortListBox LB_navigatie, 0, 1, 1

End Sub
