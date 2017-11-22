VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetActivate(ByVal sh As Object)

If LCase(sh.Name) = "voorblad" Then
sh.EnableCalculation = True
Else
Sheets("voorblad").EnableCalculation = False
End If

If (InStr(1, sh.Name, GetWorksheetFromCodeName(sh.Name), vbTextCompare) > 0 Or LCase(sh.Name = "voorblad")) And Range("menu").Value = True Then

Application.CellDragAndDrop = False
Else
With Application
.CommandBars("List Range Popup").Enabled = True
.CommandBars("Cell").Enabled = True
With .CommandBars("Column")
.Enabled = True
.Reset
End With
With .CommandBars("Row")
.Enabled = True
.Reset
End With

Application.CellDragAndDrop = True
End With
End If
End Sub

Private Sub Workbook_SheetBeforeRightClick(ByVal sh As Object, ByVal Target As Range, Cancel As Boolean)
Dim bladnaam
'Stop

Call DeletePopUpMenu

bladnaam = GetCodenameFromWorksheet(sh.Name)

If (InStr(1, bladnaam, "calc", vbTextCompare) > 0 Or LCase(bladnaam) = "voorblad") And Range("menu").Value = True Then
menus (False)
Call CreateDisplayPopUpMenu
Else
menus (True)
End If
End Sub

Private Sub Workbook_Activate()
menus (False)
End Sub

Private Sub Workbook_Deactivate()
menus (True)
With Application
.Calculation = xlCalculationAutomatic
.EnableEvents = True
.CellDragAndDrop = True
End With

End Sub

Private Sub Worbook_Open()
menus (False)
Application.EnableEvents = True
End Sub

Private Sub Workbook_Close()
menus (True)
Application.EnableEvents = True
End Sub


