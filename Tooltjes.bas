Attribute VB_Name = "Tooltjes"
Sub begintoestand_herstellen()
With Application
.ScreenUpdating = True
.Calculation = xlCalculationAutomatic
.EnableEvents = True
.Cursor = xlDefault
.CellDragAndDrop = True
End With
Call menus(True)
End Sub

Sub menus(aan As Boolean)

With Application
With .CommandBars("List Range Popup")
.Enabled = aan
If aan Then .Reset
End With

With .CommandBars("Cell")
.Enabled = aan
If aan Then .Reset
End With

With .CommandBars("Column")
.Enabled = aan
If aan Then .Reset
End With

With .CommandBars("Row")
.Enabled = aan
If aan Then .Reset
End With

End With
End Sub


Sub template_tabel_tonen()
Application.ScreenUpdating = False
    With Rows("9:19")
    beveiliging (False)
    .EntireRow.Hidden = Not .EntireRow.Hidden
    ' beveiliging (True)
    End With
    Application.ScreenUpdating = True
End Sub

Sub tonen_verbergen()
    dataregels_tonen_verbergen False, False
End Sub

'Sub tonen_verbergen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
'    dataregels_tonen_verbergen False, False
'End Sub

Sub sub_groepen_zonder_kosten_samenvouwen()
    dataregels_tonen_verbergen True, False
End Sub

Private Sub dataregels_tonen_verbergen(verbergen, alles_verbergen)
    Dim li As ListObject

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    lege_tabellen_verbergen = False

    beveiliging (False)

    'door alle tabellen heen lopen
    For Each li In ActiveSheet.ListObjects
        If li.Name <> "template_tabel" Then
            lrow = li.DataBodyRange.Cells(0, 1).Row
            If lege_tabellen_verbergen Then
                If Cells(lrow, 18) = 0 Then
                    li.Range.EntireRow.Hidden = verbergen
                    Cells(lrow, 18).EntireRow.Hidden = verbergen
                    Cells(lrow - 1, 18).EntireRow.Hidden = verbergen
                    Cells(li.Range.Row + li.Range.Rows.Count, 18).EntireRow.Hidden = verbergen
                End If
            Else
                'als het totaal van de tabel 0 of alle regels moeten worden verborgen
                If Cells(lrow, 18) = 0 Or alles_verbergen Then
                    If verbergen Then
                        li.DataBodyRange.EntireRow.RowHeight = 0.1
                    Else
                        'li.DataBodyRange.EntireRow.AutoFit
                        li.DataBodyRange.EntireRow.RowHeight = 15
                    End If
                End If
            End If
        End If
    Next li

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationAutomatic
    End With

    beveiliging (True)

End Sub

Sub rijhoogte_aanpassen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)

    For Each li In ActiveSheet.ListObjects
        If li.Name <> "template_tabel" Then
            For Each rij In li.DataBodyRange.Rows
                If rij.Row = li.DataBodyRange.Cells(0, 1).Row Then
                Else
                    rij.EntireRow.RowHeight = 15
                End If
            Next rij
        End If
    Next li

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationAutomatic
    End With

    beveiliging (True)
End Sub

Private Sub bereiken_met_fouten_verwijderen()

    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    Dim intCounter As Integer
    Dim nmTemp As Name
    Dim bereik As Name

    For Each bereik In ActiveWorkbook.Names

        naam = bereik.Name
        verw = bereik

        Debug.Print naam
        Debug.Print verw

        If InStr(1, bereik, "#REF", vbTextCompare) > 0 Then
            bereik.Delete
        Else
            Debug.Print bereik.RefersToR1C1
        End If

    Next bereik

    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With

End Sub

Sub subgroep_kop_opmaak()
    Dim intCounter As Integer
    Dim nmTemp As Name
    Dim bereik As Name
Dim naam
Dim verw
'Dim bereik As Name

    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
beveiliging (False)
    For Each bereik In ActiveSheet.Names
        naam = bereik.Name
        verw = bereik
        If InStr(1, naam, "_vast", vbTextCompare) > 0 Or _
        InStr(1, naam, "_var", vbTextCompare) > 0 Or _
        InStr(1, naam, "einde_", vbTextCompare) > 0 Then
        Range(verw).RowHeight = 22
    End If
Next bereik
beveiliging (True)

With Application
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
End With

End Sub

Sub tabel_leegmaken()
ActiveSheet.ListObjects("template_tabel").DataBodyRange.ClearContents
End Sub

Sub navigatie_scherm_starten()
FRM_Navigatie.Show
End Sub

Sub opmaak_kop()

 For Each li In ActiveSheet.ListObjects
    rij = li.DataBodyRange.Row - 1
    With Range(Cells(rij, 2), Cells(rij, 4))
    
        With .Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599963377788629
        .PatternTintAndShade = 0
    End With
    .Locked = True
    .FormulaHidden = False
    
    
    End With
    Next li
End Sub


Sub formules_aanpassen()
Application.Calculation = xlCalculationManual
For Each li In ActiveSheet.ListObjects
If li.Name = "template_tabel" Then
    Range(li.Name & "[Kolom9]").FormulaR1C1 = "=[@Kolom8]*R" & li.DataBodyRange.Row - 1 & "C5"
    Range(li.Name & "[Kolom10]").FormulaR1C1 = "=([@Kolom8]<>"""")*R4C23"
    Range(li.Name & "[Kolom11]").FormulaR1C1 = "=IF([@Kolom8]<>0,[@Kolom9]*[@Kolom10],"""")"
    Range(li.Name & "[Kolom13]").FormulaR1C1 = "=R" & li.DataBodyRange.Row - 1 & "C5*[@Kolom12]"
    Range(li.Name & "[Kolom14]").FormulaR1C1 = "=IF([@Kolom12]<>0,R4C23,"""")"
    Range(li.Name & "[Kolom15]").FormulaR1C1 = "=IF([@Kolom12],[@Kolom13]*[@Kolom14],"""")"
    Range(li.Name & "[Kolom16]").FormulaR1C1 = "=IF([@Kolom12],[@Kolom9]+[@Kolom13],0)"
    Range(li.Name & "[Kolom17]").FormulaR1C1 = "=IF([@Kolom12],IF([@Kolom11]<>"""",[@Kolom11],0)+IF([@Kolom15]<>"""",[@Kolom15],0),0)"
End If
Next li
Application.Calculation = xlCalculationAutomatic
End Sub
'---------------------------------------------------------------------------------------
' Procedure : ExportCode
' Author    : Bas Jansen
' Date      : 22-11-2017
' Purpose   : Export all code to folder (to enable version control via Github)
'---------------------------------------------------------------------------------------
'
Public Sub ExportCode()

    Dim wkbTarget As Excel.Workbook, clsCleaner As VBACodeCleaner.CodeCleaner
    Const EXPORTFOLDER As String = "\\Mac\Dropbox for Business\InfoAction Team Folder\E\engie\Calculatie"

    On Error GoTo ErrorHandler

    ''' Create an instance of the CodeCleaner object.
    Set clsCleaner = New VBACodeCleaner.CodeCleaner

    ''' Get a reference to the workbook we'll operate on.
    Set wkbTarget = ThisWorkbook

    If Not clsCleaner.ExportModules(wkbTarget, EXPORTFOLDER) Then
        Err.Raise vbObjectError, , clsCleaner.GetLastError
    End If
    
    Exit Sub

ErrorHandler:

    MsgBox Err.Description, vbCritical, "Code Cleaner Demo"

End Sub
