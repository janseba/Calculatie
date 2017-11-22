Attribute VB_Name = "calculatieblad_handelingen"
Sub opmaak_kopieren()

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Range("A26").Copy
    For Each li In ActiveSheet.ListObjects

        If li.name <> "template_tabel" Then
            With Range("A" & li.DataBodyRange.row - 1)
                Debug.Print .Address
                .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            End With
        End If
    Next li

    Application.CutCopyMode = xlCopy

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

End Sub

Sub calculatie_nieuw_kopieren(handeling)
Dim index_nummer
Dim blad
Dim t

instellingen_ophalen

nieuw = handeling

If handeling = 0 And InStr(1, ActiveSheet.name, "calculatie_", vbTextCompare) = 0 Then
MsgBox "Het geselecteerde blad is geen calculatie!", vbCritical
End
End If

With Application
.ScreenUpdating = False
.EnableEvents = False
.Calculation = xlCalculationManual
End With

For calculatienummer = 1 To max_aantal_calc_bladen
sName = "calculatie_" & calculatienummer
  If Evaluate("ISREF('" & sName & "'!A1)") = False Then Exit For
Next calculatienummer

If calculatienummer > max_aantal_calc_bladen Then
MsgBox "Het maximaal aantal calculaties is toegevoegd", vbInformation
Else
bladnaam = IIf(nieuw, "calculatie_", ActiveSheet.name)
Sheets(bladnaam).Visible = True
Sheets(bladnaam).Copy Before:=Sheets(Sheets("calculatie_").Index + 1)
ActiveSheet.name = "calculatie_" & calculatienummer
Sheets(bladnaam).Visible = Not (bladnaam = "calculatie_")

calculaties_sorteren
End If

With Application
.ScreenUpdating = True
.EnableEvents = True
.Calculation = xlCalculationAutomatic
End With

End Sub

Sub calculaties_sorteren()
Dim N As Integer
Dim M As Integer
Dim FirstWSToSort As Integer
Dim LastWSToSort As Integer
Dim SortDescending As Boolean
te_selecteren_blad = ""

If InStr(1, ActiveSheet.name, "calculatie_", vbTextCompare) = 0 Then
te_selecteren_blad = ActiveSheet.name
Sheets("calculatie_1").Select
End If

    For Each sh In ThisWorkbook.Worksheets
        If InStr(1, sh.name, "calculatie_", vbTextCompare) > 0 And Len(sh.name) > Len("calculatie_") And sh.Visible = 1 Then sh.Select False
    Next sh

If ActiveWindow.SelectedSheets.Count = 1 Then
    FirstWSToSort = 1
    LastWSToSort = Worksheets.Count
Else
    With ActiveWindow.SelectedSheets
        FirstWSToSort = .Item(1).Index
        LastWSToSort = .Item(.Count).Index
     End With
End If

For M = FirstWSToSort To LastWSToSort
    For N = M To LastWSToSort
        If SortDescending = True Then
            If UCase(Worksheets(N).name) > UCase(Worksheets(M).name) Then
                Worksheets(N).Move Before:=Worksheets(M)
            End If
        Else
            If UCase(Worksheets(N).name) < UCase(Worksheets(M).name) Then
               Worksheets(N).Move Before:=Worksheets(M)
            End If
        End If
     Next N
Next M

Worksheets("Voorblad").Move Before:=Worksheets("Calculatie_1")
Worksheets("Project BO").Move Before:=Worksheets("Voorblad")

If te_selecteren_blad <> "" Then Sheets(te_selecteren_blad).Select

End Sub

Sub calculatieblad_kopieren()
Call calculatie_nieuw_kopieren(0)
End Sub

Sub calculatieblad_aanmaken()
Call calculatie_nieuw_kopieren(1)
End Sub

