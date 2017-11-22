Attribute VB_Name = "instellingen_en_hulpfuncties"
Public groepnaam
Public Groeptitel
Public max_groep_rij
Public max_aantal_calc_bladen
Public NewRange As String
Public OldRange As String
Public SaveRange As String
Public ChangeRange As Boolean
Public naam_kolom As String

Sub instellingen_ophalen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    naam_kolom = "B"
    groepnaam = "z"
    Groeptitel = "aa"
    max_groep_rij = 11
    max_aantal_calc_bladen = 4
End Sub

Sub bladen_tonen_verbergen(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Set shp = ActiveSheet.Shapes(Application.Caller)
      worksheet_naam = Application.Caller
      
With Application
.ScreenUpdating = False
.Calculation = xlCalculationManual
.EnableEvents = False
End With

    begin_rij = Range("begin_" & Application.Caller).Row
    If Right(Application.Caller, 1) < 4 Then
    eind_rij = Range("begin_" & Left(Application.Caller, Len(Application.Caller) - 1) & Right(Application.Caller, 1) + 1).Row - 1
    Else
    eind_rij = Range("einde_calculatie").Row - 1
    End If
            If Evaluate("ISREF('" & worksheet_naam & "'!A1)") Then

    zichtbaar = ActiveSheet.Range(shp.ControlFormat.LinkedCell).Value
      
    Sheets(worksheet_naam).Visible = zichtbaar

    
    Rows(begin_rij & ":" & eind_rij).EntireRow.Hidden = Not ActiveSheet.Range(shp.ControlFormat.LinkedCell).Value

With Application
.ScreenUpdating = True
.Calculation = xlCalculationAutomatic
.EnableEvents = True
End With

Else

    Rows(begin_rij & ":" & eind_rij).EntireRow.Hidden = True
    ActiveSheet.Range(shp.ControlFormat.LinkedCell).Value = False
End If
   
End Sub

Function GetWorksheetFromCodeName(CodeName As String)
    Application.Volatile
    Dim WS As Worksheet
    For Each WS In ThisWorkbook.Worksheets
        If StrComp(WS.CodeName, CodeName, vbTextCompare) = 0 Then
            GetWorksheetFromCodeName = WS.Name
            Exit Function
        End If
    Next WS
End Function

Function GetCodenameFromWorksheet(Worksheet As String)
    Application.Volatile
    Dim Cdname
    For Each WS In ThisWorkbook.Worksheets
        If StrComp(Worksheet, WS.Name, vbTextCompare) = 0 Then
            GetCodenameFromWorksheet = WS.CodeName
            Exit Function
        End If
    Next WS
End Function

Sub testen_van_code(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim WS As Worksheet
    Set WS = GetWorksheetFromCodeName("calculatie_1")
    Debug.Print WS.Name
End Sub

