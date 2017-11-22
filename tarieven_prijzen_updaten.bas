Attribute VB_Name = "tarieven_prijzen_updaten"
Sub toon_versie_datums(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)

    MsgBox "Datum laatste update prijzen: 20-10-2016" & vbCrLf & "Datum laatste update tarieven: 21-10-2016", vbInformation, "Laatste updates"

End Sub

Sub versie_controle(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)

    Dim myFile As String
    Dim text As String
    Dim textline As String
    Dim posLat As Integer
    Dim posLong As Integer
    
    myFile = "\\10.58.120.3\dump\versie.txt"

    Open myFile For Input As #1

    Do Until EOF(1)
        Line Input #1, textline
        text = text & textline
        For Each Line In Split(textline, Chr(10))
            If InStr(1, Line, "tariefversie", vbTextCompare) Then
                tariefversie = Split(Line, ":")(1) * 1
            End If
            If InStr(1, Line, "artikelprijsversie", vbTextCompare) Then
                artikelprijsversie = Split(Line, ":")(1) * 1
            End If
        Next Line
    Loop
    Close #1

    If tariefversie > Range("tariefversie") Then
        melding = "nieuwe tarieven aanwezig"
    End If

    If artikelprijsversie > Range("artikelprijsversie") Then
        melding = IIf(melding = "", vbCrLf, "") & "nieuwe artikelprijzen aanwezig"
    End If

    If melding <> "" Then
        MsgBox melding, vbInformation
    End If

End Sub
