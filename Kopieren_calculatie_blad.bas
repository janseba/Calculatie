Attribute VB_Name = "Kopieren_calculatie_blad"
Option Explicit

Sub BLADEN_KOPIEREN(Optional variabele_zodat_de_macro_niet_zichtbaar_is As String)
    Dim nm As name
    Dim bron As String
    Dim x
    Dim doel
    Dim sh
    Dim sh_doel
    Dim naam
    Dim nieuwe_naam
    Dim verw
    Dim cell
    
    bron = "calculatie_2"

    For x = 3 To 10
        doel = "calculatie_" & x

        For Each sh In ThisWorkbook.Sheets
            If sh.name = doel Then
                Application.DisplayAlerts = False
                sh.Delete
                Application.DisplayAlerts = True
            End If
        Next sh

        Sheets(bron).Copy after:=Sheets(Sheets(bron).Index + x - 2)
        Set sh_doel = ActiveSheet
        sh_doel.name = doel
        For Each nm In ActiveWorkbook.Names
            If InStr(1, LCase(nm.name), "calculatie", vbTextCompare) Then
                If InStr(1, LCase(nm.RefersTo), doel, vbTextCompare) Then
                    naam = nm.name
                    If Right(naam, Len(bron)) = bron Then
                        nieuwe_naam = Left(naam, Len(naam) - Len(bron))
                        nieuwe_naam = Replace(nieuwe_naam, doel & "!", "")
                        nieuwe_naam = nieuwe_naam & doel
                        verw = Replace(nm.RefersTo, doel & "!", "")
                        Set cell = Worksheets(doel).Range(verw)
                        ThisWorkbook.Names.Add name:=nieuwe_naam, RefersTo:=cell
                        nm.Delete
                    End If
                End If
            End If
        Next nm
    Next x
End Sub
