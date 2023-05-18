Sub Macro6()
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp

    Rows("1:1").Select
    Selection.AutoFilter

    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:K").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:M").Select
    Selection.Delete Shift:=xlToLeft
    Columns("N:AD").Select
    Selection.Delete Shift:=xlToLeft

    Columns("I:M").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight

    Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft

    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("K:K").Select
    Selection.Cut
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight

    Columns("B:B").Select
    Selection.NumberFormat = "m/d/yyyy"

    Columns("B:B").Select
    Selection.Replace What:=" de", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    Columns("B:B").Select
    Selection.Replace What:=" hs.", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    Columns("K:K").Select
    Selection.Cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight

    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("L2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-2]"
    Range("L2").Select
    Selection.AutoFill Destination:=Range("L2:L358"), Type:=xlFillDefault
    Range("L2:L358").Select
    Columns("K:L").Select
    Range("L1").Activate
    Selection.Style = "Currency"
    Selection.NumberFormat = "_-$ * #,##0.0_-;-$ * #,##0.0_-;_-$ * ""-""??_-;_-@_-"
    Selection.NumberFormat = "_-$ * #,##0_-;-$ * #,##0_-;_-$ * ""-""??_-;_-@_-"

    Dim i as Integer
    i = 2
    Do While Range("A" & i).Value <> ""
        If Range("G" & i).Value = "Bogotá D.C." Then
            Range("F" & i).Value = "Bogotá D.C."     
        End If
        i = i + 1 
    Loop

    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft

    i = 2
    Do While Range("A" & i).Value <> ""
        If (Range("G" & i).Value <> "Entregado") Then
            Range("G" & i).EntireRow.Delete
        Else
            i = i + 1
        End If
    Loop

    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft

End Sub
