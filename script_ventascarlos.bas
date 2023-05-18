Sub Macro4()
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Columns("K:K").Select
    Selection.Copy
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight

    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft

    Columns("L:M").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight

    Columns("P:P").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight

    Columns("O:R").Select
    Selection.Delete Shift:=xlToLeft

    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft

    Columns("K:M").Select
    Selection.Delete Shift:=xlToLeft
End Sub