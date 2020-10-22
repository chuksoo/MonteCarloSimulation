Attribute VB_Name = "Module2"
Sub Reservoir_Parameter_data()
'
' Reservoir_Parameter_data Macro
' Data for Reservoir Parameters
'
' Keyboard Shortcut: Ctrl+r

    Sheets("ReservoirEstimation Parameter").Select
    
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Property"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "Area, A (acres)"
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "Height, h (ft)"
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "Porosity, pu"
    With ActiveCell.Characters(Start:=1, Length:=10).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With ActiveCell.Characters(Start:=11, Length:=1).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "FVF, Bo (RB/STB)"
    With ActiveCell.Characters(Start:=1, Length:=6).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With ActiveCell.Characters(Start:=7, Length:=1).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = True
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With ActiveCell.Characters(Start:=8, Length:=9).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "Minimum (x1)"
    With ActiveCell.Characters(Start:=1, Length:=10).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With ActiveCell.Characters(Start:=11, Length:=1).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = True
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With ActiveCell.Characters(Start:=12, Length:=1).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "2500"
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "200"
    Range("D6").Select
    ActiveCell.FormulaR1C1 = "0.15"
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "1.2"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "Most Likely (x2)"
    With ActiveCell.Characters(Start:=1, Length:=14).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With ActiveCell.Characters(Start:=15, Length:=1).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = True
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With ActiveCell.Characters(Start:=16, Length:=1).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "6000"
    Range("E5").Select
    ActiveCell.FormulaR1C1 = "300"
    Range("E6").Select
    ActiveCell.FormulaR1C1 = "0.25"
    Range("E7").Select
    ActiveCell.FormulaR1C1 = "1.3"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "Maximum (x3)"
    With ActiveCell.Characters(Start:=1, Length:=10).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With ActiveCell.Characters(Start:=11, Length:=1).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = True
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With ActiveCell.Characters(Start:=12, Length:=1).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("F4").Select
    ActiveCell.FormulaR1C1 = "9000"
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "500"
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "0.35"
    Range("F7").Select
    ActiveCell.FormulaR1C1 = "1.35"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "Probability Distibuion"
    Range("G4").Select
    ActiveCell.FormulaR1C1 = "Triangular"
    Range("G5").Select
    ActiveCell.FormulaR1C1 = "Triangular"
    Range("G6").Select
    ActiveCell.FormulaR1C1 = "Triangular"
    Range("G7").Select
    ActiveCell.FormulaR1C1 = "Triangular"
    Range("K3:L3").Select
    ActiveCell.FormulaR1C1 = "Percentiles"
    Range("K4").Select
    ActiveCell.FormulaR1C1 = "P10"
    Range("K5").Select
    ActiveCell.FormulaR1C1 = "P50"
    Range("K6").Select
    ActiveCell.FormulaR1C1 = "P90"
    
        
End Sub


