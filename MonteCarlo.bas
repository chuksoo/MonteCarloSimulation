Attribute VB_Name = "Module1"
Sub MonteCarlo()
'
' This is a macro used to input values and prepare MonteCarloSimulation sheet
' MonteCarlo Macro
' Macro for Monte Carlo Simulation
'
' Keyboard Shortcut: Ctrl+m

' Written by Chukwuemeka Okoli
'
    Sheets("MonteCarloSimulation").Select
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Wells"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Rand"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "Area,h"
    Range("B1:C1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("B1:C1").Select
    ActiveCell.FormulaR1C1 = "Area in ft"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "Rand"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = ""
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "Area, A"
    Range("B1:C1").Select
    ActiveCell.FormulaR1C1 = "Area in acres"
    Range("D1:E1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Thickness in ft"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "Thickness, h"
    Range("F1:G1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Porosity (fraction)"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "Rand"
    Range("G2").Select
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
    
    Range("H1:I1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("H1:I1").Select
    ActiveCell.FormulaR1C1 = "Formation Volume factor, RB/STB"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "Rand"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "FVF, Bo"
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
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "Swi, "
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Water Saturation"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "Swi (fraction)"
    Range("K6").Select
    Columns("J:J").ColumnWidth = 15.86
    Range("J1").Select
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Stock Tank Barrels, STB"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "Nr, STB"
    ActiveCell.FormulaR1C1 = "Stock Tank Barrels, STB"
    Range("L2").Select
    
    With ActiveCell.Characters(Start:=1, Length:=1).Font
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
    With ActiveCell.Characters(Start:=2, Length:=1).Font
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
    With ActiveCell.Characters(Start:=3, Length:=5).Font
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
   
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("H5").Select
    Columns("E:E").ColumnWidth = 11.86
    Columns("G:G").ColumnWidth = 12.57
    Columns("I:I").ColumnWidth = 15.71
    Columns("I:I").ColumnWidth = 20.43
    Range("A1:A2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("A1:K2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1:K2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A3:K3").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("J10").Select
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("I9").Select
End Sub


Sub MonteCarloSimulation()

' MonteCarloSimulation Macro
' Macros to run simulation for various run
'
' Keyboard Shortcut: Ctrl+e
' Written by Okoli Chukwuemeka
'


Dim i As Integer
Dim myValue As Variant
Dim Count As Integer


'This code allows user to select integer value for the simulation run
    Sheets("MonteCarloSimulation").Select
    myValue = InputBox("How many number of iterations do you want to run?", "Monte Carlo Simulation")
    Range("A3:CA100000").Select
    Selection.ClearContents
    
    
        If myValue <= 0 Then
        
                myValue = MsgBox("Parameter cannot be negative. Enter a positive number. Do you wish to continue?", vbYesNo + vbQuestion, "Monte Carlo Simulation")
                MsgBox ("Iteration will be aborted")
        Exit Sub
                   Else
                Range("A3:CA100000").Select
                Selection.ClearContents
        
        End If
       
       
'This code will run iteration based on input from the user
For i = 1 To myValue

    
    Cells(i + 2, 1).Value = i
    Cells(i + 2, 2).FormulaR1C1 = "=RAND()"
    Cells(i + 2, 3).FormulaR1C1 = "=IF(RC[-1]<=(('ReservoirEstimation Parameter'!R4C5-'ReservoirEstimation Parameter'!R4C4)/('ReservoirEstimation Parameter'!R4C6-'ReservoirEstimation Parameter'!R4C4)), ('ReservoirEstimation Parameter'!R4C4+SQRT((MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R4C6-'ReservoirEstimation Parameter'!R4C4)*('ReservoirEstimation Parameter'!R4C5-'ReservoirEsti" & _
        "mation Parameter'!R4C4))),  ('ReservoirEstimation Parameter'!R4C6-SQRT((1-MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R4C6-'ReservoirEstimation Parameter'!R4C4)*('ReservoirEstimation Parameter'!R4C6-'ReservoirEstimation Parameter'!R4C5))))" & _
        ""
    Cells(i + 2, 4).FormulaR1C1 = "=RAND()"
    Cells(i + 2, 5).FormulaR1C1 = "=IF(RC[-1]<=(('ReservoirEstimation Parameter'!R5C5-'ReservoirEstimation Parameter'!R5C4)/('ReservoirEstimation Parameter'!R5C6-'ReservoirEstimation Parameter'!R5C4)), ('ReservoirEstimation Parameter'!R5C4+SQRT((MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R5C6-'ReservoirEstimation Parameter'!R5C4)*('ReservoirEstimation Parameter'!R5C5-'ReservoirEsti" & _
        "mation Parameter'!R5C4))),  ('ReservoirEstimation Parameter'!R5C6-SQRT((1-MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R5C6-'ReservoirEstimation Parameter'!R5C4)*('ReservoirEstimation Parameter'!R5C6-'ReservoirEstimation Parameter'!R5C5))))" & _
        ""
    Cells(i + 2, 6).FormulaR1C1 = "=RAND()"
    Cells(i + 2, 7).FormulaR1C1 = "=IF(RC[-1]<=(('ReservoirEstimation Parameter'!R6C5-'ReservoirEstimation Parameter'!R6C4)/('ReservoirEstimation Parameter'!R6C6-'ReservoirEstimation Parameter'!R6C4)), ('ReservoirEstimation Parameter'!R6C4+SQRT((MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R6C6-'ReservoirEstimation Parameter'!R6C4)*('ReservoirEstimation Parameter'!R6C5-'ReservoirEsti" & _
        "mation Parameter'!R6C4))),  ('ReservoirEstimation Parameter'!R6C6-SQRT((1-MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R6C6-'ReservoirEstimation Parameter'!R6C4)*('ReservoirEstimation Parameter'!R6C6-'ReservoirEstimation Parameter'!R6C5))))" & _
        ""
    Cells(i + 2, 8).FormulaR1C1 = "=RAND()"
    Cells(i + 2, 9).FormulaR1C1 = "=IF(RC[-1]<=(('ReservoirEstimation Parameter'!R7C5-'ReservoirEstimation Parameter'!R7C4)/('ReservoirEstimation Parameter'!R7C6-'ReservoirEstimation Parameter'!R7C4)), ('ReservoirEstimation Parameter'!R7C4+SQRT((MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R7C6-'ReservoirEstimation Parameter'!R7C4)*('ReservoirEstimation Parameter'!R7C5-'ReservoirEsti" & _
        "mation Parameter'!R7C4))),  ('ReservoirEstimation Parameter'!R7C6-SQRT((1-MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R7C6-'ReservoirEstimation Parameter'!R7C4)*('ReservoirEstimation Parameter'!R7C6-'ReservoirEstimation Parameter'!R7C5))))" & _
        ""
    Cells(i + 2, 10).FormulaR1C1 = "=0.325-(0.5*RC[-3])"
    Cells(i + 2, 11).FormulaR1C1 = "=(7758*RC[-8]*RC[-6]*RC[-4]*(1-RC[-1]))/RC[-2]"
       
    'Code for calculating f(x) for the distribution
    Cells(i + 2, 12).Value = (i) / myValue
    
    'Code for Calculating Uniform distribution for the given data
    Cells(i + 2, 19).FormulaR1C1 = "='ReservoirEstimation Parameter'!R4C4+(RC[-17]*('ReservoirEstimation Parameter'!R4C6-'ReservoirEstimation Parameter'!R4C4))"
    Cells(i + 2, 20).FormulaR1C1 = "='ReservoirEstimation Parameter'!R5C4+(RC[-16]*('ReservoirEstimation Parameter'!R5C6-'ReservoirEstimation Parameter'!R5C4))"
    Cells(i + 2, 21).FormulaR1C1 = "='ReservoirEstimation Parameter'!R6C4+(RC[-15]*('ReservoirEstimation Parameter'!R6C6-'ReservoirEstimation Parameter'!R6C4))"
    Cells(i + 2, 22).FormulaR1C1 = "='ReservoirEstimation Parameter'!R7C4+(RC[-14]*('ReservoirEstimation Parameter'!R7C6-'ReservoirEstimation Parameter'!R7C4))"
    Cells(i + 2, 23).FormulaR1C1 = "=0.325-(0.5*'MonteCarloSimulation'!RC21)"
    Cells(i + 2, 24).FormulaR1C1 = "=(7758*'MonteCarloSimulation'!RC19*'MonteCarloSimulation'!RC20*'MonteCarloSimulation'!RC21*(1-'MonteCarloSimulation'!RC23))/('MonteCarloSimulation'!RC22)"
    
Next i

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'We need to copy, sort and paste the distribution in order to plot the curve for both triangular and uniform distribution
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    'To copy from iterated Triangular distribution and paste only values in new cell for use for plotting the distribution curves
    '
    Range("A3:L100000").Select
    Selection.Copy
    Range("AA3:AL3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AA3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AA3:AL100000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AA3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AA3:AL100000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AB3:AB100000").Select
    Selection.NumberFormat = "0.000"
    Range("AB23").Select
    Range("AB3:AB100000").Select
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AB3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AB3:AB100000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AC3:AC100000").Select
    Selection.NumberFormat = "0.00"
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AC3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AC3:AC100000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AD3:AD100000").Select
    Selection.NumberFormat = "0.000"
    Range("AE3:AE100000").Select
    Selection.NumberFormat = "0.00"
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AE3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AE3:AE100000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AF3:AF100000").Select
    Selection.NumberFormat = "0.000"
    Range("AG3:AG100000").Select
    Selection.NumberFormat = "0.00"
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AG3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AG3:AG100000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AH3:AH100000").Select
    Selection.NumberFormat = "0.000"
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AH3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AH3:AH100000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AI3:AI100000").Select
    Selection.NumberFormat = "0.00"
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AI3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AI3:AI100000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AJ3:AJ100000").Select
    Selection.NumberFormat = "0.00"
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AJ3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AJ3:AJ100000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AK3:AK100000").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AK3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AK3:AK100000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AL3:AL100000").Select
    Selection.NumberFormat = "0.00"
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AL3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AL3:AL100000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'The code for generating the Uniform distribution values
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Range("A3").Select
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "=RAND()"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]<=(('ReservoirEstimation Parameter'!R[1]C[2]-'ReservoirEstimation Parameter'!R[1]C[1])/('ReservoirEstimation Parameter'!R[1]C[3]-'ReservoirEstimation Parameter'!R[1]C[1])), ('ReservoirEstimation Parameter'!R[1]C[1]+SQRT((MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R[1]C[3]-'ReservoirEstimation Parameter'!R[1]C[1])*('ReservoirEstimation Pa" & _
        "rameter'!R[1]C[2]-'ReservoirEstimation Parameter'!R[1]C[1]))),  ('ReservoirEstimation Parameter'!R[1]C[3]-SQRT((1-MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R[1]C[3]-'ReservoirEstimation Parameter'!R[1]C[1])*('ReservoirEstimation Parameter'!R[1]C[3]-'ReservoirEstimation Parameter'!R[1]C[2]))))" & _
        ""
    Range("C3").Select
    
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "=RAND()"
    Range("D4").Select
    Sheets("MonteCarloSimulation").Select
    Range("C3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]<=(('ReservoirEstimation Parameter'!R4C5-'ReservoirEstimation Parameter'!R4C4)/('ReservoirEstimation Parameter'!R4C6-'ReservoirEstimation Parameter'!R4C4)), ('ReservoirEstimation Parameter'!R4C4+SQRT((MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R4C6-'ReservoirEstimation Parameter'!R4C4)*('ReservoirEstimation Parameter'!R4C5-'ReservoirEsti" & _
        "mation Parameter'!R4C4))),  ('ReservoirEstimation Parameter'!R4C6-SQRT((1-MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R4C6-'ReservoirEstimation Parameter'!R4C4)*('ReservoirEstimation Parameter'!R4C6-'ReservoirEstimation Parameter'!R4C5))))" & _
        ""
    Range("C4").Select
    ActiveWindow.ScrollColumn = 1
    Sheets("MonteCarloSimulation").Select
    Range("C3").Select
    Selection.Copy
    Range("E3").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=-57
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]<=(('ReservoirEstimation Parameter'!R5C5-'ReservoirEstimation Parameter'!R5C4)/('ReservoirEstimation Parameter'!R5C6-'ReservoirEstimation Parameter'!R5C4)), ('ReservoirEstimation Parameter'!R5C4+SQRT((MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R5C6-'ReservoirEstimation Parameter'!R5C4)*('ReservoirEstimation Parameter'!R5C5-'ReservoirEsti" & _
        "mation Parameter'!R5C4))),  ('ReservoirEstimation Parameter'!R5C6-SQRT((1-MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R5C6-'ReservoirEstimation Parameter'!R5C4)*('ReservoirEstimation Parameter'!R5C6-'ReservoirEstimation Parameter'!R5C5))))" & _
        ""
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "=RAND()"
    Range("E3").Select
    Selection.Copy
    Range("G3").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]<=(('ReservoirEstimation Parameter'!R6C5-'ReservoirEstimation Parameter'!R6C4)/('ReservoirEstimation Parameter'!R6C6-'ReservoirEstimation Parameter'!R6C4)), ('ReservoirEstimation Parameter'!R6C4+SQRT((MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R6C6-'ReservoirEstimation Parameter'!R6C4)*('ReservoirEstimation Parameter'!R6C5-'ReservoirEsti" & _
        "mation Parameter'!R6C4))),  ('ReservoirEstimation Parameter'!R6C6-SQRT((1-MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R6C6-'ReservoirEstimation Parameter'!R6C4)*('ReservoirEstimation Parameter'!R6C6-'ReservoirEstimation Parameter'!R6C5))))" & _
        ""
    Range("G4").Select
    Range("H3").Select
    ActiveCell.FormulaR1C1 = "=RAND()"
    Range("G3").Select
    Selection.Copy
    Range("I3").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]<=(('ReservoirEstimation Parameter'!R7C5-'ReservoirEstimation Parameter'!R7C4)/('ReservoirEstimation Parameter'!R7C6-'ReservoirEstimation Parameter'!R7C4)), ('ReservoirEstimation Parameter'!R7C4+SQRT((MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R7C6-'ReservoirEstimation Parameter'!R7C4)*('ReservoirEstimation Parameter'!R7C5-'ReservoirEsti" & _
        "mation Parameter'!R7C4))),  ('ReservoirEstimation Parameter'!R7C6-SQRT((1-MonteCarloSimulation!RC[-1])*('ReservoirEstimation Parameter'!R7C6-'ReservoirEstimation Parameter'!R7C4)*('ReservoirEstimation Parameter'!R7C6-'ReservoirEstimation Parameter'!R7C5))))" & _
        ""
    Range("I4").Select
   
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "=0.325-(0.5*RC[-3])"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "=(7758*RC[-8]*RC[-6]*RC[-4]*(1-RC[-1]))/RC[-2]"
    Range("K4").Select
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Creates a copy of Uniform Distribution
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Range("L1:L2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A3:K3").Select
    Range("K3").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A1:L2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Range("K19").Select
  
    Range("S1:X1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.UnMerge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "Uniform Distribution"
    Range("S2:X2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "Area"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "Thickness, h"
    Range("U2").Select
    ActiveCell.FormulaR1C1 = "Porosity, pu"
    Range("V2").Select
    ActiveCell.FormulaR1C1 = "FVF, Bo"
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
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "Swi (fraction)"
    Range("X2").Select
    ActiveCell.FormulaR1C1 = "Nr, STB"
    Range("X2").Select
    ActiveCell.FormulaR1C1 = "Nr, STB"
    With ActiveCell.Characters(Start:=1, Length:=1).Font
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
    With ActiveCell.Characters(Start:=2, Length:=1).Font
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
    With ActiveCell.Characters(Start:=3, Length:=5).Font
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
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "f(x)"
    Range("S2:X2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    Range("AA1").Select
    ActiveCell.FormulaR1C1 = "Well"
    Range("AB1:AC1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Area in acres"
    Range("AD1:AE1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Thickness in ft"
    Range("AF1:AG1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Porosity (fraction)"
    Range("AH1:AI1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Formation Volume Factor"
    Range("AH2").Select
    Columns("AI:AI").ColumnWidth = 16.14
    Columns("AG:AG").ColumnWidth = 10.71
    Columns("AE:AE").ColumnWidth = 10
    Columns("AE:AE").ColumnWidth = 13.29
    Columns("AJ:AJ").ColumnWidth = 17.43
    Range("AJ1").Select
    ActiveCell.FormulaR1C1 = "Water Saturation"
    Range("AK1").Select
    ActiveCell.FormulaR1C1 = "Stock Tank Barrels, STB"
    Range("AB2").Select
    ActiveCell.FormulaR1C1 = "Rand"
    Range("AC2").Select
    ActiveCell.FormulaR1C1 = "Area, A"
    Range("AD2").Select
    ActiveCell.FormulaR1C1 = "Rand"
    Range("AE2").Select
    ActiveCell.FormulaR1C1 = "Thickness, h "
    Range("AD31").Select

    Range("S1:X2").Select
    Selection.Copy
    Range("AO1").Select
    ActiveSheet.Paste
    Columns("AO:AO").ColumnWidth = 11.57
    Columns("AP:AP").ColumnWidth = 12.43
    Columns("AQ:AQ").ColumnWidth = 13.43
    Columns("AR:AR").ColumnWidth = 13.71
    Columns("AS:AS").ColumnWidth = 20.86
    Columns("AT:AT").ColumnWidth = 12.57
    Range("AO1:AT1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Sorted Uniform Distribution"
    Range("AN14").Select
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'To copy from iterated Uniform distribution and paste only values in new cell for use for plotting the uniform distribution curves
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Range("S3:X100000").Select
    Selection.Copy
    Range("AO3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AO3:AO100000").Select
    Application.CutCopyMode = False
    Selection.NumberFormat = "0.00"
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AO3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AO3:AO100000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AP3:AP100000").Select
    Selection.NumberFormat = "0.00"
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AP3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AP3:AP100000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AQ3:AQ100000").Select
    Selection.NumberFormat = "0.000"
    Application.WindowState = xlNormal
    Selection.NumberFormat = "0.00"
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AQ3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AQ3:AQ100000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AR3:AR100000").Select
    Selection.NumberFormat = "0.00"
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AR3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AR3:AR100000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AS3:AS100000").Select
    Selection.NumberFormat = "0.00"
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AS3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AS3:AS100000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AT3:AT100000").Select
    Selection.NumberFormat = "#,##0.00"
    Columns("AT:AT").ColumnWidth = 19.71
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort.SortFields.Add Key:= _
        Range("AT3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MonteCarloSimulation").Sort
        .SetRange Range("AT3:AT100000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    Range("P8").Select
    End With
    
        
End Sub


