Attribute VB_Name = "Module3"
Sub Generate_Plot_percentiles()
'
' Generate_Plot_percentiles Macro
' This macro generates plots for each distribution and the percentiles for the given data
'
' Keyboard Shortcut: Ctrl+p
'
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Generating Percentiles for the distribution
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Range("L4").Select
    ActiveCell.FormulaR1C1 = _
        "=PERCENTILE(MonteCarloSimulation!R3C11:R100000C11, 0.1)"
    Range("L5").Select
    ActiveCell.FormulaR1C1 = _
        "=PERCENTILE(MonteCarloSimulation!R3C11:R100000C11, 0.5)"
    Range("L6").Select
    ActiveCell.FormulaR1C1 = _
        "=PERCENTILE(MonteCarloSimulation!R3C11:R100000C11, 0.9)"
    Range("L4:L6").Select
    Selection.NumberFormat = "#,##0.00"


 
End Sub


