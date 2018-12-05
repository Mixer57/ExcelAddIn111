Sub Macro1()
'
' Macro1 Macro
'

'
    Range("B10:C10").Select
    ActiveCell.FormulaR1C1 = "5678"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = "5678"
    ActiveCell.FormulaR1C1 = "78"
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "5678"
    Range("C9").Select
    ActiveCell.FormulaR1C1 = "5678"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "5678"
    Range("D6").Select
    ActiveCell.FormulaR1C1 = "78"
    Range("F5").Select
    ActiveWorkbook.Save
End Sub