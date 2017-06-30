Sub Macro9()
'
' Macro9 Macro
'
' Keyboard Shortcut: Ctrl+Shift+T
'
    ActiveCell.FormulaR1C1 = Left(ActiveCell.FormulaR1C1, Len(ActiveCell.FormulaR1C1) - 1)
    Selection.Offset(1, 0).Select
End Sub