Sub Cellextractor()
'
' Cellextractor Macro
'
' Keyboard Shortcut: Ctrl+Shift+L
'
Dim pos As Integer
Dim temp As String
Dim Counter As Integer

Range("A99").Select

temp = ActiveCell.Text

For Counter = 1 To Len(MyString)
    'do something to each character in string
    'here we'll msgbox each character

pos = InStr(temp, ",")



ActiveCell.FormulaR1C1 = Left(ActiveCell.FormulaR1C1, Len(ActiveCell.FormulaR1C1) - 1)

End Sub
