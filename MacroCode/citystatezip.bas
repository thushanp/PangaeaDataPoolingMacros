Sub Statezip()
'
' Statezip Macro
'
' Keyboard Shortcut: Ctrl+d
'

Dim State As String
Dim Zip As String


    State = Left(ActiveCell.FormulaR1C1, 2)
    Zip = Right(ActiveCell.FormulaR1C1, 5)
    
    ActiveCell.Offset(0, 1).Activate
    ActiveCell.Value = Zip
    ActiveCell.Offset(0, -1).Activate
    ActiveCell.Value = State
    
    
End Sub
Sub CityStateZip()
'
' CityStateZip Macro
'
' Keyboard Shortcut: Ctrl+e
'

Dim State As String
Dim Zip As String
Dim City As String
Dim Temp As String



    Zip = Right(ActiveCell.FormulaR1C1, 5)
    City = Left(ActiveCell.FormulaR1C1, Len(ActiveCell.FormulaR1C1) - 10)
    Temp = Left(ActiveCell.FormulaR1C1, Len(ActiveCell.FormulaR1C1) - 6)
    State = Right(Temp, 2)
    
    ActiveCell.Offset(0, 2).Activate
    ActiveCell.Value = Zip
    ActiveCell.Offset(0, -1).Activate
    ActiveCell.Value = State
    ActiveCell.Offset(0, -1).Activate
    ActiveCell.Value = City
    
    
End Sub
