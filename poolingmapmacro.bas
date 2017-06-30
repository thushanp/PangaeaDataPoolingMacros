Sub Macro9()
'
' Macro9 Macro
'
' Keyboard Shortcut: Ctrl+Shift+T
'
    ActiveCell.FormulaR1C1 = Left(ActiveCell.FormulaR1C1, Len(ActiveCell.FormulaR1C1) - 1)
    Selection.Offset(1, 0).Select
End Sub
Sub MapGenerator()
'
' MapGenerator Macro
'
' Keyboard Shortcut: Ctrl+Shift+E
'
    Dim SectNum As Integer
    Dim StatusVal As String
    Dim OrderNum As String
    Dim OrderStatus As String
    Dim North As Integer
    Dim West As Integer
    Dim CountyVal As String
    Dim CoordVal As String
    Dim InputVal As String
    Dim HearingContinuedVal As String
    Dim i As Integer
    
    
    Application.ScreenUpdating = False
    
    Range("C2").Select

    'This part of the code looks for CD number and operator
    SectNum = ActiveCell.Value
    ActiveCell.Offset(0, 1).Activate
    North = ActiveCell.Value
    ActiveCell.Offset(0, 1).Activate
    West = ActiveCell.Value
    Sheets("Clean Data").[C2].Value = SectNum
    Sheets("Clean Data").[D2].Value = North
    Sheets("Clean Data").[E2].Value = West
    
    Sheets("Clean Data").Range("A2").EntireRow.Insert
    
    Range("C2").EntireRow.Delete
    
    Worksheets("TrimmedMap").Activate
    
    Range("A1").Select
    ' Highlight the active cell

    i = 19
    
    Do While i > North
        ActiveCell.Offset(6, 0).Activate
        i = i - 1
    Loop
    
    i = 14

    Do While i > West
        ActiveCell.Offset(0, 6).Activate
        i = i - 1
    Loop

    ActiveCell.Offset(0, 6).Activate

    If SectNum < 7 Then
            ActiveCell.Offset(0, -SectNum).Activate
    ElseIf SectNum < 13 Then
            ActiveCell.Offset(1, 0).Activate
            ActiveCell.Offset(0, -6).Activate
            SectNum = SectNum - 7
            ActiveCell.Offset(0, SectNum).Activate
    ElseIf SectNum < 19 Then
            ActiveCell.Offset(1, 0).Activate
            ActiveCell.Offset(1, 0).Activate
            SectNum = SectNum - 13
            ActiveCell.Offset(0, -SectNum - 1).Activate
    ElseIf SectNum < 25 Then
            ActiveCell.Offset(1, 0).Activate
            ActiveCell.Offset(1, 0).Activate
            ActiveCell.Offset(1, 0).Activate
            ActiveCell.Offset(0, -6).Activate
            SectNum = SectNum - 19
            ActiveCell.Offset(0, SectNum).Activate
    ElseIf SectNum < 31 Then
            ActiveCell.Offset(1, 0).Activate
            ActiveCell.Offset(1, 0).Activate
            ActiveCell.Offset(1, 0).Activate
            ActiveCell.Offset(1, 0).Activate
            SectNum = SectNum - 25
            ActiveCell.Offset(0, -SectNum - 1).Activate
    ElseIf SectNum < 37 Then
            ActiveCell.Offset(1, 0).Activate
            ActiveCell.Offset(1, 0).Activate
            ActiveCell.Offset(1, 0).Activate
            ActiveCell.Offset(1, 0).Activate
            ActiveCell.Offset(1, 0).Activate
            ActiveCell.Offset(0, -6).Activate
            SectNum = SectNum - 31
            ActiveCell.Offset(0, SectNum).Activate
    End If

    
    ActiveCell.Interior.ColorIndex = 8
    Worksheets("Data").Activate
    

End Sub
