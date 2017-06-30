Sub Macro2()
'
' Macro2 Macro
'

'
    Dim CDNum As String
    Dim StatusVal As String
    Dim OrderNum As String
    Dim OrderStatus As String
    Dim OpVal As String
    Dim CountyVal As String
    Dim CoordVal As String
    Dim InputVal As String
    Dim HearingContinuedVal As String
    
    
    'remember to paste CD: into cell C2 and the actual number into D2
    
    Range("C2:O15").Select
    
    If IsEmpty(Range("C12").Value) = False Then
        Range("C12").EntireRow.Insert
        Range("C12").EntireRow.Insert
        Range("C12").EntireRow.Insert
        Range("C12").EntireRow.Insert
    
    ElseIf IsEmpty(Range("C13").Value) = False Then
        Range("C13").EntireRow.Insert
        Range("C13").EntireRow.Insert
        Range("C13").EntireRow.Insert
        
    ElseIf IsEmpty(Range("C14").Value) = False Then
        Range("C14").EntireRow.Insert
        Range("C14").EntireRow.Insert
        
    ElseIf IsEmpty(Range("C15").Value) = False Then
        Range("C15").EntireRow.Insert
    
    End If
    
    'This part of the code looks for CD number and operator
    CD = Range("C2:O15").Find("CD:").Select
    ActiveCell.Offset(0, 1).Activate
    CDNum = ActiveCell.Value
    ActiveCell.Offset(0, 4).Activate
    OpVal = ActiveCell.Value
    Sheets("CleanData").[B2].Value = CDNum
    Sheets("CleanData").[F2].Value = OpVal
    
    'This part of the code looks for Status
    Status = Range("C2:O15").Find("Status:").Select
    ActiveCell.Offset(0, 1).Activate
    StatusVal = ActiveCell.Value
    Sheets("CleanData").[C2].Value = StatusVal
    
    'This part of the code looks for Order Number & Order Status if Final
    Order = Range("C2:O15").Find("Order(s):").Select
    ActiveCell.Offset(0, 2).Activate
    OrderNum = ActiveCell.Value
    ActiveCell.Offset(0, 1).Activate
    OrderStatus = ActiveCell.Value
    Sheets("CleanData").[E2].Value = OrderNum
    Sheets("CleanData").[D2].Value = OrderStatus
    
    'This part of the code looks for County
    County = Range("C2:O15").Find("County:").Select
    ActiveCell.Offset(0, 1).Activate
    CountyVal = ActiveCell.Value
    Sheets("CleanData").[G2].Value = CountyVal
    
    'This part of the code looks for Coordinates/Sections
    Coordinates = Range("C2:O15").Find("Section:").Select
    ActiveCell.Offset(0, 1).Activate
    CoordVal = ActiveCell.Value
    Sheets("CleanData").[H2].Value = CoordVal
        
    'This part of the code looks for input date
    InputDate = Range("C2:O15").Find("Input Date:").Select
    ActiveCell.Offset(0, 1).Activate
    InputVal = ActiveCell.Value
    ActiveCell.Offset(5, 0).Activate
    HearingContinuedVal = ActiveCell.Value
    Sheets("CleanData").[L2].Value = InputVal
    Sheets("CleanData").[M2].Value = HearingContinuedVal
    
    Sheets("CleanData").Range("A2").EntireRow.Insert
    
    Range("C2:O15").EntireRow.Delete
    
    
End Sub
