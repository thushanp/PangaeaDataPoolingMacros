Sub sectioner()
'
' sectioner Macro
'
' Keyboard Shortcut: Ctrl+g
'
    Dim APINum As String
    Dim DateVal As String
    Dim SectVal As String
    Dim CoordVal As String
    
    Dim CDNum As String
    Dim StatusVal As String
    Dim OrderNum As String
    Dim OrderStatus As String
    Dim OpVal As String
    Dim CountyVal As String
    Dim InputVal As String
    Dim HearingContinuedVal As String
    
    Dim cl As Object
    
    'remember to paste API: into cell C2 and the actual number into D2
    
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

    'This part of the code looks for API number, Spud Date and Section Number
    
    With Worksheets("Data").Cells
        Set cl = Range("C2:O15").Find("Spud Date:")
        If Not cl Is Nothing Then
            cl.Select
            ActiveCell.Offset(0, 1).Activate
            DateVal = ActiveCell.Value
            Sheets("CleanData").[C2].Value = DateVal
            
            Range("C2:O15").Find("API:").Select
            ActiveCell.Offset(0, 1).Activate
            APINum = ActiveCell.Value
            Sheets("CleanData").[B2].Value = APINum
    
            Range("C2:O15").Find("Section:").Select
            ActiveCell.Offset(0, 1).Activate
            SectVal = ActiveCell.Value
            Sheets("CleanData").[A2].Value = SectVal
            
            Sheets("CleanData").Range("A2").EntireRow.Insert
        End If
    End With
    
    Range("C2:O15").EntireRow.Delete
    
    
End Sub
