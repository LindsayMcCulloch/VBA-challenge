Sub Module2Challenge()

'DataRow variable is indicating the row for our loops
Dim DataRow As Long
'AnnualChange variable is indicating the yearly change between ticker open and close stocks
Dim AnnualChange As Double
'PercentChange variable is indicating the percentage change between ticker open and stocks
Dim PercentChange As Double
'TotalVolume variable is indicating the totalstock volume per ticker
Dim TotalVolume As Double

Dim a As Double
Dim b As Double
Dim c As Double

Dim LastRow As Long
Dim IncreaseTicker As String
Dim DecreaseTicker As String
Dim TotalTicker As String
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestTotal As Double
   
   'To loop through each worksheet
   For Each ws In Worksheets
   'to make sure active worksheet starts the loop
    ws.Activate
  'creating summary columns and formatting headings
   Cells(1, 9).Value = "Ticker"
   Cells(1, 9).Font.Bold = True
   Cells(1, 10).Value = "Annual Change"
   Cells(1, 10).Font.Bold = True
   Cells(1, 11).Value = "Percentage Change"
   Cells(1, 11).Font.Bold = True
   Cells(1, 12).Value = "Total Volume"
   Cells(1, 12).Font.Bold = True
   Cells(1, 14).Value = "BONUS"
   Cells(1, 14).Font.Bold = True
   Cells(1, 14).Interior.ColorIndex = 6
   Cells(2, 14).Value = "Ticker"
   Cells(2, 14).Font.Bold = True
   Cells(3, 14).Value = "Greatest % Increase"
   Cells(3, 14).Font.Bold = True
   Cells(4, 14).Value = "Greatest % Decrease"
   Cells(4, 14).Font.Bold = True
   Cells(5, 14).Value = "Greatest Total Volume"
   Cells(5, 14).Font.Bold = True
   Cells(2, 15).Value = "Value"
   Cells(2, 15).Font.Bold = True
   
   'starting values for each variable
   a = 2
   b = 0
   AnnualChange = 0
   TotalVolume = 0
   'to calculate last row of data
   LastRow = Cells(Rows.Count, 1).End(xlUp).Row
   For DataRow = 2 To LastRow
   
   'comparing cell value (ticker) with all other cells
    If Cells(DataRow + 1, 1).Value <> Cells(DataRow, 1).Value Then
           TotalVolume = TotalVolume + Cells(DataRow, 7).Value
          
           If TotalVolume = 0 Then
               
               
               Cells(2, 9).Value = Cells(DataRow, 1).Value
               Cells(2, 10).Value = 0
               Cells(2, 11).Value = "%" & 0
               Cells(2, 12).Value = 0
             
            Else
              'Find non zero starting value
              If Cells(a, 3) = 0 Then
               For Value = a To DataRow
                       If Cells(Value, 3).Value <> 0 Then
                           a = Value
                           Exit For
                       End If
                Next Value
              End If
               
            
            AnnualChange = (Cells(DataRow, 6) - Cells(a, 3))
            PercentChange = AnnualChange / Cells(a, 3)
            'formula to repeat the loop using next cell value (ticker)
            a = DataRow + 1
            Range("I" & 2 + b).Value = Cells(DataRow, 1).Value
            Range("J" & 2 + b).Value = AnnualChange
            Range("J" & 2 + b).NumberFormat = "0.00"
            Range("K" & 2 + b).Value = PercentChange
            Range("K" & 2 + b).NumberFormat = "0.00%"
            Range("L" & 2 + b).Value = TotalVolume
            
                'Formatting
                If (AnnualChange > 0) Then
                Range("J" & 2 + b).Interior.ColorIndex = 4
                
                ElseIf (AnnualChange <= 0) Then
                Range("J" & 2 + b).Interior.ColorIndex = 3
                
                End If
                
                If (PercentChange > 0) Then
                Range("K" & 2 + b).Interior.ColorIndex = 4
                
                ElseIf (PercentChange <= 0) Then
                Range("K" & 2 + b).Interior.ColorIndex = 3
                
                End If
               
               
            End If
                         
           'New starting values for next loop
           TotalVolume = 0
           AnnualChange = 0
           b = b + 1
           c = 0
      
       Else
           TotalVolume = TotalVolume + Cells(DataRow, 7).Value
    End If
    
          
    Next DataRow
    'starting values
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestTotal = 0
    'place actual value in WS cells
    IncreaseTicker = ""
    DecreaseTicker = ""
    TotalTicker = ""
    
    For DataRow = 2 To LastRow
    
        If Cells(DataRow, 11).Value > GreatestIncrease Then
            GreatestIncrease = Cells(DataRow, 11).Value
            IncreaseTicker = Cells(DataRow, 9).Value
        End If
        
        If Cells(DataRow, 11).Value < GreatestDecrease Then
            GreatestDecrease = Cells(DataRow, 11).Value
            DecreaseTicker = Cells(DataRow, 9).Value
            
        End If
        
        If Cells(DataRow, 12).Value > GreatestTotal Then
            GreatestTotal = Cells(DataRow, 12).Value
            TotalTicker = Cells(DataRow, 9).Value
            
        End If
         
    Next
    'values for BONUS
    Cells(3, 16).Value = IncreaseTicker
    Cells(4, 16).Value = DecreaseTicker
    Cells(5, 16).Value = TotalTicker
    Cells(3, 17).Value = GreatestIncrease
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 17).Value = GreatestDecrease
    Cells(4, 17).NumberFormat = "0.00%"
    Cells(5, 17).Value = GreatestTotal
    Cells(5, 17).NumberFormat = "#"

   
    Next ws

End Sub
