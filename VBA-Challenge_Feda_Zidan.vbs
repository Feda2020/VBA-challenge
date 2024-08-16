Sub stockdatasummary()
    Dim ws As Worksheet
    For Each ws In Worksheets
    'Variable for holding the ticker name
    Dim ticker_name As String

    'Variable to hold the total per ticker
    Dim total_volume As Double
    total_volume = 0

    'Ticker summery table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
    'Write the summery headers
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
  
    'To get to the end row
    EndRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    'For loop to go through all records
       For i = 2 To EndRow
        'If previous ticker and current ticker are different
        If ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then
        opening_price = ws.Cells(i, 3)
        
        'If next ticker and current ticker are different
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        ticker_name = ws.Cells(i, 1).Value

        'Add to the total volume
        total_volume = total_volume + ws.Cells(i, 7).Value
      
        'The closing price
        closing_price = ws.Cells(i, 6).Value
      
        'Calculate the difference between opening and closing price
        Yearly_Change = closing_price - opening_price
      
        'Calculate the percentage difference between opening and closing price
       
        percent_change = ((closing_price - opening_price) / opening_price)
        'To avoid division by zero
        On Error Resume Next

        ' Print the the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = ticker_name
      
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      
        ws.Range("K" & Summary_Table_Row).Value = percent_change
        'Print in percentage format
        ws.Columns("K:K").NumberFormat = "0.00%"

        ws.Range("L" & Summary_Table_Row).Value = total_volume
      
        'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
        'Reset the total volume
        total_volume = 0

        Else
        'Add to the ticker total
        total_volume = total_volume + ws.Cells(i, 7).Value

        End If
             
       Next i
        
        'The grand summary loop
       Dim greatest_increase, greatest_decrease As Double
       greatest_increase = ws.Cells(2, 11)
       greatest_decrease = ws.Cells(2, 11)
       greatest_volume = ws.Cells(2, 12)
       EndRow_summary = ws.Cells(Rows.Count, 10).End(xlUp).Row
  
       For j = 2 To EndRow_summary
        'Change the format based on the value green for positive and red for negative
        If ws.Cells(j, 10) >= 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
  
        ElseIf ws.Cells(j, 10) < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
   
        End If
        
        'Loop to get greatest inc. value
        If ws.Cells(j, 11) > greatest_increase Then
        greatest_increase = ws.Cells(j, 11)
        ws.Cells(2, 17) = greatest_increase
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 16) = ws.Cells(j, 9)
   
        End If
   
        'Loop to get greatest dec. value
        If ws.Cells(j, 11) < greatest_decrease Then
        greatest_decrease = ws.Cells(j, 11)
        ws.Cells(3, 17) = greatest_decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16) = ws.Cells(j, 9)
  
        End If
        
        If ws.Cells(j, 12) > greatest_volume Then
        greatest_volume = ws.Cells(j, 12)
        ws.Cells(4, 17) = greatest_volume
        ws.Cells(4, 16) = ws.Cells(j, 9)
   
        End If
   
       Next j
 
 Next

End Sub