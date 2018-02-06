Sub stockmarket_analysis()

'Declare my worksheet
Dim ws As Worksheet

' Keep track of the location for each credit card brand in the summary table
Dim Summary_Table_Row As Integer
  
'create a loop to go through each worksheet
For Each ws In Worksheets

    ' Declare my variables
    Dim ticker As String
    Dim yearly_change As Double
    Dim vol_total As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim percent_change As Double
    Dim current_volume As Double
    Dim rng As Range
    Dim dblMin As Double
    Dim dblMax As Double
    vol_total = 0
    Summary_Table_Row = 2
   
    'create a loop to go to the last row
    Dim LastRow As Double
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Create my headers
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
   
    ' Loop through all credit card purchases
    For i = 2 To LastRow
        
        'get the open_price of each stock
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                open_price = ws.Cells(i, 3).Value
            End If
        'To see if we are still in the same ticker. If not, then next loop.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Getting the stock volume
            vol_total = vol_total + ws.Cells(i, 7).Value
            ' Set the ticker
            ticker = ws.Cells(i, 1).Value

            '  Print the ticker in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = ticker
            
            'Find the close price value
            close_price = ws.Cells(i, 6).value
    
            'Print the yearly change in the Summary Table
            ws.Range("j" & Summary_Table_Row).Value = close_price - open_price
            If open_price = 0 Or close_price = 0 Or close_price - open_price = 0 Then
                ws.Range("k" & Summary_Table_Row).Value = 0
            Else
                ws.Range("k" & Summary_Table_Row).Value = (close_price - open_price) / open_price
            End If
            'formatting % and colors
             ws.Range("k" & Summary_Table_Row).NumberFormat = "0.00%"
            If ws.Range("j" & Summary_Table_Row).Value >= 0 Then
                ws.Range("j" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf ws.Range("j" & Summary_Table_Row).Value < 0 Then
                ws.Range("j" & Summary_Table_Row).Interior.ColorIndex = 3
            End If

            
        '  Print the volume total to the Summary Table
        ws.Range("l" & Summary_Table_Row).Value = vol_total

        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the volume
        vol_total = 0

        ' If the cell immediately following a row is the same ticker...
        Else
        
        ' Add to the volume Total
        vol_total = vol_total + ws.Cells(i, 7).Value

    End If

    Next i
    
    'declare variables for getting the ticker sybol and value
    Dim max_incease As Double
    Dim max_decease As Double
    Dim max_vol As Double
    Dim max_inc_ticker As String
    Dim max_dec_ticker As String
    Dim max_vol_ticker As String
    max_inc = 0
    max_dec = 0
    max_vol = 0

'Get the Greated % increase, Greatest % decrease and Greatest total volume
    For i = 2 To LastRow
        If ws.Cells(i, 11).Value > max_incease Then
            max_incease = ws.Cells(i, 11).Value
            max_inc_ticker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 11).Value < max_decease Then
            max_decease = ws.Cells(i, 11).Value
            max_dec_ticker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 12).Value > max_vol Then
            max_vol = ws.Cells(i, 12).Value
            max_vol_ticker = ws.Cells(i, 9).Value
        End If
    Next i
    
'Print to get values and ticker to show
    ws.Cells(2, 17).Value = max_incease_ticker
    ws.Cells(3, 17).Value = max_dec_ticker
    ws.Cells(4, 17).Value = max_vol_ticker
    ws.Cells(2, 18).Value = max_incease
    ws.Cells(3, 18).Value = max_decease
    ws.Cells(4, 18).Value = max_vol


Next ws
End Sub





