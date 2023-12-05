Attribute VB_Name = "Module1"
Sub stock_calculations()

For Each ws In Worksheets
    'Variables
    Dim columnALength As Long
    Dim ticker_row_count As Long
    Dim opening_val As Double
    Dim closing_val As Double
    Dim yearly_change As Double
    Dim yc_row_counter As Long
    Dim pc_row_counter As Long
    Dim percent_change As Double
    Dim total_vol As Variant
    Dim vol_row_counter As Long
    Dim columnILength As Long
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxVol As Double
    Dim ticker1 As String
    Dim ticker2 As String
    Dim ticker3 As String
    
    'Code
    'MsgBox ws.Name
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    columnALength = ws.Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox ("Name: " & ws.Name & " Column A Length: " & columnALength)
    
    ' Figure out the TickerNames
    ws.Cells(2, 9).Value = ws.Cells(2, 1).Value
    opening_val = ws.Cells(2, 3).Value
    ticker_row_count = 3
    yc_row_counter = 2
    pc_row_counter = 2
    vol_row_counter = 2
    total_vol = 0
    
    For i = 2 To columnALength
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
             'Add the Value to the Ticker list (We don't want duplicates; so, we're comparing the values of current and next cells in column 1. Add only if they are different)
             ws.Cells(ticker_row_count, 9).Value = ws.Cells(i + 1, 1).Value
             ticker_row_count = ticker_row_count + 1
             
             'Set value of the closing value
             closing_val = ws.Cells(i, 6).Value
             
             
             'Calculate Yearly Change
             yearly_change = closing_val - opening_val
             ws.Cells(yc_row_counter, 10).Value = yearly_change
             If yearly_change < 0 Then
               ws.Range("J" & yc_row_counter).Interior.ColorIndex = 3
             Else
               ws.Range("J" & yc_row_counter).Interior.ColorIndex = 4
             End If
             
             yc_row_counter = yc_row_counter + 1
             
             'Calculate Percent Change
             percent_change = (yearly_change / opening_val)
             ws.Cells(pc_row_counter, 11).Value = percent_change
             pc_row_counter = pc_row_counter + 1
             
             'Calculate the total_vol for the last row in the section
             total_vol = total_vol + ws.Cells(i, 7).Value
             ws.Cells(vol_row_counter, 12).Value = total_vol
             vol_row_counter = vol_row_counter + 1
             
             'Opening value should change everytime the ticker symbol is different
             opening_val = ws.Cells(i + 1, 3).Value
             total_vol = 0
        Else
              'Keep Calculating the total
              total_vol = total_vol + ws.Cells(i, 7).Value
         End If
    Next i
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    maxPercentIncrease = Cells(2, 11).Value
    maxPercentDecrease = Cells(2, 11).Value
    maxVol = Cells(2, 12).Value
    
    ticker1 = " "
    ticker2 = " "
    ticker3 = " "
    
    'Calculate the Greatest % increase, Greatest % Decrease and Greatest Total Volume
     columnILength = ws.Cells(Rows.Count, 9).End(xlUp).Row
     For n = 3 To columnILength
     
       'Greatest % Increase
       If ws.Cells(n, 11).Value > maxPercentIncrease Then
         maxPercentIncrease = ws.Cells(n, 11).Value
         ticker1 = ws.Cells(n, 9).Value
       End If
       
       If ws.Cells(n, 11).Value < maxPercentDecrease Then
          maxPercentDecrease = ws.Cells(n, 11).Value
          ticker2 = ws.Cells(n, 9).Value
       End If
                
       If ws.Cells(n, 12).Value > maxVol Then
          maxVol = ws.Cells(n, 12).Value
          ticker3 = ws.Cells(n, 9).Value
       End If

       
     Next n
     
    ws.Range("P2").Value = ticker1
    ws.Range("Q2").Value = maxPercentIncrease
    ws.Range("P3").Value = ticker2
    ws.Range("Q3").Value = maxPercentDecrease
    ws.Range("P4").Value = ticker3
    ws.Range("Q4").Value = maxVol
Next ws

End Sub
