Attribute VB_Name = "Module1"
Sub TickerPerform()

' Setup loop to go through all worksheets in the spreadsheet
Dim ws As Worksheet
For Each ws In Worksheets

    ' Create labels for Summary Table's header
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Set initial variable for counting the number of rows needed in Summary Table
    Dim TickerRow As Integer
    TickerRow = 1
    
    'Set initial variable for tracking each ticker's opening amount, starting with the first ticker's
    Dim OpenAmt As Double
    OpenAmt = ws.Cells(2, 3).Value
    
    'Set initial variable for tracking each ticker's cumulative sum
    Dim VolSum As Double
    VolSum = 0
            
    'Count the number of rows in the data set to determine our loop's length
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through all tickers
    For i = 2 To lastrow
    
        'Check if we are in the same ticker, if we are...
        If (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) Then
            
            'Add to the ticker's total volume
            VolSum = VolSum + ws.Cells(i, 7).Value
        
        'If the cell immediately following a row is a new ticker...
        ElseIf (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
        
            'Add one row to the Summary Table
            TickerRow = TickerRow + 1
            
            'Print the ticker's name in the Summary Table
            ws.Cells(TickerRow, 9).Value = ws.Cells(i, 1).Value
            
            'Calculate/print the ticker's Yearly Change in the Summary Table in two-digit decimal format
            ws.Cells(TickerRow, 10).NumberFormat = "0.00"
            ws.Cells(TickerRow, 10).Value = ws.Cells(i, 6).Value - OpenAmt
            
            'Color the Yearly Change in the Summary Table: red for negative numbers, green for positive
            If (ws.Cells(TickerRow, 10).Value < 0) Then
                ws.Cells(TickerRow, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(TickerRow, 10).Interior.ColorIndex = 4
            End If
                  
            'Calculate/print the ticker's Percent Change in the Summary Table in percent format
            ws.Cells(TickerRow, 11).NumberFormat = "0.00%"
            ws.Cells(TickerRow, 11).Value = ws.Cells(TickerRow, 10).Value / OpenAmt
            
            'Calculate/print the ticker's total volume in the Summary Table
            VolSum = VolSum + ws.Cells(i, 7).Value
            ws.Cells(TickerRow, 12).Value = VolSum
          
            'Reset the total volume and opening amount for the next ticker
            VolSum = 0
            OpenAmt = ws.Cells(i + 1, 3).Value
            
        End If
           
    Next i

    ' Create labels for the Greatest Value table's header and rows
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    'Set initial variables for holding ticker names for the highest/lowest percents and largest volume
    Dim HighPercTicker As String
    Dim LowPercTicker As String
    Dim BigVolTicker As String
    
    'Set initial variable for tracking the highest percent ticker value
    Dim HighPerc As Double
    HighPerc = 0
    
    'Set initial variable for tracking the lowest percent ticker value
    Dim LowPerc As Double
    LowPerc = 0
    
    'Set initial variable for tracking the largest volume ticker value
    Dim BigVol As Double
    BigVol = 0
                   
    'Count the number of tickers in the Summary Table to determine our loop's length
    TickerCount = ws.Cells(Rows.Count, 9).End(xlUp).Row
      
    'Loop through all tickers in the Summary Table
    For j = 2 To TickerCount
    
        'Check if the ticker's percent is greater than the High Percent variable's value, if it is...
        If (ws.Cells(j, 11).Value > HighPerc) Then
        
            'Replace the ticker's name and percent values in the High Percent label and value variables
            HighPercTicker = ws.Cells(j, 9).Value
            HighPerc = ws.Cells(j, 11).Value
            
        End If
        
        'Check if the ticker's percent is lower than the Low Percent variable's value, if it is...
        If (ws.Cells(j, 11).Value < LowPerc) Then
        
            'Replace the ticker's name and percent values in the Low Percent label and value variables
            LowPercTicker = ws.Cells(j, 9).Value
            LowPerc = ws.Cells(j, 11).Value
                       
        End If
        
        'Check if the ticker's volume is greater than the High Volume variable's value, if it is...
        If (ws.Cells(j, 12).Value > BigVol) Then
        
            BigVolTicker = ws.Cells(j, 9).Value
            BigVol = ws.Cells(j, 12).Value
        
        End If
            
    Next j
    
    'Print final High Percent ticker name and percent to the Greatest table, the number formatted as percent
    ws.Cells(2, 16).Value = HighPercTicker
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(2, 17).Value = HighPerc
    
    'Print final Low Percent ticker name and percent to the Greatest table, the number formatted as percent
    ws.Cells(3, 16).Value = LowPercTicker
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).Value = LowPerc
    
    'Print final High Volume ticker name and value to the Greatest table
    ws.Cells(4, 16).Value = BigVolTicker
    ws.Cells(4, 17).Value = BigVol
    
Next ws
        
End Sub

