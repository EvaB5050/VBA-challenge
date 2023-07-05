Attribute VB_Name = "Module1"
Sub tickerloop_challenge()
    'Setting up the range
        'Set a variable for holding the ticker name
         Dim tickername As String
        
    '***Set a variable to work through each Worksheet***
        Dim ws As Worksheet
        
    '***Loop through each Worksheet
    For Each ws In Worksheets
    
        'Set a variable for holding a total count on the volume of trade
         Dim tickervolume As Double
            tickervolume = 0
         
        'Keep track of location for each ticker name in summary table
         Dim summary_ticker_row As Integer
            summary_ticker_row = 2
        
        'Set initial open_price
         Dim open_price As Double
            open_price = ws.Cells(2, 3).Value
         
         Dim close_price As Double
         Dim yearly_change As Double
         Dim percent_change As Double
         
    '***Set variable and count for Greatest Percent Increase,
    '***Greatest Percent Decrease and Greatest Total Volume
         Dim Greatest_Percent_Increase As Double
         Dim Greatest_Percent_Decrease As Double
         Dim Greatest_Total_Volume As Double
         
         Greatest_Percent_Increase = 0
         Greatest_Percent_Decrease = 999999
         Greatest_Total_Volume = 0
         
  
                     
        'Label Summary Table Headers
         ws.Cells(1, 9).Value = "Ticker"
         ws.Cells(1, 10).Value = "Yearly Change"
         ws.Cells(1, 11).Value = "Percent Change"
         ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Count number of rows in first column
         Dim lastrow As Long
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
        'Loop through the rows by ticker name
         For i = 2 To lastrow
         
        'When value of next cell differs from current cell
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         
        'Set the ticker name
         tickername = ws.Cells(i, 1).Value
         
        'Add the volume of trade
        tickervolume = tickervolume + ws.Cells(i, 7).Value
        
        'Print the ticker name in the summary table
         ws.Range("I" & summary_ticker_row).Value = tickername
         
        'Print the trade volume for each ticker in the summary table
         ws.Range("L" & summary_ticker_row).Value = tickervolume
         
        'Collect info about closing price
         close_price = ws.Cells(i, 6).Value
         
        'Calculate yearly change
         yearly_change = (close_price - open_price)
         
        'Print the yearly change for each ticker in the summary table
         ws.Range("J" & summary_ticker_row).Value = yearly_change
         
        'Check percent change
         If (open_price = 0) Then
            percent_change = 0
            
            Else
                percent_change = yearly_change / open_price
         
         End If
         
    '***Calculate Greatest % Increase***
         If percent_change > Greatest_Percent_Increase Then
         Greatest_Percent_Increase = percent_change
            greatest_ticker = tickername
         End If
         
    '***Print Greatest % Increase***
         ws.Range("O2").Value = "Greatest % Increase"
         ws.Range("Q2").Value = Greatest_Percent_Increase
         ws.Range("Q2").NumberFormat = "0.00%"
         ws.Range("P2").Value = greatest_ticker
         
    '***Calculate Greatest % Decrease***
         If percent_change < Greatest_Percent_Decrease Then
         Greatest_Percent_Decrease = percent_change
            lowest_ticker = tickername
         End If
                        
    '***Print Greatest % Decrease***
         ws.Range("O3").Value = "Greatest % Decrease"
         ws.Range("Q3").Value = Greatest_Percent_Decrease
         ws.Range("Q3").NumberFormat = "0.00%"
         ws.Range("P3").Value = lowest_ticker
            
    '***Calculate Greatest Total Volume***
         If tickervolume > Greatest_Total_Volume Then
         Greatest_Total_Volume = tickervolume
            volume_ticker = tickername
         End If
         
    '***Print Greatest Total Volume***
         ws.Range("O4").Value = "Greatest Total Volume"
         ws.Range("Q4").Value = Greatest_Total_Volume
         ws.Range("P4").Value = volume_ticker
         
         
             
        'Print the yearly change for each ticker in the summary table
         ws.Range("K" & summary_ticker_row).Value = percent_change
         ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
         
        'Reset the row counter, add to summary_ticker_row
         summary_ticker_row = summary_ticker_row + 1
         
        'Reset volume of trade to zero
         tickervolume = 0
         
        'Reset opening price
         open_price = ws.Cells(i + 1, 3)
         
            Else
                tickervolume = tickervolume + ws.Cells(i, 7).Value
        
        End If
        
    Next i
    
    'Setting up conditional formatting to highlight positive change in Green
    'and negative change in Red
     lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
    'colour code Yearly Change
     For i = 2 To lastrow_summary_table
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
        
     'Format Yearly Change to 2 decimal points
      ws.Cells(i, 10).NumberFormat = "0.00"
        
    Next i
    
   Next
        
End Sub


