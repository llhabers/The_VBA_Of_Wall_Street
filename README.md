# The_VBA_Of_Wall_Street

Sub testinghw()

' Create a script that will loop through all the stocks for one year and output the following information.

    ' The ticker symbol.
    
    ' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    
    ' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    
    ' The total stock volume of the stock.
    
' You should also have conditional formatting that will highlight positive change in green and negative change in red.

'----------------------------------------------------------------------------------------------------

 ' Set a variable for worksheet
 Dim ws As Worksheet
 
 For Each ws In ThisWorkbook.Worksheets
 
    ' Set all initial variables.
Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim opening_price As Double
Dim closing_price As Double
Dim LastRow As Long
Dim PreviousAmount As Long
    PreviousAmount = 2
Dim total_volume As Double
    total_volume = 0
Dim ticker_table_row As Integer
    ticker_table_row = 2

    ' In case of a overflow
    'On Error Resume Next

'----------------------------------------------------------------------------------------------------
        
        ' Create title names for each worksheet
        
  'set title
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

'----------------------------------------------------------------------------------------------------

' Determine the Last Row
     LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    ' LastRow = ActiveSheet.UsedRange.Rows.Count
    
' Loop through all tickers.
    For i = 2 To LastRow
    
' Check if we are still on the ticker, if we are not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Set ticker
    ticker = ws.Cells(i, 1).Value
      
' Add to the total_volume.
    total_volume = total_volume + ws.Cells(i, 7).Value
      
' Print the ticker in the summary table.
    ws.Range("I" & ticker_table_row).Value = ticker
    
' Print the total_volume to the Summary Table
    ws.Range("L" & ticker_table_row).Value = total_volume
      
' Set opening_price and closing_price. Other opening prices will be determined in the conditional loop.
    opening_price = ws.Range("C" & PreviousAmount)
    closing_price = ws.Range("F" & i)
      
' Now we can calulate the yearly change from the beginning of the
    yearly_change = closing_price - opening_price
    
' Print the yearly change to the Summary Table
    ws.Range("J" & ticker_table_row).Value = yearly_change
      
' Check for the non-divisibilty condition when calculating the percent change
    If opening_price = 0 Then
        percent_change = 0
                
    Else

        percent_change = (yearly_change / opening_price)
                
    End If

' Print the percent change for each ticker in the summary table
    ws.Range("K" & ticker_table_row).Value = percent_change
    ws.Range("K" & ticker_table_row).NumberFormat = "0.00%"
      
' Add one to the summary table row
    ticker_table_row = ticker_table_row + 1
      
' Reset the opening price
    opening_price = ws.Cells(i + 1, 3)
      
' Reset the ticker Total
    total_volume = 0
    
' If the cell immediately following a row is the same brand...
    
    Else
    
' Add to the ticker Total
    total_volume = total_volume + ws.Cells(i, 7).Value
      
    End If
        
    Next i
     
' Add to the percent change
      
' Conditional formatting that will highlight positive change in green and negative change in red
' First find the last row of the ticker table

    lastrow_ticker_table_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
' Color code yearly change
    For i = 2 To lastrow_ticker_table_row
            
    If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 10
            
    Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
    End If
    Next i
    Next ws
        
End Sub
