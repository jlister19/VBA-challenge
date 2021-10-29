Attribute VB_Name = "Every_Wksht"
Sub Every_WkshtInTest()

Application.ScreenUpdating = False

'Loop through all sheets
For Each ws In Worksheets

ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Pct Change"
ws.Cells(1, 12) = "Total Stock Volume"

Dim ticker As String

'create a running total for each ticker's volume
Dim vol_total As Variant

'set the initial vol_total at 0; we will reset it at end of loop before next i
vol_total = 0

'set the inital BegYrOpenPrice; we'll resent it as we go in the loop
Dim BegYrOpenPrice As Double
BegYrOpenPrice = ws.Cells(2, 3).Value

Dim EndYrClosePrice As Double

Dim summary_table_row As Integer
summary_table_row = 2

'define the last row of the entire data range
LastRowData = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRowData

    'if the value in the row after this one is diff, assume this is last row for this ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'set the ticker variable to the value of the current row's ticker
        ticker = ws.Cells(i, 1).Value
        
        'add this row's vol to the running total vol variable from the summary table
        vol_total = CDec(vol_total) + ws.Cells(i, 7).Value
        
        'get the year end price for this ticker
        EndYrClosePrice = ws.Cells(i, 6).Value
                
        'add ticker to the summary table
        ws.Cells(summary_table_row, 9).Value = ticker
                
        'populate the yearly change value
        ws.Cells(summary_table_row, 10).Value = EndYrClosePrice - BegYrOpenPrice
        
        'apply green and red color based on price change direction
        If ws.Cells(summary_table_row, 10).Value > 0 Then
            ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(summary_table_row, 10).Value < 0 Then
            ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
        End If
                
        'populate the yearly Pct change value, the added if stmt accounts for dividing by 0 error in the test sheet letter P
        'If ws.Cells(summary_table_row, 10).Value <> 0 Then
        If BegYrOpenPrice <> 0 Then
        ws.Cells(summary_table_row, 11).Value = ((EndYrClosePrice / BegYrOpenPrice) - 1)
        Else: ws.Cells(summary_table_row, 11).Value = 0
        End If
            
        ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
                  
        'set the last row in the summary table to the new vol_total
        ws.Cells(summary_table_row, 12).Value = vol_total
        
        'go one row down on the summary table
        summary_table_row = summary_table_row + 1
        
        'reset the vol_total because that was the last row
        vol_total = 0
    
        'and reset the beg of yr price for the next ticker
        BegYrOpenPrice = ws.Cells(i + 1, 3).Value
    
    Else
        
        'otherwise just add the current row's volume to the running running total
        vol_total = CDec(vol_total) + ws.Cells(i, 7).Value
            
    End If
        
Next i

ws.Cells(2, 15) = "Greatest % Increase" 'O2
ws.Cells(3, 15) = "Greatest % Decrease" 'O3
ws.Cells(4, 15) = "Greatest Total Volume" '04
ws.Cells(1, 16) = "Ticker" 'P1
ws.Cells(1, 17) = "Value" 'Q1

Dim summarytablerow2 As Integer
summarytablerow2 = 2

'define the last row of the summary data range
LastRowSummaryRange = ws.Cells(Rows.Count, 9).End(xlUp).Row

Dim HighestReturnSoFarSummaryRg As Double
Dim LowestReturnSoFarSummaryRg As Double
Dim HighestVolumeSoFarSummaryRg As Variant 'will this datatype work?

HighestReturnSoFarSummaryRg = ws.Cells(2, 11).Value
LowestReturnSoFarSummaryRg = ws.Cells(2, 11).Value
HighestVolumeSoFarSummaryRg = ws.Cells(2, 12).Value

For n = 3 To LastRowSummaryRange
            
        'check to see if this ticker's return is higher than last row's
        If ws.Cells(n, 11).Value > HighestReturnSoFarSummaryRg Then
            HighestReturnSoFarSummaryRg = ws.Cells(n, 11).Value 'set the highest return variable
            ws.Cells(2, 16) = ws.Cells(n, 9).Value 'add the ticker
            ws.Cells(2, 17) = ws.Cells(n, 11).Value 'populate the return itself
            
        'check to see if this ticker's return is lower than last row's
        ElseIf ws.Cells(n, 11).Value < LowestReturnSoFarSummaryRg Then
            LowestReturnSoFarSummaryRg = ws.Cells(n, 11).Value 'set the lowest return in summary range variable
            ws.Cells(3, 16) = ws.Cells(n, 9).Value 'add the ticker
            ws.Cells(3, 17) = ws.Cells(n, 11).Value 'populate the return itself
                                                
        End If
        
        'check to see if this ticker's volume is higher than last row'
        If CDec(ws.Cells(n, 12).Value) > CDec(HighestVolumeSoFarSummaryRg) Then
            HighestVolumeSoFarSummaryRg = CDec(ws.Cells(n, 12).Value)
            ws.Cells(4, 16) = ws.Cells(n, 9).Value 'add the ticker
            ws.Cells(4, 17) = CDec(ws.Cells(n, 12).Value)
        End If
        
Next n

ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"

Next ws

Application.ScreenUpdating = True

End Sub











