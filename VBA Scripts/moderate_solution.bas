Attribute VB_Name = "moderate_solution"
Sub Easy_Solution()

Application.ScreenUpdating = False

Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Pct Change"
Cells(1, 12) = "Total Stock Volume"

Dim ticker As String

'create a running total for each ticker's volume
Dim vol_total As Variant

'set the initial vol_total at 0; we will reset it at end of loop before next i
vol_total = 0

'set the inital BegYrOpenPrice; we'll resent it as we go in the loop
Dim BegYrOpenPrice As Double
BegYrOpenPrice = Cells(2, 3).Value

Dim EndYrClosePrice As Double

Dim summary_table_row As Integer
summary_table_row = 2

'define the last row of the entire data range
LastRowData = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRowData

    'if the value in the row after this one is diff, assume this is last row for this ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'set the ticker variable to the value of the current row's ticker
        ticker = Cells(i, 1).Value
        
        'add this row's vol to the running total vol variable from the summary table
        vol_total = CDec(vol_total) + Cells(i, 7).Value
        
        'get the year end price for this ticker
        EndYrClosePrice = Cells(i, 6).Value
                
        'add ticker to the summary table
        Cells(summary_table_row, 9).Value = ticker
                
        'populate the yearly change value
        Cells(summary_table_row, 10).Value = EndYrClosePrice - BegYrOpenPrice
        
        'apply green and red color based on price change direction
        If Cells(summary_table_row, 10).Value > 0 Then
            Cells(summary_table_row, 10).Interior.ColorIndex = 4
        ElseIf Cells(summary_table_row, 10).Value < 0 Then
            Cells(summary_table_row, 10).Interior.ColorIndex = 3
        End If
                
        'populate the yearly Pct change value
        Cells(summary_table_row, 11).Value = ((EndYrClosePrice / BegYrOpenPrice) - 1)
        Cells(summary_table_row, 11).NumberFormat = "0.00%"
                  
        'set the last row in the summary table to the new vol_total
        Cells(summary_table_row, 12).Value = vol_total
        
        'go one row down on the summary table
        summary_table_row = summary_table_row + 1
        
        'reset the vol_total because that was the last row
        vol_total = 0
    
        'and reset the beg of yr price for the next ticker
        BegYrOpenPrice = Cells(i + 1, 3).Value
    
    Else
        
        'otherwise just add the current row's volume to the running running total
        vol_total = CDec(vol_total) + Cells(i, 7).Value
            
    End If
        
Next i

Application.ScreenUpdating = True

End Sub



