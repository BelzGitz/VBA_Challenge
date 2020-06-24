Attribute VB_Name = "Module1"
Sub Worksheet():
'variable declaration'
Dim ticker As String
Dim yearly_change  As Double
Dim percentage_change As Double
Dim total As Double
Dim year_close As Double
Dim year_open As Double
Dim start_open As Double

start_open = 2


'assigning cell headers'
Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "yearly_change"
Cells(1, 11).Value = "percentage chnage"
Cells(1, 12).Value = "total stock volume"

'set an initial variable for holdingtotal stock volume'
total = 0

'Keep track of each ticker in the summary table'

Dim summary_table_row As Long
summary_table_row = 2

'Determine last row'
    last_row = Cells(Rows.Count, 1).End(xlUp).Row

'loop through all ticker'

    For i = 2 To last_row
    
'check if we are with in same ticker'
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
'set ticker symbol'
ticker = Cells(i, 1).Value

'add to total volume
total = total + Cells(i, 7).Value


'print ticker symbol in the summary table'
Range("I" & summary_table_row).Value = ticker

'print the total volume amount'
Range("L" & summary_table_row).Value = total

'set yearly_change and percentage_change

year_close = Cells(i, 6).Value
year_open = Cells(start_open, 3).Value

start_open = i + 1

yearly_change = year_close - year_open

If year_open <> 0 Then

    percentage_change = yearly_change / year_open * 100
Else
    percentage_change = 0
End If
  
 'add values in to summary table
 Cells(summary_table_row, 10).Value = yearly_change
 Cells(summary_table_row, 11).Value = "%" & percentage_change
 
 
 'color range for summary table condition'
   If Cells(summary_table_row, 10).Value > 0 Then
      Cells(summary_table_row, 10).Interior.ColorIndex = 4
    Else
      Cells(summary_table_row, 10).Interior.ColorIndex = 3
    
    End If

'add 1 to the summary table'
summary_table_row = summary_table_row + 1

'reset total volume'
total = 0

'if the cell immediately following a row is the same ticker..'
Else

'add to the total'
total = total + Cells(i, 7).Value

End If



          
Next i

End Sub
