Sub Stocks()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                 '
'CWRU Data Analytics                                  Robert Wood '
'                                                                 '
'                                                                 '
'Unit 2 | Assignment - The VBA of Wall Street           2/23/2019 '
'                                                                 '
'                                                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Loop through all shets
For Each ws In Worksheets

'Declare variable to store total rows on sheet
Dim ItemCount As Long

'Declare variable to store current Ticker
Dim Ticker As String

'Declare Double to store Volume total for current Ticker
Dim TotalVolume As Double

'Declare Single to store Opening value
Dim Opening As Single

'Declare Single to store Closing value
Dim Closing As Single

'Declare Integer to store current row number in summary table
Dim SummaryRow As Integer

'Assign ItemCount variable as total rows on sheet
ItemCount = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Assign first Opening value
Opening = ws.Range("C2")

'Assign starting row for summary table
SummaryRow = 2

'Assign starting value for Ticker
Ticker = ws.Range("A2")

'Assign initial values for Closing and Total Volume
Closing = 0
TotalVolume = 0

'Print column headers for results table
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Start looping through rows, ignoring header row
For i = 2 To ItemCount

    'If the next row Ticker is the same as this row, add to our total and continue
    If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
        Ticker = ws.Cells(i, 1)
        TotalVolume = TotalVolume + ws.Cells(i, 7)

    'Otherwise, we are at the last row for this ticker
    Else
        'Update our Total Volume to include this last row
        TotalVolume = TotalVolume + ws.Cells(i, 7)
        
        'Assign a Closing value
        Closing = ws.Cells(i, 6)
        
        'Print the Ticker symbol
        ws.Cells(SummaryRow, 9) = Ticker
        
        'Print the Yearly Change
        ws.Cells(SummaryRow, 10) = Closing - Opening
        
        'Print the Percent Change
        'Guard against divide by zero error
        If Opening = 0 Then
            ws.Cells(SummaryRow, 11) = "N/A"
        Else
            ws.Cells(SummaryRow, 11) = (Closing - Opening) / Opening
        End If
        
        'Print the Total Volume
        ws.Cells(SummaryRow, 12) = TotalVolume
        
        'Format Cell Color for Yearly Change (green for positive, red for negative)
        If ws.Cells(SummaryRow, 10).Value > 0 Then
            ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
        End If
        
        If ws.Cells(SummaryRow, 10).Value < 0 Then
            ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
        End If
        
        'Format Percent Change column as percentages
        ws.Range("K:K").NumberFormat = "0.00%"
        
        'Assign new Ticker name
        Ticker = ws.Cells(i + 1, 1)

        'Reset Total Volume
        TotalVolume = 0
        
        'Set new Opening value for upcoming Ticker
        Opening = ws.Cells(i + 1, 3)
        
        'Increment row for summary table
        SummaryRow = SummaryRow + 1

    End If
    
Next i
        
'Print row and column labels for secondary (summary) results table
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Loop through Summary table
For i = 2 To SummaryRow - 1

    'If Percent Change of current row is greater than current Greatest % Increase (and not N/A),
    'overwrite Greatest % Increase
    If ws.Cells(i, 11).Value > ws.Range("Q2").Value And ws.Cells(i, 11).Value <> "N/A" Then
        ws.Range("P2").Value = ws.Cells(i, 9)
        ws.Range("Q2").Value = ws.Cells(i, 11)
    End If
    
    'If Percent Change of current row is less than current Greatest % Decrease (and not N/A),
    'overwrite Greatest % Decrease
    If ws.Cells(i, 11).Value < ws.Range("Q3").Value And ws.Cells(i, 11).Value <> "N/A" Then
        ws.Range("P3").Value = ws.Cells(i, 9)
        ws.Range("Q3").Value = ws.Cells(i, 11)
    End If
    
    'If Total Stock Volume of current row is more than current Greatest Total Volume,
    'overwrite Greatest Total Volume
    If ws.Cells(i, 12).Value > ws.Range("Q4").Value Then
        ws.Range("P4").Value = ws.Cells(i, 9)
        ws.Range("Q4").Value = ws.Cells(i, 12)
    
    End If

Next i

'Format cell styles as percentages
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"
        
'Resize columns to auto-fit
ws.Range("I1:Q4").Columns.AutoFit

'Move on to next sheet
Next ws

End Sub