# vba-challenge

'Create a script that loops through all the stocks for one year and outputs the following information:
'
'    The ticker symbol
'
'    Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'
'    The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'
'    The total stock volume of the stock.
'
'    Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.
'
'    Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

Sub stonks():

'loop through every worksheet
For Each ws In Worksheets

    'Declare variables
    Dim ticker, tickerInc, tickerDec, tickerVol As String
    Dim openPrice, closePrice, yearlyChange, percentChange, greatestInc, greatestDec, greatestVol As Double
    Dim totalStock As LongLong
    
    'Set total stock as 0 so we can aggregate total
    totalStock = 0
    
    'Keep track of location for each ticker in summary table
    Dim summaryTableRow As Integer
    summaryTableRow = 2
    
    'get LastRow/LastColumn since all worksheet have different number of rows/columns of original data
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    'get LastRow/LastColumn from summary table; will need this later when checking the max/min for the greatest % increase/decrease
    LastRowSummary = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    
    'Set the opening price for the first ticker
    openPrice = ws.Cells(2, 3).Value
        
        For i = 2 To LastRow
            
        
            'Check if next ticker is the same as current ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                ticker = ws.Cells(i, 1).Value
                
                'To grab the last close value to change
                closePrice = ws.Cells(i, 6).Value
            
                'Calculate values for yearly/percent change
                yearlyChange = closePrice - openPrice
                percentageChange = yearlyChange / openPrice
            
                'Print the ticker into the summary table
                ws.Range("I" & summaryTableRow).Value = ticker
                ws.Range("J" & summaryTableRow).Value = yearlyChange
                
                
                'Change interior colour for yearly change and percentage change just like the example
                'Logically, if yearly change is greater/less than 0, percentage change also moves in the same way
                'Therefore I am able to put this into a single if statement
                If yearlyChange > 0 Then
        
                    ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
                    ws.Range("K" & summaryTableRow).Interior.ColorIndex = 4
            
                ElseIf yearlyChange < 0 Then
                    ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
                    ws.Range("K" & summaryTableRow).Interior.ColorIndex = 3
                   
                End If
                
                
                
                ws.Range("K" & summaryTableRow).Value = percentageChange
                ws.Range("K" & summaryTableRow).NumberFormat = "0.00%" 'Changes the formatting to %
                ws.Range("Q2").Value = greatestInc
                ws.Range("L" & summaryTableRow).Value = totalStock
            
                'Update the open yearly price with next ticker
                openPrice = ws.Cells(i + 1, 3).Value
                
                'add one to summary table row so it moves to next row when it outputs
                summaryTableRow = summaryTableRow + 1
                
                'resets the total stock to 0 when moving to next ticker symbol
                totalStock = 0
            
            Else
                'Add the current stock to the total
                totalStock = totalStock + ws.Cells(i, 7).Value
            
            End If
                
        Next i
        
        'Declare a starting value for the greatest percentage increase/decrease as well as for ticker symbols and greatest volume
        'Create Initial conditions
        greatestInc = ws.Cells(2, 11).Value
        greatestDec = ws.Cells(2, 11).Value
        greatestVol = ws.Cells(2, 12).Value

        
        'Loop through the summary table to grab the greatest increase and decrease
        For i = 2 To LastRowSummary
                
            If ws.Cells(i + 1, 11).Value > greatestInc Then
            
                greatestInc = ws.Cells(i + 1, 11).Value 'grabs the greatest increase
                tickerInc = ws.Cells(i + 1, 9).Value 'grabs the ticker symbol
                
            ElseIf ws.Cells(i + 1, 11).Value < greatestDec Then
            
                greatestDec = ws.Cells(i + 1, 11).Value 'grabs the greatest decrease
                tickerDec = ws.Cells(i + 1, 9).Value 'grabs the ticker symbol
                
            End If
            
            'Compare with a value on the list. If current value is bigger than what it's compared against, then move to next i.
            'if it's smaller than the value it's being compared to, replace and move to next i
            If ws.Cells(i + 1, 12).Value > greatestVol Then
            
                greatestVol = ws.Cells(i + 1, 12).Value 'grabs the greatest volume
                tickerVol = ws.Cells(i + 1, 9).Value 'grabs the ticker symbol
                
            Else
                End If
            
        Next i
        
        'Output into the secondary summary table
        ws.Range("P2").Value = tickerInc
        ws.Range("Q2").Value = greatestInc
        ws.Range("Q2").NumberFormat = "0.00%" 'Changes the formatting to %
        
        ws.Range("P3").Value = tickerDec
        ws.Range("Q3").Value = greatestDec
        ws.Range("Q3").NumberFormat = "0.00%" 'Changes the formatting to %
        
        ws.Range("P4").Value = tickerVol
        ws.Range("Q4").Value = greatestVol
    
    Next ws
    
End Sub








