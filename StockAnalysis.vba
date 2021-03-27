Attribute VB_Name = "Module1"
Sub stockAnalysis()

    Dim ws As Worksheet
    
    'Remembers the first worksheet as the active sheet
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet
    
    'Loops through every worksheet
    For Each ws In Worksheets

        ws.Activate
        'Sets the headers of cells I1 to L1, and bonus data
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
    
        'Autofits text to the columns
        Range("I1:L1").EntireColumn.AutoFit
        Columns("O").ColumnWidth = 19
        Columns("P").ColumnWidth = 6
        Columns("Q").ColumnWidth = 10
        
        'Formats % Increase/Decrease to Percent
        Range("Q2:Q3").NumberFormat = "0.00%"
    
        'Gets collective data for all tickers
        'Get value of last row
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Dim currentTicker As String
        
        'Total stock volume for each ticker
        'Opening price and closing price for the year
        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim totalVolume As Double
        
        'Row in which data for current ticker will be entered
        Dim summaryTableRow As Integer
        summaryTableRow = 2
        totalVolume = 0
    
        'Gets stock data for volume and price change
        For Row = 2 To lastRow
            'Checks if an opening price has been set, and sets it to the first price if not
            If openingPrice = 0 Then
                openingPrice = Cells(Row, 3).Value
            Else
            End If
        
            'Check if we have changed tickers
            If (Cells(Row + 1, 1).Value <> Cells(Row, 1).Value) Then
            
                'Change ticker, add the total stock volume, and get closing price
                currentTicker = Cells(Row, 1).Value
                closingPrice = Cells(Row, 6).Value
                'Update data table with ticker, volume, and changes
                'Add ticker to new row
                Range("I" & summaryTableRow).Value = currentTicker
                'Calculate yearly change and adds it to the yearly change column
                Range("J" & summaryTableRow).Value = closingPrice - openingPrice
                Range("J" & summaryTableRow).NumberFormat = "0.00"
                
                Range("L" & summaryTableRow).Value = totalVolume
                
                'Colors positive change green and negative change red
                If Range("J" & summaryTableRow).Value > 0 Then
                    Range("J" & summaryTableRow).Interior.ColorIndex = 4
                Else
                    Range("J" & summaryTableRow).Interior.ColorIndex = 3
                End If
                'Calculate percentage change
                'Check to ensure it doesn't divide by zero, otherwise change equals zero
                If openingPrice <> 0 Then
                    Range("K" & summaryTableRow).Value = (closingPrice - openingPrice) / openingPrice
                Else
                    Range("K" & summaryTableRow).Value = 0
                End If
                Range("K" & summaryTableRow).NumberFormat = "0.00%"
                
                'Reset numbers and move to next row of data table
                openingPrice = 0
                summaryTableRow = summaryTableRow + 1
                totalVolume = 0
            
            Else
            
                'Add up current stock volume to the total
                totalVolume = totalVolume + Cells(Row, 7).Value
        
            End If
        
        Next Row
        
        'Variables for bonus analysis
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume
        Dim increaseTicker As String
        Dim decreaseTicker As String
        Dim volumeTicker As String
        'Increase and decrease initialized as the first value in the range, and updated accordingly
        greatestIncrease = Cells(Row, 11).Value
        greatestDecrease = Cells(Row, 11).Value
        greatestVolume = 0
        
        'Finds the ticker and value for the greatest % increase/decrease, and greatest total volume
        For Row = 2 To lastRow
            
            'Looks at the cells for each variable, saves the correct number, and grabs ticker
            'Finds greatest % increase
            If Cells(Row, 11).Value > greatestIncrease Then
                greatestIncrease = Cells(Row, 11).Value
                increaseTicker = Cells(Row, 9)
            Else
            End If
            
            'Finds greatest % decrease
            If Cells(Row, 11).Value < greatestDecrease Then
                greatestDecrease = Cells(Row, 11).Value
                decreaseTicker = Cells(Row, 9)
            Else
            End If
            
            'Finds greatest total Volume
            If Cells(Row, 12).Value > greatestVolume Then
                greatestVolume = Cells(Row, 12).Value
                volumeTicker = Cells(Row, 9)
            Else
            End If
        
        Next Row
        
        'Enter the values into respective cells
        Range("Q2").Value = greatestIncrease
        Range("Q3").Value = greatestDecrease
        Range("Q4").Value = greatestVolume
        Range("P2").Value = increaseTicker
        Range("P3").Value = decreaseTicker
        Range("P4").Value = volumeTicker
        
    Next
    
    'Go back to the first worksheet
    starting_ws.Activate

End Sub

Sub Reset()

    Dim ws As Worksheet
    
    'Remembers the first worksheet as the active sheet
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet
    
    'Loops through every worksheet, deletes columns I to L
    For Each ws In Worksheets

        ws.Activate
        Columns("I:L").ClearContents
        Columns("O:Q").ClearContents
        Columns("J").Interior.ColorIndex = 0
    
    Next
    
    'Go back to the first worksheet
        starting_ws.Activate

End Sub
