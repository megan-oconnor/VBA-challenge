Sub tickerinfo()
    'loop through all sheets
    For Each ws In Worksheets
        'set initial variables for calculated columns
        Dim Ticker_Name As String
        Dim Quarter_Close As Double
        Dim Quarter_Open As Double: Quarter_Open = Cells(2, 3).Value
        Dim Percent_Change As Double
        Dim Total_Volume As Double: Total_Volume = 0

        'define summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'summary table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'get unique tickers
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        
        For i = 2 To lastRow
            'check if ticker is same as next entry
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                'set the ticker name
                Ticker_Name = ws.Cells(i, 1).Value
                
                'define quarter close value
                Quarter_Close = ws.Cells(i, 6).Value
                
                'add to the total volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
                'calculate the quarter change
                Quarter_Change = Quarter_Close - Quarter_Open
                
                'calculate the percent change
                Percent_Change = Quarter_Change / Quarter_Open
                
                'print the ticker in the summary table
                ws.Cells(Summary_Table_Row, 9).Value = Ticker_Name
                
                'print the quarter change in the summary table
                ws.Cells(Summary_Table_Row, 10).Value = Quarter_Change

                'format cell color for quarter change in summary table
                If Quarter_Change > 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                ElseIf Quarter_Change < 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                End If
                
                'print the percent change in the summary table
                ws.Cells(Summary_Table_Row, 11).Value = Percent_Change

                'format percent change in summary table
                ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                
                'print the total volume in the summary table
                ws.Cells(Summary_Table_Row, 12).Value = Total_Volume
                
                'reset the quarter open value
                Quarter_Open = ws.Cells(i + 1, 3).Value
                
                'add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

                'reset the Total Volume
                Total_Volume = 0
            
            'if the cell immediately following a row is the same
            Else
                
                'add to the Total Volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
            End If
        
        Next i

        'insert new table
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest % Increase"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        'define variables for stock performance
        Dim grt_increase As Double
        Dim grt_decrease As Double
        Dim grt_volume As Double

        'define search range for percent change
        Dim r As Range
        Set r = ws.Range("K2:K" & Rows.Count)

        'get the greatest percent decrease
        grt_decrease = Application.WorksheetFunction.Min(r)

        'add to new table
        ws.Range("Q3").Value = grt_decrease

        'get the greatest percent increase
        grt_increase = Application.WorksheetFunction.Max(r)

        'add to new table
        ws.Range("Q2").Value = grt_increase

        'define search range for total volume
        Set r = ws.Range("L2:L" & Rows.Count)

        'get the greatest total volume
        grt_volume = Application.WorksheetFunction.Max(r)

        'add to new table
        ws.Range("Q4").Value = grt_volume

        'get the number of tickers in the summary table
        lastRowTickers = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        'find the tickers associated with the highlighted values we just pulled
        For i = 2 To lastRowTickers
            'find the ticker with the matching percent decrease
            If ws.Cells(i,11).Value = grt_decrease Then
                ws.Cells(3,16).Value = ws.Cells(i,9).Value
            'find the ticker with the matching percent increase    
            ElseIf ws.Cells(i,11).Value = grt_increase Then
                ws.Cells(2,16).Value = ws.Cells(i,9).Value
            End If

            'find the ticker with the matching total volume
            If ws.Cells(i,12).Value = grt_volume Then
                ws.Cells(4,16).Value = ws.Cells(i,9).Value
            End If
        Next i

    Next ws

End Sub