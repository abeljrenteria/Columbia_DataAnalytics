Sub stock_volume()

    ' Loop through all sheets
    Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
        WS.Activate

        ' Set initial variable for holding the stock ticker
        Dim Stock_Ticker As String

        ' Set initial variable for holding the total stock volume
        Dim StockVolume_Total As Double

        ' Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        ' Set initial variables for year prices
        Dim year_open As Double
        Dim year_close As Double
        Dim year_difference As Double
        
        ' Set Header for summary table
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"

        ' Set initial year open price of the first ticker
        year_open = Cells(2, 3).Value

        ' Loop through all stocks
        For i = 2 To 70926

            ' Check if we are still within the same ticker, if not
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                ' Set the ticker name
                Ticker_Name = Cells(i, 1).Value
            
                ' Add to the total stock volume
                Ticker_Total = Ticker_Total + Cells(i, 7).Value
            
                ' Set the Close price
                year_close = Cells(i, 6).Value
            
                ' Calculate the price difference
                year_difference = year_close - year_open

                ' Print the Ticker name in the Summary Table
                Range("I" & Summary_Table_Row).Value = Ticker_Name
            
                ' Print the Total Stock Volume
                Range("L" & Summary_Table_Row).Value = Ticker_Total
            
                'Print the difference b/w year open and closing price
                Range("J" & Summary_Table_Row).Value = year_difference
            
                ' Print percent change
                Range("K" & Summary_Table_Row).Value = (year_difference / year_open)
                Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
                Summary_Table_Row = Summary_Table_Row + 1
            
                Ticker_Total = 0
            
                ' Set the open price for the rest of the tickers
                year_open = Cells(i + 1, 3)
            
                ' If a cell immediately following a row is the same ticker then
            Else
            
                ' Add to Ticker Total
                Ticker_Total = Ticker_Total + Cells(i, 7).Value

            End If

        Next i
        
        ' --Conditionaly Format the cells--
        ' Find last row of summary table
        Summary_Last_Row = WS.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Set the colors of the cells
        For j = 2 To Summary_Last_Row
            If Cells(j, 10).Value >= 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j

    Next WS
        
End Sub