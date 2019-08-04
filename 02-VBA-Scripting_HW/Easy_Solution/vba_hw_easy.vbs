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

        ' Set Header for summary table
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Total Stock Volume"

        ' Loop through all stocks
        For i = 2 To 70926

            ' Check if we are still within the same ticker, if not
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                ' Set the ticker name
                Ticker_Name = Cells(i, 1).Value
            
                ' Add to the total stock volume
                Ticker_Total = Ticker_Total + Cells(i, 7).Value

                ' Print the Ticker name in the Summary Table
                Range("I" & Summary_Table_Row).Value = Ticker_Name
            
                ' Print the Total Stock Volume
                Range("J" & Summary_Table_Row).Value = Ticker_Total
        
                Summary_Table_Row = Summary_Table_Row + 1
            
                Ticker_Total = 0
            
                ' If a cell immediately following a row is the same ticker then
            Else
            
                ' Add to Ticker Total
                Ticker_Total = Ticker_Total + Cells(i, 7).Value

            End If

        Next i

    Next WS
        
End Sub