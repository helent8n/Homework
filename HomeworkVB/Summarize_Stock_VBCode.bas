Attribute VB_Name = "Module1"
Sub SummarizeStock1()

    ''--------------------------------
    ''Loop through all worksheets
    ''--------------------------------
    For Each ws In Worksheets

    
        ''--------------------------------------------------------------------
        ''Summarize ticker symbol and calculate total volume per ticker symbol
        ''--------------------------------------------------------------------

            'Set an initial variable for holding the ticker symbol
             Dim Ticker_Name As String

            'Set initial variable for holding the total per ticker symbol
            Dim Ticker_Volume_Total As Double
            Ticker_Volume_Total = 0

            'Keep track of the location for each ticker symbol in the summary table
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
    
            'Add Column Header : Ticker and Total Stock Volume
            ws.Cells(1, 12).Value = "Ticker"
            ws.Cells(1, 13).Value = "Total Stock Volume"
    
            'Loop through all ticker symbol
            For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

            'Check if we are still within the same ticker symbol, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Set the Ticker name
            Ticker_Name = ws.Cells(i, 1).Value

            'Add to the Volume Total
            Ticker_Volume_Total = Ticker_Volume_Total + ws.Cells(i, 7).Value

            'Print the Ticker symbol in the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Ticker_Name

            'Print the Volume total to the Summary Table
            ws.Range("M" & Summary_Table_Row).Value = Ticker_Volume_Total

            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
            'Reset the Ticker Volume Total for different ticker symbol
            Ticker_Volume_Total = 0

            ' If the cell immediately following a row is the same ticker symbol...
            Else

            ' Add to the Volume Total for different ticker symbol
            Ticker_Volume_Total = Ticker_Volume_Total + ws.Cells(i, 7).Value

            End If

        Next i
        
    ''--------------------------
    ''Summarize next worksheets
    ''--------------------------
    Next ws
    
End Sub






