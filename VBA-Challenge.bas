Attribute VB_Name = "Module1"
'Set up function
Sub StockPrice()

    ' Loop through all sheets
    For Each ws In Worksheets
        
        'Create a Varaible to hold Last Row, Ticker, Quarterly change and percent change
        Dim WorksheetName As String
        Dim Ticker_Name As String
        Dim Quarterly_Change As Double
        Dim FirstOpenPrice As Double
        Dim LastClosePrice As Double
        Dim Percent_Change As Double
        Dim Total_Stock_Volume As Double
        
        'Set initial variable for holding total vol per ticker
        Total_Stock_Volume = 0

       
   ' Keep track of the location for each credit card brand in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        
        ' Determine the Last Row of each worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
         ' Add the word Ticker, Quarterly change, percent change, volume to the 9th, 10th, 11th, 12th Column Header
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        

        ' Loop through all tickers in a worksheet
        For i = 2 To LastRow
        
            'Check if we are still within the same ticker, if not...
            
            If i = 2 Then
            
                'record first open price
                FirstOpenPrice = ws.Cells(i, 3).Value
                
                'Set ticker name
                Ticker_Name = ws.Cells(i, 1).Value
                
                ' Print ticker in Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                        
            ElseIf i = LastRow Then
            
                'record last close price and calc quarterly and percent change
                LastClosePrice = ws.Cells(i, 6).Value
                Quarterly_Change = LastClosePrice - FirstOpenPrice
                Percent_Change = (LastClosePrice / FirstOpenPrice - 1)
                
                ' Print the quarterly price change and percent change to the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            
            ElseIf i > 2 And i <> LastRow Then
            
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            
               'Set ticker name
                    Ticker_Name = ws.Cells(i, 1).Value
                
                ' Print ticker in Summary Table
                    ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                'record first open price
                    FirstOpenPrice = ws.Cells(i, 3).Value
                
    
                ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                    'record last close price and calc quarterly and percent change
                    LastClosePrice = ws.Cells(i, 6).Value
                    Quarterly_Change = LastClosePrice - FirstOpenPrice
                    Percent_Change = (LastClosePrice / FirstOpenPrice - 1)
                
                    ' Print the quarterly price change and percent change to the Summary Table
                    ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
                    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                    
                    'Add one to summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
            
            
                End If
                
            End If
        
        Next i
        
            ws.Range("K2: K" & Summary_Table_Row).NumberFormat = "0.00%"
            ws.Range("J2: J" & Summary_Table_Row).NumberFormat = "0.00"

    Next ws


End Sub



'Set up function
Sub StockVol()

    ' Loop through all sheets
    For Each ws In Worksheets
        
        'Create a Varaible to hold Last Row, Ticker, Quarterly change and percent change
        Dim WorksheetName As String
        Dim Total_Stock_Volume As Double
        
        'Set initial variable for holding total vol per ticker
        Total_Stock_Volume = 0

       
   ' Keep track of the location for each credit card brand in the summary table
        Dim Summary_Table_Row_v2 As Integer
        Summary_Table_Row_v2 = 2
        
        
        ' Determine the Last Row of each worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

        ' Loop through all tickers in a worksheet
        For i = 2 To LastRow
        
            'Check if we are still within the same ticker, if not...
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
               'Add to total stock vol
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                ' Print ticker in Summary Table
                ws.Range("L" & Summary_Table_Row_v2).Value = Total_Stock_Volume
                
                'Add one to summary table row
                Summary_Table_Row_v2 = Summary_Table_Row_v2 + 1
                    
                'Reset total stock volume
                Total_Stock_Volume = 0
    
            Else
                
                'Add to total stock vol
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
            End If
        
        Next i

    Next ws

End Sub

'Set up function
Sub StockMaxSummary()

    ' Loop through all sheets
    For Each ws In Worksheets
        
        'Create a Varaible to hold Last Row, Ticker, Quarterly change and percent change
        Dim WorksheetName As String
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        
        ' Add the word Ticker, Quarterly change, percent change, volume to the 9th, 10th, 11th, 12th Column Header
        ws.Cells(2, 14).Value = "Greatest % increase"
        ws.Cells(4, 14).Value = "Greatest total volume"
        ws.Cells(1, 15).Value = "Ticker"
         ws.Cells(1, 16).Value = "Value"
        
        ' Determine the Last Row of each worksheet - REVISION
        LastNewRow = ws.Range("I" & Rows.Count).End(xlUp).Row
    
       GreatestIncrease = 0
        ' Loop through all tickers in a worksheet
        For i = 2 To LastNewRow
        
            'Compare percent change
            
            If ws.Cells(i, 11).Value > GreatestIncrease Then
            
                GreatestIncrease = ws.Cells(i, 11).Value
                
                ' Print percent change in Summary Table
                ws.Cells(2, 16).Value = GreatestIncrease
                ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                      
            End If
        
        Next i
            ws.Cells(2, 16).NumberFormat = "0.00%"

    Next ws

End Sub

'Set up function
Sub StockMinSummary()

    ' Loop through all sheets
    For Each ws In Worksheets
        
        'Create a Varaible to hold Last Row, Ticker, Quarterly change and percent change
        Dim WorksheetName As String
        Dim GreatestDecrease As Double
        
        ' Add the word Ticker, Quarterly change, percent change, volume to the 9th, 10th, 11th, 12th Column Header
        ws.Cells(3, 14).Value = "Greatest % decrease"
        
        ' Determine the Last Row of each worksheet - REVISION
        LastNewRow = ws.Range("I" & Rows.Count).End(xlUp).Row
    
       GreatestDecrease = 0
        ' Loop through all tickers in a worksheet
        For i = 2 To LastNewRow
        
            'Compare percent change
            
            If ws.Cells(i, 11).Value < GreatestDecrease Then
            
                GreatestDecrease = ws.Cells(i, 11).Value
                
                ' Print percent change in Summary Table
                ws.Cells(3, 16).Value = GreatestDecrease
                ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
                      
            End If
        
        Next i
            ws.Cells(3, 16).NumberFormat = "0.00%"

    Next ws

End Sub


'Set up function
Sub StockVolSummary()

    ' Loop through all sheets
    For Each ws In Worksheets
        
        'Create a Varaible to hold Last Row, Ticker, Quarterly change and percent change
        Dim WorksheetName As String
        Dim GreatestTotalVol As Double
        
        ' Add the word Ticker, Quarterly change, percent change, volume to the 9th, 10th, 11th, 12th Column Header
        ws.Cells(4, 14).Value = "Greatest total volume"
        
        ' Determine the Last Row of each worksheet - REVISION
        LastNewRow = ws.Range("I" & Rows.Count).End(xlUp).Row
    
       GreatestTotalVol = 0
        ' Loop through all tickers in a worksheet
        For i = 2 To LastNewRow
        
            'Compare percent change
            
            If ws.Cells(i, 12).Value > GreatestTotalVol Then
            
                GreatestTotalVol = ws.Cells(i, 12).Value
                
                ' Print percent change in Summary Table
                ws.Cells(4, 16).Value = GreatestTotalVol
                ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
                      
            End If
        
        Next i

    Next ws

End Sub
    
    
'Set up function
Sub ConditionFormatting()

    ' Loop through all sheets
    For Each ws In Worksheets
        
        'Create a Varaible to hold Last Row, Ticker, Quarterly change and percent change
        Dim WorksheetName As String
        
        ' Determine the Last Row of each worksheet - REVISION
        LastNewRow = ws.Range("K" & Rows.Count).End(xlUp).Row
    
        ' Loop through all tickers in a worksheet
        For i = 2 To LastNewRow
        
            'Compare percent change
            
            If ws.Cells(i, 10).Value > 0 Then
        
                ws.Cells(i, 10).Interior.ColorIndex = 4
            
            ElseIf ws.Cells(i, 10).Value < 0 Then
            
                ws.Cells(i, 10).Interior.ColorIndex = 3
                      
            End If
        
        Next i

    Next ws

End Sub

