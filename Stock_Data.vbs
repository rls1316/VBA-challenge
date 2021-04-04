Sub Stock_Data()

  'Set loop for all worksheets
  For Each ws In Worksheets
  
    ' Set an initial variable for holding the ticker symbol
    Dim TickerSymbol As String

    ' Set an initial variable for holding the total stock volume
    Dim TotalStockVolume As Double
    TotalStockVolume = 0

    ' Keep track of the location for each ticker in the summary table
    Dim SummaryTableRow As Long
    SummaryTableRow = 2
  
    'Set variable for the yearly open price for each ticker
    Dim YearlyOpenPrice As Double
  
    'Set variable for the yearly close price for each ticker
    Dim YearlyClosePrice As Double
  
    'Set variable for the yearly change for each ticker
    Dim YearlyChange As Double
  
    'Set variable to identify the Previous Amount for each ticker
    Dim PreviousAmount As Long
    PreviousAmount = 2
  
    'Set variable to identify the percentage of change for each ticker from open to close
    Dim PercentageChange As Double
  
    'Set last row variables
    Dim LastRow As Long
    Dim LastRowValue As Long
  
    'Find the last row for the summary table
    LastSummaryRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    'Add column headers for summary outputs
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ' Loop through all the tickers
     For i = 2 To LastSummaryRow
    
        'Start total stock volume counter
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
        ' Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            'Add ticker symbol value to the summary row table
            TickerSymbol = ws.Cells(i, 1).Value
            ws.Range("I" & SummaryTableRow).Value = TickerSymbol
        
            'Find the Yearly Change Value and format with colors
            YearlyOpenPrice = ws.Range("C" & PreviousAmount)
            YearlyClosePrice = ws.Range("F" & i)
            YearlyChange = YearlyClosePrice - YearlyOpenPrice
            ws.Range("J" & SummaryTableRow).Value = YearlyChange
            If ws.Range("J" & SummaryTableRow).Value >= 0 Then
               ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
            Else
               ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            End If
        
            'Find the percentage of change value and format with % symbol
            If YearlyOpenPrice = 0 Then
               PercentageChange = 0
            ElseIf YearlyOpenPrice = ws.Range("C" & PreviousAmount) Then
                   PercentageChange = YearlyChange / YearlyOpenPrice
            End If
            ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
            ws.Range("K" & SummaryTableRow).Value = PercentageChange
                
            'Add total stock volume amount to summary row table and reset counter
            ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
            TotalStockVolume = 0
    
            ' Add one to the summary table row
            SummaryTableRow = SummaryTableRow + 1
            PreviousAmount = i + 1
        End If
    Next i
      
    'Find the last row for the bonus table
    LastBonusRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    'Set variable to find the Greatest Increase for each ticker
    Dim GreatestIncrease As Double
    GreatestIncrease = 0
  
    'Set variable to find the Greatest Decrease for each ticker
    Dim GreatestDecrease As Double
    GreatestDecrease = 0
  
    'Set variable for the Greatest Total Volume for each ticker
    Dim GreatestTotalVolume As Double
    GreatestTotalVolume = 0
    
    'Add column headers and row labels for bonus outputs
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
      
    ' Loop through the summary table
    For i = 2 To LastBonusRow
  
        'Find the Greatest % Increase and format cell; compare every "K" value to the Value field, and update both the Value and Ticker columns if "K" is greater than the Value field until the end of column K is reached
        If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
           ws.Range("Q2").Value = ws.Range("K" & i).Value
           ws.Range("P2").Value = ws.Range("I" & i).Value
        End If
        ws.Range("Q2").NumberFormat = "0.00%"
        
        'Find the Greatest % Decrease and format cell; compare every "K" value to the Value field, and update both the Value and Ticker columns if "K" is greater than the Value field until the end of column K is reached
        If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
           ws.Range("Q3").Value = ws.Range("K" & i).Value
           ws.Range("P3").Value = ws.Range("I" & i).Value
        End If
        ws.Range("Q3").NumberFormat = "0.00%"
        
        'Find the Greatest Total Volume; ; compare every "L" value to the Value field, and update both the Value and Ticker columns if "L" is greater than the Value field until the end of column L is reached
        If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
           ws.Range("Q4").Value = ws.Range("L" & i).Value
           ws.Range("P4").Value = ws.Range("I" & i).Value
        End If
    Next i

  Next ws
           
End Sub

