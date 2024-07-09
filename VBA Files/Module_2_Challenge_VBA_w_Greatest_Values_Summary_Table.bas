Attribute VB_Name = "Module2"
Sub Stock_Data_w_Greatest_Summary():
    Dim ws As Worksheet
    Dim Row As Long
    Dim lastRow As Long
    Dim Open_Val As Double
    Dim Close_val As Double
    Dim Summary_Table_Row As Long
    Dim Ticker As Long
    Dim Quarterly_Change_POS As Long
    Dim Percent_Change_POS As Long
    Dim Total_Stock_Volume_POS As Long
    Dim Quarterly_Change_VAL As Double
    Dim Percent_Change_VAL As Double
    Dim Total_Stock_Volume_VAL As Double
    Dim First_Open As Double
    Dim open_pos As Long
    Dim Close_pos As Long
    Dim vol_pos As Long
    Dim maxVolume As Double
    Dim maxPercent As Double
    Dim minPercent As Double
    Dim greatestTicker As Integer
    Dim greatestValue As Double
    Dim greatest_per_up As Integer
    Dim greatest_per_down As Integer
    Dim greatestVolume As Integer
    Dim greatest_per_up_val As Double
    Dim greatest_per_down_val As Double
    Dim greatestVolume_val As Double
    Dim minValue As Double
    Dim maxValue As Double
    Dim Maxticker As Variant
    
    'apply vba to all worksheets
    For Each ws In ThisWorkbook.Worksheets
    
    
    'Identify open and close price columns
        open_pos = 3
        Close_pos = 6
        vol_pos = 7
        
    'Calculate lastrow
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    'Set output positions
        Summary_Table_Row = 2
        Ticker = 9
        Quarterly_Change_POS = 10
        Percent_Change_POS = 11
        Total_Stock_Volume_POS = 12
        greatestTicker = 16
        greatestValue = 17
        greatest_per_up = 2
        greatest_per_down = 3
        greatestVolume = 4
        
    'Set output values
        Quarterly_Change_VAL = 0
        Percent_Change_VAL = 0
        Total_Stock_Volume_VAL = 0
    
    'Label new column headers
        ws.Cells(1, Ticker).Value = "Ticker"
        ws.Cells(1, Quarterly_Change_POS).Value = "Quarterly Change"
        ws.Cells(1, Percent_Change_POS).Value = "Percent Change"
        ws.Cells(1, Total_Stock_Volume_POS).Value = "Total Stock Volume"
        ws.Cells(1, greatestTicker).Value = "Ticker"
        ws.Cells(1, greatestValue).Value = "Value"
    'Label "greatest" summary table  names
        ws.Cells(greatest_per_up, 15).Value = "Greatest % Increase"
        ws.Cells(greatest_per_down, 15).Value = "Greatest % Decrease"
        ws.Cells(greatestVolume, 15).Value = "Greatest Total Volume"
    
    'format number output columns
        ws.Columns(Quarterly_Change_POS).NumberFormat = "0.00"
        ws.Columns(Percent_Change_POS).NumberFormat = "0.00%"
        ws.Cells(greatest_per_up, greatestValue).NumberFormat = "0.00%"
        ws.Cells(greatest_per_down, greatestValue).NumberFormat = "0.00%"
        ws.Cells(greatestVolume, greatestValue).NumberFormat = "0.00E+00"
       

    'Setup the loop
        For Row = 2 To lastRow
            'setup tickers first open value
                If ws.Cells(Row, 1).Value <> ws.Cells(Row - 1, 1).Value Then
                    First_Open = ws.Cells(Row, open_pos).Value
                End If
            
            'find next unique ticker and calculate/populate previous tickers totals and calc values
                If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
                
                'Grab ticker symbol's quarter close value
                    Close_val = ws.Cells(Row, Close_pos).Value
                
                    'Populate ticker symbol in summary table
                        ws.Cells(Summary_Table_Row, Ticker).Value = ws.Cells(Row, 1).Value
                    'Populate quarterly change
                        ws.Cells(Summary_Table_Row, Quarterly_Change_POS).Value = Close_val - First_Open
                        'color output value
                            If ws.Cells(Summary_Table_Row, Quarterly_Change_POS).Value >= 0.005 Then
                            'color positive cells green
                            ws.Cells(Summary_Table_Row, Quarterly_Change_POS).Interior.ColorIndex = 4
                            End If
                            
                            If ws.Cells(Summary_Table_Row, Quarterly_Change_POS).Value <= -0.005 Then
                            'color negative value cells red
                            ws.Cells(Summary_Table_Row, Quarterly_Change_POS).Interior.ColorIndex = 3
                            End If
                             
                    'Populate percent change
                        ws.Cells(Summary_Table_Row, Percent_Change_POS).Value = (((Close_val - First_Open) / First_Open) * 100) / 100
                    'Populate total stock volume
                        ws.Cells(Summary_Table_Row, Total_Stock_Volume_POS).Value = Total_Stock_Volume_VAL + ws.Cells(Row, vol_pos).Value
          
                'Move to next ticker symbol
                Summary_Table_Row = Summary_Table_Row + 1
                'Reset Value total
                Total_Stock_Volume_VAL = 0
                
                'if Tickers match, add stock volume value to total
                Else
                Total_Stock_Volume_VAL = Total_Stock_Volume_VAL + ws.Cells(Row, vol_pos).Value
        
                End If
         
        Next Row
        
    ' find max and min percents and max total volume
    
        'set new lastrow
            lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
        'find max percentage increase value
            maxValue = Application.WorksheetFunction.Max(ws.Range("K1:K" & lastRow))
            ws.Cells(greatest_per_up, greatestValue).Value = maxValue
            maxRow = Application.WorksheetFunction.Match(maxValue, ws.Range("K1:K" & lastRow), 0)
            Maxticker = ws.Cells(maxRow, Ticker).Value
            ws.Cells(greatest_per_up, greatestTicker).Value = Maxticker
        'find min percentage increase value
            minValue = Application.WorksheetFunction.Min(ws.Range("K1:K" & lastRow))
            ws.Cells(greatest_per_down, greatestValue).Value = minValue
            minRow = Application.WorksheetFunction.Match(minValue, ws.Range("K1:K" & lastRow), 0)
            Maxticker = ws.Cells(minRow, Ticker).Value
            ws.Cells(greatest_per_down, greatestTicker).Value = Maxticker
        'find max volume
            maxVolume = Application.WorksheetFunction.Max(ws.Range("L1:L" & lastRow))
            ws.Cells(greatestVolume, greatestValue).Value = maxVolume
            maxRow = Application.WorksheetFunction.Match(maxVolume, ws.Range("L1:L" & lastRow), 0)
            Maxticker = ws.Cells(maxRow, Ticker).Value
            ws.Cells(greatestVolume, greatestTicker).Value = Maxticker
        
    'expand cell width so totals are readable
        ws.Cells.EntireColumn.AutoFit
        
        Next ws
        
        
End Sub

