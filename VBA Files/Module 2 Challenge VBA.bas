Attribute VB_Name = "Module2"
Sub Stock_Data():
    Dim ws As Worksheet
    Dim Row As Long
    Dim lastrow As Long
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
    
    
    For Each ws In ThisWorkbook.Worksheets
    
    
    'Identify open and close price columns
        open_pos = 3
        Close_pos = 6
        vol_pos = 7
        
    'Calculate lastrow
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    'Set output positions
        Summary_Table_Row = 2
        Ticker = 9
        Quarterly_Change_POS = 10
        Percent_Change_POS = 11
        Total_Stock_Volume_POS = 12
        
    'Set output values
        Quarterly_Change_VAL = 0
        Percent_Change_VAL = 0
        Total_Stock_Volume_VAL = 0
    
    'Label new column headers
        ws.Cells(1, Ticker).Value = "Ticker"
        ws.Cells(1, Quarterly_Change_POS).Value = "Quarterly Change"
        ws.Cells(1, Percent_Change_POS).Value = "Percent Change"
        ws.Cells(1, Total_Stock_Volume_POS).Value = "Total Stock Volume"
    
    'format number output columns
        ws.Columns(Quarterly_Change_POS).NumberFormat = "0.00"
        ws.Columns(Percent_Change_POS).NumberFormat = "0.00%"

    'Setup the loop
        For Row = 2 To lastrow
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
         
             'expand cell width so totals are readable
            ws.Cells.EntireColumn.AutoFit
        
        Next ws
        
        
End Sub

