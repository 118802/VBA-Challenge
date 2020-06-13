Attribute VB_Name = "Module2"

Sub Multiple_year_stock_data()

For Each ws In Worksheets

    Dim Worksheetname As String
    
    Worksheetname = ws.Name
    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Dim Total_Stock_Volume As Double
    
            Total_Stock_Volume = 0
    
    ws.Cells(1, 9).Value = "Ticker"
    
    ws.Cells(1, 10).Value = "Yearly Change"
    
    ws.Cells(1, 11).Value = "Percent Change"
    
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    Dim i As Long
    
    Dim Stock_Ticker As String
    
    Dim Yearly_Open As Double
    
            Yearly_Open = 0
    
    Dim Yearly_Close As Double
    
            Yearly_Close = 0

    Dim Yearly_Change As Double
    
            Yearly_Change = 0
    
    Dim Percent_Change As Double
    
    Dim tickertype As Long
    
            tickertype = 2



For i = 2 To lastRow

    Yearly_Open = ws.Cells(tickertype, 3).Value
    
    Yearly_Close = ws.Cells(tickertype, 6).Value


    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        Stock_Ticker = ws.Cells(i, 1).Value
        
        Yearly_Change = Yearly_Close - Yearly_Open
        
        Percent_Change = Yearly_Change / Yearly_Open
        
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        ws.Range("L" & tickertype).Value = Total_Stock_Volume
        
        ws.Range("I" & tickertype).Value = Stock_Ticker
        
        ws.Range("J" & tickertype).Value = Yearly_Change
        
        ws.Range("K" & tickertype).Value = Percent_Change
        
        tickertype = tickertype + 1
        
        Yearly_Change = 0
        
        Total_Stock_Volume = 0
        
       
    
    Else
        
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
    End If

Next i


Dim year_change_LastRow As Long

year_change_LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row


For i = 2 To year_change_LastRow

    If ws.Cells(i, 10).Value >= 0 Then
        
        ws.Cells(i, 10).Interior.ColorIndex = 4
    
    Else
        
        ws.Cells(i, 10).Interior.ColorIndex = 3
   
   End If

Next i
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ws.Cells(1, 16).Value = "Ticker"
    
    ws.Cells(1, 17).Value = "Value"
    
    

Dim percent_change_LastRow As Long

percent_change_LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

Dim percent_change_max As Double

percent_change_max = 0

Dim percent_change_min As Double

percent_change_min = 0


For i = 2 To percent_change_LastRow

    If ws.Cells(i, 11).Value > percent_change_max Then
        
        percent_change_max = ws.Cells(i, 11).Value
        
          ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        
    ElseIf ws.Cells(i, 11).Value < percent_change_min Then
        
        percent_change_min = ws.Cells(i, 11).Value
        
           ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        
        End If

Next i

        ws.Cells(2, 17).Value = percent_change_max
        
      
        ws.Cells(3, 17).Value = percent_change_min
        
     
    


Dim Total_Stock_Volume_LastRow As Long

Total_Stock_Volume_LastRow = ws.Cells(Rows.Count, 12).End(xlUp).Row

Dim Total_Stock_Volume_Max As Double

Total_Stock_Volume_Max = 0


For i = 2 To Total_Stock_Volume_LastRow


    If Total_Stock_Volume_Max < ws.Cells(i, 12).Value Then
        
        Total_Stock_Volume_Max = ws.Cells(i, 12).Value
        
        
    End If

Next i

        ws.Cells(4, 17).Value = Total_Stock_Volume_Max
        
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    
Next ws

End Sub


