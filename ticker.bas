Sub tickerProblem():
  
  'set this for to go through each worksheet
For Each ws In Worksheets
  'setting variable to hold ticker name
  Dim ticker As String
  
  'set variable to total volume
  Dim totalVolume As Double
  totalVolume = 0
  
  'set variable to hold prices
  Dim openPrice As Double
  Dim closePrice As Double
  
  'set variable to hold change in price
  Dim changePrice As Double
    
  'set variable to hold percent change
  Dim percentChange As Double
  
  'set variable to hold max variables
  Dim maxPercent As Double
  Dim maxTicker As String
  
  'set variable to hold min variables
  Dim minPercent As Double
  Dim minTicker As String
  
  'set variables for max volume
  Dim maxTotalVolume As Double
  Dim maxVolname As String

    
  'keep track of yearly change row
  Dim yearly_change_row As Integer
  yearly_change_row = 2
  
  
  'counting rows
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
  'setting column names for additions
   ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 10).Value = "Yearly Change"
   ws.Cells(1, 11).Value = "Percent Change"
   ws.Cells(1, 12).Value = "Total Stock Volume"
   ws.Cells(2, 15).Value = "Greatest % Increase"
   ws.Cells(3, 15).Value = "Greatest % Decrease"
   ws.Cells(4, 15).Value = "Greatest Total Volume"
   ws.Cells(1, 16).Value = "Ticker"
   ws.Cells(1, 17).Value = "Value"
   
   'initializing all these cells so then as it starts at nothing and then when the for starts it begins adding values
   openPrice = Cells(2, 3).Value
   MaxChange = 0
   maxTicker = " "
   
   MinChange = 0
   minTicker = " "
   
   maxTotalVolume = 0
   maxVolname = " "
   
   
   'loop through to get ticker name
   For i = 2 To LastRow
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
       'set ticker name
       ticker = ws.Cells(i, 1).Value
       
       'set annual close
       closePrice = ws.Cells(i, 6).Value
       
       'calculate change in price
       changePrice = closePrice - openPrice
       
       'calculate percent change
       'wasn't sure how to deal with zeroes so I set it 0 but I'm sure there is another way to do this
      If openPrice <> 0 Then
       percentChange = (changePrice / openPrice)
         Else
         percentChange = 0
       End If
             
       'add to the volume
       totalVolume = totalVolume + ws.Cells(i, 7).Value
       
       'print ticker
       ws.Range("I" & yearly_change_row).Value = ticker
       
       'print yearly change
       ws.Range("J" & yearly_change_row).Value = changePrice
       If changePrice < 0 Then
         ws.Range("J" & yearly_change_row).Interior.ColorIndex = 3
         Else
         ws.Range("J" & yearly_change_row).Interior.ColorIndex = 4
       End If
       
       'print percentage change
       ws.Range("K" & yearly_change_row).Value = percentChange
       ws.Range("K" & yearly_change_row).NumberFormat = "0.00%"
       
      
       'print volume
       ws.Range("L" & yearly_change_row).Value = totalVolume

       'set to move to next row
       yearly_change_row = yearly_change_row + 1
       
       'reset to look at new set of stocks
       changePrice = 0
              
       'reset OpenPrice and increase the row because we want open price for next stock
       openPrice = Cells(i + 1, 3).Value
       
       'figure maximum percentage
       If (percentChange > MaxChange) Then
         MaxChange = percentChange
         MaxTickerName = ticker
       End If
       
       'figures minimum percentage
       If (percentChange < MinChange) Then
         MinChange = percentChange
         MinTickerName = ticker
       End If
       
       'figure out max volume amount
       If (totalVolume > maxTotalVolume) Then
         maxTotalVolume = totalVolume
         maxVolname = ticker
       End If
       
        percentChange = 0
        totalVolume = 0
       
       Else
         totalVolume = totalVolume + Cells(i, 7).Value
      End If
       
   Next i
      'prints all the max/min values to new columns
        ws.Cells(2, 16).Value = MaxTickerName
        ws.Cells(2, 17).Value = MaxChange
        
        ws.Cells(3, 16).Value = MinTickerName
        ws.Cells(3, 17).Value = MinChange
        
        ws.Cells(4, 16).Value = maxVolname
        ws.Cells(4, 17).Value = maxTotalVolume
        
Next ws

End Sub


