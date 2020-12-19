Sub ticker_data():

' create column labels

Cells(1, 10).Value = "ticker"
Cells(1, 11).Value = "yearly change"
Cells(1, 12).Value = "percent change"
Cells(1, 13).Value = "total stock volume"

total_volume = 0
open_price = 0
close_price = 0

' variable to grab ticker label in column 1

    ticker = Cells(2, 1).Value
    
    ' MsgBox (ticker)
    
' variable for opening price of stock

    open_price = Cells(2, 3).Value
    
    ' MsgBox (open_price)
    
    output = 2
    
' For Loop to loop through all rows
    
    For Row = 2 To 70926
    
    total_volume = total_volume + Cells(Row, 7).Value
    
        ' If statement: If the next row changes tickers then grab the closing price for that ticker
        
        If Cells(Row + 1, 1) <> ticker Then
        
             close_price = Cells(Row, 6).Value
        
           ' MsgBox = (close_price)
        
             ' formula to calculate the change in close - open
        
               Change = close_price - open_price
        
                 open_price = Cells(Row + 1, 3)
            
                 Cells(output, 11).Value = Change
                 Cells(output, 10).Value = Cells(Row, 1).Value
                 

                 ' how do i get the volume to sum
                 
                 Cells(output, 13).Value = total_volume
                 
                 total_volume = 0
            
    
                 ticker = Cells(Row + 1, 1).Value
        
                ' formula to calculate percent change
                
                open_price = Cells(Row, 3).Value
                
                percent_change = (close_price - open_price) / close_price
                
                Cells(output, 12).Value = percent_change
                
                
           output = output + 1


        End If
Next Row

' for loop for conditional formatting

For Color = 2 To 290

    If Cells(Color, 11).Value < 0 Then
    
   Cells(Color, 11).Interior.Color = RGB(255, 0, 0)
   
   Else:
    
    Cells(Color, 11).Interior.Color = RGB(0, 255, 0)
   
   End If
   
   Next Color


End Sub
