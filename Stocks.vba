Sub stock():
'Initializing all the requirement variables
    Dim increment As Integer
    Dim volumn As Variant
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatestvolumn As Variant
    Dim greaterticker As String
    Dim smallerticker As String
    Dim volumnticker As String
    
'starting for each loop to iterate all the sheet in the spredsheet
    For Each ws In Worksheets
        'Getting the row row number
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Declaring default value for variables
        volumn = 0
        greatestincrease = -1.8E+307 'I am declaring the lowest double value.
        greatestdecrease = 1.8E+307
        greatestvolumn = 0
        increment = 2
        openprice = 0
        
        'Setting all the column names
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volumn"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Starting a loop to iterate each row in the sheet
        For i = 2 To lastrow
        
            'This conditon is to get the open price for first of each ticker and not for the repeated tickers.
            If openprice = 0 Then
            
                openprice = ws.Cells(i, 3).Value
            End If
            
            'This If condition will iterate to each ticker untill it finds the new ticker in the next cell
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                volumn = volumn + ws.Cells(i, 7).Value
            
            'If the next ticker is new, it will excute the else block
            Else
                volumn = volumn + ws.Cells(i, 7).Value
                'Updating the ticker and total volumn
                ws.Cells(increment, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(increment, 12).Value = volumn
                closeprice = ws.Cells(i, 6).Value
                
                
                'Yearly change from what the stock opened the year at to what the closing price was.
                
                yearlychange = closeprice - openprice
                ws.Cells(increment, 10).NumberFormat = "0.000000000"
                ws.Cells(increment, 10).Value = yearlychange
                
                'If yearly change is in positive number it will change the cell color to green or it will change to red.
                If yearlychange > 0 Then
                    ws.Cells(increment, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(increment, 10).Interior.ColorIndex = 3
                End If
                
                
                
                'Percentage difference between two numbers'
                'To avoid divide by zero error i make sure the open price is not zero
                If openprice <> 0 Then
                    
                    percentchange = ((closeprice - openprice) / openprice)
                Else
                    percentchange = 0
                End If
                
                'Setting the cell format to "Pencentage and updating the percentage change
                ws.Cells(increment, 11).NumberFormat = "0.00%"
                ws.Cells(increment, 11).Value = percentchange
                
                'This will check whether is the greatest percentage change or not
                If percentchange > greatestincrease Then
                    
                    greatestincrease = percentchange
                    greaterticker = ws.Cells(increment, 9).Value
                End If
                
                'This will check whether is the lowest percentage change or not
                If percentchange < greatestdecrease Then
                    
                    greatestdecrease = percentchange
                    smallerticker = ws.Cells(increment, 9).Value
                End If
                
                'This will check whether is the greatest volumn or not
                If volumn > greatestvolumn Then
                    greatestvolumn = volumn
                    volumnticker = ws.Cells(increment, 9).Value
                End If
                
                'inrementing the value by 1 for next ticker
                increment = increment + 1
                
                'setting the variables to 0 for next ticker
                volumn = 0
                openprice = 0
                
            End If
        Next i
        
        'The below code will at the end of each sheet and update the greatest increase, decrease and volumn.
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greaterticker
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 17).Value = greatestincrease
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = smallerticker
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = greatestdecrease
        
        ws.Cells(4, 15).Value = "Greatest Total Volumn"
        ws.Cells(4, 16).Value = volumnticker
        ws.Cells(4, 17).Value = greatestvolumn
    Next ws
End Sub



