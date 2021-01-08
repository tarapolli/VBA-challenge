Sub NYSE()

'loops thru all worksheets
Dim sheet_counter As Integer
Dim Sheet_name As String

'count each sheet in the workbook
 sheet_counter = ThisWorkbook.Sheets.Count

 For j = 1 To sheet_counter
         ThisWorkbook.Sheets(j).Select
         Sheet_name = ThisWorkbook.Sheets(j).Name
        
    Dim close1 As Double

    Dim StockVolume As LongLong    ' set initial variable for holding the volume total 
    StockVolume = 0

    Dim YearChg As Double

    Dim PercentChg As Double

    Dim TickerCounter As LongLong
    TickerCounter = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row   'Count the number of rows  

    'this prevents an overflow error when open price is 0
    On Error Resume Next

    'Initialize Opening price as C2 ( Special case )
    Dim open1 As Double
    open1 = Cells(2, 3).Value

    'Initialize Ticker symbol of starting ticker ( Special case )
    Dim ticker As String
    ticker = Cells(2, 1).Value
 
        'Looping through all the rows:
        For i = 2 To LastRow

            'Keep adding the volumes for every rowm(stockvolume = stockvolume +1)
            StockVolume = StockVolume + Cells(i + 1, 7).Value
    
            'condition to check if ticker is different, then print this ticker's info
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
                'Extract the closing price
                close1 = Cells(i, 6).Value

                'Calculate the yearly change
                YearChg = (Cells(i, 6).Value - Cells(i, 3).Value)
        
                'Calculatee the yearly percentage
                PercentChg = (close1 - open1) / (open1)
        
                Range("K" & TickerCounter).NumberFormat = "0.00%"
        
                'Print out the Ticker Symbol
                Cells(TickerCounter, 9).Value = Cells(i, 1).Value
        
                'Print the Stock Volume
                'Cells(2, 12).Value = StockVolume
                Cells(TickerCounter, 12).Value = StockVolume
                    
                'Print yearly Change
                'Cells(2, 10).Value = YearChg
                'Cells(i, 10).Value = YearChg
                'Range("J" & TickerCounter).Value = YearChg
                'Cells(TickerCounter, 10).Value = (Cells(i, 6).Value - Cells(TickerCounter, 3).Value)
                Cells(TickerCounter, 10).Value = close1 - open1
                 'YearChg = close1 - open1
                     
                        'colors positives green and negatives red  Kevin#2
                        Dim num As Double
                        num = close1 - open1
                        
                        Select Case num
                            Case Is >= 0
                                Range("J" & TickerCounter).Interior.ColorIndex = 4
                            Case Else
                                Range("J" & TickerCounter).Interior.ColorIndex = 3
                        End Select
                    
                'print percent change
                Cells(TickerCounter, 11).Value = PercentChg
          
                'PRINT OPEN PRICE
                Cells(TickerCounter, 14).Value = open1
                      
                'PRINT CLOSE PRICE
                Cells(TickerCounter, 15).Value = Cells(i, 6).Value

                'Reset the volume
                StockVolume = 0
            
                'Reset the opening price to the next tickers opening price
                open1 = Cells(i + 1, 3).Value
        
                'Add the TickerCounter_rindex + 1
                TickerCounter = TickerCounter + 1
        
            End If
                 
        Next i
     
 Next j

End Sub


