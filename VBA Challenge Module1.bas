Attribute VB_Name = "Module1"
Sub Stocks()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("l1").Value = "Total Stock Volume"
        ws.Range("o1").Value = "Ticker"
        ws.Range("p1").Value = "Value"
        Range("q1").Value = "Greatest Percent Increase"
        Range("q2").Value = "Greatest Percent Decrease"
        Range("q3").Value = "Greatest Total Volume"
    
    
    
    'Initialize Variables
    Dim totalVolume As Double
    Dim rowCount As Long
    Dim percentChange As Long
    Dim yearlyChange As Double
    Dim x As Long
    Dim Start As Long
    
    
    
    
    'count number of rows
    rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    rowEnd = ws.Cells(Rows.Count, "I").End(xlUp).Row
    'Initialize totalVolume as zero
    totalVolume = 0
    j = 0
    Start = 2
    
    'MsgBox (rowCount)
    For i = 2 To rowCount
    'Ticker symbol, total volume, open price, closing price
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ' Print result if tickers changes
            'Calculate total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            'print results if total = 0
            
            
            If totalVolume = 0 Then
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = "%" & 0
                ws.Range("L" & 2 + j).Value = 0
            Else
            'Find First non zero starting value
                If ws.Cells(Start, 3) = 0 Then
                    For k = Start To i
                        If ws.Cells(k, 3).Value <> 0 Then
                            Start = k
                    
                            Exit For
                        End If
                     Next k
                     
                     
                End If
                
             'Calculate Change
                yearlyChange = (ws.Cells(i, 6) - ws.Cells(Start, 3))
                percentChange = Round((yearlyChange / ws.Cells(Start, 3) * 100), 2)
                ' start of the next stock ticker
                Start = i + 1
                ' print the results
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = Round(yearlyChange, 2)
                ws.Range("K" & 2 + j).Value = "%" & percentChange
                ws.Range("L" & 2 + j).Value = totalVolume
                
            End If
            ' reset variables for new stock ticker
            totalVolume = 0
            yearlyChange = 0
            j = j + 1
            Days = 0
                
                
            
        Else
            'Calculate total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        
        End If
    Next i
    
    'Apply Conditional formatting to yearlyChange
    For i = 2 To rowEnd
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.Color = vbGreen
        Else
            ws.Cells(i, 10).Interior.Color = vbRed
            
        End If
        
    
    Next i
    
    For i = 2 To rowEnd
        If ws.Cells(i, 11).Value > 0 Then
            ws.Cells(i, 11).Interior.Color = vbGreen
        Else
            ws.Cells(i, 11).Interior.Color = vbRed
            
        End If
    
        
    Next i
    
        
    
    
    
Next ws
    
  
  

End Sub
