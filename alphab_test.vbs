Sub percentChanges():
    
    ' open with our necessary variables
    Dim yrOpen As Double
    Dim yrClose As Double
    Dim yrDiff As Double
    
    Dim diff As Double
    
    Dim ticker As String    ' set ticker
    Dim volume As Double   ' running total for volume
    Dim tickerRows As Integer  ' first row to insert ticker names
    
    Dim nameRows As Integer ' set the first row to find yearly change
    
    
    ' reach all worksheets
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    
        ws.Cells(1, 9).Value = "Ticker" ' gives Ticker column appropriate name
    
        ' Yearly Change
        ws.Cells(1, 10).Value = "Yearly Change"
    
        ' Percent Change
        ws.Cells(1, 11).Value = "Percent Change"
    
        ' Total Stock Volume
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        volume = 0 ' start intitial total at zero
        tickerRows = 2
        
        For Row = 2 To LastRow
            
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
                ticker = ws.Cells(Row, 1).Value
                volume = volume + ws.Cells(Row, 7)
                ws.Cells(tickerRows, 9).Value = ticker
                ws.Cells(tickerRows, 12).Value = volume
                tickerRows = tickerRows + 1
                volume = 0
            
            Else
                volume = volume + Cells(Row, 7).Value
            End If
        Next Row
        
        openNum = 2
        Count = 0
        
        For Row = 2 To LastRow
            
            placeHolder1 = ws.Cells(Row, 1).Value
            'set a placeholder for the opening value
            placeHolder2 = ws.Cells(Row + 1, 1).Value
            ' set a placeholder for the secondary value to check if the ticker has changed
            
            If placeHolder1 <> placeHolder2 Then
                yrClose = ws.Cells(Row, 6).Value
                ' if the tickers are different, then set the close value and find the difference
                
                yrDiff = yrClose - yrOpen
                ws.Cells(openNum, 10) = yrDiff
                
                If yrDiff > 0 Then
                    ws.Cells(openNum, 10).Interior.ColorIndex = 4
                    
                ElseIf yrDiff < 0 Then
                    ws.Cells(openNum, 10).Interior.ColorIndex = 3
                    
                Else
                    ws.Cells(openNum, 10).Interior.ColorIndex = 6
                End If
                ' set the percent change
                diff = (yrDiff / yrOpen)
                ws.Cells(openNum, 11) = FormatPercent(diff)
                
                openNum = openNum + 1
                
                Count = 0
            
            Else
                yrOpen = ws.Cells(Row - Count, 3).Value
                Count = Count + 1
            End If
            
            
        Next Row


        Dim max As String
        Dim min As String
        Dim titleVol As String
        
        max = "Greatest % Increase"
        min = "Greatest % Decrease"
        titleVol = "Greatest Total Volume"
        
        
        Dim begin As Long
        
        For shts = 1 To Sheets.Count
            maxVal = WorksheetFunction.max(Sheets(shts).Columns("k"))
            minVal = WorksheetFunction.min(Sheets(shts).Columns("k"))
            greatVol = WorksheetFunction.max(Sheets(shts).Columns("l"))
            
            For begin = 2 To LastRow
        
                If maxVal = ws.Cells(begin, 11).Value Then
                    ws.Cells(2, 16).Value = ws.Cells(begin, 9).Value
                    ' set the title
                    ws.Cells(2, 15).Value = max
                    ' format it to a percent
                    ws.Cells(2, 17).Value = FormatPercent(maxVal)
                
                ElseIf minVal = ws.Cells(begin, 11).Value Then
                    ws.Cells(3, 16).Value = ws.Cells(begin, 9).Value
                    ' set the title
                    ws.Cells(3, 15).Value = min
                    ' format to a percent
                    ws.Cells(3, 17).Value = FormatPercent(minVal)
                
                ElseIf greatVol = ws.Cells(begin, 12).Value Then
                    ws.Cells(4, 16).Value = ws.Cells(begin, 9).Value
                    ' set the title
                    ws.Cells(4, 15).Value = titleVol
                    ws.Cells(4, 17).Value = greatVol
                
                End If
            
            Next begin
            
        Next shts
        
    Next ws
    
End Sub
