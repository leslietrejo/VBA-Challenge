Attribute VB_Name = "Module1"
Sub biggerChallengeFile():
    
    ' open with our necessary variables
    Dim yearOpen As Double   ' set the opening cells
    Dim yearClose As Double   ' set the closing cells
    Dim yearDiff As Double
    
    Dim percentChange As Double  ' set the percent change values
    Dim percentDiff As Double
    
    Dim tickName As String    ' set the ticker name
    Dim volTotal As Double   ' set a running total for volume
    Dim tickRows As Integer  ' set the first row to insert ticker names
    
    Dim nameRows As Integer ' set the first row to find yearly change
    
    
    ' reach all worksheets
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    
        
        
        ' now we need to insert all of our new rows
        ' Ticker
        ws.Cells(1, 9).Value = "Ticker" ' gives Ticker column appropriate name
    
        ' Yearly Change
        ws.Cells(1, 10).Value = "Yearly Change"
    
        ' Percent Change
        ws.Cells(1, 11).Value = "Percent Change"
    
        ' Total Stock Volume
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' now lets start to find our values
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        volTotal = 0 ' start intitial total at zero
        tickRows = 2
        
        For Row = 2 To LastRow
            
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
                tickName = ws.Cells(Row, 1).Value
                volTotal = volTotal + ws.Cells(Row, 7)
                ws.Cells(tickRows, 9).Value = tickName
                ws.Cells(tickRows, 12).Value = volTotal
                tickRows = tickRows + 1
                volTotal = 0
            
            Else
                volTotal = volTotal + Cells(Row, 7).Value
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
                yearClose = ws.Cells(Row, 6).Value
                ' if the tickers are different, then set the close value and find the difference
                
                yearDiff = yearClose - yearOpen
                ws.Cells(openNum, 10) = yearDiff
                
                If yearDiff > 0 Then
                    ws.Cells(openNum, 10).Interior.ColorIndex = 4
                    
                ElseIf yearDiff < 0 Then
                    ws.Cells(openNum, 10).Interior.ColorIndex = 3
                    
                Else
                    ws.Cells(openNum, 10).Interior.ColorIndex = 6
                End If
                ' set the percent change
                percentDiff = (yearDiff / yearOpen)
                ws.Cells(openNum, 11) = FormatPercent(percentDiff)
                
                ' add one more to the row so we keep going to the next ticker
                openNum = openNum + 1
                
                Count = 0
                ' if the placeholders aren't equal, use the count variable to reset the opening value
            
            Else
                ' if the placeholders are equal, keep the loop moving with the same opening value
                yearOpen = ws.Cells(Row - Count, 3).Value
                Count = Count + 1
                ' if the placeholders are equal, use the count variable to keep the opening value the same
            End If
            
            
        Next Row
        'set variables to make titles for calculated variables
        ' set variables for the true values of the variables
        Dim titleMax As String
        Dim titleMin As String
        Dim titleVol As String
        
        titleMax = "Greatest % Increase"
        titleMin = "Greatest % Decrease"
        titleVol = "Greatest Total Volume"
        
        
        
        ' run a loop to show the ticker values for the mins/maxs
        ' begin will serve as our starting row
        Dim begin As Long
        
        ' use a nested if loop to check which ticker contains max/min value
        ' if it does, the ticker will be stored in the specified cell
        ' if it doesn't, the loop will continue checking the next row
        ' for loop is needed to repeat the ticker - checker
        
        For shts = 1 To Sheets.Count
            ' using the max/min functions I found on WallStreetMojo.com
            maxVal = WorksheetFunction.Max(Sheets(shts).Columns("k"))
            minVal = WorksheetFunction.Min(Sheets(shts).Columns("k"))
            greatVol = WorksheetFunction.Max(Sheets(shts).Columns("l"))
            
            For begin = 2 To LastRow
        
                If maxVal = ws.Cells(begin, 11).Value Then
                    ws.Cells(2, 16).Value = ws.Cells(begin, 9).Value
                    ' set the title
                    ws.Cells(2, 15).Value = titleMax
                    ' format it to a percent
                    ws.Cells(2, 17).Value = FormatPercent(maxVal)
                
                ElseIf minVal = ws.Cells(begin, 11).Value Then
                    ws.Cells(3, 16).Value = ws.Cells(begin, 9).Value
                    ' set the title
                    ws.Cells(3, 15).Value = titleMin
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
