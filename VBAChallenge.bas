Attribute VB_Name = "Module1"
Sub biggerChallengeFile():
    
    Dim yearOpen As Double   
    Dim yearClose As Double  
    Dim yearDiff As Double
    
    Dim percentChange As Double 
    Dim percentDiff As Double
    
    Dim tickName As String   
    Dim volTotal As Double  
    Dim tickRows As Integer
    
    Dim nameRows As Integer
    
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    
        
        ws.Cells(1, 9).Value = "Ticker" ' gives Ticker column appropriate name
    
        ws.Cells(1, 10).Value = "Yearly Change"
    
        ws.Cells(1, 11).Value = "Percent Change"
    
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        volTotal = 0 
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
            
            placeHolder2 = ws.Cells(Row + 1, 1).Value
            
            If placeHolder1 <> placeHolder2 Then
                yearClose = ws.Cells(Row, 6).Value
                
                yearDiff = yearClose - yearOpen
                ws.Cells(openNum, 10) = yearDiff
                
                If yearDiff > 0 Then
                    ws.Cells(openNum, 10).Interior.ColorIndex = 4
                    
                ElseIf yearDiff < 0 Then
                    ws.Cells(openNum, 10).Interior.ColorIndex = 3
                    
                Else
                    ws.Cells(openNum, 10).Interior.ColorIndex = 6
                End If

                percentDiff = (yearDiff / yearOpen)
                ws.Cells(openNum, 11) = FormatPercent(percentDiff)
                
                openNum = openNum + 1
                
                Count = 0
            
            Else
                yearOpen = ws.Cells(Row - Count, 3).Value
                Count = Count + 1
            End If
            
            
        Next Row

        Dim titleMax As String
        Dim titleMin As String
        Dim titleVol As String
        
        titleMax = "Greatest % Increase"
        titleMin = "Greatest % Decrease"
        titleVol = "Greatest Total Volume"
        
        Dim begin As Long
        
        For shts = 1 To Sheets.Count
            maxVal = WorksheetFunction.Max(Sheets(shts).Columns("k"))
            minVal = WorksheetFunction.Min(Sheets(shts).Columns("k"))
            greatVol = WorksheetFunction.Max(Sheets(shts).Columns("l"))
            
            For begin = 2 To LastRow
        
                If maxVal = ws.Cells(begin, 11).Value Then
                    ws.Cells(2, 16).Value = ws.Cells(begin, 9).Value
                    
                    ws.Cells(2, 15).Value = titleMax
                    
                    ws.Cells(2, 17).Value = FormatPercent(maxVal)
                
                ElseIf minVal = ws.Cells(begin, 11).Value Then
                    ws.Cells(3, 16).Value = ws.Cells(begin, 9).Value
                    
                    ws.Cells(3, 15).Value = titleMin
                    
                    ws.Cells(3, 17).Value = FormatPercent(minVal)
                
                ElseIf greatVol = ws.Cells(begin, 12).Value Then
                    ws.Cells(4, 16).Value = ws.Cells(begin, 9).Value
                    
                    ws.Cells(4, 15).Value = titleVol
                    ws.Cells(4, 17).Value = greatVol
                
                End If
            
            Next begin
            
        Next shts
        
    Next ws
    
End Sub
