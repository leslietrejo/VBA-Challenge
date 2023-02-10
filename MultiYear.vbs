Attribute VB_Name = "Module1"

Sub multiYear():


' Part 1
    ' create script that loops through all stocks for one year
    ' track changes in ticker from column A
    ' track yearly change from opening price in column C to closing price in column F
    ' calculate percentage change from opening price in column C to closing price in column F
    ' calculate total stock volume from column G based on changes in column A
    
    ' each time the ticker changes in colum A,
    ' populate the name of the ticker in column I
    ' display the total stock volume in column L
    ' reset the total and start tracking the next ticker
    
    ' delcare variable to hold worksheet
    Dim ws As Worksheet
    
    ' loop throught all worksheets
    For Each ws In Worksheets
    
    ' declare variable for last row
    Dim lastRow As Long

    ' find last row in the sheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        ' check ticker names
        Dim tickName As String
        
        ' variable to hold the total
        Dim stockTotal As Double
        stockTotal = 0
        
        ' variable to hold the rows in total columns
        Dim stockRows As Long
        stockRows = 2
        
        ' declare variable to hold the row
        Dim row As Long
        
        ' declare variable to hold first open value row
        Dim openValue As Long
        openValue = 2
        
        ' declare variable to hold first value in column C
        Dim firstValue As Long
        
        ' declare variable to find
        Dim findValue As Long
        
        
        ' declare variable the holds yearly change
        Dim yearlyChange As Double
        yearlyChange = 0
        
      
        
        ' declare variable to hold percent change
        Dim percentChange As Double
        percentChange = 0
        
         
       
        
        ' loop through the rows and check the changes in ticker
        For row = 2 To lastRow

        
            ' check the changes in the tickers
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                    
                    ' if the ticker name changes the show the change
                    'MsgBox (Cells(row, 1).Value + " -> " + Cells(row + 1, 1).Value)
             
                    ' set the ticker name
                    tickName = ws.Cells(row, 1).Value
                    
                    ' add to stock volume total
                    stockTotal = stockTotal + ws.Cells(row, 7).Value
                    
                    ' display the ticker name on column I
                    ws.Cells(stockRows, 9).Value = tickName
                    
                    ' display the total stock volume on column L
                    ws.Cells(stockRows, 12).Value = stockTotal
                    
                    ' add 1 to the ticker row to keep moving
                    stockRows = stockRows + 1
                  
                ' check for zero values before continuing
                If stockTotal = 0 Then
               
                    '  display the yearly change on column J
                    ws.Cells(stockRows, 10).Value = yearlyChange
                
                    ' display the formated percent change on column K
                    ws.Cells(stockRows, 11).Value = percentChange
                   ws.Cells(stockRows, 11).NumberFormat = "0.00%"
                   
                   
            Else
        
            ' find first open
                If ws.Cells(stockRows, 3).Value = 0 Then
                        For firstValue = stockRows To row
                        
                        ' loop through data to see if next value is a zero
                        If ws.Cells(firstValue, 3).Value <> 0 Then
                        stockRows = findValue
                        
                        ' leave loop when value is not a zero
                        Exit For
                        
                        End If
                        
                        Next firstValue
                
                 End If
             
                    ' calculate the yearly changes
                    yearlyChange = ws.Cells(row, 6).Value - ws.Cells(stockRows, 3).Value
                  
                    ' calculate percent changes based on yearly changes
                    percentageChange = yearlyChange / ws.Cells(stockRows, 3).Value
                    
                    ' displays the yearly changes in column J
                    ws.Cells(stockRows, 10).Value = yearlyChange
                    
                    ' displays the percent changes in column K
                    ws.Cells(stockRows, 11).Value = percentageChange
                    
                    
                    
                        
                        
                        
    'Part II: formating yearly change in column J
        If yearlyChange > 0 Then
            ws.Cells(stockRows, 10).Interior.ColorIndex = 4
        ElseIf yearlyChange < 0 Then ws.Cells(stockRows + 2, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(stockRows, 10).Interior.ColorIndex = 0
        

        End If
                End If
                    
                    ' reset all total values to move to next
                        stockTotal = 0
                        yearlyChange = 0
                        percentChange = 0
                        
                    ' once reset move to next row
                        stockRows = stockRows + 1
                        
                        
                
         Else
         ' if no changes
            stockTotal = stockTotal + ws.Cells(row, 7).Value
            
            
                
                End If
              
        Next row
       
       
    Next ws
    
    End Sub
    
    
    Sub AlphaTesting():


' Part 1
    ' create script that loops through all stocks for one year
    ' track changes in ticker from column A
    ' track yearly change from opening price in column C to closing price in column F
    ' calculate percentage change from opening price in column C to closing price in column F
    ' calculate total stock volume from column G based on changes in column A
    
    ' each time the ticker changes in colum A,
    ' populate the name of the ticker in column I
    ' display the total stock volume in column L
    ' reset the total and start tracking the next ticker
    
    ' delcare variable to hold worksheet
    Dim ws As Worksheet
    
    ' loop throught all worksheets
    For Each ws In Worksheets
    
    ' declare variable for last row
    Dim lastRow As Long

    ' find last row in the sheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        ' check ticker names
        Dim tickName As String
        
        ' variable to hold the total
        Dim stockTotal As Double
        stockTotal = 0
        
        ' variable to hold the rows in total columns
        Dim stockRows As Long
        stockRows = 2
        
        ' declare variable to hold the row
        Dim row As Long
        
        ' declare variable to hold first open value row
        Dim openValue As Long
        openValue = 2
        
        ' declare variable to hold first value in column C
        Dim firstValue As Long
        
        ' declare variable to find
        Dim findValue As Long
        
        
        ' declare variable the holds yearly change
        Dim yearlyChange As Double
        yearlyChange = 0
        
      
        
        ' declare variable to hold percent change
        Dim percentChange As Double
        percentChange = 0
        
         
       
        
        ' loop through the rows and check the changes in ticker
        For row = 2 To lastRow

        
            ' check the changes in the tickers
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                    
                    ' if the ticker name changes the show the change
                    'MsgBox (Cells(row, 1).Value + " -> " + Cells(row + 1, 1).Value)
             
                    ' set the ticker name
                    tickName = ws.Cells(row, 1).Value
                    
                    ' add to stock volume total
                    stockTotal = stockTotal + ws.Cells(row, 7).Value
                    
                    ' display the ticker name on column I
                    ws.Cells(stockRows, 9).Value = tickName
                    
                    ' display the total stock volume on column L
                    ws.Cells(stockRows, 12).Value = stockTotal
                    
                    ' add 1 to the ticker row to keep moving
                    stockRows = stockRows + 1
                  
                ' check for zero values before continuing
                If stockTotal = 0 Then
               
                    '  display the yearly change on column J
                    ws.Cells(stockRows, 10).Value = yearlyChange
                
                    ' display the formated percent change on column K
                    ws.Cells(stockRows, 11).Value = percentChange
                   ws.Cells(stockRows, 11).NumberFormat = "0.00%"
                   
                   
            Else
        
            ' find first open
                If ws.Cells(stockRows, 3).Value = 0 Then
                        For firstValue = stockRows To row
                        
                        ' loop through data to see if next value is a zero
                        If ws.Cells(firstValue, 3).Value <> 0 Then
                        stockRows = findValue
                        
                        ' leave loop when value is not a zero
                        Exit For
                        
                        End If
                        
                        Next firstValue
                
                 End If
             
                    ' calculate the yearly changes
                    yearlyChange = ws.Cells(row, 6).Value - ws.Cells(stockRows, 3).Value
                  
                    ' calculate percent changes based on yearly changes
                    percentageChange = yearlyChange / ws.Cells(stockRows, 3).Value
                    
                    ' displays the yearly changes in column J
                    ws.Cells(stockRows, 10).Value = yearlyChange
                    
                    ' displays the percent changes in column K
                    ws.Cells(stockRows, 11).Value = percentageChange
                    
                    
                    
                        
                        
                        
    'Part II: formating yearly change in column J
        If yearlyChange > 0 Then
            ws.Cells(stockRows, 10).Interior.ColorIndex = 4
        ElseIf yearlyChange < 0 Then ws.Cells(stockRows + 2, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(stockRows, 10).Interior.ColorIndex = 0
        

        End If
                End If
                    
                    ' reset all total values to move to next
                        stockTotal = 0
                        yearlyChange = 0
                        percentChange = 0
                        
                    ' once reset move to next row
                        stockRows = stockRows + 1
                        
                        
                
         Else
         ' if no changes
            stockTotal = stockTotal + ws.Cells(row, 7).Value
            
            
                
                End If
              
        Next row
       
       
    Next ws
    
    End Sub
    
    
    
    
        
        
        
        
        
        
        
        
        
        
        
        
    
    
    
    
    

    

    
        
        
        
        
        
        
        
        
        
        
        
        
    
    
    
    
    

    

