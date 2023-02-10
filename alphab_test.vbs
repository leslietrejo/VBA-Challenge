Attribute VB_Name = "Module1"
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
    
    
        ' find last row in the sheet
        lastRow = Cells(Rows.Count, 1).End(xlUp).row
        
        ' check ticker names
        Dim tickName As String
        
        ' variable to hold the total
        Dim stockTotal As Double
        stockTotal = 0
        
        ' variable to hold the rows in total columns (column I and L)
        Dim stockRows As Integer
        stockRows = 2
        
        ' declare variable to hold the row
        Dim row As Integer
        
        ' loop through the rows ans check the changes in ticker
        For row = 2 To lastRow
        
            ' check the changes in the tickers
            If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
                    
                    ' if the ticker name changes the show the change
                    'MsgBox (Cells(row, 1).Value + " -> " + Cells(row + 1, 1).Value)
                    
                    ' set the ticker name
                    tickName = Cells(row, 1).Value
                    
                    ' add to stock volume total
                    stockTotal = stockTotal + Cells(row, 7).Value
                    
                    ' display the ticker name on column I
                    Cells(stockRows, 9).Value = tickName
                    
                    ' display the total stock volume on column L
                    Cells(stockRows, 12).Value = stockTotal
                    
                    ' add 1 to the ticker row to keep moving
                    stockRows = stockRows + 1
                    
                    ' reset the total stock volume for the next ticker
                    stockTotal = 0
                    
                Else
                        
                        ' If there is no change in the ticker then keep adding to the total stock volume
                        stockTotal = stockTotal + Cells(row, 7).Value
                
                    End If
                    
        Next row
       
            
                
    End Sub
    
