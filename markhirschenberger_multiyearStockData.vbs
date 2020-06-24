Sub multiyearStockData()
Dim ws As Worksheet

    For Each ws In Worksheets
           Dim WorksheetName As String
           
           'Finding the Last Row
           LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
           
          
           
           
           
           WorksheetName = ws.Name
           Dim ticker As String
           Dim yearlyChange As Double
           Dim yearOpen As Double
           Dim yearClose As Double
           Dim percentChange As Double
           Dim stockVol As Double
           Dim printIndex As Integer
           
           
           
           'Define the extra box value
           
           Dim greatestPerIncrease As Double
           Dim greatestPerDecrease As Double
           Dim greatestTotalVolume As Double
           
        
           
           printIndex = 2
        
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
            Cells(1, 16).Value = "Ticker"
            Cells(1, 17).Value = "Value"
            Cells(2, 15).Value = "Greatest % Increase"
            Cells(3, 15).Value = "Greatest % Decrease"
            Cells(4, 15).Value = "Greatest Total Volume"
            Range("Q2:Q3").Style = "Percent"
            
           
           
           
            
            For i = 2 To LastRow
                If i = 2 Then
                    ticker = Cells(i, 1).Value
                    yearOpen = Cells(i, 3).Value
                    stockVol = stockVol + Cells(i, 7).Value
                    
             
                    
                
                  ElseIf Cells(i + 1, 1).Value <> ticker And Cells(i + 1, 3).Value <> 0 Then
                    yearClose = Cells(i, 6).Value
                    yearlyChange = yearClose - yearOpen
                    
                    
                    
                    percentChange = (yearlyChange / yearOpen)
                    stockVol = stockVol + Cells(i, 7).Value
                    Cells(printIndex, 9).Value = ticker
                    Cells(printIndex, 10).Value = yearlyChange
                    Cells(printIndex, 11).Value = percentChange
                    Cells(printIndex, 11).Style = "Percent"
                    Cells(printIndex, 12).Value = stockVol
                    printIndex = printIndex + 1
                    ticker = Cells(i + 1, 1).Value
                    yearOpen = Cells(i + 1, 3).Value
                    Cells(i, 11).Style = "Percent"
                    
                    
                    
                    stockVol = 0
                    

                    
                  Else
                    stockVol = stockVol + Cells(i, 7).Value
                    Cells(printIndex, 12).Value = stockVol
                       
                    
                  End If
                  


            
            Next i
            
            
            
            
            
        'Conditional Formatting
           
           
           
           Dim rowColors
           Set rowColors = Range("J2:J3500")
        
        
            For Each cell In rowColors
                  
                If cell.Value > 0 Then
                    cell.Interior.ColorIndex = 4
                ElseIf cell.Value < 0 Then
                   cell.Interior.ColorIndex = 3
                Else
                    cell.Interior.ColorIndex = 2
                End If
            Next cell
            
            
            
           
            greatestPerIncrease = WorksheetFunction.Max(Range("K2:K290"))
            greatestPerDecrease = WorksheetFunction.Min(Range("K2:K290"))
            greatestTotalVolume = WorksheetFunction.Max(Range("L2:L290"))
            
            Range("Q2").Value = greatestPerIncrease
            
            Range("Q3").Value = greatestPerDecrease
            Range("Q4").Value = greatestTotalVolume
         
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
            Cells(1, 16).Value = "Ticker"
            Cells(1, 17).Value = "Value"
            Cells(2, 15).Value = "Greatest % Increase"
            Cells(3, 15).Value = "Greatest % Decrease"
            Cells(4, 15).Value = "Greatest Total Volume"
            Range("Q2:Q3").Style = "Percent"
            

                       
    Next ws
        

End Sub
