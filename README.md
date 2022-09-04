# stock-analysis
# Overview of the project:

 Our client, is a financial advisor, and his clients asked him reccomendations about green energy stocks. To create an investment strategy, our client decided to analyze and compare total daily volume and yearly return of each stock. That's why I created a VBA macro to automate the analyses with an easiness of pressing a button.

Using the green_stocks dataset we can refactor a Microsoft Excel VBA code to collect certain stoc information for the year 2017 and 2018 and determine which stocks had a positive yearly return and how active each stock was traded.
 
Daily volume is equal to the total number of shares traded throughout the day; it measures how actively a stock is traded. 
The yearly return is the percentage difference in price from the beginning of the year to the end of the year.

*The following steps applied to create the VBA*

1. Create a worksheet to hold the data. Adding a header and assiging cell row values.
2. Calculate the total daily volume using loops, conditionals and code pattern.
3. Calculate the yearly return of stocks by determining the first closing price and the last closing price.
4. Format the output sheet to make it easier to visualize.
5. Repurpose the VBA macros to analyze multiple stocks.

# Purpose of the project:

To analyze the provided stocks based on 2017 - 2018 data. 

Applied code:

Sub AllStocksAnalysis()

Dim startTime As Single
    Dim endTime  As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer

'1).Format the output sheet on the "All Stocks Analysis" worksheet.

Worksheets("All Stocks Analysis").Activate


Range("A1").Value = "All Stocks (2018)"


    'Create aHeader Row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
   

'2).Initialize an array of all tickers.

    Dim Stringtext

   Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
'3).Prepare for the analysis of tickers.


    '3a).Initialize variables for the starting price and ending price.
        Dim startingprice As Double
        Dim endingPrice As Double
        
        '3b).Activate the data worksheet.


Worksheets("2018").Activate


            '3c).Find the number of rows to loop over.
                RowCount = Cells(Rows.Count, "A").End(xlUp).Row
                

'4).Loop through the tickers.
    For i = 0 To 11
            
            ticker = tickers(i)
            totalVolume = 0

    '5).Loop through rows in the data.
Worksheets("2018").Activate
    
    For J = 2 To RowCount
    
    
        '5a).Find the total volume for the current ticker.
        If Cells(J, 1).Value = ticker Then
        totalVolume = totalVolume + Cells(J, 8).Value
        End If
        
            '5b).Find the starting price for the current ticker.
             If Cells(J - 1, 1).Value <> ticker And Cells(J, 1).Value = ticker Then
             
                startingprice = Cells(J, 6).Value
                End If
                
                '5c).Find the ending price for the current ticker.
                If Cells(J + 1, 1).Value <> ticker And Cells(J, 1).Value = ticker Then
                
                        endingPrice = Cells(J, 6).Value
                        
                        End If
                        Next J
                        
                
'6).Output the data for the current ticker
Worksheets("All Stocks Analysis").Activate
Cells(4 + i, 1).Value = ticker
Cells(4 + i, 2).Value = totalVolume
Cells(4 + i, 3).Value = endingPrice / startingprice - 1


    Next i
       endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    


End Sub

   Sub ClearWorksheet()

Cells.Clear

End Sub
    

Sub formatAllStocksAnalysisTable()

'Formatting

Worksheets("All Stocks Analysis").Activate
    
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$#,##0"
    Range("C4:C15").NumberFormat = "0,0%"
    
    Columns("B").AutoFit
    
        If Cells(4, 3) > 0 Then
'Color the cell green
            Cells(4, 3).Interior.Color = vbGreen
         
         ElseIf Cells(4, 3) < 0 Then
            'Color the cell red
                Cells(4, 3).Interior.Color = vbRed
        Else
            'Clear the cell color
            Cells(4, 3).Interior.Color = xlNone
            
            End If
            
            
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
    
            If Cells(i, 3) > 0 Then
                'Change cell color to green
                Cells(i, 3).Interior.Color = vbGreen
            ElseIf Cells(i, 3) < 0 Then
            
                    'Change cell clor to red
                    Cells(i, 3).Interior.Color = vbRed
            Else
            
                    'Clear the cell color
                    Cells(i, 3).Interior.Color = xlNone
                    
            End If
            
            
            
                
    Next i
    
    
       
        
End Sub
Sub yearValueAnalysis()

Dim startTime As Single
Dim endTime As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
 startTime = Timer

 
'1).Format the output sheet on the "All Stocks Analysis" worksheet.

Worksheets("All Stocks Analysis").Activate

 Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create aHeader Row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
   

'2).Initialize an array of all tickers.

    Dim Stringtext

   Dim tickers(11) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

    
'3).Prepare for the analysis of tickers.


    '3a).Initialize variables for the starting price and ending price.
        Dim startingprice As Double
        Dim endingPrice As Double
        
        '3b).Activate the data worksheet.


Worksheets(yearValue).Activate


            '3c).Find the number of rows to loop over.
                RowCount = Cells(Rows.Count, "A").End(xlUp).Row
                

'4).Loop through the tickers.
    For i = 0 To 11
            
            ticker = tickers(i)
            totalVolume = 0

    '5).Loop through rows in the data.

    Worksheets(yearValue).Activate
    
    For J = 2 To RowCount
    
    
        '5a).Find the total volume for the current ticker.
        If Cells(J, 1).Value = ticker Then
        totalVolume = totalVolume + Cells(J, 8).Value
        End If
        
            '5b).Find the starting price for the current ticker.
             If Cells(J - 1, 1).Value <> ticker And Cells(J, 1).Value = ticker Then
             
                startingprice = Cells(J, 6).Value
                End If
                
                '5c).Find the ending price for the current ticker.
                If Cells(J + 1, 1).Value <> ticker And Cells(J, 1).Value = ticker Then
                
                        endingPrice = Cells(J, 6).Value
                        
                        End If
                        Next J
                        
                
'6).Output the data for the current ticker
Worksheets("All Stocks Analysis").Activate
Cells(4 + i, 1).Value = ticker
Cells(4 + i, 2).Value = totalVolume
Cells(4 + i, 3).Value = endingPrice / startingprice - 1


    Next i
       endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    


End Sub



Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
        tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
''2a) Create a for loop to initialize the tickerVolumes to zero.
  For i = 0 To 11
        
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
  
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
     
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
             
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowÕs ticker doesnÕt match, increase the tickerIndex.
        'If  Then
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                        
            End If
            

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
                End If
                
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        Worksheets("All Stocks Analysis").Activate
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
# Results

10 out of 12 chosen stocks performed negative yearly return. Only "ENPH" and "RUN" performed positive yearly return. Especially, "RUN" performed the highest yearly return compared to the rest of the 11 stocks. The yearly return of "ENPH" changed from 129.5% to 81.9%, althoug the yearly return decreased over a year, but still performed positive. "RUN" stocks changed from 5.5% to 84% which is a mjore increase. 
 
By looking at the Microsoft Excel sheet " All Stock Analysis" we can conclude that in the year 2017 out of the 12 tickers analyzed only one "TERP" had a negative yearly return. While the other 11 tickers yearly return percent change varies between 8.9% to 199.4%. While during the year 2018, 10 out of the 12 tickers show a negative yearly returns, with only "ENPH" and "RUN" showing a positive percent change of 81% and 84% respectively. It is important to note that the by refactoring the VBA script total macro run time decreased by approximately 0.47 seconds.

*Changes after Refactoring:*
Aftre Refactoring, VBA running time changed significantly. Before the Refactoring, the analysis of 2017 took around 1.6 second, and after Refactoring, the same analysis took 0.26 second (Please see below)![VBA_Challenge_2017](https://user-images.githubusercontent.com/109055148/188336960-c2feb1d9-ae5c-4201-9337-49245e0dabc0.png)
