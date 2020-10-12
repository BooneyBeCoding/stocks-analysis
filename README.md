# VBA Stocks-Analysis

## Overview: 
  Using VBA to create Macros to read and execute a review of 2017 & 2018 stock data.

### Results: 
  While 2017 saw a larger return for most of the stocks analyized, 2018 only had 2 stock with a postivie return in ENPH & RUN.  From the data analysed, it would seem those are the   2 to invest in based on prior returns.
  By utilizing the provided dada, creating ticker values for each stock, and looping the macro through we were able to create a table with Total Daily Volume & Return for wach year designated:
     
     Code Example:
      '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
            '5a) Get total volume for current ticker
                If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
                
                End If
                
            '5b) get starting price for current ticker
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                    startingPrice = Cells(j, 6).Value
                    
                End If
    
            '5c) get ending price for current ticker
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                endingPrice = Cells(j, 6).Value
                End If

       Next j 
  
 #### Results of Macro:  
     2017 Results:
   ![image](https://user-images.githubusercontent.com/71455991/95695696-f0d42b00-0bfd-11eb-8f8f-ba3c3f7fac02.png)

    2018 Results
   ![image](https://user-images.githubusercontent.com/71455991/95695748-0cd7cc80-0bfe-11eb-9e7d-8d51653b73a0.png)

### Summary:
While the macro worked as intended, as always it could improved.  By creating a ticker index, we were able to not only make the macro run much faster, but also make it so that it can be continued with more years of data added.  

Refacoring allows us to make sure the code is more efficiant (first attempts are rarely efficiant).  In terms of this, the original code was limited to the specific tickers of the data.  Now, with the ticker index, we are able to keep the macro going if more are added to analyse in 2017 & 2018 as well as other years.

    Code Example:
    '1a) Create a ticker Index
    ticketIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
    End If
    
    Next i
    
#### Results of Refactored Code:
    2017 Results
   ![image](https://user-images.githubusercontent.com/71455991/95696232-efa3fd80-0bff-11eb-86dd-7f572ce3ebe2.png)

    2018 Results
  ![image](https://user-images.githubusercontent.com/71455991/95696253-021e3700-0c00-11eb-9935-934a107ffdb0.png)
 
