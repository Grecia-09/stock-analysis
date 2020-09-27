# An analysis of Stocks

## Overview of the Project.

### *Help Steve to look into different green energy stocks to analyze them and be able to provide with better information and insights to his parents so they can make the best investment decision. Steve had created an automated analysis in VBA but I wanted to help Steve to go through all the data from each stock by reducing the chance of errors and the time of running his current analysis.*
---
## Results.
---
First, I wanted to focus on the Daily Volume to see how actively a stock is traded throughout the day, and the yearly return of every stock since it will show us the percentage difference in prices from the beginning to end of the year. I created a worksheet where I am going to output my analysis from both 2017 and 2018, called “All Stock Analysis”.

<img width="413" alt="VBA_Challenge_code" src="https://user-images.githubusercontent.com/70611325/94376308-a10a4580-00ce-11eb-9ecd-a34355517549.png">

---
To find the Total Daily Volume for each stock I used a for loop to go over every row of the data sheet, but before doing that I needed to create an array to hold every stock as variables and named it  “Tickers”:

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


---
Once defined all tickers, I created a variable named TickerIndex to return the reference for the entire column or row, respectively, and set it equal to zero. This will allow me to add to it inside the loop in order to hold the sum of all three output arrays created next: 

    tickerIndex = 0

    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single


To find the Total Daily Volume, I set the TickerVolumes variable equal to zero so it can hold the sum of all the volume: 

     For i = 0 To 11
  
     tickerVolumes(i) = 0
 
     Next i
---
Since Volume is in the 8th column of the spreadsheet, I increased the Ticker Volumes by the value in Cells(i, 8):

    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
---
Next, to find the Yearly Return I searched for the starting prices and ending prices of each ticker. To do so, I checked if the current row is the first row of the ticker and if so, set it to the be the starting price. To set the starting prices and ending prices I used Conditionals, and since the prices are in column 6th, I am interested in using what is in Cells(i, 6) :

    If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then

          tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
          End If
---
Same thing to find the ending prices:

    If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

           tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
           End If
---
After finding both the starting and ending prices for each ticker, if the next row’s ticker doesn’t match the previous row’s ticker, then the TickerIndex will increase:

    If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerIndex = tickerIndex + 1
             
            End If
        
        Next i
---
Finally, the analysis need to reflected in our "All Stocks Analysis" worksheet, for this, I used a new for loop to go over the three arrays I created before and output all results there:

    For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
        
      Next i
---
After running the code, I created to help Steve, we can compare each year returns and see that tickers had a better performance in 2017, since the only one that had a decrease over 7.2% was TERP. In 2018, only ENPH and Run had a positive performance in comparison with the rest of tickers that had dropped, being the highest one DQ over -62.6%.

<img width="245" alt="VBA_Challenge_results2017" src="https://user-images.githubusercontent.com/70611325/94376438-589f5780-00cf-11eb-9095-82229948a330.png">
<img width="242" alt="VBA_Challenge_results2018" src="https://user-images.githubusercontent.com/70611325/94376452-6523b000-00cf-11eb-9f3f-e4b8c3dccaf6.png">

It seems like the best two options for Steve’s parents, after seeing their positive performance in both years, are ENPH or Run.

Another thing to noticed is that before the refactoring of Steve’s code, it took 0.7421875 seconds to execute the original script for 2017 and 1.054688 second to execute the original script for 2018. After refactoring his original code, the execution time for both years dropped significantly. See images below for reference:

<img width="434" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/70611325/94376407-37d70200-00cf-11eb-8d84-ffd410e3e640.png">
<img width="422" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/70611325/94376422-47564b00-00cf-11eb-93c6-65d66b6e9c9d.png">

---

## Summary.

---
### *1.	What are the advantages or disadvantages of refactoring code?*

**Advantages:**
-	It makes the code execution faster.
-	It makes it easier to modify and adapt. For example, if any data were to change, we can confidence that this modification won’t alter in any way or create any error in our code.
-	It makes the code much simpler and easier to understand for anyone who didn’t create the code itself.

**Disadvantages:**
-	It causes a lot of bug to fix, so if you want to refactor a code make sure that you’ll have the time to do it.
-	If not doing right, can create more complex code and confuse the reader.


### *2.	How do these pros and cons apply to refactoring the original VBA script?*

**PROS:**
Refactoring the original code clearly helped me to drop the time of execution from the original script. 
Creating different variables helps to assure that any little modification in the data, will be contemplated in the code.
Adding comments at the beginning of every step, helped me to organize my code which I believed, it would also help understand anyone who read it.
 
**CONS:**
While refactoring, it caused a lot a bug fixing, more than the original code. It took a lot of time to go over, and over the script again sometimes even restarted from scratch. When I fixed one bug, it caused a another one in the next step of my script.
