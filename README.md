# Stock-analysis
To view the results first hand, please click this link to access the Excel file: [VBA Challenge - Stock Analysis](https://github.com/yaakoum/stock-analysis/blob/main/VBA_Challenge.xlsm)

## Overview of Project
Steve has recently graduated with his finance degree. His parents are very proud of him and would like to become his first clients. They have a problem with putting all their eggs in one basket and Steve has become concerned with their lack of diversification. He promised he would look into DQ stock that they're invested in as well as a list of other stocks that use renewable energy. 

### Purpose
Steve has provided a large list of stocks with extensive data. The dataset is simply too large to analyze manually and can open room for error. Hence why he has come to ask for our assistance in automating these processes. With the use of VBA and functions like "for loops" and "if functions", we were able to completely automate the process for him. Not only that but we went a step further to refactor the code and make it run quicker and more efficiently. 

## Results
As menioned in the overview, there were two major takeaways from this project:
1. The analysis has been automated and Steve can now assist his parents in diversifying their portfolio;
2. The code has been refactored and optimized making it more efficient and able to take on more data if Steve chooses to analyze more stocks.

### Automation 
<img src="https://github.com/yaakoum/stock-analysis/blob/main/Stock%20performance/Stock_Performance_2017.png" width="200" height="200" />                 <img src="https://github.com/yaakoum/stock-analysis/blob/main/Stock%20performance/Stock_Performance_2018.png" width="200" height="200" />

Through the use of Visual Basics for Application (VBA), we were able to utilize many functions, variables, and arguments to automatically analyze the stocks Steve provided. As seen in the photo above there were major differences in the performance of most of these stocks between 2017 and 2018. The only stocks that continued to perform well in consecutive years were ENPH and RUN. The other Major finding is that DQ's performance dropped severly and is the very reason for Steve's concern of his parent's lack of diversification. However, this automated analysis now provides Steve with a few options to consider when offering new investment options to his parents.

### Refactorization 
<img src="https://github.com/yaakoum/stock-analysis/blob/main/Pre-Refactoring%20Timers/Pre_Refactoring_2018.png" width="450" height="300" />                 <img src="https://github.com/yaakoum/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png" width="400" height="300" />

As you can tell, there was a major boost in perforamnce after refactoring the code. Above is a screenshot for the timer of the same year analysis before and after the refactoring. The improved code is now more than 3.5 times faster. The major difference in the code was the use of more arrays and refraining from using a nested for loop. To visualize what this would look like, here is an example of how a nested for loop works:


       For i = 1 to 10
       
         'a line of code here will run 10 times

        For j = 1 To 20

            'a line of code here will run 200 times

        Next j

       Next i

In addition to make it work this way, we would have to activate different sheets multiple times and flip between them. The refactored code is much more clever in allowing us to activate the sheet to input all the data we need once, then activate the other sheet to output all that data for visualization once. The bulk of the change had to do with avoided nested for loops and flipping between sheets multiple times. This was achieved by creating a variable for a tickerIndex and creating 3 more arrays to be placeholders for each ticker symbol. Here is a preview of the code that made this possible:


    Dim tickerIndex As Integer
    tickerIndex = 0

    Dim tickerVolumes(12) As Single
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    For i = 0 To 11
       tickerVolumes(i) = 0
        
    Next i
        
    For i = 2 To RowCount
    
       tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
  
      If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

          tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

      End If
      
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

           tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

        End If


            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

               tickerIndex = tickerIndex + 1

            End If
    
    Next i

Most of the remaining code remained relatively the same aside from the above. What we see above is that as the code run through the for loop, it gathers the information for the first ticker symbol and saves it in the array rather than copying and pasting between sheets every time it loops. 

## Summary

### Pros and Cons of Using Refactored Code
- Refactoring code is extremely beneficial for many reasons. Refactoring makes the code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Overall, it allows a code to run better and smoother.
- On the other hand, there are some disadvantages to refactoring code and it is primarily to do with the fact that its so complex. The complexity can also result in more time consumption. Although very rewarding, if it is not needed it may not always be the best decision.
### How These Pros and Cons Impacted the Original Data
- Luckily in this case, it was not too difficult nor time consuming and as such only resulted in benefits. We encountered a code that is simpler to understand, quicker to run, and can easily be reused. Now if steve would like to add more data to the list by including more stocks, a simiple change can be made for the code to function. More importantly, it will be able to run through more data with 3.5 times the speed. This can become more significant once the dataset becomes very large and uses a lot of memory.

