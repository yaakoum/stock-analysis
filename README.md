# stock-analysis
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

'1a) Create a ticker Index
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
- The highest likelyhood for Louise to succeed with her campaign launch would be for her to launch in May. Based on the data used, she would be twice as likely to succeed than fail or cancel her campaign.
- On the contrary, the worst time for her to launch her campaign would be in December. Compared to May, there was less than a third of the number of successful campaigns in December and was almost the same as the number of failed that month.
### How These Pros and Cons Impacted the Original Data
- It is very clear that Louise cuts her chances for success significantly if she chooses to make her goal higher or near $15k. As per the chart analysis 50% failed at that goal range. Although 67% succeeded between the $35k-$45k range, the quantity of campaigns in that range only totaled to 9. Meaning if she expects her budget to be higher than $10k, I would highly suggest to stay as close to the $10k range or even try to go lower to help her chances to succeed.

