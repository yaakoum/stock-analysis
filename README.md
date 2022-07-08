# stock-analysis
To view the results first hand, please click this link to access the Excel file: [VBA Challenge - Stock Analysis](https://github.com/yaakoum/stock-analysis/blob/main/VBA_Challenge.xlsm)

## Overview of Project
Steve has recently graduated with his finance degree. His parents are very proud of him and would like to become his first clients. They have a problem with putting all their eggs in one basket and Steve has become concerned with their lack of diversification. He promised he would look into DQ stock that they're invested in as well as a list of other stocks that use renewable energy. 

### Purpose
Steve has provided a large list of stocks with extensive data. The dataset is simply too large to analyze manually and can open room for error. Hence why he has come to ask for our assistance in automating these processes. With the use of VBA and functions like "for loops" and "if functions", we were able to completely automate the process for him. Not only that but we went a step further to refactor the code and make it run quicker and more efficiently. 

## Analysis and Challenges

### Overview of analysis
There were two main analysis that were done for this project. As per below, you will find one analysis that compares the number of of Successful, Failed, and Canceled Theater campaigns relative to their launch month. The second analysis focused on the percentage of successfull, Failed, and Cancelled Campaigns relating to Plays based on their goals. The main findings of this analysis were as follows:

#### Outcome of Theater Campaigns Based on Launch Date
![Outcomes vs Launch Date](https://github.com/yaakoum/kickstarter-analysis/blob/main/Resources/Theater_Outcomes_vs_Launch.png)
As seen on the line chart above, we can see that most theater campaigns were successful and there was a correlation between the month they were launched on. Most Successfull ones were launched in the month of May while the failed campaigns remained relatively consistant throughout the year.
#### Outcome of Campaigns related to Plays Based on Goals
![Outcomes vs Goals](https://github.com/yaakoum/kickstarter-analysis/blob/main/Resources/Outcomes_vs_Goals.png)
The second analysis looked at the percentage of successfull, failed, and canceled play related campaigns based on their goals amounts. What we find here is a clear distinction in the percentage of successfull and failed campaigns with the $15k-$15.9k range being split exactly by 50%. Most higher amounts did not find success and achieving their goals which is to be expected except a few of the being successful in the higher end of the price range.

### Challenges and/or difficulties
Here we discribe different challenges we may or may not have encountered and different solutions. Listed below are some of these examples:
  1. One of the main issues that can be expected is of course issues with Excel. This is especially true as well when making reference to different sheets and copying formulas where the reference can also change. Additionally, there are some formulas that are long and comnplicated and over time one can lose focus and overlook certain details. 
  2. Although Pivot Tables provide many benefits, manipulating the data is not always as clear cut as one can expect. What this means is that achieving the correct analysis with the data you have can be difficult and manipulating to achieve your goals can sometimes prove to be troublesome.

These are two examples of challenges one can expect to see. To manage these concerns, I would periodically ensure that my data and formulas are working correctly. Secondly I would take mental checks where I would look at the data and ensure it makes sense. Meaning if I look at the data vs what I have on a chart, I would ask myself "does this make sense". Simple mitigating strategies definitely go a long way.

## Results
There were many findings and below they are displayed for Louise to make the most educated decision for her playwrite kickstarter.

### Conlusions based on Theater Outcomes by Launch Date
- The highest likelyhood for Louise to succeed with her campaign launch would be for her to launch in May. Based on the data used, she would be twice as likely to succeed than fail or cancel her campaign.
- On the contrary, the worst time for her to launch her campaign would be in December. Compared to May, there was less than a third of the number of successful campaigns in December and was almost the same as the number of failed that month.
### Conclusion based on the Outcomes based on Goals
- It is very clear that Louise cuts her chances for success significantly if she chooses to make her goal higher or near $15k. As per the chart analysis 50% failed at that goal range. Although 67% succeeded between the $35k-$45k range, the quantity of campaigns in that range only totaled to 9. Meaning if she expects her budget to be higher than $10k, I would highly suggest to stay as close to the $10k range or even try to go lower to help her chances to succeed.
### Possible limitations of the dataset, and future recommendations
- The biggest limitation I found with the dataset would be the type of plays and their location of being launched. Although, with a larger dataset we gain better insight with varying data, it may be at the expense of relevancy. To gain the highest level of accuracy, I believe we would need to curate a list of campaigns that are not only in the same region of launch and with similarity in the genre of plays.
- As mentioned some relavent charts I would recommend to use would have to do with comparing some of the data we use based on country launched in. The other potential chart comparison can be to explore the percentage of plays in the range Louise is that exceeded their goals to gauge and understand better how successfull these campaigns really were.
