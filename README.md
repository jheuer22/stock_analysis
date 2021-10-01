# Stock Analysis Using Refactored VBA Code

## Overview of Project: 
In this analysis we refactored an existing VBA code that was created to evaluate the performance of a set of stocks over the course of 2017 and 2018. We refactored this code in order to make it more efficient. 


## Results: 

In this VBA code, we were analyzing the performance of 12 stocks during 2017 and 2018. The original code was written with nested "For Loops" to loop through all of the rows of stocks in the data file to create an output chart. This output table showed the Total Daily Volume of trades for each stock, as well as the return rate for the year. Below are the output charts from the 2017 and 2018 stock analysis. 

![Output-2017](/Output-2017.png). 
![Output-2018](/Output-2018.png).

As you can see, overall the stocks had better return rates in 2017 vs. 2018 based on the number of postive retun rates in green vs. negative return rated highlighted in red.  All of the stock but one ("TERP") had an overall positive return rate in 2017. On the other hand, only 2 stocks in 2018 had an overall increase for the year ("RUN" and "ENPH") while the rest had an overall negative return rate.

Both the original code and the refactored code were able to produce the same outcomes into the table, but the refactored code was able to perform the analysis significantly faster. The original code took approximately 0.504 and 0.492 seconds to run the code for 2017 and 2018, respectively. The refactored code was able to produce the same outcome for 2017 and 2018 in 0.0898 and 0.0859 seconds respectively. You can see that the refactored code would be much more efficient if used with an even larger original data file of stocks to analyze. 

## Summary: 

In this Analysis we refactored an existing VBA stock analysis code in order to make the code more efficient. We streamlined the code by making arrays of the variables so that we could loop through the data more quickly compared to the original code. When a code is refactored, the same results are produced but it is done with a more efficient method. The benefits of refactoring a code are that you can usually make the code less complicated, more efficient, and you can use less computational power. One drawback of refactoring code is that it can sometimes take more time to refactor an existing code, especially if you are not the one who initially wrote the code. It can take time to acquaint yourself with the logic and flow of the original code then come up with a more efficient alternative way to reach the same end goal. In this case, the refactored VBA code was significantly faster than the original code. However the process of refactoring the code was considerably more time consuming since I ran into issues on a portion of the code that took additional time to resolve.


