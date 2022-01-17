# VBA-Challenge
Refactoring stock analysis to determine if VBA script runs faster

## Overview 

### Purpose
This analysis is prepared to refactor the code used for the All Stocks Analysis previously performed. Although that code achieves accurate results, sometimes there are faster ways to accomplish the analysis. For this case, the code was changed to use a ticker index to loop through the data. Pop-ups that are written into the code are used to determine the script run time before and after the refactoring

## Analysis

### Process
The code from the original analysis was run to obtain the script run time. A new module was then created to hold the refactored code and the starter code was copied into that module. The code was rewritten, using a ticker index, reusing some code from the original analysis and then debugging. The debugging process was a challenge as the loop through arrays to output the return kept returning an error. After several attempts to correct the error, a new data set was downloaded and the code rewritten in a fresh copy. The code finally ran with no errors. 

### Analysis
As exhibited with the pop-ups below, the refactored code for both year 2017 and 2018 ran faster than the previous code. The same data was returned in the table for both the original code and the refactored code, but the changes help the VBA script to run faster. 

![All_Stocks_Analysis_2017](https://github.com/Dainita/VBA-Challenge/blob/main/All_Stocks_Analysis%202017.png)
![All_Stocks_Analysis_2017](https://github.com/Dainita/VBA-Challenge/blob/main/All_Stocks_Analysis%202018.png)

![VBA_Challeng)2017](https://github.com/Dainita/VBA-Challenge/blob/main/VBA_Challenge_2017.png)
![VBA_Challeng)2017](https://github.com/Dainita/VBA-Challenge/blob/main/VBA_Challenge_2018.png)

![All_Stocks_2017](https://github.com/Dainita/VBA-Challenge/blob/main/All_Stocks_2017.png)
![All_Stocks_2018](https://github.com/Dainita/VBA-Challenge/blob/main/All_Stocks_2018.png)


## Results
Refactoring the code successfully ran the VBA script faster. This code is now more easily understandable and maintainable. The cleaner code can save time and money in the future with easier updates. A major disadvantage was the additional debugging and retesting that was required to make the code run properly. 
