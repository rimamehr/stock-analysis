# Analyzing all stock returns

## Overview of Project/Purpose
Steve wants to find the total daily volume and yearly return for each stock in his dataset to help his parents make a sound decision on which is a good stock to invest. We originally helped write a code to go through all the ticker values to find the total volume and return on each stock, but our code was not running efficiently enough. We then refactored the code to improve efficiency  and provided Steve with a VBA script that can later be used on a larger dataset as well.


## Results

### Refactored Code
The yearly return is the percentage increase or decrease in price from the beginning of the year to the end of the year. In our images below we can see that the same set of 12 stocks performed considerably better in 2017 in comparison to 2018. In 2017 only one stock (TERP) had a negative yearly return and rest all of the stocks performed really well. When we compare that to our stock performance in 2018, they all faired quite poorly with exception of 2 stocks (ENPH, RUN). These results help Steve recommend some better stock choices for his parents to invest in. 

We have also provided below a snapshot of the time it took our code to run. These run times were after we refactored our VBA code for efficiency so that if Steve wants to run this analysis for a largest dataset he can do so quite effectively. 


<img src="/Resources/VBA_Challenge_2017.png" width="400"/> <img src="/Resources/VBA_Challenge_2018.png" width="400"/>

### Original Code
In our images below are the results from our original code prior to the refactoring. We can see from our timer window on these images compared to the ones above that our code ran much faster after refactoring. We also noticed that in our original code our highlighting of the return column for 2017 was not showing the correct color coding. In our previous code our formatting macro was built as a separate macro and not integrated into our stock analysis macro that could cause someone to forget running it. Our refactored code had all the formatting built into it at once and so when we run the subroutine it executes the entire code together. 


<img src="/Resources/VBA_Challenge_2017(old code).png" width="400"/> <img src="/Resources/VBA_Challenge_2018(old code).png" width="400"/>


### A Deeper look into our Original Code vs Refactored code
#### Original VBA code
<img src="/Resources/Original Script.png" width="400"/> 

#### Refactored VBA code
<img src="/Resources/Refactored Script.png" width="500"/>

## Summary
### What exactly is Refactoring?
Code Refactoring is a way of restructuring and optimizing existing code without changing its behavior. It is a way to improve the code quality. 
#### Advantages of Refactoring 
A refactored code is easier to understand or read, less complex and easier to maintain. If we see the code repeating itself, then we know there is a better way to writing the code and that it needs refactoring. Refactoring the code lessens our run time making it more efficient specially when we are working with larger data sets.
### Disadvantages of Refactoring
With every advantage there has to be some disadvantage to it. Refactoring the code can get sometimes messy and specially when we are bound by a deadline, its better to deliver a working code than trying to refactor a code last minute and deliver a broken one if refactoring is not completed timely. 
### Advantage to Refactoring our VBA code
From the images provided in our analysis above we can see that refactoring our original script for our stock analysis project made our code run considerably faster. We might not be able to realize the advantage of the time savings here because we were working with small dataset but it's necessary when working with larger data.
