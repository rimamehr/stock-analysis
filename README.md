# Analyzing all stock returns

## Overview of Project/Purpose
Steve wants to find the total daily volume and yearly return for each stock in his dataset to help his parents make a sound decision on which is a good stock to invest. We originally helped write a code to go through all the ticker values to find the total volume and return on each stock but our code was not running efficiently enough. We then refactored the code to improve efficency and provided Steve with a VBA script that can later be used on a larger dataset as well.


## Results

### Refactored Code
The yearly return is the percentage increase or decrease in price from the beginning of the year to the end of the year. In our images below we can see that the same set of 12 stocks performed considerablaly better in 2017 in comparison to 2018. In 2017 only one stock (TERP) had a negative yearly return and rest all of the stocks performed really well. When we compare that to our stock performace in 2018, they all faired quite poorly with exception of 2 stocks (ENPH, RUN). These results help steve recomend some better stock choices for his parents to invest in. 

We have also provided below a snapshot of the time it took our code to run. These run times were after we refactored our VBA code for efficiency so that if Steve wants to run this analysis for a largest dataset he can do so quite effectively. 


<img src="/Resources/VBA_Challenge_2017.png" width="400"/> <img src="/Resources/VBA_Challenge_2018.png" width="400"/>

### Original Code
In our images below are the results from our original code prior to the refactoring. We can see from our timer window on these images compared to the ones above that our code ran much faster after refactoring. We also noticed that in our original code our highlighting of the return coulmn for 2017 was not showing the correct color coding. In our previous code our formatting macro was built as a seperate macro and not integrated into our stock analysis macro that could cause someone to forget running it. Our refactored code had all the formatting built into it at once and so when we run the subroutine it executes the entire code together. 


<img src="/Resources/VBA_Challenge_2017(old code).png" width="400"/> <img src="/Resources/VBA_Challenge_2018(old code).png" width="400"/>


### A Deeper look into our Original Code vs Refactored code
Original Code <img src="/Resources/Original Script.png" width="400"/> Refactored Code <img src="/Resources/Refactored Script.png" width="500"/>
