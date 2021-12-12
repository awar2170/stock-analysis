# Project Overview 
The purpose of this analysis was to write and refactor VBA code that effeciently outputted the "Total Daily Volume" and "Return" for 12 different stocks for the year of 2017 and 2018 to optimize the run time.  The code is interactive and able to analyze either year, depending on the user's input.  The modules specifically focused on writing code that analyzed stocks with the use of nested loops, conditional if statements, and interactive macro buttons.  The challenge specifically focused on refactoring inefficent code by reusing code and logic learned from the modules.  

# Results 
This section will compare the stock performance between 2017 and 2018 and discuss the execution times of the orignal script and the refactored script. 

## Stock Years Analysis 
The 2017 stocks have a maximum total daily volume of $782,187,000 from SPWR's stock, a minimum total daily volume of $35,796,200 from DQ's stock, a maximum return of 199.4% from DQ's stock, and a minimum return of -7.2% from TERP's stock.  The 2017 stock has average total daily volume of $263,886,592 and an average return rate of 67.3%. 

The 2018 stocks have a a maximum total daily volume of $607,473,500 from ENPH's stock, a minimum total daily volume of $83,079,900 from AY's stock, a maximum return of 84.0% from DQ's stock, and a minimum return of -62.6% from TERP's stock.  The 2018 stock has an average total daily volume of $275,503,183 and an average return rate of -8.5%.

2017 is the more profitable year, despite having a lower average total daily value, because the year has a positive average return rate.  This means that, on average, the portfolio made money that year. 

## Macro Runtime Analysis
In the orignal script, the 2017 stocks analysis macro had an 0.828125 second run time and the 2018 stocks analysis macro had and 0.84375 runtime.  In the refactored code, the 2017 stocks analysis macro had an 0.828125 second run time and the 2018 stocks analysis macro had and 0.84375 runtime. 

![2017 Refactored Runtime](https://github.com/awar2170/stock-analysis/blob/main/VBA_Challenge_2017.PNG)
![2018 Refactored Runtime](https://github.com/awar2170/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

The orignal code relied on nested for loops that looped through each ticker and then looped through each row of the data.  Instead of saving the values it calculated into an array like the refactored code, the orignal code outputs values into the "All Stocks Analysis" worksheet as it calculated the values.  This means that the computer had to spend additional time throwing away values each time it looped through the nested for loops in the orignal code.  The refactored code most likely ran faster due to using separate loops and assigning the values from these loops to arrays.  This removed the computer's need to throw out values and do unnecessary computations since after each loop completed, the computer stored the loop's outputs into an array.  This made outputting the values faster because the computer outputs the values stored in the arrays instead of outputting values as it calculates them.        

# Summary
This section will consider the advantages and disadvantages of refactoring code in general and a comparison between the advantages and disadvantages of using the original code compared to the refactored code. 

## The Advantages of Refactoring Code 
The advantages of refactoring code for effeciency are that it requires less computer computational effort and the coder can implement different learned practices.  Using less computper power is useful because this makes the code more likely to be able to run on other types of computers that may not have the same amount of power as the computer where the code was written.  The different learned practices may not always be more effecient, but there may be a more universal way of approaching a certain problem in code that will make it easier to understand for anyone else who looks at the code at a later date.

## The Disadvantages of Refactoring Code 
The disadvantages of refactoring code is that 

## Applying the Advantages and Disadvantages of Refactoring Code
