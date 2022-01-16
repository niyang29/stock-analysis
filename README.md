# Green Stock Analysis

---
## Overview of Project

The purpose of this analysis is to assess which stocks Steve's parents should invest in. In order to determine which stocks to invest in, stock performances in 2017 and 2018 were assessed by looking at daily volumes and yearly returns of each stock. Steve then requested for a way to review thousands of stocks in a short period of time, thus, resulting in the need to refactor the original code. 
---
## Results

### Original v. Refactored Code

The original code definitely helped get us started. However, because Steve requested for a way to review thousands of stocks in a short period as opposed to 12 stocks, we needed to find a way to enhance the original code by refactoring it to measure performances of the stocks by first creating a tickerIndex variable and then creating three (3) output arrays to the ticker array - tickerVolumes(12), tickerStartingPrices(12), and tickerEndingPrices(12). A nested for loops and variables to loop through the arrays to output the “Ticker,” “Total Daily Volume,” and “Return” data.
Original code
https://github.com/niyang29/stock-analysis/blob/main/Original_code.png
Refactored code
https://github.com/niyang29/stock-analysis/blob/main/Refactored_code.png

### Stock Performances between 2017 and 2018
2017 stocks generally yield more daily volumes and return of investments for many of the stocks, whereas, stocks in 2018 experienced a decline in volume and return of investments. TERP was the only stock in the red, whereas eight (8) out of the ten (10) stocks in 2018 were in the red. ENPH and RUN both provided a postive return of investments. 
https://github.com/niyang29/stock-analysis/blob/main/Resources/Stock_Analysis_2017.png
https://github.com/niyang29/stock-analysis/blob/main/Resources/Stock_Analysis_2018.png
 
### Execution times of Original Script v. Refactored Script
Utilizing the the refactored scripts for 2017 and 2018 improved the execution time. See below. 

Execution time for 2017 Original Script
https://github.com/niyang29/stock-analysis/blob/main/Exec_time_original_script_2017.png
compared to 
Execution time for refactored 2017 script
https://github.com/niyang29/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png

Execution time for 2018 Original Script
https://github.com/niyang29/stock-analysis/blob/main/Exec_time_original_script_2018.png
compared to 
Execution time for refactored 2018 script
https://github.com/niyang29/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png
---
## Summary
### Advantages & disadvantages of refactoring code
Refactoring code is more efficient as a code already exist to edit meaning that the coder would have to take fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. In addition, refactoring is common as first attempts at code can produce more errors. The disadvantage is the on-going retesting of the functions. 
### How pros and cons apply to refactoring the original VBA script
The advantage of refactoring was that the logic was already set up, I needed to add in nested for loops to test and retest the functionality. I had the original code to compare results to. The disadvantage was having to test and retest and debug where I had errors. 
