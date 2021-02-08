# Kickstarting with Excel

## Overview of Project

### Purpose
- This project delivers on creating an analysis for stock trading data from 2017 and 2018. The analysis gives the total daily volume for the yearly data while also giving the return based on the starting and ending price. The purpose of this project is to help Steve refactor the previous written code for efficiency and also make sure that the code can be run with any amount of data,number of tickets.  

## Analysis and Challenges

### 2017 Analysis
Ticker    Total Daily Volume    Return
AY    309,500    0.0%
CSIQ    310,592,800    33.1%
DQ    35,796,200    199.4%
ENPH    221,772,100    129.5%
FSLR    684,181,400    101.3%
HASI    80,949,300    25.8%
JKS    191,632,200    53.9%
RUN    267,681,300    5.5%
SEDG    206,885,200    184.5%
SPWR    782,187,000    23.1%
TERP    139,402,800    -7.2%
VSLR    109,487,900    50.0%

Looking at the above results the majority of the tickets in 2017 had an increase in return with DQ having the greatest increase at 199.4%.

### 2018 Analysis
Ticker    Total Daily Volume    Return
AY    83,079,900    -7.3%
CSIQ    200,879,900    -16.3%
DQ    107,873,900    -62.6%
ENPH    607,473,500    81.9%
FSLR    478,113,900    -39.7%
HASI    104,340,600    -20.7%
JKS    158,309,000    -60.5%
RUN    502,757,100    84.0%
SEDG    237,212,300    -7.8%
SPWR    538,024,300    -44.6%
TERP    151,434,700    -5.0%
VSLR    136,539,100    -3.5%

Looking at the 2018 results the majority of the interested stocks had a decrese in return, but only JKS, SPWR, and TERP had net decrease over the span of 2017 and 2018.

### Efficiency
- https://github.com/ssyed21/stocks-analysis/blob/main/resources/VBA_Challenge_2017.png
https://github.com/ssyed21/stocks-analysis/blob/main/resources/VBA_Challenge_2018.png

Looking at the resulting screenshots the 2017 runtime was barely more efficient.

### Challenges and Difficulties Encountered
- The biggest challenge for myself was learning the VBA language and steer away from switching between python and java while I was coding. A big problem was that I constanly was using != instead of <>.

## Results
- What are the advantages or disadvantages of refactoring code?
The advantages of refactorign code is that by doing so you can make it more efficient and also add more functionality to it. Refactoring can make code run in less steps, use less steps, and improve the logic.
The disadvantages to refactoring is that if not careful you can break the code and waste a good amount of time. Also refactoring could cause future bugs to occur if proper test cases aren't given.

- How do these pros and cons apply to refactoring the original VBA script?
For refactoring the original VBA script into it's current commit, we added major functionality to it so that we can use the script with any amount of sorted stock data. The only cons that I see, could be from the amount of time spent writing the new code could have been used for a different analysis of the DQ Analysis results and that a bug may occur with data not properly sorted.
