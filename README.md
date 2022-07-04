# An Analysis of Green Energy Stocks

## Overview of Project
My good friend, Steve, just obtained his degree in Finance, and his parents are so supportive that they’ve agreed to be his first clients. Steve’s parents are very passionate about green energy and are interested in investing all their funds into one company, DAQO New Energy Corporation, but have admittedly done no research regarding the industry or DAQO’s performance. Steve has a greed not only to research DAQO stock but also other similar green energy stocks in order to diversify his parent’s investment. Steve has created an excel file with 2017 and 2018 stock data for a handful of green energy companies to analyze.
### Purpose of Project
Steve previously requested my assistance in analyzing the stock data he has compiled. I have already used VBA to create code automation to analyze Steve’s dataset with the click of a button. However, he was so impressed with this analysis that he’s decided to expand the dataset to include the entire stock market over the last few years. Unfortunately, my original code will not be able to execute this analysis efficiently; therefore, I’ve decided to refactor my code to enable VBS script to run faster and ultimately more effectively.

## Results
My analysis involved the review of stock data for 12 green energy companies over the years 2017 and 2018. 
### DAQO New Energy Corporation (DQ) Performance
Although DAQO New Energy Corporation had an extremely high percentage return in 2017, this performance did not last, and the company experienced negative return in 2018. This indicates that DAQO New Energy Corporation may not be a good company to invest in at all as they may continue a downward trend.  
### All Other Green Energy Stocks Performance
Unfortunately, for the green energy market my analysis showed that most other stocks experienced a similar performance trend, in that, they performed very well in 2017, only to see negative returns in 2018. Two companies were the exception, however, and it is ideal for Steve’s parent to invest in both for diversification. These companies include ENPH, which had returns of 129.5% and 81.9 % in 2017 and 2018 respectively; and RUN, which had returns of 5.5% and 84 % in 2017 and 2018 respectively. Because ENPH experienced high return rates in both years, they may have better stability and it would make more sense for Steve’s parents to invest a majority of their investment funds into this company and a minority into RUN.

Images of 2017 and 2018 analysis for all companies are shown below:

<picture>
 <source media="(prefers-color-scheme: light)" srcset="https://github.com/ODaniels852/stock-analysis/raw/main/Resources/VBA_Challenge_2017.png">
 <img alt=" Shows 2017 Stock Analysis Results."/>
</picture>
 
<picture>
 <source media="(prefers-color-scheme: light)" srcset="https://github.com/ODaniels852/stock-analysis/raw/main/Resources/VBA_Challenge_2018.png">
 <img alt=" Shows 2018 Stock Analysis Results."/>
</picture> 

### Refactoring Results
Overall, I am extremely satisfied with the result of my refactoring efforts. One of the big areas of improvement involved the use of multiple arrays and a ‘tickerIndex’ variable that was used in four different arrays throughout the code. Utilizing the tickerIndex variable as the index in other arrays, like shown in the example of my code below, was also a bit easier to conceptually understand. 
<picture>
 <source media="(prefers-color-scheme: light)" srcset="https://github.com/ODaniels852/stock-analysis/raw/main/Resources/tickerIndex_Code_Example.png">
 <img alt=" Shows refactor code example."/>
</picture> 

Prior to refactoring my code, it took 1.28 seconds to analysis results for 2017 and 1.47 seconds to analyze results for 2018. After refactoring, it took .46 seconds to analysis results for 2017 and .45 seconds to analyze results for 2018, proving that my efforts made this code phenomenally more efficient.
Images of the times to run my analysis before and after refactoring are shown below:

#### **BEFORE Code Refactor**
#####_2017_
<picture>
 <source media="(prefers-color-scheme: light)" srcset="https://github.com/ODaniels852/stock-analysis/raw/main/Resources/BEFORE_Refactor_2017_AnalysisTime.png">
 <img alt=" Shows time to run analysis on 2017 stock data before refactoring code."/>
</picture>
#####_2018_ 
<picture>
 <source media="(prefers-color-scheme: light)" srcset="https://github.com/ODaniels852/stock-analysis/raw/main/Resources/BEFORE_Refactor_2018_AnalysisTime.png">
 <img alt=" Shows time to run analysis on 2018 stock data before refactoring code."/>
</picture> 

#### **AFTER Code Refactor**
#####_2017_
<picture>
 <source media="(prefers-color-scheme: light)" srcset="https://github.com/ODaniels852/stock-analysis/raw/main/Resources/AFTER_Refactor_2017_AnalysisTime.png">
 <img alt=" Shows time to run analysis on 2017 stock data after refactoring code."/>
</picture>
#####_2018_ 
<picture>
 <source media="(prefers-color-scheme: light)" srcset="https://github.com/ODaniels852/stock-analysis/raw/main/Resources/AFTER_Refactor_2018_AnalysisTime.png">
 <img alt=" Shows time to run analysis on 2018 stock data after refactoring code."/>
</picture> 

## Summary
In summary, I believe the advantages of refactoring code include, but are not limited to, the ability to analyze larger datasets and do so in a shorter amount of time. This can be especially important when it is known that code will be used repeatedly on newer and possibly larger datasets in the future. One disadvantage of refactoring code is that it may take a long time to do if the coder has little experience. In this exercise, while I see the immense value in refactoring the original VBA script, my inexperience with VBA led to me taking a long time to understand the purpose of some refactoring elements. Ultimately, I don’t believe refactoring is as beneficial of a use of time for someone who isn’t more experienced or who hasn’t had more practice with VBA coding.

