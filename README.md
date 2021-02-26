# stock-analysis

### Analyzing Stock Performances by Year

## Project Overview

In this project, I used Microsoft Excel VBA to create a macro that outputs the yearly returns for 12 different stocks of companies that work in the green energy space. I wrote this program as a tool for financial analysts and retail investors to quickly analyze different stocks according to their opening prices, closing prices, and daily trade volume. Users can quickly generate a results table, based on their desired input year, that display each stock's performance in the areas of total daily trading volume and yearly return. Using conditional statements, I also formatted the results output so that positive returns are highlighted in green, and negative returns are highlighted in red. The code could be modified slightly if I need to analyze more than 12 stocks at a time, with minimal refactoring. I also included a macro-button on the AllStocksAnalysis sheet for easy access to the program. Using VBA code written for a smaller-scope project, I refactored the existing code to optimize performance and decrease run time. Later, I'll discuss some of the key difference between the old code and the new code and how they impacted performance metrics. 

## Stock Analysis Results

When we look at the results output table for 2017, we see that 11/12 of the selected stocks had a positive return. 'DQ' had the highest yearly return at 199.4% and TERP yielded a -7.2% return. SPWR had the highest total daily volume at over 780 million trades, and DQ had the lowest at 35.7 million trades. One quick inference I can draw from these results, in broad terms, is that 2017 was overall a good year for green energy stocks. It may be interesting to analyze any policy that changed in that year that may have fueled these results. 

In contrast to the fortuitous year that green energy stocks saw in 2017, 2018 saw a clear overall contraction in the returns of these stocks. 10/12 green energy stocks saw declines from the beginning of the year to the end. ENPH led the way in total trading volume at over 607 million total trades, and finished the year with an 81.9% return. RUN saw the highest growth in their stock price, at 84% with a robust 502 million daily trades. AY had the lowest daily trade volume at 83 million, and DQ showed the greatest decrease in their share price at -62.6%. Note that in 2017, DQ led the pack with the highest growth, and then in the following year had the highest contraction. 

## Summary
Below are the code performance results from before and after refactoring. 
# 2017:

![Green_stocks_2017](https://user-images.githubusercontent.com/76958825/109364396-b5f52b80-785c-11eb-9660-eb894781d12b.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/76958825/109364461-dfae5280-785c-11eb-9071-5232c59db948.png)

# 2018:
![Green_stocks_2018](https://user-images.githubusercontent.com/76958825/109364569-12584b00-785d-11eb-9e42-756f1a22f07f.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/76958825/109364590-20a66700-785d-11eb-8880-7da85dedcec2.png)

I was able to greatly improve my code's performance through refactoring. The initial code took over 1.36 seconds to run for the year 2017, and by refactoring, I was able to drive that time down to .20 seconds. That's over a whole second of improvement! I achieved similar results for the year 2018, where the initial code ran in 1.36 seconds, and the refactored code ran in .21 seconds. Additionally, when I refactored I added in several lines of formatting and style to the output table, which clearly did not have a great affect on the code's speed. This leads to believe that VBA is able to process and execute basic formatting changes very quickly.

Through refactoring I made some key changes to how the program loops through the thousands of tickers and their stock information. By using a variable tickerIndex, which holds an integer that loops through an array of strings, the computer could directly reference each string's index in the array. This variable tickerIndex could then hold it's value through the loop, which also eliminated part of the conditional checks that the code ran when it was determining if a stock was the first or last instance in the spreadsheet. This is shown in the following block of code:

'''

If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then

                    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
                End If
                
'''

By optimizing the conditional check, the code did not have to verify more than one condition to know if it needed to move on the next ticker in the array of tickers.

Refactoring code is a key part of programming because it can save time, allow for collaboration between programmers, and ultimately improve the performance of your code. It can be immensely faster to take code you've written for a similar project, and refactor it to fit the current needs of the project, than to start from scratch in each program. 
However, there are downsides to refactoring as well. One problem that I encountered when refactoring was that I was regularly getting DataType mismatch errors, as I had not properly declared the tickerIndex as an integer. While this instance was relatively straightforward to fix, in a program that is only a few dozen lines of code, at scale, this could have been a major issue if I was refactoring thousands of lines of code. 
In conclusion, refactoring code is an important skill for a programmer to have on their 'tool belt', but one must tread carefully as the risks can create major problems down the road.
