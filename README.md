# Stocks Analysis - VBA Challenge
Please note: I did not create a button for the refactored analysis. I ran the refactored code manually on the VBA Editor. The button on the sheet is for the original analysis only. 

## Overview of Project
The purpose of this analysis was to review the code we wrote in Module 2 and refactor it to run at a more efficient pace. This was mainly done throught the use of arrays. Where we briefly used an array in the Module 2 walkthrough, in the challenge we were tasked to create three new output arrays and an index variable to guide our code through the 12 stocks for analysis. This will allow us to add more stocks and more information to the dataset, without our macro slowing down or breaking. We also added comments throughout the code in order to define what each line does for future reference. 
  


 ## Results   

### 2017 
    
After running our code for both 2017 and 2018 for the green stocks, it is clear there was a large difference in stock performance between the two years. In 2017, it is clear that the market was more bullish when looking at our green stocks. 4 stocks saw up to 100% returns, and all but one stock had a positive performance for the year. The average return percentage of these stocks is 67.3%. Compare that with the average return from the S&P 500 (a stock market index that measures the stock performance of 500 large companies listed on the market) in 2017 of 21.8% (according to an article from Business Insider (1)), and it is clear that this groups of stocks widely outperformed the market as a whole. See the below charts for the breakdown of 2017 and the comparison of the original and refactored code. 

![](Resources/2017moduleresults.png)

![](Resources/2017refactoredresults.png)


### 2018 
2018 saw a correction of the gains on these stocks. All but 2 stocks had a negative return on the year, with the average return being -8.5%. When referencing the same Business Insider article regarding the S&P 500, we can see an annual return of -4.4%. In fact, the article lists all annual returns for the decade, and we can see that 2018 was the only year in which there was on overall negative return rate. (1) However, I believe this was to be expected with this group of stocks. It is extremely rare to see gains of over 100%, and to think that these stocks would continue to keep that pace after 2017 was careless. However, if you bought these stocks at the beginning of 2017 and held through the end of 2018, 10 out of 12 stocks would have increased in value, with $SPWR and $TERP being the losers. $ENPH and $RUN were extremely successful, being the only two stocks to perform in the green over both years, and $ENPH realizing the biggest percentage gain over the 2 year period of 211.4%. See the below charts for the breakdown of 2018 and the comparison of the original and refactored code. 

![](Resources/2018moduleresults.png)

![](Resources/2018refactoredresults.png)


   
## Summary

1.  The advantages of refactoring code are worth taking advantage of. When refactoring, we get to add new and improved functionality, as well as make the code more efficient. Like all programs, the more data we are dealing with, the harder the program works to analyze it all. The ceiling of Excel's abilities with large amounts of data comes sooner than other related programs, and it is important to make VBA's code as efficient as possible in order to keep the spreadsheet working as intended. Refactoring also helps to improve the logic of the code, and helps newer coders to understand what is happening with existing code. 
    
One of the few disadvantages I can think of with refactoring code is the potential to break the existing program while working on it. However that can be easily mitigated by creating backups of the data and the program, and conducting the refactoring on those backups. Another disadvantage would be the time drain. Who knows if refactoring the code would be worth it until the end, where we can determine how much more efficient we had made it. 


2. These pros and cons definitely apply to our code. Our code is now more efficient, with a time decrease of 0.773438 seconds when calculating the 2017 results, and a decrease of 0.71875 seconds for the 2018 results. As the amount of data increases on this spreadsheet, that difference in time will become more important. When it comes to disadvantages, with my current skills level in VBA, I am not sure as of now how we can improve the code further. I would suggest leaving it as is for now, but if the time to calculate results increases as the amount of data increases, we can then refactor the code again. 


References: https://www.businessinsider.com/personal-finance/average-stock-market-return