# VBA-Challenge

In this worksheet we are looking at the key points of stock prices for all their active trading days for each year 2018, 2019, and 2020.  

# Defining Variables

The first thing I did you get my module started was to define all of the contstants I wanted to put in each column and Dims to reference to make my code easier to write instead of being repetative in my code. 

# Describing Code

The excutuion I wanted to run was to make sure that all of my code adn equatiuons within the module were done in each sperate worksheet.  After that I have an are on getting the first stock or the ticker symbol for each stock to line up down the column wihtout any repations. The first thing I wanted to do was have the progam look for the first row that I am strating my ticker symbol search. After that I am telling the code what column my output_row represents and the value of the day I am looking for with each ticker symbol.  After that I have my code begin the search for all the ticker symbols that do not repeate themselves all the way until the first balnk space in the second column.

# Calculations 

For my calculations I had the module pull the value of the stock from when it closed on the last day of the year and subtracted that from the open the stock had on the first day of the year.  I used that calculation to pull through the change in the stock price for the entire year.  In my other calculation I then used that yearly change variable and divided it by my price when the stock opened on the first day of the year to find the percent change throughout the year. 

# Nested IF 

Within my complete if statement that is finding all the stock tickers, I have another statment that is working to color the cells of the yearly change and percent change for each ticker.  I have my statment making the color of the cell green(4) if the change is greater than 0 and red(3) if it is less than zero meaning a loss for the year.

# End of Code

After all the above has been done I have the computer putting all the data and values in the places where I want them to go in order to see it nice and clearly. My first is the ticker symbols in the "I" column. Secondly I have the yearly change value for my tickers going in the "J" column right next to the ticker symbols.  Next to that I have my percent change value going next to my yearly change in the "K" column. And lastly I have added all the volumes for each trading added together and put that next to my percent change in column "L".

After I have all the value inputted for the first stock I have the code resent the values it is looking for and moving down to the next column to put the next values in for the next ticker symbol, going all the way down until I have run out of ticker symbols 

# Column Labling

At the end of my code I am labeling each column of the values I am putting in.  I put this last because it is the most simple part and put it at the end to set it and leave it. I also have a message box that pops up and tells me that the code is done running all of the calculations.

#Sources

Part of my code I hade help from a tutor Kourt Bailey, I spent an hour with him understanding how to write the calcualtions and using the variables to define 




