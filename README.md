# VBA-Challenge
## The following code analyses yearly stock market data using VBA 

### First part of the project is to loop through the daily stock prices for a large sample of stocks represented by their TICKERS. Data provided includes 

Ticker, Open Price, High, Low, Close, Stock Volume

Steps Include
  - Loop through the data and display name of the stocks represented by Tickers in an output table
  - Analyse the daily data for each ticker to summarise the yearly change in price i.e. Closing Price on last day - Opening price on first day 
  - Summarise the percent change in stock price over the year
  - Total stock volume that was traded in the year for each stock
  - To ensure readability, output is color coded to Green and Red for Postive and Negative change in prices 
  
### Additional functionality was added to automate the macro for worksheets: The yearly data was analysed for three years and macro was enabled to loop through all worksheets in the workbook at once
  
The output has been further analysed to give insights into: 
  - The stock with the greatest volume traded during each year
  - The Stock with the highest positive jump during the year
  - The stock with the highest negative slide during the year
  
  
### Output files are uploaded as :
  Screenshot of 2020
  Screenshot of 2019
  Screenshot of 2018
  VBA Code in file StockAnalysis.vbs which can be opened using VSCode/Xcode
  
  #### _Please note the source file with input data and the macro (.xlsm) is too big to be uploaded here
