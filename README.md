# Data Analytics Challenge - Stock Market Data Analysis with VBA

Welcome to the Data Analytics Challenge #2. This project contains the documents and VBA macros related to our second challenge in Data Analytics. As part of this challenge, we were tasked with using VBA scripting to analyze stock market data and create two summary tables.

## Program Used

I used VBA scripting to analyze data in three Excel spreadsheets for the years 2018, 2019, and 2020.

## Summary Table 1

Summary Table 1 is designed to calculate the following fields for each stock:

- **Ticker**: The stock ticker symbol.
- **Yearly Change**: The difference between the stock's opening price at the beginning of the year and the closing price at the end of the year.
- **Percent Change**: The percentage change in the stock price from the beginning to the end of the year.
- **Volume**: The total trading volume of the stock for the year.

The VBA script loops through the first column of each year's data to gather and calculate these fields for each stock and populates them in the Summary Table 1.

## Summary Table 2

Summary Table 2 builds upon the data from Summary Table 1 to find the following:

- **Greatest % Increase**: The stock with the highest percentage increase in price during the year.
- **Greatest % Decrease**: The stock with the highest percentage decrease in price during the year.
- **Greatest Total Volume**: The stock with the highest total trading volume for the year.

The VBA script loops through the data in Summary Table 1 to identify and populate Summary Table 2 with the required information.

## Other Info

To enhance the visualization of the data in the worksheets, conditional formatting was applied to highlight positive and negative values in the **Yearly Change** and **Percent Change** columns.

The VBA macro was applied to all three worksheets (2018, 2019, and 2020) to ensure comprehensive analysis and generation of the summary tables.

Thank you for exploring our Data Analytics Challenge - Stock Market Data Analysis with VBA repository! If you have any questions or suggestions, feel free to reach out.
