# VBA Homework: The VBA of Wall Street
## Summary
* A Module Challenge for University of Toronto's Data Analytics Bootcamp
* Files present an exploratory analysis of dataset created by Trilogy Education Services,LLC using Excel VBA
## Background

You are well on your way to becoming a programmer and Excel master! In this homework assignment, you will use VBA scripting to analyze generated stock market data. Depending on your comfort level with VBA, you may choose to challenge yourself with a few of the challenge tasks.

### Before You Begin

1. Create a new repository for this project called `VBA-challenge`. **Do not add this homework to an existing repository**.

2. Inside the new repository that you just created, add any VBA files that you use for this assignment. These will be the main scripts to run for each analysis.

### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - Use this while developing your scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.



## Instructions

Create a script that loops through all the stocks for one year and outputs the following information:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

**Note:** Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

The result should match the following image:



## Bonus

Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:



Make the appropriate adjustments to your VBA script to allow it to run on every worksheet (that is, every year) just by running the VBA script once.
