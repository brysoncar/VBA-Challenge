# VBA-Challenge
Module 2 Challenge 

Bryson's VBA Challenge Results

This is a program designed to parse stock value information for an excel document. The program works for an excel document with the following structure
Column A: Ticker Name
Column B: Ticker Name in Yeare/Month/Format 
Column C: The Stock's Open Price
Column D: The Stock's High Value for the Day
Column E: The Stock's Low Value for the Day
Column F: The Stock's Close Price for the Day
Column G: The Stock's Trade Volume for the Day 

The program's output is to list the following information for each ticker present in Column A

Column I: Ticker Name
Column J: The yearly change for a Tickers
Column K: The yearly % change for the stock according to the following formula:
  Yearly Change / The last Day of the Years Close Price
Column L : Total Stock Volume throughout the year

Conditional Formating: Column J cells are colored Red for a decrease in yearly price and green 
for an increase in yearly price. 

The program also finds the following information on a ticker basis for each year of data:
Greatest % Increase of all stocks with the ticker and % increase shown
Greatest % Decrease of all stocks with the ticker and % decrease shown
Greatest Total Volume of all stocks with the ticker and total volume shown.

This program will work on multiple years if the infortmation is separated by year and each year is a different tab. 

Things I wish to improve on:
1. The process to create Yearly Change and Percent Change are simple calculations on two columns I created that list the first day of the year's open price and the last day of the year's close price. I would rather store the values in variables and create Yearly Change and Percent Change from those, but I had problems getting the output on the correct row.
2. The for loops that used the output I created from column I to column L has a hard coded range instead of running to the last row in those columns. While the analysis of the data given for the project is correct and complete, if the number of tickers in column a for a new set of data exceeds 3000, the analysis will be incomplete. I plan on getting some help and working on this in the future. 
the 
