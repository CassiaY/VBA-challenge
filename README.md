# VBA-challenge
module2_challenge  
Contributor: Cassia Yoon

For this assignment, I was provided an Excel file (Multiple_year_stock_data.xlsx) containing daily stock market data from 2018 to 2020 (one year per worksheet). I was tasked with using VBA script to analyze the data in two parts:  
--The first part consisted of summarizing each ticker's yearly change, percent change, and total stock volume for the year.  
--The second part consisted of finding the tickers with the greatest percent increase, decrease, and total volume for that year.

My code can be found in the file Cyoon_Challenge2.vbs  
The script can be imported into a macro-enabled Excel file containing the raw stock market data (need to save Multiple_year_stock_data.xlsx as an .xlsm file). Screenshots of my code output for each worksheet are also included (Cyoon_Challenge2_1.png, Cyoon_Challenge2_2.png, Cyoon_Challenge2_3.png) in the same folder.

My script uses for loops to analyze the data and output the required information for the two parts described above. It automatically finds the last row of data in each worksheet to determine the end of the for loop. It also uses a nested if statement to color code the yearly changes red if negative and green if positive. The code formats the spreadsheet columns and creates headers to contain the results. The code loops through all three worksheets.

Thank you to my instructor and TAs for teaching me how to code most of what I needed to complete the assignment. I consulted https://www.mrexcel.com/board/threads/loop-until-last-row-in-spreadsheet.302362/ for reference on how to loop to the last row. I also consulted https://stackoverflow.com/questions/50989803/for-each-ws-loop-not-switching-sheets to learn how to get my code to loop through all the worksheets.
