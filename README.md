# VBA_Challenge
This repo contains completed version for week 2 module challenge for solving the multiple year stock data.

Thought on solving the problem and the bonus question:

For the basic question, I decided to use one main For loop to extract all values from the start row to the end. Then use If statment to create a filter to process the tickers in groups. After the output results. I then use two seperate For Each loop to finsh the conditional formating.

For the bonus question, I initially tried For loop as well by comparing each ticker's percent change and their value one by one to selet the biggest/smallest ticker. However this method won't work as intended due to the percent change cell value formating and very small decimals makes it faulty to select the corect one by comparing. Then I searched the built-in Max/Min/Match function to solve the problem. 

For completing this challenge, some resources I searched and used during the assignment include Excel tutorial webpages and YouTube Excel tutorials.

Here are the sources of the references I used:
#References:
Excel VBA: Referring to Range & Writing to Cells
https://www.youtube.com/watch?v=acGJb9Oojho

Loop through Cells Inside the Range(For Each Loop)
https://www.youtube.com/watch?v=5bq3N99mNPEExcel 

VBA to Search and Highlight Min and Max value in a range
https://www.youtube.com/watch?v=hwKyiq5L-JU

VBA Match Function
https://www.youtube.com/watch?v=syGKtH86ggs

How to change Format of a Cell to Text using VBA
https://stackoverflow.com/questions/8265350/how-to-change-format-of-a-cell-to-text-using-vba

Excel VBA For Complete Beginners 
https://www.homeandlearn.org/worksheet_functions.html#google_vignette
