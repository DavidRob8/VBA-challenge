# VBA-challenge

Running this macro will add new columns I, J, K, and L with headers named Ticker, Yearly Change, Percent Change, and Total Stock Volume. 
Then, it will also show the ticker and value in columns O, P, and Q for the greatest percentage increase, the greatest percentage decrease, and the greatest total volume.

The code starts with the creation of a sub routine (Stock_Market) and by ensuring that the macro run will run in each sheet in the workbook.
I then declared the necessary variables and assigned values to the variables and also added some formatting functionality (ex: percent, decimal count, column width adjustments).

I then ran a For loop and an If statement to add accurate and correct data to the new Ticker, Yearly Change, Percent Change, and Total Stock Volume columns, ensuring that all data from each Ticker was taken into account when populating these new columns.
For the Yearly Change column, I added conditional formatting that would highlight positive changes in green and negative changes in red.
Then I ended the If statement and the For Loop.

Next, I declared the necessary variables and assigned values to the variables once again for new columns O, P, and Q to show the greatest percentage increase, the greatest percentage decrease, and the greatest total volume.
I then started another For loop and a few If statements within the loop to grab the requested information from columns A to G to new columns O, P, and Q. The If statements were ended and the For loop was ended.

Finally, the sub routine was ended to complete the code.
