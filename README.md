# VBA-challenge

*Background:
In this project, I used VBA scripting to analyze generated stock market data

*Steps:

1. I Created a script that loops through all the stocks for one year and outputs the following information:
  -The ticker symbol
  -Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
  -The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
  -The total stock volume of the stock.

2. Added functionality to script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
3. Made the appropriate adjustments to VBA script to enable it to run on every worksheet (that is, every year) at once.

*Source code:
'To loop through all worksheets https://www.automateexcel.com/vba/cycle-and-update-all-worksheets/
'To loop to Dynamic last row http://www.everything-excel.com/loop-to-last-row
'Referencing https://www.wallstreetmojo.com/vba-range/#h-example-3-select-an-entire-column
'Conditional formatting https://www.wallstreetmojo.com/vba-conditional-formatting/
'Finding max and min https://www.wallstreetmojo.com/vba-max/
'Column autofit https://excelchamps.com/vba/autofit/
