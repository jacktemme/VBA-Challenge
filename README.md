Module 2 Homework

This repository has a VBA script that runs through the yearly stock data in the provided Multiple_year_stock_data.xlsx file. For each stock ticker the script outputs the yearly stock change from the opening value at the beginning of the year to the closing value at the end of the year. The percentage of this change is also outputted along with the yearly total stock volume. Then from  these added value the stock with the greatest percent increase and greatest percent decrease for that year is detrmined. The greatest yearly stock total volume is outputted as well. The repository also includes screenshots for the 2018 to 2020 findings after the VBA script was executed. 


For this script the edX Xpert Learning Assistant AI was referenced for syntax questions. The code below for rounding a value as well as calculating the maximum and minimum value from a range was referenced.

Round Value:

Round the number to two decimal places
    roundedNumber = Round(myNumber, 2)


Calculate Maximim Value:

    Dim rng As Range
    Dim maxVal As Double
    
    ' Set the range where you want to find the maximum value
    Set rng = Range("A1:A10") ' Update this range to your specific range
    
    ' Find the maximum value in the range
    maxVal = Application.WorksheetFunction.Max(rng)



Then code from this site: https://www.excelfunctions.net/vba-formatpercent-function.html was used to change cell value to a rounded percent:

pc3 = FormatPercent( 0.559, 0 )
' pc3 is now equal to the String "56%".
