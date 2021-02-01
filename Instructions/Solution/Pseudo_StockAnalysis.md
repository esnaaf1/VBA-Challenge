Declare variables

Ticker
Total

I for the loop
J for the summary rows for each ticker
FirstRow for the first row of the each ticker symbol

OpenPrice
ClosePrice
YearChange
PctChange
...


J = 2
FirstRow = 2
Total = 0

For i = 2 to LastRow

    Ticker = (Row i, Column 1)
    Total = Total + (Row i, Column 7)  (Stock Volumme)

    If Ticker <> (Row i+1, Column 1) //If this is the last row for the current ticker

            OpenPrice = (Row FirstRow, Column 3)
            ClosePrice = (Row i. Column 6)

            YearChange = ClosePrice - OpenPrice

            PctChange = YearlChange/ OpenPrice ( If OpenPrice is not zero)
            
            'Write the values to the spreadsheet

            Rest the value 

            1) Increment J for the next Ticker
            2) Set the Total = 0
            3) Set FirstRow = i+1 ( For the next Ticker Symbol)
    End If
Next i






