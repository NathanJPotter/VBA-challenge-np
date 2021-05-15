Sub Yearly_Change()

    ' Here is a script that will loop through all the stocks for one year and output:
    ' - The ticker symbol
    ' - Yearly change from opening at the begining of a given year to the closing price at the end of the year
        ' - including conditional formatting that highlights positive change in green and negative change in red.

    ' Add title and format the Ticker column

     Cells(1, 10).Value = "Ticker"

     Cells(1, 10).Font.Bold = True

    ' Add title and format to the Yearly Change column

    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 11).Font.Bold = True

    'Keep track of the location for each stock/ticker type
     Dim Ticker_row As Integer
     Ticker_row = 2

    ' Set numrows = number of rows of data.

        numrows = Range("A1", Range("A1").End(xlDown)).Rows.Count

    'Loop through all the stock/ticker types with a For loop to loop "numrows" number of times

     For i = 2 to numrows

          If Cells( i + 1, 1).Value <> Cells(i, 1).Value Then

         ' Set the stock/ticker name

            Ticker_name = Cells(i, 1).Value

         ' Print the stock/ticker name in the Ticker row

            Range("J" & Ticker_row).Value = Ticker_name

            ' Set the year_open value

            year_open = Cells(i, 3).Value

            ' Set the year_close value

            year_close = Cells(i, 6).Value

            ' Calculate yearly_change value as year_close minus year_open

            yearly_change = year_close - year_open

            ' Print yearly_change in the Yearly Change column

            Range("K" & Ticker_row).Value = yearly_change

            ' Add one to ticker row

            Ticker_row = Ticker_row + 1

         End If

    Next i   

End Sub

