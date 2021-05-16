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
     Dim Ticker_row_finish As Integer
     Ticker_row_finish = 2

     Dim Ticker_row_start As Integer
     Ticker_row_start = 2


     ' Define variables for yearly change

    Dim close_value As Double
    Dim open_value As Double 
    Dim Yearly_Change As Double

    
    ' Set numrows = number of rows of data.

        numrows = Range("A1", Range("A1").End(xlDown)).Rows.Count

    'Loop through all the stock/ticker types with a For loop to loop "numrows" number of times

     For i = 2 to numrows

          If Cells( i + 1, 1).Value <> Cells(i, 1).Value Then

         ' Set the stock/ticker name

            Ticker_name = Cells(i, 1).Value

         ' Print the stock/ticker name in the Ticker row

            Range("J" & Ticker_row_finish).Value = Ticker_name

            ' Find the last date

            close_value = Cells(1, 6).Value

                ' Find the first date

                    For j = 2 to numrows

                ' Define start of the date range
                     If Cells( j - 1, 1).Value <> Cells(j, 1).Value Then

                    open_value = Cells(j, 6).Value

                    Ticker_row_start = Ticker_row_start + 1

                End If

            Next j 

            ' Add one to ticker row finish

            Ticker_row_finish = Ticker_row_finish + 1

         End If

         ' Calculate the value of yearly change as close - open

        Yearly_Change = close_value - open_value

        ' Print Yearly_Change in the Yearly Change column

        Range("K" & Ticker_row_finish).Value = Yearly_Change

    Next i


End Sub

