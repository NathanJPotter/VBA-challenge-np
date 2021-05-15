' Here is a script that will loop through all the stocks for one year and output:
' - The ticker symbol
' - Yearly change from opening at the begining of a given year to the closing price at the end of the year
'       - including conditional formatting that highlights positive change in green and negative change in red.
' - The percentage change from opening at the begining of a given year to the closing price at the end of the year
' - The total stock volume of the stock

Sub Ticker_Symbol()

    'Keep track of the location for each stock/ticker type 

     Dim Ticker_row As Integer
     Ticker_row = 2

    ' Add title and format the Ticker column

     Cells(1, 10).Value = "Ticker"

     Cells(1, 10).Font.Bold = True

    ' Set numrows = number of rows of data.

        numrows = Range("A2", Range("A1").End(xlDown)).Rows.Count

    'Loop through all the stock/ticker types with a For loop to loop "numrows" number of times

     For i = 2 to numrows

          If Cells( i + 1, 1).Value <> Cells(i, 1).Value Then

         ' Set the stock/ticker name

            Ticker_name = Cells(i, 1).Value

         ' Print the stock/ticker name in the Ticker row

            Range("J" & Ticker_row).Value = Ticker_name

            ' Add one to ticker row

            Ticker_row = Ticker_row + 1

         End If

    Next i   

End Sub