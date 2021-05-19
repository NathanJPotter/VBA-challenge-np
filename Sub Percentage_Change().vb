Sub Percentage_Change()

   
   ' Here is a script that will loop through all the stocks for one year and output:
    ' - The ticker symbol
    ' - Yearly change from opening at the begining of a given year to the closing price at the end of the year
        ' - including conditional formatting that highlights positive change in green and negative change in red.
     ' - The percentage change from opening at the begining of a given year to the closing price at the end of the year

    ' Add title and format the Ticker column

     Cells(1, 10).Value = "Ticker"

     Cells(1, 10).Font.Bold = True

    ' Add title and format to the Yearly Change column

    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 11).Font.Bold = True

    ' Add titles and format to year open value and year close close culumns

    Cells(1, 15).Value = "Year Open"
    Cells(1, 15).Font.Bold = True

    Cells(1, 16).Value = "Year Close"
    Cells(1, 16).Font.Bold = True

     ' Add title and format to the Yearly Change column

    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 11).Font.Bold = True

    ' Add title and format to the Percentage Change column

    Cells(1, 12).Value = "Percentage Change"
    Cells(1, 12).Font.Bold = True

    ' Add title and format to the Total Stock Volume column

    Cells(1, 13).Value = "Total Stock Volume"
    Cells(1, 13).Font.Bold = True

     ' Define variables for yearly change

        Dim i As Long
        Dim j As Long
        Dim Ticker_tracker As Long
        Dim close_value As Double
        Dim open_value As Double 
        Dim yearly_change As Double
        Dim percentage_change As Double
        Dim total_stock As Double
        Dim close_volume As Double
        Dim open_volume As Double

    'Keep track of the location for each stock/ticker type
         
        Ticker_tracker = 2
    
    ' Set numrows = number of rows of data.

        numrows = Range("A1", Range("A1").End(xlDown)).Rows.Count

    'Use "j" to set beginning of the row to 2

        j = 2
    
    'Loop through all the stock/ticker types with a For loop to loop "numrows" number of times

     For i = 2 to numrows

          If Cells( i + 1, 1).Value <> Cells(i, 1).Value Then

         ' Set the stock/ticker name and print to column J (ie column 10). Here "i" is the last row.

            Cells(Ticker_tracker, 10).Value = Cells(i, 1).Value

        ' For yearly change, set the open_value and close_value

                open_value = Cells(j, 3).Value
               
                close_value = Cells(i, 6).Value

            ' Print open_value to column O (ie col 15) and close_value to column P (ie col 16) to check the number is correct (not part of the exercise)

                Cells(Ticker_tracker, 15).Value = open_value
                
                Cells(Ticker_tracker, 16).Value = close_value

            ' Calculate the value of yearly change as close minus open

                yearly_change = close_value - open_value
                
            ' Print Yearly_Change in the Yearly Change, column K (ie col 11)

                Cells(Ticker_tracker, 11).Value = yearly_change

            ' Calculate the yearly volume, set the open_volume and close_volume
               
                open_volume = Cells(j, 7).Value

                close_volume = Cells(i, 7).Value

            ' Calculate the value of percentage change

                percentage_change = (yearly_change / open_value) * 100

            'Print percentage_change in Percentage Change column L (ie col 12)

                Cells(Ticker_tracker, 12).Value = percentage_change
                        
            ' Calculate the value of Total Stock volume

                            

            'Print total_stock in Total Stock column

                           

            ' Add one to ticker_tracker

                Ticker_tracker = Ticker_tracker + 1

            ' Add one to the beginning row (j)

             j = i + 1

         End If

    Next i    
         
End Sub