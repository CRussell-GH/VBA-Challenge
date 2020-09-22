Sub Stocks()

    'Loops through all worksheets
    For Each ws In Worksheets

        '--------- Declare Variables ---------

        ' Sets an initial variable for stock name
        Dim Stock_Name As String
        ' Sets an initial variable for total per stock
        Dim Stock_Total As Double
        ' Sets an initial variable for opening value
        Dim Stock_Open As Double
        ' Sets an initial variable for closing value
        Dim Stock_Close As Double
        ' Sets an initial variable for change in stock value
        Dim Stock_Change As Double
        ' Sets an initial variable for percentage of stock value change
        Dim Percent_Change As Double
        ' Sets an initial variable for the row to output summary
        Dim Summary_Table_Row As Integer
        
        '-------- Set initial values before looping ------

        ' Sets Last Row for Stocks Data
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' Sets initail total to 0
        Stock_Total = 0
        ' Sets the initial opening value to the first stock
        Stock_Open = ws.Range("C2").Value
        ' Start the summary table under the header row
        Summary_Table_Row = 2
    
        ' ------- Print the Headers to the sheet -------

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"

        ' Loop through all stocks
        For i = 2 To LastRow

            ' Check if we are still within the same stock, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the stock name
                Stock_Name = Cells(i, 1).Value
                ' Add to the Brand Total
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
                ' Set the closing value of the stock
                Stock_Close = ws.Cells(i, 5).Value
                ' Measure the change in the stock
                Stock_Change = Stock_Close - Stock_Open
                ' Avoid Div/0 error with conditional  
                If Stock_Open <> 0 Then
                    ' Calculate change in stock 
                    Percent_Change = Stock_Change / Stock_Open
                        Else
                            Percent_Change = 0
                        End If

                ' Print the Stock Name in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Stock_Name
                ' Print the Stock Total to the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Stock_Total
                ' Print the Stock Total to the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = Stock_Change

                ' change color of cell to reflect gains or losses    
                If Stock_Change > 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf Stock_Change < 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 2
                End If
                ' Change format of change to percentage    
                ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%" 
                ' Print the change in percentage to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Percent_Change

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                ' Reset the Stock Total
                Stock_Total = 0
                ' Set the opening of the next stock
                Stock_Open = Cells(i + 1, 3)
            
            ' If the cell immediately following a row is the same stock...
            Else
                ' Add to the Stock Total
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            End If

        Next i

        '--------- Declare Variables ---------

        ' Sets an initial variable for Greatest Increase
        Dim Increase As Double
        ' Sets an initial variable for Greatest Decrease
        Dim Decrease As Double
        ' Sets an initial variable for Greatest Total Volume
        Dim Total_Volume As Double
        ' Sets an initial variable for the name of the Greatest Increase stock
        Dim IncTicker As String
        ' Sets an initial variable for the name of the Greatest Decrease stock
        Dim DecTicker As String
        ' Sets an initial variable for the name of the Greatest Volume stock
        Dim VolTicker As String

        '-------- Set initial values before looping ------

        ' Sets the Last Row for summary table
        LastTotalRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        ' Sets greatest increase to first stock
        Increase = ws.Cells(2, 12).Value
        ' Sets greatest decrease to first stock
        Decrease = ws.Cells(2, 12).Value
        ' Sets greatest volume to first stock
        Total_Volume = ws.Cells(2, 10).Value
        ' Sets stock name of greatest increase to first stock
        IncTicker = ws.Cells(2, 1).Value
        ' Sets stock name of greatest decrease to first stock
        DecTicker = ws.Cells(2, 1).Value
        ' Sets stock name of greatest volume to first stock
        VolTicker = ws.Cells(2, 1).Value

        ' Start loop on 2nd stock (3rd row under header) because vars initially set to first stock
        For i = 3 To LastTotalRow
            ' Check if stock is greatest increase so far
            If ws.Cells(i, 12).Value > Increase Then
                ' Store increase value
                Increase = ws.Cells(i, 12).Value
                ' Store increase stock name
                IncTicker = ws.Cells(i, 9).Value
            ' Check if stock is greatest decrease so far
            ElseIf ws.Cells(i, 12).Value < Decrease Then
                ' Store decrease value
                Decrease = ws.Cells(i, 12).Value
                ' Store decrease stock name
                DecTicker = ws.Cells(i, 9).Value
                
            End If

            ' Check if stock is greatest increase so far
            If ws.Cells(i, 10).Value > Total_Volume Then
                ' Stote volume value
                Total_Volume = ws.Cells(i, 10).Value
                ' Store volume stock name
                VolTicker = ws.Cells(i, 9).Value
            End If

        Next i

        ' Set format to percentage for Increase and Decrease outputs
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").NumberFormat = "0.00%"

        ' Output values
        ws.Range("O2").Value = IncTicker
        ws.Range("O3").Value = DecTicker
        ws.Range("O4").Value = VolTicker
        ws.Range("P2").Value = Increase
        ws.Range("P3").Value = Decrease
        ws.Range("P4").Value = Total_Volume

    Next ws

End Sub


