Sub StockCalculations()

For Each ws In Worksheets                                                           ' Loop throught all worksheets in workbook

                                                                                    ' Declare variables calculate yearly total, percent change,
                                                                                    ' and total volume by UniqueTicker
    Dim i As Long
    Dim lastRow As Long
    Dim UniqueTickerUpdateRow As Long
    Dim yearLastRow As Long
    Dim percentLastRow As Long
    Dim totalVolumeRow As Long

    Dim UniqueTicker As String

    Dim YearlyChange As Double
    Dim totalVolume As Double
    Dim totalYearly As Double
    Dim percentChange As Double
    Dim percent_max As Double
    Dim percent_min As Double
    Dim totalVolumeMax As Double


    ws.Cells(1, 9).Value = "Ticker"                                                 ' Add labels to specific columns
    ws.Cells(1, 16).Value = "Ticker"                                                ' and cells on each worksheet
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"




                                                                                    ' set values for variables in for later calculations
totalVolume = 0                                                                     ' set Total Stock Volume to zero to clear previous calcuations
UniqueTickerUpdateRow = 1                                                           ' set starting location for location iteration
totalYearly = 0                                                                     ' set Total Yearly Change to zero to clear previous calcuations

lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row                                     ' define where the last row of the current worksheet i

                                                                                    ' loop through all rows
    For i = 2 To lastRow                                                            ' define how long the For Loop has to run


        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then                    ' use condition to compare the value of vertically
                                                                                    ' concurrent cells in the
                                                                                    ' cell location (row (i), column(1))[ i.e. Column "A"]
                                                                                    ' starting on row two
    
        
            UniqueTickerUpdateRow = UniqueTickerUpdateRow + 1                       ' define the target row that will get updated
        
        
        
        
            UniqueTicker = ws.Cells(i, 1).Value                                     ' UniqueTicker is initally set as the upper value
                                                                                    ' until the lower value is different from the top
                                                                                    ' and the variable's (UniqueTicker) value is updated with the lower value.
        
        
        
            YearlyChange = ws.Cells(UniqueTickerUpdateRow, 3).Value                 ' YearlyChange is variable initally set with
                                                                                    ' the first Open Value of the unique UniqueTicker
        
        
        
            ws.Cells(UniqueTickerUpdateRow, 9) = UniqueTicker                       ' update the value of the current worksheet's target
                                                                                    ' Cell Location of row (value of UniqueTickerUpdateRos)
                                                                                    ' and column 9 [i.e. Column "I"]

        
        
        
        
            totalYearly = totalYearly + (ws.Cells(i, 6).Value - YearlyChange)       ' update totalYearly value for current UniqueTicker.
                                                                                    ' Since data is structured from oldest to newest
                                                                                    ' per Ticker Symbol, the totalYearly accumulates values
                                                                                    ' if there is a change according to the calculation
            
            totalVolume = totalVolume + ws.Cells(i, 7).Value                        ' update totalVolume value for current UniqueTicker.
                                                                                    ' Since data is structured from oldest to newest
                                                                                    ' per Ticker Symbol, the TotalVolume accumulates values
                                                                                    ' if there is a change according to the calculation
            
            percentChange = (totalYearly / YearlyChange)                            ' cacluates the percent change for that Ticker Symbol
                                                                                    ' for that year.
        
        
            ws.Cells(UniqueTickerUpdateRow, 10) = totalYearly                       ' sets the value of column 10 and the row of the UniqueTicker to
                                                                                    ' the value of totalYearly
                                                                                
            ws.Cells(UniqueTickerUpdateRow, 12).Value = totalVolume                 ' sets the value of column 12 and the row of the UniqueTicker to
                                                                                    ' the value of totalVolume
                                                                                
            ws.Cells(UniqueTickerUpdateRow, 11) = percentChange                     ' sets the value of column 11 and the row of the UniqueTicker to
                                                                                    ' the value of percentChange
            
            ws.Cells(UniqueTickerUpdateRow, 11).Style = "Percent"                   ' format the value of column 11 and the row of the UniqueTracker
                                                                                    ' to display as a percent
            

        
        
        
                                                                                    ' reset values to zero for the next Ticker symbol's cacluations
            totalYearly = 0
            totalVolume = 0

        Else
            totalVolume = totalVolume + ws.Cells(i, 7).Value                        ' set the value of totalVolume while the UniqueTicker value is the
                                                                                    ' same as its vertically concurrent previous cell
        End If
    Next i


yearLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row                                ' determine where the last row is for the Yearly Change column

    For i = 2 To yearLastRow                                                        ' loop through each row in column 10

        If ws.Cells(i, 10).Value >= 0 Then                                          ' set conditional formatting for a positive number to fill the background as green
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3                                 ' set conditional formatting for a number not positive to fill the background as red
        End If
    Next i
    

percent_max = 0                                                                     ' set max percent to zero to clear out any previous calcuations that were run
percent_min = 0                                                                     ' set min percent to zero to clear out any previous calcuations that were run


percentLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row                             ' determine where the last row is for the Percent Change column




    For i = 2 To percentLastRow                                                     ' loop through each row in column 10


        If percent_max < ws.Cells(i, 11).Value Then                                 ' check precent max value against current UniqueTicker
                                                                                    ' value in column 11
        
            percent_max = ws.Cells(i, 11).Value                                     ' Sets new percent max value
            ws.Cells(2, 17).Value = percent_max
        
            ws.Cells(2, 17).Style = "Percent"                                       ' format new value as percent
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value                            ' sets new ticker value
    
        ElseIf percent_min > ws.Cells(i, 11).Value Then                             ' check percent min value against current UniqueTicker
                                                                                    ' value in column 11
            percent_min = ws.Cells(i, 11).Value
        
            ws.Cells(3, 17).Value = percent_min                                     ' Sets new percent min value
        
            ws.Cells(3, 17).Style = "Percent"                                       ' format new value as percent
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value                            ' sets new ticker value
        End If
    Next i

    For i = 2 To percentLastRow                                                     ' loop through each row in column 11

        If ws.Cells(i, 11).Value >= 0 Then                                          ' set conditional formatting for a positive number to fill the background as green
            ws.Cells(i, 11).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 11).Interior.ColorIndex = 3                                 ' set conditional formatting for a number not positive to fill the background as red
        End If
    Next i



totalVolumeRow = ws.Cells(Rows.Count, 12).End(xlUp).Row                             ' determine where the last row is for the Total Stock Volume column

totalVolumeMax = 0                                                                  ' set the totalVolumeMax to zero to clear previous calculations

    For i = 2 To totalVolumeRow                                                     ' loop through each row in column 12


        If totalVolumeMax < ws.Cells(i, 12).Value Then                              ' use condition to determine if current iteration of Total Stock volume
                                                                                    ' is greater than the stored value
            
            totalVolumeMax = ws.Cells(i, 12).Value                                  ' set totalVolumeMax to current iteration's Total Stock Volume value
            ws.Cells(4, 17).Value = totalVolumeMax                                  ' set new total volume max value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value                            ' sets new ticker value
        End If
    Next i
                                                                                         
        
Next ws

End Sub
