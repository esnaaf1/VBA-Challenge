Sub StockAnalysis()

    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim TotalVolume As Double
    Dim YearChange As Double
    Dim PctChange As Double
    
    Dim Ticker As String
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim LastRow As Long
      
    Dim EndYearPrice As Double
    
    For Each ws In Worksheets
    
              ws.Activate
    
            ' Initialize the variables
            j = 2
            k = 2
            
            TotalVolume = 0
            
            
            ' Write Summary Column Headers
            
            Range("I" & 1).Value = "Ticker"
            Range("J" & 1).Value = "Yearly Change"
            Range("K" & 1).Value = "Percent Change"
            Range("L" & 1).Value = "Total Wtock Volume"
            
            'Determine the last non-empty row
            LastRow = Cells(Rows.Count, 1).End(xlUp).Row
            
            ' Loop through the rows
            
            For i = 2 To LastRow
            
                
                Ticker = Cells(i, 1).Value
                TotalVolume = TotalVolume + Cells(i, 7).Value
                
                ' Check to see if we next row has different value
                
                If Ticker <> Cells(i + 1, 1) Then
                
                    ' We want to write the summary data for the ticker in this "IF Then Block"
                    
                    ' End Year Price
                    BegyearPrice = Range("C" & k).Value
                    EndYearPrice = Range("F" & i).Value
                    YearChange = EndYearPrice - BegyearPrice
                    
                    ' account for zero values
                    If (BegyearPrice = 0 Or EndYearPrice = 0) Then
                        PctChange = 0
                    Else
                        PctChange = EndYearPrice / BegyearPrice - 1
                    End If
                    
                    ' Write the ticker summary to columns I - L
                    Range("I" & j).Value = Ticker
                    Range("J" & j).Value = YearChange
                    Range("K" & j).Value = PctChange
                    Range("L" & j).Value = TotalVolume
                   
                   ' Conditionally color the YearlyChange Column
                   
                   If YearChange < 0 Then
                   
                        Range("J" & j).Interior.Color = vbRed
                 
                   Else
                    
                        Range("J" & j).Interior.Color = vbGreen
                    
                    End If
                      
                     ' After we write the Summary Ticker data, we increment the row for the summary data,
                     ' And we reset the Total Volume, and set the value of the k to the first row of the next Ticker
                     ' In order to keep track of the Openning Stock Price
                    j = j + 1
                    TotalVolume = 0
                    k = i + 1
                    
                End If
            
            Next i
            
            ' Format the Percent Change Column K
            
            Range("K:K").NumberFormat = "0.00%"
            Columns("L").AutoFit
            
            
            '**********Bonus Section **********
            LastSumRow = Cells(Rows.Count, 9).End(xlUp).Row

            GIncrease = Range("K" & 2).Value
            GDecrease = Range("K" & 2).Value
            GVolume = Range("L" & 2).Value

            For i = 3 To LastSumRow

                If Range("K" & i).Value > GIncrease Then
                    GIncrease = Range("K" & i).Value
                    ITicker = Range("I" & i).Value
                End If

                If Range("K" & i).Value < GDecrease Then
                    GDecrease = Range("K" & i).Value
                    DTicker = Range("I" & i).Value
                End If

                If Range("L" & i).Value > GVolume Then
                    GVolume = Range("L" & i).Value
                    VTicker = Range("I" & i).Value
                End If

            Next i
            
            ' Write the Greatest/Least Percent Change and the Greatest Total Volume
            
            Range("P" & 1).Value = "Ticker"
            Range("Q" & 1).Value = "Value"
            
            Range("O" & 2).Value = "Greatest % Incease"
            Range("P" & 2).Value = ITicker
            Range("Q" & 2).Value = GIncrease
            
            Range("O" & 3).Value = "Greatest % Decrease"
            Range("P" & 3).Value = DTicker
            Range("Q" & 3).Value = GDecrease
            
            Range("O" & 4).Value = "Greatest Total Volume"
            Range("P" & 4).Value = VTicker
            Range("Q" & 4).Value = GVolume
            
            ' Format the columns
            Range("Q2:Q3").NumberFormat = "0.00%"
            Columns("O").AutoFit
            Columns("P").AutoFit
            Columns("Q").AutoFit
    
    Next ws
    
    
End Sub
