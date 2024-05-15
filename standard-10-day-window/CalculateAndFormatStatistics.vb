Sub CalculateAndFormatStatistics()
    Dim i As Integer
    Dim lastRow As Long
    Dim rangeD As Range, rangeE As Range
    Dim outputRow As Integer
    Dim portfolioValue As Double

    
    ' Example portfolio value
    portfolioValue = 1000000  ' Change as per your actual portfolio value

    ' Define the last row of data
    lastRow = WorksheetFunction.Max(Cells(Rows.Count, 4).End(xlUp).Row, _
                                    Cells(Rows.Count, 5).End(xlUp).Row)

    ' Set headers
    With Range("H10:Z10")
        .Value = Array("Window", "Date", "AUD/USD Mean", "AUD/USD Median", "AUD/USD StdDev", "AUD/USD Variance", _
                       "AUD/JPY Mean", "AUD/JPY Median", "AUD/JPY StdDev", "AUD/JPY Variance", _
                       "Portfolio StdDev", "Portfolio Returns", "90% VaR %", "90% VaR $", _
                       "95% VaR %", "95% VaR $", "99% VaR %", "99% VaR $", "Cash Rate")
        .Interior.Color = RGB(0, 112, 192) ' Dark Blue background
        .Font.Color = RGB(255, 255, 255) ' White font color
        .Font.Bold = True
    End With

    outputRow = 11 ' Starting output row

    ' Loop through every 10 rows
    For i = 11 To lastRow Step 10
        Set rangeD = Range("D" & i & ":D" & WorksheetFunction.Min(i + 9, lastRow))
        Set rangeE = Range("E" & i & ":E" & WorksheetFunction.Min(i + 9, lastRow))

        ' Portfolio calculations
        Dim stdDevD As Double, stdDevE As Double, meanD As Double, meanE As Double
        Dim covDE As Double, portfolioStdDev As Double, portfolioReturns As Double
        Dim varD As Double, varE As Double
        
        stdDevD = Application.WorksheetFunction.StDev(rangeD)
        stdDevE = Application.WorksheetFunction.StDev(rangeE)
        meanD = Application.WorksheetFunction.Average(rangeD)
        meanE = Application.WorksheetFunction.Average(rangeE)
        varD = Application.WorksheetFunction.Var(rangeD)
        varE = Application.WorksheetFunction.Var(rangeE)
        
        covDE = Application.WorksheetFunction.Covar(rangeD, rangeE)
        portfolioStdDev = Sqr(0.25 * varD + 0.25 * varE + 2 * 0.5 * 0.5 * covDE)
        portfolioReturns = 0.5 * meanD + 0.5 * meanE

 ' Output window number and Date
        Cells(outputRow, 8).Value = (outputRow - 11) + 1 ' Window number
        Cells(outputRow, 9).Value = Cells(i, 1).Value ' Date from Column A
        
        ' Calculate and output statistics for AUD/USD
        Cells(outputRow, 10).Value = Application.WorksheetFunction.Average(rangeD)
        Cells(outputRow, 11).Value = Application.WorksheetFunction.Median(rangeD)
        Cells(outputRow, 12).Value = Application.WorksheetFunction.StDev(rangeD)
        Cells(outputRow, 13).Value = Application.WorksheetFunction.Var(rangeD)
        
        ' Calculate and output statistics for AUD/JPY
        Cells(outputRow, 14).Value = Application.WorksheetFunction.Average(rangeE)
        Cells(outputRow, 15).Value = Application.WorksheetFunction.Median(rangeE)
        Cells(outputRow, 16).Value = Application.WorksheetFunction.StDev(rangeE)
        Cells(outputRow, 17).Value = Application.WorksheetFunction.Var(rangeE)
        ' Output the calculated statistics
        Cells(outputRow, 18).Value = portfolioStdDev
        Cells(outputRow, 19).Value = portfolioReturns
        
        ' VaR calculations for each confidence level
        Dim confLevels As Variant, zs As Variant
        confLevels = Array(0.9, 0.95, 0.99)
        zs = Array(-1.282, -1.645, -2.326)  ' Z-scores for 90%, 95%, and 99%
        
        Dim j As Integer
        For j = 0 To 2
                ' Set VaR percentage, adjusted to be displayed as percentage
                With Cells(outputRow, 20 + j * 2)
                    .Value = -(zs(j) * portfolioStdDev) ' Assuming this is a proportion
                    .NumberFormat = "0.00%" ' Formats the cell as a percentage
                End With
        
            ' Set VaR monetary value
            With Cells(outputRow, 21 + j * 2)
                .Value = -(zs(j) * portfolioStdDev) * portfolioValue
                .NumberFormat = "#,##0.00" ' Formats the cell as a number with two decimal places
            End With
    
          '  Cells(outputRow, 20 + j * 4).Value = confLevels(j)
          '  Cells(outputRow, 21 + j * 4).Value = zs(j)
           ' Cells(outputRow, 20 + j * 2).Value = Format(-(zs(j) * portfolioStdDev), "0.00")      '- (zs(j) * portfolioStdDev)
           ' Cells(outputRow, 21 + j * 2).Value = Format(-(zs(j) * portfolioStdDev) * portfolioValue)
        Next j
        
        ' Output Cash Rate from column F
        Cells(outputRow, 26).Value = Cells(i, 6).Value ' Cash Rate from Column F
        
        ' Increment the row for output
        outputRow = outputRow + 1
    Next i

    ' Format the output table
    With Range("H10:Z" & outputRow - 1)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
End Sub




