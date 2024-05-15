Sub BackTesting()
    Dim lastRow As Long, outputRow As Integer, windowFrame As Integer
    Dim startDate As Date, endDate As Date
    Dim portfolioValue As Double
    Dim searchRange As Range
    
    portfolioValue = Range("INVESTMENT_AMT").Value
    windowFrame = Range("WINDOW_FRAME").Value  ' Window size for calculations
    
    lastRow = WorksheetFunction.Max(Cells(Rows.Count, 4).End(xlUp).Row, _
                                    Cells(Rows.Count, 5).End(xlUp).Row)

    SetupHeaders  ' Call subroutine to setup headers

    outputRow = 11 ' Starting output row just below the headers
    Dim i As Integer
    Dim index As Integer
    Dim endDateRow As Variant
    Dim cashRate As Double
    i = 45  ' Start from H45
    
   For Each dateCell In Range(Range("K7").Value)
        endDate = dateCell.Value
        If IsDate(endDate) Then
             endDateRow = Application.Match(dateCell, Range("A10:A" & lastRow), 0)
             ' endDateRow = Application.Match(endDate, Range("A10:A" & Cells(Rows.Count, "A").End(xlUp).Row), 0)
            cashRate = 0 'Cells(i, 6).Value
            PerformCalculations endDateRow - 1, lastRow, outputRow, windowFrame, portfolioValue, False, cashRate ' One day before
            PerformCalculations endDateRow, lastRow, outputRow, windowFrame, portfolioValue, True, cashRate
            PerformCalculations endDateRow + 11, lastRow, outputRow, windowFrame, portfolioValue, False, cashRate ' One day after
        End If
        i = i + 1
        
    Next dateCell

    FormatOutputTable outputRow ' Call subroutine to format output table
End Sub

Sub PerformCalculations(ByVal endDateRow As Variant, ByVal lastRow As Long, ByRef outputRow As Integer _
, ByVal windowFrame As Integer, ByVal portfolioValue As Double, ByVal isBoardMeetingDay As Boolean, cashRate As Double)
    Dim startDateRow As Long
    ' endDateRow = Application.Match(endDate, Range(Cells(i, 8) & lastRow), 0)
    If Not IsError(endDateRow) Then
        endDateRow = endDateRow + 9  ' Adjust for base at A10
        If endDateRow - windowFrame >= 9 Then  ' Ensure enough days are available
            startDateRow = endDateRow - windowFrame + 1
            CalculateAndDisplayStatistics Range("D" & startDateRow & ":D" & endDateRow), _
                                           Range("E" & startDateRow & ":E" & endDateRow), _
                                           outputRow, startDateRow, endDateRow, portfolioValue, isBoardMeetingDay, windowFrame, cashRate
        End If
    End If
End Sub

Sub CalculateAndDisplayStatistics(rangeD As Range, rangeE As Range, ByRef outputRow As Integer, _
                                  ByVal startDateRow As Long, ByVal endDateRow As Long, ByVal portfolioValue As Double, _
                                  ByVal isBoardMeetingDay As Boolean, ByVal windowFrame As Integer, cashRate As Double)
    ' Calculation of statistics
    Dim varD As Double, varE As Double, covDE As Double, portfolioStdDev As Double, portfolioReturns As Double
    varD = Application.WorksheetFunction.Var(rangeD)
    varE = Application.WorksheetFunction.Var(rangeE)
    covDE = Application.WorksheetFunction.Covar(rangeD, rangeE)
    portfolioStdDev = Sqr(0.25 * varD + 0.25 * varE + 2 * 0.5 * 0.5 * covDE)
    portfolioReturns = 0.5 * Application.WorksheetFunction.Average(rangeD) + 0.5 * Application.WorksheetFunction.Average(rangeE)
    
    ' Display results
    With activeSheet
    .Cells(outputRow, 8).Value = outputRow - windowFrame
    .Cells(outputRow, 9).Value = .Cells(startDateRow, 1).Value & " to " & .Cells(endDateRow, 1).Value
     ' Perform calculations
                    .Cells(outputRow, 10).Value = Application.WorksheetFunction.Average(rangeD)
                    .Cells(outputRow, 11).Value = Application.WorksheetFunction.Median(rangeD)
                    .Cells(outputRow, 12).Value = Application.WorksheetFunction.StDev(rangeD)
                    .Cells(outputRow, 13).Value = Application.WorksheetFunction.Var(rangeD)
                    .Cells(outputRow, 14).Value = Application.WorksheetFunction.Average(rangeE)
                    .Cells(outputRow, 15).Value = Application.WorksheetFunction.Median(rangeE)
                    .Cells(outputRow, 16).Value = Application.WorksheetFunction.StDev(rangeE)
                    .Cells(outputRow, 17).Value = Application.WorksheetFunction.Var(rangeE)
                    .Cells(outputRow, 18).Value = portfolioStdDev
                    .Cells(outputRow, 19).Value = portfolioReturns

    ' Highlight the endDate row only within the specified columns
    If isBoardMeetingDay Then
        .Range("H" & outputRow & ":Z" & outputRow).Interior.Color = RGB(209, 241, 218)   ' Light pink color
    End If
    DisplayVaR .Cells, portfolioStdDev, portfolioValue, outputRow  ' Assuming VaR-related cells start at column 20
     ' Output Cash Rate from column F
    .Cells(outputRow, 26).Value = .Cells(endDateRow, 6).Value ' Cash Rate from Column F
                    
    outputRow = outputRow + 1
End With

 
End Sub

Sub SetupHeaders()
    With Range("H10:Z10")
        .Value = Array("Window", "Date Range", "AUD/USD Mean", "AUD/USD Median", "AUD/USD StdDev", "AUD/USD Variance", _
                       "AUD/JPY Mean", "AUD/JPY Median", "AUD/JPY StdDev", "AUD/JPY Variance", _
                       "Portfolio StdDev", "Portfolio Returns", "90% VaR %", "90% VaR $", _
                       "95% VaR %", "95% VaR $", "99% VaR %", "99% VaR $", "Cash Rate")
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
End Sub

Sub FormatOutputTable(ByVal outputRow As Integer)
    With Range("H10:Z" & outputRow - 1)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
End Sub

Sub DisplayVaR(targetCell As Variant, portfolioStdDev As Double, portfolioValue As Double, outputRow As Integer)
     ' VaR calculations for each confidence level
                    Dim confLevels As Variant, zs As Variant
                    confLevels = Array(0.9, 0.95, 0.99)
                    zs = Array(-1.282, -1.645, -2.326)  ' Z-scores for 90%, 95%, and 99%
                    
                    Dim j As Integer
                    For j = 0 To 2
                            ' Set VaR percentage, adjusted to be displayed as percentage
                            With targetCell(outputRow, 20 + j * 2)
                                .Value = -(zs(j) * portfolioStdDev) ' Assuming this is a proportion
                                .NumberFormat = "0.00%" ' Formats the cell as a percentage
                            End With
                     'targetCell(outputRow, 20 + j * 2).Value = portfolioStdDev  ' Example calculation
    '               targetCell.NumberFormat = "0.00%"
                        ' Set VaR monetary value
                        With targetCell(outputRow, 21 + j * 2)
                            .Value = -(zs(j) * portfolioStdDev) * portfolioValue
                            .NumberFormat = "#,##0.00" ' Formats the cell as a number with two decimal places
                        End With
                    Next j
                    
    
    ' targetCell.Value = portfolioStdDev  ' Example calculation
    ' targetCell.NumberFormat = "0.00%"
    ' Extend for $ calculations and other VaR metrics
End Sub