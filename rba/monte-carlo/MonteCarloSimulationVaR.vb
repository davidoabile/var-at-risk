Sub MonteCarloSimulationVaR()
 Dim numSimulations As Long
    Dim i As Long
    Dim cashRateRange As Range
    Dim VaR90Range As Range, VaR95Range As Range, VaR99Range As Range
    Dim simulatedCashRate As Double
    Dim stdDevCashRate As Double, meanCashRate As Double
    Dim stdDevVaR90 As Double, meanVaR90 As Double
    Dim stdDevVaR95 As Double, meanVaR95 As Double
    Dim stdDevVaR99 As Double, meanVaR99 As Double
    Dim totalSimulatedVaR90 As Double, totalSimulatedVaR95 As Double, totalSimulatedVaR99 As Double
    Dim totalSimulatedVaR_Dollar90 As Double, totalSimulatedVaR_Dollar95 As Double, totalSimulatedVaR_Dollar99 As Double
    Dim squaredDiffSumVaR90 As Double, squaredDiffSumVaR95 As Double, squaredDiffSumVaR99 As Double
    Dim avgSimulatedVaR90 As Double, avgSimulatedVaR95 As Double, avgSimulatedVaR99 As Double
    Dim stdDevSimulatedVaR90 As Double, stdDevSimulatedVaR95 As Double, stdDevSimulatedVaR99 As Double
    Dim avgSimulatedVaR_Dollar90 As Double, avgSimulatedVaR_Dollar95 As Double, avgSimulatedVaR_Dollar99 As Double
    Dim stdDevSimulatedVaR_Dollar90 As Double, stdDevSimulatedVaR_Dollar95 As Double, stdDevSimulatedVaR_Dollar99 As Double
    Dim portfolioValue As Double
    Dim simulatedVaR90 As Double, simulatedVaR95 As Double, simulatedVaR99 As Double
    Dim simulatedVaR_Dollar90 As Double, simulatedVaR_Dollar95 As Double, simulatedVaR_Dollar99 As Double
    Dim rateSensitivity90 As Double, rateSensitivity95 As Double, rateSensitivity99 As Double


    ' Set the number of simulations, this can be changed in the worksheet Inputs & Outputs
    numSimulations = Range("NUM_SIMULATIONS").Value
    portfolioValue = Range("INVESTMENT_AMT").Value

    ' Define ranges for original data
    Set cashRateRange = Range("Z11:Z" & Cells(Rows.Count, 26).End(xlUp).Row)
    Set VaR90Range = Range("T11:T" & Cells(Rows.Count, 20).End(xlUp).Row)
    Set VaR95Range = Range("V11:V" & Cells(Rows.Count, 22).End(xlUp).Row)
    Set VaR99Range = Range("X11:X" & Cells(Rows.Count, 24).End(xlUp).Row)

   ' Calculate mean and standard deviation for Cash Rate and VaR
    meanCashRate = Application.WorksheetFunction.Average(cashRateRange)
    stdDevCashRate = Application.WorksheetFunction.StDev_S(cashRateRange)
    meanVaR90 = Application.WorksheetFunction.Average(VaR90Range)
    stdDevVaR90 = Application.WorksheetFunction.StDev_S(VaR90Range)
    meanVaR95 = Application.WorksheetFunction.Average(VaR95Range)
    stdDevVaR95 = Application.WorksheetFunction.StDev_S(VaR95Range)
    meanVaR99 = Application.WorksheetFunction.Average(VaR99Range)
    stdDevVaR99 = Application.WorksheetFunction.StDev_S(VaR99Range)

    ' Calculate sensitivity of VaR to interest rate changes
    If stdDevCashRate = 0 Then
        stdDevCashRate = 1
    End If
    
    rateSensitivity90 = stdDevVaR90 / stdDevCashRate
    rateSensitivity95 = stdDevVaR95 / stdDevCashRate
    rateSensitivity99 = stdDevVaR99 / stdDevCashRate


    ' Initialize random seed
    Randomize

    ' Initialize total for averages
    totalSimulatedVaR90 = 0
    totalSimulatedVaR95 = 0
    totalSimulatedVaR99 = 0
    totalSimulatedVaR_Dollar90 = 0
    totalSimulatedVaR_Dollar95 = 0
    totalSimulatedVaR_Dollar99 = 0

    ' First pass: Calculate total sums for mean calculation
    For i = 1 To numSimulations
        ' Generate random changes for cash rate using a normal distribution assumption
        ' Here, Norm_S_Inv is used to generate a percentile from a standard normal distribution, 
        ' which is then scaled by the standard deviation of the cash rate and adjusted by its mean.
        ' This simulates possible future cash rate changes based on historical volatility and central tendency.
        simulatedCashRate = meanCashRate + stdDevCashRate * Application.WorksheetFunction.Norm_S_Inv(Rnd)

        ' Calculate simulated VaR based on the new simulated cash rate
        ' The new VaR is calculated by adjusting the mean VaR by a factor that represents how much the 
        ' simulated change in the cash rate deviates from the mean, scaled by the previously calculated sensitivity.
        ' This models the impact of an unexpected change in interest rates on the portfolio's VaR.
        simulatedVaR90 = meanVaR90 + rateSensitivity90 * (simulatedCashRate - meanCashRate) * Application.WorksheetFunction.Norm_S_Inv(0.1)
        simulatedVaR95 = meanVaR95 + rateSensitivity95 * (simulatedCashRate - meanCashRate) * Application.WorksheetFunction.Norm_S_Inv(0.05)
        simulatedVaR99 = meanVaR99 + rateSensitivity99 * (simulatedCashRate - meanCashRate) * Application.WorksheetFunction.Norm_S_Inv(0.01)
        
        
        simulatedVaR_Dollar90 = simulatedVaR90 * portfolioValue
        simulatedVaR_Dollar95 = simulatedVaR95 * portfolioValue
        simulatedVaR_Dollar99 = simulatedVaR99 * portfolioValue

        ' Accumulate totals for averaging
        totalSimulatedVaR90 = totalSimulatedVaR90 + simulatedVaR90
        totalSimulatedVaR95 = totalSimulatedVaR95 + simulatedVaR95
        totalSimulatedVaR99 = totalSimulatedVaR99 + simulatedVaR99
        totalSimulatedVaR_Dollar90 = totalSimulatedVaR_Dollar90 + simulatedVaR_Dollar90
        totalSimulatedVaR_Dollar95 = totalSimulatedVaR_Dollar95 + simulatedVaR_Dollar95
        totalSimulatedVaR_Dollar99 = totalSimulatedVaR_Dollar99 + simulatedVaR_Dollar99
    Next i

    ' Calculate averages
    avgSimulatedVaR90 = totalSimulatedVaR90 / numSimulations
    avgSimulatedVaR95 = totalSimulatedVaR95 / numSimulations
    avgSimulatedVaR99 = totalSimulatedVaR99 / numSimulations
    avgSimulatedVaR_Dollar90 = totalSimulatedVaR_Dollar90 / numSimulations
    avgSimulatedVaR_Dollar95 = totalSimulatedVaR_Dollar95 / numSimulations
    avgSimulatedVaR_Dollar99 = totalSimulatedVaR_Dollar99 / numSimulations

    ' Second pass: Calculate squared differences for standard deviation
    squaredDiffSumVaR90 = 0
    squaredDiffSumVaR95 = 0
    squaredDiffSumVaR99 = 0
    For i = 1 To numSimulations
        ' Generate random changes for cash rate using normal distribution
        simulatedCashRate = meanCashRate + stdDevCashRate * Application.WorksheetFunction.Norm_S_Inv(Rnd)

        ' Calculate simulated VaR for each confidence level based on the new simulated cash rate
        ' We are not using RND here, to avoid generating different random numbers for the same iteration
        ' which may chnage the results of the standard deviation calculation.
        simulatedVaR90 = meanVaR90 + rateSensitivity90 * (simulatedCashRate - meanCashRate)
        simulatedVaR95 = meanVaR95 + rateSensitivity95 * (simulatedCashRate - meanCashRate)
        simulatedVaR99 = meanVaR99 + rateSensitivity99 * (simulatedCashRate - meanCashRate)
        
        
        ' Accumulate squared differences for standard deviation calculation
        squaredDiffSumVaR90 = squaredDiffSumVaR90 + (simulatedVaR90 - avgSimulatedVaR90) ^ 2
        squaredDiffSumVaR95 = squaredDiffSumVaR95 + (simulatedVaR95 - avgSimulatedVaR95) ^ 2
        squaredDiffSumVaR99 = squaredDiffSumVaR99 + (simulatedVaR99 - avgSimulatedVaR99) ^ 2
    Next i

    ' Calculate standard deviations
    stdDevSimulatedVaR90 = Sqr(squaredDiffSumVaR90 / (numSimulations - 1))
    stdDevSimulatedVaR95 = Sqr(squaredDiffSumVaR95 / (numSimulations - 1))
    stdDevSimulatedVaR99 = Sqr(squaredDiffSumVaR99 / (numSimulations - 1))
    stdDevSimulatedVaR_Dollar90 = stdDevSimulatedVaR90 * portfolioValue
    stdDevSimulatedVaR_Dollar95 = stdDevSimulatedVaR95 * portfolioValue
    stdDevSimulatedVaR_Dollar99 = stdDevSimulatedVaR99 * portfolioValue


    ' Output the averages to the specific table starting from AC10
    With activeSheet
        .Cells(14, 30).Value = avgSimulatedVaR_Dollar90
        .Cells(14, 31).Value = avgSimulatedVaR_Dollar95
        .Cells(14, 32).Value = avgSimulatedVaR_Dollar99
        .Range("AC14:AF14").Font.Bold = True
        .Range("AC14:AF14").Interior.Color = RGB(255, 255, 0) ' Yellow background
        .Cells(15, 30).Value = stdDevSimulatedVaR_Dollar90
        .Cells(15, 31).Value = stdDevSimulatedVaR_Dollar95
        .Cells(15, 32).Value = stdDevSimulatedVaR_Dollar99
       
    End With

    MsgBox "Monte Carlo simulation completed successfully!", vbInformation
End Sub