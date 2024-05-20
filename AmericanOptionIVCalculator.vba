Option Base 0

Function AmericanPutPrice(CurrentPrice As Double, StrikePrice As Double, TimeToMaturity As Double, RiskFreeRate As Double, Volatility As Double) As Double
    'Crank Nicolson Method for American Put Option Pricing'
    Dim TimeStepSize As Double, PriceStepSize As Double
    Dim TimeSteps As Integer, PriceSteps As Integer
    Dim PriceIndex As Integer, TimeIndex As Integer
    Dim AlphaCoefficient As Double, BetaCoefficient As Double
    Dim LowerDiagonal() As Double, MainDiagonal() As Double, UpperDiagonal() As Double, RightHandSide() As Double, SolutionVector() As Double
    
    'Set the time and price step sizes'
    TimeStepSize = TimeToMaturity / 100
    PriceStepSize = 0.05 * CurrentPrice
    
    'Calculate grid sizes for time and price dimensions'
    TimeSteps = Round(TimeToMaturity / TimeStepSize, 0)
    PriceSteps = Round(2 * CurrentPrice / PriceStepSize, 0)
    
    'Initialize arrays for matrix and vector components'
    ReDim LowerDiagonal(1 To PriceSteps), MainDiagonal(1 To PriceSteps), UpperDiagonal(1 To PriceSteps), RightHandSide(1 To PriceSteps), SolutionVector(1 To PriceSteps)
    
    'Calculate coefficients for the finite difference scheme'
    AlphaCoefficient = 0.25 * (Volatility * Volatility) * TimeStepSize / (PriceStepSize * PriceStepSize)
    BetaCoefficient = 0.5 * RiskFreeRate * TimeStepSize / PriceStepSize
    
    'Initial condition based on the intrinsic value of the put option'
    For PriceIndex = 1 To PriceSteps
        SolutionVector(PriceIndex) = Application.Max(StrikePrice - PriceIndex * PriceStepSize, 0)
    Next PriceIndex
    
    'Iterate through time steps to update option prices'
    For TimeIndex = 1 To TimeSteps
        'Set boundary conditions for the start and end of the price grid'
        LowerDiagonal(1) = 0
        MainDiagonal(1) = -1
        UpperDiagonal(1) = 1
        RightHandSide(1) = StrikePrice * Exp(-RiskFreeRate * (TimeToMaturity - TimeIndex * TimeStepSize))
        
        LowerDiagonal(PriceSteps) = -1
        MainDiagonal(PriceSteps) = 1
        UpperDiagonal(PriceSteps) = 0
        RightHandSide(PriceSteps) = 0
        
        'Update coefficients for interior grid points based on the Crank-Nicolson scheme'
        For PriceIndex = 2 To PriceSteps - 1
            LowerDiagonal(PriceIndex) = -AlphaCoefficient * (PriceIndex - 0.5) * (PriceIndex - 1)
            MainDiagonal(PriceIndex) = 1 + 2 * AlphaCoefficient * (PriceIndex - 1)
            UpperDiagonal(PriceIndex) = -AlphaCoefficient * (PriceIndex - 0.5) * PriceIndex
            RightHandSide(PriceIndex) = -BetaCoefficient * (PriceIndex - 1) * SolutionVector(PriceIndex - 1) + (1 - BetaCoefficient * (PriceIndex - 1)) * SolutionVector(PriceIndex) - BetaCoefficient * (PriceIndex - 1) * SolutionVector(PriceIndex + 1)
        Next PriceIndex
        
        'Solve the tridiagonal system for the next time step'
        Call SolveTridiagonal(PriceSteps, LowerDiagonal, MainDiagonal, UpperDiagonal, RightHandSide, SolutionVector)
    Next TimeIndex
    
    'Interpolate to find the option price at the current stock price'
    AmericanPutPrice = SolutionVector(CurrentPrice / PriceStepSize)
    
End Function



Sub SolveTridiagonal(numberOfPoints As Integer, diagonalBelow() As Double, diagonal() As Double, diagonalAbove() As Double, RightHandSide() As Double, ByRef SolutionVector() As Double)
    ' Thomas algorithm for solving tridiagonal systems of linear equations.
    ' This method is used for solving linear equations of the form Ax = b where A is a tridiagonal matrix.
    '
    ' Parameters:
    ' numberOfPoints: The size of the tridiagonal system.
    ' diagonalBelow: The sub-diagonal elements of the tridiagonal matrix A.
    ' diagonal: The main diagonal elements of the tridiagonal matrix A.
    ' diagonalAbove: The super-diagonal elements of the tridiagonal matrix A.
    ' rightHandSide: The vector b in the equation Ax = b.
    ' solutionVector: The solution vector x where Ax = b.
    
    Dim modifiedDiagonal() As Double, modifiedRightHandSide() As Double
    Dim i As Integer
    
    ' Initialize modified diagonal and right-hand side arrays.
    ReDim modifiedDiagonal(1 To numberOfPoints), modifiedRightHandSide(1 To numberOfPoints)
    
    ' Forward sweep: modify the coefficients.
    ' Start by copying the first elements.
    modifiedDiagonal(1) = diagonal(1)
    modifiedRightHandSide(1) = RightHandSide(1)
    
    ' Error handling for division by zero in the modified diagonal calculation.
    If modifiedDiagonal(1) = 0 Then
        Err.Raise Number:=vbObjectError + 513, Description:="Division by zero encountered in Thomas algorithm."
    End If
    
    For i = 2 To numberOfPoints
        ' Prevent division by zero.
        If modifiedDiagonal(i - 1) = 0 Then
            Err.Raise Number:=vbObjectError + 514, Description:="Division by zero encountered in Thomas algorithm forward sweep."
        End If
        
        ' Modify the diagonal and right-hand side elements.
        modifiedDiagonal(i) = diagonal(i) - diagonalBelow(i) * diagonalAbove(i - 1) / modifiedDiagonal(i - 1)
        modifiedRightHandSide(i) = RightHandSide(i) - diagonalBelow(i) * modifiedRightHandSide(i - 1) / modifiedDiagonal(i - 1)
    Next i
    
    ' Backward substitution: compute the solution vector.
    SolutionVector(numberOfPoints) = modifiedRightHandSide(numberOfPoints) / modifiedDiagonal(numberOfPoints)
    
    For i = numberOfPoints - 1 To 1 Step -1
        ' Prevent division by zero in the solution vector calculation.
        If modifiedDiagonal(i) = 0 Then
            Err.Raise Number:=vbObjectError + 515, Description:="Division by zero encountered in Thomas algorithm backward substitution."
        End If
        
        SolutionVector(i) = (modifiedRightHandSide(i) - diagonalAbove(i) * SolutionVector(i + 1)) / modifiedDiagonal(i)
    Next i
End Sub




'=============================================================================================================
Public Function impVol(marketPrice As Double, stockPrice As Double, TimeToMaturity As Double, StrikePrice As Double, RiskFreeRate As Double, tolerance As Double) As Variant
    ' This function uses the secant method to compute the implied volatility of an American put option
    ' Inputs:
    '   marketPrice - the current market price of the option
    '   stockPrice - the current market price of the underlying stock
    '   timeToMaturity - time to maturity of the option in years
    '   strikePrice - the strike price of the option
    '   riskFreeRate - the risk-free interest rate as a decimal
    '   tolerance - the tolerance for the convergence of the secant method
    
    Dim lowerVolatilityBound As Double, upperVolatilityBound As Double
    Dim lowerPrice As Double, upperPrice As Double
    Dim midVolatility As Double, midPrice As Double
    Dim iterationIndex As Integer
    Dim maximumIterations As Integer
    
    ' Initialize the lower and upper bounds for the implied volatility
    lowerVolatilityBound = 0.001 ' Setting a small positive number to avoid division by zero
    upperVolatilityBound = impPutVol(marketPrice, stockPrice, TimeToMaturity, StrikePrice, RiskFreeRate) ' Using European put volatility as an initial upper bound
    
    ' Calculate the option price using American pricing model at the initial volatility bounds
    lowerPrice = AmericanPutPrice(stockPrice, StrikePrice, TimeToMaturity, RiskFreeRate, lowerVolatilityBound)
    upperPrice = AmericanPutPrice(stockPrice, StrikePrice, TimeToMaturity, RiskFreeRate, upperVolatilityBound)
    
    ' Set the maximum number of iterations to prevent infinite loops
    maximumIterations = 1000
    
    ' Begin iterations of the secant method
    For iterationIndex = 1 To maximumIterations
        ' Update the mid volatility estimate using the secant formula
        midVolatility = upperVolatilityBound - ((upperPrice - marketPrice) * (upperVolatilityBound - lowerVolatilityBound)) / (upperPrice - lowerPrice)
        ' Calculate the option price at mid volatility
        midPrice = AmericanPutPrice(stockPrice, StrikePrice, TimeToMaturity, RiskFreeRate, midVolatility)
        
        ' Check if the current estimate is within the specified tolerance
        If Abs(midPrice - marketPrice) < tolerance Then
            impVol = midVolatility
            Exit Function
        End If
        
        ' Update bounds for the next iteration based on where the mid price falls
        If midPrice < marketPrice Then
            lowerVolatilityBound = midVolatility
            lowerPrice = midPrice
        Else
            upperVolatilityBound = midVolatility
            upperPrice = midPrice
        End If
    Next iterationIndex
    
    ' If the method fails to converge, return an error
    impVol = CVErr(xlErrNA) ' Return Excel error #N/A to signify non-convergence
End Function


'=============================================================================================================
Function bsput(S As Double, sigma As Double, T As Double, X As Double, r As Double, Optional Div As Double, _
                        Optional ExDate As Date = #1/1/2000#) As Double
'=============================================================================================================

    Dim D1 As Double
    Dim D2 As Double

    D1 = (Log(S / X) + (r + sigma * sigma / 2) * T) / (sigma * T ^ 0.5)
    D2 = D1 - sigma * T ^ 0.5

    bsput = (-S * Application.NormSDist(-D1) + X * Exp(-r * T) * Application.NormSDist(-D2))

'=============================================================================================================
End Function
'=============================================================================================================
Function impPutVol(MktPrice As Double, S As Double, T As Double, X As Double, r As Double, _
                        Optional Div As Double, Optional ExDate As Date = #1/1/2000#) As Double
'=============================================================================================================

    Niter = 10
    If ExDate = #1/1/2000# Then
        impPutVol = (2 * Abs(Log(S / X) + (r - Div) * T) / T) ^ 0.5
    Else
        TDiv = (ExDate - Application.WorksheetFunction.today()) / 256
    End If
    For iter = 1 To Niter
        impPutVol = impPutVol - (bsput(S, impPutVol, T, X, r, Div, ExDate) - _
                                MktPrice) / BSPutVega(S, impPutVol, T, X, r, Div, ExDate)
    Next iter

'=============================================================================================================
End Function
'=============================================================================================================
Function BSPutVega(S As Double, sigma As Double, T As Double, X As Double, r As Double, _
                        Optional Div As Double, Optional ExDate As Date = #1/1/2000#) As Double
'=============================================================================================================
    Dim D1 As Double

    If ExDate = #1/1/2000# Then
        D1 = (Log(S / X) + (r - Div + sigma * sigma / 2) * T) / (sigma * T ^ 0.5)
        BSPutVega = S * T ^ 0.5 * Exp(-D1 * D1 / 2) / (2 * Application.Pi()) ^ 0.5
    Else
    End If
    
'=============================================================================================================
End Function
'=============================================================================================================
