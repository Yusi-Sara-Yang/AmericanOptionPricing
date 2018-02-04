Option Base 0

'=============================================================================================================
Public Function impVol(MktPrice As Double, S As Double, sigma As Double, T As Double, X As Double, R As Double, pstep As Integer, tstep As Integer, Smax As Double, Smin As Double) As Double
'=============================================================================================================
    
    'set initial values for volatility
    Dim Lvol As Double, Uvol As Double
    Lvol = 0        'volatility is non-negative
    Uvol = impPutVol(MktPrice, S, T, X, R)      'volatility of american option is smaller than that of european option
    
    'secant method for implied volatility
    Dim tmpvol As Double, tmpOptv As Double
    Dim Optv1 As Double, Optv2 As Double
    'iterate until Am_put converges to the market price (difference < 0.00001)
    Do Until Abs(tmpOptv - MktPrice) < 1E-05
        Optv1 = Am_put(Lvol, T, X, R, pstep, tstep, Smax, Smin)
        Optv2 = Am_put(Uvol, T, X, R, pstep, tstep, Smax, Smin)
        tmpvol = Uvol - (Uvol - Lvol) * (Optv2 - MktPrice) / (Optv2 - Optv1)    'use intercept to find root
        Lvol = Uvol
        Uvol = tmpvol
        tmpOptv = Am_put(tmpvol, T, X, R, pstep, tstep, Smax, Smin)     'give the root to new Am_put price
    Loop
    
    impVol = tmpvol

End Function
'=============================================================================================================
Function Am_put(sigma As Double, T As Double, X As Double, R As Double, pstep As Integer, tstep As Integer, Smax As Double, Smin As Double) As Double
'=============================================================================================================

    Dim i As Integer, j As Integer
    
    ReDim Stock(pstep) As Double
    ReDim Optv(pstep) As Double
    'define each step
    Dim ds As Double, dt As Double
    ds = Log(Smax / Smin) / pstep
    dt = T / tstep
    
    'Boundary value
    For i = 0 To pstep
        Stock(i) = Exp(Log(Smin) + i * ds)
        Optv(i) = WorksheetFunction.Max(X - Stock(i), 0)
    Next i

    'Crank-Nicolson Scheme
    ReDim a(1 To pstep - 1)
    ReDim b(1 To pstep - 1)
    ReDim c(1 To pstep - 1)
    'set the grid in matrix R
    a(1) = -0.25 * ((sigma ^ 2 / ds ^ 2) - (R - 0.5 * sigma ^ 2) / ds)
    b(1) = 1 / dt + 0.5 * R + 0.5 * (sigma / ds) ^ 2
    c(1) = -0.25 * ((sigma ^ 2 / ds ^ 2) + (R - 0.5 * sigma ^ 2) / ds)
    astar = -a(1)
    bstar = 2 / dt - b(1)
    cstar = -c(1)

    'set vector U() for R.V = U, where V is vector Optv
    ReDim U(0 To pstep) As Double
    U(0) = Optv(0)
    U(pstep) = Optv(pstep)
    'set the transition matrix Y
    ReDim Y(0 To pstep) As Double
    Y(0) = Optv(0)

    For i = 1 To tstep
        'calculate of grid in R
        For j = 2 To pstep - 1
            c(j) = c(1)
            a(j) = a(1) / b(j - 1)
            b(j) = -a(j) * c(1) + b(1)
        Next j
       
       'finite difference equation
        For j = 1 To pstep - 1
            U(j) = astar * Optv(j - 1) + bstar * Optv(j) + cstar * Optv(j + 1)
            Y(j) = U(j) - a(j) * Y(j - 1)
        Next j
        Y(pstep) = U(pstep)
        For j = (pstep - 1) To 1 Step (-1)
            U(j) = (Y(j) - c(j) * U(j + 1)) / b(j)
        Next j

        'value of amarican option price for each stock(j)
        For j = 0 To pstep
            Optv(j) = WorksheetFunction.Max(U(j), Optv(j))
        Next j
    Next i
    
    Am_put = U(pstep / 2)

End Function

'=============================================================================================================
Function bsput(S As Double, sigma As Double, T As Double, X As Double, R As Double, Optional Div As Double, _
                        Optional ExDate As Date = #1/1/2000#) As Double
'=============================================================================================================

    Dim D1 As Double
    Dim D2 As Double

    D1 = (Log(S / X) + (R + sigma * sigma / 2) * T) / (sigma * T ^ 0.5)
    D2 = D1 - sigma * T ^ 0.5

    bsput = (-S * Application.NormSDist(-D1) + X * Exp(-R * T) * Application.NormSDist(-D2))

'=============================================================================================================
End Function
'=============================================================================================================
Function impPutVol(MktPrice As Double, S As Double, T As Double, X As Double, R As Double, _
                        Optional Div As Double, Optional ExDate As Date = #1/1/2000#) As Double
'=============================================================================================================

    Niter = 10
    If ExDate = #1/1/2000# Then
        impPutVol = (2 * Abs(Log(S / X) + (R - Div) * T) / T) ^ 0.5
    Else
        TDiv = (ExDate - Application.TODAY()) / 256
    End If
    For iter = 1 To Niter
        impPutVol = impPutVol - (bsput(S, impPutVol, T, X, R, Div, ExDate) - _
                                MktPrice) / BSPutVega(S, impPutVol, T, X, R, Div, ExDate)
    Next iter

'=============================================================================================================
End Function
'=============================================================================================================
Function BSPutVega(S As Double, sigma As Double, T As Double, X As Double, R As Double, _
                        Optional Div As Double, Optional ExDate As Date = #1/1/2000#) As Double
'=============================================================================================================
    Dim D1 As Double

    If ExDate = #1/1/2000# Then
        D1 = (Log(S / X) + (R - Div + sigma * sigma / 2) * T) / (sigma * T ^ 0.5)
        BSPutVega = S * T ^ 0.5 * Exp(-D1 * D1 / 2) / (2 * Application.Pi()) ^ 0.5
    Else
    End If
    
'=============================================================================================================
End Function
'=============================================================================================================
