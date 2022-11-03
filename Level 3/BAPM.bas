Attribute VB_Name = "BAPM"
'Student name: Jianwei Su
'Date: 10/31/2022
'HW 3
'BAPM option pricer function


Option Explicit

Function BAPMPrice(flavor As String, S As Double, q As Double, r As Double, sigma As Double, period As Double, T As Double, N As Integer, K As Double) As Double


'flavor: choose call or put option
'S: underlying asset stock spot price
'q: dividend yield. It is used to be equivalent to the dividend payment rate
'r: risk free rate. It is used to be equivalent to the continuous discount rate
'sigma: volatility of the underlying asset
'period: current period
'T: time period between current and maturity date
'N: number of nods between current period and maturity date
'K: strike price

'INITIALIZING PARAMETERS AND ARRAYS___
Dim i As Integer, j As Integer, steps As Integer


Dim deltaT As Double, u As Double, d As Double, Pu As Double, Pd As Double, DF As Double

deltaT = (T - period) / (N * 1#)
u = Exp(sigma * (deltaT) ^ (0.5))
d = 1 / u
Pu = (Exp((r - q) * deltaT) - d) / (u - d)
Pd = 1 - Pu
DF = Exp(-deltaT * r)



steps = N + 1 'Because we need the extra step in our array to include first node
Dim arr As Variant
ReDim arr(1 To steps)


For i = 1 To steps
     
     'Compute possible future stock price
     arr(i) = S * (u ^ (steps - i))
     arr(i) = arr(i) * (d ^ (i - 1))
     
     'Compute possible future option payoff
     'LCase: Returns a String that has been converted to lowercase.
     If LCase(flavor) = "call" Or LCase(flavor) = "c" Then
        arr(i) = WorksheetFunction.Max(arr(i) - K, 0#)
     Else
        arr(i) = WorksheetFunction.Max(K - arr(i), 0#)
    
     End If
     
Next i

steps = steps - 1 'We are done initializing. We now need N in its true form.


'This loop will now keep taking the expected value of up and down until the price is reached
'At each step backeards we ignores one more cell of the array

'for j = steps to 1 Step -1 would mean loop back from steps to 1 subtracting 1 from the j in each loop cycle.
For j = steps To 1 Step -1
    
    For i = 1 To j 'Put back to j
        
        arr(i) = (arr(i + 1) * Pd + arr(i) * Pu) * DF
        

    Next i
    
    
Next j

    
BAPMPrice = arr(1)
End Function

Function BAPM_help() As String

BAPM_help = "flavor As String, S As Double, q As Double, r As Double, sigma As Double, period As Double, T As Double, N As Double, K As Double"
End Function
