Attribute VB_Name = "BSPrice"
'Student name: Jianwei Su
'Date: 10/31/2022
'HW 3
'Black Scholes option pricer function

Option Explicit


Function BS_Price(flavor As String, S As Double, q As Double, r As Double, sigma As Double, period As Double, T As Double, K As Double) As Double

'flavor: choose call or put option
'S: underlying asset stock spot price
'q: dividend yield. It is used to be equivalent to the dividend payment rate
'r: risk free rate. It is used to be equivalent to the continuous discount rate
'sigma: volatility of the underlying asset
'period: current period
'T: time period between current and maturity date
'K: strike price

Dim d1 As Double, d2 As Double
Dim price As Double
price = 0#


d1 = WorksheetFunction.Ln(S / K)
d1 = d1 + (r - q + 0.5 * sigma ^ 2) * (T - period)
d1 = d1 / (sigma * (T - period) ^ (1 / 2))
d2 = d1 - sigma * (T - period) ^ (1 / 2)

If LCase(flavor) = "call" Or LCase(flavor) = "c" Then
    price = S * Exp(-q * (T - period)) * WorksheetFunction.NormDist(d1, 0, 1, True) - K * Exp(-r * (T - period)) * WorksheetFunction.NormDist(d2, 0, 1, True)

Else
    price = K * Exp(-r * (T - period)) * WorksheetFunction.NormDist(-d2, 0, 1, True) - S * Exp(-q * (T - period)) * WorksheetFunction.NormDist(-d1, 0, 1, True)

End If





BS_Price = price

End Function

Function BSPrice_helper() As String

BSPrice_helper = "flavor As String, S As Double, q As Double, r As Double, sigma As Double, period As Double, T As Double, K As Double"

End Function

