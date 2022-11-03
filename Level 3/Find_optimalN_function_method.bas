Attribute VB_Name = "Find_optimalN_function_method"
'Student name: Jianwei Su
'Date: 11/01/2022
'HW 3
'Find Optimal N





Option Explicit


Function FindN(flavor As String, S As Double, period As Double, T As Double, r As Double, sigma As Double, K As Double, q As Double) As Variant


'flavor: choose call or put option
'S: underlying asset stock spot price
'q: dividend yield. It is used to be equivalent to the dividend payment rate
'r: risk free rate. It is used to be equivalent to the continuous discount rate
'sigma: volatility of the underlying asset
'period: current period
'T: time period between current and maturity date
'K: strike price


Dim N As Integer, price As Double, BSPrice As Double

Dim arr As Variant
ReDim arr(1 To 2, 0)

N = 1#
BSPrice = BS_Price(flavor, S, q, r, sigma, period, T, K)

Do While Abs(BAPMPrice(flavor, S, q, r, sigma, period, T, N, K) - BSPrice) >= 0.01
        N = N + 1
        price = BAPMPrice(flavor, S, q, r, sigma, period, T, N, K)
Loop

arr(1, 0) = N
arr(2, 0) = price

FindN = arr

End Function


Function FindN_helper() As String

FindN_helper = "flavor As String, S As Double, period As Double, T As Double, r As Double, sigma As Double, K As Double, q As Double"

End Function




