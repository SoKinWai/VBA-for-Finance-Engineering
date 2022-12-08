Attribute VB_Name = "Pricer"
'Name: Jianwei Su
'Date: 11/19/2022
'Final Project



Option Explicit

Function FRN(discountRate As Variant, freq As Double, notional As Double, forwardRate As Variant, coupon_spread As Double) As Double

'discountRate: array of money market rate
'freq: frequency of payment
'notional: notional amount
'forwardRate: array of implied forward rate
'coupon_spread: coupon rate=forward rate+ coupon spread, coupon spread= coupon rate-forward rate

Dim pv As Double, coupon As Double, i As Integer, discount_factor As Double, nper As Integer

nper = discountRate.Rows.Count

pv = 0#

For i = 1 To nper
    coupon = notional * (coupon_spread + forwardRate(i, 1)) / freq
    discount_factor = (1 + discountRate(i, 1) / freq) ^ (-i)
    pv = pv + coupon * discount_factor
Next i

FRN = pv + notional * discount_factor

End Function

Function FRN_helper() As String

FRN_helper = "discountRate As Variant, freq As Double, notional As Double, forwardRate As Variant, coupon_spread As Double"

End Function

Function BS_Price(flavor As String, S As Double, q As Double, r As Double, sigma As Double, period As Double, T As Double, k As Double) As Double

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

d1 = WorksheetFunction.Ln(S / k)
d1 = d1 + (r - q + 0.5 * sigma ^ 2) * (T - period)
d1 = d1 / (sigma * (T - period) ^ (1 / 2))
d2 = d1 - sigma * (T - period) ^ (1 / 2)

If LCase(flavor) = "call" Or LCase(flavor) = "c" Then
    price = S * Exp(-q * (T - period)) * WorksheetFunction.NormDist(d1, 0, 1, True) - k * Exp(-r * (T - period)) * WorksheetFunction.NormDist(d2, 0, 1, True)

Else
    price = k * Exp(-r * (T - period)) * WorksheetFunction.NormDist(-d2, 0, 1, True) - S * Exp(-q * (T - period)) * WorksheetFunction.NormDist(-d1, 0, 1, True)

End If

BS_Price = price

End Function

Function BSPrice_helper() As String

BSPrice_helper = "flavor As String, S As Double, q As Double, r As Double, sigma As Double, period As Double, T As Double, K As Double"

End Function


Function BondPrice(discountRate As Double, freq As Double, nper As Integer, notional As Double, coupon_rate As Double) As Double

'discountRate: discount rate used to discount CF to PV
'freq: frequency of payment
'nper: number of payments
'notional: notional amount
'coupon_rate: coupon rate
' RETURNS:
' BondPrice     --double, price of the bond
Dim pv As Double, coupon As Double, i As Integer, discount_factor As Double
pv = 0#

coupon = notional * coupon_rate / freq
For i = 1 To nper
    discount_factor = (1 + discountRate / freq) ^ (-i)
    pv = pv + coupon * discount_factor
Next i

BondPrice = pv + notional * discount_factor

End Function

Function BondPrice_helper() As String

BondPrice_helper = "discountRate As Double, freq As Double, nper As Integer, notional As Double, coupon_spread As Double"

End Function
