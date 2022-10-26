Attribute VB_Name = "Extra_Credits"
'Name: Jianwei Su
'Date: 10/25/2022
'HW 2
'Extra Credits



Option Explicit

Function FRN(oldRates As Variant, freq As Double, notional As Double, nper As Integer, coupon_spread As Double) As Double

'oldRates: variant, a matrix of data input
'freq: frequency of payment
'notional: notional amount
'nper: number of payments
'coupon_spread: coupon rate=forward rate+ coupon spread, coupon spread= coupon rate-forward rate

Dim discount_rates As Variant, forward_rates As Variant
discount_rates = Interpolate(oldRates, freq)
forward_rates = ForwardRate(freq, oldRates)

Dim pv As Double, coupon As Double, i As Integer, discount_factor As Double


pv = 0#

For i = 1 To nper
    coupon = notional * (coupon_spread + forward_rates(i, 2)) / freq
    discount_factor = (1 + discount_rates(i, 2) / freq) ^ (-i)
    pv = pv + coupon * discount_factor
    
Next i



FRN = pv + 100 * discount_factor

End Function
