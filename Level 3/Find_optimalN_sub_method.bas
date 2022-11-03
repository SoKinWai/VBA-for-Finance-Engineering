Attribute VB_Name = "Find_optimalN_sub_method"
'Student name: Jianwei Su
'Date: 11/01/2022
'HW 3
'Find Optimal N



Sub FindoptimalN()

Dim flavor As String, S As Double, period As Double, T As Double, r As Double, sigma As Double, K As Double, q As Double
Dim N As Integer, price As Double, BSPrice As Double


flavor = Sheets("Homework").Range("B8").Value
S = Sheets("Homework").Range("B9").Value
period = Sheets("Homework").Range("B10").Value
T = Sheets("Homework").Range("B11").Value
r = Sheets("Homework").Range("B12").Value
sigma = Sheets("Homework").Range("B13").Value
K = Sheets("Homework").Range("B14").Value
q = Sheets("Homework").Range("B15").Value

N = 1#
BSPrice = BS_Price(flavor, S, q, r, sigma, period, T, K)

Do While Abs(BAPMPrice(flavor, S, q, r, sigma, period, T, N, K) - BSPrice) >= 0.01
        N = N + 1
        price = BAPMPrice(flavor, S, q, r, sigma, period, T, N, K)
Loop



Sheets("Homework").Range("B22").Value = N
Sheets("Homework").Range("B23").Value = price



End Sub



