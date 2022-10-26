Attribute VB_Name = "Question_3"
'Name: Jianwei Su
'Date: 10/25/2022
'HW 2
'Question 3


Option Explicit

Function ForwardRate(freq As Double, oldRates As Variant) As Variant

'freq: frequency of interpolation
'oldRates : variant, a matrix of data input

Dim date1 As Double, date2 As Double, rate1 As Double, rate2 As Double
Dim Compound_rate1 As Double, Compound_rate2 As Double
Dim yield As Variant

'yield is the interpolate matrix which contains years' periods and yield for each period
yield = Interpolate(oldRates, freq)


Dim i As Integer, Maturity As Integer

'LBound: Returns a Long containing the smallest available subscript for the indicated dimension of an array.
'UBound: Returns a Long data type containing the largest available subscript for the indicated dimension of an array.
Maturity = UBound(yield, 1) - LBound(yield, 1) + 1

Dim Mat As Variant
ReDim Mat(1 To Maturity, 1 To 2)

Mat = yield
    

For i = 2 To Maturity
    date1 = Mat(i - 1, 1)
    date2 = Mat(i, 1)
    rate1 = yield(i - 1, 2)
    rate2 = yield(i, 2)
    
    
    'Compound_rate2=Compound_rate1*(1+ForwardRate/freq)^(freq*(date2-date1))

    Compound_rate1 = (1 + rate1 / freq) ^ (freq * date1)
    Compound_rate2 = (1 + rate2 / freq) ^ (freq * date2)


    Mat(i, 2) = (Compound_rate2 / Compound_rate1) ^ (1# / (freq * (date2 - date1))) - 1
    Mat(i, 2) = Mat(i, 2) * freq
Next i


ForwardRate = Mat

End Function

