Attribute VB_Name = "Question_2"
'Name: Jianwei Su
'Date: 10/24/2022
'HW 2
'Question 2


Option Explicit


Function Interpolate(oldRates As Variant, freq As Double) As Variant
'oldRates : variant, a matrix of data input

Dim i As Integer, j As Integer, oldCount As Integer, Maturity As Integer

'We can also use Maturity = UBound(oldRates, 1) here
Maturity = oldRates(oldRates.Rows.Count, 1)


'The hash mark (#) is a Type Declaration Character (TDC) and forces the literal "1" to the type Double
'Maturity is how many periods, which means Years times the freq
Maturity = Maturity * (freq * 1#)

Dim Mat As Variant
ReDim Mat(1 To Maturity, 1 To 2)

' Fills in the years
For i = 1 To Maturity
    Mat(i, 1) = i / (freq * 1#)
    Mat(i, 2) = 0
Next i


'COPIES OVER KNOWN RATES
For i = 1 To Maturity
    For j = 1 To oldRates.Rows.Count
         If oldRates(j, 1) = i / (freq * 1#) Then
            Mat(i, 2) = oldRates(j, 2)
            
         End If
    Next j
Next i



'Rates interpolation
For i = 2 To Maturity
    If Mat(i, 2) = 0 Then
        
        'FIND THE NEXT FILLED RATE TO INTERPOLATE WITH
        For j = i + 1 To Maturity
        
         '<> is the "not equal" operator
         If Mat(j, 2) <> 0 Then Exit For
         
         Next j
         
        Mat(i, 2) = Mat(i - 1, 2) + (Mat(j, 2) - Mat(i - 1, 2)) / (j - i + 1)
        
        
    End If
    
Next i







Interpolate = Mat

End Function


