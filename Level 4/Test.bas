Attribute VB_Name = "Test"
'Student Name:Jianwei Su
'Date: 11/10/2022
'Test the array

Option Explicit


Function PopulateArray() As Collection
' ==================================================================================================
' RETURNS:
' PopulateArray --Collection, collection of Derivatives
' ==================================================================================================
    
Dim cDers As Collection
Dim Der As Derivative
Set cDers = New Collection
    
Dim i As Integer
    
'Loads in all the instruments and values from spreadsheet
For i = 5 To 54
    Set Der = New Derivative
    Der.InstrumentType = Sheets("MarketData").Cells(i, 1)
    Der.COB = Sheets("MarketData").Cells(i, 2)
    Der.Value = Sheets("MarketData").Cells(i, 3)

    cDers.Add Der
Next i
    
    Set PopulateArray = cDers
End Function

Sub TestGetPortfolioValue()
' ==================================================================================================
' This Sub is to test the weight and VaR for different instruments
' ==================================================================================================

Dim newPort As Portfolio
Set newPort = New Portfolio
Set newPort.Derivatives = PopulateArray()

Dim i As Integer, j As Integer
Dim Mat As Variant
Mat = newPort.GetData(Range("INSTRUMENT"))
    
'Load output range with matrix from GetData() function
For i = 1 To 10
    For j = 1 To 3
        Range("OUTPUT").Cells(i, j) = Mat(i, j)
    Next j
Next i
    
'Calculate and return instrument weight
Range("WEIGHT") = newPort.GetWeight(Range("INSTRUMENT"), Range("COB"))
    
Range("Value_at_risk") = newPort.VaR(Range("INSTRUMENT"), Range("pct"))
    
End Sub
         
    
