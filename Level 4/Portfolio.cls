VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Portfolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Student Name: Jianwei Su
'Date: 11/07/2022
'HW 4
'Question 2


Option Explicit
' ===================================================================================================
' File content:
' Class Portfolio
' ===================================================================================================

' Class member
Private pDerivatives As Collection

' Getter and setter
Public Property Get Derivatives() As Collection
    Set Derivatives = pDerivatives
End Property


Public Property Set Derivatives(lDerivatives As Collection)
    Set pDerivatives = lDerivatives
End Property

'This function is to count how many dates for the same instrument type
Function count_dates(strInstrumentType As String) As Variant
'===================================================================
' strInstrumentType -- string, name of the Instrument Type

' RETURNS:
' count_dates    -- integer, number of datesof the Instrument Type
' =================================================================
Dim i As Integer
Dim Der As Derivative
i = 0

For Each Der In Derivatives
    If strInstrumentType = Der.InstrumentType Then
        i = i + 1
    End If
Next Der

count_dates = i

End Function

'This function is to calculate the weight via one instrument and one date
Function GetWeight(strInstrument As String, COB As Date) As Double
' ==================================================================================================
' strInstrumentType             --string, the Instrument Type
' COB                           --double, the date

' RETURNS:
' GetWeight    --double, the weight of one asset at one date
' ==================================================================================================
Dim Der As Derivative
Dim totalWeight As Double
Dim marginalWeight As Double
   
totalWeight = 0#
For Each Der In Derivatives
    'Sums up total weight of this date
    If COB = Der.COB Then totalWeight = totalWeight + Der.Value
        
    'Sums up total weight of this date where is correct instrument
    If COB = Der.COB And strInstrument = Der.InstrumentType Then marginalWeight = Der.Value
    
Next Der
    
GetWeight = marginalWeight / totalWeight
    
End Function


Function GetData(strInstrumentType As String) As Variant
' ==================================================================================================
' strInstrumentType -- string, name of the Instrument Type

' RETURNS:
' GetData  -- variant, a table of historical data of one Instrument Type
                    ' e.g. format Date| Value| Returns
' ==================================================================================================

Dim Mat As Variant
Dim i As Integer
Dim Der As Derivative
Dim j As Integer
Dim weight As Double
Dim typ As String
    
'Here is to deal with when instrumentType is Portfolio
'I just use one of the instruments to count how many dates for the portfolio
If strInstrumentType = "Portfolio" Then
    strInstrumentType = "Real assets"
    typ = "Portfolio"
End If
    
j = count_dates(strInstrumentType)
    
'j is to count how many dates(rows) and there are 3 columns.
ReDim Mat(1 To j, 1 To 3)
    
i = 1
    
'Load the array with values and dates for that instrument type
For Each Der In Derivatives
    If strInstrumentType = Der.InstrumentType Then
        Mat(i, 1) = Der.COB
        Mat(i, 2) = Der.Value
        If typ = "Portfolio" Then
            weight = GetWeight(strInstrumentType, Der.COB)
            
            'This is the portfolio weight
            Mat(i, 2) = Der.Value / (1# * weight) - Der.Value
            
        End If
            
        i = i + 1
    End If
Next Der
    
    
'Returns
For i = 2 To j
    Mat(i, 3) = (Mat(i, 2) / Mat(i - 1, 2)) - 1
Next i
    
GetData = Mat
    
End Function

'VaR function
'pct is the percentage of the VaR, e.g. 0.01 for 99%
Function VaR(strInstrumentType As String, pct As Double) As Double
' ==================================================================================================
' strInstrumentType --string, the Instrument Type
' pct               --double, the percentage of the VaR, e.g. 0.01 for 99%

' RETURNS:
' VaR               --double, the VaR of one asset of percentage given
' ==================================================================================================

Dim returns As Variant
Dim i As Integer
Dim j As Integer
Dim Mat As Variant
Dim k As Double

Mat = GetData(strInstrumentType)

If strInstrumentType = "Portfolio" Then strInstrumentType = "Real assets"

j = count_dates(strInstrumentType)

'This is the array for daily returns
ReDim returns(1 To j - 1)

For i = 1 To j - 1
    returns(i) = Mat(i + 1, 3)
Next i

k = WorksheetFunction.Percentile(returns, pct)

VaR = k

End Function

'This function is to calculate the average returns for each instrument
Function aver_return(strInstrumentType As String) As Double
' ==================================================================================================
' ARGUMENTS:
' strInstrumentType --string, the Instrument Type

' RETURNS:
' aver_return               --double, the average return of one asset of dates given
' ==================================================================================================
Dim returns As Variant
Dim Mat As Variant
Dim i As Integer, j As Integer
Dim k As Double

Mat = GetData(strInstrumentType)

If strInstrumentType = "Portfolio" Then strInstrumentType = "Real assets"

j = count_dates(strInstrumentType)

'This is the array for daily returns
ReDim returns(1 To j - 1)

For i = 1 To j - 1
    returns(i) = Mat(i + 1, 3)
Next i

k = WorksheetFunction.Average(returns)

aver_return = k

End Function

