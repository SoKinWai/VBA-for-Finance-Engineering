Attribute VB_Name = "utils"
'Student name:Jianwei Su
'Date: 11/18/2022
'Final project


Option Explicit

'Get the market data
Sub GetDataSQL(path As String, strSQL As String)

Dim cn As Object
Dim rs As Object
Dim strFile As String
Dim strCon As String

strFile = path & "\MarketData.accdb"
strCon = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strFile

Set cn = CreateObject("ADODB.Connection")
cn.Open strCon

Set rs = CreateObject("ADODB.RECORDSET")
rs.activeconnection = cn

rs.Open strSQL
Sheets("Market_Data").Range("A2").CopyFromRecordset rs

rs.Close
cn.Close
Set cn = Nothing

End Sub

Function Interpolate(oldRates As Variant, freq As Double) As Variant
'oldRates : variant, a matrix of data input

Dim i As Integer, j As Integer, oldCount As Integer, maturity As Integer

'We can also use Maturity = UBound(oldRates, 1) here
maturity = oldRates(oldRates.Rows.Count, 1)

'The hash mark (#) is a Type Declaration Character (TDC) and forces the literal "1" to the type Double
'Maturity is how many periods, which means Years times the freq
maturity = maturity * (freq * 1#)

Dim Mat As Variant
ReDim Mat(1 To maturity, 1 To 2)

' Fills in the years
For i = 1 To maturity
    Mat(i, 1) = i / (freq * 1#)
    Mat(i, 2) = 0
Next i

'COPIES OVER KNOWN RATES
For i = 1 To maturity
    For j = 1 To oldRates.Rows.Count
         If oldRates(j, 1) = i / (freq * 1#) Then
            Mat(i, 2) = oldRates(j, 2)
         End If
    Next j
Next i

'Rates interpolation
For i = 2 To maturity
    If Mat(i, 2) = 0 Then
        'FIND THE NEXT FILLED RATE TO INTERPOLATE WITH
        For j = i + 1 To maturity
         '<> is the "not equal" operator
         If Mat(j, 2) <> 0 Then Exit For
         
        Next j
        Mat(i, 2) = Mat(i - 1, 2) + (Mat(j, 2) - Mat(i - 1, 2)) / (j - i + 1)
    End If
Next i

Interpolate = Mat

End Function

Function Interpolate_helper() As String

Interpolate_helper = "oldRates is variant, a matrix of data input, freq is an integer "

End Function

Function forwardRate(freq As Double, date1 As Double, date2 As Double, rate1 As Double, rate2 As Double) As Double

'freq: frequency of interpolation
'date1: the eailer date
'date2: the latter date
'rate1: forward rate of date1
'rate2: forward rate of date2

 Dim Compound_rate1 As Double, Compound_rate2 As Double

'Compound_rate2=Compound_rate1*(1+ForwardRate/freq)^(freq*(date2-date1))

Compound_rate1 = (1 + rate1 / freq) ^ (freq * date1)
Compound_rate2 = (1 + rate2 / freq) ^ (freq * date2)
forwardRate = (Compound_rate2 / Compound_rate1) ^ (1# / (freq * (date2 - date1))) - 1
forwardRate = forwardRate * freq


End Function


Function ForwardRate_helper() As String

ForwardRate_helper = "freq is Double, date1 is string, date2 is string, rate1 is double, rate2 is double"

End Function

