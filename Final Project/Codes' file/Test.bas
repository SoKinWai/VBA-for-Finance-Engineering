Attribute VB_Name = "Test"
'Name: Jianwei Su
'Date: 11/23/2022
'Final Project



Option Explicit

Function PopulateArray() As Collection
' ==================================================================================================
' RETURNS:
' PopulateArray     --Collection, collection of Derivatives
' ==================================================================================================
Dim cDers As Collection
Dim Der As Derivative
Set cDers = New Collection

Dim i As Integer
Dim row_count As Integer

'Count how many rows in column B
row_count = Sheets("Output").Cells(Rows.Count, "b").End(xlUp).Row

For i = 2 To row_count
    Set Der = New Derivative
    Der.InstrumentType = Sheets("Output").Cells(i, 3)
    Der.COB = Sheets("Output").Cells(i, 1)
    Der.Value = Sheets("Output").Cells(i, 4)
    cDers.Add Der
Next i

' Checking Phrase
Set PopulateArray = cDers
End Function

Sub Get_output()
' ==================================================================================================
' OPERATIONS:
' Create a chart with specific format
' ==================================================================================================


Dim i As Integer
Dim path As String, strSQL As String
Dim row_count As Integer
Dim freq As Double, discount_rate As Double, coupon_rate As Double, notional As Double, nper As Integer, coupon_margin As Double
Dim ticker As String
 
For i = 2 To 502
   If Sheets("Output").Range("C" & i).Value = "Stock" Then
        path = Sheets("Input").Range("Path").Value
        strSQL = "SELECT DATE,TICKER,CLOSE_PRICE From Market_Data "
        utils.GetDataSQL path, strSQL
        
        ticker = Sheets("Input").Range("C2")
        Sheets("Output").Range("D" & i).Value = WorksheetFunction.Index(Range(ticker & "_Price"), WorksheetFunction.Match(Sheets("Output").Range("A" & i), Range(ticker & "_Date"), 0))
    End If
    
    If Sheets("Output").Range("C" & i).Value = "Bond" Then
        discount_rate = WorksheetFunction.Index(Range("Discount_Rate"), WorksheetFunction.Match(Sheets("Output").Range("A" & i), Range("Riskfree_Date"), 0))
        coupon_rate = WorksheetFunction.Index(Range("Coupon"), WorksheetFunction.Match(Sheets("Output").Range("B" & i), Range("Position_ID"), 0))
        notional = WorksheetFunction.Index(Range("Notional"), WorksheetFunction.Match(Sheets("Output").Range("B" & i), Range("Position_ID"), 0))
        freq = 2#
        nper = 0.5 * freq
        Sheets("Output").Range("D" & i).Value = Pricer.BondPrice(discount_rate, freq, nper, notional, coupon_rate)
    End If
    
    If Sheets("Output").Range("C" & i).Value = "Cash" Then
        Sheets("Output").Range("D" & i).Value = WorksheetFunction.Index(Range("Quantity"), WorksheetFunction.Match(Sheets("Output").Range("B" & i), Range("Position_ID"), 0))
    End If
    
    If Sheets("Output").Range("C" & i).Value = "Floater(1)" Then
        coupon_margin = WorksheetFunction.Index(Range("Coupon_Margin"), WorksheetFunction.Match(Sheets("Output").Range("B" & i), Range("Position_ID"), 0))
        notional = WorksheetFunction.Index(Range("Notional"), WorksheetFunction.Match(Sheets("Output").Range("B" & i), Range("Position_ID"), 0))
        freq = 360
        Sheets("Output").Range("D" & i).Value = Pricer.FRN(WorksheetFunction.Index(Range("discount"), WorksheetFunction.Match(Sheets("Output").Range("A" & i), Range("FRN_Date"), 0)), freq, notional, WorksheetFunction.Index(Range("Forward_Rate"), WorksheetFunction.Match(Sheets("Output").Range("A" & i), Range("FRN_Date"), 0)), coupon_margin)
    End If
    
    If Sheets("Output").Range("C" & i).Value = "Floater(2)" Then
        coupon_margin = WorksheetFunction.Index(Range("Coupon_Margin"), WorksheetFunction.Match(Sheets("Output").Range("B" & i), Range("Position_ID"), 0))
        notional = WorksheetFunction.Index(Range("Notional"), WorksheetFunction.Match(Sheets("Output").Range("B" & i), Range("Position_ID"), 0))
        freq = 360
        Sheets("Output").Range("D" & i).Value = Pricer.FRN(WorksheetFunction.Index(Range("discount"), WorksheetFunction.Match(Sheets("Output").Range("A" & i), Range("FRN_Date"), 0)), freq, notional, WorksheetFunction.Index(Range("Forward_Rate"), WorksheetFunction.Match(Sheets("Output").Range("A" & i), Range("FRN_Date"), 0)), coupon_margin)
    End If



Next i

End Sub

Sub Get_report()
' ==================================================================================================
' OPERATIONS:
' compute portfolio and marginal VaR
' ==================================================================================================
Dim newPort As Portfolio
Set newPort = New Portfolio
Set newPort.Derivatives = PopulateArray()

Dim i As Integer

For i = 9 To 14:
    'Average returns of Stock, Bond, Floaters, Cash and Portfolio
    Sheets("Report").Range("D" & i).Value = newPort.aver_return(Sheets("Report").Range("B" & i).Value)
    
    '90% VaR of Stock, Bond, Floaters, Cash and Portfolio
    Sheets("Report").Range("E" & i).Value = newPort.VaR(Sheets("Report").Range("B" & i).Value, 0.1)
    
    '95% VaR of Stock, Bond, Floaters, Cash and Portfolio
    Sheets("Report").Range("F" & i).Value = newPort.VaR(Sheets("Report").Range("B" & i).Value, 0.05)

    '99% VaR of Stock, Bond, Floaters, Cash and Portfolio
    Sheets("Report").Range("G" & i).Value = newPort.VaR(Sheets("Report").Range("B" & i).Value, 0.01)
Next i

End Sub

Sub Run()
' ==================================================================================================
' OPERATIONS:
' Get data outputs and the final report result
' ==================================================================================================

Get_output
Get_report

End Sub

Sub ClearContent()
' ==================================================================================================
' OPERATIONS:
' Clear out all data inputs
' ==================================================================================================
Sheets("Market_Data").Range("A2:C1001").ClearContents
Sheets("Output").Range("D2:D501").ClearContents
Sheets("Report").Range("D9:G14").ClearContents

End Sub
