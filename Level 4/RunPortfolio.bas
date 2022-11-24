Attribute VB_Name = "RunPortfolio"
'Student Name:Jianwei Su
'Date: 11/11/2022
'Run the portfolio

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
For i = 5 To 64
    Set Der = New Derivative
    Der.InstrumentType = Sheets("MarketData").Cells(i, 1)
    Der.COB = Sheets("MarketData").Cells(i, 2)
    Der.Value = Sheets("MarketData").Cells(i, 3)

    cDers.Add Der
Next i

Set PopulateArray = cDers
End Function

'This function is to remove one of the instruments.
Function PopulateArray_b(remove As String) As Collection
' ==================================================================================================
' remove            --string, name of the asset to be removed

' RETURNS:
' PopulateArray_b     --Collection, collection of Derivatives
' ==================================================================================================
Dim cDers As Collection
Dim Der As Derivative
Set cDers = New Collection

Dim a As Integer
    
If remove <> "Commodity" Then
    For a = 5 To 14
     Set Der = New Derivative
     Der.InstrumentType = Sheets("MarketData").Cells(a, 1)
     Der.COB = Sheets("MarketData").Cells(a, 2)
     Der.Value = Sheets("MarketData").Cells(a, 3)
     cDers.Add Der
    Next a
End If
            
If remove <> "Equity" Then
    For a = 15 To 24
     Set Der = New Derivative
     Der.InstrumentType = Sheets("MarketData").Cells(a, 1)
     Der.COB = Sheets("MarketData").Cells(a, 2)
     Der.Value = Sheets("MarketData").Cells(a, 3)
     cDers.Add Der
    Next a
End If
    
If remove <> "Fixed Income" Then
    For a = 25 To 34
     Set Der = New Derivative
     Der.InstrumentType = Sheets("MarketData").Cells(a, 1)
     Der.COB = Sheets("MarketData").Cells(a, 2)
     Der.Value = Sheets("MarketData").Cells(a, 3)
     cDers.Add Der
    Next a
End If
    
If remove <> "Futures" Then
    For a = 35 To 44
     Set Der = New Derivative
     Der.InstrumentType = Sheets("MarketData").Cells(a, 1)
     Der.COB = Sheets("MarketData").Cells(a, 2)
     Der.Value = Sheets("MarketData").Cells(a, 3)
     cDers.Add Der
    Next a
End If
    
If remove <> "CDS" Then
    For a = 45 To 54
     Set Der = New Derivative
     Der.InstrumentType = Sheets("MarketData").Cells(a, 1)
     Der.COB = Sheets("MarketData").Cells(a, 2)
     Der.Value = Sheets("MarketData").Cells(a, 3)
     cDers.Add Der
    Next a
End If
    
If remove <> "Real assets" Then
    For a = 55 To 64
     Set Der = New Derivative
     Der.InstrumentType = Sheets("MarketData").Cells(a, 1)
     Der.COB = Sheets("MarketData").Cells(a, 2)
     Der.Value = Sheets("MarketData").Cells(a, 3)
     cDers.Add Der
    Next a
End If
    
Set PopulateArray_b = cDers

End Function


Sub RunPortfolio()
' ==================================================================================================
' OPERATIONS:
' Remove Instrument with best return or VaR
' ==================================================================================================
Dim newPort As Portfolio
Set newPort = New Portfolio
Set newPort.Derivatives = PopulateArray()
    
Sheets("Homework").Range("C14").Value = newPort.VaR("Commodity", 0.1)
Sheets("Homework").Range("D14").Value = newPort.VaR("Commodity", 0.05)
Sheets("Homework").Range("E14").Value = newPort.VaR("Commodity", 0.01)

Sheets("Homework").Range("C15").Value = newPort.VaR("Equity", 0.1)
Sheets("Homework").Range("D15").Value = newPort.VaR("Equity", 0.05)
Sheets("Homework").Range("E15").Value = newPort.VaR("Equity", 0.01)

Sheets("Homework").Range("C16").Value = newPort.VaR("Fixed Income", 0.1)
Sheets("Homework").Range("D16").Value = newPort.VaR("Fixed Income", 0.05)
Sheets("Homework").Range("E16").Value = newPort.VaR("Fixed Income", 0.01)

Sheets("Homework").Range("C17").Value = newPort.VaR("Futures", 0.1)
Sheets("Homework").Range("D17").Value = newPort.VaR("Futures", 0.05)
Sheets("Homework").Range("E17").Value = newPort.VaR("Futures", 0.01)

Sheets("Homework").Range("C18").Value = newPort.VaR("CDS", 0.1)
Sheets("Homework").Range("D18").Value = newPort.VaR("CDS", 0.05)
Sheets("Homework").Range("E18").Value = newPort.VaR("CDS", 0.01)

Sheets("Homework").Range("C19").Value = newPort.VaR("Portfolio", 0.1)
Sheets("Homework").Range("D19").Value = newPort.VaR("Portfolio", 0.05)
Sheets("Homework").Range("E19").Value = newPort.VaR("Portfolio", 0.01)

Dim newPort2 As Portfolio

Set newPort2 = New Portfolio
Set newPort2.Derivatives = PopulateArray_b("Commodity")
Sheets("Homework").Range("K13").Value = newPort2.aver_return("Portfolio")
Sheets("Homework").Range("L13").Value = newPort2.VaR("Portfolio", 0.1)
Sheets("Homework").Range("M13").Value = newPort2.VaR("Portfolio", 0.05)
Sheets("Homework").Range("N13").Value = newPort2.VaR("Portfolio", 0.01)

Set newPort2 = New Portfolio
Set newPort2.Derivatives = PopulateArray_b("Equity")
Sheets("Homework").Range("K14").Value = newPort2.aver_return("Portfolio")
Sheets("Homework").Range("L14").Value = newPort2.VaR("Portfolio", 0.1)
Sheets("Homework").Range("M14").Value = newPort2.VaR("Portfolio", 0.05)
Sheets("Homework").Range("N14").Value = newPort2.VaR("Portfolio", 0.01)

Set newPort2 = New Portfolio
Set newPort2.Derivatives = PopulateArray_b("Fixed Income")
Sheets("Homework").Range("K15").Value = newPort2.aver_return("Portfolio")
Sheets("Homework").Range("L15").Value = newPort2.VaR("Portfolio", 0.1)
Sheets("Homework").Range("M15").Value = newPort2.VaR("Portfolio", 0.05)
Sheets("Homework").Range("N15").Value = newPort2.VaR("Portfolio", 0.01)

Set newPort2 = New Portfolio
Set newPort2.Derivatives = PopulateArray_b("Futures")
Sheets("Homework").Range("K16").Value = newPort2.aver_return("Portfolio")
Sheets("Homework").Range("L16").Value = newPort2.VaR("Portfolio", 0.1)
Sheets("Homework").Range("M16").Value = newPort2.VaR("Portfolio", 0.05)
Sheets("Homework").Range("N16").Value = newPort2.VaR("Portfolio", 0.01)

Set newPort2 = New Portfolio
Set newPort2.Derivatives = PopulateArray_b("CDS")
Sheets("Homework").Range("K17").Value = newPort2.aver_return("Portfolio")
Sheets("Homework").Range("L17").Value = newPort2.VaR("Portfolio", 0.1)
Sheets("Homework").Range("M17").Value = newPort2.VaR("Portfolio", 0.05)
Sheets("Homework").Range("N17").Value = newPort2.VaR("Portfolio", 0.01)

End Sub


Sub ClearContent()

Sheets("Homework").Range("C14:E19").ClearContents
Sheets("Homework").Range("K13:N17").ClearContents

End Sub
