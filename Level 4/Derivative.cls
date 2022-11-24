VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Derivative"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Student Name: Jianwei Su
'Date: 11/07/2022
'HW 4
'Question 1

Option Explicit
' ===================================================================================================
' File content:
' Class Derivative
' ===================================================================================================

' Class member
Private pValue As Double
Private pInstrumentType As String
Private pCOB As Date


' Getter and setter
Public Property Get Value() As Double
    Value = pValue
End Property

Public Property Let Value(lValue As Double)
    pValue = lValue
End Property

Public Property Get InstrumentType() As String
    InstrumentType = pInstrumentType
End Property

Public Property Let InstrumentType(lInstrumentType As String)
    pInstrumentType = lInstrumentType
End Property

Public Property Get COB() As Double
    COB = pCOB
End Property

Public Property Let COB(lCOB As Double)
    pCOB = lCOB
End Property







