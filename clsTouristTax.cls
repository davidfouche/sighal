VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTouristTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' CLASS_MODULE:   TouristTax
'''
'''                 - Implements the methods for the calculation and export of tourist taxes
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 24/10/19        David Fouche            Created
'''

'Public properties

'Private attributes

Const LocalTaxRate2019 As Variant = 0.03
Const RegionTaxRate2019 As Variant = 0.1

'Defaults

'***************************************************************************
'Purpose: Class constructor of clsTouristTax
'Inputs:  None
'Outputs: An instance of the class
'***************************************************************************
Private Sub Class_Initialize()
    On Error GoTo ErrHandler:
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsAESMethod::Class_Initialize()" & vbCrLf & "Error Description: " & Err.Description)
End Sub

' ###########################################################################################
' #
' #                                     PUBLIC METHODS
' #
' ###########################################################################################


'***************************************************************************
'Purpose: Valid method for the calculation as of 01/01/2019
'Inputs:  Total amount of the booking
'Outputs: The stay tax of the record
'***************************************************************************
Public Function TouristTax2019(ByRef TotalAmount As Variant)
    TouristTax2019 = TotalAmount * LocalTaxRate2019 * (1 + RegionTaxRate2019)
End Function
