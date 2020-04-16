Attribute VB_Name = "TestGPRD"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' TEST_CODE:   TestGPRD
'''                 - Test GPRD data processing
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 31/05/2019      David FOUCHE            Created
'''

'***************************************************************************
'Purpose: Test the call to the method that remove all the privacy data ruled by GPRD regulations
'Inputs:  None
'Outputs: All GPRD data must be removed or scrambeled
'***************************************************************************
Sub TestPurge()
    EventClearData ("Réservations")
End Sub
