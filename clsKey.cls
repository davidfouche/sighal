VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' CLASS_MODULE:   Key
'''
'''                 - use Macro 'Test' from 'TestCrypto' for testing and as an example
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 27/05/19        David Fouche            Created
'''

Public Identifier As Integer
Public hashvalue As String
Public hashmethod As Integer
Public cryptomethod As Integer
Public status As Integer
Public timestamp As Date

'Define key status
Public Enum KeyStatus
  ACTIVEKEYSTATUS = 0
  NEWKEYSTATUS = 1
  OBSOLETEKEYSTATUS = 2
End Enum


'***************************************************************************
'Purpose: Class constructor of the key class
'Inputs:  None
'Outputs: An instance of the Key class
'***************************************************************************
Private Sub Class_Initialize()
    On Error GoTo ErrHandler:
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsKey::Class_Initialize()" & vbCrLf & "Error Description: " & Err.Description)
End Sub

' ###########################################################################################
' #
' #                                     PUBLIC METHODS
' #
' ###########################################################################################


' ###########################################################################################
' #
' #                                     PRIVATE METHODS
' #
' ###########################################################################################


