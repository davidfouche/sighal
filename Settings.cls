VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' SHEET_CODE:   Settings
'''                 - Handle sheet events
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 31/05/2019      David FOUCHE            Created
'''

Private Sub GPRDPassword_Change()
    On Error GoTo ErrHandler:
    GPRD.defaultCrypto.GPRDPassphrase = GPRDPassword.Value
    GPRDPassword.BackColor = vbGreen
    Logging.logINFO ("Settings::GPRDPassword_Change GRPD password has been verified successfully")
    Exit Sub
ErrHandler:
    GPRDPassword.BackColor = vbRed
    Logging.logINFO ("Settings::GPRDPassword_Change GRPD password has not been verified successfully")
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsBookingOrder::Class_Initialize()" & vbCrLf & "Error Description: " & Err.Description)
End Sub

Private Sub ChangePassphraseButton_Click()
    NewPassphraseForm.Show
    Logging.logINFO ("Settings::ChangePassphraseButton_Click User clicked on the button to change the password")
End Sub

