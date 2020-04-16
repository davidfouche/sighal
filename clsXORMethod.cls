VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsXORMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' CLASS_MODULE:   XOR
'''
'''                 - use Macro 'Test' from 'TestCrypto' for testing and as an example
'''                 - this method shall only be called to mask a data with a random value before to encipher with a secure method (AES)
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 27/05/19        David Fouche            Created
'''

'Public properties


'Private attributes

'***************************************************************************
'Purpose: Class constructor of clsXORMethod
'Inputs:  None
'Outputs: An instance of the class
'***************************************************************************
Private Sub Class_Initialize()
    On Error GoTo ErrHandler:
    Exit Sub
ErrHandler:
End Sub


' ###########################################################################################
' #
' #                                     PUBLIC METHODS
' #
' ###########################################################################################

'***************************************************************************
'Purpose: Encrypt a string of data with a key
'Inputs:  The data as a string, the key as a string
'Outputs: The encrypted data as a string
'***************************************************************************
Public Function Encrypt(ByVal sdata As String, ByVal sKey As String) As String
    Encrypt = XorC(sdata, sKey)
End Function

'***************************************************************************
'Purpose: Decrypt a string of data with a key
'Inputs:  The encrypted data as a string, the key as a string
'Outputs: The decrypted data as a string
'***************************************************************************
Public Function Decrypt(ByVal sdata As String, ByVal sKey As String) As String
    Decrypt = XorC(sdata, sKey)
End Function


' ###########################################################################################
' #
' #                                     PRIVATE METHODS
' #
' ###########################################################################################

Private Function XorC(ByVal sdata As String, ByVal sKey As String) As String
    Dim L As Long, i As Long, byIn() As Byte, byOut() As Byte, byKey() As Byte
    Dim bEncOrDec As Boolean
     
    On Error GoTo ErrHandler:
     
     'confirm valid string and key input:
    If Len(sdata) = 0 Or Len(sKey) = 0 Then XorC = "Invalid argument(s) used": Exit Function
     'check whether running encryption or decryption (flagged by presence of "xxx" at start of sData):
    If Left$(sdata, 3) = "xxx" Then
        bEncOrDec = False 'decryption
        sdata = Mid$(sdata, 4)
    Else
        bEncOrDec = True 'encryption
    End If
     'assign strings to byte arrays (unicode)
    byIn = sdata
    byOut = sdata
    byKey = sKey
    L = LBound(byKey)
    For i = LBound(byIn) To UBound(byIn) - 1 Step 2
        byOut(i) = ((byIn(i) + Not bEncOrDec) Xor byKey(L)) - bEncOrDec 'avoid Chr$(0) by using bEncOrDec flag
        L = L + 2
        If L > UBound(byKey) Then L = LBound(byKey) 'ensure stay within bounds of Key
    Next i
    XorC = byOut
    If bEncOrDec Then XorC = "xxx" & XorC 'add "xxx" onto encrypted text
    
    Exit Function
ErrHandler:

End Function
