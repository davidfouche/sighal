VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCryptoEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' CLASS_MODULE:   CryptoEngine
'''
'''                 - use Macro 'Test' from 'TestCrypto' for testing and as an example
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 27/05/19        David Fouche            Created
'''

'Private properties
Private scurrentGPRDPassphrase As String
Private snewGPRDPassphrase As String
Private ocurrentkey As New clsKey
Private onewkey As New clsKey

'False if the password was never checked
Public isReady As Boolean

'Define key and method types
Public Enum CryptoType
  XORMethod = 0
  TDESMETHOD = 1
  AESCBCMETHOD = 2
  AESCBCMETHODLIGHT = 3
End Enum

'***************************************************************************
'Purpose: Class constructor of clsCryptoEngine
'Inputs:  None
'Outputs: An instance of the class
'***************************************************************************
Private Sub Class_Initialize()
    On Error GoTo ErrHandler:
    isReady = False
    'Read the active key from the key list
    Call Readactivekey
    Call Readnewkey
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsCryptoEngine::Class_Initialize()" & vbCrLf & "Error Description: " & Err.Description)
End Sub

' ###########################################################################################
' #
' #                                     PUBLIC METHODS
' #
' ###########################################################################################

'***************************************************************************
'Purpose: Calculates the hash of a passphrase
'Inputs:  The passphrase to be hashed
'Outputs: The hash of the passphrase
'***************************************************************************
Public Function HashPassphrase(ByVal sPassphrase As String)
    Dim oHashHandler As New clsHashMethod
    
    On Error GoTo ErrHandler:
    HashPassphrase = oHashHandler.Gethash(sPassphrase, ocurrentkey.hashmethod)
    Exit Function
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsCryptoEngine::HashPassphrase()" & vbCrLf & "Error Description: " & Err.Description)
End Function

'***************************************************************************
'Purpose: Calculates the SHA256 hash of a data
'Inputs:  The data as a string
'Outputs: The hash of the data
'***************************************************************************
Public Function HashData(ByVal sdata As String) As String
    Dim oHashHandler As New clsHashMethod
    
    On Error GoTo ErrHandler:
    HashData = oHashHandler.Gethash(sdata, SHA256Hash)
    Exit Function
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsCryptoEngine::HashData()" & vbCrLf & "Error Description: " & Err.Description)
End Function

'***************************************************************************
'Purpose: Encrypts a data accordingly with the key method
'Inputs:  The data as a string
'Outputs: The encrypted data as a string
'***************************************************************************
Public Function Encrypt(ByVal sin As String) As String
    Dim encryptmethod As Object
        
    On Error GoTo ErrHandler:
    'Check the passphrase on every call
    If CheckGPRDPassphrase(scurrentGPRDPassphrase) Then
        Select Case ocurrentkey.cryptomethod
            Case XORMethod:
                Set encryptmethod = New clsXORMethod
            Case TDESMETHOD:
                Set encryptmethod = New clsTDESMethod
            Case AESCBCMETHOD:
                Set encryptmethod = New clsAESMethod
                encryptmethod.strength = AESSTRONG
            Case AESCBCMETHODLIGHT:
                Set encryptmethod = New clsAESMethod
                encryptmethod.strength = AESLIGHT
            Case Else:
                Error.Raise 1, "Encrypt::Unknown Encrypt method", "Cryptomethod " & ocurrentkey.cryptomethod & " is unknown"
        End Select
        'You shall not encrypt an empty string
        If Len(Trim(sin)) = 0 Then
            Encrypt = ""
        Else
            Encrypt = encryptmethod.Encrypt(sin, scurrentGPRDPassphrase)
        End If
    End If
    Exit Function
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsCryptoEngine::Encrypt()" & vbCrLf & "Error Description: " & Err.Description)
End Function

'***************************************************************************
'Purpose: Decrypts a data accordingly with the key method
'Inputs:  The data as a string
'Outputs: The decrypted data as a string
'***************************************************************************
Public Function Decrypt(ByVal sin As String) As String
    Dim decryptmethod As Object
    
    On Error GoTo ErrHandler:
    'Check the passphrase on every call
    If CheckGPRDPassphrase(scurrentGPRDPassphrase) Then
        Select Case ocurrentkey.cryptomethod
            Case XORMethod:
                Set decryptmethod = New clsXORMethod
            Case TDESMETHOD:
                Set decryptmethod = New clsTDESMethod
            Case AESCBCMETHOD:
                Set decryptmethod = New clsAESMethod
                decryptmethod.strength = AESSTRONG
            Case AESCBCMETHODLIGHT:
                Set decryptmethod = New clsAESMethod
                decryptmethod.strength = AESLIGHT
            Case Else:
                Error.Raise 1, "Decrypt::Unknown decrypt method", "Cryptomethod " & ocurrentkey.cryptomethod & " is unknown"
        End Select
        'You shall not encrypt an empty string
        If Len(Trim(sin)) = 0 Then
            Decrypt = ""
        Else
            Decrypt = decryptmethod.Decrypt(sin, scurrentGPRDPassphrase)
        End If
    End If
    Exit Function
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsCryptoEngine::Decrypt()" & vbCrLf & "Error Description: " & Err.Description)
End Function

'***************************************************************************
'Purpose: Decrypts with the currently active key and encrypts accordingly with the new key method
'Inputs:  The data as a string
'Outputs: The encrypted data as a string
'***************************************************************************
Public Sub CryptWithNewKey(ByRef keyid As String, ByRef coll As Collection)
    'get the last new key id
    'for each item of the collection
    'get the keyid of every item
    'decrypt the item fields with the key[keyid]
    'encrypt the item fields with the new key
    'update the key identifier of the item
    Dim sdata, splain As Variant
    Dim sGPRDPassphrase As String
    Dim okey As New clsKey
    Dim localcoll As New Collection
    
    On Error GoTo ErrHandler:
    
    sGPRDPassphrase = scurrentGPRDPassphrase
    Set okey = ocurrentkey
    
    For Each sdata In coll
        scurrentGPRDPassphrase = sGPRDPassphrase
        Set ocurrentkey = okey
        splain = Me.Decrypt(sdata)
        scurrentGPRDPassphrase = snewGPRDPassphrase
        Set ocurrentkey = onewkey
        sdata = Me.Encrypt(splain)
        localcoll.Add sdata
    Next sdata
    
    Set coll = localcoll
    keyid = onewkey.Identifier
    
    ' Do not update the key and passphrase yet
    Set ocurrentkey = okey
    scurrentGPRDPassphrase = sGPRDPassphrase
    
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsCryptoEngine::CryptWithNewKey()" & vbCrLf & "Error Description: " & Err.Description)
End Sub


'***************************************************************************
'Purpose: Set the passphrase used for the GPRD access restriction to private data
'Inputs:  The passphrase as an input string
'Outputs: The passphrase as a private property or an exception if the entered passphrase is not correct
'***************************************************************************
Property Let GPRDPassphrase(passphrase As String)
    On Error GoTo ErrHandler:
    If CheckGPRDPassphrase(passphrase) Then
        scurrentGPRDPassphrase = passphrase
    End If
    Exit Property
ErrHandler:
    scurrentGPRDPassphrase = ""
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsCryptoEngine::GPRDPassphrase()" & vbCrLf & "Error Description: " & Err.Description)
    Error.Raise 1, "GPRDPassphrase::Invalid password", "Passphrase is invalid"
End Property

'***************************************************************************
'Purpose: Set the new passphrase used for the GPRD access restriction to private data
'Inputs:  The new passphrase as an input string
'Outputs: The new passphrase as a private property or an exception if the entered passphrase is not correct
'***************************************************************************
Property Let newGPRDPassphrase(passphrase As String)
    On Error GoTo ErrHandler:
    snewGPRDPassphrase = passphrase
    Exit Property
ErrHandler:
    snewGPRDPassphrase = ""
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsCryptoEngine::newGPRDPassphrase()" & vbCrLf & "Error Description: " & Err.Description)
    Error.Raise 1, "GPRDPassphrase::new password is not valid", "New passphrase is invalid"
End Property

'***************************************************************************
'Purpose: Get the current key identifier that will be used to crypt or hash data
'Inputs:  None
'Outputs: The current key identifier as a private property
'***************************************************************************
Property Get activekeyIndex() As String
    activekeyIndex = ocurrentkey.Identifier
End Property


' ###########################################################################################
' #
' #                                     PRIVATE METHODS
' #
' ###########################################################################################

'***************************************************************************
'Purpose: Returns False if the passphrase is not checked successfully
'Inputs:  The passphrase as an input string
'Outputs: True or False
'***************************************************************************
Private Function CheckGPRDPassphrase(passphrase As String)
    If Len(passphrase) = 0 Then
        CheckGPRDPassphrase = False
        Error.Raise 1, "SIGHAL::No passphrase entered", "Pass phrase is empty"
    Else
        If HashPassphrase(passphrase) = ocurrentkey.hashvalue Then
            CheckGPRDPassphrase = True
            isReady = True
        Else
            CheckGPRDPassphrase = False
            Error.Raise 1, "SIGHAL::Passphrase not valid", "Pass phrase entered is not correct"
        End If
    End If
End Function

'***************************************************************************
'Purpose: Read the new key from the key list
'Inputs:  None
'Outputs: The new key as a private attribute of the class
'***************************************************************************
Private Sub Readnewkey()
    Dim i, rowcount As Integer
    
    On Error GoTo ErrHandler:

    rowcount = KeyList.Cells(Cells.Rows.Count, 1).End(xlUp).row
    
    For i = 2 To rowcount
        If KeyList.Cells(i, Range("KeyStatus").Column).Value = NEWKEYSTATUS Then
            onewkey.Identifier = KeyList.Cells(i, Range("Id").Column).Value
            onewkey.hashvalue = KeyList.Cells(i, Range("HashValue").Column).Value
            onewkey.hashmethod = KeyList.Cells(i, Range("HashMethod").Column).Value
            onewkey.cryptomethod = KeyList.Cells(i, Range("CryptoAlgo").Column).Value
            onewkey.timestamp = KeyList.Cells(i, Range("Timestamp").Column).Value
        End If
    Next i

    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsCryptoEngine::Readnewkey()" & vbCrLf & "Error Description: " & Err.Description)
    Error.Raise 1, "Readnewkey::Unable to set active key", "No new key has been found"
End Sub


'***************************************************************************
'Purpose: Read the active key from the key list
'Inputs:  None
'Outputs: The active key as a private attribute of the class
'***************************************************************************
Private Sub Readactivekey()
    Dim i, rowcount As Integer
    
    On Error GoTo ErrHandler:

    rowcount = KeyList.Cells(Cells.Rows.Count, 1).End(xlUp).row
    
    For i = 2 To rowcount
        If KeyList.Cells(i, Range("KeyStatus").Column).Value = ACTIVEKEYSTATUS Then
            ocurrentkey.Identifier = KeyList.Cells(i, Range("Id").Column).Value
            ocurrentkey.hashvalue = KeyList.Cells(i, Range("HashValue").Column).Value
            ocurrentkey.hashmethod = KeyList.Cells(i, Range("HashMethod").Column).Value
            ocurrentkey.cryptomethod = KeyList.Cells(i, Range("CryptoAlgo").Column).Value
            ocurrentkey.timestamp = KeyList.Cells(i, Range("Timestamp").Column).Value
        End If
    Next i

    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsCryptoEngine::Readactivekey()" & vbCrLf & "Error Description: " & Err.Description)
    Error.Raise 1, "Readactivekey::Unable to set active key", "Active key can not be found"
End Sub

'***************************************************************************
'Purpose: Draw a random as a string
'Inputs:  Range for the length of the random (between Cnt1 and Cnt2), character set in use to make the random string
'Outputs: The random string
'***************************************************************************
Public Function DrawRandom(Cnt1 As Integer, Cnt2 As Integer, MySet As Integer)
    Dim Rand As String
    Dim i As Integer, RndNo As Integer, XSet As Integer
    Dim MyCase As Integer
    
    On Error GoTo ErrHandler:
    
    Application.Volatile
    
    Select Case MySet
    Case Is = "1" 'Upper case
        MyCase = 65: XSet = 26
    Case Is = "2" 'Lower Case
        MyCase = 97: XSet = 26
    Case Is = "3" 'Leading Capital
        MyCase = 97: XSet = 26
    Case Is = "4" 'Text digits
        MyCase = 48: XSet = 10
    Case Is = "5" 'Numeric digits
        MyCase = 48: XSet = 10
    End Select
     
    If MySet = 3 Then 'Set leading character of "Name"
        i = i + 1
        Randomize
        Rand = Rand & Chr(Int((26) * Rnd + 65))
    End If
     'Set random length of string
    RndNo = Int((Cnt2 + 1 - Cnt1) * Rnd + Cnt1)
    Do
        i = i + 1
        Randomize
        Rand = Rand & Chr(Int((XSet) * Rnd + MyCase))
    Loop Until i = RndNo
    
    DrawRandom = Rand
    
     'Convert string to number
    If MySet = 5 Then DrawRandom = DrawRandom * 1
    
    Exit Function
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsCryptoEngine::DrawRandom()" & vbCrLf & "Error Description: " & Err.Description)
End Function

