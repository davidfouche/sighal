Attribute VB_Name = "TestCrypto"
Option Explicit
Option Base 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' TEST_CODE:   TestCrypto
'''                 - Test crypto functions
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 31/05/2019      David FOUCHE            Created
'''

'***************************************************************************
'Purpose: Test the add of a new key in the key list
'Inputs:  None
'Outputs: a new key must be added into the key list and its hash and identifier updated accordingly
'***************************************************************************
Sub TestAddNewKey()
    KeyList.AddNewKey ("1234567890ABCDEF%")
End Sub

'***************************************************************************
'Purpose: Test the call to the method that hashes the passphrase
'Inputs:  None
'Outputs: the hash of the passphrase
'***************************************************************************
Sub TestHashPassphrase()
    Dim sPassphrase As String
    Dim sHashOut As String
    
    Call thiscryptoinit
    
    sPassphrase = "blabla"
    sHashOut = defaultCrypto.HashPassphrase(sPassphrase)
    
    Logging.logINFO ("defaultHash (" & sPassphrase & "): " & sHashOut)
End Sub

'***************************************************************************
'Purpose: Test the call to the methods that encrypt and decrypt with the AES method
'Inputs:  None
'Outputs: The decrypted sentence must be equal to the input sentence
'Source : https://gist.github.com/JonTheNiceGuy/44d8aa94da0a56841202
'***************************************************************************
Sub TestAES()

    Dim EncryptedText As String
    Dim DecryptedText As String
    Dim PlainText As String

    Dim oCryptoEngine As New clsCryptoEngine

    On Error GoTo ErrHandler:
    
    PlainText = "A sentence or data that needs to be encrypted."
    EncryptedText = oCryptoEngine.Encrypt(PlainText)
    
    MsgBox ("Encrypted text is: " + EncryptedText)
    
    DecryptedText = oCryptoEngine.RetrieveDecryptAES(EncryptedText)
    
    MsgBox ("Decrypted text is: " + DecryptedText)

    Exit Sub
ErrHandler:
    
End Sub


'***************************************************************************
'Purpose: Test the call to the method that masks by XOR any word or sentence
'Inputs:  None
'Outputs: The word or sentence must be masked by the GPRD passphrase
'***************************************************************************
Sub TestXorC()
    Dim gender, PlainLastName, PlainFirstName As String
    Dim EncLastName As String
    Dim GPRDPassphrase As String

    Call thiscryptoinit
    defaultCrypto.GPRDPassphrase = Settings.GPRDPassword.Value
    
    'Encrypt ANDRE in Guests.Cells(10, Range("LastName").Column)
    PlainLastName = "ANDRE"
    EncLastName = defaultCrypto.Encrypt(PlainLastName)
    Guests.Cells(10, Range("LastName").Column) = EncLastName
End Sub

'***************************************************************************
'Purpose: Test the call to the method that draws random sentences or strings
'Inputs:  None
'Outputs: The randoms must differ
'***************************************************************************
Sub TestRandom()
    Dim rdstr(6) As String
    Dim sRand As Variant
    
    Call thiscryptoinit
    rdstr(1) = defaultCrypto.DrawRandom(24, 24, 1)
    rdstr(2) = defaultCrypto.DrawRandom(10, 10, 1)
    rdstr(3) = defaultCrypto.DrawRandom(4, 4, 1)
    rdstr(4) = defaultCrypto.DrawRandom(24, 24, 1)
    rdstr(5) = defaultCrypto.DrawRandom(10, 10, 1)
    rdstr(6) = defaultCrypto.DrawRandom(4, 4, 1)
    
    For Each sRand In rdstr
        Logging.logINFO ("Test random : " & sRand)
    Next
    
End Sub

'***************************************************************************
'Purpose: Test the call to the method that check the status of the keys
'Inputs:  None
'Outputs: Only one active key must be found
'***************************************************************************
Sub TestKey()
    Dim keycoll As New Collection
    Dim guestcoll As New Collection
    Dim key As New clsKey
    Dim guest As New clsGuest
    Dim i, rowcount As Integer
    Dim activekey As New clsKey
    
    rowcount = KeyList.Cells(Cells.Rows.Count, 1).End(xlUp).row
    For i = 2 To rowcount
        key.status = KeyList.Cells(i, Range("KeyStatus").Column).Value
        If key.status <> OBSOLETEKEYSTATUS Then
            key.Identifier = KeyList.Cells(i, Range("Id").Column).Value
            key.hashvalue = KeyList.Cells(i, Range("HashValue").Column).Value
            key.hashmethod = KeyList.Cells(i, Range("HashMethod").Column).Value
            key.cryptomethod = KeyList.Cells(i, Range("CryptoAlgo").Column).Value
            key.timestamp = KeyList.Cells(i, Range("Timestamp").Column).Value
            keycoll.Add key
        End If
        If key.status = ACTIVEKEYSTATUS Then
            activekey.Identifier = KeyList.Cells(i, Range("Id").Column).Value
            activekey.hashvalue = KeyList.Cells(i, Range("HashValue").Column).Value
            activekey.hashmethod = KeyList.Cells(i, Range("HashMethod").Column).Value
            activekey.cryptomethod = KeyList.Cells(i, Range("CryptoAlgo").Column).Value
            activekey.timestamp = KeyList.Cells(i, Range("Timestamp").Column).Value
        End If
    Next i
    Call thiscryptoinit
    
    guest.Lastname = Guests.Cells(11, Range("LastName").Column)
    guest.Firstname = Guests.Cells(11, Range("LastName").Column)
    guestcoll.Add guest
End Sub

'***************************************************************************
'Purpose: Test the call to the methods that trans-crypt the data of the sheet Guests
'Inputs:  None
'Outputs: The data of the sheet Guests must be trans-ciphered
'***************************************************************************
Sub TestXCryptAll()
    'Test iteration through the booking sheet
    Dim mapper As clsMapper

    On Error GoTo ErrHandler:

    Set mapper = New clsMapper
    
    Call mapper.Map(Guests, XCRYPTOPE)
    
    Exit Sub
ErrHandler:
    Logging.logFATAL ("clsMapper::Map " & Err.Description)

End Sub

'***************************************************************************
'Purpose: Test the call to the methods that update the hash of the sheet Guests
'Inputs:  None
'Outputs: The data hashes of the sheet Guests must be updated
'***************************************************************************
Sub TestHashAll()
    'Test iteration through the booking sheet
    Dim mapper As clsMapper

    On Error GoTo ErrHandler:

    Set mapper = New clsMapper
    
    Call mapper.Map(Guests, UPDATEHASHOPE)
    
    Exit Sub
ErrHandler:
    Logging.logFATAL ("clsMapper::Map " & Err.Description)

End Sub
