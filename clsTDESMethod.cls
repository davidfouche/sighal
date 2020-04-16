VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTDESMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' CLASS_MODULE:   TDES
'''
'''                 - use Macro 'Test' from 'TestCrypto' for testing and as an example
'''                 - the TDES method is deprecated and considered insecure
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 27/05/19        David Fouche            Created
'''

'Public properties


'Private attributes
Private TRIPLE_DES_KEY As String
Private INITIALIZATION_VECTOR As String

'***************************************************************************
'Purpose: Class constructor of clsTDESMethod
'Inputs:  None
'Outputs: An instance of the class
'***************************************************************************
Private Sub Class_Initialize()
    On Error GoTo ErrHandler:
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsTDESMethod::Class_Initialize()" & vbCrLf & "Error Description: " & Err.Description)
End Sub


' ###########################################################################################
' #
' #                                     PUBLIC METHODS
' #
' ###########################################################################################

'***************************************************************************
'Purpose: Encrypt a string of data with a key
'Inputs:  The data as a string, the key as a string, an initialization vector
'Outputs: The encrypted data as a string
'***************************************************************************
Public Function Encrypt(ByVal sdata As String, ByVal sKey As String) As String
    TRIPLE_DES_KEY = sKey
    INITIALIZATION_VECTOR = "012345678"
    Encrypt = EncryptStringTripleDES(sdata)
End Function

'***************************************************************************
'Purpose: Decrypt a string of encrypted data with a key
'Inputs:  The data as a string, the key as a string, an initialization vector
'Outputs: The decrypted data as a string
'***************************************************************************
Public Function Decrypt(ByVal sdata As String, ByVal sKey As String) As String
    TRIPLE_DES_KEY = sKey
    INITIALIZATION_VECTOR = "012345678"
    Decrypt = DecryptStringTripleDES(sdata)
End Function

' ###########################################################################################
' #
' #                                     PRIVATE METHODS
' #
' ###########################################################################################

'***************************************************************************
'Purpose: Encrypt a string of data with a key
'Inputs:  The data as a string, the key as a string
'Outputs: The encrypted data as a string
'***************************************************************************
Private Function EncryptStringTripleDES(plain_string As String) As Variant
    Dim encryption_object As Object
    Dim plain_byte_data() As Byte
    Dim encrypted_byte_data() As Byte
    Dim encrypted_base64_string As String

    EncryptStringTripleDES = Null
    On Error GoTo ErrHandler:

    plain_byte_data = CreateObject("System.Text.UTF8Encoding").Getbytes_4(plain_string)

    Set encryption_object = CreateObject("System.Security.Cryptography.TripleDESCryptoServiceProvider")
    encryption_object.Padding = 3
    encryption_object.key = CreateObject("System.Text.UTF8Encoding").Getbytes_4(TRIPLE_DES_KEY)
    encryption_object.IV = CreateObject("System.Text.UTF8Encoding").Getbytes_4(INITIALIZATION_VECTOR)
    encrypted_byte_data = _
            encryption_object.CreateEncryptor().TransformFinalBlock(plain_byte_data, 0, UBound(plain_byte_data) + 1)

    encrypted_base64_string = BytesToBase64(encrypted_byte_data)

    EncryptStringTripleDES = encrypted_base64_string

    Exit Function
ErrHandler:

    MsgBox "TripleDES encryption failed"

End Function

'***************************************************************************
'Purpose: Decrypt a string of encrypted data with a key
'Inputs:  The data as a string, the key as a string
'Outputs: The decrypted data as a string
'***************************************************************************
Private Function DecryptStringTripleDES(encrypted_string As String) As Variant

    Dim encryption_object As Object
    Dim encrypted_byte_data() As Byte
    Dim plain_byte_data() As Byte
    Dim plain_string As String

    DecryptStringTripleDES = Null

    On Error GoTo ErrHandler:

    encrypted_byte_data = Base64toBytes(encrypted_string)

    Set encryption_object = CreateObject("System.Security.Cryptography.TripleDESCryptoServiceProvider")
    encryption_object.Padding = 3
    encryption_object.key = CreateObject("System.Text.UTF8Encoding").Getbytes_4(TRIPLE_DES_KEY)
    encryption_object.IV = CreateObject("System.Text.UTF8Encoding").Getbytes_4(INITIALIZATION_VECTOR)
    plain_byte_data = encryption_object.CreateDecryptor().TransformFinalBlock(encrypted_byte_data, 0, UBound(encrypted_byte_data) + 1)

    plain_string = CreateObject("System.Text.UTF8Encoding").GetString(plain_byte_data)

    DecryptStringTripleDES = plain_string

    Exit Function
ErrHandler:

    MsgBox "TripleDES decryption failed"

End Function


'***************************************************************************
'Purpose: Convert a string in bytes into a base 64 string
'Inputs:  The data as an array of bytes
'Outputs: The data as a string
'***************************************************************************
Private Function BytesToBase64(varBytes() As Byte) As String
    With CreateObject("MSXML2.DomDocument").createElement("b64")
        .DataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64 = Replace(.Text, vbLf, "")
    End With
End Function


'***************************************************************************
'Purpose: Convert a string in base 64 string into an array of bytes
'Inputs:  The data as a base 64 string
'Outputs: The data as an array of bytes
'***************************************************************************
Private Function Base64toBytes(varStr As String) As Byte()
    With CreateObject("MSXML2.DOMDocument").createElement("b64")
         .DataType = "bin.base64"
         .Text = varStr
         Base64toBytes = .nodeTypedValue
    End With
End Function

