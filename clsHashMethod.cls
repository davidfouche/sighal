VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHashMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' CLASS_MODULE:   HashMethod
'''
'''                 - use Macro 'Test' from 'TestCrypto' for testing and as an example
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 27/05/19        David Fouche            Created
'''

'Define key and method types
Public Enum HashType
  MD5Hash = 0
  SHA1HASH = 1
  SHA256Hash = 2
  SHA384Hash = 3
  SHA512Hash = 4
  SHA512SaltHash = 5
End Enum


'Class constructor
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

Public Function Gethash(ByVal sin As String, Optional sMethodeType As Integer, Optional ByVal sKey As String)
    Select Case sMethodeType
        Case MD5Hash:
            Gethash = MD5(sin)
        Case SHA1HASH:
            Gethash = SHA1(sin)
        Case SHA256Hash:
            Gethash = SHA256(sin)
        Case SHA384Hash:
            Gethash = SHA384(sin)
        Case SHA512Hash:
            Gethash = SHA512(sin)
        Case SHA512SaltHash:
            Gethash = StrToSHA512Salt(sin, sKey)
        Case Else:
            'By default SHA512
            Gethash = SHA512(sin)
    End Select
End Function


' ###########################################################################################
' #
' #                                     PRIVATE METHODS
' #
' ###########################################################################################

' MD5 Hash
Private Function MD5(ByVal sin As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    'Hex:   d41d8cd98f00...etc
    'Base-64: 1B2M2Y8Asg...etc
        
    Dim oT As Object, oMD5 As Object
    Dim TextToHash() As Byte
    Dim bytes() As Byte
    
    On Error GoTo ErrHandler:

    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
 
    TextToHash = oT.Getbytes_4(sin)
    bytes = oMD5.ComputeHash_2((TextToHash))
 
    If bB64 = True Then
       MD5 = ConvToBase64String(bytes)
    Else
       MD5 = ConvToHexString(bytes)
    End If
        
    Set oT = Nothing
    Set oMD5 = Nothing

    Exit Function
ErrHandler:

End Function

' SHA1 Hash
Private Function SHA1(sin As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    '40 Hex:   da39a3ee5e6...etc
    '28 Base-64:   2jmj7l5rSw0yVb...etc
    
    Dim oT As Object, oSHA1 As Object
    Dim TextToHash() As Byte
    Dim bytes() As Byte

    On Error GoTo ErrHandler:
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA1 = CreateObject("System.Security.Cryptography.SHA1Managed")
    
    If Len(sin) = 0 Then
        Exit Function
    End If
    
    TextToHash = oT.Getbytes_4(sin)
    bytes = oSHA1.ComputeHash_2((TextToHash))
        
    If bB64 = True Then
       SHA1 = ConvToBase64String(bytes)
    Else
       SHA1 = ConvToHexString(bytes)
    End If
            
    Set oT = Nothing
    Set oSHA1 = Nothing

    Exit Function
ErrHandler:

End Function

' SHA256 Hash
Private Function SHA256(sin As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    '64 Hex:   e3b0c44298f...etc
    '44 Base-64:   47DEQpj8HBSa+/...etc
    
    Dim oT As Object, oSHA256 As Object
    Dim TextToHash() As Byte, bytes() As Byte
    
    On Error GoTo ErrHandler:
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    TextToHash = oT.Getbytes_4(sin)
    bytes = oSHA256.ComputeHash_2((TextToHash))
    
    If bB64 = True Then
       SHA256 = ConvToBase64String(bytes)
    Else
       SHA256 = ConvToHexString(bytes)
    End If
    
    Set oT = Nothing
    Set oSHA256 = Nothing

    Exit Function
ErrHandler:

End Function

' SHA384 Hash
Private Function SHA384(sin As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    '96 Hex:   38b060a751ac...etc
    '64 Base-64:   OLBgp1GsljhM2T...etc
    
    Dim oT As Object, oSHA384 As Object
    Dim TextToHash() As Byte, bytes() As Byte
    
    On Error GoTo ErrHandler:

    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA384 = CreateObject("System.Security.Cryptography.SHA384Managed")
    
    TextToHash = oT.Getbytes_4(sin)
    bytes = oSHA384.ComputeHash_2((TextToHash))
    
    If bB64 = True Then
       SHA384 = ConvToBase64String(bytes)
    Else
       SHA384 = ConvToHexString(bytes)
    End If
    
    Set oT = Nothing
    Set oSHA384 = Nothing

    Exit Function
ErrHandler:

End Function

' SHA512 Hash
Private Function SHA512(sin As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    '128 Hex:   cf83e1357eefb8bd...etc
    '88 Base-64:   z4PhNX7vuL3xVChQ...etc
    
    Dim oT As Object, oSHA512 As Object
    Dim TextToHash() As Byte, bytes() As Byte

    On Error GoTo ErrHandler:

    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")
    
    TextToHash = oT.Getbytes_4(sin)
    bytes = oSHA512.ComputeHash_2((TextToHash))
    
    If bB64 = True Then
       SHA512 = ConvToBase64String(bytes)
    Else
       SHA512 = ConvToHexString(bytes)
    End If
    
    Set oT = Nothing
    Set oSHA512 = Nothing

    Exit Function
ErrHandler:

End Function

' SHA512 with a secret key Hash
Private Function StrToSHA512Salt(ByVal sin As String, ByVal sSecretKey As String, _
                           Optional ByVal b64 As Boolean = False) As String
    'Returns a sha512 STRING HASH in function name, modified by the parameter sSecretKey.
    'This hash differs from that of SHA512 using the SHA512Managed class.
    'HMAC class inputs are hashed twice;first input and key are mixed before hashing,
    'then the key is mixed with the result and hashed again.
    
    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SecretKey() As Byte
    Dim bytes() As Byte

    On Error GoTo ErrHandler:

    'Test results with both strings empty:
    '128 Hex:    b936cee86c9f...etc
    '88 Base-64:   uTbO6Gyfh6pd...etc
    
    'create text and crypto objects
    Set asc = CreateObject("System.Text.UTF8Encoding")
    
    'Any of HMACSHAMD5,HMACSHA1,HMACSHA256,HMACSHA384,or HMACSHA512 can be used
    'for corresponding hashes, albeit not matching those of managed classes.
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA512")

    'make a byte array of the text to hash
    bytes = asc.Getbytes_4(sin)
    'make a byte array of the private key
    SecretKey = asc.Getbytes_4(sSecretKey)
    'add the private key property to the encryption object
    enc.key = SecretKey

    'make a byte array of the hash
    bytes = enc.ComputeHash_2((bytes))
    
    'convert the byte array to string
    If b64 = True Then
       StrToSHA512Salt = ConvToBase64String(bytes)
    Else
       StrToSHA512Salt = ConvToHexString(bytes)
    End If
    
    'release object variables
    Set asc = Nothing
    Set enc = Nothing

    Exit Function
ErrHandler:

End Function

' Conversion functions used in hash methods
Private Function ConvToBase64String(vIn As Variant) As Variant

    Dim oD As Object

    On Error GoTo ErrHandler:

    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

    Exit Function
ErrHandler:

End Function

Private Function ConvToHexString(vIn As Variant) As Variant

    Dim oD As Object

    On Error GoTo ErrHandler:
    
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

    Exit Function
ErrHandler:

End Function

'***************************************************************************
'Purpose: Convert a string in base 64 but without using the plugin MSXML2.DOMDocument
'Inputs:  The data as a base 64 string
'Outputs: The hash of the data as a string
'***************************************************************************
Private Function BASE64SHA1(ByVal sTextToHash As String)

    Dim asc As Object
    Dim enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Dim bytes() As Byte
    Const cutoff As Integer = 5

    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")

    TextToHash = asc.Getbytes_4(sTextToHash)
    SharedSecretKey = asc.Getbytes_4(sTextToHash)
    enc.key = SharedSecretKey

    bytes = enc.ComputeHash_2((TextToHash))
    BASE64SHA1 = EncodeBase64(bytes)
    BASE64SHA1 = Left(BASE64SHA1, cutoff)

    Set asc = Nothing
    Set enc = Nothing

End Function

'***************************************************************************
'Purpose: Convert a string in base 64
'Inputs:  The data as an array of bytes
'Outputs: The data as a string encoded in Base 64
'***************************************************************************
Private Function EncodeBase64(ByRef arrData() As Byte) As String

    Dim objXML As Object
    Dim objNode As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text

    Set objNode = Nothing
    Set objXML = Nothing

End Function


