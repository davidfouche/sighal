VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

'Add a new key
Public Sub AddNewKey(ByVal sNewPassphrase As String)
    Dim lastrow As Long
    Dim newkey As New clsKey
    Dim oHashHandler As New clsHashMethod
    
    On Error GoTo ErrHandler:

    lastrow = KeyList.Cells(Rows.Count, 1).End(xlUp).row + 1
    With newkey
        .Identifier = lastrow - 1
        .status = NEWKEYSTATUS
        .hashmethod = SHA512Hash
        .hashvalue = oHashHandler.Gethash(sNewPassphrase, .hashmethod)
        .cryptomethod = AESCBCMETHODLIGHT
        .timestamp = Now()
    End With
    
    'check if another new key already exists
    If CheckNewKeyExists() Then
        Error.Raise 1, "KeyList::AddNewKey ", "A new key already exists"
    End If
    
    'Update the sheet
    KeyList.Cells(lastrow, Range("Id").Column).Value = newkey.Identifier
    KeyList.Cells(lastrow, Range("HashValue").Column).Value = newkey.hashvalue
    KeyList.Cells(lastrow, Range("HashMethod").Column).Value = newkey.hashmethod
    KeyList.Cells(lastrow, Range("CryptoAlgo").Column).Value = newkey.cryptomethod
    KeyList.Cells(lastrow, Range("KeyStatus").Column).Value = newkey.status
    KeyList.Cells(lastrow, Range("Timestamp").Column).Value = newkey.timestamp
    
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: KeyList::AddNewKey()" & vbCrLf & "Error Description: " & Err.Description)
    Error.Raise 1, "KeyList::AddNewKey", "Failed"
End Sub


'Add a new key
Public Function CheckNewKeyExists() As Boolean
    Dim i, rowcount As Integer
    
    CheckNewKeyExists = False
    rowcount = KeyList.Cells(Cells.Rows.Count, 1).End(xlUp).row
    
    For i = 2 To rowcount
        If KeyList.Cells(i, Range("KeyStatus").Column).Value = NEWKEYSTATUS Then
            CheckNewKeyExists = True
        End If
    Next i
End Function

