VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' CLASS_MODULE:   Mapper
'''
'''                 - use Macro 'Test' for testing and as an example
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 27/05/19        David Fouche            Created
'''

'Public properties
Public Enum MapOperations
  CLEARALLOPE = 0
  XCRYPTOPE = 1
  SEARCHOPE = 2
  UPDATEKEY = 3
  UPDATEBOOKINGORDER = 4
  UPDATEHASHOPE = 5
  UPDATEORDERAMOUNT = 6
End Enum

'Private attributes
Private rownumber, rowid As Integer

'Defaults
Const COLUMNNBRID = 1    ' Always consider column #1 of every sheet as the column that must be filled with the identifcation of the data set
Const ROWBLOCKMAX = 20

'***************************************************************************
'Purpose: Class constructor of clsMapper
'Inputs:  None
'Outputs: An instance of the class
'***************************************************************************
Private Sub Class_Initialize()
    On Error GoTo ErrHandler:
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsMapper::Class_Initialize()" & vbCrLf & "Error Description: " & Err.Description)
End Sub

' ###########################################################################################
' #
' #                                     PUBLIC METHODS
' #
' ###########################################################################################

'***************************************************************************
'Purpose: A factory method that applies a change to every range in the target sheet
'Inputs:  the reference of the sheet, the type of operation
'Outputs: the target sheet modified by the operation
'***************************************************************************
Public Sub Map(ByRef wks As Worksheet, ByVal operationtype As Integer)
    Dim rowRange As Range
    Dim rrow As Range
    Dim sid As String
    Dim dummy As String

    On Error GoTo ErrHandler:
        Set rowRange = wks.usedrange.Rows
        rownumber = rowRange.Count
    
        Logging.logINFO ("clsMapper::Map number of rows is " & rownumber)
        Logging.logINFO ("clsMapper::Map proceeding with operation #" & operationtype & " on sheet " & wks.name)
        
        rowid = 1
        
        For Each rrow In rowRange.offset(1, 0)
            ' Extra conditions to Exit loop
            sid = rrow.Cells(1, COLUMNNBRID).Value
            If Len(sid) = 0 Then
                Logging.logINFO ("clsMapper::Map Empty line " & rowid & " rows of " & rownumber & " have been processed")
                Exit For
            End If
            If rowid = rownumber Then
                Logging.logINFO ("clsMapper::Map index reached the number of rows : " & rowid - 1 & " rows of " & rownumber & " have been processed")
                Exit For
            End If
            
            'Call the operation
            Select Case operationtype
                Case CLEARALLOPE:
                    
                Case XCRYPTOPE:
                    Call XCrypt(rrow)
                    
                Case UPDATEKEY:
                    Call UpdateKeyStatus(rrow)
                    
                Case UPDATEORDERAMOUNT:
                    Call RecalculatePrices(rrow)
                
                Case UPDATEBOOKINGORDER:
                    Call RecalculatePrices(rrow)
                    Call UpdateBookingnumber(rrow)
                    
                Case UPDATEHASHOPE:
                    Call RecalculateHash(rrow)
                    
                Case Else:
                    Error.Raise 1, "clsMapper::Map ", "Map method " & operationtype & " is unknown"
            End Select

            'Progress and increment
            Call NotifyProgress
            rowid = rowid + 1
            
        Next rrow
        
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsMapper::Map()" & vbCrLf & "Error Description: " & Err.Description)
End Sub

' ###########################################################################################
' #
' #                                     PRIVATE METHODS
' #
' ###########################################################################################

'***************************************************************************
'Purpose: Notify progress through the log file
'Inputs:  None
'Outputs: The log file updated with a statement of the progress
'***************************************************************************
Private Sub NotifyProgress()
    If (rowid Mod ROWBLOCKMAX = 0) Then
        Logging.logINFO ("clsMapper::NotifyProgress " & rowid & " rows of " & rownumber & " have been processed")
    End If
End Sub

'***************************************************************************
'Purpose: The method that reciphers the identity of every guest who are still recorded in the sheet Guests
'Inputs:  None
'Outputs: The sheet updated
'***************************************************************************
Private Sub XCrypt(rrow As Range)
    Dim keyid As String
    Dim datatoxcrypt As New Collection

    On Error GoTo ErrHandler:
        defaultCrypto.GPRDPassphrase = Settings.GPRDPassword.Value
        defaultCrypto.newGPRDPassphrase = NewPassphraseForm.NewPassphraseTBox.Value
        
        'Read the cells
        keyid = str(rrow.Cells(1, Range("GuestKeyId").Column).Value)
        datatoxcrypt.Add (rrow.Cells(1, Range("LastName").Column).Value)
        datatoxcrypt.Add (rrow.Cells(1, Range("FirstName").Column).Value)
        datatoxcrypt.Add (rrow.Cells(1, Range("Address").Column).Value)
        datatoxcrypt.Add (rrow.Cells(1, Range("ZipCode").Column).Value)
        datatoxcrypt.Add (rrow.Cells(1, Range("Email").Column).Value)
        datatoxcrypt.Add (rrow.Cells(1, Range("Phone").Column).Value)
        
        'Call the transcryption and update the key id
        Call defaultCrypto.CryptWithNewKey(keyid, datatoxcrypt)
        
        'Update the cells
        rrow.Cells(1, Range("LastName").Column).Value = datatoxcrypt(1)
        rrow.Cells(1, Range("FirstName").Column).Value = datatoxcrypt(2)
        rrow.Cells(1, Range("Address").Column).Value = datatoxcrypt(3)
        rrow.Cells(1, Range("ZipCode").Column).Value = datatoxcrypt(4)
        rrow.Cells(1, Range("Email").Column).Value = datatoxcrypt(5)
        rrow.Cells(1, Range("Phone").Column).Value = datatoxcrypt(6)
        rrow.Cells(1, Range("GuestKeyId").Column).Value = keyid
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsMapper::XCrypt()" & vbCrLf & "Error Description: " & Err.Description)
    Error.Raise 1, "Map::XCrypt ", "Failed"
End Sub

'***************************************************************************
'Purpose: Update the hash of the guest identities
'Inputs:  None
'Outputs: The hashes updated with the new algotithm or new hash key
'***************************************************************************
Private Sub RecalculateHash(rrow As Range)
    Dim guest As New clsGuest
    Dim prevguestidentifier As String
    
    On Error GoTo ErrHandler:
        'Keep the previous value of the guest identification for replacement in Booking sheet
        With guest
            .gender = rrow.Cells(1, Range("Gender").Column).Value
            .Firstname = defaultCrypto.Decrypt(rrow.Cells(1, Range("FirstName").Column).Value)
            .Lastname = defaultCrypto.Decrypt(rrow.Cells(1, Range("LastName").Column).Value)
            .postaladdress = defaultCrypto.Decrypt(rrow.Cells(1, Range("Address").Column).Value)
            .Zipcode = defaultCrypto.Decrypt(rrow.Cells(1, Range("ZipCode").Column).Value)
            .City = rrow.Cells(1, Range("City").Column).Value
            .state = rrow.Cells(1, Range("State").Column).Value
            .Country = rrow.Cells(1, Range("Country").Column).Value
            .Email = defaultCrypto.Decrypt(rrow.Cells(1, Range("Email").Column).Value)
            .Phone = defaultCrypto.Decrypt(rrow.Cells(1, Range("Phone").Column).Value)
        End With

        prevguestidentifier = rrow.Cells(1, Range("Identification").Column).Value
        'Replace the prevguestidentifier by guest.Identifier in the column Guestid of the sheet Bookings
        Bookings.Range("Guestid").Replace What:=prevguestidentifier, Replacement:=guest.Identifier, LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        rrow.Cells(1, Range("Identification").Column).Value = guest.Identifier
        
        rrow.Cells(1, Range("FirstNameHash").Column).Value = guest.FirstnameHash
        rrow.Cells(1, Range("LastNameHash").Column).Value = guest.LastnameHash
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsMapper::RecalculateHash()" & vbCrLf & "Error Description: " & Err.Description)
    Error.Raise 1, "Map::RecalculateHash ", "Failed"
End Sub

'***************************************************************************
'Purpose: Update the price of the reservations
'Inputs:  None
'Outputs: The amounts of the reservation updated
'***************************************************************************
Private Sub RecalculatePrices(rrow As Range)
    Dim bookingorder As New clsBookingOrder
    Dim touristtax As New clsTouristTax
    
    On Error GoTo ErrHandler:

    bookingorder.checkinDate = rrow.Cells(1, Range(Bookings.gCheckinColName).Column).Value
    bookingorder.checkoutDate = rrow.Cells(1, Range(Bookings.gCheckoutColName).Column).Value
    bookingorder.nbrcategory1 = rrow.Cells(1, Range(Bookings.gGuestType1ColName).Column).Value
    bookingorder.nbrcategory2 = rrow.Cells(1, Range(Bookings.gGuestType2ColName).Column).Value
    bookingorder.nbrcategory3 = rrow.Cells(1, Range(Bookings.gGuestType3ColName).Column).Value
    bookingorder.nbrcategory4 = rrow.Cells(1, Range(Bookings.gGuestType4ColName).Column).Value
    'TODO : refactor the class to make the fares and the columns and rows of the sheet BookingRate updated by the init of the class
    'TODO : replace the code below that is only written to make things work but is very anti-pattern
    Select Case rrow.Cells(1, Range(Bookings.gBookModeColName).Column).Value:
        Case "Gestion libre":
            bookingorder.bookingmode = COMPREHENSIVEBOOKING
            
        Case "Location individuelle":
            bookingorder.bookingmode = INDIVIDUALBOOKING
            
        Case "Camping":
            bookingorder.bookingmode = CAMPINGBOOKING
    End Select
    
    rrow.Cells(1, Range(Bookings.gTotalColName).Column).Value = bookingorder.TotalAmount
    rrow.Cells(1, Range(Bookings.gDepositColName).Column).Value = bookingorder.DepositAmount
    rrow.Cells(1, Range(Bookings.gTaxColName).Column).Value = touristtax.TouristTax2019(bookingorder.TotalAmount)
    
    Exit Sub

ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsMapper::UpdateBookingnumber()" & vbCrLf & "Error Description: " & Err.Description)
    Error.Raise 1, "Map::RecalculatePrices", "Failed"

End Sub

'***************************************************************************
'Purpose: Refresh the booking numbers when missing
'Inputs:  None
'Outputs: The sheet updated with the booking numbers
'***************************************************************************
Private Sub UpdateBookingnumber(rrow As Range)
    Dim bookingorder As New clsBookingOrder

    On Error GoTo ErrHandler:

    bookingorder.checkinDate = rrow.Cells(1, Range("Checkin").Column).Value
    rrow.Cells(1, Range("BookingNumber").Column).Value = bookingorder.OrderId
    
    Exit Sub

ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsMapper::UpdateBookingnumber()" & vbCrLf & "Error Description: " & Err.Description)
    Error.Raise 1, "Map::UpdateBookingnumber", "Failed"

End Sub

'***************************************************************************
'Purpose: Update the key status: a new key becomes active, an active key turns into obsolete
'Inputs:  None
'Outputs: The status of the key changed in the sheet Keys
'***************************************************************************
Private Sub UpdateKeyStatus(rrow As Range)

    On Error GoTo ErrHandler:

    Select Case rrow.Cells(1, Range("KeyStatus").Column).Value:
        Case ACTIVEKEYSTATUS:
            rrow.Cells(1, Range("KeyStatus").Column).Value = OBSOLETEKEYSTATUS
            
        Case NEWKEYSTATUS:
            rrow.Cells(1, Range("KeyStatus").Column).Value = ACTIVEKEYSTATUS
            
    End Select
    
    Exit Sub

ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsMapper::UpdateKeyStatus()" & vbCrLf & "Error Description: " & Err.Description)
    Error.Raise 1, "Map::UpdateKeyStatus", "Failed"

End Sub

