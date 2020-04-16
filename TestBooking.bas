Attribute VB_Name = "TestBooking"
Option Explicit
Option Base 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' TEST_CODE:   TestBooking
'''                 - Test booking data processing
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 31/05/2019      David FOUCHE            Created
'''

'***************************************************************************
'Purpose: Test that the booking number is correctly generated
'Inputs:  None
'Outputs: the booking number bookingorderid is generated accordingly to the check-in date
'***************************************************************************
Sub TestBookingNumber()
    Dim savednumber As Integer
    Dim obookingorder As New clsBookingOrder
    Dim bookingorderid As String
    
    obookingorder.checkinDate = Now()
    
    savednumber = Settings.Range("LastOrderNumber").Value
    
    bookingorderid = obookingorder.OrderId
    
    Settings.Range("LastOrderNumber").Value = savednumber
End Sub

'***************************************************************************
'Purpose: Test the call to the method to print the booking orders
'Inputs:  None
'Outputs: a booking order in PDF format must be saved
'***************************************************************************
Sub TestPrintBookingOrder()
    Dim savednumber As Integer
    Dim obookingorder As New clsBookingOrder
    Dim bookingorderid As String
    
    With obookingorder.oRequester
        .Firstname = "Alain"
        .Lastname = "Dupont"
        .Email = "a.dupont@toto.fr"
        .City = "Besançon"
        .postaladdress = "100 AVENUE DES PLATANES"
        .Phone = "0102030405"
        .Zipcode = "12345"
        .state = "Zimbabwe"
    End With
    obookingorder.checkinDate = CDate("01/01/2015") + 1
    obookingorder.checkoutDate = CDate("01/01/2015") + 2
    obookingorder.sidentifier = "150085"
    obookingorder.bookingmode = COMPREHENSIVEBOOKING
    obookingorder.nbrcategory1 = 3
    obookingorder.nbrcategory2 = 1
    obookingorder.nbrcategory3 = 4
    obookingorder.nbrcategory4 = 5
    
    Call obookingorder.PrintOrder
    
    With obookingorder.oRequester
        .Firstname = "Alice"
        .Lastname = "Durand"
        .Email = "a.durand@Yahoo.fr"
        .City = "Auchs"
        .postaladdress = "5 BOULEVARD DES CIGALES"
        .Phone = "0607080910"
        .Zipcode = "010101"
        .state = "Polynésie"
    End With
    obookingorder.checkinDate = CDate("03/05/2016") + 1
    obookingorder.checkoutDate = CDate("03/05/2016") + 7
    obookingorder.sidentifier = "160017"
    obookingorder.bookingmode = INDIVIDUALBOOKING
    obookingorder.nbrcategory1 = 2
    obookingorder.nbrcategory2 = 4
    obookingorder.nbrcategory3 = 3
    obookingorder.nbrcategory4 = 1
    
    Call obookingorder.PrintOrder
    
End Sub

'***************************************************************************
'Purpose: Test the call to the iterator through the guests sheet to update the hash of identities
'Inputs:  None
'Outputs: the hash of the last and first names must be updated
'***************************************************************************
Sub TestUpdateHashAll()
    'Test iteration through the booking sheet
    Dim mapper As clsMapper

    On Error GoTo ErrHandler:

    Set mapper = New clsMapper
    
    Call mapper.Map(Guests, UPDATEHASHOPE)
    
    Exit Sub
ErrHandler:
    Logging.logFATAL ("clsMapper::Map " & Err.Description)

End Sub

'***************************************************************************
'Purpose: Test the call to the iterator through the booking sheet to update booking number and prices
'Inputs:  None
'Outputs: the prices and numbers of booking orders in the sheet Bookings must be updated
'***************************************************************************
Sub TestUpdateBookingOrder()
    Dim mapper As clsMapper

    On Error GoTo ErrHandler:

    Set mapper = New clsMapper
    
    Call mapper.Map(Bookings, UPDATEBOOKINGORDER)
    
    Exit Sub
ErrHandler:
    Logging.logFATAL ("clsMapper::Map " & Err.Description)

End Sub

'***************************************************************************
'Purpose: Test the look up of a given first name and last name through the sheet Guests
'Inputs:  None
'Outputs: Result of the search
'***************************************************************************
Sub TestSearchByName()
    Dim Firstname, Lastname As String
    Dim guest As New clsGuest
    Dim RowNum As Long
    Dim collsearchresult As New Collection
    
    Call thiscryptoinit

    Firstname = "josiane"
    Lastname = "DEGREy"
    
    On Error GoTo ErrHandler:
    'Search the first occurence
    With guest
        .Firstname = Firstname
        .Lastname = Lastname
    End With
    
    'Search through columns firstnamehash and lastnamehash

    RowNum = 1
        
    Do Until Guests.Cells(RowNum, Range("Identification").Column).Value = ""
    
        If Guests.Cells(RowNum, Range("FirstNameHash").Column).Value & Guests.Cells(RowNum, Range("LastNameHash").Column).Value = guest.FirstnameHash & guest.LastnameHash Then
        On Error GoTo next1
            collsearchresult.Add (Guests.Cells(RowNum, Range("Identification").Column).Value)
        End If
next1:
        RowNum = RowNum + 1
    Loop

    
    Exit Sub
ErrHandler:
    Logging.logFATAL ("TestSearchByName: " & Err.Description)
    
End Sub

'***************************************************************************
'Purpose: Test the call to the iterator through the booking sheet to update booking prices
'Inputs:  None
'Outputs: the prices of booking orders and tourist taxes in the sheet Bookings must be updated
'***************************************************************************
Sub TestUpdateBookingAmounts()
    Dim mapper As clsMapper

    On Error GoTo ErrHandler:

    Set mapper = New clsMapper
    
    Call mapper.Map(Bookings, UPDATEORDERAMOUNT)
    
    Exit Sub
ErrHandler:
    Logging.logFATAL ("clsMapper::Map " & Err.Description)

End Sub

