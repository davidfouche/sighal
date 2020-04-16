VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReservationForm 
   Caption         =   "Formulaire de réservation"
   ClientHeight    =   7120
   ClientLeft      =   50
   ClientTop       =   370
   ClientWidth     =   10200
   OleObjectBlob   =   "ReservationForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReservationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' FORM_CODE:   ReservationForm
'''                 - Handle form events
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 01/05/2017      David FOUCHE            Created
''' 31/05/2019      David FOUCHE            Changed
'''

Private guest As New clsGuest
Private arinfoguest(5, 2) As String '5 selections max
Private aridguest(5) As String '5 selections max

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub SearchGuest()
    Dim RowNum As Long
    Dim i As Integer
    
    On Error GoTo ErrHandler:
    'Clear the combobox and the arrays
    ListBookingGuest.Clear
    Erase arinfoguest
    Erase aridguest
    
    'Search through columns firstnamehash and lastnamehash
    RowNum = 1
    i = 1
        
    Do Until Guests.Cells(RowNum, Range("Identification").Column).Value = ""
        If Guests.Cells(RowNum, Range("FirstNameHash").Column).Value & Guests.Cells(RowNum, Range("LastNameHash").Column).Value = guest.FirstnameHash & guest.LastnameHash Then
        On Error GoTo next1
            aridguest(i) = Guests.Cells(RowNum, Range("Identification").Column).Value
            arinfoguest(i, 1) = defaultCrypto.Decrypt(Guests.Cells(RowNum, Range("Address").Column).Value)
            arinfoguest(i, 2) = Guests.Cells(RowNum, Range("City").Column).Value
            i = i + 1
        End If
        
next1:
        RowNum = RowNum + 1
    Loop
    
    'Update the list if arinfoguest is not empty
    ListBookingGuest.List = arinfoguest
    
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: ReservationForm::SearchGuest()" & vbCrLf & "Error Description: " & Err.Description)

End Sub

Private Sub FirstNamePattern_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    guest.Firstname = FirstNamePattern.Value
    Call SearchGuest
End Sub

Private Sub LastNamePattern_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    guest.Lastname = LastNamePattern.Value
    Call SearchGuest
End Sub

Private Sub PrintoutButton_Click()
    Dim bookingorder As New clsBookingOrder
    
    If defaultCrypto.isReady Then
        'complete the bookingorder before print
        bookingorder.sidentifier = Range(Bookings.selection_precedente).Cells(1, Range(Bookings.gBkNbrColName).Column).Value
        bookingorder.checkinDate = DateSerial(ListBoxCheckinYear.ListIndex + Bookings.GetYearIndex0, ListBoxCheckinMonth.ListIndex + 1, ListBoxCheckinDay.ListIndex + 1)
        bookingorder.checkoutDate = DateSerial(ListBoxCheckoutYear.ListIndex + Bookings.GetYearIndex0, ListBoxCheckoutMonth.ListIndex + 1, ListBoxCheckoutDay.ListIndex + 1)
        bookingorder.nbrcategory1 = TextBoxAdhEnfantNbr.Value
        bookingorder.nbrcategory2 = TextBoxAdhAdulteNbr.Value
        bookingorder.nbrcategory3 = TextBoxNadhEnfantNbr.Value
        bookingorder.nbrcategory4 = TextBoxNadhAdulteNbr.Value
        Select Case Range(Bookings.selection_precedente).Cells(1, Range(Bookings.gBookModeColName).Column).Value:
            Case "Gestion libre":
                bookingorder.bookingmode = COMPREHENSIVEBOOKING
        
            Case "Location individuelle":
                bookingorder.bookingmode = INDIVIDUALBOOKING
        
            Case "Camping":
                bookingorder.bookingmode = CAMPINGBOOKING
                
        End Select

        With bookingorder.oRequester
            .Firstname = defaultCrypto.Decrypt(Range(Bookings.selection_precedente).Cells(1, Range("EncFirstName").Column).Value)
            .Lastname = defaultCrypto.Decrypt(Range(Bookings.selection_precedente).Cells(1, Range("EncLastName").Column).Value)
            .Zipcode = defaultCrypto.Decrypt(Range(Bookings.selection_precedente).Cells(1, Range("EncZipCode").Column).Value)
            .postaladdress = defaultCrypto.Decrypt(Range(Bookings.selection_precedente).Cells(1, Range("EncAddress").Column).Value)
            .City = Range(Bookings.selection_precedente).Cells(1, Range("BookingCity").Column).Value
            .Country = Range(Bookings.selection_precedente).Cells(1, Range("BookingCountry").Column).Value
        End With

        'print the booking order
        Call bookingorder.PrintOrder
        
        'Notify the user
        MsgBox "Le bon de réservation " & bookingorder.sidentifier & " a été créé.", vbOKOnly + vbInformation, "Bon de réservation créé"

    End If
    
End Sub

Private Sub SaveButton_Click()
    Dim resarow As Integer
    Dim checkinDate As Date
    Dim bookingorder As New clsBookingOrder

    'TODO: refactor this code which is really bad
    
    'Make the sheet active
    Bookings.Activate
    
    'Copy information into the cells
    With Range(Bookings.selection_precedente)
        .Cells(1, Range(Bookings.gCheckinColName).Column).Value = DateSerial(ListBoxCheckinYear.ListIndex + Bookings.GetYearIndex0, ListBoxCheckinMonth.ListIndex + 1, ListBoxCheckinDay.ListIndex + 1)
        .Cells(1, Range(Bookings.gCheckoutColName).Column).Value = DateSerial(ListBoxCheckoutYear.ListIndex + Bookings.GetYearIndex0, ListBoxCheckoutMonth.ListIndex + 1, ListBoxCheckoutDay.ListIndex + 1)
        'First index of the listbookingguest is 0
        .Cells(1, Range(Bookings.gGuestColName).Column).Value = aridguest(ListBookingGuest.ListIndex + 1)
        .Cells(1, Range(Bookings.gGuestType1ColName).Column).Value = Me.TextBoxAdhEnfantNbr.Value
        .Cells(1, Range(Bookings.gGuestType2ColName).Column).Value = Me.TextBoxAdhAdulteNbr.Value
        .Cells(1, Range(Bookings.gGuestType3ColName).Column).Value = Me.TextBoxNadhEnfantNbr.Value
        .Cells(1, Range(Bookings.gGuestType4ColName).Column).Value = Me.TextBoxNadhAdulteNbr.Value
        .Cells(1, Range(Bookings.gBookModeColName).Column).Value = ListBookingMode.Value
        
        bookingorder.checkinDate = .Cells(1, Range(Bookings.gCheckinColName).Column).Value
        bookingorder.checkoutDate = .Cells(1, Range(Bookings.gCheckoutColName).Column).Value
        bookingorder.nbrcategory1 = .Cells(1, Range(Bookings.gGuestType1ColName).Column).Value
        bookingorder.nbrcategory2 = .Cells(1, Range(Bookings.gGuestType2ColName).Column).Value
        bookingorder.nbrcategory3 = .Cells(1, Range(Bookings.gGuestType3ColName).Column).Value
        bookingorder.nbrcategory4 = .Cells(1, Range(Bookings.gGuestType4ColName).Column).Value
        Select Case .Cells(1, Range(Bookings.gBookModeColName).Column).Value:
            Case "Gestion libre":
                bookingorder.bookingmode = COMPREHENSIVEBOOKING
        
            Case "Location individuelle":
                bookingorder.bookingmode = INDIVIDUALBOOKING
        
            Case "Camping":
                bookingorder.bookingmode = CAMPINGBOOKING
                
            'TODO: remove the case else
            Case Else:
                bookingorder.bookingmode = INDIVIDUALBOOKING
        End Select
        
        .Cells(1, Range(Bookings.gTotalColName).Column).Value = bookingorder.TotalAmount
        .Cells(1, Range(Bookings.gDepositColName).Column).Value = bookingorder.DepositAmount
        
        'Set the booking number if the booking order is new
        If Len(.Cells(1, Range(Bookings.gBkNbrColName).Column).Value) = 0 Then
            .Cells(1, Range(Bookings.gBkNbrColName).Column) = bookingorder.OrderId
        End If
        
        If Me.CheckCancelled Then
            .Cells(1, Range(Bookings.gCancelColName).Column).Value = Now()
        Else
            .Cells(1, Range(Bookings.gCancelColName).Column).Clear
        End If
        If Me.CheckArrhes Then
            .Cells(1, Range(Bookings.gDepositPayColName).Column).Value = Now()
        Else
            .Cells(1, Range(Bookings.gDepositPayColName).Column).Clear
        End If
        If Me.CheckInvoice Then
            .Cells(1, Range(Bookings.gInvoicePayColName).Column).Value = Now()
        Else
            .Cells(1, Range(Bookings.gInvoicePayColName).Column).Clear
        End If
    End With
    
    Unload Me
    
End Sub

Private Sub TextBoxAdhAdulteNbr_Change()
    If Not IsNumeric(Me.TextBoxAdhAdulteNbr.Value) And Me.TextBoxAdhAdulteNbr.Value <> vbNullString Then
        MsgBox "Veuillez saisir un nombre entre 0 et 37"
        Me.TextBoxAdhAdulteNbr.Value = vbNullString
    End If
End Sub

Private Sub TextBoxAdhEnfantNbr_Change()
    If Not IsNumeric(Me.TextBoxAdhEnfantNbr.Value) And Me.TextBoxAdhEnfantNbr.Value <> vbNullString Then
        MsgBox "Veuillez saisir un nombre entre 0 et 37"
        Me.TextBoxAdhEnfantNbr.Value = vbNullString
    End If
End Sub

Private Sub TextBoxNadhAdulteNbr_Change()
    If Not IsNumeric(Me.TextBoxNadhAdulteNbr.Value) And Me.TextBoxNadhAdulteNbr.Value <> vbNullString Then
        MsgBox "Veuillez saisir un nombre entre 0 et 37"
        Me.TextBoxNadhAdulteNbr.Value = vbNullString
    End If
End Sub

Private Sub TextBoxNadhEnfantNbr_Change()
    If Not IsNumeric(Me.TextBoxNadhEnfantNbr.Value) And Me.TextBoxNadhEnfantNbr.Value <> vbNullString Then
        MsgBox "Veuillez saisir un nombre entre 0 et 37"
        Me.TextBoxNadhEnfantNbr.Value = vbNullString
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    Dim k, tmp As String
    Dim rngbkmode, c As Range
    Dim d As Object
    
    With ReservationForm
        .ListBookingMode.Clear
        .ListBoxCheckinDay.Clear
        .ListBoxCheckinMonth.Clear
        .ListBoxCheckinYear.Clear
        .ListBoxCheckoutDay.Clear
        .ListBoxCheckoutMonth.Clear
        .ListBoxCheckoutYear.Clear

        Set rngbkmode = Worksheets("Tarifs").Range("A2:A11")
        Set d = CreateObject("scripting.dictionary")
        
        With .ListBookingMode
            For Each c In rngbkmode
                tmp = Trim(c.Value)
                If Len(tmp) > 0 Then d(tmp) = d(tmp) + 1
            Next c

            For Each k In d.keys
                .AddItem k
            Next k
            
            Set d = Nothing
        End With
        
        For i = 1 To 31
            .ListBoxCheckinDay.AddItem i
            .ListBoxCheckoutDay.AddItem i
        Next i
        For i = 1 To 12
            .ListBoxCheckinMonth.AddItem MonthName(i)
            .ListBoxCheckoutMonth.AddItem MonthName(i)
        Next i
        For i = Bookings.GetYearIndex0 To 2050
            .ListBoxCheckinYear.AddItem i
            .ListBoxCheckoutYear.AddItem i
        Next i
        .ListBoxCheckinDay.ListIndex = 1
        .ListBoxCheckoutDay.ListIndex = 1
        .ListBoxCheckinMonth.ListIndex = 1
        .ListBoxCheckoutMonth.ListIndex = 1
        .ListBoxCheckinYear.ListIndex = 0
        .ListBoxCheckoutYear.ListIndex = 0
    End With
End Sub

