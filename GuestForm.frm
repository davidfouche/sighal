VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GuestForm 
   Caption         =   "Enregistrement résident"
   ClientHeight    =   5200
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   9930
   OleObjectBlob   =   "GuestForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GuestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' FORM_CODE:   GuestForm
'''                 - Handle form events
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 01/05/2017      David FOUCHE            Created
''' 31/05/2019      David FOUCHE            Changed
'''

' ###########################################################################################
' #
' #                                     PUBLIC METHODS
' #
' ###########################################################################################


' ###########################################################################################
' #
' #                                     PRIVATE METHODS
' #
' ###########################################################################################

Private Sub CancelButton_Click()
    Unload Me
End Sub


Private Sub SaveAndBookButton_Click()
    On Error GoTo ErrHandler:
    
ErrHandler:
    Unload Me
    Logging.logFATAL ("SaveAndBookButton_Click : " & Err.Description)
    Err.Raise 1, "GuestForm::SaveButton_Click", "Error"
End Sub

Private Sub SaveButton_Click()
    On Error GoTo ErrHandler:

    Call UpdateGuestRecord
    
    Unload Me
    Exit Sub
ErrHandler:
    Unload Me
    Logging.logFATAL ("SaveButton_Click : " & Err.Description)
    Err.Raise 1, "GuestForm::SaveButton_Click", "Error"
End Sub

Private Sub UpdateGuestRecord()
    Dim guest As New clsGuest
    
    On Error GoTo ErrHandler:
    
    'Update information in the Guests sheet in the selected row
    With Range(Guests.selection_precedente)
        If OptionMale.Value = True Then
            guest.gender = Guests.gManGender
            .Cells(1, Range(Guests.gGenderColName).Column).Value = Guests.gManGender
        End If
        If OptionFemale.Value = True Then
            guest.gender = Guests.gWomanGender
            .Cells(1, Range(Guests.gGenderColName).Column).Value = Guests.gWomanGender
        End If
        If OptionAssociation.Value = True Then
            guest.gender = Guests.gAssociation
            .Cells(1, Range(Guests.gGenderColName).Column).Value = Guests.gAssociation
        End If
        With guest
            .Lastname = Lastname.Value
            .Firstname = Firstname.Value
            .postaladdress = Address1.Value & " " & Address2.Value
            .Zipcode = Zipcode.Value
            .state = ListDept.Value
            .City = City.Value
            .Country = ListCountry.Value
            .Phone = Phone.Value
            .Email = Email.Value
        End With
        .Cells(1, Range(Guests.gLastNameColName).Column).Value = defaultCrypto.Encrypt(guest.Lastname)
        .Cells(1, Range(Guests.gFirstNameColName).Column).Value = defaultCrypto.Encrypt(guest.Firstname)
        .Cells(1, Range(Guests.gAddressColName).Column).Value = defaultCrypto.Encrypt(guest.postaladdress)
        .Cells(1, Range(Guests.gZipCodeColName).Column).Value = defaultCrypto.Encrypt(guest.Zipcode)
        .Cells(1, Range(Guests.gStateColName).Column).Value = guest.state
        .Cells(1, Range(Guests.gCityColName).Column).Value = guest.City
        .Cells(1, Range(Guests.gCountryColName).Column).Value = guest.Country
        .Cells(1, Range(Guests.gPhoneColName).Column).Value = defaultCrypto.Encrypt(guest.Phone)
        .Cells(1, Range(Guests.gEmailColName).Column).Value = defaultCrypto.Encrypt(guest.Email)
        .Cells(1, Range(Guests.gLastNameHashColName).Column).Value = guest.LastnameHash
        .Cells(1, Range(Guests.gFirstNameHashColName).Column).Value = guest.FirstnameHash
        .Cells(1, Range(Guests.gKeyIdColName).Column).Value = defaultCrypto.activekeyIndex
        If Len(.Cells(1, Range(Guests.gIdColName).Column).Value) <> 0 And .Cells(1, Range(Guests.gIdColName).Column).Value <> guest.Identifier Then
            Call Bookings.UpdateHash(.Cells(1, Range(Guests.gIdColName).Column).Value, guest.Identifier)
        End If
        .Cells(1, Range(Guests.gIdColName).Column).Value = guest.Identifier
    End With
    Exit Sub
ErrHandler:
    Logging.logFATAL ("UpdateGuestRecord : " & Err.Description)
    Err.Raise 1, "GuestForm::UpdateGuestRecord", "Error"

End Sub
