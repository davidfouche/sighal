VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Guests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' SHEET_CODE:   Guests
'''                 - Handle sheet events
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 01/05/2017      David FOUCHE            Created
''' 31/05/2019      David FOUCHE            Changed
'''

' The column names
Const IdColName As String = "Identification"
Const KeyIdColName As String = "GuestKeyId"
Const GenderColName As String = "Gender"
Const LastNameHashColName As String = "LastNameHash"
Const FirstNameHashColName As String = "FirstNameHash"
Const LastNameColName As String = "LastName"
Const FirstNameColName As String = "FirstName"
Const AddressColName As String = "Address"
Const ZipCodeColName As String = "ZipCode"
Const StateColName As String = "State"
Const CityColName As String = "City"
Const CountryColName As String = "Country"
Const PhoneColName As String = "Phone"
Const EmailColName As String = "Email"

'Gender const
Const ManGender As String = "MR"
Const WomanGender As String = "MME"
Const AssociationGender As String = "ASSOCIATION"

'Default
Const DefaultCountryIndex = 74   'France
Const DefaultStateIndex = 1      'Ain

Public selection_precedente As String


' ###########################################################################################
' #
' #                                     PUBLIC METHODS
' #
' ###########################################################################################

Public Property Get gIdentification() As String
    gIdentification = IdColName
End Property

Public Property Get gManGender() As String
    gManGender = ManGender
End Property

Public Property Get gWomanGender() As String
    gWomanGender = WomanGender
End Property

Public Property Get gAssociation() As String
    gAssociation = AssociationGender
End Property

Public Property Get gIdColName() As String
    gIdColName = IdColName
End Property

Public Property Get gGenderColName() As String
    gGenderColName = GenderColName
End Property

Public Property Get gLastNameHashColName() As String
    gLastNameHashColName = LastNameHashColName
End Property

Public Property Get gFirstNameHashColName() As String
    gFirstNameHashColName = FirstNameHashColName
End Property

Public Property Get gLastNameColName() As String
    gLastNameColName = LastNameColName
End Property

Public Property Get gFirstNameColName() As String
    gFirstNameColName = FirstNameColName
End Property

Public Property Get gAddressColName() As String
    gAddressColName = AddressColName
End Property

Public Property Get gZipCodeColName() As String
    gZipCodeColName = ZipCodeColName
End Property

Public Property Get gStateColName() As String
    gStateColName = StateColName
End Property

Public Property Get gCityColName() As String
    gCityColName = CityColName
End Property

Public Property Get gCountryColName() As String
    gCountryColName = CountryColName
End Property

Public Property Get gPhoneColName() As String
    gPhoneColName = PhoneColName
End Property

Public Property Get gEmailColName() As String
    gEmailColName = EmailColName
End Property

Public Property Get gKeyIdColName() As String
    gKeyIdColName = KeyIdColName
End Property


' ###########################################################################################
' #
' #                                     PRIVATE METHODS
' #
' ###########################################################################################

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim guest As clsGuest
    
    On Error GoTo ErrHandler:

        If Not defaultCrypto.isReady Then
            'Notify the user
            MsgBox "Veuillez saisir le mot de passe RGPD dans l'onglet Configuration.", vbOKOnly + vbInformation, "Mot de passe RGPD invalide"
            Error.Raise 1, "Guests::Worksheet_BeforeDoubleClick", "Crypto engine not ready"
        End If
        
        Select Case Cells(Target.row, Range(GenderColName).Column).Value
            Case ManGender:
                GuestForm.OptionMale.Value = True
            Case WomanGender:
                GuestForm.OptionFemale.Value = True
            Case AssociationGender:
                GuestForm.OptionAssociation.Value = True
        End Select
        With GuestForm
            .Lastname.Value = defaultCrypto.Decrypt(Cells(Target.row, Range(LastNameColName).Column).Value)
            .Firstname.Value = defaultCrypto.Decrypt(Cells(Target.row, Range(FirstNameColName).Column).Value)
            .Address1.Value = defaultCrypto.Decrypt(Cells(Target.row, Range(AddressColName).Column).Value)
            .Zipcode.Value = defaultCrypto.Decrypt(Cells(Target.row, Range(ZipCodeColName).Column).Value)
            .ListDept.ListIndex = 0
            .City.Value = Cells(Target.row, Range(CityColName).Column).Value
            'Default country selected
            .ListCountry.ListIndex = DefaultCountryIndex
            'Default state selected
            .ListDept.ListIndex = DefaultStateIndex
            .Phone.Value = defaultCrypto.Decrypt(Cells(Target.row, Range(PhoneColName).Column).Value)
            .Email.Value = defaultCrypto.Decrypt(Cells(Target.row, Range(EmailColName).Column).Value)
            .Show
        End With
    Cancel = True
    Exit Sub
    
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: Guests::Worksheet_BeforeDoubleClick()" & vbCrLf & "Error Description: " & Err.Description)
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call ViewSelection.UpdateSelection(selection_precedente, Target)
End Sub
