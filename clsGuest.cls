VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' CLASS_MODULE:   Guest
'''
'''                 - use Macro 'Test' for testing and as an example
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 27/05/19        David Fouche            Created
'''

Private pkeyid As String
Private pencryptedrandom As String
Private pgender As String
Private pFirstname As String
Private pLastname As String
Private ppostaladdress As String
Private pZipcode As String
Private pCity As String
Private pstate As String
Private pCountry As String
Private pPhone As String
Private pEmail As String

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

Property Get keyid() As String
    keyid = pkeyid
End Property

Property Get Firstname() As String
    Firstname = pFirstname
End Property

Property Get Lastname() As String
    Lastname = pLastname
End Property

Property Get gender() As String
    gender = pgender
End Property

Property Get postaladdress() As String
    postaladdress = ppostaladdress
End Property

Property Get Zipcode() As String
    Zipcode = pZipcode
End Property

Property Get City() As String
    City = pCity
End Property

Property Get Country() As String
    Country = pCountry
End Property

Property Get state() As String
    state = pstate
End Property

Property Get Phone() As String
    Phone = pPhone
End Property

Property Get Email() As String
    Email = pEmail
End Property

Property Let Firstname(ByVal sin As String)
    pFirstname = UCase(sin)
End Property

Property Let Lastname(ByVal sin As String)
    pLastname = UCase(sin)
End Property

Property Let gender(ByVal sin As String)
    pgender = UCase(sin)
End Property

Property Let postaladdress(ByVal sin As String)
    ppostaladdress = UCase(sin)
End Property

Property Let Zipcode(ByVal sin As String)
    pZipcode = UCase(sin)
End Property

Property Let City(ByVal sin As String)
    pCity = UCase(sin)
End Property

Property Let Country(ByVal sin As String)
    pCountry = sin
End Property

Property Let state(ByVal sin As String)
    pstate = UCase(sin)
End Property

Property Let Phone(ByVal sin As String)
    pPhone = sin
End Property

Property Let Email(ByVal sin As String)
    pEmail = sin
End Property

Property Get FirstnameHash() As String
    FirstnameHash = defaultCrypto.HashData(Firstname)
End Property

Property Get LastnameHash() As String
    LastnameHash = defaultCrypto.HashData(Lastname)
End Property

Property Get Identifier() As String
    Dim sfullIdentity As String
    
    sfullIdentity = pgender & pLastname & pFirstname & ppostaladdress & pZipcode & pCity & pstate & pCountry
    Identifier = defaultCrypto.HashData(sfullIdentity)
End Property
' ###########################################################################################
' #
' #                                     PRIVATE METHODS
' #
' ###########################################################################################


