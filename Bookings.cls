VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Bookings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' SHEET_CODE:   Bookings
'''                 - Handle sheet events
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 01/05/2017      David FOUCHE            Created
''' 31/05/2019      David FOUCHE            Changed
'''


' The column names
Const CheckinColName = "Checkin"
Const CheckoutColName = "Checkout"
Const DepositColName = "DepositAmount"
Const TotalColName = "TotalAmount"
Const BkNbrColName = "BookingNumber"
Const BookModeColName = "BookingMode"
Const GuestColName = "GuestId"
Const InvoiceColName = "InvoiceNumber"
Const TaxColName = "TaxAmount"
Const CancelColName = "CancelDate"
Const DepositPayColName = "DepositPayDate"
Const InvoicePayColName = "InvoicePayDate"
Const GuestType1ColName = "Category1"
Const GuestType2ColName = "Category2"
Const GuestType3ColName = "Category3"
Const GuestType4ColName = "Category4"
Const EncFirstNameColName = "EncFirstName"
Const EncLastNameColName = "EncLastName"
Const EncAddressColName = "EncAddress"
Const EncZipCodeColName = "EncZipCode"

Public selection_precedente As String

Private Const YearIndex0 As Integer = 2010


' ###########################################################################################
' #
' #                                     PUBLIC METHODS
' #
' ###########################################################################################

Public Property Get gIdentification() As String
    gIdentification = BkNbrColName
End Property


Public Property Get gCheckinColName() As String
    gCheckinColName = CheckinColName
End Property

Public Property Get gCheckoutColName() As String
    gCheckoutColName = CheckoutColName
End Property

Public Property Get gDepositColName() As String
    gDepositColName = DepositColName
End Property

Public Property Get gTotalColName() As String
    gTotalColName = TotalColName
End Property

Public Property Get gBkNbrColName() As String
    gBkNbrColName = BkNbrColName
End Property

Public Property Get gBookModeColName() As String
    gBookModeColName = BookModeColName
End Property

Public Property Get gGuestColName() As String
    gGuestColName = GuestColName
End Property

Public Property Get gInvoiceColName() As String
    gInvoiceColName = InvoiceColName
End Property

Public Property Get gTaxColName() As String
    gTaxColName = TaxColName
End Property

Public Property Get gCancelColName() As String
    gCancelColName = CancelColName
End Property

Public Property Get gDepositPayColName() As String
    gDepositPayColName = DepositPayColName
End Property

Public Property Get gInvoicePayColName() As String
    gInvoicePayColName = InvoicePayColName
End Property

Public Property Get gGuestType1ColName() As String
    gGuestType1ColName = GuestType1ColName
End Property

Public Property Get gGuestType2ColName() As String
    gGuestType2ColName = GuestType2ColName
End Property

Public Property Get gGuestType3ColName() As String
    gGuestType3ColName = GuestType3ColName
End Property

Public Property Get gGuestType4ColName() As String
    gGuestType4ColName = GuestType4ColName
End Property

Public Function GetYearIndex0() As Integer
    GetYearIndex0 = YearIndex0
End Function

'Update each hash of guest when required
Public Sub UpdateHash(ByVal prevguestidentifier As String, ByVal newguestidentifier As String)

    On Error GoTo ErrHandler:
    
    Bookings.Range("Guestid").Replace What:=prevguestidentifier, Replacement:=newguestidentifier, LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: Bookings::UpdateHash()" & vbCrLf & "Error Description: " & Err.Description)
    Error.Raise 1, "Map::RecalculatePrices", "Failed"

End Sub

' ###########################################################################################
' #
' #                                     PRIVATE METHODS
' #
' ###########################################################################################

Private Sub Worksheet_Activate()
    Application.EnableEvents = False
    
    ' Add events on key pressed
    Application.OnKey "{ENTER}", "ResaModify_EventHandler"
    
    Application.EnableEvents = True
    
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Call ResaModify_EventHandler(Target)
    Cancel = True
End Sub

Private Sub ResaModify_EventHandler(ByRef Target As Range)
    With ReservationForm
        If IsDate(Cells(Target.row, Range(CheckinColName).Column).Value) Then
            'The reservation must be updated
            .ListBoxCheckinDay.ListIndex = Day(Cells(Target.row, Range(CheckinColName).Column).Value) - 1
            .ListBoxCheckinMonth.ListIndex = Month(Cells(Target.row, Range(CheckinColName).Column).Value) - 1
            .ListBoxCheckinYear.ListIndex = Year(Cells(Target.row, Range(CheckinColName).Column).Value) - YearIndex0
            .ListBoxCheckoutDay.ListIndex = Day(Cells(Target.row, Range(CheckoutColName).Column).Value) - 1
            .ListBoxCheckoutMonth.ListIndex = Month(Cells(Target.row, Range(CheckoutColName).Column).Value) - 1
            .ListBoxCheckoutYear.ListIndex = Year(Cells(Target.row, Range(CheckoutColName).Column).Value) - YearIndex0
            .ListBookingGuest.Value = Cells(Target.row, Range(GuestColName).Column).Value
            .ListBookingMode.Value = Cells(Target.row, Range(BookModeColName).Column).Value
            .TextBoxAdhEnfantNbr.Value = Cells(Target.row, Range(GuestType1ColName).Column).Value
            .TextBoxAdhAdulteNbr.Value = Cells(Target.row, Range(GuestType2ColName).Column).Value
            .TextBoxNadhEnfantNbr.Value = Cells(Target.row, Range(GuestType3ColName).Column).Value
            .TextBoxNadhAdulteNbr.Value = Cells(Target.row, Range(GuestType4ColName).Column).Value
            If defaultCrypto.isReady Then
                .FirstNameTBox.Value = defaultCrypto.Decrypt(Cells(Target.row, Range(EncFirstNameColName).Column).Value)
                .LastNameTBox.Value = defaultCrypto.Decrypt(Cells(Target.row, Range(EncLastNameColName).Column).Value)
            End If
            If IsDate(Cells(Target.row, Range(CancelColName).Column).Value) Then
                .CheckCancelled.Value = True
            End If
            If IsDate(Cells(Target.row, Range(DepositPayColName).Column).Value) Then
                .CheckArrhes.Value = True
            End If
            If IsDate(Cells(Target.row, Range(InvoicePayColName).Column).Value) Then
                .CheckInvoice.Value = True
            End If
        Else
            'The reservation must be created
            .ListBoxCheckinDay.ListIndex = Day(Now()) - 1
            .ListBoxCheckinMonth.ListIndex = Month(Now()) - 1
            .ListBoxCheckinYear.ListIndex = Year(Now()) - YearIndex0
            .ListBoxCheckoutDay.ListIndex = Day(Now()) - 1
            .ListBoxCheckoutMonth.ListIndex = Month(Now()) - 1
            .ListBoxCheckoutYear.ListIndex = Year(Now()) - YearIndex0
            .ListBookingMode.ListIndex = 1
            .TextBoxAdhEnfantNbr.Value = 0
            .TextBoxAdhAdulteNbr.Value = 0
            .TextBoxNadhEnfantNbr.Value = 0
            .TextBoxNadhAdulteNbr.Value = 0
        End If
        .Show
    End With
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Hide changes from being viewed on the screen during processing
    Application.ScreenUpdating = False
    Call ViewSelection.UpdateSelection(selection_precedente, Target)
    'Display changes
    Application.ScreenUpdating = True
End Sub

