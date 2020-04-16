VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBookingOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' CLASS_MODULE:   BookingOrder
'''
'''                 - use Macro 'Test' for testing and as an example
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 27/05/19        David Fouche            Created
'''

'Public attributes
Public sidentifier As String
Public checkinDate As Date
Public checkoutDate As Date
Public oRequester As clsGuest

Public nbrcategory1, nbrcategory2, nbrcategory3, nbrcategory4 As Integer
Public amountcategory1, amountcategory2, amountcategory3, amountcategory4, amountTotal, amountDeposit As Variant
Public bookingmode As BookingModeType

'Private attributes
Private currentNumber As Integer
Private sordertemplatepath As String
Private sorderpath As String

'Define key and method types
Public Enum BookingModeType
  COMPREHENSIVEBOOKING = 0
  INDIVIDUALBOOKING = 1
  CAMPINGBOOKING = 2
End Enum

Const MAXBOOKINGID As Integer = 1000
Const ENTETEBONRESERVATION As String = "BonReservation_LOUBECO_"

'***************************************************************************
'Purpose: Class constructor of clsBookingOrder
'Inputs:  None
'Outputs: An instance of the class
'***************************************************************************
Private Sub Class_Initialize()
    On Error GoTo ErrHandler:
    Set oRequester = New clsGuest
    currentNumber = Settings.Range("LastOrderNumber").Value
    sordertemplatepath = ActiveWorkbook.path & "\" & Settings.Range("BOTemplatePath").Value
    sorderpath = ActiveWorkbook.path & "\" & Settings.Range("BookingOrderPath").Value & "\"
    If Dir(sorderpath, vbDirectory) = "" Then
        Shell ("cmd /c mkdir """ & sorderpath & """")
    End If

    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsBookingOrder::Class_Initialize()" & vbCrLf & "Error Description: " & Err.Description)
End Sub

' ###########################################################################################
' #
' #                                     PUBLIC METHODS
' #
' ###########################################################################################

'***************************************************************************
'Purpose: Calculate the sale amount of the booking order
'Inputs:  None
'Outputs: The Totalamount calculated from the overnights and, if needed, the number of guests for each category
'***************************************************************************
Property Get TotalAmount() As Variant
    Call CalculateTotalAmount
    TotalAmount = amountTotal
End Property


'***************************************************************************
'Purpose: Calculate the deposit amount
'Inputs:  None
'Outputs: The Depositamount calculated
'***************************************************************************
Property Get DepositAmount() As Variant
    Call CalculateDepositAmount
    DepositAmount = amountDeposit
End Property

'***************************************************************************
'Purpose: Provide the next booking order identifier
'Inputs:  None
'Outputs: The order identifier that will be used the reference of the next booking
'***************************************************************************
Property Get OrderId() As String
    On Error GoTo ErrHandler:
    OrderId = Format(checkinDate, "YY") & Format(currentNumber, "000")
    OrderId = OrderId & LuhnChecksum(OrderId)
    Settings.Range("LastOrderNumber").Value = Settings.Range("LastOrderNumber").Value + 1
    Settings.Range("LastOrderNumber").Value = Settings.Range("LastOrderNumber").Value Mod MAXBOOKINGID
    currentNumber = Settings.Range("LastOrderNumber").Value
    Exit Property
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsBookingOrder::OrderId()" & vbCrLf & "Error Description: " & Err.Description)
End Property

'***************************************************************************
'Purpose: Print the booking order
'Inputs:  None
'Outputs: The DOCX file of the booking order updated with the data from the current order of the Booking sheet
'***************************************************************************
Public Sub PrintOrder()
    Dim wApp As Object
    Dim wDoc As Object
    Dim sbookingOrderFileName As String
    Dim orderamount As Variant
    Dim overnights As Integer
    
    ' you must declare your variable with class' name, not instanced object's name
    Dim cc As Object
    Dim srateColumn As String

    On Error GoTo ErrHandler:
    
    srateColumn = "Bookingrate" & Format(Me.checkinDate, "YYYY")
    
    Set wApp = CreateObject("Word.Application")
    Set wDoc = CreateObject("Word.Document")
    Set wApp = GetObject(, "Word.Application")
    If wApp Is Nothing Then
        Set wApp = CreateObject("Word.Application")
    End If

    wApp.DisplayAlerts = False
    
    'Opens template to create document
    Set wDoc = wApp.Documents.Add(Template:=sordertemplatepath)
    
    sbookingOrderFileName = sorderpath & ENTETEBONRESERVATION & Format(Me.checkinDate, "YYMMDD") & "_" & Me.sidentifier & ".docx"
    
    For Each cc In wDoc.SelectContentControlsByTag("BookingNumber")
        cc.Range.Text = sidentifier
    Next
    
    For Each cc In wDoc.SelectContentControlsByTag("StayStartDate")
        cc.Range.Text = Format(checkinDate, "DD/MM/YYYY")
    Next

    For Each cc In wDoc.SelectContentControlsByTag("StayEndDate")
        cc.Range.Text = Format(checkoutDate, "DD/MM/YYYY")
    Next

    For Each cc In wDoc.SelectContentControlsByTag("Category1PricePerNight")
        cc.Range.Text = Format(Bookingrate.Cells(Range("Category1Price").row, Range(srateColumn).Column).Value, "#,##0.00")
    Next

    For Each cc In wDoc.SelectContentControlsByTag("Category2PricePerNight")
        cc.Range.Text = Format(Bookingrate.Cells(Range("Category2Price").row, Range(srateColumn).Column).Value, "#,##0.00")
    Next

    For Each cc In wDoc.SelectContentControlsByTag("Category3PricePerNight")
        cc.Range.Text = Format(Bookingrate.Cells(Range("Category3Price").row, Range(srateColumn).Column).Value, "#,##0.00")
    Next

    For Each cc In wDoc.SelectContentControlsByTag("Category4PricePerNight")
        cc.Range.Text = Format(Bookingrate.Cells(Range("Category4Price").row, Range(srateColumn).Column).Value, "#,##0.00")
    Next

    For Each cc In wDoc.SelectContentControlsByTag("RequesterLastName")
        cc.Range.Text = Me.oRequester.Lastname
    Next

    For Each cc In wDoc.SelectContentControlsByTag("RequesterFirstName")
        cc.Range.Text = Me.oRequester.Firstname
    Next

    For Each cc In wDoc.SelectContentControlsByTag("RequesterAddress")
        cc.Range.Text = Me.oRequester.postaladdress & " " & oRequester.Zipcode & " " & oRequester.City & " " & oRequester.state
    Next
    
    For Each cc In wDoc.SelectContentControlsByTag("RequesterPhone")
        cc.Range.Text = Me.oRequester.Phone
    Next

    For Each cc In wDoc.SelectContentControlsByTag("RequesterEmail")
        cc.Range.Text = Me.oRequester.Email
    Next

    For Each cc In wDoc.SelectContentControlsByTag("ReceiverAddress")
        cc.Range.Text = Settings.Range("BookingReceiverAddress").Value
    Next

    For Each cc In wDoc.SelectContentControlsByTag("ReceiverPhone")
        cc.Range.Text = Settings.Range("BookingReceiverPhone").Value
    Next

    For Each cc In wDoc.SelectContentControlsByTag("ReceiverEmail")
        cc.Range.Text = Settings.Range("BookingReceiverEmail").Value
    Next

    'Update the amounts in the order
    overnights = Me.checkoutDate - Me.checkinDate
    Call CalculateTotalAmount
    Call CalculateDepositAmount
    
    For Each cc In wDoc.SelectContentControlsByTag("TotalAmount")
        cc.Range.Text = Format(amountTotal, "#,##0.00")
    Next
    
    For Each cc In wDoc.SelectContentControlsByTag("Category1TotalAmount")
        cc.Range.Text = Format(amountcategory1, "#,##0.00")
    Next
    
    For Each cc In wDoc.SelectContentControlsByTag("Category2TotalAmount")
        cc.Range.Text = Format(amountcategory2, "#,##0.00")
    Next
    
    For Each cc In wDoc.SelectContentControlsByTag("Category3TotalAmount")
        cc.Range.Text = Format(amountcategory3, "#,##0.00")
    Next
    
    For Each cc In wDoc.SelectContentControlsByTag("Category4TotalAmount")
        cc.Range.Text = Format(amountcategory4, "#,##0.00")
    Next
    
    For Each cc In wDoc.SelectContentControlsByTag("Category1Number")
        cc.Range.Text = Me.nbrcategory1
    Next
    
    For Each cc In wDoc.SelectContentControlsByTag("Category2Number")
        cc.Range.Text = Me.nbrcategory2
    Next
    
    For Each cc In wDoc.SelectContentControlsByTag("Category3Number")
        cc.Range.Text = Me.nbrcategory3
    Next
    
    For Each cc In wDoc.SelectContentControlsByTag("Category4Number")
        cc.Range.Text = Me.nbrcategory4
    Next
    
    'The same number of overnights to all category
    For Each cc In wDoc.SelectContentControlsByTag("Category1Overnights")
        cc.Range.Text = overnights
    Next
    
    For Each cc In wDoc.SelectContentControlsByTag("Category2Overnights")
        cc.Range.Text = overnights
    Next
    
    For Each cc In wDoc.SelectContentControlsByTag("Category3Overnights")
        cc.Range.Text = overnights
    Next
    
    For Each cc In wDoc.SelectContentControlsByTag("Category4Overnights")
        cc.Range.Text = overnights
    Next
    
    For Each cc In wDoc.SelectContentControlsByTag("DepositAmount")
        cc.Range.Text = Format(DepositAmount, "#,##0.00")
    Next
    
    wDoc.SaveAs2 sbookingOrderFileName
    wDoc.Close False

    If wApp.Documents.Count = 0 Then wApp.Quit False
    
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsBookingOrder::PrintOrder()" & vbCrLf & "Error Description: " & Err.Description)
End Sub

' ###########################################################################################
' #
' #                                     PRIVATE METHODS
' #
' ###########################################################################################

'***************************************************************************
'Purpose: Calculate the total amount of the booking order
'Inputs:  None
'Outputs: An error and the private data amountTotal
'***************************************************************************
Private Sub CalculateTotalAmount()
    Dim overnights As Integer
    Dim srateColumn As String
    
    srateColumn = "Bookingrate" & Format(Me.checkinDate, "YYYY")
    overnights = Me.checkoutDate - Me.checkinDate
    
    Select Case bookingmode
        Case COMPREHENSIVEBOOKING:
            amountcategory1 = "forfait"
            amountcategory2 = "forfait"
            amountcategory3 = "forfait"
            amountcategory4 = "forfait"
            Select Case overnights
                Case 1:
                    amountTotal = Bookingrate.Cells(Range("FullComprehensive1nightPrice").row, Range(srateColumn).Column).Value
                Case 2:
                    amountTotal = Bookingrate.Cells(Range("FullComprehensive2nightPrice").row, Range(srateColumn).Column).Value
                Case 3:
                    amountTotal = Bookingrate.Cells(Range("FullComprehensive3nightPrice").row, Range(srateColumn).Column).Value
                Case 4:
                    amountTotal = Bookingrate.Cells(Range("FullComprehensive4nightPrice").row, Range(srateColumn).Column).Value
                Case Else:
                    Error.Raise 1, "totalAmount::invalid number of nights for comprehensive booking", "Coprehensive overnights " & overnights & " is not valid"
            End Select
        
        Case INDIVIDUALBOOKING:
            amountcategory1 = category1TotalAmount()
            amountcategory2 = category2TotalAmount()
            amountcategory3 = category3TotalAmount()
            amountcategory4 = category4TotalAmount()
            amountTotal = amountcategory1 + amountcategory2 + amountcategory3 + amountcategory4
        
        Case CAMPINGBOOKING:
            amountTotal = 0
            amountcategory1 = 0
            amountcategory2 = category2CampingTotalAmount()
            amountcategory3 = 0
            amountcategory4 = category4CampingTotalAmount()
        
        Case Else:
            Error.Raise 1, "totalAmount::Unknown booking type", "Bookingmode " & Me.bookingmode & " is unknown"
    End Select
End Sub

'***************************************************************************
'Purpose: Calculate the total amount for the guests in category 2 that booked a camping place
'Inputs:  None
'Outputs: the private data category2CampingTotalAmount
'***************************************************************************
Private Function category2CampingTotalAmount() As Variant
    Dim overnights As Integer
    Dim srateColumn As String
    Dim fare As Double
    
    srateColumn = "Bookingrate" & Format(Me.checkinDate, "YYYY")
    overnights = Me.checkoutDate - Me.checkinDate
    
    fare = Bookingrate.Cells(Range("CampingCategory2Price").row, Range(srateColumn).Column).Value
    category2CampingTotalAmount = nbrcategory2 * fare * overnights
End Function

'***************************************************************************
'Purpose: Calculate the total amount for the guests in category 4 that booked a camping place
'Inputs:  None
'Outputs: the private data category4CampingTotalAmount
'***************************************************************************
Private Function category4CampingTotalAmount() As Variant
    Dim overnights As Integer
    Dim srateColumn As String
    Dim fare As Double
    
    srateColumn = "Bookingrate" & Format(Me.checkinDate, "YYYY")
    overnights = Me.checkoutDate - Me.checkinDate
    
    fare = Bookingrate.Cells(Range("CampingCategory4Price").row, Range(srateColumn).Column).Value
    category4CampingTotalAmount = nbrcategory4 * fare * overnights
End Function


'***************************************************************************
'Purpose: Calculate the total amount for the guests in category 1 that booked in-house beds
'Inputs:  None
'Outputs: the private data category1TotalAmount
'***************************************************************************
Private Function category1TotalAmount() As Variant
    Dim overnights As Integer
    Dim srateColumn As String
    Dim fare As Double
    
    srateColumn = "Bookingrate" & Format(Me.checkinDate, "YYYY")
    overnights = Me.checkoutDate - Me.checkinDate
    
    fare = Bookingrate.Cells(Range("Category1Price").row, Range(srateColumn).Column).Value
    category1TotalAmount = nbrcategory1 * fare * overnights
End Function

'***************************************************************************
'Purpose: Calculate the total amount for the guests in category 2 that booked in-house beds
'Inputs:  None
'Outputs: the private data category2TotalAmount
'***************************************************************************
Private Function category2TotalAmount() As Variant
    Dim overnights As Integer
    Dim srateColumn As String
    Dim fare As Double
    
    srateColumn = "Bookingrate" & Format(Me.checkinDate, "YYYY")
    overnights = Me.checkoutDate - Me.checkinDate
    
    fare = Bookingrate.Cells(Range("Category2Price").row, Range(srateColumn).Column).Value
    category2TotalAmount = nbrcategory2 * fare * overnights
End Function

'***************************************************************************
'Purpose: Calculate the total amount for the guests in category 3 that booked in-house beds
'Inputs:  None
'Outputs: the private data category3TotalAmount
'***************************************************************************
Private Function category3TotalAmount() As Variant
    Dim overnights As Integer
    Dim srateColumn As String
    Dim fare As Double
    
    srateColumn = "Bookingrate" & Format(Me.checkinDate, "YYYY")
    overnights = Me.checkoutDate - Me.checkinDate
    
    fare = Bookingrate.Cells(Range("Category3Price").row, Range(srateColumn).Column).Value
    category3TotalAmount = nbrcategory3 * fare * overnights
End Function

'***************************************************************************
'Purpose: Calculate the total amount for the guests in category 4 that booked in-house beds
'Inputs:  None
'Outputs: the private data category4TotalAmount
'***************************************************************************
Private Function category4TotalAmount() As Variant
    Dim overnights As Integer
    Dim srateColumn As String
    Dim fare As Double
    
    srateColumn = "Bookingrate" & Format(Me.checkinDate, "YYYY")
    overnights = Me.checkoutDate - Me.checkinDate
    
    fare = Bookingrate.Cells(Range("Category4Price").row, Range(srateColumn).Column).Value
    category4TotalAmount = nbrcategory4 * fare * overnights
End Function

'***************************************************************************
'Purpose: Calculate the deposit amount accordingly with the value of the cell named DepositRate
'Inputs:  None
'Outputs: the private data amountDeposit
'***************************************************************************
Private Sub CalculateDepositAmount()
    Dim srateColumn As String
    
    srateColumn = "Bookingrate" & Format(Me.checkinDate, "YYYY")
    
    amountDeposit = Round(Bookingrate.Cells(Range("DepositRate").row, Range(srateColumn).Column).Value * amountTotal)
End Sub

'***************************************************************************
'Purpose: Calculate the Luhn checksum of a string. This method is called by the method that outputs the order identifier
'Inputs:  None
'Outputs: the checksum of the string base on the Luhn algotithm
'***************************************************************************
Private Function LuhnChecksum(strNum As String) As Integer
    Dim i As Integer
    Dim iEven As Integer
    Dim iOdd As Integer
    Dim iTotal As Integer
    Dim strOneChar As String
    Dim iTemp As Integer
    
    ' Add digits in even ordinal positions
    ' starting from rightmost
    For i = Len(strNum) - 1 To 2 Step -2
    
    strOneChar = Mid$(strNum, i, 1)
    If IsNumeric(strOneChar) Then
    iEven = iEven + CInt(strOneChar)
    End If
    Next i
    
    ' Process digits in odd ordinal positions
    ' starting from rightmost
    For i = Len(strNum) To 1 Step -2
    strOneChar = Mid$(strNum, i, 1)
    If IsNumeric(strOneChar) Then
    ' Double it
    iTemp = CInt(strOneChar) * 2
    If iTemp > 9 Then
    ' Break the digits (e.g., 19 becomes 1+9)
    iOdd = iOdd + (iTemp \ 10) + (iTemp - 10)
    Else
    iOdd = iOdd + iTemp
    End If
    End If
    Next i
    
    ' Add even and odd
    iTotal = iEven + iOdd
    
    ' Return the 10's complement
    LuhnChecksum = 10 - (iTotal Mod 10)
    If LuhnChecksum = 10 Then LuhnChecksum = 0

End Function

