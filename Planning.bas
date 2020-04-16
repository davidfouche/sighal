Attribute VB_Name = "Planning"
Option Explicit
Option Base 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' MODULE_CODE:   Planning
'''                 - All function or sub called to handle planning and calendar
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 31/05/2019      David FOUCHE            Created
'''

Private Const planninglinemax As Integer = 20

Dim planningline(planninglinemax) As Date
Dim planningRange As String
Dim totalBookedRange As String
Dim selectedPlanningDate As Date

' ###########################################################################################
' #
' #                                     PUBLIC METHODS
' #
' ###########################################################################################

'***************************************************************************
'Purpose: To display the booking planning of a month
'Inputs : selectedPlanningDate, the date that is selected in a planning
'Outputs: display of the planning of the month that includes the selected date
'***************************************************************************
Public Sub render()
    ' Read the reservation periods
    ' Select only reservation periods that contains the selected date
    Dim rowRange, rrow As Range
    Dim planninglinenbr As Integer
    Dim n As Long
    Dim col As Range
    Dim firstdayofmonth, LastDayOfMonth, begindate, enddate As Date
    Dim resanbrmax As Integer
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrHandler:
        firstdayofmonth = DateSerial(Year(selectedPlanningDate), Month(selectedPlanningDate), 1)
        LastDayOfMonth = DateSerial(Year(selectedPlanningDate), Month(selectedPlanningDate) + 1, 0)
        
        Set rowRange = Bookings.usedrange.Rows
        For Each rrow In rowRange.offset(1, 0)
            begindate = rrow.Cells(1, Range(Bookings.gCheckinColName).Column).Value
            enddate = rrow.Cells(1, Range(Bookings.gCheckoutColName).Column).Value
            If firstdayofmonth <= enddate And LastDayOfMonth >= begindate Then
                If begindate < firstdayofmonth Then
                    begindate = firstdayofmonth
                End If
                planninglinenbr = getPlanningline(begindate, enddate)
                Call displaySegAtLine(planninglinenbr, rrow)
            End If
            If Len(rrow.Cells(1, Range(Bookings.gBkNbrColName).Column).Value) = 0 Then
                Exit For
            End If
        Next rrow
    
        n = 1
        For Each col In Range(planningRange).Columns
            Range(totalBookedRange).Cells(1, n).Value = calculateSum(col)
            n = n + 1
        Next col
        
        Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: Planning::render()" & vbCrLf & "Error Description: " & Err.Description)
    Application.ScreenUpdating = True
End Sub

'***************************************************************************
'Purpose: Clear the planning before to render and refresh
'Inputs : selectedDate, the date of a month that has been selected for rendering
'Outputs: A planning that is cleared and ready for update
'***************************************************************************
Public Sub initialize(ByVal selectedDate As Date)
    ' Reinitialize all cells of the planning
    Dim i As Integer
    Dim firstdayofmonth As Date
    Dim name As String
    
    Application.ScreenUpdating = False
    selectedPlanningDate = selectedDate
    
    ' Resize planning accordingly with the last day of month
    firstdayofmonth = DateSerial(Year(selectedPlanningDate), Month(selectedPlanningDate), 1)
    Call resizeRange(planningRange, totalBookedRange, selectedPlanningDate)

    ActiveSheet.Range(planningRange).Select
    With selection
        .UnMerge
        .ClearContents
        .ClearComments
        .Interior.ColorIndex = xlColorIndexNone
        .Borders.LineStyle = xlContinuous
    End With
    With Range(totalBookedRange)
        .ClearContents
    End With
        
    For i = 1 To planninglinemax
        planningline(i) = firstdayofmonth - 1
    Next i

    'Clear current selection
    ActiveSheet.Range("A1").Select
    
    Application.ScreenUpdating = True
End Sub


' ###########################################################################################
' #
' #                                     PRIVATE METHODS
' #
' ###########################################################################################

'***************************************************************************
'Purpose: To set the title of a planning
'Inputs : The selected date of the month for which the planning will be updated
'Outputs: The title of the planning
'***************************************************************************
Private Function updateTitle(ByVal selectedDate)
    ' Update the tab of the sheet
    updateTitle = MonthName(Month(selectedDate)) & " " & Year(selectedDate)
End Function

'***************************************************************************
'Purpose: To create the spreadsheet tabs containing the months of the planning
'Inputs : Range of the months [NmonthsBefore, NmonthsAfter]
'Outputs: The tabs of the planning months in the spreadsheet
'***************************************************************************
Private Sub AddCopyPlanningTemplate(NmonthsBefore As Integer, NmonthsAfter As Integer)
    Dim offset As Integer
    Dim ws As Worksheet
    Dim found As Boolean
    
    For offset = NmonthsAfter To -NmonthsBefore Step -1
        found = False
        For Each ws In ThisWorkbook.Sheets
        If ws.name = updateTitle(DateSerial(Year(Now()), Month(Now()) + offset, 1)) Then
            found = True
            Exit For
        End If
        Next
        If Not found Then
            Planning_template.Copy After:=MailConfirmation_template
            ActiveSheet.name = updateTitle(DateSerial(Year(Now()), Month(Now()) + offset, 1))
            ActiveSheet.SetMonthOffset (offset)
        End If
    Next offset
End Sub

'***************************************************************************
'Purpose: To add the tabs accordingly with the settings
'Inputs : The settings
'Outputs: The tabs of the planning months in the spreadsheet
'***************************************************************************
Private Sub UpdatePlanningTabs()
    Call AddCopyPlanningTemplate(Settings.Range("PlanningStartOffset"), Settings.Range("PlanningEndOffset"))
End Sub

'***************************************************************************
'Purpose: To compute the total number of beds with overnights
'Inputs : The range of cells in the planning that must be involved in the sum
'Outputs: The sum of the booked beds
'***************************************************************************
Private Function calculateSum(ByRef rng As Range)
    'Calculate the total occupation in a given range
    Dim rCell As Range
    Dim c As Range
    calculateSum = 0
    
    For Each rCell In rng.Cells
        If rCell.MergeCells = True Then
            calculateSum = calculateSum + rCell.MergeArea.Cells(1, 1).Value
        Else
            'Not a merged cell
            calculateSum = calculateSum + rCell.Value
        End If
    Next rCell

End Function

'***************************************************************************
'Purpose: To search for a row where cells are free for the display of a period of time
'Inputs : The period of time [segmentstartdate, segmentenddate]
'Outputs: The row number of the planning where the period can be displayed
'***************************************************************************
Private Function getPlanningline(ByVal segmentstartdate As Date, ByVal segmentenddate As Date) As Integer
    Dim i As Integer
    
    'Search for the earliest segment available
    For i = 1 To planninglinemax
        If segmentstartdate > planningline(i) Then
            planningline(i) = segmentenddate
            Exit For
        End If
    Next i
    
    getPlanningline = i
End Function

'***************************************************************************
'Purpose: To set the comment of a segment of time in the planning
'Inputs : The Bookings row from which the data will be displayed in the comment
'Outputs: A string that will be displayed as comment of a segment of time
'***************************************************************************
Private Function getcomment(ByRef row As Range) As String
    Dim cancelled As String
    cancelled = row.Cells(1, Range(Bookings.gCancelColName).Column).Value
    If IsDate(row.Cells(1, Range(Bookings.gCancelColName).Column).Value) Then
        getcomment = "Id : " & row.Cells(1, Range(Bookings.gIdentification).Column).Value & vbLf
        getcomment = getcomment & "Réservation annulée"
    Else
        ' Build the comment of a cell of the planning table
        getcomment = "Id : " & row.Cells(1, Range(Bookings.gIdentification).Column).Value & vbLf
        getcomment = getcomment & "Deposit : " & row.Cells(1, Range(Bookings.gDepositColName).Column).Value & " EUR" & vbLf
        getcomment = getcomment & "Total : " & row.Cells(1, Range(Bookings.gTotalColName).Column).Value & " EUR" & vbLf
    End If
End Function

'***************************************************************************
'Purpose: To hide the cells of the planning that are out of bound (greater than the last day of the month)
'Inputs : The range of cells for the planning, the bottom line of the planning and the selected date
'Outputs: The planning range resized
'***************************************************************************
Private Sub resizeRange(ByRef planningRange As String, ByRef totalBookedRange As String, ByVal selectedDate As Date)
    Dim numberofdaysinmonth As Date
    numberofdaysinmonth = Day(DateSerial(Year(selectedDate), Month(selectedDate) + 1, 0))
    ' Display all columns
    Range("B6:AF24").Select
    selection.EntireColumn.Hidden = False
    ' For each day in month update every header in line 4 with the day of week
    ' Set columns in gray where days are out of bounds of the month
    Select Case numberofdaysinmonth
        Case 31
            planningRange = "B6:AF23"
            totalBookedRange = "B24:AF24"
        Case 30
            planningRange = "B6:AE23"
            totalBookedRange = "B24:AE24"
            Range("AF6:AF24").Select
            selection.EntireColumn.Hidden = True
        Case 29
            planningRange = "B6:AD23"
            totalBookedRange = "B24:AD24"
            Range("AE6:AF24").Select
            selection.EntireColumn.Hidden = True
        Case 28
            planningRange = "B6:AC23"
            totalBookedRange = "B24:AC24"
            Range("AD6:AF24").Select
            selection.EntireColumn.Hidden = True
    End Select
End Sub

'***************************************************************************
'Purpose: To display the book orders in the planning
'Inputs : Line number where space is available to display the booking and the row of the tab Bookings
'Outputs: The segment displayed in the planning
'***************************************************************************
Private Sub displaySegAtLine(ByVal planninglinenbr As Integer, ByRef row As Range)
    'Display a period in yellow in the planning at row number given in parameter
    Dim segmentstart, segmentend, guestsnbr As Integer
    Dim comment As String
    Dim arrival, departure As Date
    
    If IsDate(row.Cells(1, Range(Bookings.gCancelColName).Column).Value) Then
        'Booking has been cancelled
        guestsnbr = 0
    Else
        guestsnbr = row.Cells(1, Range(Bookings.gGuestType1ColName).Column).Value + _
            row.Cells(1, Range(Bookings.gGuestType2ColName).Column).Value + _
            row.Cells(1, Range(Bookings.gGuestType3ColName).Column).Value + _
            row.Cells(1, Range(Bookings.gGuestType4ColName).Column).Value
    End If
    'TODO : update the comment
    'comment = PlanningHandler.getcomment(row)
    comment = getcomment(row)
    arrival = row.Cells(1, Range(Bookings.gCheckinColName).Column).Value
    departure = row.Cells(1, Range(Bookings.gCheckoutColName).Column).Value
    
    segmentstart = 0
    segmentend = 0
    
    If Year(arrival) = Year(selectedPlanningDate) Then
        If Month(arrival) = Month(selectedPlanningDate) Then
            segmentstart = Day(arrival) + 1
        End If
        If Month(departure) = Month(selectedPlanningDate) Then
            segmentend = Day(departure) + 1
        End If
        If Month(arrival) < Month(selectedPlanningDate) Then
            segmentstart = 2
        End If
        If Month(departure) > Month(selectedPlanningDate) Then
            segmentend = 32
        End If
        If segmentstart <> 0 And segmentend <> 0 Then
            Range(ActiveSheet.Cells(planninglinenbr + 5, segmentstart), ActiveSheet.Cells(planninglinenbr + 5, segmentend)).Select
            With selection
                .HorizontalAlignment = xlCenter
                .Merge
                .Value = guestsnbr
                .ClearComments
                .Cells(1, 1).AddComment comment
            End With
            ' Change the color by calling the procedure SetInterior from ViewSelection module
            Call SetInterior(selection, row)
        End If
    End If
End Sub

