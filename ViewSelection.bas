Attribute VB_Name = "ViewSelection"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' MODULE_CODE:   ViewSelection
'''                 - Any function or sub called to handle direct selection of data
'''             from the Excel tab
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 31/05/2019      David FOUCHE            Created
'''

' ###########################################################################################
' #
' #                                     PUBLIC METHODS
' #
' ###########################################################################################

'***************************************************************************
'Purpose: To color and highlight a selection in the Bookings sheet. Select color according to the status of the booking.
'Inputs : Range of the cells that will be colored, highlighted. The row containing the booking data in the tab Bookings.
'Outputs: The range of cells colored and highlighted.
'***************************************************************************

Public Sub SetInterior(ByRef rng As Range, ByRef resarow As Range)
    'Depending on the status returns the proper shape
    'Reservation status
    If IsDate(resarow.Cells(1, Range("Checkin").Column).Value) Then
        rng.Interior.Color = Settings.Range("BookingPendingColorCode").Interior.Color
        rng.Interior.Pattern = Settings.Cells(1, Range("BookingPendingColorCode").Column).Interior.Pattern
        rng.Interior.PatternColor = Settings.Cells(1, Range("BookingPendingColorCode").Column).Interior.PatternColor
        'Cancel status
        If IsDate(resarow.Cells(1, Range("CancelDate").Column).Value) Then
            'Deposit status
            If IsDate(resarow.Cells(1, Range("DepositPayDate").Column).Value) Then
                rng.Interior.Color = Settings.Range("CancelledDepositPaidColorCode").Interior.Color
                rng.Interior.Pattern = Settings.Range("CancelledDepositPaidColorCode").Interior.Pattern
                rng.Interior.PatternColor = Settings.Range("CancelledDepositPaidColorCode").Interior.PatternColor
            Else
                rng.Interior.Color = Settings.Range("CancelledBeforeDepositColorCode").Interior.Color
                rng.Interior.Pattern = Settings.Range("CancelledBeforeDepositColorCode").Interior.Pattern
                rng.Interior.PatternColor = Settings.Range("CancelledBeforeDepositColorCode").Interior.PatternColor
            End If
        Else
            'Deposit status
            If IsDate(resarow.Cells(1, Range("DepositPayDate").Column).Value) Then
                rng.Interior.Color = Settings.Range("DepositPaidColorCode").Interior.Color
                rng.Interior.PatternColor = Settings.Range("DepositPaidColorCode").Interior.PatternColor
                rng.Interior.Pattern = Settings.Range("DepositPaidColorCode").Interior.Pattern
            End If
            'Invoice status
            If IsDate(resarow.Cells(1, Range("InvoicePayDate").Column).Value) Then
                rng.Interior.Color = Settings.Range("InvoicePaidColorCode").Interior.Color
                rng.Interior.Pattern = Settings.Range("InvoicePaidColorCode").Interior.Pattern
                rng.Interior.PatternColor = Settings.Range("InvoicePaidColorCode").Interior.PatternColor
            End If
        End If
    Else
        rng.Interior.ColorIndex = xlColorIndexNone
        rng.Interior.PatternColor = xlColorIndexNone
    End If
End Sub

'***************************************************************************
'Purpose: To bold and highlight a selection in the Guests sheet.
'Inputs : The row containing the booking dats in the tab Bookings.
'Outputs: The range of cells in bold.
'***************************************************************************

Public Sub UpdateSelection(ByRef selection As String, ByRef Target As Range)
        If selection <> "" Then
        If Range(selection).row <> Target.row Then
            With Range(selection)
                'Suppression de la couleur de fond de la sélection précédente :
                .Interior.ColorIndex = xlColorIndexNone
                .Font.Bold = False
                'Default font size is 11
                .Font.Size = 11
            End With
            With Target
                'Coloration de la sélection actuelle :
                .EntireRow.Font.Bold = True
                .EntireRow.Font.Size = 15
                Call SetInterior(.EntireRow, .EntireRow)
            End With
        End If
    End If
    'Enregistrement de l'adresse de la sélection actuelle :
    selection = Target.EntireRow.Address

End Sub
