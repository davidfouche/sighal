VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planning_template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' SHEET_CODE:   Planning_template
'''                 - Handle sheet events
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 01/05/2017      David FOUCHE            Created

Private MonthOffset As Integer

Public Sub SetMonthOffset(offset As Integer)
    MonthOffset = offset
    Cells(1, 1).Value = offset
End Sub

Private Sub Worksheet_Activate()
    Dim i As Integer
    Dim selectDate As Date
    Dim ResetListboxesDone As Boolean
    
    MonthOffset = Cells(1, 1).Value
    
    'Prevent Listboxes event handler
    ResetListboxesDone = False
    
    selectDate = DateSerial(Year(Now()), Month(Now()) + MonthOffset, 1)
    
    Call Planning.initialize(selectDate)
    Call Planning.render

    ResetListboxesDone = True
End Sub

