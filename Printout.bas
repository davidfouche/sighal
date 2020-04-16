Attribute VB_Name = "Printout"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' MODULE_CODE:   Printout
'''                 - Any function or sub called to process printings
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
'Purpose: To printout the active sheet in PDF format
'Inputs : The active sheet of the Excel spreadsheet
'Outputs: The PDF printing of the active Excel spreadsheet
'***************************************************************************

Public Sub PDFActiveSheet()
    'www.contextures.com
    'for Excel 2010 and later
    Dim wsA As Worksheet
    Dim wbA As Workbook
    Dim wksAllSheets As Variant
    Dim strTime As String
    Dim strName As String
    Dim strPath As String
    Dim strFile As String
    Dim strPathFile As String
    Dim myFile As Variant
    On Error GoTo ErrHandler
    
    Set wbA = ActiveWorkbook
    wksAllSheets = Array("Réservations")
    ActiveWorkbook.Sheets(wksAllSheets).Select
    'strTime = Year(Now()) & Month(Now()) & Day(Now()) & Hour(Now())
    
    'get active workbook folder, if saved
    strPath = wbA.path
    If strPath = "" Then
      strPath = Application.DefaultFilePath
    End If
    strPath = strPath & "\"
    
    'replace spaces and periods in sheet name
    strName = Replace(wbA.name, " ", "")
    strName = Replace(strName, ".", "_")
    
    'create default name for savng file
    strFile = ".pdf"
    'strFile = "_" & strTime & strFile
    strFile = strName & strFile
    strPathFile = strPath & strFile
    
    'use can enter name and
    ' select folder for file
    myFile = Application.GetSaveAsFilename _
        (InitialFileName:=strPathFile, _
            FileFilter:="PDF Files (*.pdf), *.pdf", _
            Title:="Select Folder and FileName to save")
    
    'export to PDF if a folder was selected
    If myFile <> "False" Then
        wbA.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            filename:=myFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        'confirmation message with file info
        MsgBox "PDF file has been created: " _
          & vbCrLf _
          & myFile
    End If
    
exitHandler:
        Exit Sub
ErrHandler:
        MsgBox "Could not create PDF file"
        Resume exitHandler
End Sub

