VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBackupHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' CLASS_MODULE:   Backup
'''
'''                 - use Macro 'Test' from 'TestBackup' for testing and as an example
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 27/05/19        David Fouche            Created
'''

'Public properties
Private iCSVIndex As Integer
Private sBackupPath As String

'Defaults
Const DEFAULT_CSVINDEXMAX = 10

'***************************************************************************
'Purpose: Class constructor of clsBackupHandler
'Inputs:  None
'Outputs: An instance of the class
'***************************************************************************
Private Sub Class_Initialize()
    On Error GoTo ErrHandler:
    iCSVIndex = Int(Settings.Range("CurrentExportNumber").Value)
    Call UpdateIndex
    sBackupPath = ActiveWorkbook.path & "\" & Settings.Range("CSVBackupPath") & "\"
    If Dir(sBackupPath, vbDirectory) = "" Then
        Shell ("cmd /c mkdir """ & sBackupPath & """")
    End If

    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsBackupHandler::Class_Initialize()" & vbCrLf & "Error Description: " & Err.Description)
End Sub

' ###########################################################################################
' #
' #                                     PUBLIC METHODS
' #
' ###########################################################################################

'***************************************************************************
'Purpose: Method that saves a sheet to a CSV file
'Inputs:  Name of the sheet as a string
'Outputs: A trace in the log file and, if no error occured, a CSV file fed with the data of the sheet
'***************************************************************************
Public Sub ToCSV(sheetname As String)
    Dim WorkbookName As String
        
    On Error GoTo ErrHandler:

    Logging.logINFO "-logINFO-", "Processing the backup of " & sheetname
    
    WorkbookName = ActiveWorkbook.FullName
    
    Application.DisplayAlerts = False
    Sheets(sheetname).Select
    Cells.Select
    ActiveWorkbook.SaveAs filename:=sBackupPath & NewCSVFileName(sheetname), _
        FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    
    'Reset selection
    ActiveSheet.name = sheetname
    Sheets(sheetname).Cells.Select
    Sheets(sheetname).Range("A1").Select
    
    'Reset the file format to xlsm and enable backup of the file
    ActiveWorkbook.SaveAs filename:=WorkbookName, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=True
    
    'Reset display alerts
    Application.DisplayAlerts = True

    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: clsBackupHandler::ToCSV()" & vbCrLf & "Error Description: " & Err.Description)
End Sub

' ###########################################################################################
' #
' #                                     PRIVATE METHODS
' #
' ###########################################################################################


'***************************************************************************
'Purpose: Method to update the index of the CSV file
'Inputs:  The index of the CSV file read from the cell named CSVBackupNumber
'Outputs: The updated index
'***************************************************************************
Private Sub UpdateIndex()
    If iCSVIndex = Int(Settings.Range("CSVBackupNumber").Value) - 1 Then
        iCSVIndex = 0
    Else
        iCSVIndex = iCSVIndex + 1
    End If
    Settings.Range("CurrentExportNumber").Value = str(iCSVIndex)
End Sub

'***************************************************************************
'Purpose: Provide the name of the CSV file
'Inputs:  The name of the sheet that will be saved
'Outputs: The name of the CSV file that will be saved as a backup of the sheet
'***************************************************************************
Private Function NewCSVFileName(sheetname As String) As String
    NewCSVFileName = "Export" & sheetname & "." & iCSVIndex & ".csv"
End Function

