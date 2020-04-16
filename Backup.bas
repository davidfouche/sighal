Attribute VB_Name = "Backup"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' MODULE_CODE:   Backup
'''                 - Any function or sub called for the processing of backup
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 31/05/2019      David FOUCHE            Created
'''

Public defaultBackup As clsBackupHandler

'***************************************************************************
'Purpose: Create a static class to handle the backup of the spreadsheet
'Inputs:  Sheet name that will be used to name the CSV files
'Outputs: None
'***************************************************************************
Public Sub thisbackup(ByVal sheetname As String)
  Static staticBackup As New clsBackupHandler 'singelton
  staticBackup.ToCSV (sheetname)  'Save a CSV file of the spreadsheet
  Set defaultBackup = staticBackup  'refence to static object
End Sub
