Attribute VB_Name = "TestBackup"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' TEST_CODE:   TestBackup
'''                 - Test backup and recovery
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 31/05/2019      David FOUCHE            Created
'''

'***************************************************************************
'Purpose: Test the backup of some sheets
'Inputs:  None
'Outputs: the CSV folder must be added with the CSV files of the sheets Bookings, Guests and Keys
'***************************************************************************
Sub TestBackup()
    thisbackup (Bookings.name)
    thisbackup (Guests.name)
    thisbackup (KeyList.name)
End Sub

'***************************************************************************
'Purpose: Test the call of the method used to create CSV files
'Inputs:  None
'Outputs: the CSV folder must contain the CSV files with the expected number and name
'***************************************************************************
Sub TestNewCSVFileName()
    Settings.Range("CurrentExportNumber").Value = ""
    NewCSVFileName ("TestBackup")
    Settings.Range("CurrentExportNumber").Value = "2"
    NewCSVFileName ("TestBackup")
    Settings.Range("CurrentExportNumber").Value = Settings.Range("CSVBackupNumber").Value
    NewCSVFileName ("TestBackup")
End Sub
