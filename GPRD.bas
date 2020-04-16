Attribute VB_Name = "GPRD"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' MODULE_CODE:   GPRD
'''                 - Any function or sub called for the processing of data
'''             in compliance with GPRD rules statements
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 31/05/2019      David FOUCHE            Created
'''

Public defaultCrypto As clsCryptoEngine
Public defaultHash As clsHashMethod

'***************************************************************************
'Purpose: To create a static class in charge of handling the enciphering and deciphering operations.
'Inputs : None
'Outputs: A single instance of the class clsCryptoEngine
'***************************************************************************
Public Sub thiscryptoinit()
  Static staticCrypto As New clsCryptoEngine 'singelton
  On Error GoTo ErrHandler:
    'Read the active key from the key list
    Set defaultCrypto = staticCrypto  'refence to static object
    Logging.logINFO ("thiscryptoinit::defaultCrypto static object is initialized")
  Exit Sub
ErrHandler:
    Logging.logFATAL ("thiscryptoinit::Initialisation of defaultCrypto has failed")
End Sub

'***************************************************************************
'Purpose: To create a static class in charge of handling the hash of the data
'Inputs : None
'Outputs: A single instance of the class clsHashMethod
'***************************************************************************
Public Sub thishashinit()
  Static statichash As New clsHashMethod 'singelton
  Set defaultHash = statichash  'refence to static object
End Sub

'***************************************************************************
'Purpose: To clear the data of a sheetname without clearing the header or the format of the cells.
'Inputs : The name of the sheet where the clear must be performed.
'Outputs: None
'***************************************************************************
Sub ClearDataOnly(sheetname As String)
    'Select any cell with input data in it, without selecting the header or any cell with formula inside

    With Sheets(sheetname).usedrange
        .Cells(2, 1).Resize(.Rows.Count - 1, .Columns.Count).SpecialCells(xlCellTypeConstants).Select
    End With
    selection.ClearContents

End Sub

'***************************************************************************
'Purpose: The event called when the button "Clear all data" is pressed
'Inputs : The name of the sheet where the data clear must be performed.
'Outputs: None
'***************************************************************************
Sub EventClearData(sheetname As String)
    If MsgBox("Etes-vous sûr de vouloir effacer les données de " & sheetname & " ?", vbYesNo, "Effacer les données") = vbYes Then
        ClearDataOnly (sheetname)
        MsgBox "Les données de la feuille " & sheetname & " ont été effacées !"
    End If

End Sub
