VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewPassphraseForm 
   Caption         =   "Nouveau mot de passe"
   ClientHeight    =   3450
   ClientLeft      =   30
   ClientTop       =   370
   ClientWidth     =   6940
   OleObjectBlob   =   "NewPassphraseForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewPassphraseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' SHEET_CODE:   NewPassphraseForm
'''                 - Handle sheet events
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 31/05/2019      David FOUCHE            Created
'''

Const PASSPHRASEMINLENGTH = 12

''' https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/strings/walkthrough-validating-that-passwords-are-complex
''' <summary>Determines if a password is sufficiently complex.</summary>
''' <param name="pwd">Password to validate</param>
''' <param name="minLength">Minimum number of password characters.</param>
''' <param name="numUpper">Minimum number of uppercase characters.</param>
''' <param name="numLower">Minimum number of lowercase characters.</param>
''' <param name="numNumbers">Minimum number of numeric characters.</param>
''' <param name="numSpecial">Minimum number of special characters.</param>
''' <returns>True if the password is sufficiently complex.</returns>
Private Function ValidatePassword(ByVal psw As String) As Boolean
    Dim hasNum, hasUpper, hasLower As Boolean
    Dim i As Integer, j As Integer, k As Integer
    
    hasNum = False
    hasUpper = False
    hasLower = False
    
    For k = 48 To 57
        If (InStr(1, psw, Chr(k))) Then
            hasNum = True
            Exit For
        End If
    Next k

    'See if there is an upper case
    For i = 65 To 90
        If (InStr(1, psw, Chr(i))) Then
            hasUpper = True
            Exit For
        End If
    Next i

    'See if there is a lower case
    For j = 97 To 122
        If (InStr(1, psw, Chr(j))) Then
            hasLower = True
            Exit For
        End If
    Next j
    
    If Not hasLower Or Not hasUpper Or Not hasNum Or (Len(psw) < PASSPHRASEMINLENGTH) Then
        ValidatePassword = False
    Else
        ValidatePassword = True
    End If
    
End Function

Private Sub CurrentPassphraseTBox_Change()
    On Error GoTo ErrHandler:
    GPRD.defaultCrypto.GPRDPassphrase = CurrentPassphraseTBox.Value
    CurrentPassphraseTBox.BackColor = vbGreen
    NewPassphraseTBox.Enabled = True
    NewPassphraseTBox.BackColor = vbRed
    Exit Sub
ErrHandler:
    CurrentPassphraseTBox.BackColor = vbRed
    NewPassphraseValidButton.Enabled = False
    NewPassphraseTBox.Enabled = False
    NewPassphraseTBox.BackColor = &HE0E0E0
    NewPassphraseRepeatTBox.Enabled = False
    NewPassphraseRepeatTBox.BackColor = &HE0E0E0
    Logging.logWARN ("CurrentPassphraseTBox_Change::Password entered is invalid")
End Sub

Private Sub NewPassphraseRepeatTBox_Change()
    If NewPassphraseRepeatTBox.Value = NewPassphraseTBox.Value Then
        NewPassphraseValidButton.Enabled = True
        NewPassphraseRepeatTBox.BackColor = vbGreen
    Else
        NewPassphraseValidButton.Enabled = False
        NewPassphraseRepeatTBox.BackColor = vbRed
        Logging.logWARN ("NewPassphraseRepeatTBox_Change::New password entered is not correctly repeated")
    End If
End Sub

Private Sub NewPassphraseTBox_Change()
    If ValidatePassword(NewPassphraseTBox.Value) Then
        NewPassphraseTBox.BackColor = vbGreen
        NewPassphraseRepeatTBox.Enabled = True
        NewPassphraseRepeatTBox.BackColor = vbRed
    Else
        NewPassphraseTBox.BackColor = vbRed
        NewPassphraseRepeatTBox.Enabled = False
        NewPassphraseRepeatTBox.BackColor = &HE0E0E0
        MsgBox "Le nouveau mot de passe doit contenir au moins une minuscule, une majuscule, un chiffre et 12 caractères minimum.", vbOKOnly + vbInformation, "Mot de passe invalide"
        Logging.logWARN ("NewPassphraseTBox_Change::New password entered is not compliant with minimal requirements")
    End If
End Sub

Private Sub NewPassphraseValidButton_Click()
    Dim mapper As clsMapper
    
    On Error GoTo ErrHandler:
    
    'Add the new key into the key list
    Call KeyList.AddNewKey(NewPassphraseRepeatTBox.Value)
    'Launch the xcrypting of the sheet Guests
    Set mapper = New clsMapper
    
    'Notify the user
    MsgBox "Merci de patienter jusqu'au rechiffrement complet de l'onglet des résidents. Un message vous sera affiché à la fin de ce rechiffrement.", vbOKOnly + vbInformation, "Patientez svp..."
    
    'Apply the new passphrase to the Guests sheet
    Call mapper.Map(Guests, XCRYPTOPE)
    
    'Update the key sheet
    Call mapper.Map(KeyList, UPDATEKEY)
    
    'Notify the user
    MsgBox "Votre nouveau mot de passe a été appliqué. Merci de le conserver dans un carnet papier.", vbOKOnly + vbInformation, "Fin de traitement"
    Exit Sub
ErrHandler:
    Logging.logFATAL ("The following error has occured." & vbCrLf & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Source: NewPassphraseForm::NewPassphraseValidButton_Click()" & vbCrLf & "Error Description: " & Err.Description)
    'Notify the user
    MsgBox "Une erreur est survenue. Merci d'envoyer votre fichier et les logs à david.fouche@gmail.com", vbOKOnly + vbCritical, "Erreur changement de mot de passe"
End Sub

