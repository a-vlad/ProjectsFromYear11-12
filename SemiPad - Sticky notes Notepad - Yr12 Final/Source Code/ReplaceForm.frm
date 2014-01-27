VERSION 5.00
Begin VB.Form ReplaceForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Replace"
   ClientHeight    =   1920
   ClientLeft      =   5445
   ClientTop       =   4725
   ClientWidth     =   5160
   ClipControls    =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   Icon            =   "ReplaceForm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1325.218
   ScaleMode       =   0  'User
   ScaleWidth      =   4845.507
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "March case"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtWith 
      Height          =   310
      Left            =   1200
      TabIndex        =   7
      Top             =   600
      Width           =   2490
   End
   Begin VB.TextBox txtFind 
      Height          =   310
      Left            =   1200
      TabIndex        =   0
      Top             =   189
      Width           =   2490
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3840
      TabIndex        =   4
      Top             =   1335
      Width           =   1140
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace All"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3840
      TabIndex        =   3
      Top             =   932
      Width           =   1140
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3840
      TabIndex        =   2
      Top             =   522
      Width           =   1140
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label lblReplace 
      Caption         =   "Replace with:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   667
      Width           =   1095
   End
   Begin VB.Label lblFind 
      Caption         =   "Find what:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "ReplaceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Me.Hide 'hides the form
End Sub

Private Sub cmdReplace_Click()
Dim result As Integer

mainform.SetFocus 'sets focus to the mainform text entry form
mainform.Text.SelStart = 0 'sets the text input to the front of the file

'this changes the search flags if the match case box is ticked
If chkMatchCase.Value = 1 Then
findFlags = 4
Else: findFlags = 0: End If

result = mainform.Text.Find(txtFind.Text, mainform.Text.SelStart, Len(mainform.Text.Text), findFlags) 'does the search
mainform.Text.SelText = txtWith.Text 'sends the replace text which is replaced over the highlighted text

'shows a message if there was no text replaced or found in the document
If result = "-1" Then MsgBox "Cannot find: " & txtFind & "", vbInformation, "Semipad"
End Sub

Private Sub cmdReplaceAll_Click()
On Error Resume Next
Dim result As Integer
Dim Replacements As Integer

mainform.SetFocus 'sets focus to the mainform text entry window
mainform.Text.SelStart = 0 'sets the text entry position back to zero
Replacements = 0 'resets the number of preplaced words to 0
findFlags = 0

'this changes the search flags if the match case box is ticked
If chkMatchCase.Value = 1 Then
findFlags = 4
Else: findFlags = 0: End If

'this is a loop which replaces words methodically based on the replace function
Do
    mainform.SetFocus  'sets focus to the mainform form each time it loops
    result = mainform.Text.Find(txtFind.Text, mainform.Text.SelStart, Len(mainform.Text.Text), findFlags) 'finds the text to replace
    If result <> "-1" Then mainform.Text.SelText = txtWith.Text 'if text is found it replaces it
    
    mainform.Text.SelStart = mainform.Text.SelStart + Len(txtWith) - 1 'it moves the current start replace mosition up by the length of the replaced word to avoid a loop
    Replacements = Replacements + 1 'increments the words replaced counter
    
Loop Until result = "-1" ' post test loop allows for the counter to count the last replacement

Replacements = Replacements - 1 ' this accounts for a misscount which occurs when no words are found yet the counter still counts 1 replacement

'this code displayes the according message to the user based on the replacement results
If Replacements = 0 Then 'when no replacements made the user is told so
        MsgBox "Cannot find: " & txtFind & "", vbInformation, "Semipad": Exit Sub
ElseIf Replacements > 0 Then 'else the user is told the number of replacements made
        MsgBox Replacements & " replacements made.", vbInformation, "Replace"
        Exit Sub
End If
End Sub

Private Sub txtFind_Change()
'this code enables the replace and find buttons if text is present in the search field
If txtFind.Text = "" Then
    cmdFind.Enabled = False
    cmdReplace.Enabled = False
    cmdReplaceAll.Enabled = False
Else
    cmdFind.Enabled = True
    cmdReplace.Enabled = True
    cmdReplaceAll.Enabled = True
End If
End Sub

Private Sub cmdFind_Click()
Dim result As Integer

mainform.SetFocus 'sets focus to the mainform text form
mainform.Text.SelStart = 0 'sets the current start text position back to start

'this changes the search flags if the match case box is ticked
If chkMatchCase.Value = 1 Then
findFlags = 4
Else: findFlags = 0: End If

result = mainform.Text.Find(txtFind.Text, mainform.Text.SelStart, Len(mainform.Text.Text), findFlags) 'performs the search

'if no texx is found a msg box is shown
If result = "-1" Then MsgBox "Cannot find: " & txtFind & "", vbInformation, "Semipad"
End Sub

Private Sub Form_Load()
'this code enables the replace and find buttons if text is present in the search field
If txtFind.Text = "" Then
    cmdFind.Enabled = False
    cmdReplace.Enabled = False
    cmdReplaceAll.Enabled = False
Else
    cmdFind.Enabled = True
    cmdReplace.Enabled = True
    cmdReplaceAll.Enabled = True
End If
End Sub
