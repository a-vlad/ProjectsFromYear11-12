VERSION 5.00
Begin VB.Form FindDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1530
   ClientLeft      =   5340
   ClientTop       =   5250
   ClientWidth     =   5310
   Icon            =   "FindDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "March case"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1060
      Width           =   1935
   End
   Begin VB.TextBox FindText 
      Height          =   315
      Left            =   990
      TabIndex        =   3
      Top             =   165
      Width           =   2920
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4080
      TabIndex        =   1
      Top             =   550
      Width           =   1095
   End
   Begin VB.CommandButton FindButton 
      Caption         =   "Find Next"
      Enabled         =   0   'False
      Height          =   330
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label FindWhat 
      Caption         =   "Find what:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   210
      Width           =   975
   End
End
Attribute VB_Name = "FindDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()
FindDialog.Hide 'hides the box
End Sub

Private Sub FindButton_Click()
Dim result As Integer
mainform.SetFocus
mainform.Text.SelStart = 0 'sets the start position back to 0

'this changes the search flags if the match case box is ticked
If chkMatchCase.Value = 1 Then
findFlags = 4
Else: findFlags = 0: End If
'performs the search
result = mainform.Text.Find(FindText.Text, mainform.Text.SelStart, Len(mainform.Text.Text), findFlags)

'if nothing is found the result is -1 and a dialog box is displayed
If result = "-1" Then MsgBox "Cannot find: " & FindText & "", vbInformation, "Semipad": Exit Sub

Me.Hide 'the form disapears
End Sub

Private Sub FindText_Change()
If FindText.Text = "" Then 'enables the buttons based on the text input
    FindButton.Enabled = False
Else
    FindButton.Enabled = True
End If
End Sub

Private Sub Form_Load()
mainform.Text.SelStart = 0 'sets the current text edit position to the start
If FindText.Text = "" Then 'enables the buttons based on the text input
    FindButton.Enabled = False
Else
    FindButton.Enabled = True
End If
End Sub
