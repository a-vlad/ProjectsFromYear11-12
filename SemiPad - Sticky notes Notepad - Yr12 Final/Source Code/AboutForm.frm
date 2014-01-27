VERSION 5.00
Begin VB.Form AboutForm 
   BackColor       =   &H00D8EAED&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Semipad"
   ClientHeight    =   4830
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6195
   ClipControls    =   0   'False
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "AboutForm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3333.752
   ScaleMode       =   0  'User
   ScaleWidth      =   5817.425
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox NotepadPic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   405
      Picture         =   "AboutForm.frx":000C
      ScaleHeight     =   570
      ScaleWidth      =   510
      TabIndex        =   10
      Top             =   1335
      Width           =   510
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00CCAE93&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H00CCAE93&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   0
      Picture         =   "AboutForm.frx":0FBE
      ScaleHeight     =   811.195
      ScaleMode       =   0  'User
      ScaleWidth      =   4350.955
      TabIndex        =   1
      Top             =   0
      Width           =   6195
      Begin VB.Label lblMainTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Semipad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   1800
         TabIndex        =   11
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00DCE6EB&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4800
      TabIndex        =   0
      Top             =   4320
      Width           =   1260
   End
   Begin VB.Label lblMemory 
      BackStyle       =   0  'Transparent
      Caption         =   "Windows"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label lblLicenceAgreement 
      BackStyle       =   0  'Transparent
      Caption         =   "End-User License Agreement"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A5030F&
      Height          =   240
      Left            =   1736
      TabIndex        =   9
      Top             =   2760
      Width           =   2115
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 1999-2006 V-Unit Software"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1440
      TabIndex        =   8
      Top             =   1920
      Width           =   3795
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "V-Unit ®  Semipad"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1440
      TabIndex        =   7
      Top             =   1440
      Width           =   3795
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "V-Unit Software"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1680
      TabIndex        =   6
      Top             =   3120
      Width           =   3795
   End
   Begin VB.Label lblTerms2 
      BackStyle       =   0  'Transparent
      Caption         =   "the                                                 to:"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1440
      TabIndex        =   5
      Top             =   2760
      Width           =   3795
   End
   Begin VB.Label lblTerms 
      BackStyle       =   0  'Transparent
      Caption         =   "This product is licenced under the terms of"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1440
      TabIndex        =   4
      Top             =   2520
      Width           =   3795
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.1.6 Beta (Build 841.semipad.2399-49)"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Width           =   3795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   1352.234
      X2              =   5676.567
      Y1              =   2484.784
      Y2              =   2484.784
   End
   Begin VB.Label lblOperatingSystem 
      BackStyle       =   0  'Transparent
      Caption         =   "Semipad is currently running on:"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1440
      TabIndex        =   2
      Top             =   3720
      Width           =   2355
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   1352.234
      X2              =   5662.481
      Y1              =   2484.784
      Y2              =   2484.784
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me 'hides the form
End Sub

Private Sub Form_Load()
Dim UserName As String

UserName = String(100, Chr$(0)) ' Fills up a string with empty spaces
GetUserName UserName, 100 ' the string is then used to store the username
lblUser.Caption = UserName 'the username is displayed

lblMemory.Caption = "Microsoft Windows " & GetWinVersion 'the windows version code is also done here

'restores the default mouse icon
lblLicenceAgreement.MouseIcon = LoadPicture(App.Path & "\Graphics\hand.cur")
lblLicenceAgreement.MousePointer = 99
End Sub

Public Function GetWinVersion() As String
Dim Ver As Long, WinVer As Long

'this function is used to retrieve the windows version
Ver = GetVersion() 'first the version number is retrived

'The version number is processed though the following algorithem which produces a major minor version
WinVer = Ver And &HFFFF&
GetWinVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed")
End Function

Private Sub Form_Unload(Cancel As Integer)
mainform.Show 'shows the mainform window
mainform.Enabled = True 'reenables the mainform window
End Sub

