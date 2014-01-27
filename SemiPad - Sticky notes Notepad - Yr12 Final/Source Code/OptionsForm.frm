VERSION 5.00
Begin VB.Form OptionsForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2625
   ClientLeft      =   6285
   ClientTop       =   4515
   ClientWidth     =   3570
   Icon            =   "OptionsForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Tabs 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   2475
      TabIndex        =   12
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2190
      TabIndex        =   1
      Top             =   2205
      Width           =   1290
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   330
      Left            =   870
      TabIndex        =   0
      Top             =   2205
      Width           =   1290
   End
   Begin VB.PictureBox imgHeaderLine 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   -75
      Picture         =   "OptionsForm.frx":000C
      ScaleHeight     =   75
      ScaleWidth      =   6195
      TabIndex        =   2
      Top             =   285
      Width           =   6195
   End
   Begin VB.PictureBox imgHeader 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   -330
      Picture         =   "OptionsForm.frx":1886
      ScaleHeight     =   660
      ScaleWidth      =   6825
      TabIndex        =   3
      Top             =   -345
      Width           =   6825
   End
   Begin VB.PictureBox TabContent1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   90
      Picture         =   "OptionsForm.frx":18DC0
      ScaleHeight     =   1455
      ScaleWidth      =   3600
      TabIndex        =   4
      Top             =   735
      Visible         =   0   'False
      Width           =   3600
      Begin VB.PictureBox tabicon2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   300
         Picture         =   "OptionsForm.frx":283D2
         ScaleHeight     =   585
         ScaleWidth      =   645
         TabIndex        =   11
         Top             =   375
         Width           =   645
      End
      Begin VB.CheckBox chkAutoLoad 
         BackColor       =   &H00F9FAFB&
         Caption         =   "Autoload with Windows"
         Height          =   240
         Left            =   1140
         TabIndex        =   6
         Top             =   330
         Width           =   1995
      End
      Begin VB.CheckBox chkReloadNotes 
         BackColor       =   &H00F2F5F7&
         Caption         =   "Reload Notes on Startup"
         Height          =   240
         Left            =   1140
         TabIndex        =   5
         Top             =   780
         Width           =   2115
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   3390
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         X1              =   3375
         X2              =   3375
         Y1              =   1335
         Y2              =   0
      End
   End
   Begin VB.PictureBox TabContent2 
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   90
      Picture         =   "OptionsForm.frx":29830
      ScaleHeight     =   1440
      ScaleWidth      =   3585
      TabIndex        =   7
      Top             =   735
      Width           =   3585
      Begin VB.PictureBox tabicon1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   150
         Picture         =   "OptionsForm.frx":38E42
         ScaleHeight     =   615
         ScaleWidth      =   870
         TabIndex        =   10
         Top             =   345
         Width           =   870
      End
      Begin VB.CheckBox chkAnimateTileEffect 
         BackColor       =   &H00F5FAFA&
         Caption         =   "Animate Tile Notes"
         Height          =   240
         Left            =   1185
         TabIndex        =   9
         Top             =   315
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CheckBox chkQuickMode 
         BackColor       =   &H00F2F5F7&
         Caption         =   "Quick-Post Mode (beta)"
         ForeColor       =   &H00A5030F&
         Height          =   240
         Left            =   1185
         TabIndex        =   8
         Top             =   795
         Width           =   2085
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         X1              =   3375
         X2              =   3375
         Y1              =   1335
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   3390
         Y1              =   1335
         Y2              =   1335
      End
   End
End
Attribute VB_Name = "OptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
'Closes the window and sets focus back to the main window
Me.Hide: mainform.Enabled = True: mainform.SetFocus
mainform.LoadOptionsOnly
End Sub

Private Sub cmdApply_Click()
'Applies the autoload setting if checked by callin the auto load/ auto unload function
If chkAutoLoad.Value = 1 Then AutoLoad Else AutoUnload
mainform.SaveEditorSettings
Me.Hide: mainform.Enabled = True: mainform.SetFocus

'This function applies the quickmode setting if enabled.
If Me.chkQuickMode.Value = 1 Then
    mainform.Hide: mainform.NewNote
    For i = 0 To UBound(NoteInstance)
        NoteInstance(i).cmdReturnToMain.Visible = True
    Next i
End If
End Sub

Private Sub Form_Load()
'loads the toolbar and changes the toolbar text on the top
Tabs.ChangeTabNames "General", "Startup Options", ""
Tabs.KillTab3 'disables the last toolbar header
Tabs.LoadTabs
Tabs_ClickTab1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'closes the program
Cancel = 1
cmdCancel_Click
End Sub

Private Sub Tabs_ClickTab1()
'This changes the picture content
TabContent2.Visible = True
TabContent1.Visible = False
End Sub

Private Sub Tabs_ClickTab2()
'This changes the picture content
TabContent2.Visible = False
TabContent1.Visible = True
End Sub
