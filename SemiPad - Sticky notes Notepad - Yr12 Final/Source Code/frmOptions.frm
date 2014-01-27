VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customise Note"
   ClientHeight    =   7845
   ClientLeft      =   4830
   ClientTop       =   4200
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Tabs 
      Height          =   315
      Left            =   345
      ScaleHeight     =   255
      ScaleWidth      =   3930
      TabIndex        =   31
      Top             =   540
      Width           =   3990
   End
   Begin VB.PictureBox imgSep 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   -165
      Picture         =   "frmOptions.frx":4D5A
      ScaleHeight     =   75
      ScaleWidth      =   6195
      TabIndex        =   32
      Top             =   375
      Width           =   6195
   End
   Begin VB.PictureBox imgHeader 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   -15
      Picture         =   "frmOptions.frx":65D4
      ScaleHeight     =   420
      ScaleWidth      =   6060
      TabIndex        =   29
      Top             =   -30
      Width           =   6060
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Post Note"
      Height          =   345
      Left            =   4845
      MaskColor       =   &H80000016&
      TabIndex        =   2
      ToolTipText     =   "A note with these settings will be poseted."
      Top             =   7365
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3780
      MaskColor       =   &H80000016&
      TabIndex        =   1
      Top             =   7365
      Width           =   1020
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   2685
      MaskColor       =   &H80000016&
      TabIndex        =   0
      Top             =   7365
      Width           =   1020
   End
   Begin VB.PictureBox ReminderTab 
      BorderStyle     =   0  'None
      Height          =   2160
      Left            =   90
      Picture         =   "frmOptions.frx":1DB0E
      ScaleHeight     =   2160
      ScaleWidth      =   5850
      TabIndex        =   16
      Top             =   3000
      Width           =   5850
      Begin VB.CheckBox chkReminder 
         BackColor       =   &H00F2F9F9&
         Caption         =   "Set a Reminder"
         Height          =   255
         Left            =   3090
         TabIndex        =   28
         ToolTipText     =   "An alarm will be set on the note for the specific time and date."
         Top             =   585
         Width           =   2355
      End
      Begin VB.Timer TimeSynch 
         Interval        =   1000
         Left            =   735
         Top             =   1470
      End
      Begin VB.PictureBox picAlarm 
         AutoSize        =   -1  'True
         BackColor       =   &H00DCE6EB&
         BorderStyle     =   0  'None
         Height          =   750
         Left            =   210
         Picture         =   "frmOptions.frx":45A38
         ScaleHeight     =   750
         ScaleWidth      =   810
         TabIndex        =   24
         Top             =   465
         Width           =   810
      End
      Begin VB.ComboBox part 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOptions.frx":47A82
         Left            =   4620
         List            =   "frmOptions.frx":47A8C
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1380
         Width           =   720
      End
      Begin VB.ComboBox min 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOptions.frx":47A98
         Left            =   3885
         List            =   "frmOptions.frx":47AC0
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1380
         Width           =   645
      End
      Begin VB.ComboBox hour 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOptions.frx":47AF4
         Left            =   3090
         List            =   "frmOptions.frx":47B1C
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1380
         Width           =   615
      End
      Begin VB.PictureBox remDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3090
         ScaleHeight     =   225
         ScaleWidth      =   2190
         TabIndex        =   17
         Top             =   990
         Width           =   2250
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D76A28&
         Height          =   285
         Left            =   2940
         TabIndex        =   26
         Top             =   195
         Width           =   2595
      End
      Begin VB.Label lblNow 
         BackStyle       =   0  'Transparent
         Caption         =   "Now"
         Height          =   270
         Left            =   1575
         TabIndex        =   25
         Top             =   225
         Width           =   1470
      End
      Begin VB.Label lblSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3750
         TabIndex        =   20
         Top             =   1380
         Width           =   120
      End
      Begin VB.Label lblAt 
         BackStyle       =   0  'Transparent
         Caption         =   "At"
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   1560
         TabIndex        =   19
         Top             =   1410
         Width           =   1500
      End
      Begin VB.Label lblReminde 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Reminde me on the"
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   1545
         TabIndex        =   18
         Top             =   1035
         Width           =   2310
      End
   End
   Begin VB.PictureBox BackgroundTab 
      BorderStyle     =   0  'None
      Height          =   2190
      Left            =   105
      Picture         =   "frmOptions.frx":47B47
      ScaleHeight     =   2190
      ScaleWidth      =   5835
      TabIndex        =   9
      Top             =   5130
      Width           =   5835
      Begin VB.PictureBox DragCorner 
         BackColor       =   &H00FFFDF9&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   1905
         ScaleHeight     =   285
         ScaleWidth      =   300
         TabIndex        =   50
         Top             =   1545
         Width           =   300
         Begin VB.Line DragLine 
            BorderColor     =   &H00D66C3F&
            Index           =   0
            X1              =   300
            X2              =   150
            Y1              =   225
            Y2              =   375
         End
         Begin VB.Line DragLine 
            BorderColor     =   &H00D66C3F&
            Index           =   1
            X1              =   330
            X2              =   105
            Y1              =   150
            Y2              =   375
         End
         Begin VB.Line DragLine 
            BorderColor     =   &H00D66C3F&
            Index           =   2
            X1              =   420
            X2              =   75
            Y1              =   15
            Y2              =   360
         End
      End
      Begin VB.PictureBox imgHolder 
         BackColor       =   &H00F3F7F8&
         BorderStyle     =   0  'None
         Height          =   1005
         Left            =   3000
         ScaleHeight     =   1005
         ScaleWidth      =   2040
         TabIndex        =   30
         Top             =   870
         Width           =   2040
         Begin VB.CheckBox chkRandCol 
            BackColor       =   &H00F3F7F8&
            Caption         =   "Random Frame Colour"
            Height          =   210
            Left            =   120
            TabIndex        =   49
            Top             =   735
            Width           =   1920
         End
         Begin VB.Label lblFrameCol 
            BackStyle       =   0  'Transparent
            Caption         =   "Note Frame Colour"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A5030F&
            Height          =   240
            Left            =   120
            MouseIcon       =   "frmOptions.frx":6FA71
            MousePointer    =   99  'Custom
            TabIndex        =   48
            ToolTipText     =   "Change the note border and header colour."
            Top             =   465
            Width           =   1335
         End
         Begin VB.Label lblBkgCol 
            BackStyle       =   0  'Transparent
            Caption         =   "Note Background Colour"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A5030F&
            Height          =   240
            Left            =   120
            MouseIcon       =   "frmOptions.frx":6FD7B
            MousePointer    =   99  'Custom
            TabIndex        =   47
            ToolTipText     =   "Change the background colour of the note content behind the text."
            Top             =   30
            Width           =   1815
         End
      End
      Begin VB.PictureBox NoteHeader 
         BackColor       =   &H00D66C3F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   180
         ScaleHeight     =   210
         ScaleWidth      =   2040
         TabIndex        =   35
         Top             =   225
         Width           =   2040
         Begin VB.Label lblExtFun 
            BackStyle       =   0  'Transparent
            Caption         =   "¤"
            BeginProperty Font 
               Name            =   "Wingdings 3"
               Size            =   8.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   390
            Index           =   0
            Left            =   1455
            TabIndex        =   39
            Top             =   15
            Width           =   225
         End
         Begin VB.Label lblCustomise 
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Wingdings 2"
               Size            =   8.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Index           =   0
            Left            =   1620
            TabIndex        =   38
            Top             =   30
            Width           =   300
         End
         Begin VB.Label lblClose 
            BackStyle       =   0  'Transparent
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Index           =   0
            Left            =   1860
            TabIndex        =   37
            Top             =   -45
            Width           =   195
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Sticky Note"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   36
            Top             =   0
            Width           =   1080
         End
      End
      Begin VB.CheckBox chkLockEdit 
         BackColor       =   &H00F9FBFB&
         Caption         =   "Lock from Edit"
         Height          =   255
         Left            =   3495
         TabIndex        =   34
         ToolTipText     =   "When locked a note's content cannot be edited."
         Top             =   240
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.PictureBox Dialog 
         Height          =   480
         Left            =   4830
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   52
         Top             =   -450
         Width           =   1200
      End
      Begin VB.Frame frmBkgOptions 
         BackColor       =   &H00F3F7F8&
         Caption         =   "Background Options"
         ForeColor       =   &H8000000D&
         Height          =   1335
         Left            =   2865
         TabIndex        =   10
         Top             =   630
         Width           =   2340
      End
      Begin VB.PictureBox sample 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFDF9&
         BorderStyle     =   0  'None
         Height          =   1395
         Left            =   195
         ScaleHeight     =   1395
         ScaleWidth      =   2010
         TabIndex        =   40
         Top             =   435
         Width           =   2010
      End
      Begin VB.Shape Border 
         BorderColor     =   &H00D66C3F&
         Height          =   1425
         Index           =   0
         Left            =   180
         Top             =   420
         Width           =   2040
      End
      Begin VB.Image imgLock 
         Height          =   615
         Left            =   2880
         Picture         =   "frmOptions.frx":70085
         Top             =   15
         Width           =   555
      End
   End
   Begin VB.PictureBox EffectsTab 
      BorderStyle     =   0  'None
      Height          =   2160
      Left            =   120
      Picture         =   "frmOptions.frx":712B7
      ScaleHeight     =   2160
      ScaleWidth      =   5835
      TabIndex        =   11
      Top             =   840
      Width           =   5835
      Begin VB.PictureBox DragCorner 
         BackColor       =   &H00FFFDF9&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   1905
         ScaleHeight     =   285
         ScaleWidth      =   300
         TabIndex        =   51
         Top             =   1545
         Width           =   300
         Begin VB.Line DragLine 
            BorderColor     =   &H00D66C3F&
            Index           =   5
            X1              =   435
            X2              =   90
            Y1              =   0
            Y2              =   345
         End
         Begin VB.Line DragLine 
            BorderColor     =   &H00D66C3F&
            Index           =   4
            X1              =   330
            X2              =   105
            Y1              =   150
            Y2              =   375
         End
         Begin VB.Line DragLine 
            BorderColor     =   &H00D66C3F&
            Index           =   3
            X1              =   315
            X2              =   165
            Y1              =   210
            Y2              =   360
         End
      End
      Begin VB.PictureBox sample2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFDF9&
         BorderStyle     =   0  'None
         Height          =   1395
         Left            =   195
         ScaleHeight     =   1395
         ScaleWidth      =   2010
         TabIndex        =   46
         Top             =   435
         Width           =   2010
      End
      Begin VB.PictureBox NoteHeader 
         BackColor       =   &H00D66C3F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   180
         ScaleHeight     =   210
         ScaleWidth      =   2040
         TabIndex        =   41
         Top             =   225
         Width           =   2040
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Sticky Note"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   1
            Left            =   90
            TabIndex        =   45
            Top             =   0
            Width           =   1080
         End
         Begin VB.Label lblClose 
            BackStyle       =   0  'Transparent
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Index           =   1
            Left            =   1860
            TabIndex        =   44
            Top             =   -45
            Width           =   195
         End
         Begin VB.Label lblCustomise 
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Wingdings 2"
               Size            =   8.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Index           =   1
            Left            =   1620
            TabIndex        =   43
            Top             =   30
            Width           =   300
         End
         Begin VB.Label lblExtFun 
            BackStyle       =   0  'Transparent
            Caption         =   "¤"
            BeginProperty Font 
               Name            =   "Wingdings 3"
               Size            =   8.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   390
            Index           =   1
            Left            =   1455
            TabIndex        =   42
            Top             =   15
            Width           =   225
         End
      End
      Begin VB.CheckBox chkAOT 
         BackColor       =   &H00F0F8F9&
         Caption         =   "Always on Top"
         Height          =   255
         Left            =   2910
         TabIndex        =   27
         Top             =   1620
         Width           =   1515
      End
      Begin VB.CheckBox chkInvisible 
         BackColor       =   &H00F0F8F9&
         Caption         =   "Make background invisible"
         Height          =   255
         Left            =   2910
         TabIndex        =   13
         ToolTipText     =   "The note background is see-through."
         Top             =   1275
         Width           =   2295
      End
      Begin VB.Frame frmTransparancy 
         BackColor       =   &H00F9FAFB&
         Caption         =   "Transparancy"
         Height          =   1095
         Left            =   2745
         TabIndex        =   12
         Top             =   90
         Width           =   2535
         Begin VB.CheckBox chkTransparent 
            BackColor       =   &H00F8FAFC&
            Caption         =   "Use Transparancy"
            Height          =   255
            Left            =   165
            TabIndex        =   14
            Top             =   300
            Width           =   1695
         End
         Begin VB.PictureBox TransparancySlider 
            Height          =   330
            Left            =   135
            ScaleHeight     =   270
            ScaleWidth      =   2145
            TabIndex        =   33
            Top             =   615
            Width           =   2205
         End
         Begin VB.Label lblTransparancy 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "100%"
            Height          =   210
            Left            =   2040
            TabIndex        =   15
            Top             =   330
            Width           =   420
         End
      End
      Begin VB.Shape Border 
         BorderColor     =   &H00D66C3F&
         Height          =   1425
         Index           =   1
         Left            =   180
         Top             =   420
         Width           =   2040
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BackgroundTab_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tabs.RefreshTabs
End Sub

Public Sub chkInvisible_Click()

If chkInvisible.Value = 1 Then
    'code used to disabple unavailable options
    chkTransparent.Enabled = False
    TransparancySlider.DisableSlide
    frmTransparancy.Enabled = False
    lblTransparancy.ForeColor = &H8000000C
    lblBkgCol.Enabled = False
    
    'Code used to change the sample note
    CapWindow Me
    
    'dissables the template of the note on the left hand side
    sample.BackColor = RGB(255, 255, 254)
    sample2.BackColor = RGB(255, 255, 254)
    DragCorner(0).BackColor = RGB(255, 255, 254)
    DragCorner(1).BackColor = RGB(255, 255, 254)
    Render.WindowShape Me, RGB(255, 255, 254)
Else
    'Code used to disable unavailable options
    chkTransparent.Enabled = True
    frmTransparancy.Enabled = True
    lblTransparancy.ForeColor = vbBlack
    lblBkgCol.Enabled = True
    
    'Code to change sample
    ReleaseWindow Me.hwnd
    sample2.BackColor = sample.BackColor
    sample.Refresh: sample2.Refresh
    AssimDragIco
End If
End Sub

Private Sub chkRandCol_Click()
If chkRandCol.Value = vbChecked Then lblFrameCol.Enabled = False Else lblFrameCol.Enabled = True
End Sub

Private Sub chkReminder_Click()
If chkReminder.Value = 0 Then
    'disables the options for setting a new alarm
    remDate.Enabled = False
    hour.Enabled = False
    min.Enabled = False
    part.Enabled = False
    lblReminde.ForeColor = &H808080
    lblAt.ForeColor = &H808080
Else
    'enables the options for the user to select
    remDate.Enabled = True
    hour.Enabled = True
    min.Enabled = True
    part.Enabled = True
    lblReminde.ForeColor = vbBlack
    lblAt.ForeColor = vbBlack
    'sets the current time and date , rounds off the time
    hour.Text = Left(Time, 2)
    min.Text = (Left(Right(Time, 8), 1) & "0")
    part.Text = Right(Time, 2)
    remDate.Year = Right(Date, 4)
    remDate.Month = Left(Right(Date, 7), 2)
    If Right(Left(Date, 2), 1) = "/" Then remDate.Day = Left(Date, 1) Else remDate.Day = Left(Date, 2)
End If
End Sub

Private Sub chkTransparent_Click()
If chkTransparent.Value = 1 Then
    'enables the slider and disables invisibility as they cannot work at the same time!
    chkInvisible.Enabled = False
    TransparancySlider.EnableSlide
    TransparancySlider_Change
Else
    'dissables the dlider and enables all options inc inv. bak
    chkInvisible.Enabled = True
    TransparancySlider.DisableSlide
    Render.ReleaseWindow Me.hwnd
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me                    'closes the window
cmdCancel.Tag = "1"          'sets a tag noting that the option has been closed used to detect the apply settings function in the note form
mainform.Enabled = True      'returns back to the mainform
If cmdPost.Enabled = True Then mainform.SetFocus
End Sub


Private Sub cmdOK_Click()
Me.Hide                   'hides but does not close the customise dialog.
mainform.Enabled = True
cmdCancel.Tag = ""
If cmdPost.Enabled = True Then mainform.SetFocus
End Sub

Private Sub cmdPost_Click()
Me.Hide
mainform.Enabled = True
mainform.NewNote            'calls the new note function routine
End Sub

Private Sub EffectsTab_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tabs.RefreshTabs
TransparancySlider.RefreshPick
End Sub

Private Sub Form_Load()
'Loads the basic settings into the dialog
Render.AlwaysOnTop Me, True 'sets as aot to overrule aot notes for screen view
min.Text = "00"
hour.Text = "12"
part.Text = "AM"
Tabs.LoadTabs
Tabs_ClickTab1
TransparancySlider.DisableSlide
End Sub

Public Function LoadValuesFromFile(NoteArrayID As Integer)
'this function synchronises options in the customise dialog box.
'applys the alarm settings from the note to all the options of the dialog box.
If Alarm(NoteArrayID).Set = True Then
    chkReminder.Value = 1
    min.Text = Alarm(NoteArrayID).min
    hour.Text = Alarm(NoteArrayID).hour
    part.Text = Alarm(NoteArrayID).part
    remDate.Day = Alarm(NoteArrayID).Day
    remDate.Month = Alarm(NoteArrayID).Month
    remDate.Year = Alarm(NoteArrayID).Year
Else
    chkReminder.Value = 0
End If

'Appplys the settings of the note to the options of the dialog box.
If NoteSetting(NoteArrayID).nsLockEdit = True Then chkLockEdit.Value = 1 Else chkLockEdit.Value = 0
If NoteSetting(NoteArrayID).nsAlwaysOnTop = True Then chkAOT.Value = 1 Else chkAOT.Value = 0
If NoteSetting(NoteArrayID).nsTransparancyEnabled = True Then chkTransparent.Value = 1: chkInvisible.Value = 0: TransparancySlider.SetNewValue (NoteSetting(NoteArrayID).nsTransparancyLevel): lblTransparancy.Caption = NoteSetting(NoteArrayID).nsTransparancyLevel & "%"
If NoteSetting(NoteArrayID).nsInvisibleBackground = True Then chkTransparent.Value = 0: chkInvisible.Value = 1
If NoteSetting(NoteArrayID).nsRandomBorderColour = True Then chkRandCol.Value = 1: lblFrameCol.Enabled = False
AssimDragIco
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tabs.RefreshTabs 'refreshes tabs
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdCancel.Tag = "1" 'sets tag to mean to apply the settings
Unload Me 'closes the window
mainform.Enabled = True 'enables main form
If cmdPost.Enabled = True Then mainform.SetFocus
End Sub

Private Sub frmTransparancy_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'Updates the transparancy slider status
TransparancySlider.RefreshPick
End Sub

Private Sub lblBkgCol_Click()
On Error Resume Next
Dialog.Color = sample.BackColor     'updates dialog selected colour to current colour
Dialog.ShowColor                    'changes the colour.
sample.BackColor = Dialog.Color     'updates the change.
AssimDragIco
End Sub

Private Sub lblBkgCol_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'updates back colours on form
lblBkgCol.ForeColor = vbBlue
lblFrameCol.ForeColor = &HA5030F
End Sub

Private Sub lblFrameCol_Click()
On Error Resume Next
Dialog.Color = NoteHeader(0).BackColor  'updates dialog selected colour to current colour
Dialog.ShowColor                        'changes the colour.
'updates the change.
Border(0).BorderColor = Dialog.Color
Border(1).BorderColor = Dialog.Color
'updates back colours on form
NoteHeader(0).BackColor = Dialog.Color
NoteHeader(1).BackColor = Dialog.Color
AssimDragIco
End Sub

Private Sub lblFrameCol_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'updates back colours on form
lblBkgCol.ForeColor = &HA5030F
lblFrameCol.ForeColor = vbBlue
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tabs.RefreshTabs 'refreshed the tabs
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'updates back colours on form
lblBkgCol.ForeColor = &HA5030F
lblFrameCol.ForeColor = &HA5030F
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tabs.RefreshTabs 'refreshed the tabs
End Sub

Private Sub ReminderTab_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tabs.RefreshTabs 'refreshed the tabs
End Sub

Private Sub Tabs_ClickTab1()
'makes all tabs invisible excent for the select tab. this is used to swap the tabs
Me.BackgroundTab.Visible = True
Me.EffectsTab.Visible = False
Me.ReminderTab.Visible = False
ReleaseWindow Me.hwnd
        
Me.sample.Refresh: Me.sample2.Refresh
        
If Me.chkInvisible.Value = 1 Then Me.chkInvisible_Click
End Sub

Private Sub Tabs_ClickTab2()
'this function tabs and switches the tab content,
'it makes only the corresponding tag visible
Me.BackgroundTab.Visible = False
Me.EffectsTab.Visible = True
Me.ReminderTab.Visible = False
        
Me.sample2.BackColor = Me.sample.BackColor
Me.sample.Refresh: Me.sample2.Refresh
                
Me.chkInvisible_Click
End Sub

Private Sub Tabs_ClickTab3()
'this function tabs and switches the tab content,
'it makes only the corresponding tag visible
Render.ReleaseWindow Me.hwnd
Me.BackgroundTab.Visible = False
Me.EffectsTab.Visible = False
Me.ReminderTab.Visible = True
       
Me.lblTime.Caption = Time & " " & Date
Me.TimeSynch.Enabled = True
End Sub

Private Sub TimeSynch_Timer()
'updates the time date in the appointments tab in real time every second
lblTime.Caption = Time & " " & Date
End Sub

Private Sub TransparancySlider_Change()
'exactly the same as the function bellow
lblTransparancy.Caption = TransparancySlider.Value & "%"
If TransparancySlider.Value = 100 Or chkTransparent.Value = 0 Then Exit Sub
Render.TranslucenthWnd Me.hwnd, 255 * (TransparancySlider.Value / 100)
End Sub

Private Sub TransparancySlider_Scroll()
'this function sets the transparacy slider value and percentage next to the slider
lblTransparancy.Caption = TransparancySlider.Value & "%"
If TransparancySlider.Value = 100 Or chkTransparent.Value = 0 Then Exit Sub
'creates semitrasparancy in real time
Render.TranslucenthWnd Me.hwnd, 255 * (TransparancySlider.Value / 100)
End Sub

Public Function AssimDragIco()
'this function assimilates the 3 line drag corner colours in the note template
'and matches the correct colour theme in the template example.
DragCorner(0).BackColor = sample.BackColor ' sets the colour of the back for both examples
DragCorner(1).BackColor = sample.BackColor
For i = 0 To 5
    DragLine(i).BorderColor = NoteHeader(0).BackColor 'themes the line colours
Next i
End Function
