VERSION 5.00
Begin VB.Form PostedNote 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   " Sticky Note"
   ClientHeight    =   2790
   ClientLeft      =   10710
   ClientTop       =   855
   ClientWidth     =   3000
   Icon            =   "PostedNote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox DragCorner 
      BackColor       =   &H00F9FAFB&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2670
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   9
      Top             =   2475
      Width           =   300
      Begin VB.Line DragLine 
         BorderColor     =   &H00D66C3F&
         BorderWidth     =   2
         Index           =   2
         X1              =   375
         X2              =   30
         Y1              =   -15
         Y2              =   330
      End
      Begin VB.Line DragLine 
         BorderColor     =   &H00D66C3F&
         BorderWidth     =   2
         Index           =   1
         X1              =   315
         X2              =   90
         Y1              =   120
         Y2              =   345
      End
      Begin VB.Line DragLine 
         BorderColor     =   &H00D66C3F&
         BorderWidth     =   2
         Index           =   0
         X1              =   300
         X2              =   150
         Y1              =   210
         Y2              =   360
      End
   End
   Begin VB.CommandButton cmdOptions 
      Height          =   300
      Left            =   2340
      TabIndex        =   8
      Top             =   2460
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox NoteHeader 
      BackColor       =   &H00D66C3F&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   2985
      TabIndex        =   1
      Top             =   -15
      Width           =   2985
      Begin VB.Label lblSpacer 
         BackStyle       =   0  'Transparent
         Height          =   60
         Left            =   2685
         TabIndex        =   7
         Top             =   15
         Width           =   315
      End
      Begin VB.Label cmdReturnToMain 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   390
         Left            =   1845
         TabIndex        =   6
         Top             =   -30
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lblExtFun 
         BackStyle       =   0  'Transparent
         Caption         =   "¤"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   11.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   390
         Left            =   2145
         TabIndex        =   5
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lblCustomise 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   11.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   2400
         TabIndex        =   4
         Top             =   60
         Width           =   300
      End
      Begin VB.Label lblClose 
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2730
         TabIndex        =   3
         Top             =   -75
         Width           =   210
      End
      Begin VB.Image AlarmIcon 
         Height          =   585
         Left            =   -45
         Picture         =   "PostedNote.frx":000C
         Top             =   -30
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Sticky Note"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   75
         TabIndex        =   2
         Top             =   60
         Width           =   1080
      End
   End
   Begin VB.PictureBox note 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F7F8&
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   675
      MousePointer    =   1  'Arrow
      ScaleHeight     =   1050
      ScaleWidth      =   1410
      TabIndex        =   0
      Top             =   675
      Width           =   1410
   End
   Begin VB.Shape Border 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00D66C3F&
      BorderWidth     =   3
      Height          =   2775
      Left            =   0
      Top             =   15
      Width           =   3000
   End
End
Attribute VB_Name = "PostedNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const BorderColOffset = -20
Const HighlightCol = &H80FF&

Private Sub AlarmIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 4 Then RaiseAlarm
End Sub

Private Sub cmdAOT_Click()
If OptionsForm.chkQuickMode = 0 Then
    Render.AlwaysOnTop Me, True 'render on top of all
    NoteSetting(cmdOptions.Tag).nsAlwaysOnTop = True  'updates setting array with AOT setting
Else
    mainform.NewNote
End If
End Sub

Public Function LoadFromSettingsDialog()
'spcial effects loaded
If frmOptions.chkReminder = 1 Then
    AlarmIcon.Visible = True
    lblTitle.Left = 520
        Alarm(Notes).min = frmOptions.min.Text
        Alarm(Notes).hour = frmOptions.hour.Text
        Alarm(Notes).part = frmOptions.part.Text
        Alarm(Notes).Day = frmOptions.remDate.Day
        Alarm(Notes).Month = frmOptions.remDate.Month
        Alarm(Notes).Year = frmOptions.remDate.Year
        Alarm(Notes).Set = True
    
    AlarmIcon.ToolTipText = "ALARM set for: " & Alarm(Notes).Day & "/" & Alarm(Notes).Month & "/" & Alarm(Notes).Year & " At " & Alarm(Notes).hour & ":" & Alarm(Notes).min & " " & Alarm(Notes).part
Else
    Alarm(Notes).Set = False:    AlarmIcon.Visible = False:    lblTitle.Left = 105
End If

'Applys the universal settings from the customise dialog box.
If frmOptions.chkLockEdit = 0 Then
    Me.note.Locked = False: NoteSetting(cmdOptions.Tag).nsLockEdit = False
Else: Me.note.Locked = True: NoteSetting(cmdOptions.Tag).nsLockEdit = True: End If

If frmOptions.chkAOT.Value = 1 Then
    Render.AlwaysOnTop Me, True: NoteSetting(cmdOptions.Tag).nsAlwaysOnTop = True
Else: NoteSetting(cmdOptions.Tag).nsAlwaysOnTop = False: End If

If frmOptions.chkTransparent = 1 Then
    Render.TranslucenthWnd Me.hwnd, 255 * (frmOptions.TransparancySlider.Value / 100): NoteSetting(cmdOptions.Tag).nsTransparancyEnabled = True: NoteSetting(cmdOptions.Tag).nsTransparancyLevel = frmOptions.TransparancySlider.Value
Else: NoteSetting(cmdOptions.Tag).nsTransparancyEnabled = False: NoteSetting(cmdOptions.Tag).nsTransparancyLevel = 100: End If

If frmOptions.chkRandCol = vbChecked Then
    Border.BorderColor = Theme.NewShade(note.BackColor, BorderColOffset): NoteHeader.BackColor = Theme.NewShade(note.BackColor, BorderColOffset): NoteSetting(cmdOptions.Tag).nsRandomBorderColour = True
Else: Border.BorderColor = frmOptions.NoteHeader(0).BackColor: NoteHeader.BackColor = frmOptions.NoteHeader(0).BackColor:: NoteSetting(cmdOptions.Tag).nsRandomBorderColour = False: End If

If frmOptions.chkInvisible = 1 Then
    CapWindow Me: Border.Visible = False: Me.BackColor = RGB(255, 255, 254): note.BackColor = RGB(255, 255, 254): DragCorner.BackColor = RGB(255, 255, 254): Border.BorderColor = RGB(255, 255, 254): WindowShape Me, RGB(255, 255, 254): NoteSetting(cmdOptions.Tag).nsInvisibleBackground = True
Else: NoteSetting(cmdOptions.Tag).nsInvisibleBackground = False: End If

AssimDragIco
End Function


Public Function LoadFromFile()
'cmdOptions.Tag holds a tag for the note which indicates the note number
'in the note array in order to match up with its corresponding noteseting
'and alarm array data.

'Applys the new settings from alarm array which is filled up from the IO file.
If Alarm(cmdOptions.Tag).Set = True Then
    AlarmIcon.Visible = True
    lblTitle.Left = 520
    AlarmIcon.ToolTipText = "ALARM set for the: " & Alarm(cmdOptions.Tag).Day & "/" & Alarm(cmdOptions.Tag).Month & "/" & Alarm(cmdOptions.Tag).Year & " At " & Alarm(cmdOptions.Tag).hour & ":" & Alarm(cmdOptions.Tag).min & " " & Alarm(cmdOptions.Tag).part
Else
    AlarmIcon.Visible = False
    lblTitle.Left = 105
End If

'Loads special effect settings from the NoteSetting array assignments set from IO file.
If NoteSetting(cmdOptions.Tag).nsLockEdit = False Then Me.note.Locked = False Else Me.note.Locked = True
If NoteSetting(cmdOptions.Tag).nsAlwaysOnTop = True Then Render.AlwaysOnTop Me, True
If NoteSetting(cmdOptions.Tag).nsTransparancyEnabled = True Then Render.TranslucenthWnd Me.hwnd, 255 * (NoteSetting(cmdOptions.Tag).nsTransparancyLevel / 100)
If NoteSetting(cmdOptions.Tag).nsInvisibleBackground = True Then CapWindow Me: Me.BackColor = RGB(255, 255, 254): Border.Visible = False: note.BackColor = RGB(255, 255, 254): DragCorner.BackColor = RGB(255, 255, 254): Border.BorderColor = RGB(255, 255, 254): WindowShape Me, RGB(255, 255, 254)
AssimDragIco
End Function


Public Sub KillAlarm()
Alarm(cmdOptions.Tag).Set = False
AlarmIcon.Visible = False
lblTitle.Left = 105
End Sub

Public Sub RaiseAlarm()
For i = 1 To 3
    ShakeForm Me, 5: Sleep 100
    ShakeForm Me, 5: Sleep 100
    ShakeForm Me, 5: Sleep 150
    DoEvents
    
    ShakeForm Me, 1: Sleep 100
    ShakeForm Me, 1: Sleep 200
Next i
Sleep 300
ShakeForm Me, 11: Sleep 100
ShakeForm Me, 5: Sleep 150
DoEvents

ShakeForm Me, 1: Sleep 100
ShakeForm Me, 1: Sleep 100
End Sub

Private Sub cmdReturnToMain_Click()
mainform.Show
OptionsForm.chkQuickMode.Value = 0
For i = 0 To UBound(NoteInstance)
    NoteInstance(i).cmdReturnToMain.Visible = False
Next i
End Sub

Private Sub cmdReturnToMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseReset
cmdReturnToMain.ForeColor = HighlightCol
End Sub

Private Sub DragCorner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim CursorPosition As POINTAPI

DragCorner.Tag = "1" ' This is a flag used to see if the mouse button is being held down by the user
X = 15.2 ' This is a conversion value between tweeps and pixels

' This is a loop which changes the size of the form based on the position of the mouse on the screen
Do
    GetCursorPos CursorPosition ' Gets the cursors positoin
        
    Me.Height = ((CursorPosition.Y - (Me.Top / X)) * X) + 30 ' The V-scaling equation
    Me.Width = ((CursorPosition.X - (Me.Left / X)) * X) - 10 ' The H-scaling equation
    
    DoEvents ' tells VB to do other stuff in between
    
    If Me.Width < 2310 Or Me.Height < 700 Then
        If Me.Width < 2310 Then Me.Width = 2360: Me.Refresh
        If Me.Height < 700 Then Me.Height = 750: Me.Refresh
        DragCorner.Tag = "0"
    End If
    
Loop Until DragCorner.Tag = "0"

Me.Refresh
End Sub

Private Sub DragCorner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragCorner.Tag = "0" ' Disables the mouse being held down tag
End Sub

Private Sub Form_Load()
Me.Enabled = True

'Sizes the textbox to fit the form exactly
note.Top = 320: note.Left = 30 '10
note.Width = Me.Width - 60: note.Height = Me.Height - 350 '- 250

'copies the content of the textbox from the editor into the note
note.TextRTF = mainform.Text.TextRTF

note.BackColor = frmOptions.sample.BackColor
Border.Visible = True
Me.BackColor = &H8000000F

'Checks and edits the note approprietly if Quick-Post is on.
If OptionsForm.chkQuickMode = 1 Then: cmdReturnToMain.Visible = True: Me.lblExtFun.Caption = "r": frmOptions.chkLockEdit.Value = False: Else Me.lblExtFun.Caption = "¤"

If cmdOptions.Tag = "" Then Exit Sub ' If the NoteID has not been assigned yet to this note then it should not proceed with loading note styles.
'spcial effects loaded
If frmOptions.chkReminder = 1 Then
    AlarmIcon.Visible = True
    lblTitle.Left = 520
        Alarm(Notes).min = frmOptions.min.Text
        Alarm(Notes).hour = frmOptions.hour.Text
        Alarm(Notes).part = frmOptions.part.Text
        Alarm(Notes).Day = frmOptions.remDate.Day
        Alarm(Notes).Month = frmOptions.remDate.Month
        Alarm(Notes).Year = frmOptions.remDate.Year
        Alarm(Notes).Set = True
    
    AlarmIcon.ToolTipText = "ALARM set for: " & Alarm(Notes).Day & "/" & Alarm(Notes).Month & "/" & Alarm(Notes).Year & " At " & Alarm(Notes).hour & ":" & Alarm(Notes).min & " " & Alarm(Notes).part
Else
    Alarm(Notes).Set = False
    AlarmIcon.Visible = False
    lblTitle.Left = 105
End If

'Applys the universal settings from the customise dialog box.
If frmOptions.chkAOT.Value = 1 Then
    Render.AlwaysOnTop Me, True: NoteSetting(cmdOptions.Tag).nsAlwaysOnTop = True
Else: NoteSetting(cmdOptions.Tag).nsAlwaysOnTop = False: End If

If frmOptions.chkTransparent = 1 Then
    Render.TranslucenthWnd Me.hwnd, 255 * (frmOptions.TransparancySlider.Value / 100): NoteSetting(cmdOptions.Tag).nsTransparancyEnabled = True: NoteSetting(cmdOptions.Tag).nsTransparancyLevel = frmOptions.TransparancySlider.Value
Else: NoteSetting(cmdOptions.Tag).nsTransparancyEnabled = False: NoteSetting(cmdOptions.Tag).nsTransparancyLevel = 100: End If

If frmOptions.chkLockEdit = 0 Then
    note.Locked = False: NoteSetting(cmdOptions.Tag).nsLockEdit = False
Else: note.Locked = True: NoteSetting(cmdOptions.Tag).nsLockEdit = True: End If

If frmOptions.chkRandCol = 1 Then
    note.Locked = False: NoteSetting(cmdOptions.Tag).nsLockEdit = False
Else: note.Locked = True: NoteSetting(cmdOptions.Tag).nsLockEdit = True: End If

If frmOptions.chkRandCol = vbChecked Then
    Border.BorderColor = Theme.NewShade(note.BackColor, BorderColOffset): NoteHeader.BackColor = Theme.NewShade(note.BackColor, BorderColOffset): NoteSetting(cmdOptions.Tag).nsRandomBorderColour = True
Else: Border.BorderColor = frmOptions.NoteHeader(0).BackColor: NoteHeader.BackColor = frmOptions.NoteHeader(0).BackColor: NoteSetting(cmdOptions.Tag).nsRandomBorderColour = False: End If
AssimDragIco

If frmOptions.chkInvisible = 1 Then
    CapWindow Me: Border.Visible = False: Me.BackColor = RGB(255, 255, 254): note.BackColor = RGB(255, 255, 254): DragCorner.BackColor = RGB(255, 255, 254): Border.BorderColor = RGB(255, 255, 254): WindowShape Me, RGB(255, 255, 254): NoteSetting(cmdOptions.Tag).nsInvisibleBackground = True: Exit Sub
Else: NoteSetting(cmdOptions.Tag).nsInvisibleBackground = False: End If
End Sub


Private Sub Form_Resize()
On Error GoTo RestoreMin:
    DragCorner.BackColor = note.BackColor
    DragCorner.Left = Me.Width - DragCorner.Width - 30
    DragCorner.Top = Me.Height - DragCorner.Height - 30

    Border.Width = Me.Width
    Border.Height = Me.Height - 20

    note.Width = Me.Width - 70
    note.Height = Me.Height - 345

    NoteHeader.Width = Me.Width

    lblClose.Left = NoteHeader.Width - 300
    lblCustomise.Left = lblClose.Left - 300
    lblExtFun.Left = lblCustomise.Left - 260
    cmdReturnToMain.Left = lblExtFun.Left - 300
    Exit Sub
RestoreMin:
Me.Width = 2360
Me.Height = 750
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseReset
End Sub

Private Sub lblClose_Click()
Unload Me ' closes window
Notes = Notes - 1
If Notes = 0 Then mainform.TileNotes.Enabled = False: mainform.CloseAllNotes.Enabled = False
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseReset
lblClose.ForeColor = HighlightCol
End Sub

Private Sub lblCustomise_Click()
If Border.Tag = "1" Then Exit Sub
Dim NoteOptions As New frmOptions

NoteOptions.LoadValuesFromFile cmdOptions.Tag 'synchronises special effects and alarm settings

'Synchronises information with the note in the options available
NoteOptions.sample.BackColor = note.BackColor
NoteOptions.sample2.BackColor = NoteOptions.sample.BackColor
NoteOptions.chkReminder.Value = Abs(Int(Alarm(cmdOptions.Tag).Set))
NoteOptions.sample.TextRTF = note.TextRTF
NoteOptions.sample2.TextRTF = note.TextRTF
NoteOptions.Border(0).BorderColor = NoteHeader.BackColor: NoteOptions.Border(1).BorderColor = NoteHeader.BackColor
NoteOptions.NoteHeader(0).BackColor = NoteHeader.BackColor: NoteOptions.NoteHeader(1).BackColor = NoteHeader.BackColor
NoteOptions.AssimDragIco

NoteOptions.Show 'shows the option box instence

Border.Tag = "1"

'gets rid of the post note button fron the new option box instence and realigns the buttons
NoteOptions.cmdPost.Visible = False: NoteOptions.cmdPost.Enabled = False   'takes care of the button
NoteOptions.cmdCancel.Left = NoteOptions.cmdCancel.Left + NoteOptions.cmdPost.Width + 70 'alligns cancel button
NoteOptions.cmdOK.Left = NoteOptions.cmdOK.Left + NoteOptions.cmdCancel.Width 'alligns the ok button


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
                    Do 'loops until the options box is closed
                            DoEvents
                    Loop Until NoteOptions.Visible = False
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Border.Tag = "0"

'Makes sure there will only be one instance of the options box
If NoteOptions.cmdCancel.Tag = "1" Then NoteOptions.cmdCancel.Tag = "": Exit Sub

'Applys the new settings from the customise dialog.
If NoteOptions.chkReminder = 1 Then
    AlarmIcon.Visible = True
    lblTitle.Left = 520
        Alarm(cmdOptions.Tag).min = NoteOptions.min.Text
        Alarm(cmdOptions.Tag).hour = NoteOptions.hour.Text
        Alarm(cmdOptions.Tag).part = NoteOptions.part.Text
        Alarm(cmdOptions.Tag).Day = NoteOptions.remDate.Day
        Alarm(cmdOptions.Tag).Month = NoteOptions.remDate.Month
        Alarm(cmdOptions.Tag).Year = NoteOptions.remDate.Year
        Alarm(cmdOptions.Tag).Set = True
     AlarmIcon.ToolTipText = "ALARM set for the: " & Alarm(cmdOptions.Tag).Day & "/" & Alarm(cmdOptions.Tag).Month & "/" & Alarm(cmdOptions.Tag).Year & " At " & Alarm(cmdOptions.Tag).hour & ":" & Alarm(cmdOptions.Tag).min & " " & Alarm(cmdOptions.Tag).part
Else
    Alarm(cmdOptions.Tag).Set = False
    AlarmIcon.Visible = False
    lblTitle.Left = 105
End If

'Loads special effect settings and saves settings into NoteSetting structure
If NoteOptions.chkAOT.Value = 1 Then
    Render.AlwaysOnTop Me, True: NoteSetting(cmdOptions.Tag).nsAlwaysOnTop = True
Else: NoteSetting(cmdOptions.Tag).nsAlwaysOnTop = False: AlwaysOnTop Me, False: End If

If NoteOptions.chkTransparent = 1 Then
    Render.TranslucenthWnd Me.hwnd, 255 * (NoteOptions.TransparancySlider.Value / 100): NoteSetting(cmdOptions.Tag).nsTransparancyEnabled = True: NoteSetting(cmdOptions.Tag).nsTransparancyLevel = NoteOptions.TransparancySlider.Value
Else: NoteSetting(cmdOptions.Tag).nsTransparancyEnabled = False: NoteSetting(cmdOptions.Tag).nsTransparancyLevel = 100: Render.ReleaseWindow Me.hwnd:   End If

If NoteOptions.chkLockEdit = 0 Then
    note.Locked = False: NoteSetting(cmdOptions.Tag).nsLockEdit = False
Else: note.Locked = True: NoteSetting(cmdOptions.Tag).nsLockEdit = True: End If

If NoteOptions.chkRandCol = vbChecked Then
    Border.BorderColor = Theme.NewShade(note.BackColor, BorderColOffset): NoteHeader.BackColor = Theme.NewShade(note.BackColor, BorderColOffset): NoteSetting(cmdOptions.Tag).nsRandomBorderColour = True
Else: Border.BorderColor = NoteOptions.NoteHeader(0).BackColor: NoteHeader.BackColor = NoteOptions.NoteHeader(0).BackColor: NoteSetting(cmdOptions.Tag).nsRandomBorderColour = False: End If

AssimDragIco

If NoteOptions.chkInvisible = 1 Then
    CapWindow Me: Border.Visible = False: Me.BackColor = RGB(255, 255, 254): note.BackColor = RGB(255, 255, 254): DragCorner.BackColor = RGB(255, 255, 254): Border.BorderColor = RGB(255, 255, 254): WindowShape Me, RGB(255, 255, 254): NoteSetting(cmdOptions.Tag).nsInvisibleBackground = True: Exit Sub
Else: NoteSetting(cmdOptions.Tag).nsInvisibleBackground = False: End If

note.BackColor = NoteOptions.sample.BackColor
Border.Visible = True
AssimDragIco
Form_Resize
Me.Refresh
note.Refresh
End Sub

Private Sub lblCustomise_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseReset
lblCustomise.ForeColor = HighlightCol
End Sub

Private Sub lblExtFun_Click()
If OptionsForm.chkQuickMode = 0 Then
    Render.AlwaysOnTop Me, True 'render on top of all
    NoteSetting(cmdOptions.Tag).nsAlwaysOnTop = True  'updates setting array with AOT setting
Else
    mainform.NewNote
End If
End Sub

Private Sub lblExtFun_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseReset
lblExtFun.ForeColor = HighlightCol
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Render.DragForm Me, Me 'used for drag effect
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = 0 'used for drag effect
End Sub


Private Sub note_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If NoteSetting(cmdOptions.Tag).nsLockEdit = True Then Render.DragForm note, Me 'used for drag effect
End Sub

Private Sub note_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseReset
End Sub

Private Sub NoteHeader_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Render.DragForm Me, Me 'used for drag effect
End Sub

Private Sub NoteHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseReset
End Sub

Public Function MouseReset()
lblClose.ForeColor = vbWhite
lblCustomise.ForeColor = vbWhite
lblExtFun.ForeColor = vbWhite
cmdReturnToMain.ForeColor = vbWhite
Me.MousePointer = 0 'used for drag effect
note.MousePointer = 0
End Function

Public Function AssimDragIco()
DragCorner.BackColor = note.BackColor
For i = 0 To 2
    DragLine(i).BorderColor = NoteHeader.BackColor
Next i
End Function
