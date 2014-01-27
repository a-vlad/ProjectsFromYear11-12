VERSION 5.00
Begin VB.Form mainform 
   Caption         =   "Untitled - Semipad"
   ClientHeight    =   6120
   ClientLeft      =   3600
   ClientTop       =   2850
   ClientWidth     =   8430
   Icon            =   "semiPad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   8430
   Begin VB.PictureBox Text 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4755
      Left            =   30
      ScaleHeight     =   4755
      ScaleWidth      =   8385
      TabIndex        =   0
      Top             =   360
      Width           =   8385
   End
   Begin VB.PictureBox Status 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   8370
      TabIndex        =   1
      Top             =   5820
      Visible         =   0   'False
      Width           =   8430
   End
   Begin VB.PictureBox Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   300
      ScaleWidth      =   8370
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   8430
      Begin VB.ComboBox ComboFontNames 
         Height          =   315
         Left            =   30
         TabIndex        =   4
         Text            =   "Arial"
         Top             =   0
         Width           =   2730
      End
      Begin VB.ComboBox ComboFontSize 
         Height          =   315
         ItemData        =   "semiPad.frx":628A
         Left            =   2790
         List            =   "semiPad.frx":62C9
         TabIndex        =   3
         Text            =   "12"
         Top             =   0
         Width           =   840
      End
   End
   Begin VB.PictureBox DefaultToolBar 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   3000
      ScaleHeight     =   420
      ScaleWidth      =   540
      TabIndex        =   5
      Top             =   5160
      Width           =   600
   End
   Begin VB.PictureBox DefaultFormatToolBar 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   3720
      ScaleHeight     =   420
      ScaleWidth      =   540
      TabIndex        =   6
      Top             =   5160
      Width           =   600
   End
   Begin VB.PictureBox DefaultWindow 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   4440
      ScaleHeight     =   420
      ScaleWidth      =   540
      TabIndex        =   7
      Top             =   5160
      Width           =   600
   End
   Begin VB.Timer AlarmCheck 
      Interval        =   1000
      Left            =   5115
      Top             =   5190
   End
   Begin VB.PictureBox DialogBox 
      Height          =   480
      Left            =   720
      ScaleHeight     =   420
      ScaleWidth      =   540
      TabIndex        =   8
      Top             =   5160
      Width           =   600
   End
   Begin VB.Timer SystemSynch 
      Interval        =   3000
      Left            =   120
      Top             =   5160
   End
   Begin VB.PictureBox ToolbarIcons 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   5640
      ScaleHeight     =   420
      ScaleWidth      =   540
      TabIndex        =   9
      Top             =   5160
      Width           =   600
   End
   Begin VB.Shape Border 
      BorderColor     =   &H00CCAE93&
      Height          =   4575
      Left            =   0
      Top             =   345
      Width           =   8385
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu Open 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu Spacer 
         Caption         =   "-"
      End
      Begin VB.Menu PageSetup 
         Caption         =   "Pa&ge Setup..."
      End
      Begin VB.Menu Print 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu spacer2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Undo 
         Caption         =   "Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu spacer3 
         Caption         =   "-"
      End
      Begin VB.Menu Cut 
         Caption         =   "Cut"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu Copy 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
         Shortcut        =   ^U
      End
      Begin VB.Menu Delete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu Spacer4 
         Caption         =   "-"
      End
      Begin VB.Menu Find 
         Caption         =   "Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu FindNext 
         Caption         =   "Find Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Replace 
         Caption         =   "Replace..."
         Shortcut        =   ^H
      End
      Begin VB.Menu WordCount 
         Caption         =   "Word Count..."
         Shortcut        =   ^W
      End
      Begin VB.Menu Spacer5 
         Caption         =   "-"
      End
      Begin VB.Menu SelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu TimeDate 
         Caption         =   "Time/Date"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu Format 
      Caption         =   "Format"
      Begin VB.Menu WordWrap 
         Caption         =   "Word Wrap"
      End
      Begin VB.Menu Spacer11 
         Caption         =   "-"
      End
      Begin VB.Menu FontColour 
         Caption         =   "Text Colour..."
      End
      Begin VB.Menu Font 
         Caption         =   "Text Font..."
      End
      Begin VB.Menu Spacer12 
         Caption         =   "-"
      End
      Begin VB.Menu InsertPicture 
         Caption         =   "Insert Picture"
      End
   End
   Begin VB.Menu View 
      Caption         =   "View"
      Begin VB.Menu Viewtoolbar 
         Caption         =   "Toolbar"
         Shortcut        =   {F2}
      End
      Begin VB.Menu CustomiseTool 
         Caption         =   "Customise"
      End
      Begin VB.Menu Spacer10 
         Caption         =   "-"
      End
      Begin VB.Menu StatusBar 
         Caption         =   "Status Bar"
      End
      Begin VB.Menu spacer9 
         Caption         =   "-"
      End
      Begin VB.Menu Options 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu NoteMenu 
      Caption         =   "Notes"
      Begin VB.Menu PostNote 
         Caption         =   "Post Note..."
         Shortcut        =   ^Q
      End
      Begin VB.Menu spacer7 
         Caption         =   "-"
      End
      Begin VB.Menu OrganiseNotes 
         Caption         =   "Organise"
         Begin VB.Menu TileNotes 
            Caption         =   "Tile Notes"
         End
         Begin VB.Menu CloseAllNotes 
            Caption         =   "Close All Notes"
         End
      End
      Begin VB.Menu spacer13 
         Caption         =   "-"
      End
      Begin VB.Menu SaveNotes 
         Caption         =   "Save Notes"
      End
      Begin VB.Menu LoadNotes 
         Caption         =   "Load Notes"
      End
      Begin VB.Menu spacer8 
         Caption         =   "-"
      End
      Begin VB.Menu Customise 
         Caption         =   "Customise"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu HelpTopics 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu Spacer6 
         Caption         =   "-"
      End
      Begin VB.Menu AboutSemiPad 
         Caption         =   "About SemiPad"
      End
   End
End
Attribute VB_Name = "mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function DeleteSelText()
Dim InputKey As GENERALINPUT

InputKey.dwType = 1 ' Sets the Type of input to Keyboard
InputKey.xi(0) = &H2E 'Sets the key to send to the keyboard which is the DEL key

SendInput 1, InputKey, Len(InputKey) ' Sends the keystroke
End Function

Private Sub AboutSemiPad_Click()
Me.Enabled = False ' disables switching back to the me form from the about form
AboutForm.Show 'shows the about form
End Sub

Private Sub AlarmCheck_Timer()
Dim AlarmDate As Date
Dim CurTime As String
For i = 0 To Notes - 1
 'scans all the notes in the array and checks if there is an alarm to ring or not
 If Alarm(i).Set = True Then
    'checks if the alarm time and date match the current time and date
    AlarmDate = DateSerial(Alarm(i).Year, Alarm(i).Month, Alarm(i).Day)
    AlarmTime = Alarm(i).hour & ":" & Alarm(i).min & ":00 " & Alarm(i).part
    
    CurTime = Time
    'if the time and date match it then sends an alarm to ring
    If AlarmDate = Date And Left(AlarmTime, 2) = Left(CurTime, 2) And _
        Right(AlarmTime, 2) = Right(CurTime, 2) And _
        Left(Right(AlarmTime, 8), 2) = Left(Right(CurTime, 8), 2) Then _
    NoteInstance(i).RaiseAlarm: NoteInstance(i).KillAlarm
    
 End If
Next i
End Sub

Private Sub CloseAllNotes_Click()
On Error Resume Next
'loops all the note instances
For i = 0 To 29
    'every note instance will be closed.
    Unload NoteInstance(i)
Next i
Notes = 0
TileNotes.Enabled = False: CloseAllNotes.Enabled = False
End Sub

Private Sub ComboFontNames_Click()
'updates the selected font
Text.SelFontName = ComboFontNames.Text
End Sub

Private Sub ComboFontSize_Click()
'updates the selected font size
Text.SelFontSize = ComboFontSize.Text
End Sub

Private Sub Copy_Click()
Clipboard.Clear 'Clears the previous text
Clipboard.SetText Text.SelText 'copies the text to the clipboard
End Sub

Private Sub Customise_Click()
'shows the customise dialog
frmOptions.Show
'updates the cusotmise templates
frmOptions.sample.TextRTF = Text.TextRTF
frmOptions.sample2.TextRTF = Text.TextRTF

'disables the main form
Me.Enabled = False
End Sub

Private Sub CustomiseTool_Click()
Toolbar.Customize ' displays the toolbar controller's inbuilt customise dialog.
End Sub

Private Sub Cut_Click()
Clipboard.Clear 'Clears the previous text
Clipboard.SetText Text.SelText 'saves the text to the clipboard
DeleteSelText ' Calls the delete text function
End Sub

Private Sub Delete_Click()
DeleteSelText ' Calls the delete text function
End Sub

Private Sub DragCorner_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    DragCorner.Tag = "0" ' Disables the mouse being held down tag
End Sub

Private Sub Exit_Click()
Unload Me 'unloads the form
End Sub

Private Sub Find_Click()
FindDialog.Show ' displayes the find dialogbox
End Sub

Private Sub FindNext_Click()
If FindDialog.FindText = "" Then FindDialog.Show: Exit Sub ' makes sure there is previous seach made
Text.Find FindDialog.FindText.Text, Text.SelStart + Len(FindDialog.FindText.Text), Len(Text.Text) 'user the previous search
End Sub

Private Sub Font_Click()
On Error GoTo SkipLoadFont:
'Sets up the font chooser box properties before launch
DialogBox.flags = cdlCFApply Or cdlCFEffects Or cdlCFBoth 'Tells it to show the effects chooser

'Loads the current fornt settings into the font boz
DialogBox.FontName = Text.SelFontName
DialogBox.FontSize = Text.SelFontSize
DialogBox.Color = Text.SelColor
DialogBox.FontBold = Text.SelBold
DialogBox.FontItalic = Text.SelItalic
DialogBox.FontUnderline = Text.SelUnderline
DialogBox.FontStrikethru = Text.SelStrikeThru

SkipLoadFont:
If Err.Number = 32755 Then Exit Sub
'Shows the font box
DialogBox.ShowFont

'Updates the new settings of the fornt box
Text.SelFontName = DialogBox.FontName
Text.SelFontSize = DialogBox.FontSize
'Text.SelColor = DialogBox.Color
Text.SelBold = DialogBox.FontBold
Text.SelItalic = DialogBox.FontItalic
Text.SelUnderline = DialogBox.FontUnderline
Text.SelStrikeThru = DialogBox.FontStrikethru
End Sub

Private Sub FontColour_Click()
On Error Resume Next
DialogBox.ShowColor
Text.SelColor = DialogBox.Color
End Sub

Private Sub Form_Load()
'Sets up universal File System Variables
SavedChangesChecksum = 0
Notes = 0
DocumentTitle = "Untitled"
FilePath = "Untitled"
If Notes = 0 Then TileNotes.Enabled = False: CloseAllNotes.Enabled = False

'Loads editor settings from file
LoadEditorSettings
If OptionsForm.chkReloadNotes.Value = 1 Then LoadNotes_Click 'loads previously saved notes if the option is ticked

'Loads Font list
For FontCount = 0 To Screen.FontCount - 1
   ComboFontNames.AddItem Screen.Fonts(FontCount)
Next

'Turns the text blue
SelectAll_Click
Text.SelColor = RGB(0, 0, 0)
Text.SelStart = 0
Status.Panels.Item(1).Text = "For Help, press F1"

'Text.SelColor = vbBlue
'me.ScaleMode = 3 'Sets the scale mode of the form to pixels NOT TWEEPS
If Clipboard.GetText = "" Then Paste.Enabled = False 'Checks if the clipboard contails any text
End Sub

Public Function SaveEditorSettings()
Open App.Path & "\settings.dat" For Output As #3
    ' this function writes to file the toolbar settings
    Write #3, WordWrap.Checked
    Write #3, Viewtoolbar.Checked
    Write #3, StatusBar.Checked
    'the option form settings
    Write #3, OptionsForm.chkAutoLoad.Value
    Write #3, OptionsForm.chkReloadNotes.Value
    Write #3, OptionsForm.chkAnimateTileEffect.Value
    'and the main form editor size and position on the user's screen
    Write #3, Me.Top
    Write #3, Me.Left
    Write #3, Me.Width
    Write #3, Me.Height
Close #3
End Function

Public Function LoadEditorSettings()
On Error GoTo CreateNewFile
Dim OutDataVal As Variant
Open App.Path & "\settings.dat" For Input As #4
    'This function will read the date saved from before in a the same order
    'and apply the values  back into the corresponding locations
    Input #4, OutDataVal: If OutDataVal = True Then WordWrap_Click
    Input #4, OutDataVal: If OutDataVal = True Then Viewtoolbar_Click
    Input #4, OutDataVal: If OutDataVal = True Then StatusBar_Click
    Input #4, OutDataVal: OptionsForm.chkAutoLoad.Value = OutDataVal
    Input #4, OutDataVal: OptionsForm.chkReloadNotes.Value = OutDataVal
    Input #4, OutDataVal: OptionsForm.chkAnimateTileEffect.Value = OutDataVal
    Input #4, InputVar: Me.Top = InputVar
    Input #4, InputVar: Me.Left = InputVar
    Input #4, InputVar: Me.Width = InputVar
    Input #4, InputVar: Me.Height = InputVar
    Me.Refresh
Close #4
Exit Function
CreateNewFile:
SaveEditorSettings
End Function

Public Function LoadOptionsOnly()
Dim OutDataVal As Variant
Open App.Path & "\settings.dat" For Input As #5
    'this function is used for the cancel button in order to reload the previous settings.
    Input #5, a, b, c
    Input #5, OutDataVal: OptionsForm.chkAutoLoad.Value = OutDataVal
    Input #5, OutDataVal: OptionsForm.chkReloadNotes.Value = OutDataVal
    Input #5, OutDataVal: OptionsForm.chkAnimateTileEffect.Value = OutDataVal
Close #5
Exit Function
End Function

Private Sub Form_Resize()
On Error Resume Next

'TextBox resize Position calebration Code
If StatusBar.Checked = True Then
    If Toolbar.Visible = False Then Text.Height = Me.ScaleHeight - Status.Height: Text.Top = Me.ScaleTop
    If Toolbar.Visible = True Then Text.Height = Me.ScaleHeight - Status.Height - Toolbar.Height: Text.Top = Toolbar.Height + 5
    Status.Panels.Item(1).Width = Abs(Me.Width - 2000)
Else
    If Toolbar.Visible = False Then Text.Height = Me.ScaleHeight: Text.Top = Me.ScaleTop
    If Toolbar.Visible = True Then Text.Height = Me.ScaleHeight - Toolbar.Height: Text.Top = Toolbar.Height + 5
End If

Text.Width = Me.ScaleWidth + 10
Text.Left = 0

''Frame Border resize Position calebration Code
Border.Height = Me.ScaleHeight
Border.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim MsgBoxReturn As String
If SavedChangesChecksum = GetFileChecksum Then SaveEditorSettings: End
'displayes the save changes box to the user
MsgBoxReturn = MsgBox("The text in the text file is unsaved." & vbCrLf & vbCrLf & "Do you wish to save the changes?", vbExclamation + vbYesNoCancel, "Semipad")

'checks the users response
If MsgBoxReturn = vbNo Then 'closes the program if the user doesn not wish to save
ElseIf MsgBoxReturn = vbYes Then 'saves if the user wants to save
 DialogBox.ShowSave
Else: Cancel = 1: Exit Sub
End If ' if the users clicks cancel the program does nothing

SaveEditorSettings
End
End Sub

Private Sub InsertPicture_Click()
On Error Resume Next
' this is the code used for the open dialog
DialogBox.DialogTitle = "Insert Picture"
DialogBox.Filter = "JPEG|*.JPG|Microsoft Bitmap|*.BMP|GIF|*.GIF"

'opens file
DialogBox.ShowOpen
Clipboard.Clear
Clipboard.SetData LoadPicture(DialogBox.FileName)
SendMessage Text.hwnd, &H302, 0, 0
End Sub

Private Sub LoadNotes_Click()
Dim InputVar As Variant, NewNotes As Integer
Open App.Path & "\notes.nts" For Input As #2
Input #2, NewNotes: Notes = Notes + NewNotes        'appends note index on top of the currently created notes.
If Notes > 0 Then TileNotes.Enabled = True: CloseAllNotes.Enabled = True 'enables note menu features

For i = Notes - NewNotes To Notes - 1
    NoteInstance(i).cmdOptions.Tag = i
    NoteInstance(i).Show
   
    'Tab 1
    Input #2, InputVar: NoteInstance(i).note.BackColor = InputVar
    Input #2, InputVar: NoteInstance(i).NoteHeader.BackColor = InputVar: NoteInstance(i).Border.BorderColor = InputVar
    
    'Tab 2
    Input #2, InputVar: NoteSetting(i).nsTransparancyEnabled = InputVar
    Input #2, InputVar: NoteSetting(i).nsTransparancyLevel = InputVar
    Input #2, InputVar: NoteSetting(i).nsInvisibleBackground = InputVar
    Input #2, InputVar: NoteSetting(i).nsAlwaysOnTop = InputVar
    Input #2, InputVar: NoteSetting(i).nsLockEdit = InputVar
    
    'Tab 3:
    Input #2, Alarm(i).Set
    Input #2, Alarm(i).Year
    Input #2, Alarm(i).Month
    Input #2, Alarm(i).Day
    Input #2, Alarm(i).hour
    Input #2, Alarm(i).min
    Input #2, Alarm(i).part
    
    NoteInstance(i).LoadFromFile 'applys the newly assigned variables Alarm and NoteSetting to the note style.
    
    'Content RTF:
    Input #2, InputVar: NoteInstance(i).note.TextRTF = InputVar
    
    'Screen Position:
    Input #2, InputVar: NoteInstance(i).Top = InputVar
    Input #2, InputVar: NoteInstance(i).Left = InputVar
    Input #2, InputVar: NoteInstance(i).Width = InputVar
    Input #2, InputVar: NoteInstance(i).Height = InputVar
    NoteInstance(i).Refresh
Next i
Close #2
End Sub

Private Sub New_Click()
CheckSavedChanges
End Sub

Private Sub NoteOptions_Click()
frmOptions.Show
Me.Enabled = False
End Sub

Private Sub Open_Click()
On Error Resume Next
' this is the code used for the open dialog
If CheckSavedChanges = "-1" Then Exit Sub   'Saves changes to document
DialogBox.DialogTitle = "Open Document"
DialogBox.Filter = "Text Document|*.TXT|Rich Text Format|*.rtf"

'opens file
DialogBox.ShowOpen
Text.LoadFile DialogBox.FileName

'Updates FileSystem Variables
FilePath = DialogBox.FileName
DocumentTitle = DialogBox.FileTitle
SavedChangesChecksum = GetFileChecksum

If DocumentTitle <> "" Then Me.Caption = DocumentTitle & " - Semipad"
End Sub

Private Sub Options_Click()
Me.Enabled = False
OptionsForm.Show
End Sub

Private Sub PageSetup_Click()
'sets up the properties of the dialog, these must be set
PageSetupDialog.lStructSize = Len(PageSetupDialog)
PageSetupDialog.hwndOwner = Me.hwnd

PageSetupDlg PageSetupDialog ' calls the dialog dunction from the windows api
End Sub

Private Sub Paste_Click()
Text.SelText = Clipboard.GetText ' This gets the text from the clipboard
End Sub

Public Function Post()
Me.BorderStyle = 0
End Function

Private Sub PostNote_Click()
NewNote
End Sub

Public Function NewNote()
'On Error Resume Next
NoteInstance(Notes).Show
NoteInstance(Notes).cmdOptions.Tag = Notes 'setes the NoteID index number of the note in the note array.
NoteInstance(Notes).LoadFromSettingsDialog

Notes = Notes + 1
If Notes > 0 Then TileNotes.Enabled = True: CloseAllNotes.Enabled = True
End Function


Private Sub Print_Click()
On Error Resume Next
'print file code
DialogBox.CancelError = True
DialogBox.ShowPrinter
End Sub

Private Sub Replace_Click()
'this shows the replace form, all the replace coding is done there
ReplaceForm.Show
End Sub

Private Sub Save_Click()
'saves the file with the current file location or shows the save as box if there has not been any saves
If FilePath = "Untitled" Then
SaveAs_Click
Else: Text.SaveFile FilePath: SavedChangesChecksum = GetFileChecksum
End If
End Sub

Private Sub SaveAs_Click()
On Error Resume Next
'Initializes the Common dialog box with the save as settings
DialogBox.DialogTitle = "Save As"
DialogBox.Filter = "Text Document|*.TXT|Rich Text Format|*.rtf"

'Launches the dialog to get the user file name
DialogBox.ShowSave

'saves the file
Text.SaveFile DialogBox.FileName
SavedChangesChecksum = GetFileChecksum
End Sub

Private Sub SaveNotes_Click()
MsgBoxReturn = MsgBox("Any previously saved notes will be overwritten." & vbCrLf & vbCrLf & "Do you want to save anyways?", vbExclamation + vbYesNo, "Semipad"): If MsgBoxReturn = vbNo Then Exit Sub
Open App.Path & "\notes.nts" For Output As #1 'Opens notefile for saving purposes.
Write #1, Notes 'Writes to the file header (first line of data) the number of notes stored in the file.

For i = 0 To Notes - 1 'Loops for the number of opened notes in order to save each individual note's settings and content.
    'Tab 1: Writes the colour settings from the customise dialog.
    Write #1, NoteInstance(i).note.BackColor
    Write #1, NoteInstance(i).NoteHeader.BackColor
    
    'Tab 2: Writes the special effects settings from the customise dialog
    Write #1, NoteSetting(i).nsTransparancyEnabled                    'Booleon if transparancy is enabled in the note
    Write #1, NoteSetting(i).nsTransparancyLevel                      'The transparency level, if its disabled this will be 100 and wount be taken in account on reloading of note file.
    Write #1, NoteSetting(i).nsInvisibleBackground                    'Booleon, if true then the note background is seethrough.
    Write #1, NoteSetting(i).nsAlwaysOnTop                            'Always on top booleon, if true the note is always on top.
    Write #1, NoteSetting(i).nsLockEdit                               'Set to true if the note is ineditable
    
    'Tab 3: Writes the alarm settings from the customise dialog which stores them in the appropriate alarm array to correspond to the specific note.
    Write #1, Alarm(i).Set                  'Booleon, true if an alarm has been set for the specific note.
    Write #1, Alarm(i).Year                 'Alarm Year
    Write #1, Alarm(i).Month                'Alarm Month
    Write #1, Alarm(i).Day                  'Alarm Day
    Write #1, Alarm(i).hour                 'Alarm Hour
    Write #1, Alarm(i).min                  'Alarm Minute
    Write #1, Alarm(i).part                 'Alarm part of day: AM or PM
    
    'Content RTF: The content of the note itself.
    Write #1, NoteInstance(i).note.TextRTF          'RTF data of note
    
    'Screen Position: The screen positioning and size of the note for replacement on reload.
    NoteInstance(i).Refresh
    Write #1, NoteInstance(i).Top               'Position from the top of the screen
    Write #1, NoteInstance(i).Left              'Position from the left of the screen
    Write #1, NoteInstance(i).Width             'Width of the note
    Write #1, NoteInstance(i).Height            'Height of the note
Next i
Close #1 'Closes the file.
End Sub

Private Sub SelectAll_Click()
Text.SelStart = 0
Text.SelLength = Len(Text.Text)
Status.Panels(2).Text = "Chr: " & Text.SelStart & "  Sel: " & Text.SelLength
If Text.SelLength = 0 Then
Cut.Enabled = False: Copy.Enabled = False: Delete.Enabled = False
Else
Cut.Enabled = True: Copy.Enabled = True: Delete.Enabled = True
End If
End Sub

Private Sub StatusBar_Click()
On Error Resume Next
'This shows the status bar
If StatusBar.Checked = False Then ' this is done based on the tick of the option
    Status.Visible = True 'shows the bar
    StatusBar.Checked = True ' changes the tickbox
    Form_Resize
Else
    Status.Visible = False 'hides the bar
    StatusBar.Checked = False 'unchecks the option box
    Form_Resize
End If
End Sub

Private Sub SystemSynch_Timer()
'On Error Resume Next
'Checks if the clipboard contails any text
If Clipboard.GetText = "" Then
Paste.Enabled = False
Else
Paste.Enabled = True
End If
End Sub

Private Sub Text_Change()
Status.Panels(2).Text = "Chr: " & Text.SelStart & "  Sel: " & Text.SelLength
Undo.Enabled = True ' enables undo option in the edit menu

'Checks if there is any text to copy or cut or delete in window
If Text.SelLength = 0 Then
Cut.Enabled = False: Copy.Enabled = False: Delete.Enabled = False
Else
Cut.Enabled = True: Copy.Enabled = True: Delete.Enabled = True
End If
If Toolbar.Visible = True Then UpdateToolbar
End Sub

Private Sub Text_Click()
'This updates the statusbar information of the selection length and character position in text document
Status.Panels(2).Text = "Chr: " & Text.SelStart & "  Sel: " & Text.SelLength

'this checks if there is anything to cut, copy or delete (selected text in document)
If Text.SelLength = 0 Then 'if none the options are desabled
    Cut.Enabled = False: Copy.Enabled = False: Delete.Enabled = False
Else ' if there is selected text options are enables in the menu
    Cut.Enabled = True: Copy.Enabled = True: Delete.Enabled = True
End If
If Toolbar.Visible = True Then UpdateToolbar
End Sub

Public Function UpdateToolbar()
On Error Resume Next
'Updates the toolbar based on the selected text
If Text.SelBold = True Then Toolbar.Buttons.Item(2).Value = tbrPressed Else Toolbar.Buttons.Item(2).Value = tbrUnpressed
If Text.SelItalic = True Then Toolbar.Buttons.Item(3).Value = tbrPressed Else Toolbar.Buttons.Item(3).Value = tbrUnpressed
If Text.SelUnderline = True Then Toolbar.Buttons.Item(4).Value = tbrPressed Else Toolbar.Buttons.Item(4).Value = tbrUnpressed

Toolbar.Buttons.Item(6).Value = tbrUnpressed: Toolbar.Buttons.Item(7).Value = tbrUnpressed: Toolbar.Buttons.Item(8).Value = tbrUnpressed
If Text.SelAlignment = 0 Then
Toolbar.Buttons.Item(6).Value = tbrPressed
ElseIf Text.SelAlignment = 2 Then Toolbar.Buttons.Item(7).Value = tbrPressed
Else: Toolbar.Buttons.Item(8).Value = tbrPressed
End If

ComboFontNames.Text = Text.SelFontName
ComboFontSize.Text = Text.SelFontSize

If Text.SelBullet = 0 Then Toolbar.Buttons.Item(10).Value = tbrUnpressed Else Toolbar.Buttons.Item(10).Value = tbrPressed
End Function


Private Sub Text_KeyDown(KeyCode As Integer, Shift As Integer)
'This updates the statusbar information of the selection length and character position in text document
Status.Panels(2).Text = "Chr: " & Text.SelStart & "  Sel: " & Text.SelLength

'this checks if there is anything to cut, copy or delete (selected text in document)
If Text.SelLength = 0 Then 'if none the options are desabled
    Cut.Enabled = False: Copy.Enabled = False: Delete.Enabled = False
Else ' if there is selected text options are enables in the menu
    Cut.Enabled = True: Copy.Enabled = True: Delete.Enabled = True
End If

'Easter Egg 1 Enbedded : To View Press CTRL + F12 while a selection of "enter the matrix" is done
If KeyCode = vbKeyF12 And Shift = vbCtrlMask And Text.SelText = "enter the matrix" Then: Text.BackColor = vbBlack: Text.SelColor = vbGreen: For i = 1 To 10000:  Text.SelText = Len(Text.Text): DoEvents: Next i

'Easter Egg Enbedded 2 : To View Press CTRL + F12 while a selection of "spank the monkey" is done
If KeyCode = vbKeyF12 And Shift = vbCtrlMask And Text.SelText = "spank the monkey" Then For i = 0 To Notes - 1: Render.ShakeForm NoteInstance(i), 2: Next i
End Sub

Private Sub TileNotes_Click()
Dim Columns As Integer, Rows As Integer
Dim NotePos As RECT
Dim PrePos As RECT
Dim AnimateMovement As Boolean

pix = 15.2

AnimateMovement = OptionsForm.chkAnimateTileEffect.Value   'sets animation options on and off.

If Notes = 0 Then Exit Sub ' checks if there are any opened notes

'Positions the first note posted in the corner of the screen
If AnimateMovement = False Then
SetWindowPos NoteInstance(0).hwnd, -2, 0, 0, 200, 186, &H40:
Else: MoveTo NoteInstance(0), 0, 0: End If

Rows = Abs((GetDeviceCaps(Me.hdc, 10)) / 186) ' 4 autocalculates the maximum height (rows)
Columns = Abs((GetDeviceCaps(Me.hdc, 8)) / 200) ' 5 autocalculates the maximum width (coloumns)

For i = 1 To Rows
    For ii = (i * Columns) - (Columns - 1) To (i * Columns) - 1
        'Gets the note positions for the current note to be positions and the previously positions note
        GetWindowRect NoteInstance(ii - 1).hwnd, PrePos  'previous note
        GetWindowRect NoteInstance(ii).hwnd, NotePos  'current note to be positioned

        'makes sure there is visible notes to tile
        If NoteInstance(ii).Visible = False Then
            'Shifts the start position for the next row down the y axis
            'If AnimateMovement = False Then
            'SetWindowPos NoteInstance(ii).hwnd, -2, PrePos.Left + 200, (i - 1) * 186, 200, 186, &H40:
            'Else: MoveTo NoteInstance(ii), (i - 1) * 186, PrePos.Left + 200: End If
            'NoteInstance(ii).Visible = False
        Else
            'uses the information and calculates the next position for the next note.
            If AnimateMovement = False Then
            SetWindowPos NoteInstance(ii).hwnd, -2, PrePos.Left + (NoteInstance(ii).Width / pix), (i - 1) * (NoteInstance(ii).Height / pix), (NoteInstance(ii).Width / pix), (NoteInstance(ii).Height / pix), &H40:
            Else: MoveTo NoteInstance(ii), (i - 1) * (NoteInstance(ii).Height / pix), PrePos.Left + (NoteInstance(ii).Width / pix): End If
        End If
    Next ii
    
        'makes sure there is visible notes to tile
    If NoteInstance(i * Columns).Visible = False Then
        'Shifts the start position for the next row down the y axis
        'If AnimateMovement = False Then
        'SetWindowPos NoteInstance(i * Columns).hwnd, -2, 0, (i - 1) * 186, 200, 186, &H40
        'Else: MoveTo NoteInstance(i * Columns), (i - 1) * 186, 0: End If
        'NoteInstance(i * Columns).Visible = False
    Else
        'Shifts the start position for the next row down the y axis
        If AnimateMovement = False Then
        SetWindowPos NoteInstance(i * Columns).hwnd, -2, 0, (i) * (NoteInstance(i * Columns).Height / pix), (NoteInstance(i * Columns).Width / pix), (NoteInstance(i * Columns).Height / pix), &H40
        Else: MoveTo NoteInstance(i * Columns), (i) * (NoteInstance(i * Columns).Height / pix), 0: End If
    End If
Next i
End Sub


Public Function MoveTo(hForm As Form, NewTop As Integer, NewLeft As Integer)
Dim Pos As RECT
Dim NewPos As RECT
Dim CurPos As RECT
Dim PerSlide As Integer

'Positions the first note posted in the corner of the screen
GetWindowRect hForm.hwnd, Pos  'previous note

NewPos.Top = NewTop
NewPos.Left = NewLeft
PerSlide = 200 'this is the gradient used for the movement calculations

For i = 1 To PerSlide
   
    GetWindowRect hForm.hwnd, CurPos
       
    'uses the information and calculates the next position for the next note.
    SetWindowPos hForm.hwnd, -2, CurPos.Left + ((NewPos.Left - CurPos.Left) / PerSlide) * i, CurPos.Top + ((NewPos.Top - CurPos.Top) / 100) * i, 0, 0, &H40 Or &H1
    
    Sleep 0.5
    DoEvents ':hForm.Refresh:
Next i
End Function

Private Sub TimeDate_Click()
Text.SelText = Time & " " & Date ' Pastes the Date and time
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    'This function checks the corresponding button key and applys the apporpriate code
    ' to the text
    Case "Bold"
        If Button.Value = tbrPressed Then Text.SelBold = 1 Else Text.SelBold = 0
    Case "Italic"
        If Button.Value = tbrPressed Then Text.SelItalic = 1 Else Text.SelItalic = 0
    Case "Underline"
        If Button.Value = tbrPressed Then Text.SelUnderline = 1 Else Text.SelUnderline = 0
    Case "Left"
        Toolbar.Buttons.Item(6).Value = tbrUnpressed: Toolbar.Buttons.Item(7).Value = tbrUnpressed: Toolbar.Buttons.Item(8).Value = tbrUnpressed
        Text.SelAlignment = 0: Toolbar.Buttons.Item(6).Value = tbrPressed
    Case "Centre"
        Toolbar.Buttons.Item(6).Value = tbrUnpressed: Toolbar.Buttons.Item(7).Value = tbrUnpressed: Toolbar.Buttons.Item(8).Value = tbrUnpressed
        Text.SelAlignment = 2: Toolbar.Buttons.Item(7).Value = tbrPressed
    Case "Right"
        Toolbar.Buttons.Item(6).Value = tbrUnpressed: Toolbar.Buttons.Item(7).Value = tbrUnpressed: Toolbar.Buttons.Item(8).Value = tbrUnpressed
        Text.SelAlignment = 1: Toolbar.Buttons.Item(8).Value = tbrPressed
    Case "Bullet"
        If Text.SelBullet = True Then Text.SelBullet = 0: Button.Value = tbrUnpressed Else Text.SelBullet = 1
    Case "InsertDate"
        TimeDate_Click
    Case "FontCol"
        FontColour_Click
End Select
End Sub

Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
    'This code reads the selected colour in the drop down button on the toolbar
    ' it then uses a selection to set the corresponding colour to the selected text.
    Case Is = "vbRed": Text.SelColor = vbRed
    Case Is = "vbBlue": Text.SelColor = vbBlue
    Case Is = "vbGreen": Text.SelColor = vbGreen
    Case Is = "vbYellow": Text.SelColor = vbYellow
    Case Is = "vbCyan": Text.SelColor = vbCyan
    Case Is = "vbMagenta": Text.SelColor = vbMagenta
    Case Is = "vbBlack": Text.SelColor = vbBlack
    Case Is = "vbWhite": Text.SelColor = vbWhite
    Case Is = "Other": FontColour_Click
End Select
End Sub

Private Sub Undo_Click()
'Sends Undo request through the windows api and the common control library to the controller.
SendMessage Text.hwnd, &HC7, 0, 0&
End Sub

Private Sub Viewtoolbar_Click()
On Error Resume Next
'This shows the status bar
If Viewtoolbar.Checked = False Then ' this is done based on the tick of the option
    Toolbar.Visible = True 'shows the bar
    Viewtoolbar.Checked = True ' changes the tickbox
    Form_Resize
Else
    Toolbar.Visible = False 'hides the bar
    Viewtoolbar.Checked = False 'unchecks the option box
    Form_Resize
End If
End Sub

Private Sub WordCount_Click()
On Error Resume Next
Dim WordsCounted As Integer
Dim CountString As String
Dim TextPos As Integer

CountString = Text.Text ' Loads the text into a string variable
TextPos = 1
WordsCounted = 0 ' sets the words counted back to zero

'Checks if there is any text at all to count
If LTrim(CountString) = "" Then MsgBox "0 words present in document.", vbInformation, "Word Count": Exit Sub

'Replace the Enter, Space, Tab and NewLine With " "
CountString = VBA.Replace(CountString, Chr(32), " ")
CountString = VBA.Replace(CountString, Chr(13), " ")
CountString = VBA.Replace(CountString, Chr(10), " ")
CountString = VBA.Replace(CountString, Chr(9), " ")

CountString = Trim(CountString)

Do 'begins a word count loop
    TextPos = InStr(TextPos, CountString, " ")
    If TextPos > 0 Then
        WordsCounted = WordsCounted + 1 'undates the current spaces found (words)
        'Skips coming spaces from current position in text.
        While Mid(CountString, TextPos, 1) = " "
            TextPos = TextPos + 1
        Wend
    End If
Loop Until TextPos < "1" ' post test loop is used so that the word count is done after the last check

'tells the user of the word count though a msg box
MsgBox WordsCounted + 1 & " words present in document.", vbInformation, "Word Count"
End Sub

Private Sub WordWrap_Click()
'changes the wordwrap option of the textbox based on the checkstatus of the option
If WordWrap.Checked = False Then Text.RightMargin = 0: WordWrap.Checked = True: Exit Sub
If WordWrap.Checked = True Then Text.RightMargin = 10000: WordWrap.Checked = False: Exit Sub
End Sub

Public Function CheckSavedChanges() As Integer
'checks for changes in the document based on checksum
If SavedChangesChecksum = GetFileChecksum Then GoTo CreateNewDocument

'Displays the save changes dialog to the user.
MsgBoxReturn = MsgBox("The text in the " & FilePath & " file is unsaved." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbExclamation + vbYesNoCancel, "Semipad")

'Processed the user's response to the dialog.
If MsgBoxReturn = vbNo Then
    'It does not save the document hence it will skip this and go straight to the last code batch
ElseIf MsgBoxReturn = vbYes Then
    'Saves the document before a new one is created
    If FilePath = "Untitled" Then
    SaveAs_Click
    Else: Text.SaveFile FilePath
    End If
Else
    'Cancels the new document process and does not execute the last batch of code
    CheckSavedChanges = -1
    Exit Function
End If

CreateNewDocument:
'Creates a new document
Text.Text = ""
mainform.Caption = "Untitled - Semipad"
DocumentTitle = "Untitled"
FilePath = "Untitled"
SavedChangesChecksum = GetFileChecksum
End Function

Public Function GetFileChecksum() As Double
Dim y As Byte, z As Byte
On Error GoTo UsePrimarySum:
'this function creates a rough checksum of the document in order to detect changes
If Len(Text.Text) <= 2 Then GetFileChecksum = 0: Exit Function

'this generates a checksum based on the character ascii code of the last character
' and second last character of the text.
Text.SelStart = Len(Text.Text) - 1: Text.SelLength = 2: y = Asc(Text.SelText)
Text.SelStart = Len(Text.Text) - 2: Text.SelLength = 2: z = Asc(Text.SelText)
Text.SelStart = 0
GetFileChecksum = (Len(Text.Text) * y) / z

Exit Function

UsePrimarySum:
GetFileChecksum = Len(Text.Text)
End Function
