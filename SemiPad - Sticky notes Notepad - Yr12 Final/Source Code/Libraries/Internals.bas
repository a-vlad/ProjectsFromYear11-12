Attribute VB_Name = "Internals"
'Declares API functions that will be called later in the project
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Public Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PageSetupDlg) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetVersion Lib "kernel32" () As Long

'Declares DATA structures that will be needed later in the project
Public Type GENERALINPUT
  dwType As Long
  xi(0 To 23) As Byte
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type PageSetupDlg
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type


'This module is used as a trasnfere brindge for internal values which cannot be
'set between an array of notes and their respective array of option boxes
Public Type AlarmData
    Day As Byte
    Month As Byte
    Year As Integer
    hour As Byte
    min As String
    part As String
    Set As Boolean
End Type

Public Type NoteSettings
    nsTransparancyEnabled As Boolean
    nsTransparancyLevel As Integer
    nsInvisibleBackground As Boolean
    nsAlwaysOnTop As Boolean
    nsLockEdit As Boolean
    nsRandomBorderColour As Boolean
End Type

'This creates a public array of 30 note forms! this is required in order to
'be able to mnipulate the notes from a central origin. This will be used for
'note save feature in a loop which will take each note at a time and save the
'data in it.Changing the index will increase the maximum number of notes allowed
Public NoteInstance(29) As New PostedNote
Public Alarm(29) As AlarmData
Public NoteSetting(29) As NoteSettings
Public Notes As Integer 'This is an index of the current posted note in the array
Public Saved As Boolean
Public FilePath As String
Public DocumentTitle As String
Public SavedChangesChecksum As Double
Public PageSetupDialog As PageSetupDlg

Public Sub Main()
ThemeXP
Load mainform
mainform.Show
End Sub
