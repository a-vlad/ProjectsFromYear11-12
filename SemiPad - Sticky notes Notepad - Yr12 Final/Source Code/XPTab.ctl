VERSION 5.00
Begin VB.UserControl XPTab 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   LockControls    =   -1  'True
   ScaleHeight     =   315
   ScaleWidth      =   4065
   Begin VB.Image Image1 
      Height          =   15
      Left            =   0
      Picture         =   "XPTab.ctx":0000
      Stretch         =   -1  'True
      Top             =   300
      Width           =   4155
   End
   Begin VB.Image t 
      Height          =   285
      Left            =   1410
      Picture         =   "XPTab.ctx":004A
      Top             =   3090
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image ts 
      Height          =   315
      Left            =   195
      Picture         =   "XPTab.ctx":11C4
      Top             =   3030
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image h3 
      Height          =   45
      Left            =   2910
      Picture         =   "XPTab.ctx":2466
      Stretch         =   -1  'True
      Top             =   45
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image h2 
      Height          =   45
      Left            =   1395
      Picture         =   "XPTab.ctx":2718
      Stretch         =   -1  'True
      Top             =   45
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image h1 
      Height          =   45
      Left            =   30
      Picture         =   "XPTab.ctx":29CA
      Stretch         =   -1  'True
      Top             =   45
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      Height          =   300
      Left            =   2880
      TabIndex        =   3
      Top             =   15
      Width           =   15
   End
   Begin VB.Label l3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reminder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2910
      TabIndex        =   2
      Top             =   90
      Width           =   1050
   End
   Begin VB.Image t3 
      Height          =   270
      Left            =   2910
      Picture         =   "XPTab.ctx":2C7C
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1065
   End
   Begin VB.Label l2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Special Effects"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1395
      TabIndex        =   1
      Top             =   90
      Width           =   1485
   End
   Begin VB.Image t2 
      Height          =   270
      Left            =   1395
      Picture         =   "XPTab.ctx":3DF6
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1515
   End
   Begin VB.Label l1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Background"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   1335
   End
   Begin VB.Image t1 
      Height          =   270
      Left            =   30
      Picture         =   "XPTab.ctx":4F70
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1350
   End
End
Attribute VB_Name = "XPTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Event ClickTab1()
Public Event ClickTab2()
Public Event ClickTab3()

Private Sub h1_Click()
'calls close and open tab subfunction to render a new tab, raises a new event
OpenTab t1, l1: CloseTab t2, l2: CloseTab t3, l3: RaiseEvent ClickTab1
End Sub

Private Sub h2_Click()
'calls close and open tab subfunction to render a new tab, raises a new event
OpenTab t2, l2: CloseTab t1, l1: CloseTab t3, l3: RaiseEvent ClickTab2
End Sub

Private Sub h3_Click()
'calls close and open tab subfunction to render a new tab, raises a new event
OpenTab t3, l3: CloseTab t2, l2: CloseTab t1, l1: RaiseEvent ClickTab3
End Sub


Private Sub l1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'calls close and open tab subfunction to render a new tab mouse over effect
OpenTab t1, l1: CloseTab t2, l2: CloseTab t3, l3
'raises a new event
RaiseEvent ClickTab1
End Sub

Private Sub l1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'calls close and open tab subfunction to render a new tab mouse over effect
h1.Visible = True: h2.Visible = False: h3.Visible = False
End Sub

Private Sub l2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'calls close and open tab subfunction to render a new tab mouse over effect
OpenTab t2, l2: CloseTab t1, l1: CloseTab t3, l3: RaiseEvent ClickTab2
End Sub

Private Sub l2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'calls close and open tab subfunction to render a new tab mouse over effect
h1.Visible = False: h2.Visible = True: h3.Visible = False
End Sub

Private Sub l3_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'calls close and open tab subfunction to render a new tab mouse over effect
OpenTab t3, l3: CloseTab t2, l2: CloseTab t1, l1: RaiseEvent ClickTab3
End Sub

Private Sub l3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'calls close and open tab subfunction to render a new tab mouse over effect
h1.Visible = False: h2.Visible = False: h3.Visible = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'calls close and open tab subfunction to render a new tab mouse over effect
h1.Visible = False: h2.Visible = False: h3.Visible = False
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'calls close and open tab subfunction to render a new tab mouse over effect
RefreshTabs
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'This refreshes the mouse over effect
RefreshTabs
End Sub

Private Sub OpenTab(tTab As Image, lLabel As Label)
If tTab.Height = 315 Then Exit Sub
tTab.Picture = ts.Picture
tTab.ZOrder 0
tTab.Top = tTab.Top - (315 - tTab.Height)
tTab.Height = 315
tTab.Left = tTab.Left - 30
tTab.Width = tTab.Width + 60
lLabel.ZOrder 0
End Sub

Private Sub CloseTab(tTab As Image, lLabel As Label)
If tTab.Height = 270 Then Exit Sub
tTab.Picture = t.Picture
tTab.ZOrder 1
tTab.Top = tTab.Top + (315 - 270)
tTab.Height = 270

tTab.Left = tTab.Left + 30
tTab.Width = tTab.Width - 60
lLabel.ZOrder 0
End Sub
Public Function KillTab3()
'makes the last tab invisible
h3.Visible = False
t3.Visible = False
l3.Visible = False
End Function

Public Function ChangeTabNames(Tab1Name As String, Tab2Name As String, Tab3Name As String)
l1.Caption = Tab1Name: l2.Caption = Tab2Name: l3.Caption = Tab3Name
End Function
Public Function RefreshTabs()
h1.Visible = False: h2.Visible = False: h3.Visible = False
End Function

Public Function LoadTabs()
OpenTab t1, l1: CloseTab t2, l2: CloseTab t3, l3

frmOptions.BackgroundTab.Visible = True
frmOptions.EffectsTab.Visible = False
frmOptions.ReminderTab.Visible = False
ReleaseWindow frmOptions.hwnd
        
frmOptions.sample.Refresh: frmOptions.sample2.Refresh
        
If frmOptions.chkInvisible.Value = 1 Then frmOptions.chkInvisible_Click
End Function
