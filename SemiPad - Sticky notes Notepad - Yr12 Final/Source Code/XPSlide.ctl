VERSION 5.00
Begin VB.UserControl XPSlide 
   BackColor       =   &H00F8FAFC&
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2220
   LockControls    =   -1  'True
   ScaleHeight     =   330
   ScaleWidth      =   2220
   Begin VB.PictureBox SlideFrame 
      BackColor       =   &H00F8FAFC&
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   -105
      ScaleHeight     =   420
      ScaleWidth      =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   2400
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2125
         Picture         =   "XPSlide.ctx":0000
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   11
         TabIndex        =   1
         Top             =   0
         Width           =   165
      End
      Begin VB.Image Image1 
         Height          =   60
         Left            =   105
         Picture         =   "XPSlide.ctx":0336
         Top             =   105
         Width           =   45
      End
      Begin VB.Image sliderrail 
         Height          =   60
         Left            =   150
         Picture         =   "XPSlide.ctx":03A8
         Stretch         =   -1  'True
         Top             =   105
         Width           =   2130
      End
      Begin VB.Image Image3 
         Height          =   60
         Left            =   2280
         Picture         =   "XPSlide.ctx":040A
         Top             =   105
         Width           =   45
      End
   End
   Begin VB.Image dis_pic 
      Height          =   315
      Left            =   1620
      Picture         =   "XPSlide.ctx":047C
      Top             =   1065
      Width           =   165
   End
   Begin VB.Image norm_pic 
      Height          =   315
      Left            =   1245
      Picture         =   "XPSlide.ctx":07B2
      Top             =   1035
      Width           =   165
   End
   Begin VB.Image md_pic 
      Height          =   315
      Left            =   840
      Picture         =   "XPSlide.ctx":0AE8
      Top             =   1065
      Width           =   165
   End
   Begin VB.Image sel_pic 
      Height          =   315
      Left            =   345
      Picture         =   "XPSlide.ctx":0E1E
      Top             =   1035
      Width           =   165
   End
End
Attribute VB_Name = "XPSlide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Change()
Public Event Scroll()

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If pic.Tag <> "D" Then pic.Picture = sel_pic.Picture
End Sub

Private Sub Pic_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
pic.Tag = ""
End Sub


Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim Cur As POINTAPI, WinPos As RECT
If pic.Left > 2125 Then pic.Left = 2125: Exit Sub
If pic.Left < 100 Then pic.Left = 100: Exit Sub

pic.Picture = md_pic.Picture

pic.Tag = "D"

Do
    GetCursorPos Cur
    GetWindowRect SlideFrame.hwnd, WinPos
    
    RaiseEvent Change
    
    pic.Left = ((Cur.X - WinPos.Left) - 7) * 15.3
    pic.Refresh
    
    If pic.Left > 2125 Then pic.Left = 2124: pic.Refresh ': Exit Sub
    If pic.Left < 100 Then pic.Left = 101: pic.Refresh ': Exit Sub
    
    DoEvents
Loop Until pic.Tag = ""

pic.Picture = norm_pic.Picture
End Sub

Private Sub SlideFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If pic.Enabled = True Then pic.Picture = norm_pic.Picture
End Sub

Public Function RefreshPick()
If pic.Enabled = True Then pic.Picture = norm_pic.Picture
End Function

Public Function DisableSlide()
pic.Picture = dis_pic.Picture
pic.Enabled = False
End Function
Public Function EnableSlide()
pic.Picture = norm_pic.Picture
pic.Enabled = True
End Function

Public Function Value() As Integer
Value = Abs((pic.Left / 2125) * 100)
End Function

Public Function SetNewValue(Value As Integer)
pic.Left = Value * 21.25
End Function

Private Sub sliderrail_Click()
Dim Cur As POINTAPI, WinPos As RECT
    GetCursorPos Cur
    GetWindowRect SlideFrame.hwnd, WinPos
    RaiseEvent Scroll
    RaiseEvent Change
    pic.Left = ((Cur.X - WinPos.Left) - 7) * 15.3
End Sub

