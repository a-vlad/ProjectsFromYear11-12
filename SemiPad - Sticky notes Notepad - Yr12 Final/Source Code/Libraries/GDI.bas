Attribute VB_Name = "Render"
'[==============================================]
'[RENDER Function Library                       ]
'[Based on the Windows GDI MSIMG32 Libraries    ]
'[Fully coded and written by VLAD PARASCHIV     ]
'[No external code or clipings were used        ]
'[                                              ]
'[All UI code is under the Software Name:       ]
'[WaterColour Graphics Engine coded by Vlad P.  ]
'[                                              ]
'[                                 versio 0.9   ]
'[______________________________________________]

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long

Private Declare Function SetMenuInfo Lib "user32" (ByVal hmenu As Long, mi As MENUINFO) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long

Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Type BlendFunction
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private Type MENUINFO
   cbSize As Long
   fMask As Long
   dwStyle As Long
   cyMax As Long
   hbrBack As Long
   dwContextHelpID As Long
   dwMenuData As Long
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type ChooseColor
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const MIM_APPLYTOSUBMENUS As Long = &H80000000
Private Const MIM_BACKGROUND As Long = &H2
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const ICC_USEREX_CLASSES = &H200
Const AC_SRC_OVER = &H0

Public Function ThemeXP() As Boolean
'this code is used to apply the windows xp theme, it uses windows api to theme the form
'i got this from the interent
Dim ControlTag As tagInitCommonControlsEx
    ControlTag.lngSize = LenB(ControlTag)
    ControlTag.lngICC = ICC_USEREX_CLASSES
    
InitCommonControlsEx ControlTag
End Function

Public Function TranslucentForm(Form, TranslucenceLevel As Byte)
On Error Resume Next
'checks if the windows version is compatible
If Not GetWinVer = 2000 Then Exit Function
'renders the semipatrasparent to the form
SetWindowLong Form.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
SetLayeredWindowAttributes Form.hwnd, 0, TranslucenceLevel, LWA_ALPHA
End Function


Public Function WindowShape(Form, TransparentColour As Long)
On Error Resume Next
'checks if the windows version is compatible
If Not GetWinVer = 2000 Then Exit Function
'renders the trasparent to the form
SetWindowLong Form.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
SetLayeredWindowAttributes Form.hwnd, TransparentColour, 255, LWA_COLORKEY Or LWA_ALPHA
End Function


Public Function DragForm(Form As Object, FormToMove As Object)
Dim curPath As String
curPath = App.Path & "\Graphics\hmove.cur"

'this code sets the cursor for the form to drag
Form.MouseIcon = LoadPicture(curPath, 0, 0, 0, 0)
Form.MousePointer = 99

'this code enables form dragging to move through windows api
ReleaseCapture
SendMessage FormToMove.hwnd, &HA1, 2, 0&
End Function


Public Function GetWinVer() As Integer
Dim verDetails As OSVERSIONINFOEX
'checks the windows version currently being used
verDetails.dwOSVersionInfoSize = Len(verDetails)
GetVersionEx verDetails

'the code identifies windows as compatible or not for the special effects
If verDetails.dwMajorVersion = "5" Then GetWinVer = "2000": Exit Function
If verDetails.dwMajorVersion = "4" And verDetails.dwPlatformId = "2" Then GetWinVer = "4.0": Exit Function
If verDetails.dwMajorVersion = "4" And verDetails.dwMinorVersion = "0" Then GetWinVer = "98": Exit Function

GetWinVer = "0"
End Function

Public Function AlwaysOnTop(Form As Form, Status As Boolean)
'based on the status it renders the form to always on top using the windows api call
If Status = True Then SetWindowPos Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
If Status = False Then SetWindowPos Form.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Function

Public Function ReleaseWindow(hwnd As Long)
'This function checks if the OS is compatible, if it is it proceeds
If Not GetWinVer = 2000 Then Exit Function
'this destroys the last 64 bits of the header
SetWindowLong hwnd, GWL_EXSTYLE, 0
End Function

Public Function CapWindow(Form As Object)
'This function checks if the OS is compatible, if it is it proceeds
If Not GetWinVer = 2000 Then Exit Function
'Caps the window, this adds 64 bits of extra handle data to the window handle
'this enables special effects to the window
SetWindowLong Form.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
SetLayeredWindowAttributes Form.hwnd, 0, 0, 0
End Function

Public Function TranslucenthWnd(hwnd As Long, TranslucenceLevel As Byte)
On Error Resume Next
'This function checks if the OS is compatible, if it is it proceeds
If Not GetWinVer = 2000 Then Exit Function
'makes the form semi transparent through windows api call
SetWindowLong hwnd, GWL_EXSTYLE, WS_EX_LAYERED
SetLayeredWindowAttributes hwnd, 0, TranslucenceLevel, LWA_ALPHA
End Function

Public Function ShakeForm(shake_form As Form, Number_of_Shakes As Integer)
Dim OriPos As POINTAPI
OriPos.X = shake_form.Left
OriPos.y = shake_form.Top

For i = 1 To Number_of_Shakes
    'Moves form randomly toward bottom right screen corner
    shake_form.Top = shake_form.Top + GenShakeKey
    shake_form.Left = shake_form.Left + GenShakeKey
       
    'Moves form randomly toward top right screen corner
    shake_form.Top = shake_form.Top - GenShakeKey
    shake_form.Left = shake_form.Left + GenShakeKey
    
    'Moves form randomly toward top left screen corner
    shake_form.Top = shake_form.Top - GenShakeKey
    shake_form.Left = shake_form.Left - GenShakeKey
    
    'Moves form randomly toward bottom left screen corner
    shake_form.Top = shake_form.Top + GenShakeKey
    shake_form.Left = shake_form.Left - GenShakeKey
    Beep
    'Slows the effect down
    Sleep 25
    DoEvents
Next i
shake_form.Left = OriPos.X: shake_form.Top = OriPos.y
End Function

Public Function GenShakeKey() As Integer
OffSet = 900
RndGen = 5

Randomize
'generates a random shake patttern through the first algorithem, the second algorithem generates a more vigurous shaking motion
GenShakeKey = Mid(Rnd(RndGen) * OffSet, 2, 2)
'GenShakeKey = Rnd(RndGen) * OffSet
End Function

