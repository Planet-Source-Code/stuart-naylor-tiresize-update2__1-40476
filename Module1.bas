Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetBoundsRect Lib "gdi32" (ByVal Hdc As Long, lprcBounds As RECT, ByVal Flags As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetUpdateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long


     'API Declarations used for subclassing.
      Public Declare Sub CopyMemory _
         Lib "kernel32" Alias "RtlMoveMemory" _
            (pDest As Any, _
            pSrc As Any, _
            ByVal ByteLen As Long)

      Public Declare Function SetWindowLong _
         Lib "user32" Alias "SetWindowLongA" _
            (ByVal hwnd As Long, _
            ByVal nIndex As Long, _
            ByVal dwNewLong As Long) As Long

      Public Declare Function GetWindowLong _
         Lib "user32" Alias "GetWindowLongA" _
            (ByVal hwnd As Long, _
            ByVal nIndex As Long) As Long

      Public Declare Function CallWindowProc _
         Lib "user32" Alias "CallWindowProcA" _
            (ByVal lpPrevWndFunc As Long, _
            ByVal hwnd As Long, _
            ByVal Msg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long
            
        Public Declare Function ChildWindowFromPoint _
            Lib "user32" (ByVal hwnd As Long, ByVal xPoint As Long, _
            ByVal yPoint As Long) As Long


' You can find more o these (lower) in the API Viewer.  Here
' they are used only for resizing the left and right
Public Const HTCLIENT = 1
Public Const HTCAPTION = 2
Public Const HTSYSMENU = 3
Public Const HTGROWBOX = 4
Public Const HTMENU = 5
Public Const HTHSCROLL = 6
Public Const HTVSCROLL = 7
Public Const HTMINBUTTON = 8
Public Const HTMAXBUTTON = 9
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTBORDER = 18
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_SYSCOMMAND = &H112
Public Const SC_MOVE = &HF010&
Public Const SC_SIZE = &HF000&
Public Const WM_SIZING = &H214
Public Const WM_PRINTCLIENT = &H318
Public Const WM_PRINT = &H317
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CAPTURECHANGED = &H215
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_ERASEBKGND = &H14
'This message is fired whilst the window is being resized. The lParam of the message points to a RECT structure containing the desired position of the window. Any modifications you make to this rectangle are passed back to Windows, which moves or sizes the window directly to the size and position you specify.
Public Const WM_MOVING = &H216
'This message works the same way as WM_SIZING except it is fired whilst the window is being moved.
Public Const WM_ENTERSIZEMOVE = &H231
'This message is fired when your window is about to start moving or sizing.
Public Const WM_EXITSIZEMOVE = &H232
'This message is fired when a moving or sizing operation on your window has completed.
Public Const WM_SIZE = &H5
'This message is fired whenever your window has its size changed by the SetWindowPos function, for example when windows minimizes, maximizes or restores your window, or when you call a VB function which changes the size of the window.
'The sample application shows how you can subclass these messages for a window and respond correctly to them, providing the following new
Public Const WM_MOVE = &H3

      'Constants for GetWindowLong() and SetWindowLong() APIs.
        Public Const GWL_WNDPROC = (-4)
        Public Const GWL_USERDATA = (-21)
        Public Const WM_MENUSELECT = &H11F
        Public Const WM_PARENTNOTIFY = &H210
        Public Const WM_MOUSEACTIVATE = &H21
        Public Const WM_NOTIFY As Long = &H4E&
        Public Const WM_HSCROLL = &H114
        Public Const WM_VSCROLL = &H115
        Public Const NM_RCLICK = -5
        Public Const WM_LBUTTONDBLCLK = &H203
        Public Const WM_LBUTTONDOWN = &H201
        Public Const WM_LBUTTONUP = &H202
        Public Const WM_MBUTTONDBLCLK = &H209
        Public Const WM_MBUTTONDOWN = &H207
        Public Const WM_MBUTTONUP = &H208
        Public Const WM_RBUTTONDBLCLK = &H206
        Public Const WM_RBUTTONDOWN = &H204
        Public Const WM_RBUTTONUP = &H205
        Public Const WM_MOUSEFIRST = &H200
        Public Const WM_PAINT = &HF
        Public Const WM_COMMAND = &H111
Public Const WM_SETCURSOR = &H20
Public Const WM_SETFOCUS = &H7
Public Const WM_SETHOTKEY = &H32
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_SHOWWINDOW = &H18
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCHITTEST = &H84
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCPAINT = &H85
Public Const WM_NCRBUTTONDOWN = &HA4
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

    Public Type Subclass
        hwnd As Long
        ProcessId As Long
    End Type
    'Used to hold a reference to the control to call its procedure.
      'NOTE: "UserControl1" is the UserControl.Name Property at
      '      design-time of the .CTL file.
      '      ('As Object' or 'As Control' does not work)
      Dim ctlShadowControl As TiResize

      'Used as a pointer to the UserData section of a window.
      Public mWndSubClass(1) As Subclass
      
      'Used as a pointer to the UserData section of a window.
      Dim ptrObject As Long

      'The address of this function is used for subclassing.
      'Messages will be sent here and then forwarded to the
      'UserControl's WindowProc function. The HWND determines
      'to which control the message is sent.
      Public Function SubWndProc( _
         ByVal hwnd As Long, _
         ByVal Msg As Long, _
         ByVal wParam As Long, _
         ByVal lParam As Long) As Long

         On Error Resume Next

         'Get pointer to the control's VTable from the
         'window's UserData section. The VTable is an internal
         'structure that contains pointers to the methods and
         'properties of the control.
         ptrObject = GetWindowLong(mWndSubClass(0).hwnd, GWL_USERDATA)

         'Copy the memory that points to the VTable of our original
         'control to the shadow copy of the control you use to
         'call the original control's WindowProc Function.
         'This way, when you call the method of the shadow control,
         'you are actually calling the original controls' method.
         CopyMemory ctlShadowControl, ptrObject, 4

         'Call the WindowProc function in the instance of the UserControl.
         SubWndProc = ctlShadowControl.WindowProc(hwnd, Msg, _
            wParam, lParam)

         'Destroy the Shadow Control Copy
         CopyMemory ctlShadowControl, 0&, 4
         Set ctlShadowControl = Nothing
      End Function


Public Function HiWord(Param As Long) As Integer
  
 Dim WordHex As String
 Dim offset As Long
 
 WordHex = Hex$(Param)
 offset = Len(WordHex) - 4
 If offset > 0 Then
 HiWord = CInt("&H" & Left(WordHex, offset))
    Else
HiWord = 0
End If
End Function
Public Function LoWord(Param As Long) As Integer
  
 Dim WordHex As String
  
 WordHex = Hex$(Param)
 LoWord = CInt("&H" & Right(WordHex, 4))
End Function

Public Function MakeLong(LoWord As Integer, HiWord As Integer) As Long
'Creates a Long value using Low and High integers
'Useful when converting code from C++

  Dim nLoWord As Long
  
  If LoWord% < 0 Then
    nLoWord& = LoWord% + &H10000
  Else
    nLoWord& = LoWord%
  End If

  MakeLong& = CLng(nLoWord&) Or (HiWord% * &H10000)
End Function

Public Function MakeWord(LoByte As Byte, HiByte As Byte) As Integer
'Creates an integer value using Low and High bytes
'Useful when converting code from C++
  Dim nLoByte As Integer

  If LoByte < 0 Then
    nLoByte = LoByte + &H100
  Else
    nLoByte = LoByte
  End If

  MakeWord = CInt(nLoByte) Or (HiByte * &H100)
End Function
 



