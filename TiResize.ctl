VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl TiResize 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "TiResize.ctx":0000
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   2310
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1890
      Top             =   2025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3915
      Top             =   2295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3735
      Top             =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   0
      Left            =   15
      Top             =   420
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   1
      Left            =   435
      Top             =   420
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   2
      Left            =   885
      Top             =   420
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   3
      Left            =   1290
      Top             =   435
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   4
      Left            =   1725
      Top             =   435
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   5
      Left            =   2145
      Top             =   435
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   6
      Left            =   2610
      Top             =   450
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Index           =   7
      Left            =   3015
      Top             =   450
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   0
      Left            =   15
      MousePointer    =   8  'Size NW SE
      TabIndex        =   7
      Top             =   0
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   1
      Left            =   435
      MousePointer    =   7  'Size N S
      TabIndex        =   6
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   2
      Left            =   885
      MousePointer    =   6  'Size NE SW
      TabIndex        =   5
      Top             =   15
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   3
      Left            =   1290
      MousePointer    =   9  'Size W E
      TabIndex        =   4
      Top             =   45
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   4
      Left            =   1740
      MousePointer    =   8  'Size NW SE
      TabIndex        =   3
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   5
      Left            =   2145
      MousePointer    =   7  'Size N S
      TabIndex        =   2
      Top             =   0
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   6
      Left            =   2580
      MousePointer    =   6  'Size NE SW
      TabIndex        =   1
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   7
      Left            =   3015
      MousePointer    =   9  'Size W E
      TabIndex        =   0
      Top             =   30
      Width           =   360
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Edit"
      Begin VB.Menu mnuPopUpItem 
         Caption         =   "Edit Font"
         Index           =   0
      End
      Begin VB.Menu mnuPopUpItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPopUpItem 
         Caption         =   "Move To Top"
         Index           =   2
      End
      Begin VB.Menu mnuPopUpItem 
         Caption         =   "Move To Bottom"
         Index           =   3
      End
   End
End
Attribute VB_Name = "TiResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This is the first version of the TiResize Control and it is at the 80 /20 stage so it will do for now
'The code has been extracted from various sources from MSDN to PSC so thanks to the shoulders of giants and all that
'If I was going to provide move on from the 80 / 20 then the subclassing routine would be changed for message filtering
'Also the move method needs to be rehashed, as it is it works
'This is the original draft of the code so really it also could do with a spring clean and comments
'The Ocx is activated in a mouse down proc ...
'Set Me.TiResize1.ResizeControl = Text1(Index)
'Me.TiResize1.Visible = True
' Stuart Naylor Industrial Technology

Option Explicit
      'mWndProcOrg holds the original address of the
      'Window Procedure for this window. This is used to
      'route messages to the original procedure after you
      'process them.
      Private mWndProcOrg As Long
      Private IsSubClassed As Boolean
      'Handle (hWnd) of the subclassed window.
      Private mHWndSubClassed As Long
      Private mActiveControl As Object
      Private SizeMove As enumSizeMove
      Private Enum enumSizeMove
      Error = 0
      Sizing = 1
      Moving = 2
      SizeAndMove = 3
      End Enum
      Private Type CoOrds
      Left As Single
      Top As Single
      Width As Single
      Height As Single
      End Type
      Private mActiveCoOrds As CoOrds
     Private Sub Subclass()
         '-------------------------------------------------------------
         'Initiates the subclassing of this UserControl's window (hwnd).
         'Records the original WinProc of the window in mWndProcOrg.
         'Places a pointer to the object in the window's UserData area.
         '-------------------------------------------------------------

         'Exit if the window is already subclassed.
         If mWndSubClass(0).ProcessId Then Exit Sub

            'Redirect the window's messages from this control's default
            'Window Procedure to the SubWndProc function in your .BAS
            'module and record the address of the previous Window
            'Procedure for this window in mWndProcOrg.
            
            mWndSubClass(0).ProcessId = SetWindowLong(hwnd, GWL_WNDPROC, _
                                         AddressOf SubWndProc)

            'Record your window handle in case SetWindowLong gave you a
            'new one. You will need this handle so that you can unsubclass.
            mWndSubClass(0).hwnd = hwnd

            'Store a pointer to this object in the UserData section of
            'this window that will be used later to get the pointer to
            'the control based on the handle (hwnd) of the window getting
            'the message.
            Call SetWindowLong(hwnd, GWL_USERDATA, ObjPtr(Me))
            
            mWndSubClass(1).ProcessId = SetWindowLong(mActiveControl.hwnd, GWL_WNDPROC, _
                                         AddressOf SubWndProc)
            mWndSubClass(1).hwnd = mActiveControl.hwnd
            IsSubClassed = True
            
            
      End Sub

      Private Sub UnSubClass()
         '-----------------------------------------------------------
         'Unsubclasses this UserControl's window (hwnd), setting the
         'address of the Windows Procedure back to the address it was
         'at before it was subclassed.
         '-----------------------------------------------------------

         'Ensures that you don't try to unsubclass the window when
         'it is not subclassed.
         If mWndSubClass(0).ProcessId = 0 Then Exit Sub

         'Reset the window's function back to the original address.
         SetWindowLong mWndSubClass(0).hwnd, GWL_WNDPROC, mWndSubClass(0).ProcessId
         '0 Indicates that you are no longer subclassed.
         mWndSubClass(0).ProcessId = 0

         'Ensures that you don't try to unsubclass the window when
         'it is not subclassed.
         If mWndSubClass(1).ProcessId = 0 Then Exit Sub

         'Reset the window's function back to the original address.
         SetWindowLong mWndSubClass(1).hwnd, GWL_WNDPROC, mWndSubClass(1).ProcessId
         '0 Indicates that you are no longer subclassed.
         mWndSubClass(1).ProcessId = 0
      End Sub

      Friend Function WindowProc(ByVal hwnd As Long, _
         ByVal uMsg As Long, ByVal wParam As Long, _
         ByVal lParam As Long) As Long
         '--------------------------------------------------------------
         'Process the window's messages that are sent to your UserControl.
         'The WindowProc function is declared as a "Friend" function so
         'that the .BAS module can call the function but the function
         'cannot be seen from outside the UserControl project.
         '--------------------------------------------------------------
        Dim Consume As Boolean
        Dim CurrenthWnd As Long
        If hwnd = mWndSubClass(1).hwnd Then
        CurrenthWnd = 1
        End If
        Select Case uMsg
        Case WM_EXITSIZEMOVE
        Debug.Print "WM_EXITSIZEMOVE", "TiResize", CurrenthWnd
        Timer1.Enabled = True
        Case WM_LBUTTONDOWN
        Debug.Print "WM_LBUTTONDOWN", "TiResize", CurrenthWnd
        Case WM_LBUTTONUP
        Timer2.Enabled = False
        Debug.Print "WM_LBUTTONUP", "TiResize", CurrenthWnd
        Case WM_SIZING
        Debug.Print "WM_SIZING", "TiResize", CurrenthWnd
        Timer2.Enabled = False
        Case WM_MOVING
        Debug.Print "WM_MOVING", "TiResize", CurrenthWnd
        Timer2.Enabled = False
        Case WM_ENTERSIZEMOVE
        Timer2.Enabled = False
        Debug.Print "WM_ENTERSIZEMOVE", "TiResize", CurrenthWnd
        SizeMove = Error
        mActiveCoOrds.Left = 0
        mActiveCoOrds.Top = 0
        mActiveCoOrds.Width = 0
        mActiveCoOrds.Height = 0
        Case WM_SIZE
        Debug.Print "WM_SIZE", "TiResize", CurrenthWnd, HiWord(lParam), LoWord(lParam)
        SizeMove = SizeMove Or Sizing
        mActiveCoOrds.Width = ScaleX(LoWord(lParam), vbPixels, vbTwips)
        mActiveCoOrds.Height = ScaleY(HiWord(lParam), vbPixels, vbTwips)
        Case WM_MOVE
        Debug.Print "WM_MOVE", "TiResize", CurrenthWnd, HiWord(lParam), LoWord(lParam)
        SizeMove = SizeMove Or Moving
        mActiveCoOrds.Left = ScaleX(LoWord(lParam), vbPixels, vbTwips)
        mActiveCoOrds.Top = ScaleY(HiWord(lParam), vbPixels, vbTwips)
        Case WM_RBUTTONDOWN
        Debug.Print "WM_RBUTTONDOWN", "TiResize", CurrenthWnd
        Consume = True
        Case WM_RBUTTONUP
        Debug.Print "WM_RBUTTONUP", "TiResize", CurrenthWnd
        Timer3.Enabled = True
        Consume = True
        Case WM_PARENTNOTIFY
        Debug.Print "WM_PARENTNOTIFY", "TiResize", CurrenthWnd
        Case WM_MOUSEACTIVATE
        Debug.Print "WM_MOUSEACTIVATE", "TiResize", CurrenthWnd
        Case WM_NOTIFY
        Debug.Print "WM_NOTIFY", "TiResize", CurrenthWnd
        Case WM_MOUSEFIRST
        Debug.Print "WM_MOUSEFIRST", "TiResize", CurrenthWnd
        Case WM_PAINT
        Debug.Print "WM_PAINT", "TiResize", CurrenthWnd
        Case WM_COMMAND
        Debug.Print "WM_COMMAND", "TiResize", CurrenthWnd
        Case WM_NCLBUTTONDOWN
        Debug.Print "WM_NCLBUTTONDOWN", "TiResize", CurrenthWnd
        Case WM_SYSCOMMAND
        Debug.Print "WM_SYSCOMMAND", "TiResize", CurrenthWnd
        Case WM_SETCURSOR
        Debug.Print "WM_SETCURSOR", "TiResize", CurrenthWnd
        Case WM_SETFOCUS
        Debug.Print "WM_SETFOCUS", "TiResize", CurrenthWnd
        Case WM_SETHOTKEY
        Debug.Print "WM_SETHOTKEY", "TiResize", CurrenthWnd
        Case WM_SETREDRAW
        Debug.Print "WM_SETREDRAW", "TiResize", CurrenthWnd
        Case WM_SETTEXT
        Debug.Print "WM_SETTEXT", "TiResize", CurrenthWnd
        Case WM_SHOWWINDOW
        Debug.Print "WM_SHOWWINDOW", "TiResize", CurrenthWnd
        Case WM_WINDOWPOSCHANGED
        Debug.Print "WM_WINDOWPOSCHANGED", "TiResize", CurrenthWnd
        Case WM_WINDOWPOSCHANGING
        Debug.Print "WM_WINDOWPOSCHANGING", "TiResize", CurrenthWnd
        Case WM_CHILDACTIVATE
        Debug.Print "WM_CHILDACTIVATE", "TiResize", CurrenthWnd
        Case WM_NCCALCSIZE
        Debug.Print "WM_NCCALCSIZE", "TiResize", CurrenthWnd, wParam, lParam
        Case WM_NCCREATE
        Debug.Print "WM_NCCREATE", "TiResize", CurrenthWnd
        Case WM_NCDESTROY
        Debug.Print "WM_NCDESTROY", "TiResize", CurrenthWnd
        Case WM_NCHITTEST
        Debug.Print "WM_NCHITTEST", "TiResize", CurrenthWnd
        Case WM_NCACTIVATE
        Debug.Print "WM_NCACTIVATE", "TiResize", CurrenthWnd
        Case WM_NCLBUTTONUP
        Debug.Print "WM_NCLBUTTONUP", "TiResize", CurrenthWnd
        Case WM_NCMBUTTONDOWN
        Debug.Print "WM_NCMBUTTONDOWN", "TiResize", CurrenthWnd
        Case WM_NCMOUSEMOVE
        Debug.Print "WM_NCMOUSEMOVE", "TiResize", CurrenthWnd
        Case WM_NCPAINT
        Debug.Print "WM_NCPAINT", "TiResize", CurrenthWnd
        Case WM_NCRBUTTONDOWN
        Debug.Print "WM_NCRBUTTONDOWN", "TiResize", CurrenthWnd
        Case WM_PRINTCLIENT
        Debug.Print "WM_PRINTCLIENT", "TiResize", CurrenthWnd
        Case WM_PRINT
        Debug.Print "WM_PRINT", "TiResize", CurrenthWnd
        Case WM_CTLCOLOREDIT
        Debug.Print "WM_CTLCOLOREDIT", "TiResize", CurrenthWnd
        Case WM_CAPTURECHANGED
        Debug.Print "WM_CAPTURECHANGED", "TiResize", CurrenthWnd
        Case WM_GETMINMAXINFO
        Debug.Print "WM_GETMINMAXINFO", "TiResize", CurrenthWnd
        Case WM_ERASEBKGND
        Debug.Print "WM_ERASEBKGND", "TiResize", CurrenthWnd
        Case Else
        Debug.Print uMsg, "TiResize", CurrenthWnd
        End Select
        
        
        
        
         'Start Demo Code: Changes the color of the UserControl each
         'time the control is clicked in design-time from red to blue
         'or from blue to red.
              'End Demo Code.

            'Forwards the window's messages that came in to the original
            'Window Procedure that handles the messages and returns
            'the result back to the SubWndProc function.
            If Not Consume Then
            Select Case hwnd
            
            Case mWndSubClass(0).hwnd
            WindowProc = CallWindowProc(mWndSubClass(0).ProcessId, hwnd, _
                          uMsg, wParam, ByVal lParam)
            Case mWndSubClass(1).hwnd
            WindowProc = CallWindowProc(mWndSubClass(1).ProcessId, hwnd, _
                          uMsg, wParam, ByVal lParam)
            Case Else
            Debug.Print ("What was that?")
            End Select
            End If
            
      End Function '

Private Sub mnuPopUpItem_Click(Index As Integer)
Select Case Index
Case 0 'EditFont
CommonDialog1.FontBold = mActiveControl.FontBold
CommonDialog1.FontItalic = mActiveControl.FontItalic
CommonDialog1.FontName = mActiveControl.FontName
CommonDialog1.FontSize = mActiveControl.FontSize
CommonDialog1.FontStrikethru = mActiveControl.FontStrikethru
CommonDialog1.FontUnderline = mActiveControl.FontUnderline
CommonDialog1.Color = mActiveControl.ForeColor

CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects
CommonDialog1.ShowFont
If CommonDialog1.CancelError = False Then
mActiveControl.FontBold = CommonDialog1.FontBold
mActiveControl.FontItalic = CommonDialog1.FontItalic
mActiveControl.FontName = CommonDialog1.FontName
mActiveControl.FontSize = CommonDialog1.FontSize
mActiveControl.FontStrikethru = CommonDialog1.FontStrikethru
mActiveControl.FontUnderline = CommonDialog1.FontUnderline
mActiveControl.ForeColor = CommonDialog1.Color
End If
Case 2 'MoveToTop
mActiveControl.ZOrder 0
Case 3 'MoveToBottom
mActiveControl.ZOrder 1
End Select
End Sub

Private Sub UserControl_Resize()
PlaceAnchors
End Sub
Private Sub Timer1_Timer()
'Exit_Size_Move
Timer1.Enabled = False
Dim mRect As RECT
If TypeName(mActiveControl) = "TiLabel" Then
Select Case SizeMove
Case Moving
mActiveControl.Move mActiveCoOrds.Left, mActiveCoOrds.Top, mActiveControl.Width, mActiveControl.Height
Debug.Print "Move Only"
Case Sizing
mActiveControl.Move mActiveControl.Left, mActiveControl.Top, mActiveCoOrds.Width, mActiveCoOrds.Height
Debug.Print "Size Only"
Case SizeAndMove
mActiveControl.Move mActiveCoOrds.Left, mActiveCoOrds.Top, mActiveCoOrds.Width, mActiveCoOrds.Height
Debug.Print "Size And Move"
End Select
End If
SizeMove = Error
mActiveCoOrds.Left = 0
mActiveCoOrds.Top = 0
mActiveCoOrds.Width = 0
mActiveCoOrds.Height = 0
CheckSameContainer
MoveUserControl
SetControlOnTop

End Sub
Private Sub Timer2_Timer()
Timer2.Enabled = False

Dim I As Long
Dim OBJ As Object
On Error Resume Next
For I = 0 To UserControl.ParentControls.Count - 1

    If UserControl.hwnd = UserControl.ParentControls.Item(I).hwnd Then
    Set OBJ = UserControl.ParentControls.Item(I)
            'OBJ.Visible = False
        Exit For
    End If
Next I
ReleaseCapture
Debug.Print "SendMessage mActiveControl.hWnd, WM_SYSCOMMAND, SC_MOVE + 2, 0"
SendMessage mActiveControl.hwnd, WM_SYSCOMMAND, SC_MOVE + 2, 0


End Sub

Public Property Set ResizeControl(ByVal ActiveControl As Object)

    On Error Resume Next

    
    If ActiveControl Is Nothing Then
    UnSubClass
    Dim I As Long
    Dim OBJ As Object
    
    For I = 0 To UserControl.ParentControls.Count - 1

    If UserControl.hwnd = UserControl.ParentControls.Item(I).hwnd Then
    Set OBJ = UserControl.ParentControls.Item(I)
        OBJ.Visible = False
        Exit For
    End If
    Next I
    Else
    Set mActiveControl = ActiveControl
    PropertyChanged "ActiveControl"
    If ActiveControl.hwnd <> mWndSubClass(1).hwnd And IsSubClassed = True Then
    UnSubClass
    IsSubClassed = False
    End If
    CheckSameContainer
    MoveUserControl
    SetControlOnTop
    Subclass
    Timer2.Enabled = True
    End If

End Property

Private Sub CheckSameContainer()
'Make sure the usercontrol is in the same container as the bound control
Dim I As Long
Dim OBJ As Object
On Error Resume Next
For I = 0 To UserControl.ParentControls.Count - 1

    If UserControl.hwnd = UserControl.ParentControls.Item(I).hwnd Then
    Set OBJ = UserControl.ParentControls.Item(I)
        Exit For
    End If
Next I
If mActiveControl Is Nothing Then Exit Sub
If OBJ Is Nothing Then Exit Sub
    If mActiveControl.Container <> OBJ.Container Then
        Set OBJ.Container = mActiveControl.Container
    End If
End Sub
Private Sub MoveUserControl()
'find this control and move it on the parent container
Dim I As Long
Dim OBJ As Object
On Error Resume Next
For I = 0 To UserControl.ParentControls.Count - 1

    If UserControl.hwnd = UserControl.ParentControls.Item(I).hwnd Then
    Set OBJ = UserControl.ParentControls.Item(I)
            OBJ.Move mActiveControl.Left - 100, mActiveControl.Top - 100, mActiveControl.Width + 200, mActiveControl.Height + 200
            OBJ.Visible = True
        Exit For
    End If
Next I
End Sub

Private Sub SetControlOnTop()
Dim I As Long
Dim OBJ As Object

On Error Resume Next
For I = 0 To UserControl.ParentControls.Count - 1
    'get this usercontrol from the parent
    If UserControl.hwnd = UserControl.ParentControls.Item(I).hwnd Then
        Set OBJ = UserControl.ParentControls.Item(I)
        OBJ.ZOrder 0
        Exit For
    End If
Next I
End Sub

Private Sub PlaceAnchors()
Label1(0).Move 0, 0, 100, 100
Shape1(0).Move 0, 0, 100, 100
Label1(1).Move (UserControl.Width / 2) - 50, 0, 100, 100
Shape1(1).Move (UserControl.Width / 2) - 50, 0, 100, 100
Label1(2).Move UserControl.Width - 100, 0, 100, 100
Shape1(2).Move UserControl.Width - 100, 0, 100, 100
Label1(3).Move UserControl.Width - 100, (UserControl.Height / 2) - 50, 100, 100
Shape1(3).Move UserControl.Width - 100, (UserControl.Height / 2) - 50, 100, 100
Label1(4).Move UserControl.Width - 100, UserControl.Height - 100, 100, 100
Shape1(4).Move UserControl.Width - 100, UserControl.Height - 100, 100, 100
Label1(5).Move (UserControl.Width / 2) - 50, UserControl.Height - 100, 100, 100
Shape1(5).Move (UserControl.Width / 2) - 50, UserControl.Height - 100, 100, 100
Label1(6).Move 0, UserControl.Height - 100, 100, 100
Shape1(6).Move 0, UserControl.Height - 100, 100, 100
Label1(7).Move 0, (UserControl.Height / 2) - 50, 100, 100
Shape1(7).Move 0, (UserControl.Height / 2) - 50, 100, 100
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim wParam As Long

If Button = 1 Then
Select Case Index
Case 0 'NorthWest
wParam = HTTOPLEFT
Case 1 'North
wParam = HTTOP
Case 2 'NorthEast
wParam = HTTOPRIGHT
Case 3 'East
wParam = HTRIGHT
Case 4 'SouthEast
wParam = HTBOTTOMRIGHT
Case 5 'South
wParam = HTBOTTOM
Case 6 'SouthWest
wParam = HTBOTTOMLEFT
Case 7 'West
wParam = HTLEFT
End Select

ReleaseCapture
SendMessage mActiveControl.hwnd, WM_NCLBUTTONDOWN, wParam, 0
End If
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
PopupMenu mnuPopUp
End Sub



Private Sub UserControl_Terminate()
UnSubClass
End Sub
