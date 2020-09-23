VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSubclass 
   BackColor       =   &H00D0D0D0&
   Caption         =   "Subclass..."
   ClientHeight    =   7710
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11625
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00404080&
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   514
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   775
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picHeader 
      Height          =   300
      Left            =   -15
      ScaleHeight     =   240
      ScaleWidth      =   8505
      TabIndex        =   11
      Top             =   0
      Width           =   8565
      Begin VB.Label lblHeader 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "######## When.. lReturn. hWnd.... uMsg.... wParam.. lParam.. Message name....... "
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -15
         TabIndex        =   12
         Top             =   -30
         UseMnemonic     =   0   'False
         Width           =   8580
      End
   End
   Begin VB.PictureBox picMsgSel 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7710
      Left            =   8550
      ScaleHeight     =   7650
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   0
      Width           =   3075
      Begin VB.CheckBox chkAfter 
         Caption         =   "After original WndProc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   9
         Top             =   3990
         Width           =   1950
      End
      Begin VB.PictureBox picOptAfter 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   150
         ScaleHeight     =   465
         ScaleWidth      =   2190
         TabIndex        =   6
         Top             =   4320
         Width           =   2190
         Begin VB.OptionButton optAfter 
            Caption         =   "All messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optAfter 
            Caption         =   "Selected messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   7
            Top             =   270
            Width           =   2175
         End
      End
      Begin VB.CheckBox chkBefore 
         Caption         =   "Before original WndProc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   5
         Top             =   150
         Width           =   2040
      End
      Begin VB.PictureBox picOptBefore 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   150
         ScaleHeight     =   465
         ScaleWidth      =   2190
         TabIndex        =   2
         Top             =   480
         Width           =   2190
         Begin VB.OptionButton optBefore 
            Caption         =   "Selected messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   4
            Top             =   270
            Width           =   2175
         End
         Begin VB.OptionButton optBefore 
            Caption         =   "All messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView lvBefore 
         Height          =   2640
         Left            =   150
         TabIndex        =   1
         Top             =   1050
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   4657
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4260
         EndProperty
      End
      Begin MSComctlLib.ListView lvAfter 
         Height          =   2640
         Left            =   150
         TabIndex        =   10
         Top             =   4875
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   4657
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4260
         EndProperty
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuItm 
         Caption         =   "Do nothing"
         Index           =   0
      End
      Begin VB.Menu mnuItm 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuItm 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================================
' Self-subclassing sample - demonstrates adding and removing individual messages, ALL_MESSAGES and
' illustrates the range of windows messages aavailable to the programmer.
'
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
' v1.1.0005 20040620 Self-subclassing version of a WinSubHook2 sample..............................

Option Explicit

'==================================================================================================
'Application declarations

'Windows message enumeration - see also ..\Utility\mMsg.bas
Private Enum eMsg
  WM_NULL = &H0
  WM_CREATE = &H1
  WM_DESTROY = &H2
  WM_MOVE = &H3
  WM_SIZE = &H5
  WM_ACTIVATE = &H6
  WM_SETFOCUS = &H7
  WM_KILLFOCUS = &H8
  WM_ENABLE = &HA
  WM_SETREDRAW = &HB
  WM_SETTEXT = &HC
  WM_GETTEXT = &HD
  WM_GETTEXTLENGTH = &HE
  WM_PAINT = &HF
  WM_CLOSE = &H10
  WM_QUERYENDSESSION = &H11
  WM_QUIT = &H12
  WM_QUERYOPEN = &H13
  WM_ERASEBKGND = &H14
  WM_SYSCOLORCHANGE = &H15
  WM_ENDSESSION = &H16
  WM_SHOWWINDOW = &H18
  WM_WININICHANGE = &H1A
  WM_SETTINGCHANGE = &H1A
  WM_DEVMODECHANGE = &H1B
  WM_ACTIVATEAPP = &H1C
  WM_FONTCHANGE = &H1D
  WM_TIMECHANGE = &H1E
  WM_CANCELMODE = &H1F
  WM_SETCURSOR = &H20
  WM_MOUSEACTIVATE = &H21
  WM_CHILDACTIVATE = &H22
  WM_QUEUESYNC = &H23
  WM_GETMINMAXINFO = &H24
  WM_PAINTICON = &H26
  WM_ICONERASEBKGND = &H27
  WM_NEXTDLGCTL = &H28
  WM_SPOOLERSTATUS = &H2A
  WM_DRAWITEM = &H2B
  WM_MEASUREITEM = &H2C
  WM_DELETEITEM = &H2D
  WM_VKEYTOITEM = &H2E
  WM_CHARTOITEM = &H2F
  WM_SETFONT = &H30
  WM_GETFONT = &H31
  WM_SETHOTKEY = &H32
  WM_GETHOTKEY = &H33
  WM_QUERYDRAGICON = &H37
  WM_COMPAREITEM = &H39
  WM_GETOBJECT = &H3D
  WM_COMPACTING = &H41
  WM_WINDOWPOSCHANGING = &H46
  WM_WINDOWPOSCHANGED = &H47
  WM_POWER = &H48
  WM_COPYDATA = &H4A
  WM_CANCELJOURNAL = &H4B
  WM_NOTIFY = &H4E
  WM_INPUTLANGCHANGEREQUEST = &H50
  WM_INPUTLANGCHANGE = &H51
  WM_TCARD = &H52
  WM_HELP = &H53
  WM_USERCHANGED = &H54
  WM_NOTIFYFORMAT = &H55
  WM_CONTEXTMENU = &H7B
  WM_STYLECHANGING = &H7C
  WM_STYLECHANGED = &H7D
  WM_DISPLAYCHANGE = &H7E
  WM_GETICON = &H7F
  WM_SETICON = &H80
  WM_NCCREATE = &H81
  WM_NCDESTROY = &H82
  WM_NCCALCSIZE = &H83
  WM_NCHITTEST = &H84
  WM_NCPAINT = &H85
  WM_NCACTIVATE = &H86
  WM_GETDLGCODE = &H87
  WM_SYNCPAINT = &H88
  WM_NCMOUSEMOVE = &HA0
  WM_NCLBUTTONDOWN = &HA1
  WM_NCLBUTTONUP = &HA2
  WM_NCLBUTTONDBLCLK = &HA3
  WM_NCRBUTTONDOWN = &HA4
  WM_NCRBUTTONUP = &HA5
  WM_NCRBUTTONDBLCLK = &HA6
  WM_NCMBUTTONDOWN = &HA7
  WM_NCMBUTTONUP = &HA8
  WM_NCMBUTTONDBLCLK = &HA9
  WM_KEYFIRST = &H100
  WM_KEYDOWN = &H100
  WM_KEYUP = &H101
  WM_CHAR = &H102
  WM_DEADCHAR = &H103
  WM_SYSKEYDOWN = &H104
  WM_SYSKEYUP = &H105
  WM_SYSCHAR = &H106
  WM_SYSDEADCHAR = &H107
  WM_KEYLAST = &H108
  WM_IME_STARTCOMPOSITION = &H10D
  WM_IME_ENDCOMPOSITION = &H10E
  WM_IME_COMPOSITION = &H10F
  WM_IME_KEYLAST = &H10F
  WM_INITDIALOG = &H110
  WM_COMMAND = &H111
  WM_SYSCOMMAND = &H112
  WM_TIMER = &H113
  WM_HSCROLL = &H114
  WM_VSCROLL = &H115
  WM_INITMENU = &H116
  WM_INITMENUPOPUP = &H117
  WM_MENUSELECT = &H11F
  WM_MENUCHAR = &H120
  WM_ENTERIDLE = &H121
  WM_MENURBUTTONUP = &H122
  WM_MENUDRAG = &H123
  WM_MENUGETOBJECT = &H124
  WM_UNINITMENUPOPUP = &H125
  WM_MENUCOMMAND = &H126
  WM_CTLCOLORMSGBOX = &H132
  WM_CTLCOLOREDIT = &H133
  WM_CTLCOLORLISTBOX = &H134
  WM_CTLCOLORBTN = &H135
  WM_CTLCOLORDLG = &H136
  WM_CTLCOLORSCROLLBAR = &H137
  WM_CTLCOLORSTATIC = &H138
  WM_MOUSEFIRST = &H200
  WM_MOUSEMOVE = &H200
  WM_LBUTTONDOWN = &H201
  WM_LBUTTONUP = &H202
  WM_LBUTTONDBLCLK = &H203
  WM_RBUTTONDOWN = &H204
  WM_RBUTTONUP = &H205
  WM_RBUTTONDBLCLK = &H206
  WM_MBUTTONDOWN = &H207
  WM_MBUTTONUP = &H208
  WM_MBUTTONDBLCLK = &H209
  WM_MOUSEWHEEL = &H20A
  WM_PARENTNOTIFY = &H210
  WM_ENTERMENULOOP = &H211
  WM_EXITMENULOOP = &H212
  WM_NEXTMENU = &H213
  WM_SIZING = &H214
  WM_CAPTURECHANGED = &H215
  WM_MOVING = &H216
  WM_DEVICECHANGE = &H219
  WM_MDICREATE = &H220
  WM_MDIDESTROY = &H221
  WM_MDIACTIVATE = &H222
  WM_MDIRESTORE = &H223
  WM_MDINEXT = &H224
  WM_MDIMAXIMIZE = &H225
  WM_MDITILE = &H226
  WM_MDICASCADE = &H227
  WM_MDIICONARRANGE = &H228
  WM_MDIGETACTIVE = &H229
  WM_MDISETMENU = &H230
  WM_ENTERSIZEMOVE = &H231
  WM_EXITSIZEMOVE = &H232
  WM_DROPFILES = &H233
  WM_MDIREFRESHMENU = &H234
  WM_IME_SETCONTEXT = &H281
  WM_IME_NOTIFY = &H282
  WM_IME_CONTROL = &H283
  WM_IME_COMPOSITIONFULL = &H284
  WM_IME_SELECT = &H285
  WM_IME_CHAR = &H286
  WM_IME_REQUEST = &H288
  WM_IME_KEYDOWN = &H290
  WM_IME_KEYUP = &H291
  WM_MOUSEHOVER = &H2A1
  WM_MOUSELEAVE = &H2A3
  WM_CUT = &H300
  WM_COPY = &H301
  WM_PASTE = &H302
  WM_CLEAR = &H303
  WM_UNDO = &H304
  WM_RENDERFORMAT = &H305
  WM_RENDERALLFORMATS = &H306
  WM_DESTROYCLIPBOARD = &H307
  WM_DRAWCLIPBOARD = &H308
  WM_PAINTCLIPBOARD = &H309
  WM_VSCROLLCLIPBOARD = &H30A
  WM_SIZECLIPBOARD = &H30B
  WM_ASKCBFORMATNAME = &H30C
  WM_CHANGECBCHAIN = &H30D
  WM_HSCROLLCLIPBOARD = &H30E
  WM_QUERYNEWPALETTE = &H30F
  WM_PALETTEISCHANGING = &H310
  WM_PALETTECHANGED = &H311
  WM_HOTKEY = &H312
  WM_PRINT = &H317
  WM_PRINTCLIENT = &H318
  WM_THEMECHANGED = &H31A
  WM_HANDHELDFIRST = &H358
  WM_HANDHELDLAST = &H35F
  WM_AFXFIRST = &H360
  WM_AFXLAST = &H37F
  WM_PENWINFIRST = &H380
  WM_PENWINLAST = &H38F
  WM_USER = &H400
  WM_APP = &H8000
End Enum

Private Const SW_INVALIDATE      As Long = &H2

Private Type RECT
  Left                           As Long
  Top                            As Long
  Right                          As Long
  Bottom                         As Long
End Type

Private nTxtHeight               As Long        'Height of a text line
Private nMsgNo                   As Long        'Just a message counter
Private rc                       As RECT        'Scrolling rectangle

'Api declares
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function ScrollWindowEx Lib "user32" (ByVal hWnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As Any, ByVal fuScroll As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long

'==================================================================================================
'Subclasser declarations

Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Type tSubData                                                                   'Subclass data type
  hWnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'==================================================================================================
'Subclass handler - MUST be the first Public routine in this file.

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
'Parameters:
  'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
  'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  'hWnd     - The window handle
  'uMsg     - The message number
  'wParam   - Message related data
  'lParam   - Message related data
'Notes:
  'If you really know what you're doing, it's possible to change the values of the
  'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
  'values get passed to the default handler.. and optionaly, the 'after' callback
  Dim sWhen As String

  If uMsg = eMsg.WM_PAINT Then
    'If we try to display the paint message we'll just cause another paint message... vicious circle.
    Exit Sub
  End If

  If bBefore Then
    sWhen = "Before "
  Else
    sWhen = "After  "
  End If

  Call Display(sWhen, lReturn, lng_hWnd, uMsg, wParam, lParam)
End Sub

'Display a line
Private Sub Display(ByVal sWhen As String, ByVal lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  nMsgNo = nMsgNo + 1
  
  With Me
    Call ScrollWindowEx(.hWnd, 0, -nTxtHeight, rc, rc, 0, ByVal 0&, SW_INVALIDATE)
    Call UpdateWindow(.hWnd)
    .CurrentY = .ScaleHeight - nTxtHeight
    Print FmtHex(nMsgNo) & sWhen & FmtHex(lReturn) & FmtHex(lng_hWnd) & FmtHex(uMsg) & FmtHex(wParam) & FmtHex(lParam) & GetMsgName(uMsg)
  End With
End Sub

'Return the passed Long value as a hex string with leading zeros, if required, to a width of eight characters, plus a trailing space
Private Function FmtHex(ByVal nValue As Long) As String
  FmtHex = Hex$(nValue)
  FmtHex = Right$(String$(7, "0") & FmtHex, 8) & " "
End Function

'Return the name of the passed messag number
Private Function GetMsgName(ByVal uMsg As Long) As String
  Select Case uMsg
    Case eMsg.WM_ACTIVATE:              GetMsgName = "WM_ACTIVATE"
    Case eMsg.WM_ACTIVATEAPP:           GetMsgName = "WM_ACTIVATEAPP"
    Case eMsg.WM_ASKCBFORMATNAME:       GetMsgName = "WM_ASKCBFORMATNAME"
    Case eMsg.WM_CANCELJOURNAL:         GetMsgName = "WM_CANCELJOURNAL"
    Case eMsg.WM_CANCELMODE:            GetMsgName = "WM_CANCELMODE"
    Case eMsg.WM_CAPTURECHANGED:        GetMsgName = "WM_CAPTURECHANGED"
    Case eMsg.WM_CHANGECBCHAIN:         GetMsgName = "WM_CHANGECBCHAIN"
    Case eMsg.WM_CHAR:                  GetMsgName = "WM_CHAR"
    Case eMsg.WM_CHARTOITEM:            GetMsgName = "WM_CHARTOITEM"
    Case eMsg.WM_CHILDACTIVATE:         GetMsgName = "WM_CHILDACTIVATE"
    Case eMsg.WM_CLEAR:                 GetMsgName = "WM_CLEAR"
    Case eMsg.WM_CLOSE:                 GetMsgName = "WM_CLOSE"
    Case eMsg.WM_COMMAND:               GetMsgName = "WM_COMMAND"
    Case eMsg.WM_COMPACTING:            GetMsgName = "WM_COMPACTING"
    Case eMsg.WM_COMPAREITEM:           GetMsgName = "WM_COMPAREITEM"
    Case eMsg.WM_COPY:                  GetMsgName = "WM_COPY"
    Case eMsg.WM_COPYDATA:              GetMsgName = "WM_COPYDATA"
    Case eMsg.WM_CREATE:                GetMsgName = "WM_CREATE"
    Case eMsg.WM_CTLCOLORBTN:           GetMsgName = "WM_CTLCOLORBTN"
    Case eMsg.WM_CTLCOLORDLG:           GetMsgName = "WM_CTLCOLORDLG"
    Case eMsg.WM_CTLCOLOREDIT:          GetMsgName = "WM_CTLCOLOREDIT"
    Case eMsg.WM_CTLCOLORLISTBOX:       GetMsgName = "WM_CTLCOLORLISTBOX"
    Case eMsg.WM_CTLCOLORMSGBOX:        GetMsgName = "WM_CTLCOLORMSGBOX"
    Case eMsg.WM_CTLCOLORSCROLLBAR:     GetMsgName = "WM_CTLCOLORSCROLLBAR"
    Case eMsg.WM_CTLCOLORSTATIC:        GetMsgName = "WM_CTLCOLORSTATIC"
    Case eMsg.WM_CUT:                   GetMsgName = "WM_CUT"
    Case eMsg.WM_DEADCHAR:              GetMsgName = "WM_DEADCHAR"
    Case eMsg.WM_DELETEITEM:            GetMsgName = "WM_DELETEITEM"
    Case eMsg.WM_DESTROY:               GetMsgName = "WM_DESTROY"
    Case eMsg.WM_DESTROYCLIPBOARD:      GetMsgName = "WM_DESTROYCLIPBOARD"
    Case eMsg.WM_DRAWCLIPBOARD:         GetMsgName = "WM_DRAWCLIPBOARD"
    Case eMsg.WM_DRAWITEM:              GetMsgName = "WM_DRAWITEM"
    Case eMsg.WM_DROPFILES:             GetMsgName = "WM_DROPFILES"
    Case eMsg.WM_ENABLE:                GetMsgName = "WM_ENABLE"
    Case eMsg.WM_ENDSESSION:            GetMsgName = "WM_ENDSESSION"
    Case eMsg.WM_ENTERIDLE:             GetMsgName = "WM_ENTERIDLE"
    Case eMsg.WM_ENTERMENULOOP:         GetMsgName = "WM_ENTERMENULOOP"
    Case eMsg.WM_ENTERSIZEMOVE:         GetMsgName = "WM_ENTERSIZEMOVE"
    Case eMsg.WM_ERASEBKGND:            GetMsgName = "WM_ERASEBKGND"
    Case eMsg.WM_EXITMENULOOP:          GetMsgName = "WM_EXITMENULOOP"
    Case eMsg.WM_EXITSIZEMOVE:          GetMsgName = "WM_EXITSIZEMOVE"
    Case eMsg.WM_FONTCHANGE:            GetMsgName = "WM_FONTCHANGE"
    Case eMsg.WM_GETDLGCODE:            GetMsgName = "WM_GETDLGCODE"
    Case eMsg.WM_GETFONT:               GetMsgName = "WM_GETFONT"
    Case eMsg.WM_GETHOTKEY:             GetMsgName = "WM_GETHOTKEY"
    Case eMsg.WM_GETMINMAXINFO:         GetMsgName = "WM_GETMINMAXINFO"
    Case eMsg.WM_GETTEXT:               GetMsgName = "WM_GETTEXT"
    Case eMsg.WM_GETTEXTLENGTH:         GetMsgName = "WM_GETTEXTLENGTH"
    Case eMsg.WM_HOTKEY:                GetMsgName = "WM_HOTKEY"
    Case eMsg.WM_HSCROLL:               GetMsgName = "WM_HSCROLL"
    Case eMsg.WM_HSCROLLCLIPBOARD:      GetMsgName = "WM_HSCROLLCLIPBOARD"
    Case eMsg.WM_ICONERASEBKGND:        GetMsgName = "WM_ICONERASEBKGND"
    Case eMsg.WM_IME_CHAR:              GetMsgName = "WM_IME_CHAR"
    Case eMsg.WM_IME_COMPOSITION:       GetMsgName = "WM_IME_COMPOSITION"
    Case eMsg.WM_IME_COMPOSITIONFULL:   GetMsgName = "WM_IME_COMPOSITIONFULL"
    Case eMsg.WM_IME_CONTROL:           GetMsgName = "WM_IME_CONTROL"
    Case eMsg.WM_IME_ENDCOMPOSITION:    GetMsgName = "WM_IME_ENDCOMPOSITION"
    Case eMsg.WM_IME_KEYDOWN:           GetMsgName = "WM_IME_KEYDOWN"
    Case eMsg.WM_IME_KEYLAST:           GetMsgName = "WM_IME_KEYLAST"
    Case eMsg.WM_IME_KEYUP:             GetMsgName = "WM_IME_KEYUP"
    Case eMsg.WM_IME_NOTIFY:            GetMsgName = "WM_IME_NOTIFY"
    Case eMsg.WM_IME_SELECT:            GetMsgName = "WM_IME_SELECT"
    Case eMsg.WM_IME_SETCONTEXT:        GetMsgName = "WM_IME_SETCONTEXT"
    Case eMsg.WM_IME_STARTCOMPOSITION:  GetMsgName = "WM_IME_STARTCOMPOSITION"
    Case eMsg.WM_INITDIALOG:            GetMsgName = "WM_INITDIALOG"
    Case eMsg.WM_INITMENU:              GetMsgName = "WM_INITMENU"
    Case eMsg.WM_INITMENUPOPUP:         GetMsgName = "WM_INITMENUPOPUP"
    Case eMsg.WM_KEYDOWN:               GetMsgName = "WM_KEYDOWN"
    Case eMsg.WM_KEYFIRST:              GetMsgName = "WM_KEYFIRST"
    Case eMsg.WM_KEYLAST:               GetMsgName = "WM_KEYLAST"
    Case eMsg.WM_KEYUP:                 GetMsgName = "WM_KEYUP"
    Case eMsg.WM_KILLFOCUS:             GetMsgName = "WM_KILLFOCUS"
    Case eMsg.WM_LBUTTONDBLCLK:         GetMsgName = "WM_LBUTTONDBLCLK"
    Case eMsg.WM_LBUTTONDOWN:           GetMsgName = "WM_LBUTTONDOWN"
    Case eMsg.WM_LBUTTONUP:             GetMsgName = "WM_LBUTTONUP"
    Case eMsg.WM_MBUTTONDBLCLK:         GetMsgName = "WM_MBUTTONDBLCLK"
    Case eMsg.WM_MBUTTONDOWN:           GetMsgName = "WM_MBUTTONDOWN"
    Case eMsg.WM_MBUTTONUP:             GetMsgName = "WM_MBUTTONUP"
    Case eMsg.WM_MDIACTIVATE:           GetMsgName = "WM_MDIACTIVATE"
    Case eMsg.WM_MDICASCADE:            GetMsgName = "WM_MDICASCADE"
    Case eMsg.WM_MDICREATE:             GetMsgName = "WM_MDICREATE"
    Case eMsg.WM_MDIDESTROY:            GetMsgName = "WM_MDIDESTROY"
    Case eMsg.WM_MDIGETACTIVE:          GetMsgName = "WM_MDIGETACTIVE"
    Case eMsg.WM_MDIICONARRANGE:        GetMsgName = "WM_MDIICONARRANGE"
    Case eMsg.WM_MDIMAXIMIZE:           GetMsgName = "WM_MDIMAXIMIZE"
    Case eMsg.WM_MDINEXT:               GetMsgName = "WM_MDINEXT"
    Case eMsg.WM_MDIREFRESHMENU:        GetMsgName = "WM_MDIREFRESHMENU"
    Case eMsg.WM_MDIRESTORE:            GetMsgName = "WM_MDIRESTORE"
    Case eMsg.WM_MDISETMENU:            GetMsgName = "WM_MDISETMENU"
    Case eMsg.WM_MDITILE:               GetMsgName = "WM_MDITILE"
    Case eMsg.WM_MEASUREITEM:           GetMsgName = "WM_MEASUREITEM"
    Case eMsg.WM_MENUCHAR:              GetMsgName = "WM_MENUCHAR"
    Case eMsg.WM_MENUSELECT:            GetMsgName = "WM_MENUSELECT"
    Case eMsg.WM_MOUSEACTIVATE:         GetMsgName = "WM_MOUSEACTIVATE"
    Case eMsg.WM_MOUSEMOVE:             GetMsgName = "WM_MOUSEMOVE"
    Case eMsg.WM_MOUSEWHEEL:            GetMsgName = "WM_MOUSEWHEEL"
    Case eMsg.WM_MOVE:                  GetMsgName = "WM_MOVE"
    Case eMsg.WM_MOVING:                GetMsgName = "WM_MOVING"
    Case eMsg.WM_NCACTIVATE:            GetMsgName = "WM_NCACTIVATE"
    Case eMsg.WM_NCCALCSIZE:            GetMsgName = "WM_NCCALCSIZE"
    Case eMsg.WM_NCCREATE:              GetMsgName = "WM_NCCREATE"
    Case eMsg.WM_NCDESTROY:             GetMsgName = "WM_NCDESTROY"
    Case eMsg.WM_NCHITTEST:             GetMsgName = "WM_NCHITTEST"
    Case eMsg.WM_NCLBUTTONDBLCLK:       GetMsgName = "WM_NCLBUTTONDBLCLK"
    Case eMsg.WM_NCLBUTTONDOWN:         GetMsgName = "WM_NCLBUTTONDOWN"
    Case eMsg.WM_NCLBUTTONUP:           GetMsgName = "WM_NCLBUTTONUP"
    Case eMsg.WM_NCMBUTTONDBLCLK:       GetMsgName = "WM_NCMBUTTONDBLCLK"
    Case eMsg.WM_NCMBUTTONDOWN:         GetMsgName = "WM_NCMBUTTONDOWN"
    Case eMsg.WM_NCMBUTTONUP:           GetMsgName = "WM_NCMBUTTONUP"
    Case eMsg.WM_NCMOUSEMOVE:           GetMsgName = "WM_NCMOUSEMOVE"
    Case eMsg.WM_NCPAINT:               GetMsgName = "WM_NCPAINT"
    Case eMsg.WM_NCRBUTTONDBLCLK:       GetMsgName = "WM_NCRBUTTONDBLCLK"
    Case eMsg.WM_NCRBUTTONDOWN:         GetMsgName = "WM_NCRBUTTONDOWN"
    Case eMsg.WM_NCRBUTTONUP:           GetMsgName = "WM_NCRBUTTONUP"
    Case eMsg.WM_NEXTDLGCTL:            GetMsgName = "WM_NEXTDLGCTL"
    Case eMsg.WM_NULL:                  GetMsgName = "WM_NULL"
    Case eMsg.WM_PAINT:                 GetMsgName = "WM_PAINT"
    Case eMsg.WM_PAINTCLIPBOARD:        GetMsgName = "WM_PAINTCLIPBOARD"
    Case eMsg.WM_PAINTICON:             GetMsgName = "WM_PAINTICON"
    Case eMsg.WM_PALETTECHANGED:        GetMsgName = "WM_PALETTECHANGED"
    Case eMsg.WM_PALETTEISCHANGING:     GetMsgName = "WM_PALETTEISCHANGING"
    Case eMsg.WM_PARENTNOTIFY:          GetMsgName = "WM_PARENTNOTIFY"
    Case eMsg.WM_PASTE:                 GetMsgName = "WM_PASTE"
    Case eMsg.WM_PENWINFIRST:           GetMsgName = "WM_PENWINFIRST"
    Case eMsg.WM_PENWINLAST:            GetMsgName = "WM_PENWINLAST"
    Case eMsg.WM_POWER:                 GetMsgName = "WM_POWER"
    Case eMsg.WM_QUERYDRAGICON:         GetMsgName = "WM_QUERYDRAGICON"
    Case eMsg.WM_QUERYENDSESSION:       GetMsgName = "WM_QUERYENDSESSION"
    Case eMsg.WM_QUERYNEWPALETTE:       GetMsgName = "WM_QUERYNEWPALETTE"
    Case eMsg.WM_QUERYOPEN:             GetMsgName = "WM_QUERYOPEN"
    Case eMsg.WM_QUEUESYNC:             GetMsgName = "WM_QUEUESYNC"
    Case eMsg.WM_QUIT:                  GetMsgName = "WM_QUIT"
    Case eMsg.WM_RBUTTONDBLCLK:         GetMsgName = "WM_RBUTTONDBLCLK"
    Case eMsg.WM_RBUTTONDOWN:           GetMsgName = "WM_RBUTTONDOWN"
    Case eMsg.WM_RBUTTONUP:             GetMsgName = "WM_RBUTTONUP"
    Case eMsg.WM_RENDERALLFORMATS:      GetMsgName = "WM_RENDERALLFORMATS"
    Case eMsg.WM_RENDERFORMAT:          GetMsgName = "WM_RENDERFORMAT"
    Case eMsg.WM_SETCURSOR:             GetMsgName = "WM_SETCURSOR"
    Case eMsg.WM_SETFOCUS:              GetMsgName = "WM_SETFOCUS"
    Case eMsg.WM_SETFONT:               GetMsgName = "WM_SETFONT"
    Case eMsg.WM_SETHOTKEY:             GetMsgName = "WM_SETHOTKEY"
    Case eMsg.WM_SETREDRAW:             GetMsgName = "WM_SETREDRAW"
    Case eMsg.WM_SETTEXT:               GetMsgName = "WM_SETTEXT"
    Case eMsg.WM_SHOWWINDOW:            GetMsgName = "WM_SHOWWINDOW"
    Case eMsg.WM_SIZE:                  GetMsgName = "WM_SIZE"
    Case eMsg.WM_SIZECLIPBOARD:         GetMsgName = "WM_SIZECLIPBOARD"
    Case eMsg.WM_SIZING:                GetMsgName = "WM_SIZING"
    Case eMsg.WM_SPOOLERSTATUS:         GetMsgName = "WM_SPOOLERSTATUS"
    Case eMsg.WM_SYSCHAR:               GetMsgName = "WM_SYSCHAR"
    Case eMsg.WM_SYSCOLORCHANGE:        GetMsgName = "WM_SYSCOLORCHANGE"
    Case eMsg.WM_SYSCOMMAND:            GetMsgName = "WM_SYSCOMMAND"
    Case eMsg.WM_SYSDEADCHAR:           GetMsgName = "WM_SYSDEADCHAR"
    Case eMsg.WM_SYSKEYDOWN:            GetMsgName = "WM_SYSKEYDOWN"
    Case eMsg.WM_SYSKEYUP:              GetMsgName = "WM_SYSKEYUP"
    Case eMsg.WM_TIMECHANGE:            GetMsgName = "WM_TIMECHANGE"
    Case eMsg.WM_TIMER:                 GetMsgName = "WM_TIMER"
    Case eMsg.WM_UNDO:                  GetMsgName = "WM_UNDO"
    Case eMsg.WM_USER:                  GetMsgName = "WM_USER"
    Case eMsg.WM_VKEYTOITEM:            GetMsgName = "WM_VKEYTOITEM"
    Case eMsg.WM_VSCROLL:               GetMsgName = "WM_VSCROLL"
    Case eMsg.WM_VSCROLLCLIPBOARD:      GetMsgName = "WM_VSCROLL"
    Case eMsg.WM_WINDOWPOSCHANGED:      GetMsgName = "WM_WINDOWPOSCHANGED"
    Case eMsg.WM_WINDOWPOSCHANGING:     GetMsgName = "WM_WINDOWPOSCHANGING"
    Case eMsg.WM_WININICHANGE:          GetMsgName = "WM_WININICHANGE"
    Case Else:                          GetMsgName = FmtHex(uMsg)
  End Select
End Function

'Enable/disable *after* subclassing
Private Sub chkAfter_Click()
  If chkAfter = 0 Then
    optAfter(0).Enabled = False
    optAfter(0).Value = False
    optAfter(1).Enabled = False
    optAfter(1).Value = False
    lvAfter.Enabled = False
    lvAfter.TextBackground = lvwTransparent
    Call Deselect(lvAfter)
    Call Subclass_DelMsg(Me.hWnd, ALL_MESSAGES)
  Else
    optAfter(0).Enabled = True
    optAfter(0).Value = False
    optAfter(1).Enabled = True
    optAfter(1).Value = False
  End If
End Sub

'Enable/disable *before* subclassing
Private Sub chkBefore_Click()
  If chkBefore.Value = 0 Then
    optBefore(0).Enabled = False
    optBefore(0).Value = False
    optBefore(1).Enabled = False
    optBefore(1).Value = False
    lvBefore.Enabled = False
    lvBefore.TextBackground = lvwTransparent
    Call Deselect(lvBefore)
    Call Subclass_DelMsg(Me.hWnd, ALL_MESSAGES, MSG_BEFORE)
  Else
    optBefore(0).Enabled = True
    optBefore(0).Value = False
    optBefore(1).Enabled = True
    optBefore(1).Value = False
  End If
End Sub

'Uncheck all of the messages in the passed listview
Private Sub Deselect(ByVal lv As ListView)
  Dim itm As MSComctlLib.ListItem

  For Each itm In lv.ListItems
    itm.Checked = False
  Next itm
End Sub

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_Load()
  Dim i As Long
  Dim s As String

  With Me
    nTxtHeight = .TextHeight("My")
    rc.Top = .picHeader.Height
    Call .Move(.Left, .Top, 11745, 7965)
  End With

  'Populate the listview controls
  For i = 0 To &H400
    s = GetMsgName(i)
    If Asc(Left$(s, 1)) <> 48 Then
      Call lvBefore.ListItems.Add(, "k" & i, s)
      Call lvAfter.ListItems.Add(, "k" & i, s)
    End If
  Next i

  lvBefore.Sorted = True
  lvAfter.Sorted = True

  lvBefore.ColumnHeaders(1).Width = 2430#
  lvAfter.ColumnHeaders(1).Width = 2430#

  Call Subclass_Start(Me.hWnd)
End Sub

Private Sub Form_Resize()
  With Me
    If .WindowState <> vbMinimized Then
      .picHeader.Width = .ScaleWidth - .picMsgSel.Width + 1#
      .lblHeader.Width = .picHeader.ScaleWidth + 30#
      rc.Right = .picMsgSel.Left - 1
      rc.Bottom = .ScaleHeight
    End If
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call Subclass_Stop(Me.hWnd)
End Sub

'After list check box set/unset
Private Sub lvAfter_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  Dim nMsg As eMsg

  nMsg = Val(Mid$(Item.Key, 2))

  If Item.Checked Then
    Call Subclass_AddMsg(Me.hWnd, nMsg)
  Else
    Call Subclass_DelMsg(Me.hWnd, Val(Mid$(Item.Key, 2)))
  End If

  If nMsg = eMsg.WM_MOUSEWHEEL Then
    'The mousewheel events will be captured/stolen  by the listview, so set the focus elsewhere
    Call chkAfter.SetFocus
  End If
End Sub

'Before list check box set/unset
Private Sub lvBefore_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  Dim nMsg As eMsg

  nMsg = Val(Mid$(Item.Key, 2))

  If Item.Checked Then
    Call Subclass_AddMsg(Me.hWnd, nMsg, MSG_BEFORE)
  Else
    Call Subclass_DelMsg(Me.hWnd, nMsg, MSG_BEFORE)
  End If

  If nMsg = eMsg.WM_MOUSEWHEEL Then
    'The mousewheel events will be captured/stolen  by the listview, so set the focus elsewhere
    Call chkBefore.SetFocus
  End If
End Sub

Private Sub mnuItm_Click(Index As Integer)
  If Index = 2 Then
    Call Unload(Me)
  End If
End Sub

'After all or selected
Private Sub optAfter_Click(Index As Integer)
  If Index = 0 Then
    lvAfter.Enabled = False
    lvAfter.TextBackground = lvwTransparent
    Call Deselect(lvAfter)
    Call Subclass_AddMsg(Me.hWnd, ALL_MESSAGES)
  Else
    lvAfter.TextBackground = lvwOpaque
    lvAfter.Enabled = True
    Call Subclass_DelMsg(Me.hWnd, ALL_MESSAGES)
  End If
End Sub

'Before all or selected
Private Sub optBefore_Click(Index As Integer)
  If Index = 0 Then
    lvBefore.Enabled = False
    lvBefore.TextBackground = lvwTransparent
    Call Deselect(lvBefore)
    Call Subclass_AddMsg(Me.hWnd, ALL_MESSAGES, MSG_BEFORE)
  Else
    lvBefore.TextBackground = lvwOpaque
    lvBefore.Enabled = True
    Call Subclass_DelMsg(Me.hWnd, ALL_MESSAGES, MSG_BEFORE)
  End If
End Sub

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim i                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
           "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
           "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
           "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90F8060000C3"
           
'Convert the string from hex pairs to bytes and store in the static machine code buffer
    i = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      i = i + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hWnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
  Dim i As Long
  
  i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While i >= 0                                                                       'Iterate through each element
    With sc_aSubData(i)
      If .hWnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hWnd)                                                       'Subclass_Stop
      End If
    End With
    
    i = i - 1                                                                           'Next element
  Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hWnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hWnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hWnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
  If Not bAdd Then
    Debug.Assert False                                                                  'hWnd not found, programmer error
  End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function


