VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const WM_USER              As Integer = &H400
Private Const SWP_NOSIZE           As Integer = &H1
Private Const SWP_NOACTIVATE       As Integer = &H10
Private Const SWP_NOMOVE           As Integer = &H2
Private Const HWND_TOPMOST         As Integer = -1
Private Const TTS_NOPREFIX         As Integer = &H2
Private Const TTF_TRANSPARENT      As Integer = &H100
Private Const TTF_CENTERTIP        As Integer = &H2
Private Const TTM_ADDTOOLA         As Integer = (WM_USER + 4)
Private Const TTM_ACTIVATE         As Integer = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA   As Integer = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH   As Integer = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR    As Integer = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR  As Integer = (WM_USER + 20)
Private Const TTM_SETTITLE         As Integer = (WM_USER + 32)
Private Const TTS_BALLOON          As Integer = &H40
Private Const TTS_ALWAYSTIP        As Integer = &H1
Private Const TTF_SUBCLASS         As Integer = &H10
Private Const CW_USEDEFAULT        As Long = &H80000000
Private Const TOOLTIPS_CLASSA      As String = "tooltips_class32"
Private Type RECT
    Left                           As Long
    Top                            As Long
    Right                          As Long
    Bottom                         As Long
End Type
Private Type TOOLINFO
    lSize                          As Long
    lFlags                         As Long
    lHwnd                          As Long
    lId                            As Long
    lpRect                         As RECT
    hInstance                      As Long
    lpStr                          As String
    lParam                         As Long
End Type
Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum
Public Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum
    '    #If False Then
    '    Private TTStandard, TTBalloon
    '    Private TTNoIcon, TTIconInfo, TTIconWarning, _
    '            TTIconError
    '    #End If
Private mvarBackColor              As Long
Private mvarTitle                  As String
Private mvarForeColor              As Long
Private mvarParentControl          As Object
Private mvarIcon                   As ttIconType
Private mvarCentered               As Boolean
Private mvarStyle                  As ttStyleEnum
Private lHwnd                      As Long
Private ti                         As TOOLINFO
Private mvarTipText                As String
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                          ByVal lpClassName As String, _
                          ByVal lpWindowName As String, _
                          ByVal dwStyle As Long, _
                          ByVal X As Long, _
                          ByVal Y As Long, _
                          ByVal nWidth As Long, _
                          ByVal nHeight As Long, _
                          ByVal hWndParent As Long, _
                          ByVal hMenu As Long, _
                          ByVal hInstance As Long, _
                          lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                          ByVal wMsg As Long, _
                          ByVal wParam As Long, _
                          lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                          ByVal hWndInsertAfter As Long, _
                          ByVal X As Long, _
                          ByVal Y As Long, _
                          ByVal cx As Long, _
                          ByVal cy As Long, _
                          ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                          lpRect As RECT) As Long
Private Sub Class_Terminate()
    If lHwnd <> 0 Then DestroyWindow lHwnd
End Sub
Public Function Create() As Boolean
  Dim lpRect    As RECT, lWinStyle As Long
    If lHwnd <> 0 Then DestroyWindow lHwnd
    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    If mvarStyle = TTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON
    If Not mvarParentControl Is Nothing Then
        lHwnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, mvarParentControl.hwnd, 0&, App.hInstance, 0&)
        SetWindowPos lHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
        GetClientRect mvarParentControl.hwnd, lpRect
        With ti
            If mvarCentered Then
                .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
              Else
                .lFlags = TTF_SUBCLASS
            End If
            .lHwnd = mvarParentControl.hwnd
            .lId = 0
            .hInstance = App.hInstance
            .lpRect = lpRect
        End With
        SendMessage lHwnd, TTM_ADDTOOLA, 0&, ti
        If mvarTitle <> vbNullString Or mvarIcon <> TTNoIcon Then _
         SendMessage lHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
        If mvarForeColor <> Empty Then _
         SendMessage lHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
        If mvarBackColor <> Empty Then _
         SendMessage lHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
    End If
End Function
Public Property Get msgTipText() As String
    msgTipText = mvarTipText
End Property
Public Property Let msgTipText(ByVal vData As String)
    ti.lpStr = vData
    If lHwnd <> 0 Then SendMessage lHwnd, TTM_UPDATETIPTEXTA, 0&, ti
End Property
Public Property Get ParentControl() As Object
    Set ParentControl = mvarParentControl
End Property
Public Property Set ParentControl(ByVal objData As Object)
    Set mvarParentControl = objData
End Property
Public Property Get ttStyle() As ttStyleEnum
    ttStyle = mvarStyle
End Property
Public Property Let ttStyle(ByVal vData As ttStyleEnum)
    mvarStyle = vData
End Property
Public Property Let ttBackColor(ByVal vData As Long)
    mvarBackColor = vData
    If lHwnd <> 0 Then _
     SendMessage lHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
End Property
Public Property Get ttBackColor() As Long
    ttBackColor = mvarBackColor
End Property
Public Property Let ttCentered(ByVal vData As Boolean)
    mvarCentered = vData
End Property
Public Property Get ttCentered() As Boolean
    ttCentered = mvarCentered
End Property
Public Property Let ttForeColor(ByVal vData As Long)
    mvarForeColor = vData
    If lHwnd <> 0 Then _
     SendMessage lHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
End Property
Public Property Get ttForeColor() As Long
    ttForeColor = mvarForeColor
End Property
Public Property Get ttIcon() As ttIconType
    ttIcon = mvarIcon
End Property
Public Property Let ttIcon(ByVal vData As ttIconType)
    mvarIcon = vData
    If lHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
     SendMessage lHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
    End If
End Property
Public Property Get ttTitle() As String
    ttTitle = ti.lpStr
End Property
Public Property Let ttTitle(ByVal vData As String)
    mvarTitle = vData
    If lHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
     SendMessage lHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
    End If
End Property






