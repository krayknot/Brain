VERSION 5.00
Begin VB.Form FrmBooks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BRAIN: Books"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   ControlBox      =   0   'False
   Icon            =   "FrmBooks.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   13215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   0
      ScaleHeight     =   1200
      ScaleWidth      =   13215
      TabIndex        =   13
      Top             =   0
      Width           =   13215
      Begin VB.Image CoolBarImageBw 
         Height          =   480
         Index           =   11
         Left            =   9960
         Picture         =   "FrmBooks.frx":0ECA
         Top             =   240
         Width           =   480
      End
      Begin VB.Label CooLBarLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         Height          =   255
         Index           =   11
         Left            =   9840
         TabIndex        =   24
         Top             =   720
         Width           =   735
      End
      Begin VB.Image CoolBarImageColor 
         Height          =   480
         Index           =   11
         Left            =   9960
         Picture         =   "FrmBooks.frx":130C
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image CoolBar 
         Height          =   975
         Index           =   11
         Left            =   9720
         Top             =   120
         Width           =   975
      End
      Begin VB.Image CoolBarImageBw 
         Height          =   480
         Index           =   0
         Left            =   360
         Picture         =   "FrmBooks.frx":174E
         Top             =   240
         Width           =   480
      End
      Begin VB.Image CoolBarImageBw 
         Height          =   480
         Index           =   1
         Left            =   1320
         Picture         =   "FrmBooks.frx":1A58
         Top             =   240
         Width           =   480
      End
      Begin VB.Image CoolBarImageBw 
         Height          =   480
         Index           =   2
         Left            =   2280
         Picture         =   "FrmBooks.frx":1D62
         Top             =   240
         Width           =   480
      End
      Begin VB.Image CoolBarImageBw 
         Height          =   480
         Index           =   3
         Left            =   3240
         Picture         =   "FrmBooks.frx":206C
         Top             =   240
         Width           =   480
      End
      Begin VB.Image CoolBarImageBw 
         Height          =   480
         Index           =   4
         Left            =   4200
         Picture         =   "FrmBooks.frx":2376
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image CoolBarImageBw 
         Height          =   480
         Index           =   5
         Left            =   4200
         Picture         =   "FrmBooks.frx":2680
         Top             =   240
         Width           =   480
      End
      Begin VB.Image CoolBarImageBw 
         Height          =   480
         Index           =   6
         Left            =   6120
         Picture         =   "FrmBooks.frx":298A
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image CoolBarImageBw 
         Height          =   480
         Index           =   7
         Left            =   7080
         Picture         =   "FrmBooks.frx":2C94
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image CoolBarImageBw 
         Height          =   480
         Index           =   8
         Left            =   8040
         Picture         =   "FrmBooks.frx":2F9E
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image CoolBarImageBw 
         Height          =   480
         Index           =   9
         Left            =   9000
         Picture         =   "FrmBooks.frx":32A8
         Top             =   240
         Width           =   480
      End
      Begin VB.Image CoolBar 
         Height          =   975
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   975
      End
      Begin VB.Image CoolBar 
         Height          =   975
         Index           =   1
         Left            =   1080
         Top             =   120
         Width           =   975
      End
      Begin VB.Image CoolBar 
         Height          =   975
         Index           =   2
         Left            =   2040
         Top             =   120
         Width           =   975
      End
      Begin VB.Image CoolBar 
         Height          =   975
         Index           =   3
         Left            =   3000
         Top             =   120
         Width           =   975
      End
      Begin VB.Image CoolBar 
         Height          =   975
         Index           =   4
         Left            =   4920
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image CoolBar 
         Height          =   975
         Index           =   5
         Left            =   3960
         Top             =   120
         Width           =   975
      End
      Begin VB.Image CoolBar 
         Height          =   975
         Index           =   6
         Left            =   5880
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image CoolBar 
         Height          =   975
         Index           =   7
         Left            =   6840
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image CoolBar 
         Height          =   975
         Index           =   8
         Left            =   7800
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image CoolBar 
         Height          =   975
         Index           =   9
         Left            =   8760
         Top             =   120
         Width           =   975
      End
      Begin VB.Label CooLBarLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Back"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   735
      End
      Begin VB.Label CooLBarLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Forward"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin VB.Label CooLBarLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stop"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.Label CooLBarLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.Label CooLBarLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Home"
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label CooLBarLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.Label CooLBarLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Favorites"
         Height          =   255
         Index           =   6
         Left            =   6000
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label CooLBarLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Print"
         Height          =   255
         Index           =   7
         Left            =   6960
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label CooLBarLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Font"
         Height          =   255
         Index           =   8
         Left            =   7920
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label CooLBarLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mail"
         Height          =   255
         Index           =   9
         Left            =   8880
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.Image CoolBarImageColor 
         Height          =   480
         Index           =   0
         Left            =   360
         Picture         =   "FrmBooks.frx":35B2
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image CoolBarImageColor 
         Height          =   480
         Index           =   1
         Left            =   1320
         Picture         =   "FrmBooks.frx":38BC
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image CoolBarImageColor 
         Height          =   480
         Index           =   2
         Left            =   2280
         Picture         =   "FrmBooks.frx":3BC6
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image CoolBarImageColor 
         Height          =   480
         Index           =   3
         Left            =   3240
         Picture         =   "FrmBooks.frx":3ED0
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image CoolBarImageColor 
         Height          =   480
         Index           =   4
         Left            =   4200
         Picture         =   "FrmBooks.frx":41DA
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image CoolBarImageColor 
         Height          =   480
         Index           =   5
         Left            =   4200
         Picture         =   "FrmBooks.frx":44E4
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image CoolBarImageColor 
         Height          =   480
         Index           =   6
         Left            =   6120
         Picture         =   "FrmBooks.frx":47EE
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image CoolBarImageColor 
         Height          =   480
         Index           =   7
         Left            =   7080
         Picture         =   "FrmBooks.frx":4AF8
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image CoolBarImageColor 
         Height          =   480
         Index           =   8
         Left            =   8040
         Picture         =   "FrmBooks.frx":4E02
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image CoolBarImageColor 
         Height          =   480
         Index           =   9
         Left            =   9000
         Picture         =   "FrmBooks.frx":510C
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgBackground 
         Height          =   1200
         Left            =   0
         Picture         =   "FrmBooks.frx":5416
         Stretch         =   -1  'True
         Top             =   0
         Width           =   13215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   0
      TabIndex        =   2
      Top             =   7680
      Width           =   13215
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   12975
         TabIndex        =   3
         Top             =   240
         Width           =   12975
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   900
            Index           =   6
            Left            =   6600
            ScaleHeight     =   900
            ScaleWidth      =   975
            TabIndex        =   12
            Top             =   120
            Width           =   975
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   900
            Index           =   5
            Left            =   5520
            ScaleHeight     =   900
            ScaleWidth      =   975
            TabIndex        =   11
            Top             =   120
            Width           =   975
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   900
            Index           =   4
            Left            =   4440
            ScaleHeight     =   900
            ScaleWidth      =   975
            TabIndex        =   10
            Top             =   120
            Width           =   975
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Index           =   9
            Left            =   1200
            ScaleHeight     =   975
            ScaleWidth      =   975
            TabIndex        =   9
            Top             =   4440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Index           =   8
            Left            =   120
            ScaleHeight     =   975
            ScaleWidth      =   975
            TabIndex        =   8
            Top             =   4440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   900
            Index           =   3
            Left            =   3360
            ScaleHeight     =   900
            ScaleWidth      =   975
            TabIndex        =   7
            Top             =   120
            Width           =   975
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   900
            Index           =   2
            Left            =   2280
            ScaleHeight     =   900
            ScaleWidth      =   975
            TabIndex        =   6
            Top             =   120
            Width           =   975
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   900
            Index           =   1
            Left            =   1200
            ScaleHeight     =   900
            ScaleWidth      =   975
            TabIndex        =   5
            Top             =   120
            Width           =   975
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   900
            Index           =   0
            Left            =   120
            ScaleHeight     =   900
            ScaleWidth      =   975
            TabIndex        =   4
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   2400
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   13215
      Begin VB.Timer Timer3 
         Interval        =   1
         Left            =   9360
         Top             =   5760
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   6
      Left            =   3720
      Picture         =   "FrmBooks.frx":2F258
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   5
      Left            =   3240
      Picture         =   "FrmBooks.frx":354E2
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   4
      Left            =   2640
      Picture         =   "FrmBooks.frx":3B76C
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   2040
      Picture         =   "FrmBooks.frx":3C636
      Top             =   9360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   720
      Index           =   2
      Left            =   1440
      Picture         =   "FrmBooks.frx":51798
      Top             =   9360
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   840
      Picture         =   "FrmBooks.frx":52662
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   720
      Index           =   0
      Left            =   120
      Picture         =   "FrmBooks.frx":5352C
      Top             =   9360
      Width           =   720
   End
End
Attribute VB_Name = "FrmBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PrevButton

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim IntIndex As Integer

Dim lpPoint As POINTAPI, mHwnd As Long, lHwnd As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long


Private Sub Form_Load()
FrmSelection.Skin1.ApplySkin Me.hwnd
Dir1.Path = App.Path & "\Books\"
Web.Navigate App.Path & "\books\html\1.htm"

Me.Top = (Screen.Height - Me.Height) \ 2
Me.Left = (Screen.Width - Me.Width) \ 2

End Sub

Private Sub Picture9_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
   Web.Navigate App.Path & "\books\html\1.htm"
ElseIf Index = 1 Then
       Web.Navigate App.Path & "\books\win32\1.htm"
ElseIf Index = 2 Then
       Web.Navigate App.Path & "\books\vb\1.htm"
ElseIf Index = 3 Then
       Web.Navigate App.Path & "\books\crack\1.htm"
ElseIf Index = 4 Then
       Web.Navigate App.Path & "\books\database\1.htm"
ElseIf Index = 5 Then
       Web.Navigate App.Path & "\books\c\1.htm"
ElseIf Index = 6 Then
       Web.Navigate App.Path & "\books\cyber\1.htm"
End If

End Sub

Private Sub Picture9_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
IntIndex = Index
End Sub
Private Sub Timer3_Timer()
Dim i
GetCursorPos lpPoint
mHwnd = WindowFromPoint(lpPoint.X, lpPoint.Y)

If mHwnd = Picture9(IntIndex).hwnd Then
   Picture9(IntIndex).Cls
   Picture9(IntIndex).PaintPicture Image2(IntIndex).Picture, 250, 250, 600, 600
   Picture9(IntIndex).CurrentX = 100
   Picture9(IntIndex).CurrentY = 50
   Picture9(IntIndex).Font.Name = "Tahoma"
   Picture9(IntIndex).FontBold = True
   If IntIndex = 0 Then
          Picture9(IntIndex).Print "  HTML"
   ElseIf IntIndex = 1 Then
          Picture9(IntIndex).Print "Win - 32"
   ElseIf IntIndex = 2 Then
          Picture9(IntIndex).Print "    VB"
   ElseIf IntIndex = 3 Then
          Picture9(IntIndex).Print "Crack"
   ElseIf IntIndex = 4 Then
          Picture9(IntIndex).Print "DataBase"
   ElseIf IntIndex = 5 Then
          Picture9(IntIndex).Print "   C"
   ElseIf IntIndex = 6 Then
          Picture9(IntIndex).Print "Cyber"
   End If
Else
   For i = 0 To 6
     If mHwnd <> Picture9(i).hwnd Then
     Picture9(i).Cls
     Picture9(i).PaintPicture Image2(i).Picture, 220, 50, 500, 500
     Picture9(i).CurrentX = 220
     Picture9(i).CurrentY = 600
     Picture9(i).FontBold = False
     Picture9(i).Font.Name = "Tahoma"
     End If
   Next i
   
    Picture9(0).Print "  HTML"
    Picture9(1).Print "Win - 32"
    Picture9(2).Print "    VB"
    Picture9(3).Print " Crack"
    Picture9(4).Print "DataBase"
    Picture9(5).Print "   C"
    Picture9(6).Print "Cyber"
    
End If
End Sub



Sub ButtonDown(Index)

'* Make sure coolbar button doesn't resize again if user is repeatedly & quickly clicking
If CoolBar(Index).Height <> 975 Or CoolBar(Index).Width <> 975 Then
Exit Sub
End If

'* Shrink & move coolbar button to give impression of a button being pushed
CoolBarImageColor(Index).Left = CoolBarImageColor(Index).Left + 10
CoolBarImageColor(Index).Top = CoolBarImageColor(Index).Top + 10
CooLBarLabel(Index).Left = CooLBarLabel(Index).Left + 10
CooLBarLabel(Index).Top = CooLBarLabel(Index).Top + 10
CoolBar(Index).Left = CoolBar(Index).Left + 10
CoolBar(Index).Top = CoolBar(Index).Top + 10
CoolBar(Index).Height = CoolBar(Index).Height - 40
CoolBar(Index).Width = CoolBar(Index).Width - 40
    
End Sub

Sub ButtonUp(Index)

'* Make sure coolbar button doesn't resize again if user is repeatedly & quickly clicking
If CoolBar(Index).Height <> 935 Or CoolBar(Index).Width <> 935 Then
Exit Sub
End If

'* Expand & move coolbar button to give impression of a button being lifted
CoolBarImageColor(Index).Left = CoolBarImageColor(Index).Left - 10
CoolBarImageColor(Index).Top = CoolBarImageColor(Index).Top - 10
CooLBarLabel(Index).Left = CooLBarLabel(Index).Left - 10
CooLBarLabel(Index).Top = CooLBarLabel(Index).Top - 10
CoolBar(Index).Left = CoolBar(Index).Left - 10
CoolBar(Index).Top = CoolBar(Index).Top - 10
CoolBar(Index).Height = CoolBar(Index).Height + 40
CoolBar(Index).Width = CoolBar(Index).Width + 40
    
End Sub
Function MoveMouse(Index)

'* If the mouse is no longer over the same button then make the previous buttons
'* grayscale icon visible and the color icon invisible and turn it's border off
If Index <> PrevButton Then
    On Error Resume Next
        CoolBar(PrevButton).BorderStyle = 0
        CoolBarImageBw(PrevButton).Visible = True
        CoolBarImageColor(PrevButton).Visible = False
End If
    
'* If mouse has moved to another button, update Prevbutton so that if mouse is
'* moved again program will know which button to change
    PrevButton = Index
    
'* If the mouse is on this button turn on the border for this button (image)
    CoolBar(Index).BorderStyle = 1
    
'* If mouse is on this button then make the grayscale icon invisible and make
'* the color icon visible
    CoolBarImageBw(Index).Visible = False
    CoolBarImageColor(Index).Visible = True

End Function

Private Sub CoolBar_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
   Web.GoBack
ElseIf Index = 1 Then
       Web.GoForward
ElseIf Index = 2 Then
       Web.Stop
ElseIf Index = 3 Then
       Web.Refresh
ElseIf Index = 4 Then
       Web.GoHome
ElseIf Index = 5 Then
       Web.GoSearch
ElseIf Index = 9 Then
       SendNewMail "krayknot@yahoo.com", "", ""
End If

End Sub

Private Sub CoolBar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'* User pressed a mouse button
ButtonDown (Index)

End Sub

Private Sub CoolBar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'* User moved the mouse pointer
MoveMouse (Index)

End Sub

Private Sub CoolBar_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'* User released the mouse button
ButtonUp (Index)

End Sub

Private Sub CoolBarImageBw_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
   Web.GoBack
ElseIf Index = 1 Then
       Web.GoForward
ElseIf Index = 2 Then
       Web.Stop
ElseIf Index = 3 Then
       Web.Refresh
ElseIf Index = 5 Then
       Web.GoSearch
ElseIf Index = 9 Then
       SendNewMail "krayknot@yahoo.com", "", ""
End If
End Sub

Private Sub CoolBarImageBw_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'* User pressed a mouse button
ButtonDown (Index)

End Sub

Private Sub CoolBarImageBw_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'* User moved the mouse pointer
MoveMouse (Index)

End Sub


Private Sub CoolBarImageBw_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'* User released the mouse button
ButtonUp (Index)

End Sub

Private Sub CoolBarImageColor_Click(Index As Integer)
If Index = 11 Then
Unload Me
End If


End Sub

Private Sub CoolBarImageColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'* User pressed a mouse button
ButtonDown (Index)

End Sub

Private Sub CoolBarImageColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'* User moved the mouse pointer
MoveMouse (Index)

End Sub


Private Sub CoolBarImageColor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'* User released the mouse button
ButtonUp (Index)

End Sub


Private Sub CooLBarLabel_Click(Index As Integer)
If Index = 0 Then
   Web.GoBack
ElseIf Index = 1 Then
       Web.GoForward
End If
End Sub

Private Sub CooLBarLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'* User pressed a mouse button
   ButtonDown (Index)
End Sub

Private Sub CooLBarLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'* User moved the mouse pointer
MoveMouse (Index)

End Sub


Private Sub CooLBarLabel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'* User released the mouse button
ButtonUp (Index)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'* If the mouse is on the form and not on a button, then turn off the border on the previous
'* button and make the grayscale icon visible and the color icon invisible
    On Error Resume Next
    CoolBar(PrevButton).BorderStyle = 0
    CoolBarImageBw(PrevButton).Visible = True
    CoolBarImageColor(PrevButton).Visible = False
    PrevButton = -1

End Sub

Private Sub ImgBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'* If the mouse is on the background and not on a button, then turn off the border on the previous
'* button and make the grayscale icon visible and the color icon invisible
    On Error Resume Next
    CoolBar(PrevButton).BorderStyle = 0
    CoolBarImageBw(PrevButton).Visible = True
    CoolBarImageColor(PrevButton).Visible = False
    PrevButton = -1

End Sub
