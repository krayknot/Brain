Attribute VB_Name = "functions"
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4

Enum StyleButtons
 OkOnly = 0
 OkCancelonly = 1
 CloseOnly = 2
 YesOnly = 3
 YesNoonly = 4
End Enum

'Function to get default windows directory
 Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Function that will drops down the list portion of a combobox control whenever it receives focus
 Declare Function SendMessage Lib "user32" Alias _
  "sendmessagea" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
  
Function MESSAGE(messagetext As String, buttons As StyleButtons, Title As String)
Frmmessage.Lblmessage.Caption = messagetext
If buttons = OkOnly Then
   Frmmessage.CmdCancel.Visible = False
ElseIf buttons = OkCancelonly Then
       Frmmessage.CmdCancel.Visible = True
       Frmmessage.CmdOk.Visible = True
ElseIf buttons = CloseOnly Then
       Frmmessage.CmdCancel.Visible = True
       Frmmessage.CmdCancel.Caption = "Close"
       Frmmessage.CmdOk.Visible = False
ElseIf buttons = YesOnly Then
       Frmmessage.CmdCancel.Visible = False
       Frmmessage.CmdOk.Visible = True
       Frmmessage.CmdOk.Caption = "Yes"
ElseIf buttons = YesNoonly Then
       Frmmessage.CmdCancel.Visible = True
       Frmmessage.CmdOk.Visible = True
       Frmmessage.CmdCancel.Caption = "No"
       Frmmessage.CmdOk.Caption = "Yes"
End If
Frmmessage.Caption = Title

Frmmessage.Show vbModal
     
End Function

Function Extract(Data As String) 'Function that reverses a string
Dim t As String
Dim t1 As String
Dim i As Integer

For i = 1 To Len(Data)
    t = Right$(Data, i)
    t = Mid(t, 1, 1)
    
    If t = Trim("\") Then
       Extract = t1
       Exit Function
    Else
       t1 = t & t1
    End If
Next
End Function

Function WindowsDirectory() As String
Dim WinPath As String
    WinPath = String(145, Chr(0))
    WindowsDirectory = Left(WinPath, GetWindowsDirectory(WinPath, 145))
End Function

Function properfilename(filename)
'Extracts the file name from the full path
Dim result, temp, temp1
Dim i As Integer

'code that will extract the file name from the last in reverse order'
 For i = Len(filename) To 1 Step -1
  result = Mid(filename, i, 1)
  If result = "\" Then
   Exit For
  End If
  temp = temp & result
 Next i

'code that will again the reverse the temp variable to make the name readable
 For i = Len(temp) To 1 Step -1
  result = Mid(temp, i, 1)
  temp1 = temp1 & result
 Next i
properfilename = temp1
End Function
Public Sub SendNewMail(ByVal MailTo As String, _
                       ByVal Subject As String, _
                       ByVal Body As String)
    Dim Buff As String
    
    Buff = "mailto:" & MailTo & "?Subject=" & _
           Subject & "&Body=" & Body
    
    Call ShellExecute(0&, "Open", Buff, "", "", 1)

End Sub
