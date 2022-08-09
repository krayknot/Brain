VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStep2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BRAIN: Uninstaller"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar bar 
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Please wait while we configure your system and uninstall BRAIN"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmStep2.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmStep2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FSO As New FileSystemObject
            
Private Sub Form_Load()
On Error Resume Next
If frmStep2.Visible = False Then frmStep2.Visible = True
frmStep2.Refresh
'determines whether BRAIN has already uninstalled or not
'Determine the location of the application folder where the BRAIN is installed
 Dim strAppPath As String
 strAppPath = App.Path

 Dim blFolder As Boolean
 Dim blFile As Boolean

 bar.Value = bar.Value + 20
 bar.Refresh
'Determine that whether the BRAIN is running or not if yes close it else close uninsaller
 Dim strProcess() As String
 Dim intCount As Integer

 ReDim Preserve strProcess(0)

      Select Case getVersion()

      Case 1 'Windows 95/98

         Dim f As Long, sname As String
         Dim hSnap As Long, proc As PROCESSENTRY32
         hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
         If hSnap = hNull Then Exit Sub
         proc.dwSize = Len(proc)
         ' Iterate through the processes
         f = Process32First(hSnap, proc)
         Do While f
           sname = StrZToStr(proc.szExeFile)
           f = Process32Next(hSnap, proc)
         Loop

      Case 2 'Windows NT

         Dim cb As Long
         Dim cbNeeded As Long
         Dim NumElements As Long
         Dim ProcessIDs() As Long
         Dim cbNeeded2 As Long
         Dim NumElements2 As Long
         Dim Modules(1 To 200) As Long
         Dim lRet As Long
         Dim ModuleName As String
         Dim nSize As Long
         Dim hProcess As Long
         Dim i As Long
         'Get the array containing the process id's for each process object
         cb = 8
         cbNeeded = 96
         Do While cb <= cbNeeded
            cb = cb * 2
            ReDim ProcessIDs(cb / 4) As Long
            lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
         Loop
         NumElements = cbNeeded / 4

         For i = 1 To NumElements
            'Get a handle to the Process
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
               Or PROCESS_VM_READ, 0, ProcessIDs(i))
            'Got a Process handle
            If hProcess <> 0 Then
                'Get an array of the module handles for the specified
                'process
                lRet = EnumProcessModules(hProcess, Modules(1), 200, _
                                             cbNeeded2)
                'If the Module Array is retrieved, Get the ModuleFileName
                If lRet <> 0 Then
                   ModuleName = Space(MAX_PATH)
                   nSize = 500
                   lRet = GetModuleFileNameExA(hProcess, Modules(1), _
                                   ModuleName, nSize)
'                   List1.AddItem Left(ModuleName, lRet)
                   ReDim Preserve strProcess(i + 1)
                   strProcess(i) = Left(ModuleName, lRet)
                End If
            End If
          'Close the handle to the process
         lRet = CloseHandle(hProcess)
         Next
      End Select

Dim intCnt As Integer
For intCnt = LBound(strProcess) To UBound(strProcess)
    If InStr(3, strProcess(intCnt), "BRAIN.exe", vbTextCompare) > 0 Then
        MsgBox "BRAIN is running. You required to close it before uninstalling.", vbCritical
        End
    End If
Next

bar.Value = bar.Value + 20
bar.Refresh
'Deletes the components inside the BRAIN folder

 If FSO.FolderExists(strAppPath & "\books") Then 'deletes the books folder
    FSO.DeleteFile strAppPath & "\books\*.*", True
    FSO.DeleteFolder strAppPath & "\books"
 End If

 If FSO.FolderExists(strAppPath & "\chars") Then 'deletes the chars folder
    FSO.DeleteFile strAppPath & "\chars\*.*", True
    FSO.DeleteFolder strAppPath & "\chars"
 End If

 If FSO.FolderExists(strAppPath & "\fonts") Then 'deletes the fonts folder
    FSO.DeleteFile strAppPath & "\fonts\*.*", True
    FSO.DeleteFolder strAppPath & "\fonts"
 End If

If FSO.FolderExists(strAppPath & "\Installation\Speech") Then 'deletes the installation folder
    FSO.DeleteFile strAppPath & "\Installation\Speech\*.*", True
    FSO.DeleteFolder strAppPath & "\Installation\Speech"
    FSO.DeleteFolder strAppPath & "\Installation"
End If

 If FSO.FolderExists(strAppPath & "\skins") Then 'deletes the skins folder
    FSO.DeleteFile strAppPath & "\skins\*.*", True
    FSO.DeleteFolder strAppPath & "\skins"
 End If

 FSO.DeleteFile strAppPath & "\brain.exe", True
 FSO.DeleteFile strAppPath & "\License.txt", True
 FSO.DeleteFile strAppPath & "\MSAGENT.EXE", True
 FSO.DeleteFile strAppPath & "\README.TXT", True
 FSO.DeleteFile strAppPath & "\SPCHAPI.EXE", True
 FSO.DeleteFile strAppPath & "\TV_ENUA.EXE", True


 bar.Value = bar.Value + 20
 bar.Refresh
'Delete the BSystems2 folder from the windows folder
 If FSO.FolderExists(WindowsDirectory & "\bsystem2\chars") Then
    FSO.DeleteFile (WindowsDirectory & "\bsystem2\chars\*.*"), True
 End If

 If FSO.FolderExists(WindowsDirectory & "\bsystem2\skins") Then
    FSO.DeleteFile (WindowsDirectory & "\bsystem2\skins\*.*"), True
 End If

 If FSO.FolderExists(WindowsDirectory & "\bsystem2\skins") Then
    FSO.DeleteFolder (WindowsDirectory & "\bsystem2\skins"), True
 End If

 If FSO.FolderExists(WindowsDirectory & "\bsystem2\chars") Then
    FSO.DeleteFolder (WindowsDirectory & "\bsystem2\chars"), True
 End If

 If FSO.FolderExists(WindowsDirectory & "\bsystem2") Then
    FSO.DeleteFolder (WindowsDirectory & "\bsystem2"), True
 End If

 bar.Value = bar.Value + 20
 bar.Refresh
'Delete the shortcut from the desktop and from programs menu
  Dim shell As New WshShell
  Dim strDesktop As String
  strDesktop = shell.SpecialFolders("Desktop") & "\BRAIN.lnk"

  If FSO.FileExists(strDesktop) Then
     FSO.DeleteFile strDesktop, True
  End If

 bar.Value = bar.Value + 20
 bar.Refresh
'Delete all the registry entries BRAIN has made
 DeleteSetting "BRAIN", "POSTINSTALL", "POSTINSTALL"
 DeleteSetting "BRAIN", "Values", "Skin"

 Dim strstart As String
 Dim strstart1 As String
 Dim lpBuff As String * 25
 Dim Ret As Long, username As String
 Dim temp As String
 Dim temp1 As String, strDemo As String
 
 Ret = GetUserName(lpBuff, 25)
 username = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
 strstart = shell.SpecialFolders(1) & "\Programs\BRAIN"
 
 temp = InStr(4, shell.SpecialFolders(1), "\", vbTextCompare)
 temp1 = Mid(shell.SpecialFolders(1), 1, temp) & username & "\Start Menu\Programs\BRAIN"
 strstart = temp1
 
 If FSO.FolderExists(strstart) Then
'    MsgBox strstart
    If FSO.FileExists(strstart & "\BRAIN.lnk") Then
        FSO.DeleteFile strstart & "\BRAIN.lnk", True
        FSO.DeleteFolder strstart, True
    End If
 End If

 If FSO.FolderExists(strDemo) Then
    If FSO.FileExists(strDemo & "\BRAIN.lnk") Then
        FSO.DeleteFile strDemo & "\BRAIN.lnk", True
        FSO.DeleteFolder strDemo, True
    End If
 End If

 Unload Me
 frmStep3.Show vbModal

End Sub

