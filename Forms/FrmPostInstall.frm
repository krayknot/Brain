VERSION 5.00
Begin VB.Form FrmPostInstall 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pre Installer of BRAIN"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4875
   ControlBox      =   0   'False
   Icon            =   "FrmPostInstall.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   3000
      Top             =   120
   End
   Begin VB.FileListBox File2 
      Height          =   1260
      Left            =   1440
      Pattern         =   "*.age"
      TabIndex        =   8
      Top             =   5040
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   3840
      TabIndex        =   7
      Top             =   5040
      Width           =   975
   End
   Begin VB.FileListBox File3 
      Height          =   1260
      Left            =   2640
      Pattern         =   "*.ski"
      TabIndex        =   6
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Exit!"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdgoinstall 
      Caption         =   "Go>>"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3480
      Top             =   120
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin Brain.UserControl1 Progressbar1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   8388608
      Scrolling       =   4
   End
   Begin VB.Label Label2 
      Caption         =   "Before execution we require to set the environment."
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Setting the BRAIN Environment [Recommended]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "FrmPostInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetVersion Lib "kernel32" () As Long

      Private Type LUID
         UsedPart As Long
         IgnoredForNowHigh32BitPart As Long
      End Type

      Private Type LUID_AND_ATTRIBUTES
         TheLuid As LUID
         Attributes As Long
      End Type

      Private Type TOKEN_PRIVILEGES
         PrivilegeCount As Long
         TheLuid As LUID
         Attributes As Long
      End Type

Dim FSO As New FileSystemObject
Dim BLGo As Boolean
Private WithEvents Huffman As ClsHuffMan
Attribute Huffman.VB_VarHelpID = -1
Dim Inttimer As Integer
Dim intChars As Integer
Dim intSkins As Integer


Private Sub CmdCLose2_Click()
Dim IntCnt As Integer
For IntCnt = Picture3.Top To 0 Step 100
    Picture3.Top = IntCnt
Next IntCnt

End Sub

Private Sub cmdclose_Click()
End
End Sub

Private Sub cmdgoinstall_Click()
 Dim strDrive As String, strPath As String, strPath1 As String
 Dim temp As TextStream, strTempSkin As TextStream
 Dim IntCount As Integer
 Dim gstrfontdir As String
 Dim dfile As String
 Dim fsoDrive As Drive
 Dim i As Integer, arem
 Dim StrDriveName As String
 

 gstrwindir = WindowsDirectory
 gstrfontdir = GetWindowsFontDir()
 List1.Clear
 Progressbar1.Value = Progressbar1.Value + 5
 List1.AddItem "Checking controls to run BRAIN: "
 List1.AddItem "Retrieving Characters Information..."
 List1.Refresh

'Checks the presence of the file and Counts the number of characters we have provided
 If Not FSO.FileExists(App.Path & "\chars\char_no.txt") Then
    List1.AddItem "Setup Aborted."
    List1.AddItem "Character information file"
    List1.AddItem "char_no.txt not found."
    Exit Sub
 End If
       
'Checks the presence of the file and Counts the number of skins we have provided
 If Not FSO.FileExists(App.Path & "\skins\skins_no.txt") Then
    List1.AddItem "Setup Aborted."
    List1.AddItem "Skins information file"
    List1.AddItem "skin_no.txt not found."
    Exit Sub
 End If
 
 Progressbar1.Value = Progressbar1.Value + 2
 
 Set temp = FSO.OpenTextFile(App.Path & "\chars\char_no.txt", ForReading) 'Opens characters file
 intChars = Val(temp.ReadLine)
 Set strTempSkin = FSO.OpenTextFile(App.Path & "\skins\skins_no.txt", ForReading) 'Opens skins file
 intSkins = Val(strTempSkin.ReadLine)
  
'code that takes the names of characters in the array from the file
 Dim StrAgentArray() As String
 Dim StrDestination() As String
 Dim StrSkinArray() As String
 Dim StrSkinDestination() As String
 Dim StrTemp As String
 Dim StrTemp1 As String
 Dim strAgentName As String
 Dim strSkinName As String
 ReDim StrAgentArray(intChars)
 ReDim StrDestination(intChars)
 ReDim StrSkinArray(intSkins)
 ReDim StrSkinDestination(intSkins)

'creates the temporary folder tempbrain in the drive strdrive
'and extract all the files their , copy them and delete them
 strDrive = Mid$(App.Path, 1, 2)
 
'Prior check if already contains the folder tempb and temps then delete them
'Deletes the character files from the Chars directory to prevent unauthorized copying
 If FSO.FolderExists(strDrive & "\tempb") Then FSO.DeleteFolder strDrive & "\tempb", True
'Deletes the skin files from the skins directory to prevent unauthorized copying
 If FSO.FolderExists(strDrive & "\temps") Then FSO.DeleteFolder strDrive & "\temps", True
 
 FSO.CreateFolder strDrive & "\tempb"
 
' MsgBox strDrive
' strPath = Trim("cd " & Mid$(App.Path, 4, Trim(Len(App.Path))) & "\CHARS")
 strPath = Trim("cd " & "tempb")
 
 'Collects the names of all the characters from the text file
 For IntCount = 1 To intChars
     strAgentName = Trim(temp.ReadLine)
     StrAgentArray(IntCount) = strDrive & "\tempb\" & strAgentName & ".age"
     StrDestination(IntCount) = gstrwindir & "\BSystem2\Chars\" & strAgentName & ".acs"
 Next
 List1.AddItem intChars & " Characters found."
 
'extracts the characters from the cab file
'makes a single batch file to operate all the things
 Open App.Path & "\chars\brain.bat" For Output As #1   ' Open file for output.
 Print #1, "cd\"   ' Print text to file.
 Print #1, strDrive
 Print #1, strPath
 Print #1, "extract.exe /y /a chars.cab *.*"   ' Print text to file.
 Close #1
 FSO.CopyFile App.Path & "\chars\*.*", strDrive & "\tempb\" 'copies file to temporary folder
 Shell strDrive & "\tempb\brain.bat", vbHide 'execute the batch file
 
 List1.AddItem "Extracting the characters..."
 List1.AddItem "Please wait..."
 List1.Selected(List1.ListCount - 1) = True
 List1.Refresh
 
File2.Path = strDrive & "\tempb"
File2.Refresh

'MsgBox "check"
Again:
  If File2.ListCount <> intChars Then
     File2.Refresh
     GoTo Again
  End If

 List1.AddItem "Characters extraction completed"
 List1.Refresh
 
 Progressbar1.Value = Progressbar1.Value + 5

'creates the temporary folder tempbrain in the drive strdrive
'and extract all the files their , copy them and delete them
 strDrive = Mid$(App.Path, 1, 2)
 FSO.CreateFolder strDrive & "\temps"
 strPath1 = Trim("cd " & "temps")
  
'Collects the names of all the skins from the text file
 For IntCount = 1 To intSkins
     strSkinName = Trim(strTempSkin.ReadLine)
     StrSkinArray(IntCount) = strDrive & "\temps\" & strSkinName & ".ski"
     StrSkinDestination(IntCount) = gstrwindir & "\BSystem2\skins\" & strSkinName & ".skn"
 Next
 List1.AddItem intSkins & " Skins found."
  
 'strDrive = Mid$(App.Path, 1, 2)
 'strPath1 = Trim("cd " & Mid$(App.Path, 4, Trim(Len(App.Path))) & "\skins")
 
 'extracts the skins from the cab file
'makes a single batch file to operate all the things
 Open App.Path & "\skins\brain.bat" For Output As #1   ' Open file for output.
 Print #1, "cd\"   ' Print text to file.
 Print #1, strDrive
 Print #1, strPath1
 Print #1, "extract.exe /y /a skins.cab *.*"   ' Print text to file.
 Close #1
 FSO.CopyFile App.Path & "\skins\*.*", strDrive & "\temps\" 'copies file to temporary folder
 Shell strDrive & "\temps\brain.bat", vbHide 'execute the batch file
 
 List1.AddItem "Extracting the characters..."
 List1.AddItem "Please wait..."
 List1.Selected(List1.ListCount - 1) = True
 List1.Refresh

 List1.AddItem "Extracting the Skins..."
 List1.AddItem "Please wait..."
 List1.Selected(List1.ListCount - 1) = True
 List1.Refresh
 
 File3.Path = strDrive & "\temps"
 File3.Refresh

Again1:
  If File3.ListCount <> intSkins Then
     File3.Refresh
     GoTo Again1
  End If
  
 List1.AddItem "Characters extraction completed"
 List1.Refresh

 Progressbar1.Value = Progressbar1.Value + 2
 
 gstrwindir = WindowsDirectory
 gstrfontdir = GetWindowsFontDir()
 
'Makes the default folder in the windows directory
 If Not FSO.FolderExists(gstrwindir & "\BSystem2") Then
    FSO.CreateFolder gstrwindir & "\BSystem2"
 End If
 
 Progressbar1.Value = Progressbar1.Value + 2

 If Not FSO.FolderExists(gstrwindir & "\BSystem2\Skins") Then
    FSO.CreateFolder gstrwindir & "\BSystem2\Skins"
 End If
 
 Progressbar1.Value = Progressbar1.Value + 2

 If Not FSO.FolderExists(gstrwindir & "\BSystem2\Chars") Then
    FSO.CreateFolder gstrwindir & "\BSystem2\Chars"
 End If
  
 Progressbar1.Value = Progressbar1.Value + 5
 
 List1.AddItem "All necessary folders created"
 List1.Selected(List1.ListCount - 1) = True
 List1.Refresh
  
 Timer1.Enabled = True
 
 For IntCount = 1 To intChars
     'Copies the animated character in the specific folder
        If FSO.FileExists(StrAgentArray(IntCount)) Then
           StrTemp = StrReverse(StrAgentArray(IntCount))
           StrTemp1 = InStr(1, StrTemp, "\", vbTextCompare)
           List1.AddItem "Installing the Agent " & StrReverse(Mid$(StrTemp, 1, Val(StrTemp1) - 1))
           FSO.CopyFile StrAgentArray(IntCount), StrDestination(IntCount), True
           List1.AddItem StrReverse(Mid$(StrTemp, 1, Val(StrTemp1) - 1)) & " installed successfully"
           List1.Selected(List1.ListCount - 1) = True
           List1.Refresh
        End If
        Progressbar1.Value = Progressbar1.Value + 5
 Next

For IntCount = 1 To intSkins
    If FSO.FileExists(StrSkinArray(IntCount)) Then
       StrTemp = StrReverse(StrSkinArray(IntCount))
       StrTemp1 = InStr(1, StrTemp, "\", vbTextCompare)
       List1.AddItem "Installing the " & StrReverse(Mid$(StrTemp, 1, Val(StrTemp1) - 1))
       FSO.CopyFile StrSkinArray(IntCount), StrSkinDestination(IntCount), True
       List1.AddItem StrReverse(Mid$(StrTemp, 1, Val(StrTemp1) - 1)) & " installed successfully"
       List1.Selected(List1.ListCount - 1) = True
       List1.Refresh
    End If
Next

 List1.AddItem "All Skins and Agents are installed successfully"
 List1.Selected(List1.ListCount - 1) = True
 List1.Refresh
 Timer1.Enabled = True

Shell App.Path & "\SPCHAPI.EXE", vbNormalFocus

'Install the appropriate fonts in  the system fonts folder
 If FSO.FileExists(App.Path & "\Fonts\BALLSONT.TTF") Then
    List1.AddItem "Installing the Font BALLSONT.TTF"
    If Not FSO.FileExists(gstrfontdir & "BALLSONT.TTF") Then
       FSO.CopyFile App.Path & "\Fonts\BALLSONT.TTF", gstrfontdir & "BALLSONT.TTF", True
    End If
    List1.AddItem "Font BALLSONT.TTF installed successfully"
    List1.Selected(List1.ListCount - 1) = True
    List1.Refresh
 End If
 Timer1.Enabled = True

 If FSO.FileExists(App.Path & "\Fonts\tahoma.ttf") Then
    List1.AddItem "Installing the Font Tahoma.ttf"
    If Not FSO.FileExists(gstrfontdir & "tahoma.ttf") Then
       FSO.CopyFile App.Path & "\Fonts\tahoma.ttf", gstrfontdir & "tahoma.ttf", True
    End If
    List1.AddItem "Font tahoma.ttf installed successfully"
    List1.Selected(List1.ListCount - 1) = True
    List1.Refresh
 End If
 Timer1.Enabled = True

 Progressbar1.Value = Progressbar1.Value + 5
 
 List1.AddItem "All Fonts are installed successfully"
 List1.Selected(List1.ListCount - 1) = True
 List1.Refresh
 Timer1.Enabled = True
 
'checks that if the operating system is below the windows 2000 then it will start the
'msagent service after extracting it
 List1.AddItem "Operating System Detected"
 Dim lngVersion As Long
 Dim strPlatform As String
 lngVersion = GetVersion()
 If ((lngVersion And &H80000000) = 0) Then
    glngWhichWindows32 = mlngWindowsNT
    strPlatform = "NT"
    List1.AddItem "Windows 2000 or Xp or NT"
 Else
     glngWhichWindows32 = mlngWindows95
     strPlatform = "simple"
     List1.AddItem "Windows 95 or 98"
     List1.AddItem "Running external tool"
     List1.AddItem "Microsoft Agent 2.0"
     List1.AddItem "Please install this service"
     List1.AddItem "to run the software properly"
     List1.Refresh
     If MsgBox("Are you ready to install Microsoft Agent 2.0 ?", vbYesNo) = vbYes Then
        Shell App.Path & "\msagent.exe"
     End If
 End If
 
 
'Deletes the character files from the Chars directory to prevent unauthorized copying
 FSO.DeleteFile strDrive & "\tempb\*.*", True
 FSO.DeleteFolder strDrive & "\tempb"
'Deletes the skin files from the skins directory to prevent unauthorized copying
 FSO.DeleteFile strDrive & "\temps\*.*", True
 FSO.DeleteFolder strDrive & "\temps"
 
'start the tv_enua service
 If MsgBox("Install American True Voice [Recommended] ?", vbYesNo) = vbYes Then
        Shell App.Path & "\tv_enua.exe", vbNormalFocus
 End If
 
 Progressbar1.Value = 100
 'Setting in the registry
  SaveSetting "BRAIN", "POSTINSTALL", "POSTINSTALL", True
 
 List1.AddItem "Restarting BRAIN"
 List1.Selected(List1.ListCount - 1) = True
 List1.Refresh
 Timer1.Enabled = True
 Unload Me
 MsgBox "Updation Completed. Please restart BRAIN.", vbInformation, "BRAIN updation"
 Shell App.Path & "\BRAIN.exe", vbNormalFocus
End
End Sub

Private Sub Form_Load()
Set Huffman = New ClsHuffMan
  
  If Not FSO.FolderExists(App.Path & "\Skins") Then
     MESSAGE "Skins directory not found. Cannot Proceed.", OkOnly, "Error"
     Exit Sub
  End If
  
If Not FSO.FolderExists(App.Path & "\chars") Then
   MESSAGE "Characters directory not found. Cannot Proceed.", OkOnly, "Error"
     Exit Sub
  End If
  
File1.Path = App.Path & "\chars"
File2.Path = App.Path & "\chars"
File3.Path = App.Path & "\skins"
Progressbar1.Value = 0
End Sub

'-----------------------------------------------------------
' FUNCTION: GetWindowsFontDir
'
' Calls the windows API to get the windows font directory
' and ensures that a trailing dir separator is present
'
' Returns: The windows font directory
'-----------------------------------------------------------
'
Function GetWindowsFontDir() As String
'    Dim oMalloc As IVBMalloc
'    Dim sPath   As String
'    Dim IDL     As Long
'
'    ' Fill the item id list with the pointer of each folder item, rtns 0 on success
'    If SHGetSpecialFolderLocation(0, sfidFONTS, IDL) = NOERROR Then
'        sPath = String$(gintMAX_PATH_LEN, 0)
'        SHGetPathFromIDListA IDL, sPath
'        SHGetMalloc oMalloc
'        oMalloc.Free IDL
'        sPath = StringFromBuffer(sPath)
'    End If
'    AddDirSep sPath
'
'    GetWindowsFontDir = sPath
End Function

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Picture3_Click()

End Sub

Private Sub Timer1_Timer()
Inttimer = Inttimer + 1
If Inttimer = 10000 Then Timer1.Enabled = False
  
End Sub

Function Check_Chars()
'will check that whether all the characters  are extracted or not
 If File2.ListCount = intchar Then
    Check_Chars = 0
 Else
    Check_Chars
 End If
End Function
