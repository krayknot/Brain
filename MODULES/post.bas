Attribute VB_Name = "post"
Option Explicit

Global Const gintMAX_SIZE% = 255
Global Const gstrSEP_URLDIR$ = "/"
Global Const gstrSEP_DIR$ = "\"
Global Const gintMAX_PATH_LEN% = 260

Public Enum SpecialFolderIDs
    sfidDESKTOP = &H0
    sfidPROGRAMS = &H2
    sfidPERSONAL = &H5
    sfidFAVORITES = &H6
    sfidSTARTUP = &H7
    sfidRECENT = &H8
    sfidSENDTO = &H9
    sfidSTARTMENU = &HB
    sfidDESKTOPDIRECTORY = &H10
    sfidNETHOOD = &H13
    sfidFONTS = &H14
    sfidTEMPLATES = &H15
    sfidCOMMON_STARTMENU = &H16
    sfidCOMMON_PROGRAMS = &H17
    sfidCOMMON_STARTUP = &H18
    sfidCOMMON_DESKTOPDIRECTORY = &H19
    sfidAPPDATA = &H1A
    sfidPRINTHOOD = &H1B
    sfidProgramFiles = &H10000
    sfidCommonFiles = &H10001
End Enum

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwndOwner As Long, ByVal nFolder As SpecialFolderIDs, ByRef pIdl As Long) As Long
Public Declare Function SHGetPathFromIDListA Lib "Shell32" (ByVal pIdl As Long, ByVal pszPath As String) As Long
'Public Declare Function SHGetMalloc Lib "shell32" (ByRef pMalloc As IVBMalloc) As Long



'-----------------------------------------------------------
' FUNCTION: StripTerminator
'
' Returns a string without any zero terminator.  Typically,
' this was a string returned by a Windows API call.
'
' IN: [strString] - String to remove terminator from
'
' Returns: The value of the string passed in minus any
'          terminating zero.
'-----------------------------------------------------------
'
Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

'-----------------------------------------------------------
' SUB: AddDirSep
' Add a trailing directory path separator (back slash) to the
' end of a pathname unless one already exists
'
' IN/OUT: [strPathName] - path to add separator to
'-----------------------------------------------------------
'
Sub AddDirSep(strPathName As String)
    If Right(Trim(strPathName), Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR And _
       Right(Trim(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
End Sub

Public Function StringFromBuffer(Buffer As String) As String
    Dim nPos As Long

    nPos = InStr(Buffer, Chr$(0))
    If nPos > 0 Then
        StringFromBuffer = Left$(Buffer, nPos - 1)
    Else
        StringFromBuffer = Buffer
    End If
End Function

