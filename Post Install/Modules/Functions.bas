Attribute VB_Name = "Functions"
'Function to get default windows directory
 Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

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
