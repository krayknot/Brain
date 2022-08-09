VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnixPaperCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function UNIX_FB_PaperNum()
'Number of Fill in the blanks PApers
 UNIX_FB_PaperNum = 11
End Function

Function UNIX_MTC_PaperNum()
'Number of MAtching Columns
 UNIX_MTC_PaperNum = 10
End Function

Function UNIX_TF_PaperNum()
'Number of MAtching Columns
 UNIX_TF_PaperNum = 10
End Function

Function UNIX_QUES_PaperNum()
'Number of MAtching Columns
 UNIX_QUES_PaperNum = 6
End Function
Function UNIX_MC_PaperNum()
'Number of MAtching Columns
 UNIX_MC_PaperNum = 10
End Function

Function UNIX_FB_Option_Count()
If StrPaperType = "Fill in the Blanks" Then
   If StrPaper = 1 Then UNIX_FB_Option_Count = 10
   ElseIf IntPaper = 2 Then UNIX_FB_Option_Count = 10
   ElseIf IntPaper = 3 Then UNIX_FB_Option_Count = 10
   ElseIf IntPaper = 4 Then UNIX_FB_Option_Count = 10
   ElseIf IntPaper = 5 Then UNIX_FB_Option_Count = 10
   ElseIf IntPaper = 6 Then UNIX_FB_Option_Count = 10
   ElseIf IntPaper = 7 Then UNIX_FB_Option_Count = 10
   ElseIf IntPaper = 8 Then UNIX_FB_Option_Count = 10
   ElseIf IntPaper = 9 Then UNIX_FB_Option_Count = 10
   ElseIf IntPaper = 10 Then UNIX_FB_Option_Count = 10
   ElseIf IntPaper = 11 Then UNIX_FB_Option_Count = 10
   ElseIf IntPaper = 12 Then UNIX_FB_Option_Count = 10
   ElseIf IntPaper = 13 Then UNIX_FB_Option_Count = 10
   ElseIf IntPaper = 14 Then UNIX_FB_Option_Count = 10
   ElseIf IntPaper = 15 Then UNIX_FB_Option_Count = 10
End If
End Function

Function UNIX_MTC_Option_Count()
If StrPaperType = "Matching Columns" Then
   If StrPaper = 1 Then UNIX_MTC_Option_Count = 10
   ElseIf IntPaper = 2 Then UNIX_MTC_Option_Count = 10
   ElseIf IntPaper = 3 Then UNIX_MTC_Option_Count = 10
   ElseIf IntPaper = 4 Then UNIX_MTC_Option_Count = 10
   ElseIf IntPaper = 5 Then UNIX_MTC_Option_Count = 10
   ElseIf IntPaper = 6 Then UNIX_MTC_Option_Count = 10
   ElseIf IntPaper = 7 Then UNIX_MTC_Option_Count = 10
   ElseIf IntPaper = 8 Then UNIX_MTC_Option_Count = 10
   ElseIf IntPaper = 9 Then UNIX_MTC_Option_Count = 10
   ElseIf IntPaper = 10 Then UNIX_MTC_Option_Count = 10
   ElseIf IntPaper = 11 Then UNIX_MTC_Option_Count = 10
   ElseIf IntPaper = 12 Then UNIX_MTC_Option_Count = 10
   ElseIf IntPaper = 13 Then UNIX_MTC_Option_Count = 10
   ElseIf IntPaper = 14 Then UNIX_MTC_Option_Count = 10
   ElseIf IntPaper = 15 Then UNIX_MTC_Option_Count = 10
End If

   
End Function


