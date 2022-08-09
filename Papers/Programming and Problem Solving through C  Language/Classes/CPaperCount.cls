VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPaperCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function C_FB_PaperNum()
'Number of Fill in the blanks PApers
 C_FB_PaperNum = 11
End Function

Function C_MTC_PaperNum()
'Number of MAtching Columns
 C_MTC_PaperNum = 11
End Function

Function C_TF_PaperNum()
'Number of MAtching Columns
 C_TF_PaperNum = 14
End Function

Function C_QUES_PaperNum()
'Number of MAtching Columns
 C_QUES_PaperNum = 10
End Function
Function C_MC_PaperNum()
'Number of MAtching Columns
 C_MC_PaperNum = 15
End Function

Function C_FB_Option_Count()
If StrPaperType = "Fill in the Blanks" Then
   If StrPaper = 1 Then C_FB_Option_Count = 10
   ElseIf IntPaper = 2 Then C_FB_Option_Count = 10
   ElseIf IntPaper = 3 Then C_FB_Option_Count = 10
   ElseIf IntPaper = 4 Then C_FB_Option_Count = 10
   ElseIf IntPaper = 5 Then C_FB_Option_Count = 10
   ElseIf IntPaper = 6 Then C_FB_Option_Count = 10
   ElseIf IntPaper = 7 Then C_FB_Option_Count = 10
   ElseIf IntPaper = 8 Then C_FB_Option_Count = 10
   ElseIf IntPaper = 9 Then C_FB_Option_Count = 10
   ElseIf IntPaper = 10 Then C_FB_Option_Count = 10
   ElseIf IntPaper = 11 Then C_FB_Option_Count = 10
   ElseIf IntPaper = 12 Then C_FB_Option_Count = 10
   ElseIf IntPaper = 13 Then C_FB_Option_Count = 10
   ElseIf IntPaper = 14 Then C_FB_Option_Count = 10
   ElseIf IntPaper = 15 Then C_FB_Option_Count = 10
End If
End Function

Function C_MTC_Option_Count()
If StrPaperType = "Matching Columns" Then
   If StrPaper = 1 Then C_MTC_Option_Count = 10
   ElseIf IntPaper = 2 Then C_MTC_Option_Count = 10
   ElseIf IntPaper = 3 Then C_MTC_Option_Count = 10
   ElseIf IntPaper = 4 Then C_MTC_Option_Count = 10
   ElseIf IntPaper = 5 Then C_MTC_Option_Count = 10
   ElseIf IntPaper = 6 Then C_MTC_Option_Count = 10
   ElseIf IntPaper = 7 Then C_MTC_Option_Count = 10
   ElseIf IntPaper = 8 Then C_MTC_Option_Count = 10
   ElseIf IntPaper = 9 Then C_MTC_Option_Count = 10
   ElseIf IntPaper = 10 Then C_MTC_Option_Count = 10
   ElseIf IntPaper = 11 Then C_MTC_Option_Count = 10
   ElseIf IntPaper = 12 Then C_MTC_Option_Count = 10
   ElseIf IntPaper = 13 Then C_MTC_Option_Count = 10
   ElseIf IntPaper = 14 Then C_MTC_Option_Count = 10
   ElseIf IntPaper = 15 Then C_MTC_Option_Count = 10
End If

   
End Function

