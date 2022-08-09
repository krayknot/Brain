VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BSPaperCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function BS_FB_PaperNum()
'Number of Fill in the blanks PApers
 BS_FB_PaperNum = 11
End Function

Function BS_MTC_PaperNum()
'Number of MAtching Columns
 BS_MTC_PaperNum = 11
End Function

Function BS_TF_PaperNum()
'Number of MAtching Columns
 BS_TF_PaperNum = 15
End Function

Function BS_QUES_PaperNum()
'Number of MAtching Columns
 BS_QUES_PaperNum = 11
End Function
Function BS_MC_PaperNum()
'Number of MAtching Columns
 BS_MC_PaperNum = 10
End Function

Function BS_FB_Option_Count()
If StrPaperType = "Fill in the Blanks" Then
   If StrPaper = 1 Then BS_FB_Option_Count = 10
   ElseIf IntPaper = 2 Then BS_FB_Option_Count = 10
   ElseIf IntPaper = 3 Then BS_FB_Option_Count = 10
   ElseIf IntPaper = 4 Then BS_FB_Option_Count = 10
   ElseIf IntPaper = 5 Then BS_FB_Option_Count = 10
   ElseIf IntPaper = 6 Then BS_FB_Option_Count = 10
   ElseIf IntPaper = 7 Then BS_FB_Option_Count = 10
   ElseIf IntPaper = 8 Then BS_FB_Option_Count = 10
   ElseIf IntPaper = 9 Then BS_FB_Option_Count = 10
   ElseIf IntPaper = 10 Then BS_FB_Option_Count = 10
   ElseIf IntPaper = 11 Then BS_FB_Option_Count = 10
   ElseIf IntPaper = 12 Then BS_FB_Option_Count = 10
   ElseIf IntPaper = 13 Then BS_FB_Option_Count = 10
   ElseIf IntPaper = 14 Then BS_FB_Option_Count = 10
   ElseIf IntPaper = 15 Then BS_FB_Option_Count = 10
End If
End Function

Function BS_MTC_Option_Count()
If StrPaperType = "Matching Columns" Then
   If StrPaper = 1 Then BS_MTC_Option_Count = 10
   ElseIf IntPaper = 2 Then BS_MTC_Option_Count = 10
   ElseIf IntPaper = 3 Then BS_MTC_Option_Count = 10
   ElseIf IntPaper = 4 Then BS_MTC_Option_Count = 10
   ElseIf IntPaper = 5 Then BS_MTC_Option_Count = 10
   ElseIf IntPaper = 6 Then BS_MTC_Option_Count = 10
   ElseIf IntPaper = 7 Then BS_MTC_Option_Count = 10
   ElseIf IntPaper = 8 Then BS_MTC_Option_Count = 10
   ElseIf IntPaper = 9 Then BS_MTC_Option_Count = 10
   ElseIf IntPaper = 10 Then BS_MTC_Option_Count = 10
   ElseIf IntPaper = 11 Then BS_MTC_Option_Count = 10
   ElseIf IntPaper = 12 Then BS_MTC_Option_Count = 10
   ElseIf IntPaper = 13 Then BS_MTC_Option_Count = 10
   ElseIf IntPaper = 14 Then BS_MTC_Option_Count = 10
   ElseIf IntPaper = 15 Then BS_MTC_Option_Count = 10
End If

   
End Function

