VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITPaperCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function IT_FB_PaperNum()
'Number of Fill in the blanks PApers
 IT_FB_PaperNum = 15
End Function

Function IT_MTC_PaperNum()
'Number of MAtching Columns
 IT_MTC_PaperNum = 15
End Function

Function IT_TF_PaperNum()
'Number of MAtching Columns
 IT_TF_PaperNum = 15
End Function

Function IT_QUES_PaperNum()
'Number of MAtching Columns
 IT_QUES_PaperNum = 15
End Function

Function IT_MC_PaperNum()
'Number of MAtching Columns
 IT_MC_PaperNum = 15
End Function

Function IT_FB_Option_Count()
If StrPaperType = "Fill in the Blanks" Then
   If StrPaper = 1 Then IT_FB_Option_Count = 10
   ElseIf IntPaper = 2 Then IT_FB_Option_Count = 10
   ElseIf IntPaper = 3 Then IT_FB_Option_Count = 10
   ElseIf IntPaper = 4 Then IT_FB_Option_Count = 10
   ElseIf IntPaper = 5 Then IT_FB_Option_Count = 10
   ElseIf IntPaper = 6 Then IT_FB_Option_Count = 10
   ElseIf IntPaper = 7 Then IT_FB_Option_Count = 10
   ElseIf IntPaper = 8 Then IT_FB_Option_Count = 10
   ElseIf IntPaper = 9 Then IT_FB_Option_Count = 10
   ElseIf IntPaper = 10 Then IT_FB_Option_Count = 10
   ElseIf IntPaper = 11 Then IT_FB_Option_Count = 10
   ElseIf IntPaper = 12 Then IT_FB_Option_Count = 10
   ElseIf IntPaper = 13 Then IT_FB_Option_Count = 10
   ElseIf IntPaper = 14 Then IT_FB_Option_Count = 10
   ElseIf IntPaper = 15 Then IT_FB_Option_Count = 10
End If
End Function

Function IT_MTC_Option_Count()
If StrPaperType = "Matching Columns" Then
   If StrPaper = 1 Then IT_MTC_Option_Count = 10
   ElseIf IntPaper = 2 Then IT_MTC_Option_Count = 10
   ElseIf IntPaper = 3 Then IT_MTC_Option_Count = 10
   ElseIf IntPaper = 4 Then IT_MTC_Option_Count = 10
   ElseIf IntPaper = 5 Then IT_MTC_Option_Count = 10
   ElseIf IntPaper = 6 Then IT_MTC_Option_Count = 10
   ElseIf IntPaper = 7 Then IT_MTC_Option_Count = 10
   ElseIf IntPaper = 8 Then IT_MTC_Option_Count = 10
   ElseIf IntPaper = 9 Then IT_MTC_Option_Count = 10
   ElseIf IntPaper = 10 Then IT_MTC_Option_Count = 10
   ElseIf IntPaper = 11 Then IT_MTC_Option_Count = 10
   ElseIf IntPaper = 12 Then IT_MTC_Option_Count = 10
   ElseIf IntPaper = 13 Then IT_MTC_Option_Count = 10
   ElseIf IntPaper = 14 Then IT_MTC_Option_Count = 10
   ElseIf IntPaper = 15 Then IT_MTC_Option_Count = 10
End If

   
End Function

