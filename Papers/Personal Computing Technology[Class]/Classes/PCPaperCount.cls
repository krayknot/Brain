VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCPaperCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function PC_FB_PaperNum()
'Number of Fill in the blanks PApers
 PC_FB_PaperNum = 12
End Function

Function PC_MTC_PaperNum()
'Number of MAtching Columns
 PC_MTC_PaperNum = 12
End Function

Function PC_TF_PaperNum()
'Number of MAtching Columns
 PC_TF_PaperNum = 19
End Function

Function PC_QUES_PaperNum()
'Number of MAtching Columns
 PC_QUES_PaperNum = 15
End Function

Function PC_MC_PaperNum()
'Number of MAtching Columns
 PC_MC_PaperNum = 12
End Function

Function PC_FB_Option_Count()
If StrPaperType = "Fill in the Blanks" Then
   If StrPaper = 1 Then PC_FB_Option_Count = 10
   ElseIf IntPaper = 2 Then PC_FB_Option_Count = 10
   ElseIf IntPaper = 3 Then PC_FB_Option_Count = 10
   ElseIf IntPaper = 4 Then PC_FB_Option_Count = 10
   ElseIf IntPaper = 5 Then PC_FB_Option_Count = 10
   ElseIf IntPaper = 6 Then PC_FB_Option_Count = 10
   ElseIf IntPaper = 7 Then PC_FB_Option_Count = 10
   ElseIf IntPaper = 8 Then PC_FB_Option_Count = 10
   ElseIf IntPaper = 9 Then PC_FB_Option_Count = 10
   ElseIf IntPaper = 10 Then PC_FB_Option_Count = 10
   ElseIf IntPaper = 11 Then PC_FB_Option_Count = 10
   ElseIf IntPaper = 12 Then PC_FB_Option_Count = 10
   ElseIf IntPaper = 13 Then PC_FB_Option_Count = 10
   ElseIf IntPaper = 14 Then PC_FB_Option_Count = 10
   ElseIf IntPaper = 15 Then PC_FB_Option_Count = 10
End If
End Function

Function PC_MTC_Option_Count()
If StrPaperType = "Matching Columns" Then
   If StrPaper = 1 Then PC_MTC_Option_Count = 10
   ElseIf IntPaper = 2 Then PC_MTC_Option_Count = 10
   ElseIf IntPaper = 3 Then PC_MTC_Option_Count = 10
   ElseIf IntPaper = 4 Then PC_MTC_Option_Count = 10
   ElseIf IntPaper = 5 Then PC_MTC_Option_Count = 10
   ElseIf IntPaper = 6 Then PC_MTC_Option_Count = 10
   ElseIf IntPaper = 7 Then PC_MTC_Option_Count = 10
   ElseIf IntPaper = 8 Then PC_MTC_Option_Count = 10
   ElseIf IntPaper = 9 Then PC_MTC_Option_Count = 10
   ElseIf IntPaper = 10 Then PC_MTC_Option_Count = 10
   ElseIf IntPaper = 11 Then PC_MTC_Option_Count = 10
   ElseIf IntPaper = 12 Then PC_MTC_Option_Count = 10
   ElseIf IntPaper = 13 Then PC_MTC_Option_Count = 10
   ElseIf IntPaper = 14 Then PC_MTC_Option_Count = 10
   ElseIf IntPaper = 15 Then PC_MTC_Option_Count = 10
End If

   
End Function

