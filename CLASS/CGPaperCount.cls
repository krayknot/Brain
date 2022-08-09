VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGPaperCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function CG_FB_PaperNum()
'Number of Fill in the blanks PApers
 CG_FB_PaperNum = 10
End Function

Function CG_MTC_PaperNum()
'Number of MAtching Columns
 CG_MTC_PaperNum = 9
End Function

Function CG_TF_PaperNum()
'Number of MAtching Columns
 CG_TF_PaperNum = 15
End Function

Function CG_QUES_PaperNum()
'Number of MAtching Columns
 CG_QUES_PaperNum = 7
End Function
Function CG_MC_PaperNum()
'Number of MAtching Columns
 CG_MC_PaperNum = 10
End Function

Function CG_FB_Option_Count()
If StrPaperType = "Fill in the Blanks" Then
   If StrPaper = 1 Then CG_FB_Option_Count = 10
   ElseIf IntPaper = 2 Then CG_FB_Option_Count = 10
   ElseIf IntPaper = 3 Then CG_FB_Option_Count = 10
   ElseIf IntPaper = 4 Then CG_FB_Option_Count = 10
   ElseIf IntPaper = 5 Then CG_FB_Option_Count = 10
   ElseIf IntPaper = 6 Then CG_FB_Option_Count = 10
   ElseIf IntPaper = 7 Then CG_FB_Option_Count = 10
   ElseIf IntPaper = 8 Then CG_FB_Option_Count = 10
   ElseIf IntPaper = 9 Then CG_FB_Option_Count = 10
   ElseIf IntPaper = 10 Then CG_FB_Option_Count = 10
   ElseIf IntPaper = 11 Then CG_FB_Option_Count = 10
   ElseIf IntPaper = 12 Then CG_FB_Option_Count = 10
   ElseIf IntPaper = 13 Then CG_FB_Option_Count = 10
   ElseIf IntPaper = 14 Then CG_FB_Option_Count = 10
   ElseIf IntPaper = 15 Then CG_FB_Option_Count = 10
End If
End Function

Function CG_MTC_Option_Count()
If StrPaperType = "Matching Columns" Then
   If StrPaper = 1 Then CG_MTC_Option_Count = 10
   ElseIf IntPaper = 2 Then CG_MTC_Option_Count = 10
   ElseIf IntPaper = 3 Then CG_MTC_Option_Count = 10
   ElseIf IntPaper = 4 Then CG_MTC_Option_Count = 10
   ElseIf IntPaper = 5 Then CG_MTC_Option_Count = 10
   ElseIf IntPaper = 6 Then CG_MTC_Option_Count = 10
   ElseIf IntPaper = 7 Then CG_MTC_Option_Count = 10
   ElseIf IntPaper = 8 Then CG_MTC_Option_Count = 10
   ElseIf IntPaper = 9 Then CG_MTC_Option_Count = 10
   ElseIf IntPaper = 10 Then CG_MTC_Option_Count = 10
   ElseIf IntPaper = 11 Then CG_MTC_Option_Count = 10
   ElseIf IntPaper = 12 Then CG_MTC_Option_Count = 10
   ElseIf IntPaper = 13 Then CG_MTC_Option_Count = 10
   ElseIf IntPaper = 14 Then CG_MTC_Option_Count = 10
   ElseIf IntPaper = 15 Then CG_MTC_Option_Count = 10
End If

   
End Function


