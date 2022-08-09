VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COPaperCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function CO_FB_PaperNum()
'Number of Fill in the blanks PApers
 CO_FB_PaperNum = 10
End Function

Function CO_MTC_PaperNum()
'Number of MAtching Columns
 CO_MTC_PaperNum = 10
End Function

Function CO_TF_PaperNum()
'Number of MAtching Columns
 CO_TF_PaperNum = 9
End Function

Function CO_QUES_PaperNum()
'Number of MAtching Columns
 CO_QUES_PaperNum = 4
End Function

Function CO_MC_PaperNum()
'Number of MAtching Columns
 CO_MC_PaperNum = 12
End Function

Function CO_FB_Option_Count()
If StrPaperType = "Fill in the Blanks" Then
   If StrPaper = 1 Then CO_FB_Option_Count = 10
   ElseIf IntPaper = 2 Then CO_FB_Option_Count = 10
   ElseIf IntPaper = 3 Then CO_FB_Option_Count = 10
   ElseIf IntPaper = 4 Then CO_FB_Option_Count = 10
   ElseIf IntPaper = 5 Then CO_FB_Option_Count = 10
   ElseIf IntPaper = 6 Then CO_FB_Option_Count = 10
   ElseIf IntPaper = 7 Then CO_FB_Option_Count = 10
   ElseIf IntPaper = 8 Then CO_FB_Option_Count = 10
   ElseIf IntPaper = 9 Then CO_FB_Option_Count = 10
   ElseIf IntPaper = 10 Then CO_FB_Option_Count = 10
   ElseIf IntPaper = 11 Then CO_FB_Option_Count = 10
   ElseIf IntPaper = 12 Then CO_FB_Option_Count = 10
   ElseIf IntPaper = 13 Then CO_FB_Option_Count = 10
   ElseIf IntPaper = 14 Then CO_FB_Option_Count = 10
   ElseIf IntPaper = 15 Then CO_FB_Option_Count = 10
End If
End Function

Function CO_MTC_Option_Count()
If StrPaperType = "Matching Columns" Then
   If StrPaper = 1 Then CO_MTC_Option_Count = 10
   ElseIf IntPaper = 2 Then CO_MTC_Option_Count = 10
   ElseIf IntPaper = 3 Then CO_MTC_Option_Count = 10
   ElseIf IntPaper = 4 Then CO_MTC_Option_Count = 10
   ElseIf IntPaper = 5 Then CO_MTC_Option_Count = 10
   ElseIf IntPaper = 6 Then CO_MTC_Option_Count = 10
   ElseIf IntPaper = 7 Then CO_MTC_Option_Count = 10
   ElseIf IntPaper = 8 Then CO_MTC_Option_Count = 10
   ElseIf IntPaper = 9 Then CO_MTC_Option_Count = 10
   ElseIf IntPaper = 10 Then CO_MTC_Option_Count = 10
   ElseIf IntPaper = 11 Then CO_MTC_Option_Count = 10
   ElseIf IntPaper = 12 Then CO_MTC_Option_Count = 10
   ElseIf IntPaper = 13 Then CO_MTC_Option_Count = 10
   ElseIf IntPaper = 14 Then CO_MTC_Option_Count = 10
   ElseIf IntPaper = 15 Then CO_MTC_Option_Count = 10
End If

   
End Function


