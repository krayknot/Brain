VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCNPaperCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function DCN_FB_PaperNum()
'Number of Fill in the blanks PApers
 DCN_FB_PaperNum = 10
End Function

Function DCN_MTC_PaperNum()
'Number of MAtching Columns
 DCN_MTC_PaperNum = 10
End Function

Function DCN_TF_PaperNum()
'Number of MAtching Columns
 DCN_TF_PaperNum = 10
End Function

Function DCN_QUES_PaperNum()
'Number of MAtching Columns
 DCN_QUES_PaperNum = 7
End Function

Function DCN_MC_PaperNum()
'Number of MAtching Columns
 DCN_MC_PaperNum = 10
End Function

Function DCN_FB_Option_Count()
If StrPaperType = "Fill in the Blanks" Then
   If StrPaper = 1 Then DCN_FB_Option_Count = 10
   ElseIf IntPaper = 2 Then DCN_FB_Option_Count = 10
   ElseIf IntPaper = 3 Then DCN_FB_Option_Count = 10
   ElseIf IntPaper = 4 Then DCN_FB_Option_Count = 10
   ElseIf IntPaper = 5 Then DCN_FB_Option_Count = 10
   ElseIf IntPaper = 6 Then DCN_FB_Option_Count = 10
   ElseIf IntPaper = 7 Then DCN_FB_Option_Count = 10
   ElseIf IntPaper = 8 Then DCN_FB_Option_Count = 10
   ElseIf IntPaper = 9 Then DCN_FB_Option_Count = 10
   ElseIf IntPaper = 10 Then DCN_FB_Option_Count = 10
   ElseIf IntPaper = 11 Then DCN_FB_Option_Count = 10
   ElseIf IntPaper = 12 Then DCN_FB_Option_Count = 10
   ElseIf IntPaper = 13 Then DCN_FB_Option_Count = 10
   ElseIf IntPaper = 14 Then DCN_FB_Option_Count = 10
   ElseIf IntPaper = 15 Then DCN_FB_Option_Count = 10
End If
End Function

Function DCN_MTC_Option_Count()
If StrPaperType = "Matching Columns" Then
   If StrPaper = 1 Then DCN_MTC_Option_Count = 10
   ElseIf IntPaper = 2 Then DCN_MTC_Option_Count = 10
   ElseIf IntPaper = 3 Then DCN_MTC_Option_Count = 10
   ElseIf IntPaper = 4 Then DCN_MTC_Option_Count = 10
   ElseIf IntPaper = 5 Then DCN_MTC_Option_Count = 10
   ElseIf IntPaper = 6 Then DCN_MTC_Option_Count = 10
   ElseIf IntPaper = 7 Then DCN_MTC_Option_Count = 10
   ElseIf IntPaper = 8 Then DCN_MTC_Option_Count = 10
   ElseIf IntPaper = 9 Then DCN_MTC_Option_Count = 10
   ElseIf IntPaper = 10 Then DCN_MTC_Option_Count = 10
   ElseIf IntPaper = 11 Then DCN_MTC_Option_Count = 10
   ElseIf IntPaper = 12 Then DCN_MTC_Option_Count = 10
   ElseIf IntPaper = 13 Then DCN_MTC_Option_Count = 10
   ElseIf IntPaper = 14 Then DCN_MTC_Option_Count = 10
   ElseIf IntPaper = 15 Then DCN_MTC_Option_Count = 10
End If

   
End Function


