VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SADPaperCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function SAD_FB_PaperNum()
'Number of Fill in the blanks PApers
 SAD_FB_PaperNum = 16
End Function

Function SAD_MTC_PaperNum()
'Number of MAtching Columns
 SAD_MTC_PaperNum = 15
End Function

Function SAD_TF_PaperNum()
'Number of MAtching Columns
 SAD_TF_PaperNum = 18
End Function

Function SAD_QUES_PaperNum()
'Number of MAtching Columns
 SAD_QUES_PaperNum = 4
End Function

Function SAD_MC_PaperNum()
'Number of MAtching Columns
 SAD_MC_PaperNum = 20
End Function


Function SAD_FB_Option_Count()
If StrPaperType = "Fill in the Blanks" Then
   If StrPaper = 1 Then SAD_FB_Option_Count = 10
   ElseIf IntPaper = 2 Then SAD_FB_Option_Count = 10
   ElseIf IntPaper = 3 Then SAD_FB_Option_Count = 10
   ElseIf IntPaper = 4 Then SAD_FB_Option_Count = 10
   ElseIf IntPaper = 5 Then SAD_FB_Option_Count = 10
   ElseIf IntPaper = 6 Then SAD_FB_Option_Count = 10
   ElseIf IntPaper = 7 Then SAD_FB_Option_Count = 10
   ElseIf IntPaper = 8 Then SAD_FB_Option_Count = 10
   ElseIf IntPaper = 9 Then SAD_FB_Option_Count = 10
   ElseIf IntPaper = 10 Then SAD_FB_Option_Count = 10
   ElseIf IntPaper = 11 Then SAD_FB_Option_Count = 10
   ElseIf IntPaper = 12 Then SAD_FB_Option_Count = 10
   ElseIf IntPaper = 13 Then SAD_FB_Option_Count = 10
   ElseIf IntPaper = 14 Then SAD_FB_Option_Count = 10
   ElseIf IntPaper = 15 Then SAD_FB_Option_Count = 10
End If
End Function

Function SAD_MTC_Option_Count()
If StrPaperType = "Matching Columns" Then
   If StrPaper = 1 Then SAD_MTC_Option_Count = 10
   ElseIf IntPaper = 2 Then SAD_MTC_Option_Count = 10
   ElseIf IntPaper = 3 Then SAD_MTC_Option_Count = 10
   ElseIf IntPaper = 4 Then SAD_MTC_Option_Count = 10
   ElseIf IntPaper = 5 Then SAD_MTC_Option_Count = 10
   ElseIf IntPaper = 6 Then SAD_MTC_Option_Count = 10
   ElseIf IntPaper = 7 Then SAD_MTC_Option_Count = 10
   ElseIf IntPaper = 8 Then SAD_MTC_Option_Count = 10
   ElseIf IntPaper = 9 Then SAD_MTC_Option_Count = 10
   ElseIf IntPaper = 10 Then SAD_MTC_Option_Count = 10
   ElseIf IntPaper = 11 Then SAD_MTC_Option_Count = 10
   ElseIf IntPaper = 12 Then SAD_MTC_Option_Count = 10
   ElseIf IntPaper = 13 Then SAD_MTC_Option_Count = 10
   ElseIf IntPaper = 14 Then SAD_MTC_Option_Count = 10
   ElseIf IntPaper = 15 Then SAD_MTC_Option_Count = 10
End If

   
End Function

