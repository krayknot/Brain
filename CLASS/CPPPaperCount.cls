VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPPPaperCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function CPP_FB_PaperNum()
'Number of Fill in the blanks PApers
 CPP_FB_PaperNum = 10
End Function

Function CPP_MTC_PaperNum()
'Number of MAtching Columns
 CPP_MTC_PaperNum = 10
End Function

Function CPP_TF_PaperNum()
'Number of MAtching Columns
 CPP_TF_PaperNum = 10
End Function

Function CPP_QUES_PaperNum()
'Number of MAtching Columns
 CPP_QUES_PaperNum = 9
End Function

Function CPP_MC_PaperNum()
'Number of MAtching Columns
 CPP_MC_PaperNum = 10
End Function
Function CPP_FB_Option_Count()
If StrPaperType = "Fill in the Blanks" Then
   If StrPaper = 1 Then CPP_FB_Option_Count = 10
   ElseIf IntPaper = 2 Then CPP_FB_Option_Count = 10
   ElseIf IntPaper = 3 Then CPP_FB_Option_Count = 10
   ElseIf IntPaper = 4 Then CPP_FB_Option_Count = 10
   ElseIf IntPaper = 5 Then CPP_FB_Option_Count = 10
   ElseIf IntPaper = 6 Then CPP_FB_Option_Count = 10
   ElseIf IntPaper = 7 Then CPP_FB_Option_Count = 10
   ElseIf IntPaper = 8 Then CPP_FB_Option_Count = 10
   ElseIf IntPaper = 9 Then CPP_FB_Option_Count = 10
   ElseIf IntPaper = 10 Then CPP_FB_Option_Count = 10
   ElseIf IntPaper = 11 Then CPP_FB_Option_Count = 10
   ElseIf IntPaper = 12 Then CPP_FB_Option_Count = 10
   ElseIf IntPaper = 13 Then CPP_FB_Option_Count = 10
   ElseIf IntPaper = 14 Then CPP_FB_Option_Count = 10
   ElseIf IntPaper = 15 Then CPP_FB_Option_Count = 10
End If
End Function

Function CPP_MTC_Option_Count()
If StrPaperType = "Matching Columns" Then
   If StrPaper = 1 Then CPP_MTC_Option_Count = 10
   ElseIf IntPaper = 2 Then CPP_MTC_Option_Count = 10
   ElseIf IntPaper = 3 Then CPP_MTC_Option_Count = 10
   ElseIf IntPaper = 4 Then CPP_MTC_Option_Count = 10
   ElseIf IntPaper = 5 Then CPP_MTC_Option_Count = 10
   ElseIf IntPaper = 6 Then CPP_MTC_Option_Count = 10
   ElseIf IntPaper = 7 Then CPP_MTC_Option_Count = 10
   ElseIf IntPaper = 8 Then CPP_MTC_Option_Count = 10
   ElseIf IntPaper = 9 Then CPP_MTC_Option_Count = 10
   ElseIf IntPaper = 10 Then CPP_MTC_Option_Count = 10
   ElseIf IntPaper = 11 Then CPP_MTC_Option_Count = 10
   ElseIf IntPaper = 12 Then CPP_MTC_Option_Count = 10
   ElseIf IntPaper = 13 Then CPP_MTC_Option_Count = 10
   ElseIf IntPaper = 14 Then CPP_MTC_Option_Count = 10
   ElseIf IntPaper = 15 Then CPP_MTC_Option_Count = 10
End If

   
End Function


