VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IWPDPaperCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function IWPD_FB_PaperNum()
'Number of Fill in the blanks PApers
 IWPD_FB_PaperNum = 12
End Function

Function IWPD_MTC_PaperNum()
'Number of MAtching Columns
 IWPD_MTC_PaperNum = 14
End Function

Function IWPD_TF_PaperNum()
'Number of MAtching Columns
 IWPD_TF_PaperNum = 14
End Function

Function IWPD_QUES_PaperNum()
'Number of MAtching Columns
 IWPD_QUES_PaperNum = 6
End Function
Function IWPD_MC_PaperNum()
'Number of MAtching Columns
 IWPD_MC_PaperNum = 11
End Function

Function IWPD_FB_Option_Count()
If StrPaperType = "Fill in the Blanks" Then
   If StrPaper = 1 Then IWPD_FB_Option_Count = 10
   ElseIf IntPaper = 2 Then IWPD_FB_Option_Count = 10
   ElseIf IntPaper = 3 Then IWPD_FB_Option_Count = 10
   ElseIf IntPaper = 4 Then IWPD_FB_Option_Count = 10
   ElseIf IntPaper = 5 Then IWPD_FB_Option_Count = 10
   ElseIf IntPaper = 6 Then IWPD_FB_Option_Count = 10
   ElseIf IntPaper = 7 Then IWPD_FB_Option_Count = 10
   ElseIf IntPaper = 8 Then IWPD_FB_Option_Count = 10
   ElseIf IntPaper = 9 Then IWPD_FB_Option_Count = 10
   ElseIf IntPaper = 10 Then IWPD_FB_Option_Count = 10
   ElseIf IntPaper = 11 Then IWPD_FB_Option_Count = 10
   ElseIf IntPaper = 12 Then IWPD_FB_Option_Count = 10
   ElseIf IntPaper = 13 Then IWPD_FB_Option_Count = 10
   ElseIf IntPaper = 14 Then IWPD_FB_Option_Count = 10
   ElseIf IntPaper = 15 Then IWPD_FB_Option_Count = 10
End If
End Function

Function IWPD_MTC_Option_Count()
If StrPaperType = "Matching Columns" Then
   If StrPaper = 1 Then IWPD_MTC_Option_Count = 10
   ElseIf IntPaper = 2 Then IWPD_MTC_Option_Count = 10
   ElseIf IntPaper = 3 Then IWPD_MTC_Option_Count = 10
   ElseIf IntPaper = 4 Then IWPD_MTC_Option_Count = 10
   ElseIf IntPaper = 5 Then IWPD_MTC_Option_Count = 10
   ElseIf IntPaper = 6 Then IWPD_MTC_Option_Count = 10
   ElseIf IntPaper = 7 Then IWPD_MTC_Option_Count = 10
   ElseIf IntPaper = 8 Then IWPD_MTC_Option_Count = 10
   ElseIf IntPaper = 9 Then IWPD_MTC_Option_Count = 10
   ElseIf IntPaper = 10 Then IWPD_MTC_Option_Count = 10
   ElseIf IntPaper = 11 Then IWPD_MTC_Option_Count = 10
   ElseIf IntPaper = 12 Then IWPD_MTC_Option_Count = 10
   ElseIf IntPaper = 13 Then IWPD_MTC_Option_Count = 10
   ElseIf IntPaper = 14 Then IWPD_MTC_Option_Count = 10
   ElseIf IntPaper = 15 Then IWPD_MTC_Option_Count = 10
End If

   
End Function

