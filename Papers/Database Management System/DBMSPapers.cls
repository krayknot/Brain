VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBMSPapers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim strquestions(10) As String
Dim stroptions(10) As String
Dim stranswers(10) As String
Dim strcolumnA(10) As String
Dim strcolumnB(10) As String

Enum Category
 Fill_Questions = 0
 Fill_Options = 1
 Fill_Answers = 2
 Matching_ColumnA = 3
 Matching_ColumnB = 4
 Matching_Answers = 5
End Enum

Enum Paper
 Paper1 = 1
 Paper2 = 2
 Paper3 = 3
 Paper4 = 4
 Paper5 = 5
 Paper6 = 6
 Paper7 = 7
 Paper8 = 8
 Paper9 = 9
 Paper10 = 10
 Paper11 = 11
 Paper12 = 12
 Paper13 = 13
 Paper14 = 14
 Paper15 = 15
End Enum

Public Function DBMS(PaperType As String, DBMSPaper As Paper, mCat As Category, Number As Integer)
'ptype is the paper type
'DBMSpaper is the number of paper as paper1,paper2 ...
'mcat is the category
'number is the number of questions, or options or answers to be taken

 StrPaperType = PaperType
 StrPaper = Paper
If PaperType = "Fill in the Blanks" Then
   If DBMSPaper = Paper1 Then Fill_in_the_Blanks_Paper1
   If DBMSPaper = Paper2 Then Fill_in_the_Blanks_Paper2
   If DBMSPaper = Paper3 Then Fill_in_the_Blanks_Paper3
   If DBMSPaper = Paper4 Then Fill_in_the_Blanks_Paper4
   If DBMSPaper = Paper5 Then Fill_in_the_Blanks_Paper5
   If DBMSPaper = Paper6 Then Fill_in_the_Blanks_Paper6
   If DBMSPaper = Paper7 Then Fill_in_the_Blanks_Paper7
   If DBMSPaper = Paper8 Then Fill_in_the_Blanks_Paper8
   If DBMSPaper = Paper9 Then Fill_in_the_Blanks_Paper9
   If DBMSPaper = Paper10 Then Fill_in_the_Blanks_Paper10
   If DBMSPaper = Paper11 Then Fill_in_the_Blanks_Paper11
   If DBMSPaper = Paper12 Then Fill_in_the_Blanks_Paper12
   If DBMSPaper = Paper13 Then Fill_in_the_Blanks_Paper13
   If DBMSPaper = Paper14 Then Fill_in_the_Blanks_Paper14
   If DBMSPaper = Paper15 Then Fill_in_the_Blanks_Paper15
   
ElseIf PaperType = "Matching Columns" Then
       If DBMSPaper = Paper1 Then Matching_Columns_Paper1
End If

If mCat = Fill_Questions Then
   DBMS = strquestions(Number)
ElseIf mCat = Fill_Options Then
       DBMS = stroptions(Number)
ElseIf mCat = Fill_Answers Then
       DBMS = stranswers(Number)
ElseIf mCat = Matching_ColumnA Then
       DBMS = strcolumnA(Number)
ElseIf mCat = Matching_ColumnB Then
       DBMS = strcolumnB(Number)
ElseIf mCat = Matching_Answers Then
       DBMS = stranswers(Number)
End If

End Function

Private Sub Fill_in_the_Blanks_Paper1()
strquestions(1) = "SQL queries are based on _____________ operation."
strquestions(2) = "The _____________ of a given relation is the set of tuples appearing in that relation at any given instant."
strquestions(3) = "In the relational data model, the columns of a table are called _____________."
strquestions(4) = "A multi level index leads to a _____________."
strquestions(5) = "_____________ is a specification of the portion off data which is oriented to the needs of one or more application programs."
strquestions(6) = "SET _____________ TO is used to set up an open index file as the master index, without closing other index files."
strquestions(7) = "The primary key index does not allow _____________ data in a field."
strquestions(8) = "_____________ is updated before a transaction commits."
strquestions(9) = "In a lock based schedule _____________ can occur."
strquestions(10) = "Read lock of a data item by a transaction prevents other transactions to _____________ that data."

stroptions(1) = "Tuple"
stroptions(2) = "Sets"
stroptions(3) = "Attributes"
stroptions(4) = "B-Tree"
stroptions(5) = "Binary tree"
stroptions(6) = "Mapping"
stroptions(7) = "Selection"
stroptions(8) = "Where"
stroptions(9) = "For"
stroptions(10) = "While"
stroptions(11) = "Extension"
stroptions(12) = "Schema"
stroptions(13) = "Subschema"
stroptions(14) = "Modify"
stroptions(15) = "Read"
stroptions(16) = "Log"
stroptions(17) = "Order"
stroptions(18) = "Index"
stroptions(19) = "Deadlock"
stroptions(20) = "Lovelock"
stroptions(20) = "Duplicate"

stranswers(1) = "Selection"
stranswers(2) = "Extension"
stranswers(3) = "Attributes"
stranswers(4) = "B-Tree"
stranswers(5) = "Subschema"
stranswers(6) = "Order"
stranswers(7) = "Duplicate"
stranswers(8) = "Log"
stranswers(9) = "Deadlock"
stranswers(10) = "Modify"

End Sub

Private Sub Fill_in_the_Blanks_Paper()
strquestions(1) = ""
strquestions(2) = ""
strquestions(3) = ""
strquestions(4) = ""
strquestions(5) = ""
strquestions(6) = ""
strquestions(7) = ""
strquestions(8) = ""
strquestions(9) = ""
strquestions(10) = ""

stroptions(1) = ""
stroptions(2) = ""
stroptions(3) = ""
stroptions(4) = ""
stroptions(5) = ""
stroptions(6) = ""
stroptions(7) = ""
stroptions(8) = ""
stroptions(9) = ""
stroptions(10) = ""

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""

End Sub
