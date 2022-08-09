VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BSPapers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Dim StrQuestions(10) As String * 5000
 Dim stroptions(10) As String
 Dim stranswers(10) As String
 Dim strcolumnA(10) As String
 Dim strcolumnB(10) As String

'Variable declaration for the multiple choice
 Dim strmultiple(10) As String
 Dim strmuloption(40) As String
 Dim strmulanswer(10) As String
 
Enum Category
 Fill_Questions = 0
 Fill_Options = 1
 Fill_Answers = 2
 Matching_ColumnA = 3
 Matching_ColumnB = 4
 Matching_Answers = 5
 Questions = 6
 TF_Questions = 7
 TF_Answers = 8
 MC_Questions = 9
 MC_Options_1 = 10
 MC_Options_2 = 11
 MC_Options_3 = 12
 MC_Options_4 = 13
 MC_Answer = 14
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

Public Function BS(PaperType As String, BSPaper As Paper, mCat As Category, Number As Integer)
'ptype is the paper type
'itpaper is the number of paper as paper1,paper2 ...
'mcat is the category
'number is the number of questions, or options or answers to be taken

 StrPaperType = PaperType
' StrPaper = Paper
If PaperType = "Fill in the Blanks" Then
   If BSPaper = Paper1 Then Fill_in_the_Blanks_Paper1
   If BSPaper = Paper2 Then Fill_in_the_Blanks_Paper2
   If BSPaper = Paper3 Then Fill_in_the_Blanks_Paper3
   If BSPaper = Paper4 Then Fill_in_the_Blanks_Paper4
   If BSPaper = Paper5 Then Fill_in_the_Blanks_Paper5
   If BSPaper = Paper6 Then Fill_in_the_Blanks_Paper6
   If BSPaper = Paper7 Then Fill_in_the_Blanks_Paper7
   If BSPaper = Paper8 Then Fill_in_the_Blanks_Paper8
   If BSPaper = Paper9 Then Fill_in_the_Blanks_Paper9
   If BSPaper = Paper10 Then Fill_in_the_Blanks_Paper10
   If BSPaper = Paper11 Then Fill_in_the_Blanks_Paper11
   
ElseIf PaperType = "Matching Columns" Then
       If BSPaper = Paper1 Then Matching_Columns_Paper1
       If BSPaper = Paper2 Then Matching_Columns_Paper2
       If BSPaper = Paper3 Then Matching_Columns_Paper3
       If BSPaper = Paper4 Then Matching_Columns_Paper4
       If BSPaper = Paper5 Then Matching_Columns_Paper5
       If BSPaper = Paper6 Then Matching_Columns_Paper6
       If BSPaper = Paper7 Then Matching_Columns_Paper7
       If BSPaper = Paper8 Then Matching_Columns_Paper8
       If BSPaper = Paper9 Then Matching_Columns_Paper9
       If BSPaper = Paper10 Then Matching_Columns_Paper10
       If BSPaper = Paper11 Then Matching_Columns_Paper11
       
ElseIf PaperType = "True Or False" Then
       If BSPaper = Paper1 Then TrueorFalse_Paper1
       If BSPaper = Paper2 Then TrueorFalse_Paper2
       If BSPaper = Paper3 Then TrueorFalse_Paper3
       If BSPaper = Paper4 Then TrueorFalse_Paper4
       If BSPaper = Paper5 Then TrueorFalse_Paper5
       If BSPaper = Paper6 Then TrueorFalse_Paper6
       If BSPaper = Paper7 Then TrueorFalse_Paper7
       If BSPaper = Paper8 Then TrueorFalse_Paper8
       If BSPaper = Paper9 Then TrueorFalse_Paper9
       If BSPaper = Paper10 Then TrueorFalse_Paper10
       If BSPaper = Paper11 Then TrueorFalse_Paper11
       If BSPaper = Paper12 Then TrueorFalse_Paper12
       If BSPaper = Paper13 Then TrueorFalse_Paper13
       If BSPaper = Paper14 Then TrueorFalse_Paper14
       If BSPaper = Paper15 Then TrueorFalse_Paper15
       
ElseIf PaperType = "Descriptive Questions" Then
       If BSPaper = Paper1 Then questions_paper1
       If BSPaper = Paper2 Then questions_paper2
       If BSPaper = Paper3 Then questions_paper3
       If BSPaper = Paper4 Then questions_paper4
       If BSPaper = Paper5 Then questions_paper5
       If BSPaper = Paper6 Then questions_paper6
       If BSPaper = Paper7 Then questions_paper7
       If BSPaper = Paper8 Then questions_paper8
       If BSPaper = Paper9 Then questions_paper9
       If BSPaper = Paper10 Then questions_paper10
       If BSPaper = Paper11 Then questions_paper11
       
 ElseIf PaperType = "Multiple Choice" Then
       If BSPaper = Paper1 Then multiplechoice_Paper1
       If BSPaper = Paper2 Then multiplechoice_Paper2
       If BSPaper = Paper3 Then multiplechoice_Paper3
       If BSPaper = Paper4 Then multiplechoice_Paper4
       If BSPaper = Paper5 Then multiplechoice_Paper5
       If BSPaper = Paper6 Then multiplechoice_Paper6
       If BSPaper = Paper7 Then multiplechoice_Paper7
       If BSPaper = Paper8 Then multiplechoice_Paper8
       If BSPaper = Paper9 Then multiplechoice_Paper9
       If BSPaper = Paper10 Then multiplechoice_Paper10
End If

If mCat = Fill_Questions Then
   BS = StrQuestions(Number)
ElseIf mCat = Fill_Options Then
       BS = stroptions(Number)
ElseIf mCat = Fill_Answers Then
       BS = stranswers(Number)
       
ElseIf mCat = Matching_ColumnA Then
       BS = strcolumnA(Number)
ElseIf mCat = Matching_ColumnB Then
       BS = strcolumnB(Number)
ElseIf mCat = Matching_Answers Then
       BS = stranswers(Number)
       
ElseIf mCat = Questions Then
       BS = StrQuestions(Number)
       
ElseIf mCat = TF_Questions Then
       BS = StrQuestions(Number)
ElseIf mCat = TF_Answers Then
       BS = stranswers(Number)
       
ElseIf mCat = MC_Questions Then
       BS = strmultiple(Number)
       
ElseIf mCat = MC_Options_1 Then
       BS = strmuloption(Number)
       
ElseIf mCat = MC_Options_2 Then
       BS = strmuloption(Number)
       
ElseIf mCat = MC_Options_3 Then
       BS = strmuloption(Number)
       
ElseIf mCat = MC_Options_4 Then
       BS = strmuloption(Number)
       
ElseIf mCat = MC_Answer Then
       BS = strmulanswer(Number)
       
End If

End Function

Private Sub Fill_in_the_Blanks_Paper1()
StrQuestions(1) = "A sales invoice debits customers and dash sales."
StrQuestions(2) = "A portion of program that performs a specific program function is called a dash."
StrQuestions(3) = "Pay image file is dash of pay slip."
StrQuestions(4) = "The process of ascertaining the logical conformity of data to certain prescribed characteristics is termed dash."
StrQuestions(5) = "Study of the meaning of the language structure is called the dash of the language"
StrQuestions(6) = "A suspense account reflects balances which do not have clear cut dash defined"
StrQuestions(7) = "Depreciation is considered dash item."
StrQuestions(8) = "Vendor performance can be judged by dash of the supplies."
StrQuestions(9) = "There may be minimal inventory held if the lead time is dash."
StrQuestions(10) = "Loans to employees form part of dash."

stroptions(1) = "Credits"
stroptions(2) = "Subsidiary"
stroptions(3) = "Timeliness"
stroptions(4) = "An Expense"
stroptions(5) = "Account heads"
stroptions(6) = "Zero"
stroptions(7) = "Validation"
stroptions(8) = "Module"
stroptions(9) = "Semantics"
stroptions(10) = "A replica"

stranswers(1) = "Credits"
stranswers(2) = "Module"
stranswers(3) = "A replica"
stranswers(4) = "Validation"
stranswers(5) = "Semantics"
stranswers(6) = "Accounts head"
stranswers(7) = "Expense"
stranswers(8) = "Timeliness"
stranswers(9) = "Zero"
stranswers(10) = "Subsidiary Ledger"
End Sub

Private Sub Fill_in_the_Blanks_Paper2()
StrQuestions(1) = "A special field can be placed in a dash record at the end of a file to indicate that all valid records have been processed."
StrQuestions(2) = "The process of executing the developed program with various data inputs in order to detect programming errors or omissions is called dash."
StrQuestions(3) = "A term which describes program segments written in ordinary language rather than in a programming language is dash."
StrQuestions(4) = "A dash file issued to update a master file."
StrQuestions(5) = "Computer software with which one user creates applications are called dash."
StrQuestions(6) = "Outstanding analysis requires the invoice file and dash file."
StrQuestions(7) = "A suspense account reflects balances whose dash in still to be decided."
StrQuestions(8) = "dash is the difference between gross profit and operating expenses."
StrQuestions(9) = "Ageing analysis, of the stocks held, help the company in dash its inventory."
StrQuestions(10) = "Sales to customers form part of dash."

stroptions(1) = "Operating Profit"
stroptions(2) = "Subsidiary ledgers"
stroptions(3) = "Payments receipt"
stroptions(4) = "Disposal"
stroptions(5) = "Pseudo code"
stroptions(6) = "Optimizing"
stroptions(7) = "4 GL"
stroptions(8) = "Transaction"
stroptions(9) = "Testing"
stroptions(10) = "Dummy"

stranswers(1) = "Dummy"
stranswers(2) = "Testing"
stranswers(3) = "Pseudo code"
stranswers(4) = "Transaction"
stranswers(5) = "4GL"
stranswers(6) = "Payment receipt"
stranswers(7) = "Disposal"
stranswers(8) = "Operating Profit"
stranswers(9) = "Optimizing"
stranswers(10) = "Subsidiary Ledger"

End Sub

Private Sub Fill_in_the_Blanks_Paper3()
StrQuestions(1) = "Valuation of stock can be done by dash."
StrQuestions(2) = "Each module of structured program should have dash entry point."
StrQuestions(3) = "Typically, a computer based provident fund system used inputs from the computerized dash ."
StrQuestions(4) = "Data stored on disks can be accessed in either a direct or dash way."
StrQuestions(5) = "A dash is a number of related records that are treated as a unit."
StrQuestions(6) = "A unique identifying value that is found in every record in a file is called a record dash."
StrQuestions(7) = "After every year we merge the P & L of operations with the dash."
StrQuestions(8) = "Purchases from vendors form part of dash."
StrQuestions(9) = "A journal voucher is also passed to incorporate corrections/amendments to the dash."
StrQuestions(10) = "Strict inventory controls help dash product obsolescence."

stroptions(1) = "Books of account"
stroptions(2) = "Balance sheet"
stroptions(3) = "Minimize"
stroptions(4) = "Subsidiary ledger"
stroptions(5) = "Key"
stroptions(6) = "File"
stroptions(7) = "Standard"
stroptions(8) = "Sequential"
stroptions(9) = "One"
stroptions(10) = "Payroll"

stranswers(1) = "Standard"
stranswers(2) = "One"
stranswers(3) = "Payroll"
stranswers(4) = "Sequential"
stranswers(5) = "File"
stranswers(6) = "Key"
stranswers(7) = "Balance Sheet"
stranswers(8) = "Subsidiary Ledger"
stranswers(9) = "Book of Accounts"
stranswers(10) = "Minimize"

End Sub

Private Sub Fill_in_the_Blanks_Paper4()
StrQuestions(1) = "The delimiters used to delimit subscripts, indexes, conditions or arithmetic expressions are dash."
StrQuestions(2) = "Whenever we withdraw funds from the bank to meet our day to day expenses, we dash the contra account and debit the cash book."
StrQuestions(3) = "A dash decision is a routing one that's made following the completion of a preplanned series of steps."
StrQuestions(4) = "A dash decision is a routing one that's made following the completion of a preplanned series of steps."
StrQuestions(5) = "Routine and mundane tasks such as sorting, copying etc. are done by dash ."
StrQuestions(6) = "Journal vouchers dash affect the bank or cast book balance."
StrQuestions(7) = "The balances of contra account should always be dash."
StrQuestions(8) = "After each year we dash start with the P & L statement afresh."
StrQuestions(9) = "Variation in demand of a product are taken care by carrying dash."
StrQuestions(10) = "The volume of purchases by a customer forms an important parameter for determining the dash."

stroptions(1) = "Need to"
stroptions(2) = "Do not"
stroptions(3) = "Credit worthiness"
stroptions(4) = "Credit"
stroptions(5) = "Structured"
stroptions(6) = "Left & Right parentheses"
stroptions(7) = "Buffer stock"
stroptions(8) = "System software"
stroptions(9) = "Utility programs"
stroptions(10) = "Zero"

stranswers(1) = "Left & Right Parenthesis"
stranswers(2) = "Credit"
stranswers(3) = "Structured"
stranswers(4) = "System Software"
stranswers(5) = "Utility programs"
stranswers(6) = "Do not"
stranswers(7) = "Zero"
stranswers(8) = "Need to"
stranswers(9) = "Need to"
stranswers(10) = "Need to"

End Sub

Private Sub Fill_in_the_Blanks_Paper5()
StrQuestions(1) = "A cylinder index identifies the key of the highest record stored in the dash cylinder of interest."
StrQuestions(2) = "An indexed file essentially is a dash."
StrQuestions(3) = "Schedules to annual accounts show dash regarding the working of the company."
StrQuestions(4) = "A database management system is essentially a specialized type of dash package."
StrQuestions(5) = "Schedules to annual accounts have to b prepared to comply with dash requirements of the law."
StrQuestions(6) = "The balance sheet of company dash the status of a company at any given point of time."
StrQuestions(7) = "Whenever we withdraw funds from the bank to meet our day to day expenses, we dash the contra account & credit the bank account."
StrQuestions(8) = "Valuation of stock can be done by using dash method."
StrQuestions(9) = "The memory space which is setup to transfer records to and from a file is called a dash."
StrQuestions(10) = "dash file is a master file in a sales invoicing system."

stroptions(1) = "Detailed information"
stroptions(2) = "Customer details"
stroptions(3) = "Buffer"
stroptions(4) = "Sequential"
stroptions(5) = "Debit"
stroptions(6) = "Software"
stroptions(7) = "Reflects"
stroptions(8) = "Conceptual"
stroptions(9) = "Weighted average "
stroptions(10) = "Statutory"

stranswers(1) = "Conceptual"
stranswers(2) = "Sequential"
stranswers(3) = "Details of information"
stranswers(4) = "Software"
stranswers(5) = "Statutory"
stranswers(6) = "Reflect"
stranswers(7) = "Debit"
stranswers(8) = "Weighted average"
stranswers(9) = "Buffer"
stranswers(10) = "Customer details"

End Sub

Private Sub Fill_in_the_Blanks_Paper6()
StrQuestions(1) = "dash is the graphical representation of an algorithm."
StrQuestions(2) = "The formal rules governing the construction of valid statements in a language are called dash."
StrQuestions(3) = "Logical organization of data in file is also called dash."
StrQuestions(4) = "The approach towards program design in which main program is decomposed into progressively lower level modules is called dash."
StrQuestions(5) = "Debits an credits must be dash in a double entry book keeping system."
StrQuestions(6) = "The total of all debits may not be affected when a dash is posted in the books."
StrQuestions(7) = "Reserves and surpluses constitute the dash side of the balance sheet."
StrQuestions(8) = "An ABC analysis shows that a small number of items carry the dash of the inventory."
StrQuestions(9) = "Valuation of stock can be done by dash method."
StrQuestions(10) = "Computerized Railway Reservation System is an example of dash data processing system."

stroptions(1) = "Relatively high value"
stroptions(2) = "Data structure"
stroptions(3) = "Liabilities"
stroptions(4) = "Flowchart"
stroptions(5) = "Syntax"
stroptions(6) = "Top down"
stroptions(7) = "Matching"
stroptions(8) = "Centralised"
stroptions(9) = "LIFO"
stroptions(10) = "Voucher"

stranswers(1) = "Flowchart"
stranswers(2) = "Syntax"
stranswers(3) = "Data structure"
stranswers(4) = "Top Down"
stranswers(5) = "Matching"
stranswers(6) = "Voucher"
stranswers(7) = "Liabilities"
stranswers(8) = "Relatively high value"
stranswers(9) = "LIFO"
stranswers(10) = "Centralised"

End Sub

Private Sub Fill_in_the_Blanks_Paper7()
StrQuestions(1) = "Detailed design logic of an application, developed using a package, can be represented by pseudo codes. Another way of representing this is by use of the dash."
StrQuestions(2) = "The fields within a logical record on which files are sorted are called dash."
StrQuestions(3) = "The process of identifying the causes of program errors and correcting then is called dash."
StrQuestions(4) = "If it is possible to access data in any desired order from an external storage device, then the device is said to have dash access capability."
StrQuestions(5) = "If it is possible to access data in any desired order from an external storage device, then the device is said to have dash access capability."
StrQuestions(6) = "A dash is a listing of the ledger accounts, along with their net debit or credit balances."
StrQuestions(7) = "dash Allow decision makers to measure, interpret, analyze & communicate financial information."
StrQuestions(8) = "Subsidiary ledgers of a proprietary concern may or not be linked with dash."
StrQuestions(9) = "Vendor performance can be judged by the dash supplied."
StrQuestions(10) = "Vendor performance can be judged by dash of supplies."

stroptions(1) = "Financial statements"
stroptions(2) = "Flow chart"
stroptions(3) = "Pricing of material"
stroptions(4) = "Key field"
stroptions(5) = "Debugging"
stroptions(6) = "Random"
stroptions(7) = "Quality"
stroptions(8) = "Accounting systems"
stroptions(9) = "Books of account"
stroptions(10) = "Trial Balance"

stranswers(1) = "Flowchart"
stranswers(2) = "Key field"
stranswers(3) = "Debugging"
stranswers(4) = "Random"
stranswers(5) = "Accounting systems"
stranswers(6) = "Trial Balance"
stranswers(7) = "Financial Statements"
stranswers(8) = "Books of account"
stranswers(9) = "Pricing of material"
stranswers(10) = "Quality"

End Sub

Private Sub Fill_in_the_Blanks_Paper8()
StrQuestions(1) = "The dash command is used to close a database file in use."
StrQuestions(2) = "The file name can be up to dash characters long."
StrQuestions(3) = "The dash command is used to select the default drive for accessing database and other files."
StrQuestions(4) = "To send the output of LIST to printer, use the dash clause with LIST."
StrQuestions(5) = "The dash function removes leading and trailing blanks form a character expression."
StrQuestions(6) = "dash command requires vacant space on the disk that is equal to or more than the size of the database being sorted."
StrQuestions(7) = "After a record has been located with LOCATE, the dash command is used to search the next record."
StrQuestions(8) = "To use the LOCATE command, the database file need NOT be sorted or dash."
StrQuestions(9) = "For displaying contents of a memo field with DISPLAY/EDIT, it is essential to specify dash with the command."
StrQuestions(10) = "To suppress page title (header) after the first page of the report, use the dash option with REPORT FORM."

stroptions(1) = "INDEXED"
stroptions(2) = "PLAIN"
stroptions(3) = "10"
stroptions(4) = "CONTINUE"
stroptions(5) = "TO PRINTER"
stroptions(6) = "SET DEFAULT"
stroptions(7) = "MEMO FIELDNAME"
stroptions(8) = "CLOSE DATABASE/USE"
stroptions(9) = "ALL TRIM()"
stroptions(10) = "SORT"

stranswers(1) = "CLOSE DATABASE/USE"
stranswers(2) = "10"
stranswers(3) = "SET DEFAULT"
stranswers(4) = "TO PRINTER"
stranswers(5) = "ALLTRIM()"
stranswers(6) = "SORT"
stranswers(7) = "CONTINUE"
stranswers(8) = "INDEXED"
stranswers(9) = "MEMO FIELDNAME"
stranswers(10) = "PLAIN"

End Sub

Private Sub Fill_in_the_Blanks_Paper9()
StrQuestions(1) = "If a variable NUMBER contains 11500, STR (VALUE,3) returns dash."
StrQuestions(2) = "The dash command clears all data from the current window."
StrQuestions(3) = "Regardless of the setting of SET PRINTER, dash sends output to the printer."
StrQuestions(4) = "The function (command) assigned to a function key can be changed by using the dash command."
StrQuestions(5) = "CTOD () is the reverse of the dash function."
StrQuestions(6) = "dash transfers the control to the beginning of DO WHILE - ENDDO loop."
StrQuestions(7) = "The coordinates for the top left position of a window are dash."
StrQuestions(8) = "The dash option is used with @...GET to restrict the user response within specifies limits, e.g. 10 through 30."
StrQuestions(9) = "After a variable has been displayed on the screen with @...SAY...GET , the dash command is used to enable the user to enter the value for the variable."
StrQuestions(10) = "To store a paintbrush picture, one can use dash type of field in a database file."

stroptions(1) = "READ OR READ CYCLE"
stroptions(2) = "ASTERISKS"
stroptions(3) = "GENERAL"
stroptions(4) = "RANGE"
stroptions(5) = "???"
stroptions(6) = "DO WHILE -  ENDDO"
stroptions(7) = "CLEAR"
stroptions(8) = "0,0"
stroptions(9) = "DTOC()"
stroptions(10) = "SET FUNCTION"

stranswers(1) = "ASTERISKS"
stranswers(2) = "CLEAR"
stranswers(3) = "???"
stranswers(4) = "SET FUNCTION"
stranswers(5) = "DTOC()"
stranswers(6) = "DO WHILE – ENDDO"
stranswers(7) = "0,0"
stranswers(8) = "RANGE"
stranswers(9) = "READ"
stranswers(10) = "GENERAL"

End Sub

Private Sub Fill_in_the_Blanks_Paper10()
StrQuestions(1) = "The report design window is divided in three parts. These are called page header, page footer and dash."
StrQuestions(2) = "RQBE saves the query in file. This file is given an extension name dash."
StrQuestions(3) = "To arrange the records as per some order with RQBE, the SELECT command uses the clause dash."
StrQuestions(4) = "The dash command removes all records from the current database file."
StrQuestions(5) = "The dash function returns true when the record pointer is positioned before the first record."
StrQuestions(6) = "The dash function converts a character expression to upper case."
StrQuestions(7) = "dash selects the next record."
StrQuestions(8) = "dash selects the previous record."
StrQuestions(9) = "RTRIM () and dash are similar functions."
StrQuestions(10) = "dash computers the arithmetic mean of numeric fields."

stroptions(1) = "TRIM()"
stroptions(2) = "SKIP -1"
stroptions(3) = ".QPR"
stroptions(4) = "UPPER()"
stroptions(5) = "ORDER BY"
stroptions(6) = "DETAIL"
stroptions(7) = "BOF()"
stroptions(8) = "SKIP"
stroptions(9) = "AVERAGE"
stroptions(10) = "ZAP"

stranswers(1) = "DETAIL"
stranswers(2) = ".QPR"
stranswers(3) = "ORDER BY"
stranswers(4) = "ZAP"
stranswers(5) = "BOF()"
stranswers(6) = "UPPER()"
stranswers(7) = "SKIP"
stranswers(8) = "SKIP -1"
stranswers(9) = "TRIM()"
stranswers(10) = "AVERAGE"

End Sub


Private Sub Fill_in_the_Blanks_Paper11()
StrQuestions(1) = "The default extension name used with a screen file created with CREATE SCREEN is dash."
StrQuestions(2) = "To edit the current record using the ADDRESS .spr screen file, the dash command is used."
StrQuestions(3) = "The dash option with the ? Command displays the text at specified column."
StrQuestions(4) = "The MOD () function returns dash on division of two numbers."
StrQuestions(5) = "The STR () function converts numeric data to dash."
StrQuestions(6) = "dash causes the control to unconditionally exit from a programming loop such as DO WHILE - ENDDO."
StrQuestions(7) = "dash is a computer science branch that is involved with using computers to solve problems that appear to require human imagination or intelligence."
StrQuestions(8) = "dash clause is used to specify alignment of an elementary data item to the natural boundaries of the computer memory."
StrQuestions(9) = "dash statement is not desirable in a structured program."
StrQuestions(10) = "An input procedure in a SORT verb must have a statement beginning with the verb dash."

stroptions(1) = "EXIT"
stroptions(2) = "REMAINDER"
stroptions(3) = "DO ADDRESS.SPR"
stroptions(4) = "AT"
stroptions(5) = "GO TO"
stroptions(6) = "RELEASE"
stroptions(7) = "CHARACTER DATA"
stroptions(8) = "SYNC"
stroptions(9) = ".SPR"
stroptions(10) = "AI"

stranswers(1) = ".SPR"
stranswers(2) = "DO ADDRESS.SPR"
stranswers(3) = "AT"
stranswers(4) = "REMAINDER"
stranswers(5) = "CHARACTER"
stranswers(6) = "EXIT"
stranswers(7) = "AT"
stranswers(8) = "SYNC"
stranswers(9) = "GOTO"
stranswers(10) = "RELEASE"

End Sub
Private Sub Matching_Columns_Paper1()
strcolumnA(1) = "Purchase bills"
strcolumnA(2) = "Ageing analusis"
strcolumnA(3) = "Current assets"
strcolumnA(4) = "Income tax calculations"
strcolumnA(5) = "Journal voucher"
strcolumnA(6) = "Reorder level"
strcolumnA(7) = "Carrying cost"
strcolumnA(8) = "Vensor rating"
strcolumnA(9) = "P & L operations"
strcolumnA(10) = "Recovery from failure"

strcolumnB(1) = "Timely supplies"
strcolumnB(2) = "Pay slip files"
strcolumnB(3) = "Debtor control"
strcolumnB(4) = "Merge in balance sheet"
strcolumnB(5) = "Financial accounts"
strcolumnB(6) = "Check point restart"
strcolumnB(7) = "Payroll processing & posting in financial A/C."
strcolumnB(8) = "Creditors analysis"
strcolumnB(9) = "Inventory policy"
strcolumnB(10) = "Dead stock"

stranswers(1) = "Creditors analysis"
stranswers(2) = "Debtor control"
stranswers(3) = "Financial accounts"
stranswers(4) = "Pay slip files"
stranswers(5) = "Payroll processing & posting in financial A/C."
stranswers(6) = "Inventory policy"
stranswers(7) = "Dead stock"
stranswers(8) = "Timely supplies"
stranswers(9) = "Merge in balance sheet"
stranswers(10) = "Check point restart"

End Sub

Private Sub Matching_Columns_Paper2()
strcolumnA(1) = "Creditors control account"
strcolumnA(2) = "Debtors control account"
strcolumnA(3) = "Disk file organisation"
strcolumnA(4) = "EOQ"
strcolumnA(5) = "LIFO"
strcolumnA(6) = "Vendor rating"
strcolumnA(7) = "Depreciation"
strcolumnA(8) = "Loans to employees"
strcolumnA(9) = "Suspense account"
strcolumnA(10) = "Occupation"

strcolumnB(1) = "Random access"
strcolumnB(2) = "Hotel management"
strcolumnB(3) = "Subsidiary ledger"
strcolumnB(4) = "Inventory valuation"
strcolumnB(5) = "Treatment undecided"
strcolumnB(6) = "Suppliers account"
strcolumnB(7) = "Price consideration of supplies"
strcolumnB(8) = "Cusotmers account"
strcolumnB(9) = "Expense item in books"
strcolumnB(10) = "Inventory management"

stranswers(1) = "Suppliers account"
stranswers(2) = "Cusotmers account"
stranswers(3) = "Random access"
stranswers(4) = "Inventory management"
stranswers(5) = "Inventory valuation"
stranswers(6) = "Price consideration of supplies"
stranswers(7) = "Expense item in books"
stranswers(8) = "Subsidiary ledger"
stranswers(9) = "Treatment undecided"
stranswers(10) = "Hotel management"

End Sub
Private Sub Matching_Columns_Paper3()
strcolumnA(1) = "Purchases from supplers"
strcolumnA(2) = "P & L A/C head"
strcolumnA(3) = "Journal voucher"
strcolumnA(4) = "Vendor rating"
strcolumnA(5) = "Zero inventory level"
strcolumnA(6) = "Stock valuation"
strcolumnA(7) = "Schedules to annual A/C"
strcolumnA(8) = "Balance sheet"
strcolumnA(9) = "Funds withdrawn"
strcolumnA(10) = "Diamond"

strcolumnB(1) = "Credit bank A/C & debit contra"
strcolumnB(2) = "Low invetory storage"
strcolumnB(3) = "Flowchart"
strcolumnB(4) = "Current status of company"
strcolumnB(5) = "LIFO"
strcolumnB(6) = "Detailed financial information"
strcolumnB(7) = "Quality product supplies"
strcolumnB(8) = "Subsidiary ledger"
strcolumnB(9) = "Modifies books of account"
strcolumnB(10) = "Create afresh each year"

stranswers(1) = "Subsidiary ledger"
stranswers(2) = "Create afresh each year"
stranswers(3) = "Modifies books of account"
stranswers(4) = "Quality product supplies"
stranswers(5) = "Low invetory storage"
stranswers(6) = "LIFO"
stranswers(7) = "Detailed financial information"
stranswers(8) = "Current status of company"
stranswers(9) = "Credit bank A/C & debit contra"
stranswers(10) = "Flowchart"

End Sub
Private Sub Matching_Columns_Paper4()
strcolumnA(1) = "Sales invoice system"
strcolumnA(2) = "Error report"
strcolumnA(3) = "Pseudocode"
strcolumnA(4) = "Contra account"
strcolumnA(5) = "Re-Order level"
strcolumnA(6) = "Vendor rating"
strcolumnA(7) = "Ageing analysis"
strcolumnA(8) = "Personal loans & advances"
strcolumnA(9) = "Inventory control"
strcolumnA(10) = " Payroll"

strcolumnB(1) = "ABC analysis"
strcolumnB(2) = "Reliability"
strcolumnB(3) = "Financial accounting"
strcolumnB(4) = "Liabilities"
strcolumnB(5) = "Batch processing"
strcolumnB(6) = "Always zero"
strcolumnB(7) = "Customer master"
strcolumnB(8) = "Warning signal"
strcolumnB(9) = "Check digit"
strcolumnB(10) = "Programming aid"

stranswers(1) = "Customer master"
stranswers(2) = "Check digit"
stranswers(3) = "Programming aid"
stranswers(4) = "Always zero"
stranswers(5) = "Warning signal"
stranswers(6) = "Reliability"
stranswers(7) = "Financial accounting"
stranswers(8) = "Liabilities"
stranswers(9) = "ABC analysis"
stranswers(10) = "Batch processing"

End Sub
Private Sub Matching_Columns_Paper5()
strcolumnA(1) = "Sales to customers"
strcolumnA(2) = "Hash totals"
strcolumnA(3) = "Smallest loical data entity"
strcolumnA(4) = "Stockout"
strcolumnA(5) = "Dead stock"
strcolumnA(6) = "Vendor rating"
strcolumnA(7) = "Employee advances"
strcolumnA(8) = "Pay image file"
strcolumnA(9) = "Working capital"
strcolumnA(10) = "FORTRAN"

strcolumnB(1) = "Non moving items"
strcolumnB(2) = "Compiler"
strcolumnB(3) = "Payroll"
strcolumnB(4) = "Delievery timeliness"
strcolumnB(5) = "Control account"
strcolumnB(6) = "Current assets"
strcolumnB(7) = "Validation"
strcolumnB(8) = "Data item"
strcolumnB(9) = "Inventory malfunction"
strcolumnB(10) = "Subsdiary ledger"

stranswers(1) = "Subsdiary ledger"
stranswers(2) = "Validation"
stranswers(3) = "Data item"
stranswers(4) = "Inventory malfunction"
stranswers(5) = "Non moving items"
stranswers(6) = "Delievery timeliness"
stranswers(7) = "Control account"
stranswers(8) = "Payroll"
stranswers(9) = "Current assets"
stranswers(10) = "Compiler"

End Sub
Private Sub Matching_Columns_Paper6()
strcolumnA(1) = "Payroll system"
strcolumnA(2) = "Subsidiary ledger"
strcolumnA(3) = "General ledger"
strcolumnA(4) = "Double entry bools"
strcolumnA(5) = "Vendor rating"
strcolumnA(6) = "standard cost"
strcolumnA(7) = "Strict inventory controls"
strcolumnA(8) = "Journal voucher"
strcolumnA(9) = "Walkthrough"
strcolumnA(10) = "Aid to programming"

strcolumnB(1) = "G L code"
strcolumnB(2) = "Control osolescence"
strcolumnB(3) = "Quality of material supplies"
strcolumnB(4) = "Pseudocode"
strcolumnB(5) = "Stock valuation"
strcolumnB(6) = "P F Statement"
strcolumnB(7) = "Programming"
strcolumnB(8) = "Control account"
strcolumnB(9) = "No effect on cash ool balance"
strcolumnB(10) = "Match Dr and Cr"

stranswers(1) = "P F Statement"
stranswers(2) = "Control account"
stranswers(3) = "G L code"
stranswers(4) = "Match Dr and Cr"
stranswers(5) = "Quality of material supplies"
stranswers(6) = "Stock valuation"
stranswers(7) = "Control osolescence"
stranswers(8) = "No effect on cash ool balance"
stranswers(9) = "Programming"
stranswers(10) = "Pseudocode"

End Sub
Private Sub Matching_Columns_Paper7()
strcolumnA(1) = "Voucher"
strcolumnA(2) = "Reserves"
strcolumnA(3) = "Schedules to annual accounts"
strcolumnA(4) = "Sles fluctuations"
strcolumnA(5) = "Zero lead time"
strcolumnA(6) = "TDS"
strcolumnA(7) = "Skills inventory"
strcolumnA(8) = "Funds deposited"
strcolumnA(9) = "Transaction document"
strcolumnA(10) = "Suspense account"

strcolumnB(1) = "Statutory requirement"
strcolumnB(2) = "Credit conta & debit cash]"
strcolumnB(3) = "Debits"
strcolumnB(4) = "Payroll system"
strcolumnB(5) = "Balance sheet"
strcolumnB(6) = "Unclear account heads"
strcolumnB(7) = "Buffer stock"
strcolumnB(8) = "Voucher"
strcolumnB(9) = "Personnel systems"
strcolumnB(10) = "Minimized inventory"

stranswers(1) = "Debits"
stranswers(2) = "Balance sheet"
stranswers(3) = "Statutory requirement"
stranswers(4) = "Buffer stock"
stranswers(5) = "Minimized inventory"
stranswers(6) = "Payroll system"
stranswers(7) = "Personnel systems"
stranswers(8) = "Credit conta & debit cash]"
stranswers(9) = "Voucher"
stranswers(10) = "Unclear account heads"

End Sub
Private Sub Matching_Columns_Paper8()
strcolumnA(1) = "Displays clock on screen "
strcolumnA(2) = "Selects provious record"
strcolumnA(3) = "Displays yesterday's date"
strcolumnA(4) = "Creates empty character string"
strcolumnA(5) = "Creates memory variables"
strcolumnA(6) = "elects printer"
strcolumnA(7) = "Adds comment to a command line"
strcolumnA(8) = "Comes out of a DO SHILE ENDDO loop"
strcolumnA(9) = "Used with WAIT WINDOW"
strcolumnA(10) = "Creates numeric variable"

strcolumnB(1) = "TIMEOUT"
strcolumnB(2) = "&&"
strcolumnB(3) = "INPUT"
strcolumnB(4) = "EXIT"
strcolumnB(5) = "TIME()"
strcolumnB(6) = "DATE()-1"
strcolumnB(7) = "SET PRINTER ON"
strcolumnB(8) = "SKIP - 1"
strcolumnB(9) = "SPACE()"
strcolumnB(10) = "="

stranswers(1) = "TIME()"
stranswers(2) = "SKIP - 1"
stranswers(3) = "DATE()-1"
stranswers(4) = "SPACE()"
stranswers(5) = "="
stranswers(6) = "SET PRINTER ON"
stranswers(7) = "&&"
stranswers(8) = "EXIT"
stranswers(9) = "TIMEOUT"
stranswers(10) = "INPUT"

End Sub
Private Sub Matching_Columns_Paper9()
strcolumnA(1) = "Creates character variable"
strcolumnA(2) = "Selects a work area"
strcolumnA(3) = "NOTE"
strcolumnA(4) = "NODELETE"
strcolumnA(5) = "READ"
strcolumnA(6) = "OFF"
strcolumnA(7) = "APPEND FRO<"
strcolumnA(8) = "PICTURE " & Chr(4) & "@!" & Chr(4)
strcolumnA(9) = "@...GET"
strcolumnA(10) = "RIGHT()"

strcolumnB(1) = "@..SAY..GET"
strcolumnB(2) = "Converts to uppercase"
strcolumnB(3) = "Used with LIST"
strcolumnB(4) = "Displays the right part ofacharactr expression"
strcolumnB(5) = "SELECT"
strcolumnB(6) = "VALID"
strcolumnB(7) = "BROWSE"
strcolumnB(8) = "ACCEPT/SPACE"
strcolumnB(9) = "Comment line"
strcolumnB(10) = "READS records from a file"

stranswers(1) = "ACCEPT/SPACE"
stranswers(2) = "SELECT"
stranswers(3) = "Comment line"
stranswers(4) = "BROWSE"
stranswers(5) = "@..SAY..GET"
stranswers(6) = "Used with LIST"
stranswers(7) = "READS records from a file"
stranswers(8) = "Converts to uppercase"
stranswers(9) = "VALID"
stranswers(10) = "Displays the right part ofacharactr expression"

End Sub

Private Sub Matching_Columns_Paper10()
strcolumnA(1) = "REST"
strcolumnA(2) = "WITH CDX"
strcolumnA(3) = "SET filter"
strcolumnA(4) = ".SPR"
strcolumnA(5) = "??"
strcolumnA(6) = "???"
strcolumnA(7) = "MODIFY STRUCTURE"
strcolumnA(8) = "PACK"
strcolumnA(9) = "DO"
strcolumnA(10) = "ERASE"

strcolumnB(1) = "Runs a program file"
strcolumnB(2) = "Screen file"
strcolumnB(3) = "Deletes a file"
strcolumnB(4) = "Sends to printer"
strcolumnB(5) = "Removes deleted records"
strcolumnB(6) = "Also copies index file"
strcolumnB(7) = "Displays in the same line"
strcolumnB(8) = "Current & following records"
strcolumnB(9) = "Adds or modifies a fields"
strcolumnB(10) = "Restricts number of records with commands"

stranswers(1) = "Current & following records"
stranswers(2) = "Also copies index file"
stranswers(3) = "Restricts number of records with commands"
stranswers(4) = "Screen file"
stranswers(5) = "Displays in the same line"
stranswers(6) = "Sends to printer"
stranswers(7) = "Adds or modifies a fields"
stranswers(8) = "Removes deleted records"
stranswers(9) = "Runs a program file"
stranswers(10) = "Deletes a file"

End Sub

Private Sub Matching_Columns_Paper11()
strcolumnA(1) = "RANGE"
strcolumnA(2) = "AZP"
strcolumnA(3) = "RUN"
strcolumnA(4) = "SORT"
strcolumnA(5) = "EOF()"
strcolumnA(6) = "DO CASE"
strcolumnA(7) = "SET CARRY ON"
strcolumnA(8) = "ALIAS"
strcolumnA(9) = "INDEXING"

strcolumnB(1) = "More efficient code genration thanb subscripting"
strcolumnB(2) = "Copies the contents of privious record to the text"
strcolumnB(3) = "Returns .T. when record pointer is at the last record"
strcolumnB(4) = "Removes all records"
strcolumnB(5) = "OTHERWISE"
strcolumnB(6) = "Executes a DOS program"
strcolumnB(7) = "Restricts entry in a numeric field"
strcolumnB(8) = "Creates a database file with records arranged in some order"
strcolumnB(9) = "Alternate name for a database file"

stranswers(1) = "Restricts entry in a numeric field"
stranswers(2) = "Removes all records"
stranswers(3) = "Executes a DOS program"
stranswers(4) = "Creates a database file with records arranged in some order"
stranswers(5) = "Returns .T. when record pointer is at the last record"
stranswers(6) = "OTHERWISE"
stranswers(7) = "Copies the contents of privious record to the text"
stranswers(8) = "Alternate name for a database file"
stranswers(9) = "More efficient code genration thanb subscripting"

End Sub

Private Sub TrueorFalse_Paper1()
StrQuestions(1) = "When two or more sequential files ar combined to form one sequential file, the process is called sorting."
StrQuestions(2) = " Sequential organization of adisk file is particularly a;;ropriate for files of low activity."
StrQuestions(3) = "Structured programming diminishes the efficiency of coding of programs."
StrQuestions(4) = "Data validation isgeneratlly accomplished b making use of various types of control totals."
StrQuestions(5) = "User friendly packages are a boo to end users since they do not have to call up their DP experts."
StrQuestions(6) = "Leave record, though afunction in personnel management does not form part of the payroll system."
StrQuestions(7) = "In a personnel data processing system, the number of master records ofth employeed stats invariant while th number of transaction change."
StrQuestions(8) = "Unsecured loans ar borrowings ofth firms against which specific securities have been provided."
StrQuestions(9) = "Depreciation is a journal entry & does not affect cash/bank."
StrQuestions(10) = "Bank balance is an asset of the compamny."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper2()
StrQuestions(1) = "Capital is a liability head."
StrQuestions(2) = "Profit of a company is shown under assets."
StrQuestions(3) = "Total wae bill is incorporated in the financial acconts by passing a journal voucher."
StrQuestions(4) = "Inter-departmental movement of stocks does ot affect the balance sheet ofthe company."
StrQuestions(5) = "Basic EOQ model assumes that inventory cannot be obtained without any timne lag."
StrQuestions(6) = "In an inventory systemn, reorder quantity and EOQ mean the same thing."
StrQuestions(7) = "The term stock incudes manpower."
StrQuestions(8) = "ABC analysis is useful in cateorising he fault occurence inthe maintenance of equipment."
StrQuestions(9) = "The top down approach to develop modula program involves working from the general to the specific."
StrQuestions(10) = "In a highly volatile system, incesec files are preferred over sequential file organisaiton for data retrieval. "

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"

End Sub

Private Sub TrueorFalse_Paper3()
StrQuestions(1) = "In areal time syste, all changes are recorded on an online transaction file andthe master file updated forreport production when required."
StrQuestions(2) = "The attendence record of an employee forms the transaction to thepayroll system."
StrQuestions(3) = "Ageing analysis of debtors is an important contol report of the sales analysis system."
StrQuestions(4) = "Sales day book is an output from the computerised sales analy7sis asystem."
StrQuestions(5) = "Typically, a computer based provident fund system uses inputs ffom the computerised payroll system."
StrQuestions(6) = "Balance sheet shows the company's health at an instant while profit & loss shows the compauy health at an isstant while profit & loss shows the company performance over a period of time."
StrQuestions(7) = "A purchase transaction increases the liability of a company."
StrQuestions(8) = "A sale made increasses the current assets of the company."
StrQuestions(9) = "The strength of an organisation cab ne revirewed through the profit and loss and alance sheet."
StrQuestions(10) = "Money withdraw from the bank to make cast payment would result in debiting the bank and crediting cash book."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"

End Sub

Private Sub TrueorFalse_Paper4()
StrQuestions(1) = "Year end processing of inancial accounts entail zeroing all accounts of the trial balance."
StrQuestions(2) = "A journal voucher is passed in the financial accounts to correct the financial errors."
StrQuestions(3) = "In a company invovled in manufacturing products, the term EOQ is used for the raw materials inventory."
StrQuestions(4) = "It an item does not entail any stock out cost, it implies that the company need not keep it as inventory."
StrQuestions(5) = "Prudent inventory policies dictate that all B category items should be fast moving items"
StrQuestions(6) = "Fast moving items should be stocked according to their value and lead time."
StrQuestions(7) = "Real time systems work well with magnetic tape based information systems."
StrQuestions(8) = "In a batch mode of operaion, it it impossible to maintain files reflecting he current status."
StrQuestions(9) = "Transcription errors of a;phaetic dta are detected by a check digit scheme."
StrQuestions(10) = "A sales invoice debits finished stock and cridits sales"

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"

End Sub

Private Sub TrueorFalse_Paper5()
StrQuestions(1) = "Paymens receipt file is a master file in a sales invoicin system."
StrQuestions(2) = "Invoice detail file is a master file in a sales invoicing system"
StrQuestions(3) = "We need onlu the invoicefile to carry out an outstanding nalysis."
StrQuestions(4) = "A contra entry is passed through the journal voucher."
StrQuestions(5) = "Control accounts are shown in the general ledger."
StrQuestions(6) = "The subsidiary ledger and general ledger are linked through th conto laccount."
StrQuestions(7) = "Trial balance is a superset of profit & loss and balance sheet."
StrQuestions(8) = "Books of account for a perprietory condern may or may not hav subsidiary ledgers."
StrQuestions(9) = "A double entry book keeping system dioes not require debits to equal credits."
StrQuestions(10) = "A voucher when posted in the bookd may or may ot affct the total of all debits."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"

End Sub

Private Sub TrueorFalse_Paper6()
StrQuestions(1) = "Variation of demand ofaproduct istaken care of by cattying large inventory."
StrQuestions(2) = "In a standard costing system rate is stored in the ransaction s file."
StrQuestions(3) = "For carrying out historical analsis, previous year consumption is trored in the iem master file."
StrQuestions(4) = "An ABC analysis shows that the maximum prpoortional value ofthe invetor is made up ofa large umber of items."
StrQuestions(5) = "In asequential file on a disk, it is not possble to access a record directly."
StrQuestions(6) = "In program documentation, it is incorrect to include a system flow chart."
StrQuestions(7) = "In indexed sequential operation, records are held in sequential order on the disk nd the index is kept permanently on the main storage for fast access."
StrQuestions(8) = "Program flow chart is also a system flow chart."
StrQuestions(9) = "Sales invoice isn ot a hournal voucher."
StrQuestions(10) = "Sales tax does ot form part o the total price paid by a consumer for a particular commodity he purchases."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper7()
StrQuestions(1) = "Ageing analysis isuseful to conrol bad debts."
StrQuestions(2) = "In a payroll system any voluntary PF cut opted for by an employee is not recorded."
StrQuestions(3) = "Reserves and suspluses costitute the liabilities sode of the balance sheet."
StrQuestions(4) = "The balance sheet of a company fails to relect the status of a company at any given point of time."
StrQuestions(5) = "Wheneer we withdraw fuds from the bak to meet our day to da expenses, we credit the bank account and debit contra A/C"
StrQuestions(6) = "A journal voucher does not affect the bank or cash book balance."
StrQuestions(7) = "Te balances ofcntra account need not always bezero."
StrQuestions(8) = "A journal voucher, when passed does not incorporatecorrections/amendments to he books of account."
StrQuestions(9) = "Afte each year need not start with the P & L statement aftesh."
StrQuestions(10) = "An ABC analusis shows that the masimum proportinal value ofthe inventory is made up of a small number of items."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper8()
StrQuestions(1) = "Value os stock an be done bu LIFO costing."
StrQuestions(2) = "Value ofstock cab ne done by average cost."
StrQuestions(3) = "Value of stock cannot be done by standard cost."
StrQuestions(4) = "TDS, not being pertinent is not takn car of during payroll processing."
StrQuestions(5) = "Pay image file is a replica  ofthe employee's payslip"
StrQuestions(6) = "After every year we mere the P&L of operations with the balance sheet."
StrQuestions(7) = "Depreciation is not considered an expense item."
StrQuestions(8) = "Purchases from suppliers do not form part ofthe subsidiary ledgers."
StrQuestions(9) = ""
StrQuestions(10) = "Vendor performance cannot be judged by price of material supplied."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper9()
StrQuestions(1) = "Vendor performance cab ne judged u timeliness of dupply."
StrQuestions(2) = "Vendor performance cannot be judged by quality of supplies"
StrQuestions(3) = "Ageing analusis of th stocks held does not help the company in optimising inventory."
StrQuestions(4) = "Sales to customers form part ofsubsidiary ledgers."
StrQuestions(5) = "A suspense account reflects balances which do not have clear cut account heads defined."
StrQuestions(6) = "A suspense accountreflects balances whosetreatment is trill to be decided."
StrQuestions(7) = "Strict inventory controls canno help minimise product obsolescence."
StrQuestions(8) = "In a situation of zero leadtime of supplies, we may not need to hold any inventory."
StrQuestions(9) = "Thetrue pictre of the inventory value is best reflected by standard cost."
StrQuestions(10) = "A field name must begin with an alphabetic character."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper10()
StrQuestions(1) = "It is not possile to search for text in a memo field."
StrQuestions(2) = "Up to 10 database files can be opened in a work area."
StrQuestions(3) = "The contents of the mem ofield are stored in the same /DBF file."
StrQuestions(4) = "MODIFY STRUCTURE automaticall ycrates a backup file(s), if required."
StrQuestions(5) = "A database file cannot be indexed on al ogical field."
StrQuestions(6) = "The REINDEX command aotomaticall yreindex all open index files for the database  file open in the current work area."
StrQuestions(7) = "The ZAP command onlu removes the records marked for deletion."
StrQuestions(8) = "When a database file ABC.dbf is opened, its index file ABC.cdx is opened automatically."
StrQuestions(9) = "Therecord numbers in the datanase file change when it is indexed."
StrQuestions(10) = "The RUN command is used to run a program file."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper11()
StrQuestions(1) = "The COPY FILE command cab cpy ultiple files."
StrQuestions(2) = "A function such as UPPER() cab ne used while sorting a database file with SORT."
StrQuestions(3) = "A function such as UPPER() cba ne used while indexing a catabase file with INDEX."
StrQuestions(4) = "A compound index file can include multiple index tags."
StrQuestions(5) = "A database file must be in use before using SORT and REPORT FORM commands."
StrQuestions(6) = "ROUND() cab ne used to round off a numbe to hundreds."
StrQuestions(7) = "LIST Command can be used to search for information in a database file."
StrQuestions(8) = "The COUNT command is usedto count the number of fields in a database file."
StrQuestions(9) = "Records deleted with DELETE cab nerecoered selectively."
StrQuestions(10) = "ABS(-13.98) returns 13.98 and ABS(6.45) returns  -6.45."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper12()
StrQuestions(1) = "While creating a report woth CREATE REPORT , each record cab occupy onlu one deetail line."
StrQuestions(2) = "If RATE is a numeric field, LIST RATE/1.1 is invalid."
StrQuestions(3) = "ALLTRIM() removes liading blands from the specified character expression."
StrQuestions(4) = "Only one data base file cab ne used while creating a report format file with CREATE REPORT."
StrQuestions(5) = "If you enter n in a logical field, it is not accepted."
StrQuestions(6) = "If the variable ABC contains a name, LEN() cab ne used to find out the length of the name."
StrQuestions(7) = "** is a symbol for exponentiation."
StrQuestions(8) = "RELEASE command is used to remove memory variables from memory."
StrQuestions(9) = "All or specified memory variables cab ne saved in disk file for later use."
StrQuestions(10) = "CHANGE and EDIT are identical commands."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper13()
StrQuestions(1) = "APPEND BLANK adds a blamk record and presents it for data entry."
StrQuestions(2) = "WAIT WINDOW" & Chr(4) & "Press a key" & Chr(4) & "NOWAIT displays the message andwaits for the user to press a key."
StrQuestions(3) = "The COPY FILE command cannot be used to copy the current databasefile."
StrQuestions(4) = "It is not possible to add records while using EDIT."
StrQuestions(5) = "EDIT cab display onlu one record on the screen."
StrQuestions(6) = "Edit or change window cab ne placed anywhere on the screen."
StrQuestions(7) = "RECALL indeletes records deleted from the atabase file with the PACK command."
StrQuestions(8) = "DISPLAY NEXT 3 TO PRINTER FOR PIN =" & Chr(4) & "110025" & Chr(4) & "FIELDS PIN,NAME is invalid."
StrQuestions(9) = "SET MEMOWIDTH changes the width of the memo window."
StrQuestions(10) = "SORT ON NAME + AMOUNT TO BANK is a valid command."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper14()
StrQuestions(1) = "The name ofstructural compound index File of ABC.dbf is ABC.dbf."
StrQuestions(2) = "SET INDEX command is used to open the structural compound indx file ofadatabasefile."
StrQuestions(3) = " If a database file ocntains 4 fields, its structural compound index file cab have a maximum of 4 indextags."
StrQuestions(4) = "One can create an index tag while defining the structure of the datanase file."
StrQuestions(5) = "REPORT FORM SAMPLE SCREEN displays the report on the screen."
StrQuestions(6) = "A field from a database cannot be included in the report page header (PageHd)."
StrQuestions(7) = "The TOTAL command condolidates totals in report."
StrQuestions(8) = "Aeport cannot use more than 4 details lines."
StrQuestions(9) = "The output of REPORT FORM cab ne displayed or printed, however, it cannot be saved in a disk file."
StrQuestions(10) = "When therecord pointer is at the last record ofthe database, file EOF() returns true."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub
Private Sub TrueorFalse_Paper15()
StrQuestions(1) = "STORE [20] TO NUM creates a character variable."
StrQuestions(2) = "A ate tupe variable cab necreated with STORE."
StrQuestions(3) = "More than one relation can be established out of a database file."
StrQuestions(4) = "For each@....SAY...GET used in a program, one must use a corrsponding READ or READ CYCLE."
StrQuestions(5) = "The output of all @...SAY...GET command issent to the active Window."
StrQuestions(6) = "@5.1,0.3 SAY" & Chr(4) & "RAJAN" & Chr(4) & " is an invalid command."
StrQuestions(7) = " Before using RQBE, you mus open the database files and establish their relations."
StrQuestions(8) = "While using RQBE< you can view the resultant SQL SELECT command before executing the query."
StrQuestions(9) = "A query cab ne saved in a file for later use."
StrQuestions(10) = "If the database file used in RQBE contains 6 fields, the output of the query always contains 6 fields."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub questions_paper1()
StrQuestions(1) = "What are the master and transaction files in a Payroll System? Show also the typical syuvture, including type and width."
StrQuestions(2) = "What  are the master and transaction files in a financial accounting system? Show alsothe tupical structure, including type andwidth."
StrQuestions(3) = "What are the amaster andtransaction files in a stores accolunting sytstem? Show also the tupical structure, including type and width."
StrQuestions(4) = "What arethe master andtransactin iles ina sales accountibg system? Show also the tpical structure, including tpe and width."
StrQuestions(5) = "What are the masterand transactin files in a sales analysis function? Show also the typical suructure, including type and width. Patients getadmitted and discharged at Hospitals. There will be paying/free patients, what are the master andtransaction files in an admission discharge sstem in a Hospital? Show also the tupical structure, including type andwidth."
StrQuestions(6) = "Patients getadmitted and discharged at Hospitals. There will be paying/free patients, what are the master andtransaction files in an admission discharge sstem in a Hospital? Show also the tupical structure, including type andwidth."
StrQuestions(7) = "Travel involves routes, clsses , concesions etc. What are the master and transaction files in arailway passengetreserfation sytem? Show also the tupical structure, including tyupe ans width."
StrQuestions(8) = "Hotels have rooms, getguests and carry out booking and allotment; they also charge for Services. What arethe masterandtransactin files in a Hotel front office i.e. boling in and booling out System? Show also the tupical structure, including tpe and width."
StrQuestions(9) = "Dolleges ofer courses, admit students and keep tah ontheir performance. What are the master and transaction files in a College record keeping System? Show also the tupical structure, including type and width."
StrQuestions(10) = "What are the master andtransaction files in a stores function? Show also the tupical structure, including tpe and width."
End Sub

Private Sub questions_paper2()
StrQuestions(1) = "Your city Dvelopment Authority(CDA) isrinning an amusement pak made up offive Sections, each housing a diferent type of ausement vez. , wate sports , indoor games, weings. lawn tennis, video prlout. What are hte master and transaction files in this fiuction? Show also the ypicalstructure, including tpe and width."
StrQuestions(2) = " Libraies acquire bools, isssue bols to etitled persons and keep record of wher they are located on the shelves. What are the mastr andtransaction files in a library system/ Show also the tupical structure, including type nadwidth."
StrQuestions(3) = " Every civic service like water, electricity, tlephone keeps a record ofthe consumer and bils them for comsumption/usage. What are the master and transactin files in a popoerty tax illing systen? Show also the tupical structure, including tupe and idth."
StrQuestions(4) = "Banks serve their customers through two streams: assistance to tellers andthe aback office. Tellers interface withthe baln's clients for dayt to at transactions on their Current/Savings accounts. What are the master and transaction files in this part of the Banking system? Show also the tupical structure, including tpe and width."
StrQuestions(5) = "Theatres and cimenas have gaded deating, many shows and have advanced/current booling. What ar the master andtransaction files at a cimena theatre's computerised system? Show also the tupial structure, including tpe and width."
StrQuestions(6) = "Briefly  describe in not more tha 100 words" & Chr(13) & _
                  "a.Control accounts" & Chr(13) & _
                  "b.Trail balance" & Chr(13) & _
                  "c. REserves and surplus" & Chr(13) & _
                  "d.ABC analysis" & Chr(13) & _
                  "e.Ageing analusis" & Chr(13) & _
                  "f.LIFO.FIFO valuation" & Chr(13) & _
                  "g. Pay image file" & Chr(13) & _
                  "h.Reorder levels" & Chr(13) & _
                  "i.Structured programming" & Chr(13) & _
                  "j. Program life cycle" & Chr(13) & _
                  "k.Systems audit" & Chr(13) & _
                  "l.Viruses" & Chr(13) & _
                  "m. Batch and on line real time" & Chr(13) & _
                  "n.EOQ" & Chr(13) & _
                  "o. Internal and external lead time" & Chr(13) & _
                  "p.BOM" & Chr(13) & _
                  "q. Stock out and buffer stock" & Chr(13) & _
                  "r. Dead stock and shelf life" & Chr(13) & _
                  "s. Indexing and sorting"
StrQuestions(7) = ""
StrQuestions(8) = ""
StrQuestions(9) = ""
StrQuestions(10) = ""
End Sub

Private Sub questions_paper3()
StrQuestions(1) = "In a STUDENT FILE each record contains thefollowing information regarding students." & Chr(13) & _
                  "Field       Type    Width" & Chr(13) & _
                  "RollNo  Ch  6" & Chr(13) & _
                  "Name        Ch  24  (Surname)" & Chr(13) & _
                  "Initl       Ch  4   (Initials)" & Chr(13) & _
                  "Sex     Ch  1   (M/F)" & Chr(13) & _
                  "Engmk       N   3" & Chr(13) & _
                  "Mathmk  N   3" & Chr(13) & _
                  "Hismk       N   3" & Chr(13) & _
                   Chr(13) & _
                  "Write a FOXPRO program t ocreate afile on totla marks and withn total marks on roll number. The total marks for each stucent is created by adding all the three subjects marks otained by the strident. If total marks are same, the record will be arranged F first and within sex, alphabetical order ofthe name. Each record ofthe sorted output file should comtain the followinginformation." & Chr(13) & _
                  "ROll Number" & Chr(13) & _
                  "Name" & Chr(13) & _
                  "Initials" & Chr(13) & _
                  "English marks" & Chr(13) & _
                  "Maths marks" & Chr(13) & _
                  "History marks" & Chr(13) & _
                  "Total marks" & Chr(13) & _
                  "Grading (D-over 81; A - ever 71; B- over 61; C - over 51; F - below 50)"
StrQuestions(2) = "In a big city , at a car park, there are charges for car parking. A minimum fee of Rs. 1.50 is charged to park a car upto one haour. After, that and additional Rs. 0.50 is charged for every half hout aor part there of in excess of 1 hour. A dedicated computer terminal is placed ateach parking meter. If parked over 12 hours, charges are douled and over 24 hours, terbled." & Chr(13) & _
                  "Owners/Drivers type in thefollowing information on a terminal cm keyboard: Meter Station ( 3 Numeric, as painted onthe Meter); Tine IN as hout and minutes(automatic when energised by system clock, when IN is typed); car number (alphanumeric); Owner or Driver(O/D); Name(25 characters). The car is automatically clamped. When chechking out on operating the terminal  the system generates OUT time, calculates the amount due and grnerates a slip asfollows :" & Chr(13) & _
                  "Meter station" & Chr(13) & _
                  "Car No(12 alphanumeric)" & Chr(13) & _
                  "Name(25 characters)" & Chr(13) & _
                  "Hours parked" & Chr(13) & _
                  "Normal amount( 4 numeric with 2 decimals)" & Chr(13) & _
                  "Additional charge( 5 numeric withtwo decimals)" & Chr(13) & _
                  "Total charge(5 numeric with two decimals)" & Chr(13) & _
                   Chr(13) & _
                  "When the driver/owner pays at the gate, the guard releases the clamp and allows the car to be taken out." & Chr(13) & _
                   Chr(13) & _
                  "Write a FOXPRO program for the central computer that handles early 100 meter stations."
StrQuestions(3) = "a. What is file updation? Write the process involved in file updation. Suppose the master file contains thefollowing fields: SUbscriber number, name, and address. The transaction record contains the following fields: Subscriber number, transaction code, name and address." & Chr(13) & _
                   Chr(13) & _
                  "Write a FOXPRO program to update the master file." & Chr(13) & _
                  "b. What is a check digit? Why is it advised?" & Chr(13) & _
                  "c. If the above updation is to use Modulo 11 for check digit, write the additional code in FOXPRO."
StrQuestions(4) = "Briefly explain in 60 word sabout any fove of the following" & Chr(13) & _
                  "a.Structured programming" & Chr(13) & _
                  "b. Portability of program" & Chr(13) & _
                  "c. Local Area Networks" & Chr(13) & _
                  "d.Viruses" & Chr(13) & _
                  "e. Indexed file organization" & Chr(13) & _
                  "f. Types of common errors in a FOXPRO program"
StrQuestions(5) = "An inventory master file is sorted on disk as a sequential file withaprt number as the key. The structure ofarecord is" & Chr(13) & _
                   Chr(13) & _
                   "Field Size" & Chr(13) & _
                   "Part number         9(12)" & Chr(13) & _
                   "Part description        X(20)" & Chr(13) & _
                   "Quantity on hand        9(3)" & Chr(13) & _
                   "Price per unit      99V99" & Chr(13) & _
                   "Reorder point       9(3)" & Chr(13) & _
                   "Reorder quantity        9(3)" & Chr(13) & _
                   "Quantity on order       9(3)" & Chr(13) & _
                   "The transaction file isalso sorted on part number as key and its format is" & Chr(13) & _
                   "Field Size" & Chr(13) & _
                   "Part number     9(4)" & Chr(13) & _
                   "Quantity        9(3)" & Chr(13) & _
                   "ACtion code     X" & Chr(13) & _
                   "If the action code in a transaction record is 'A' then the quantity isto be added to the inventory file. If the action code 'S' it is to be subtracted. When a reorder point is reached, the program sould note the item/s  involved and print out a notice to the stores in the following format:" & Chr(13) & _
                   "_____________________________________________________________________" & Chr(13) & _
                   "To" & Chr(13) & _
                   "Stores Manager" & Chr(13) & _
                   "XYZ Co" & Chr(13) & _
                   "REORDER STATUS OF INVENTORY ON DD MM YYYY" & Chr(13) & _
                   "Part No Description     Reorder level       In store Qty" & Chr(13) & _
                   "_____________________________________________________________________" & Chr(13) & _
                   "Write a FOXPRO program to update the master file using the transactin file." & Chr(13) & _
                   "Note : More that one transaction record with the same part number should be allowed. Diffrent transaction records with the same part number may have diffrent actin codes."
StrQuestions(6) = "Information on sales personel ar recorded on adisk file. Each record consists ofthe saleserson number and amount of sales done by hn/her. There cab ne more that one record for a person. All recors s pertaining to an individual ar grouped together. To mark the end of dta ofa salesperson, an additional recors, with sales zero, is provided after all the recorss pertaining to that sa;esperson." & Chr(13) & _
                   Chr(13) & _
                  "Write a FOXPRO program to calculate tand peint saleswise commission and also total sales and total commission for the person. Commission is to be calculated @ 10% of sales." & Chr(13) & _
                   Chr(13) & _
                  "Sample output is :" & Chr(13) & _
                  "_______________________________________________________" & Chr(13) & _
                  "SALESPERSON     SALE        COMMISSION" & Chr(13) & _
                  "1           ...     ......" & Chr(13) & _
                  "...     ......" & Chr(13) & _
                  "...     ......" & Chr(13) & _
                  "TOTAL     ____        ______" & Chr(13) & _
                  "____        ______" & Chr(13) & _
                  "2           ...     ......" & Chr(13) & _
                  "etc."
StrQuestions(7) = "A PAY-ROLL file contais a nuber ofrecords. Each records contains among other thins, a field called HOURS-WORKED and another field called RATE. Eac record relates to an individual employee ad TOTAL-PAY is calculated from the table that follows :" & Chr(13) & _
                   Chr(13) & _
                  "Write a FOXPRO program to read a record from PAY-ROLL, calculate TOTAL-PAY for each employee and to transfer control to a subroutine aftr each such calculation." & Chr(13) & _
                   Chr(13) & _
                  "TABLE FOR TOTAL PAY CALCULATION." & Chr(13) & _
                  "_____________________________________________________" & Chr(13) & _
                  "HOURS-WORKED            TOTAL-PAY" & Chr(13) & _
                  "_____________________________________________________" & Chr(13) & _
                  "Less than or equal to 36    HOURS-WORKED * RATE" & Chr(13) & _
                   Chr(13) & _
                  "Greater than 36 but less    36*RATE+1.5*RATE*" & Chr(13) & _
                  "than or equal to 45     (HOURS-WORKED - 36)" & Chr(13) & _
                   Chr(13) & _
                  "Greater than 45         36*RATE + 13.5 * RATE" & Chr(13) & _
                  "+ 2 * RATE *" & Chr(13) & _
                  "(HOURS-WORKED - 45)" & Chr(13) & _
                  "____________________________________________________"
StrQuestions(8) = "A consulting firm has, for each employee, an internal rate( what they pay an employee per hout) & an external rate ( what the carge customers for an hour ofthat employee's time)." & Chr(13) & _
                   Chr(13) & _
                  "Their computerised system has an input document" & Chr(13) & _
                   Chr(13) & _
                  "Employee ID         X(4)" & Chr(13) & _
                  "Internal Rate       9(3)V99" & Chr(13) & _
                  "External Rate       9(3)V99" & Chr(13) & _
                  "Weekly Hours WOrked 9(2)V99" & Chr(13) & _
                   Chr(13) & _
                  "Prin a payroll report, using FOXPRO showing:" & Chr(13) & _
                  "________________________________________________________________________" & Chr(13) & _
                  "Employee    Hours   Internal    Payroll External Amount   Gross" & Chr(13) & _
                  "    ID      Worked  Rate        Amount  Rate       Charged  Profit" & Chr(13) & _
                  "________________________________________________________________________" & Chr(13) & _
                  "where -" & Chr(13) & _
                  "Payroll Amount = hours * internal rate" & Chr(13) & _
                  "amount Charged = hours * external rate" & Chr(13) & _
                  "Gross Profit = amount charged - payroll amount" & Chr(13) & _
                   Chr(13) & _
                  "At the end of the report, print a line showing the total payroll amount, total amountcharged & the average gross profit per employee."
StrQuestions(9) = "Briefly explai the following" & Chr(13) & _
                  "a. What are basic tasks involved in business data processing?" & Chr(13) & _
                  "b. Classify dat elements by source and explain each of them briefly." & Chr(13) & _
                  "c. File security and dat security" & Chr(13) & _
                  "d.Systems Audit"
StrQuestions(10) = "Given the following record lay out ofa swquential file." & Chr(13) & _
                   "SALES-ID        X(5)" & Chr(13) & _
                   "CLOSING-DATE    X(6) (in the form of YYMMDD)" & Chr(13) & _
                   "AMOUNT-THIS-YEAR    S9(4)V99" & Chr(13) & _
                   "AMOUNT-LAST-YEAR    S9(4)V99" & Chr(13) & _
                    Chr(13) & _
                   "Write a FOXPRO program to print the following information." & Chr(13) & _
                   "SALES-ID    AMOUNT  AMOUNT      INCREASED" & Chr(13) & _
                   "THIS YEAR   LAST YEAR       AMOUNT" & Chr(13) & _
                    Chr(13) & _
                   "At the end of the report, print he following statistics" & Chr(13) & _
                    Chr(13) & _
                   "Total sales this year :" & Chr(13) & _
                   "Total sales last year :" & Chr(13) & _
                   "Averge sales this year :" & Chr(13) & _
                   "Average sales last year :" & Chr(13) & _
                   "Maximum increase in sales :" & Chr(13) & _
                    Chr(13) & _
                   "Note : Increased Amount = (Amount This Year - Am"
End Sub

Private Sub questions_paper4()
StrQuestions(1) = "Briefly explain in 60 words about any five ofthe following " & Chr(13) & _
    "a. Deine and diggerentiate between data entry and data sorage as elements ofan information system." & Chr(13) & _
    "b. Need for dat abackup in a data processing center." & Chr(13) & _
    "c. Distinguish between LAN and WAN" & Chr(13) & _
    "d. Need for SORTING and MERGING of sata files." & Chr(13) & _
    "e. Flow charting rules" & Chr(13) & _
    ".Proram documentatiion"
StrQuestions(2) = "A real estate company maintains a masterfile with the followig record structure." & Chr(13) & _
    "Plot Number" & Chr(13) & _
    "Owner 's name" & Chr(13) & _
    "Assessed Value" & Chr(13) & _
 Chr(13) & _
    "A tranasaction file with the following setails is used to update the master fle :" & Chr(13) & _
 Chr(13) & _
    "Plot Number" & Chr(13) & _
    "Code: 1 - Change owner; 2 - Change assessed valuation" & Chr(13) & _
        "3 - Change both name and valuation" & Chr(13) & _
        "4 - Delete the master record" & Chr(13) & _
    "Name" & Chr(13) & _
    "Assessed Value" & Chr(13) & _
    "Write PSEUDOCODE fo updating the above masterfile with the given transaction file"
StrQuestions(3) = "Briefly explain, in 6 0words, any five , of the following" & Chr(13) & _
    "a. Distinguish between indexed andrelative file organisation of DOXPRO files" & Chr(13) & _
    "b. Distinguish between an algorithm and pseudocode." & Chr(13) & _
    "c. Distinguish between distributed data processing and centralised data processing" & Chr(13) & _
    "d. Client-server operations." & Chr(13) & _
    "e.Define.Data Item, record, file" & Chr(13) & _
    "f. Data security in a data processing center"
StrQuestions(4) = "Write a pseudocode fora program to prepare a simple balancing of a chequebook transactions. The prpgram should take - " & Chr(13) & _
    "- the old balance," & Chr(13) & _
    "- a set of transaction-codes that follow )'D' for debit i.e. withdrawal and 'C' for credit i.e. deposit) and" & Chr(13) & _
    "- amounts until end of file." & Chr(13) & _
 Chr(13) & _
    "Output thetotal-deposits , total withdrawals and new balance."
StrQuestions(5) = "a. Suppose a supplier ives a 0.05 volume, discount for purchass of 2000 units or more. Let ITEM, QUANTITY and PRICE bethe names of the storage locations that contain the data." & Chr(13) & _
 Chr(13) & _
    "Write a FOXPRO program to produce the following report at the end of the day:" & Chr(13) & _
    "__________________________________________________" & Chr(13) & _
            "BRAHMA & CO" & Chr(13) & _
    "__________________________________________________" & Chr(13) & _
        "DISCOUNT LIST FOR DD MM YYYY" & Chr(13) & _
    "ITEM        QUANTITY    PRICE   DISCOUNT" & Chr(13) & _
    "____        ______  ___ _____" & Chr(13) & _
    "____        ______  ___ _____" & Chr(13) & _
    "__________________________________________________" & Chr(13) & _
    "TOTAL" & Chr(13) & _
    "__________________________________________________" & Chr(13) & _
 Chr(13) & _
    "b. An organisation markets two types of products denoted by the numbers 1 and 2. Let there be two categories of customers - regular and casual denoted bu the numbers 1 and 2 respectively. Develop a DECISION TABLE for computing the rate ofcommission on purchases, a customer is entitled to get."
StrQuestions(6) = "BOOKMAN & CO offers trade discount to establised bool sellers and other s for anewly released best seller :" & Chr(13) & _
    "- 20% fo private sustomers and librariesm; 5% discount is allowed on orders for 6 books or more ;" & Chr(13) & _
    "- 10% on orders for 20 books or more; 15% of orders for 60 or more" & Chr(13) & _
    "Trade ordres for 20 books nad over receive 10% discount in addition to the trade discount. Sukhram, the proprietor, keeps a special sales register with columns as below" & Chr(13) & _
    "BOOLMAN & CO REGISTER OF SALES AND DISCOUNTS OF 'DEATH LOOMS'" & Chr(13) & _
"__________________________________________________________________________" & Chr(13) & _
    "Date.       Sale    Customer Name   Order           Discount" & Chr(13) & _
            "code                Qty.        Normal  Special" & Chr(13) & _
      "___     ___ ____________    ___     ____        _____" & Chr(13) & _
"__________________________________________________________________________" & Chr(13) & _
    "TOTAL                               ____        _____" & Chr(13) & _
"__________________________________________________________________________" & Chr(13) & _
    "He wants to put this system of a computer and tells you that the sidths othe various headings are as Date ddmmyyyy Code        1 Alphanumeric( P private, L Library, B Booksellers, M Miscellaneous)" & Chr(13) & _
    "Name        40 Alphanumeric Order       9(4) Rate        Rs. 250.00 (Introductory offer for year)" & Chr(13) & _
    "N Disc  999V99 S Disc  9999V99" & Chr(13) & _
    "He wants thefollowing outputs starting 01 01 1998 which is a Thursday" & Chr(13) & _
    "- a weekly report o sales and discounts o the two types and also the total discounts in the following form every Monday, Sunday being a holiday -" & Chr(13) & _
    "Week ending     Sales       Discounts given      Net sale        Normal        Special TOTAL" & Chr(13) & _
    "- a weekly report on stock balance assming that he printed 25000 bools for sale, showing a reprint warning when the stock falls to 5000 in the following form as of saturdat evening closing -" & Chr(13) & _
    "Week ending     Stock position  Reprint     Opening      Closing   Y/N" & Chr(13) & _
    "The number is to be given a royalty of 5% on net sales every six months and Sukhram wants a report in the following format :" & Chr(13) & _
    "Month ending    TOTAL SALES     ROYALTY DUE" & Chr(13) & _
    "( JUN/DEC)" & Chr(13) & _
    "a. Show the structure of the master and transaction file/s  b. Draw a program flow chart/s for any two outputs."
StrQuestions(7) = "For the last question, a. Write the pweudocode or all programs. b. Write a FOXPRO program for any two outputs."
StrQuestions(8) = "For the last Question,  " & Chr(13) & _
    "a. Write the pseuocode for the programs generating the outputs" & Chr(13) & _
    "b. Write a program in FOXPRO for any two."
StrQuestions(9) = "HUMBI Transport Corporation owns 200 buses that run on 40 inter city routes foxed bu it, to carry passengers to towns nearby. The rules fortravel are as follows :" & Chr(13) & _
    "- each route had 5 buses asigned to it, one owhich will be under routing maintenance; these are so arranged that the availiability  is known well beforehand; intheevemt of a us going in for repair, the spare bus will be put into service andthe tickets issued ot the sick bus are valid for the replaced one; the passenger is informed b telephone if notified;" & Chr(13) & _
    "-each bus seats 50 passengers ofwhich 10 are fitst class;" & Chr(13) & _
    "- fares are in slabs of 25Km, Rs 10 for the fitst and Rs 5 for every additional 20 Km or less; no oute is over 125 Km and all bbusesreturn the same day;" & Chr(13) & _
    "- there are onlu two services a day, the first at 6 AM and the second at mid day." & Chr(13) & _
    "- there is aroute directory that lists the towns served and istances involved;" & Chr(13) & _
    "- return journeys atrte at 50% if booled at the same time as the outward hourney;" & Chr(13) & _
Chr(13) & _
    "a) For reservation onthis service, using a computer, whatare the Master and Transaction files tha you will plan?" & Chr(13) & _
    "b) Show the stucture of these files." & Chr(13) & _
    "c) What would be the flow chart for a program to make reservations ?"
End Sub

Private Sub questions_paper5()
StrQuestions(1) = "After  assessing the creditworthiness of the orrower by looking at various proofs ofhis technical and financial credibility, the board of SHARK & Co, a leading Fnancial Institution ofering loag term financial assistance in Indian Rupees, enters into a loan agreement with the borrower . The terms and condtions include the following:" & Chr(13) & _
  Chr(13) & _
    "- The ammount of the loan" & Chr(13) & _
    "- The number of instalments and the dates on which they are repayable" & Chr(13) & _
    "- The rates o interest for the loan" & Chr(13) & _
    "- Penal interest for overdue bills" & Chr(13) & _
  Chr(13) & _
"On all the dates on which principal amounts are due, the borrower is sent a bill for principal and penal interest if any (If any amount are overdue the penalties for these amounts are also levied.). Whenever receipts are received from the customer, theredeipuys are automaticall yadhusted with the customer's unpaid bill. Oldest bills are adjusted first. In case no amount is outstanding the receipy amount is retained with the compaay and adjusted against future bills" & Chr(13) & _
  Chr(13) & _
    "Draw a decision table for preparing thebill."
StrQuestions(2) = "In the above problem, assume that Mr Shark has computerised the systenm, with among others, one master file, one transaction file and one print file structured as below and inputs as shown" & Chr(13) & _
    "Master file has the following structure:" & Chr(13) & _
    "borrowno N(7); Borchkdi; 9 Check digit on Modulo 11; Borrini X(10); Borrname X(30)" & Chr(13) & _
    "Boradd1 X(25); Boradd2 X(25);Boradd X(25);Borcity X(20); Borstate X(20);Borpincd N(6); Borstdcd N(7) STD" & Chr(13) & _
    "Bortele1 N(7);Bortele2 N(7);Borfax N(7);Boremail X(30);creditok X; Loan 9(6)V99;instalmt 99; nxtindat        ddmmyyyy(Date next instalment);inttopay 999V99" & Chr(13) & _
    "penalty     99(percentage if overdue)" & Chr(13) & _
    "The transaction file is for payments made" & Chr(13) & _
    "borrowno N(7);borchkdi 9;datepaid ddmmyyyy; amtpaid 9(3)V99" & Chr(13) & _
    "The print file to print the bill to borrower" & Chr(13) & _
    "borrowno N(7);borchkdi 9;instalmt 99;inttopay 999V99;penalty 999V99;total 999V99;datedue ddmmyyyy" & Chr(13) & _
    "a) Write a FOXPRO program to update the master file when a patment  is received.; b) Write a FOXPRO program to write a leter to the borrower in the following format" & Chr(13) & _
                "SHARK & CO THE FINANCE COMPANY TO TURN TO FOR" & Chr(13) & _
                              "ANY FINANCING AID" & Chr(13) & _
"    Naya Shylok         Gift Edge" & Chr(13) & _
    "FA & CO         Broke Avenue" & Chr(13) & _
    "Ph: 999999          Oodlespuram" & Chr(13) & _
    "FAX:999989          Goofy Pradesh" & Chr(13) & _
    "Ref.S&CO 35/62/     dd/mm/yyyy" & Chr(13) & _
    "Dear Sir," & Chr(13) & _
    " Your nex instalment is due on ______ and the amount due is _________ mde up as below:" & Chr(13) & _
    "Normal Interest         : ___________________" & Chr(13) & _
    "Penalty             : ___________________" & Chr(13) & _
    "Total                   : ___________________" & Chr(13) & _
    "Assuring of our services always " & Chr(13) & _
    "              Yours faithfully"
StrQuestions(3) = "Write short notes on" & Chr(13) & _
    "a. Double entry system of accounting" & Chr(13) & _
    "b. Stages of payroll processing 5 marks each"
StrQuestions(4) = " Discuss how the following indicuduals in organisations may benefit from computr usage:" & Chr(13) & _
    "a.Sales manager" & Chr(13) & _
    "b.Personnel manager" & Chr(13) & _
    "c.Accountant" & Chr(13) & _
    "d.Project manager"
StrQuestions(5) = "Write short notes in about 100 words each on any four of the following :" & Chr(13) & _
    "a. Reserve and surplus" & Chr(13) & _
    "b. Profit/loss carried to the balance sheet" & Chr(13) & _
    "c. Current assets/fixed assets" & Chr(13) & _
    "d. Stock - out cost" & Chr(13) & _
    "e.TDS"
StrQuestions(6) = "A modern department store empluys over 300 persons. As corporat strategy , it was decided to intrudoce a computer based system to mabnage the human resources. Thescope oftheproposed system wourd cover, among others, thefollowing attributes" & Chr(13) & _
    "- leave status ie. the leave availed of by the employees ( the company permis 30 days earned leave, 7 days casual leave and 10 days medical leave);" & Chr(13) & _
    "- loans & advances taken by the emplotees and recoveries thereof;" & Chr(13) & _
    "- transfer from one counter to another, where  ever applicable( the store has codd a totla of 20 counters)" & Chr(13) & _
    "- basic qualifications and skills etc acquired" & Chr(13) & _
    "The system is planned to give information onthe above parameters in an on line query mode" & Chr(13) & _
    "a) What major fields wourld tha mster file (list 15) and one transcationfile contain?" & Chr(13) & _
    "b) Write the pseudocode to select and individual for transfer from A counter to B counter who has no leave left and has taken no advances in the year."
StrQuestions(7) = "Give ten important dat fields of the cash voucher in a computer based integrated financial accounts systenm."
StrQuestions(8) = "Write short notes in about 100 words each of any foiur of the following:" & Chr(13) & _
    "a. LIFO & FIFO mathods of stock valuation" & Chr(13) & _
    "b. Journal entry for posting payroll data in financial accounts" & Chr(13) & _
    "c. Debtors control A/c in financial accounts" & Chr(13) & _
    "d. Balance carry forward for new ear financial accounts data" & Chr(13) & _
    "e.Buffer stock" & Chr(13) & _
    "f.Vendor Rating"
StrQuestions(9) = "A company makes six different types ofbath soaps called A,B,C,D,E and F. These soaps have multiple packings X, Y and Z for each tupe of soap" & Chr(13) & _
 Chr(13) & _
    "The company has set up an extensive retail outlet. For the purpose of monitoring and control. following procedure has been adopted:" & Chr(13) & _
    "- theenitire countr isdivided into four zonts viz, North, West, East and South and there are 6,7,5 and 9 retail out - lets in each zone identified as M,N,O,P,Q,R,S,Tand U prefixed by the zone initial (N,E,W,S);" & Chr(13) & _
    "- each zone is headed bu a xonal manager ;" & Chr(13) & _
    "- company has apponited sales representatives in each zone." & Chr(13) & _
    "These representatives deal/work directly withthe retail outlets to sell / promote their product." & Chr(13) & _
    "The last one uyear has seen increased competition and pressure on the sales of the company products. The management hastherefore decided to computerise its sales reporting system" & Chr(13) & _
    "a. List thefields uou would expectto find in a masterfile ofthe company" & Chr(13) & _
    "b. Prepare a decision tableto progectthesles persons who are" & Chr(13) & _
    "- female," & Chr(13) & _
    "- from East Zone," & Chr(13) & _
    "- handling sales of any fory bath soaps but no procust G;" & Chr(13) & _
    "- delaing withanyretail outlet but not E4 & E7"
StrQuestions(10) = "List two iportant files of an inventory managment system along with five dat afields ofeach of these."
End Sub

Private Sub questions_paper6()
StrQuestions(1) = "The payroll system in an organisation adopts the following policy" & Chr(13) & _
    "- no houserentallowanceswould be giveb ifthe empliyee is absent formore tha 10 days in a month" & Chr(13) & _
    "- the employees contribution to providentfund would not be applicable ifthe employee does ot work for a continous stretch of more than 2 weeks ina month" & Chr(13) & _
    "- if lave is taken on both preceeding and succeeding days o announced holidays, the leave account would include the bracketed holidays." & Chr(13) & _
 Chr(13) & _
    "a) Draw a decision table reflecting this policy." & Chr(13) & _
    "b) Name twelvefields you would expect to find in the Master file."
StrQuestions(2) = "In the above problem, itis intended to calculate the waes due to an employee, in Oct, under thefollowing assumed sutuations:" & Chr(13) & _
    "i) emoluments-basic salary is Rs 2000, house rent subsidy is Rs 1000, employees contribution is Rs 200 per month ;" & Chr(13) & _
    "ii) scheduled holidas are - 2nd Oct. Gandhi Jayanti, 16th , 8th, 15th, 22nd, and 29th October;" & Chr(13) & _
    "iii) absence from duty - 1st to 7th suffixing 15;24 to 28 with prefix and sufic o holidays." & Chr(13) & _
  Chr(13) & _
    "a) Draw a flow chart forth program;" & Chr(13) & _
    "b) Write pseudo code for it."""
StrQuestions(3) = "It has been decided to compteise the stores function withthe prime objective of ontaimimg the currentand updted information on stock status. The designed computer based system sould, necessarily, be able to generte the stores ledger in the format at Figure 1, for all items in the store." & Chr(13) & _
 Chr(13) & _
        "FIGURE 1: STORES LEDGER" & Chr(13) & _
"____________________________________________________________" & Chr(13) & _
    "Item code________       Item Description______________" & Chr(13) & _
    "___________________________________________________" & Chr(13) & _
"____________________________________________________________" & Chr(13) & _
    "Date    Issue/  Qty Value   CLOSING BALANCE" & Chr(13) & _
        "Receipt         ________________" & Chr(13) & _
                        "QTY Value" & Chr(13) & _
"____________________________________________________________" & Chr(13) & _
"a. Design the data structure for masterfiles and transaction files as may berequired." & Chr(13) & _
    "EXHIBIT INV.1:  Item Master Structure:" & Chr(13) & _
"IT_CODE 6   N   Item code" & Chr(13) & _
"IT_DESC 30  Ch  Item description" & Chr(13) & _
"UNIT_MES    15  Ch  Measure of unit" & Chr(13) & _
"STD_COST    6.2 N   Standard cost" & Chr(13) & _
    "EXHIBIT INV.2: Transactions:" & Chr(13) & _
"TR_DATE 8   Date    Date of transaction" & Chr(13) & _
"TR_NO       4   N   Transaction n" & Chr(13) & _
"TR_TYPE 1   Ch  Transaction tyupe (I for ISsue, R Receipt)" & Chr(13) & _
"IT_CODE 6   N   ITem code" & Chr(13) & _
"IT_QTY  4   N   Quantity" & Chr(13) & _
"IT_VALUE    12.2    N   Value" & Chr(13) & _
"b. Given the data structures at Exhibits INV.1  , draw a mudular flow chart to print the stores ledger; c. Write FOXPRO code to generate the stores ledger."
StrQuestions(4) = "You have recently joined a Puublinc limited company. The financedirector has adked you to desig and intouduce a new voudher format ( since theearlier format has led to inadequate recording of details)." & Chr(13) & _
    "Having designed thevoucher format, you are adked to  computerise the coucher recording." & Chr(13) & _
    "a. List the important data elements that you would include in yourfile; you must specify the data type and width." & Chr(13) & _
        "EXHIBIT FIN.1 - VOUCHER FILE" & Chr(13) & _
"_____________________________________________________" & Chr(13) & _
        "VOUCHER FILE" & Chr(13) & _
"_____________________________________________________" & Chr(13) & _
    "VR_NO       4   N   Voucher no" & Chr(13) & _
    "VR_DT       8   Date    Voucher date" & Chr(13) & _
    "GL_CODE 6   N   General ledger code" & Chr(13) & _
    "D_C_CODE    1   Ch  Debit/Credit code" & Chr(13) & _
    "PARTCULR    30  Ch  Particulars" & Chr(13) & _
    "AMOUNT  12.2    N   Amount" & Chr(13) & _
"_____________________________________________________" & Chr(13) & _
    "EXHIBIT FIN.2   ACCOUNT MASTER" & Chr(13) & _
"_____________________________________________________" & Chr(13) & _
        "ACCOUNT MASTER" & Chr(13) & _
"___________________________________________________" & Chr(13) & _
    "GL_CODE 6   N   General ledger code" & Chr(13) & _
    "GL_NAME 30  Ch  General ledger name" & Chr(13) & _
    "GL_TYPE 1   Ch  Asset/Liability or" & Chr(13) & _
                    "Income/Expenditure head." & Chr(13) & _
    "b. Given the data structure at exhiot FIN.1 , draw aflowchart to append arecord in the file(you would enter the record from thekeyboard, field byfield. The flow chart should be modualr in nature as far as possible."
StrQuestions(5) = "Assuming that you are given the structure ofthe general ledger master file(Exhibot FIN.1), print thereport at figure 1(from the data ocntained in FIN.1 and FIN.2)" & Chr(13) & _
    "FIGURE 1: Output Format" & Chr(13) & _
"_____________________________________________________" & Chr(13) & _
    "GL code GL description  TOTAL BALANCE" & Chr(13) & _
                        "DEBIT CREDIT" & Chr(13) & _
"c. Draw a flow chart to write the program." & Chr(13) & _
"d. Code the program in DOXPRO, the program should be writte in a modular a\manner with separate functions to computer the total balances field."

End Sub

Private Sub questions_paper7()
StrQuestions(1) = "Which field will be most appropriate to store the following informatin in a dtabase file(table)?" & Chr(13) & _
                  "a. Name of a prohect" & Chr(13) & _
                  "b. Unit sale price of an item" & Chr(13) & _
                  "c.Employee code" & Chr(13) & _
                  "d. Sex of  worker" & Chr(13) & _
                  "e.Interest Rate" & Chr(13) & _
                  "f. Date of issue of a letter" & Chr(13) & _
                  "g.Project details" & Chr(13) & _
                  "h.Scanned Image" & Chr(13) & _
                  "i. A very large number sed in engineering calculation"
StrQuestions(2) = "Explain thefollowing in brief:" & Chr(13) & _
                  "a. Structureal compund index fole" & Chr(13) & _
                  "b.Rushmore technology"
StrQuestions(3) = "A database file BANK.dbf uses an index file BANK.cdx. If a user opens the database file without opening the index file and modifies a few records, will it correct the index file? If yes, how do you correct the index file, and if no , why?"
StrQuestions(4) = "Explain the various tupes of index files used in FOxPro."
StrQuestions(5) = "Write down the function available in FOxPro that returns:" & Chr(13) & _
                  "a. The current time" & Chr(13) & _
                  "b. The current date" & Chr(13) & _
                  "c. Name of a month from a dat expression" & Chr(13) & _
                  "d. Day of week from a date expression" & Chr(13) & _
                  "e. Minimum oftheree numbers" & Chr(13) & _
                  "f. LEngth of a character expression"
StrQuestions(6) = "NUM contains 14567.6271. Write down a functiuon that reounds off NUM to" & Chr(13) & _
                  "a.whole Number" & Chr(13) & _
                  "b. 2 decimal places" & Chr(13) & _
                  "c. Hundreds(i.e., 14600 is returned)"
StrQuestions(7) = "A database file ADDRESS contains name and address of emplouyees working in a company. Write a program to accept the nam of a person in a variable and check if it exists in the database. If yes, display the NAMe and AGE from the database, otherwise display a message -' Name not found'."
StrQuestions(8) = "a. The currently slected work area is 2, and SALES.dbf is open in work area 3. Write doen comomand(s) to sisplay NAME field from SALES.dbf without changing the work area." & Chr(13) & _
                   "b. How many database files cab ne open in each work area?" & Chr(13) & _
                   "c. Write down a command to select the third work aea."
End Sub

Private Sub questions_paper8()
StrQuestions(1) = "A database file SALES.dbf uses an index ile SALES.cdx. WRite down a command to copy the database file as wel as the index file to COMPANY.dbf and COMPANY.cdx using the COPY TO command."
StrQuestions(2) = "a. You are required to accept the age of person in the AGE variable. Write down a command such that the user can enter the age in the range 0 to 99(both inclusive)" & Chr(13) & _
                  "b. ADDRESS.dbf contians a field NAME. Write down commandsto open this database file and to display the NAMe field at coordinates 5,0."
StrQuestions(3) = "It isrequired o accept the maxium day temperature which isa variable TEMP. Ensure thatthe user cab enter onluthetempertre intherange 0 through 50. Ifthe temperature entred by the user is up to 25, display - ' Not Hot Weather',iotherwise display - 'Hot Weather'. Use the minimum number of commands."
StrQuestions(4) = "Consider the following program. Whatwill be displayed in response to the last two commands?" & Chr(13) & _
                  "Name = 'Sunder'" & Chr(13) & _
                  "CHOICE = Name" & Chr(13) & _
                  "Print Name" & Chr(13) & _
                  "?& CHOICE"
StrQuestions(5) = "Explain in brief the use of the following in a Foxpro program" & Chr(13) & _
                  "a. IF - ELSE - ENDIF" & Chr(13) & _
                  "b. DO WHILE - ENDDO" & Chr(13) & _
                  "c.SCAN -ENDSCAN"
StrQuestions(6) = "Explain the use of LOOP and EXIT in the DO WHILE - ENDDO truture in a FOxPro program."
StrQuestions(7) = "Write down a progrm to do the following" & Chr(13) & _
                  "a. Open a database file SAMPLE.dbf in work area 2." & Chr(13) & _
                  "b. Ceaate a window feom coordinates 5,5 to 20,20." & Chr(13) & _
                  "c. Display thefirst 10 records in the window degined in the previous step."
StrQuestions(8) = "Explain the use of APplication Generator. Also explain, in beirf the features provided inthe application generated by application generator."
StrQuestions(9) = "ADDRESS.dbf contains the following fields : NAME, ADDRESS, CITY and PIN. ADDRESS .dbf uses an index file ADDRESS.cdx with a tah NAME onthe NAME field. The NAME field is 20 characters wide. It is required to accept a record, it should delete that record . If no matching record is found, it should display a message- 'Record not found'"
StrQuestions(10) = "A database file BANK.dbf contains a memo field. The .CDX index file for this database contains two index tags. Write down commandsto copy the structure ofthis file to ACCOUNT.dbf. Ensure that the corresponding index and memo files are also available for ACCOUNTS.dbf."
End Sub

Private Sub questions_paper9()
StrQuestions(1) = "How can you" & Chr(13) & _
                  "a.embed" & Chr(13) & _
                  "b. link a picture created in Paintbrush in a Foxpro field."
StrQuestions(2) = "Explain the use of a memo field in a FoxPro database file. What are the operations that can beperformed on a memo field?"
StrQuestions(3) = "A label format file EMPLOYEE.lbl has been created to print address labels for the EMPLOYEE.dbf database file. EMPLOYEE.dbf contains NAME<CITY and PIN fields. All fields ar of characte tupe. How will you perform the following:" & Chr(13) & _
                  "a. print address labels for employees lining in BOmbay." & Chr(13) & _
                  "b. preview addresslabels for emplotees who live inthe PIN area 400020 or 400030."
StrQuestions(4) = "ADDRESS.dbf contains NAME,LOCALITY,CITY and PIN fields. Find out what is wron inthe following program." & Chr(13) & _
                  "USE ADDRESS" & Chr(13) & _
                  "SCAN FOR CITY = 'New Delhi'" & Chr(13) & _
                  "Print Name" & Chr(13) & _
                  "Print ALLTRIM(CITY)" & Chr(13) & _
                  "Skip" & Chr(13) & _
                  "ENDSCAN" & Chr(13) & _
                  "Close DATABASES"
StrQuestions(5) = "Explain, in brief th euse of the TEXT-ENDTEXT structure in a FoxPro program"
StrQuestions(6) = "How will you display text in a specified font and style with the @...SAY comand? Can you use font with the ? command?"
StrQuestions(7) = "How do you create a custom screen for adding and editing records? How do use the screen file to edit an existing record?"
StrQuestions(8) = "What do you understand by the term - databasefile structure? Briefly describe the various types of fields supported in FoxPro."
StrQuestions(9) = " a. What is the difference between LIST and DISPLAY ALL?" & Chr(13) & _
                  "b. A database file SALES.dbf contains the following fields: NAME,DEPARTMENT, REGION and TARGET. Write down commands to print NAME, REGION and TARGET for those salesmen sho work in 'Consumer'DEPARTMENT"
StrQuestions(10) = "a. What isthe difference between APPEND and APPEND BLAN?" & Chr(13) & _
                   "b. While isin the APPEND command, it is required toautomatically copy the contents of the previous record to the new record. How do you achieve this? How will you copy the selectd fields to the new record?"
End Sub

Private Sub questions_paper10()
StrQuestions(1) = "a. Briefly explai the term - EOF()" & Chr(13) & _
                  "b. How do you add a new record while using the BROWSE command?" & Chr(13) & _
                  "c How do you delete and recall a record whileusin the BROWSE command?"
StrQuestions(2) = "a. Adatabasefile isopen andthestatus line displays 'rec: 4/10:. If you enter DELETE REST, what will happen?" & Chr(13) & _
                  "b. Whatarethe mahor tasls you ca do with the Browse mwnu?"
StrQuestions(3) = "a. What are the differences between ZAP and PACK commands?" & Chr(13) & _
                   "b. What are thedifferences between LOCATE & SEEK commands?"
StrQuestions(4) = "Briefly descibe thedifferences between DORT and INDEX commands."
StrQuestions(5) = "ADDRESS.dbf contains ficve fields: NAME, ADDRESS, CITY PIN and PHONE. How will you create index tags inthe ascending order on NAMe and CITY fields? How will you select thetag on the CITY firld as he controlling index?"
StrQuestions(6) = "a. Briefly explain the use ofthe SET FILTER command." & Chr(13) & _
                  "b. What is a view file? What is the tupical use of the view file in FoxPro?"
StrQuestions(7) = "List the various features that ou cab use in FoxPro reports."
StrQuestions(8) = "Briefly explain how you cab achieve the following while using REPORT FORM/ CREATE REPORT:" & Chr(13) & _
                  "a. Including selected recrords in report" & Chr(13) & _
                  "b. Removing page header from the secondandsubsequent pages" & Chr(13) & _
                  "c. Displaying the current date inthe page header." & Chr(13) & _
                  "d. Saving the report in a disk file."
StrQuestions(9) = "Explain in brief the TOTAL and JOIN commands."
StrQuestions(10) = "Explain the SET RELATION comamnd. What happens when therecord pointer moves in the 'parent' file after establishing the relation?"
End Sub

Private Sub questions_paper11()
StrQuestions(1) = "ADDRESS.dbf contains NAME, LOCALITY, CITY and PIN fields. Consider the following program and rewrite it afte correcting mistakes." & Chr(13) & _
                  "USE ADDRESS" & Chr(13) & _
                  "SET TALK TO ON" & Chr(13) & _
                  "Do While Not EOF" & Chr(13) & _
                  "IF CITY = NEW DELHI" & Chr(13) & _
                  "Print Name" & Chr(13) & _
                  "Print LOCALITY" & Chr(13) & _
                  "Print CITY" & Chr(13) & _
                  "SKIP NEXT RECORD" & Chr(13) & _
                  "ENDDO" & Chr(13) & _
                  "Close ADDRESS"
StrQuestions(2) = "A databasefile PERSON.dbf contains the fields NAME, BASIC and HRA. Write down a program to print areport with emplotee name and total salaty for those employees whose BASIC is more than 2000. total Salaty = BASIC + HRA. In the end, the program should also print the total number of employees who getatotal salary more than 2000."
StrQuestions(3) = "There aretwo database files - STAff.dbf and ACCOUNTS.dbf. Both these files contain a field EMP_CODE. Write down theecessaty commandsto do the following :" & Chr(13) & _
                  "a. Open STAFF.dbf in work area 2 and ACCOUNTS.dbf in work area 3." & Chr(13) & _
                  "b. Index these files on the EMP_CODE field." & Chr(13) & _
                  "c. Link ACCOUNTS.dbf to STAFF.dbf on the EMP_CODE field."
StrQuestions(4) = "a. How many procedure files cab neused woth a program?" & Chr(13) & _
                  "b. How do you select a new procedure file?" & Chr(13) & _
                  "c. How do you close a procedure file?" & Chr(13) & _
                  "d. Can you write a procedure in any other file? If yes, how?"
StrQuestions(5) = "a. Explain, how you cab extract a specified number ofcharacters from the left of a character expression." & Chr(13) & _
                  "b. Explain PROPER() in brief."
StrQuestions(6) = "a. Explain in brief the REPLACE command and its usefulness." & Chr(13) & _
                  "b. Explain the SET DATE command in brief."
StrQuestions(7) = ""
StrQuestions(8) = ""
StrQuestions(9) = ""
StrQuestions(10) = ""
End Sub

Private Sub multiplechoice_Paper1()
strmultiple(1) = "Data validation is generaly accompished by making use of various types of totals"
strmuloption(1) = "Valiation"
strmuloption(2) = "Verification"
strmuloption(3) = "Control"
strmuloption(4) = "None of the above"

strmultiple(2) = "Outstandings analysis is carried out using"
strmuloption(5) = "Invoice file"
strmuloption(6) = "Invoice file and payments receipt file"
strmuloption(7) = "Payments receipt file and customer detail file"
strmuloption(8) = "None of the above"

strmultiple(3) = "Identify the invalid field name out ofthe following"
strmuloption(9) = "'RS.35.89'"
strmuloption(10) = "'CURRENCY'"
strmuloption(11) = "PQR 32"
strmuloption(12) = "'32 PQR'"

strmultiple(4) = "Which ofthe following is true about a disk file where WRITE statement will always place a record inthe next physical position?"
strmuloption(13) = "Sequential file"
strmuloption(14) = "Relative file"
strmuloption(15) = "Indexed file"
strmuloption(16) = "None of the above"

strmultiple(5) = "Which of the following constitute master files ina sales invoicing system?"
strmuloption(17) = "Customer details file"
strmuloption(18) = "Invoice detail file"
strmuloption(19) = "Payments receipt file"
strmuloption(20) = "All of the above"

strmultiple(6) = "In adata processing system, if the iput dat ais processed, as the transaction occurs, it is called"
strmuloption(21) = "Parallel"
strmuloption(22) = "Real time"
strmuloption(23) = "Batch"
strmuloption(24) = "Virtual"

strmultiple(7) = "Thetype of analytical tool which is useful for portraying all the possible combinations of results which cab ne encountered in a checking procedure is"
strmuloption(25) = "Confirm table"
strmuloption(26) = "Flow chart"
strmuloption(27) = "Decision table"
strmuloption(28) = "Pseudo table"

strmultiple(8) = "A sales invoice"
strmuloption(29) = "Debits customers and credits sales"
strmuloption(30) = "Credits customers and debits sales"
strmuloption(31) = "Debits finished stock andcredits sales"
strmuloption(32) = "None of the above"

strmultiple(9) = "A digit sufixed to a number for ensuring accuracy ofthe item is th"
strmuloption(33) = "Test digit"
strmuloption(34) = "Check digit"
strmuloption(35) = "Digital aspect"
strmuloption(36) = "None of the above"

strmultiple(10) = "A File Management system can typically access records from"
strmuloption(37) = "Only one file at a time"
strmuloption(38) = "Many files at a time"
strmuloption(39) = "Only twofiles at a time"
strmuloption(40) = "None of the above"


strmulanswer(1) = "Control"
strmulanswer(2) = "Invoice file and payments receipt file"
strmulanswer(3) = "PQR 32"
strmulanswer(4) = "Sequential file"
strmulanswer(5) = "Customer details file"
strmulanswer(6) = "Real time"
strmulanswer(7) = "Decision table"
strmulanswer(8) = "Debits customers and credits sales"
strmulanswer(9) = "Check digit"
strmulanswer(10) = "Only one file at a time"
End Sub

Private Sub multiplechoice_Paper2()
strmultiple(1) = "Edit checks and controls are bult into a computer based system to"
strmuloption(1) = "Ensure that transcription errors do not corrupt the result "
strmuloption(2) = "Duplicate data does not enter into the computer system"
strmuloption(3) = "Avoid transpostion errors"
strmuloption(4) = "All of the above"

strmultiple(2) = "Check lists aregenerated by the computer programs to"
strmuloption(5) = "Check manually the input data fed into the computer against input forms data"
strmuloption(6) = "Form part of the computerised processing stage."
strmuloption(7) = "Serve as aid to the management control and monitoring."
strmuloption(8) = "None of the above"

strmultiple(3) = "Check digit is usually added to"
strmuloption(9) = "serve as important tool for dat avalidation"
strmuloption(10) = " the product code in sales analysis  systemtoensure that variable information isassigned to the correct item"
strmuloption(11) = "General ledger code to ensure that correct account head isdebited/credited."
strmuloption(12) = "All of the above"

strmultiple(4) = "The pperation carried out on an input data for examining its acceptabolity is termed "
strmuloption(13) = "Validation"
strmuloption(14) = "Verification"
strmuloption(15) = "Moderation"
strmuloption(16) = "None of these"

strmultiple(5) = "A sequential file organisation is most suited for "
strmuloption(17) = "Airline reservation system"
strmuloption(18) = "A library"
strmuloption(19) = "A payroll system"
strmuloption(20) = "None of the above"

strmultiple(6) = "A suite of programs that handle an organisation's database responsibiities is called as"
strmuloption(21) = "DBMS"
strmuloption(22) = "Database Processing System(DBPS)"
strmuloption(23) = "DMS"
strmuloption(24) = "ALL of the above"

strmultiple(7) = "In preparing test data the following is not a desirable characteristic"
strmuloption(25) = "Every sttment should be executed"
strmuloption(26) = "Must be comprehensive"
strmuloption(27) = "Users do not participate"
strmuloption(28) = "None of the above"

strmultiple(8) = "Coding and testing are carried out-"
strmuloption(29) = "Top down"
strmuloption(30) = "Bottom up"
strmuloption(31) = "Ad Hoc"
strmuloption(32) = "None of these"

strmultiple(9) = "Audit can be done"
strmuloption(33) = "Through the computer"
strmuloption(34) = "Round the computer"
strmuloption(35) = "Neither(a) not (b)"
strmuloption(36) = "Both (a) and (b)"

strmultiple(10) = "Products returned bythe customers would affect the financial accounts in the following manner "
strmuloption(37) = "Debits stock & credits customer"
strmuloption(38) = "Credits stock & debits customer"
strmuloption(39) = "Credits rejected stock and debits customer"
strmuloption(40) = "None of the above"

strmulanswer(1) = "All of the above"
strmulanswer(2) = "Check manually the input data fed into the computer against input forms data"
strmulanswer(3) = "All of the above"
strmulanswer(4) = "Validation"
strmulanswer(5) = "A payroll system"
strmulanswer(6) = "ALL of the above"
strmulanswer(7) = "Users do not participate"
strmulanswer(8) = "None of these"
strmulanswer(9) = "Both (a) and (b)"
strmulanswer(10) = "Debits stock & credits customer"
End Sub
Private Sub multiplechoice_Paper3()
strmultiple(1) = "Design specifications do not normally include"
strmuloption(1) = "Blueprint showing the layout of hardware"
strmuloption(2) = "Output requirements"
strmuloption(3) = "Input and storage requirements"
strmuloption(4) = "None of the above"

strmultiple(2) = "Teh most efficient and well designed computerised payroll system would be"
strmuloption(5) = "On Line real time system"
strmuloption(6) = "Batch processing system"
strmuloption(7) = "Both a) and b)"
strmuloption(8) = "None of the above"

strmultiple(3) = "Ageing analysis is integral to which of the following?"
strmuloption(9) = "Personnel management"
strmuloption(10) = "Oil drilling"
strmuloption(11) = "Passenger reservation"
strmuloption(12) = "None of the above"

strmultiple(4) = "The term deadstock is met with in which, of the followiing systems?"
strmuloption(13) = "Cartography"
strmuloption(14) = "Personnel"
strmuloption(15) = "Both above"
strmuloption(16) = "None of these"

strmultiple(5) = "The objective in maintaining an inventory control operation is"
strmuloption(17) = "To maximise the money invested"
strmuloption(18) = "To reduce thecost of storing, handling an ensuring a large quantity o goods "
strmuloption(19) = "To have much more goods on hand thanb ordered"
strmuloption(20) = "All of the above"

strmultiple(6) = "Reported profits in a period orising prices ar depressed by the following method of inventory pricing"
strmuloption(21) = "FIFO"
strmuloption(22) = "LIFO"
strmuloption(23) = "Both of the above"
strmuloption(24) = "None of the above"

strmultiple(7) = "'A' category items in an inventory control system"
strmuloption(25) = "should  be monitored very closely"
strmuloption(26) = "should ot be reviewed"
strmuloption(27) = "should be highest stocked"
strmuloption(28) = "none of the above"

strmultiple(8) = "Issue price"
strmuloption(29) = "Stays constant throughout the system in standard costing"
strmuloption(30) = "Changes with each receipt in case of average costing"
strmuloption(31) = "Changes in the case of the LIFO"
strmuloption(32) = "all of the above"

strmultiple(9) = "Discounts and rebates are offered on"
strmuloption(33) = "the base price of product"
strmuloption(34) = "the total invoice price"
strmuloption(35) = "the Base Price Plus all Taxes"
strmuloption(36) = "all of the above"

strmultiple(10) = "In an inflationary economy, follwing will tend to undermine stock on hand"
strmuloption(37) = "Average price"
strmuloption(38) = "LIFO"
strmuloption(39) = "FIFO"
strmuloption(40) = "None of the above"

strmulanswer(1) = "Blueprint showing the layout of hardware"
strmulanswer(2) = "On Line real time system"
strmulanswer(3) = "None of the above"
strmulanswer(4) = "None of these"
strmulanswer(5) = "To reduce thecost of storing, handling an ensuring a large quantity o goods "
strmulanswer(6) = "LIFO"
strmulanswer(7) = "should  be monitored very closely"
strmulanswer(8) = "Stays constant throughout the system in standard costing"
strmulanswer(9) = "all of the above"
strmulanswer(10) = "LIFO"
End Sub
Private Sub multiplechoice_Paper4()
strmultiple(1) = "Computerised stores accounts may generate item balances differnt from actual balances because of "
strmuloption(1) = "Pilferage"
strmuloption(2) = "Auctioning stores without recording"
strmuloption(3) = "Non recording of receipts"
strmuloption(4) = "All of the above"

strmultiple(2) = "The item master contains the following fields"
strmuloption(5) = "Item code"
strmuloption(6) = "Quantity purchased"
strmuloption(7) = "Purchase rate"
strmuloption(8) = "None of the above"

strmultiple(3) = "A typicalmatrials management system ocntains following transaction files"
strmuloption(9) = "Receipts from vendors"
strmuloption(10) = "Issues to shop floors"
strmuloption(11) = " Rejected materials returned to vendors"
strmuloption(12) = "All of the above"

strmultiple(4) = "An efficient sales analysis system should be"
strmuloption(13) = "On Line processing system"
strmuloption(14) = "Batch mode processing system"
strmuloption(15) = "Both a) and b)"
strmuloption(16) = "None of the above"

strmultiple(5) = "A ales analysis system"
strmuloption(17) = "Should be integrated with financial axxounts"
strmuloption(18) = "Would prove vry beneficial if it is made on line"
strmuloption(19) = "should be integrated with finished stores system"
strmuloption(20) = "All of the above"

strmultiple(6) = "In an integrates system, regected materials returned to the vendor,would generate the following entries in financial accounts"
strmuloption(21) = "Debits creditor control account and credits raw material stock account"
strmuloption(22) = "Credits creditor control account and debits rawmaterial stock account"
strmuloption(23) = "Credits supplier account and debits raw material account "
strmuloption(24) = "None of the above"

strmultiple(7) = "An ABC analysis shows that"
strmuloption(25) = "the maximum proportional value of the inventory is made up by small number of items."
strmuloption(26) = "the maximum proportion value ofth inventory is made up by a large number of items."
strmuloption(27) = "both a) and b)"
strmuloption(28) = "neither a) or b)"

strmultiple(8) = "Value of stock can be done by"
strmuloption(29) = "LIFO"
strmuloption(30) = "Standard cost"
strmuloption(31) = "Average cost"
strmuloption(32) = "Any of the above"

strmultiple(9) = "Vendor performance can be measured in terms of"
strmuloption(33) = "Quality of supplies"
strmuloption(34) = "Timeliness of supply"
strmuloption(35) = "Materials supplied"
strmuloption(36) = "All of the above"

strmultiple(10) = "Ageing analysis ofthe stocks held in store"
strmuloption(37) = "Does ot help the company optimise its inventory"
strmuloption(38) = "Helps the company optimise its inventory"
strmuloption(39) = "Both a) and b) "
strmuloption(40) = "None of these"


strmulanswer(1) = "All of the above"
strmulanswer(2) = "None of the above"
strmulanswer(3) = "All of the above"
strmulanswer(4) = "On Line processing system"
strmulanswer(5) = "All of the above"
strmulanswer(6) = "Credits supplier account and debits raw material account "
strmulanswer(7) = "the maximum proportional value of the inventory is made up by small number of items."
strmulanswer(8) = "Any of the above"
strmulanswer(9) = "All of the above"
strmulanswer(10) = "Helps the company optimise its inventory"
End Sub
Private Sub multiplechoice_Paper5()
strmultiple(1) = "Stricter inventory controls "
strmuloption(1) = "Minimise product obsolescence"
strmuloption(2) = "maximise product obsolescence"
strmuloption(3) = "Have no effect on product obsolescence"
strmuloption(4) = "None of these"

strmultiple(2) = "Zero leadtime of supplies means"
strmuloption(5) = "We may not need to hold any inventory"
strmuloption(6) = "Reorder points will have no relevance"
strmuloption(7) = "EOQ will have no relevance"
strmuloption(8) = "All of the above"

strmultiple(3) = "Which , o the following does ot constitute a master file in a store system?"
strmuloption(9) = "Material issues file"
strmuloption(10) = " Materials recript file"
strmuloption(11) = "Purchase orders file"
strmuloption(12) = "All of these"

strmultiple(4) = "Following cab necategorised as inventory in an organisation"
strmuloption(13) = "In process matrials"
strmuloption(14) = "Dead stock"
strmuloption(15) = "Slow moving items"
strmuloption(16) = "All of the above"

strmultiple(5) = "Ageing analysis report"
strmuloption(17) = "Classigies the outstanding balances ofthe customers"
strmuloption(18) = "Is a control report for monitoring debtors"
strmuloption(19) = "Both a) and b)"
strmuloption(20) = "none of the above"

strmultiple(6) = "Which ofthe following is not a liability?"
strmuloption(21) = "Shae capital"
strmuloption(22) = "Reserves and surplus"
strmuloption(23) = "Secured loans"
strmuloption(24) = "Investments"

strmultiple(7) = "Reserves and surplus are"
strmuloption(25) = "Borrowings of the firm"
strmuloption(26) = "Profits which have neen retained in the firm"
strmuloption(27) = "Financial securities owned by the firm"
strmuloption(28) = "Amounts due to the firm"

strmultiple(8) = "A Business accounting systm includes"
strmuloption(29) = "Order entry"
strmuloption(30) = "Sales analysis"
strmuloption(31) = "Shipping/billing"
strmuloption(32) = "All of the above"

strmultiple(9) = "The accounts payable(A/P) component in an accounting package is not designed to"
strmuloption(33) = "Verify the accuracy of incoming bills"
strmuloption(34) = "Keep track of amounts owed by customers "
strmuloption(35) = "Make sure that discounts are not lost"
strmuloption(36) = "Issue payment cheques towards the end of the discount period"

strmultiple(10) = "An accounts receivable operation uses"
strmuloption(37) = "Customer invoices"
strmuloption(38) = "Purchases vouchers"
strmuloption(39) = "Challans"
strmuloption(40) = "All of the above"

strmulanswer(1) = "Have no effect on product obsolescence"
strmulanswer(2) = "All of the above"
strmulanswer(3) = "All of these"
strmulanswer(4) = "All of the above"
strmulanswer(5) = "Both a) and b)"
strmulanswer(6) = "Investments"
strmulanswer(7) = "Profits which have neen retained in the firm"
strmulanswer(8) = "All of the above"
strmulanswer(9) = "Keep track of amounts owed by customers "
strmulanswer(10) = "Customer invoices"
End Sub
Private Sub multiplechoice_Paper6()
strmultiple(1) = "An Accounting system's collectioin ofall temporaty and permanent accounts ismaintained in the"
strmuloption(1) = "Cash ournal"
strmuloption(2) = "General ledger"
strmuloption(3) = "Accounts payable component"
strmuloption(4) = "Double entry balance system"

strmultiple(2) = "High receivables as a proportion of sales, indicates "
strmuloption(5) = "Efficient credit control policy"
strmuloption(6) = "Poor credit control policy"
strmuloption(7) = "Large number ofcustomers"
strmuloption(8) = "none of these"

strmultiple(3) = "The credit control sytem should take account of"
strmuloption(9) = "Customer data bank"
strmuloption(10) = "Past payment record"
strmuloption(11) = "Volume of purchases"
strmuloption(12) = "All of the above"

strmultiple(4) = "Contra entries afect the "
strmuloption(13) = "P & L"
strmuloption(14) = "Balance sheet"
strmuloption(15) = "Capital account"
strmuloption(16) = "None of these"

strmultiple(5) = "Every year end, the computerised financial axxounts"
strmuloption(17) = "Zeroise all income heads"
strmuloption(18) = "Zeroise all expenditure heads"
strmuloption(19) = "Leaves sub ledger accounts unchanged"
strmuloption(20) = "Does all of the above"

strmultiple(6) = "Payroll system should necessarily have provision for"
strmuloption(21) = "Ensuring that duplicate pay slips are not generated for an employee in a given month"
strmuloption(22) = "Arrears ar paid to the employees in a timely manner"
strmuloption(23) = "Integrating the employee PF contribution with the PF system"
strmuloption(24) = "All of the above"

strmultiple(7) = "The master file of apayroll system contains"
strmuloption(25) = "Employee name"
strmuloption(26) = "No. of days present"
strmuloption(27) = "Basic salary"
strmuloption(28) = "(a) and (c) only"

strmultiple(8) = "The most efficient and well designed computerised payroll system would be"
strmuloption(29) = "On Line real time"
strmuloption(30) = "Batch processing"
strmuloption(31) = "Both a) and b)"
strmuloption(32) = "None of the above"

strmultiple(9) = "Computer based provident func accounting system"
strmuloption(33) = "Uses inputs from the monthly pay slips"
strmuloption(34) = "Is an yearly system"
strmuloption(35) = "Does not accept anytpe of input(transaction) directly,"
strmuloption(36) = "All of the above"

strmultiple(10) = "The bank bool generated bythe computerised financial accounts is"
strmuloption(37) = "A mirror image of the bank statement prepared by the bank"
strmuloption(38) = "To be reconciled with the statementreceived from the bank"
strmuloption(39) = "Both a) and b)"
strmuloption(40) = "None of the above"

strmulanswer(1) = "General ledger"
strmulanswer(2) = "Poor credit control policy"
strmulanswer(3) = "All of the above"
strmulanswer(4) = "None of these"
strmulanswer(5) = "Does all of the above"
strmulanswer(6) = "All of the above"
strmulanswer(7) = "(a) and (c) only"
strmulanswer(8) = "On Line real time"
strmulanswer(9) = "All of the above"
strmulanswer(10) = "Both a) and b)"
End Sub
Private Sub multiplechoice_Paper7()
strmultiple(1) = "Employee advances ar generally"
strmuloption(1) = "Kept in the subsidiary account"
strmuloption(2) = "Are part ofthe memorandum accounts"
strmuloption(3) = "Are part of the memorandum accounts"
strmuloption(4) = "None of the above"

strmultiple(2) = "Payroll  isintegrated with financial accounts through"
strmuloption(5) = "A month end hournal voucher"
strmuloption(6) = "On line updation"
strmuloption(7) = "Maintaining each employees record in a subsidiary ledger"
strmuloption(8) = "All of the above"

strmultiple(3) = "Which ofeh following is not a liability?"
strmuloption(9) = "Share Capital"
strmuloption(10) = " Reserves and surplus"
strmuloption(11) = "Secured loans"
strmuloption(12) = "None of the above"

strmultiple(4) = "A credit note is"
strmuloption(13) = "Raised n the customer"
strmuloption(14) = "Cedits customer & debits sales"
strmuloption(15) = "Both a) and b)"
strmuloption(16) = "None of the above"

strmultiple(5) = "In adouble entry book keeping system, debits and credits"
strmuloption(17) = "Need not match"
strmuloption(18) = "Must match"
strmuloption(19) = "Both a) and b)"
strmuloption(20) = "Neither a) nor b)"

strmultiple(6) = "Reported profits in a period of rising prices are depressed by the following method of Inventory Pricing"
strmuloption(21) = "FIFO"
strmuloption(22) = "LIFO"
strmuloption(23) = "Both a) and b)"
strmuloption(24) = "None of the above"

strmultiple(7) = "Depreciation is"
strmuloption(25) = "Considered an expense item"
strmuloption(26) = "An addition to the assets of teh company"
strmuloption(27) = "Provisional to inflation, the profits ofthe company"
strmuloption(28) = "None ofthe above"

strmultiple(8) = "A suspense account represents"
strmuloption(29) = "Balances which do not have clear cut account heads"
strmuloption(30) = "Balances whose treatment isstill to be decided"
strmuloption(31) = "Neither a) nor b)"
strmuloption(32) = "Both a) and b)"

strmultiple(9) = "TDS is met with in which of the following systems?"
strmuloption(33) = "Inventory control"
strmuloption(34) = "P P & C"
strmuloption(35) = "Sales analysis"
strmuloption(36) = "None of these"

strmultiple(10) = "Overtime analysis is"
strmuloption(37) = "USeful for formulating personnel policies"
strmuloption(38) = "Is derived from the payroll system"
strmuloption(39) = "Both of the above"
strmuloption(40) = "Neither of the above"

strmulanswer(1) = "Kept in the subsidiary account"
strmulanswer(2) = "A month end hournal voucher"
strmulanswer(3) = "None of the above"
strmulanswer(4) = "Both a) and b)"
strmulanswer(5) = "Must match"
strmulanswer(6) = "LIFO"
strmulanswer(7) = "Considered an expense item"
strmulanswer(8) = "Both a) and b)"
strmulanswer(9) = "None of these"
strmulanswer(10) = "Both of the above"
End Sub
Private Sub multiplechoice_Paper8()
strmultiple(1) = "Which of the following does NOT form part of the subsidiary ledger?"
strmuloption(1) = "Loans to employees"
strmuloption(2) = "Sales to customers"
strmuloption(3) = "Purchases from suppliers"
strmuloption(4) = "None of the above"

strmultiple(2) = "Which ofthe follwoing is a valid command?"
strmuloption(5) = "INPUT 'Enter your age' AGE"
strmuloption(6) = "INPUT 'Enter your age' ,AGE"
strmuloption(7) = "INPUT 'Enter your age' GET AGE"
strmuloption(8) = "INPUT 'Enter your age' TO AGE"

strmultiple(3) = "Which ofthe following commands can be used to delete all files with extensions name '.BAK' in thecurrent directory while running a DBMS software?"
strmuloption(9) = "RUN DEL *.BAK"
strmuloption(10) = "RUN ERASE *.BAK"
strmuloption(11) = "!DEL *.BAK"
strmuloption(12) = "All of the above"

strmultiple(4) = "If AGE and SALARY are numeric fields in a database file, which one of the following commands is incorrect?"
strmuloption(13) = "REPLACE ALL AGE WITH AGE +1 FOR AGE < 20"
strmuloption(14) = "REPLACE ALL AGE WITH AGE +1 WHILE AGE < 20"
strmuloption(15) = "REPLACE AGE WITH AGE +1 FOR AGE < 20 ALL"
strmuloption(16) = "All of the above"

strmultiple(5) = "Which one ofthe following is invalid?"
strmuloption(17) = "AVERAGE calculates average of a specified memory variable."
strmuloption(18) = "INPUT cab create character , numeric orlogical variable."
strmuloption(19) = "A user defined message can be spedified with WAIT."
strmuloption(20) = "None of the above"

strmultiple(6) = "Which of the following statements isfalse?"
strmuloption(21) = "FIND and SEEK commands need an index file."
strmuloption(22) = "CONTINUE cannot be used to find the next matching record after using FIND/SEEK."
strmuloption(23) = "Contents of a variable can be searched with FIND and SEEK commands"
strmuloption(24) = "None of the above"

strmultiple(7) = "The command that selects an index file/tag for a database file is"
strmuloption(25) = "SELECT INDEX"
strmuloption(26) = "SELECT ORDER"
strmuloption(27) = "SET ORDER"
strmuloption(28) = "None of the above"

strmultiple(8) = "The command to create a new dtabase file based on two open database files is"
strmuloption(29) = "SET RELATION"
strmuloption(30) = "UPDATE"
strmuloption(31) = "JOIN"
strmuloption(32) = "none of the above"

strmultiple(9) = "The command used to change the contents of a database file using contents of another file by linking them on a common key field is"
strmuloption(33) = "REPLACE"
strmuloption(34) = "UPDATE"
strmuloption(35) = "CHANGE"
strmuloption(36) = "JOIN"

strmultiple(10) = "The most appropriate commandto permanently remove all records from the currnt database file is"
strmuloption(37) = "PACK"
strmuloption(38) = "DELETE ALL"
strmuloption(39) = "ZAP"
strmuloption(40) = "DELETE FILE"

strmulanswer(1) = "None of the above"
strmulanswer(2) = "INPUT 'Enter your age' TO AGE"
strmulanswer(3) = "All of the above"
strmulanswer(4) = "All of the above"
strmulanswer(5) = "AVERAGE calculates average of a specified memory variable."
strmulanswer(6) = "CONTINUE cannot be used to find the next matching record after using FIND/SEEK."
strmulanswer(7) = "SET ORDER"
strmulanswer(8) = "JOIN"
strmulanswer(9) = "UPDATE"
strmulanswer(10) = "ZAP"
End Sub
Private Sub multiplechoice_Paper9()
strmultiple(1) = "Which one ofthe following is an invalid command in a DBMS package?"
strmuloption(1) = "RUN"
strmuloption(2) = "CLEAR"
strmuloption(3) = "REPLACE"
strmuloption(4) = "None of the above"

strmultiple(2) = "To convert a numeric to a character string, use the function "
strmuloption(5) = "ASC()"
strmuloption(6) = "CHR()"
strmuloption(7) = "STR()"
strmuloption(8) = "VAL()"

strmultiple(3) = " The function to remove leading andtrailing spaces from a character expression is"
strmuloption(9) = " TRIM()"
strmuloption(10) = "LTRIM()"
strmuloption(11) = "RTRIM()"
strmuloption(12) = "ALLTRIM()"

strmultiple(4) = "If NAME is a character field and AGE is a numeric field, which one of the following commands is invalid?"
strmuloption(13) = "REPLACE NAME WITH 'Rama' "
strmuloption(14) = "REPLACE NAME WITH [RAMA]"
strmuloption(15) = "REPLACE NAME WITH [RAMA], AGE WITH 150"
strmuloption(16) = "None of the above"

strmultiple(5) = "GO BOTTOM and SKIP -3 commands ae given in a sequence after a database file having 40 records have been opened. Then, the record pointer points to"
strmuloption(17) = "38th record"
strmuloption(18) = "37th record"
strmuloption(19) = "3rd record"
strmuloption(20) = "none of the above"

strmultiple(6) = "If SENT is  a logical field in a database file, which one ofthe following is a valid command?"
strmuloption(21) = "DISPLAY FOR SENT = 'T'"
strmuloption(22) = "DISPLAY FOR SENT = '.T.'"
strmuloption(23) = "DISPLAY FOR SENT"
strmuloption(24) = "DISPLAY FOR SENT =T"

strmultiple(7) = "Which one of the following is  an invalid command?"
strmuloption(25) = "CLOSE DATABASE"
strmuloption(26) = "SET ALTERNATE ON"
strmuloption(27) = "SET ALTERNATE TO DUMMY"
strmuloption(28) = "ERASE ALTERNATE"

strmultiple(8) = "Which one ofthe following is an invalid statement?"
strmuloption(29) = "A procedure can be entered in a procedure file or in the command File itself."
strmuloption(30) = "Each pricedure begins with PROCEDURE >procedure name> command."
strmuloption(31) = "A procddure normally ends with RETURN but we cab skip the RETURN at the end of the procedure."
strmuloption(32) = "None of the above"

strmultiple(9) = "Which ofthe following commands is incorrect for displaying a record with a value 2000 in a numeric PHONE field and 'false' in a logical field PAID?"
strmuloption(33) = "DISPLAY FOR PHONE = 2000 AND PAID = 'FALSE'"
strmuloption(34) = "DISPLAY FOR PHONE = 2000 AND PAID = "
strmuloption(35) = "DISPLAY FOR PHONE = 2000 .AND. .NOT. .PAID."
strmuloption(36) = " DISPLAY FOR PHONE = 2000 AND PAID = NOT"

strmultiple(10) = "The cmmand used to add a ew field to an active database file is"
strmuloption(37) = "MODIFY LABEL"
strmuloption(38) = "MODIFY COMMAND"
strmuloption(39) = "MODIFY STRUCTURE"
strmuloption(40) = "MODIFY REPORT"

strmulanswer(1) = "None of the above"
strmulanswer(2) = "STR()"
strmulanswer(3) = "ALLTRIM()"
strmulanswer(4) = "None of the above"
strmulanswer(5) = "37th record"
strmulanswer(6) = "DISPLAY FOR SENT"
strmulanswer(7) = "ERASE ALTERNATE"
strmulanswer(8) = "None of the above"
strmulanswer(9) = "DISPLAY FOR PHONE = 2000 .AND. .NOT. .PAID."
strmulanswer(10) = "MODIFY STRUCTURE"
End Sub
Private Sub multiplechoice_Paper10()
strmultiple(1) = "Which command is used for closing the textfilethat has been created with SET ALTERNATE TO <file name> command?"
strmuloption(1) = "SET ALTERNATE OFF"
strmuloption(2) = "CLOSE DATABASES"
strmuloption(3) = "CLOSE ALTERNATE"
strmuloption(4) = "CLEAR ALL"

strmultiple(2) = "The MODIFY command is used to"
strmuloption(5) = "Modify the structure of a database file"
strmuloption(6) = "Modify a report format file"
strmuloption(7) = "Add a new field to every record of a database file"
strmuloption(8) = "none of the above"

strmultiple(3) = "The command to find out the numbr ofrecords that meet a specified condition is"
strmuloption(9) = "CONDITION"
strmuloption(10) = "UPDATE"
strmuloption(11) = "JOIN"
strmuloption(12) = "None of the above"

strmultiple(4) = "The purpose ofthe EXIT command in FoxPro"
strmuloption(13) = "Get out of a condition loop"
strmuloption(14) = "Get out o FoxPro"
strmuloption(15) = "Transfer the control to a diffrent procedure"
strmuloption(16) = "None of the above"

strmultiple(5) = "The SET ALTERNATE TO command"
strmuloption(17) = "Selects an alternalte program file"
strmuloption(18) = "Defines another name for a database file"
strmuloption(19) = "Selects an alternate procedure file"
strmuloption(20) = "None of the above"

strmultiple(6) = "Mark the invalid statement"
strmuloption(21) = "Index files are normally smaller than sorted files"
strmuloption(22) = "Structural compound indes files are automatically opened and updated."
strmuloption(23) = "Up to seven index files can be used with a database file"
strmuloption(24) = "None of the above"

strmultiple(7) = "which one of the followig statements is not correct?"
strmuloption(25) = "The STORE command can create more than one vriable"
strmuloption(26) = "DATE() returns the current date."
strmuloption(27) = "YEAR() returns the numeric value of year from a date expression."
strmuloption(28) = "FIELD() returns the type of a specified field"

strmultiple(8) = "To unconditionally exitfrom a SCAN-ENDSCAN loop, one can use:"
strmuloption(29) = "LOOp"
strmuloption(30) = ".NOT. EOF()"
strmuloption(31) = "EXIT"
strmuloption(32) = "SKIP"

strmultiple(9) = "To establish a link between two database files, one can use the Command:"
strmuloption(33) = "JOIN"
strmuloption(34) = "UPDATE"
strmuloption(35) = "SET RELATION"
strmuloption(36) = "BROWSE"

strmultiple(10) = "Which one of the following statements is incorrect?"
strmuloption(37) = "Structural compound index files are automatically opened."
strmuloption(38) = "Structural compund indexfiles are automatically opened and updated."
strmuloption(39) = "All open inde files ar automatically updated."
strmuloption(40) = "None of the above"

strmulanswer(1) = "CLOSE ALTERNATE"
strmulanswer(2) = "Add a new field to every record of a database file"
strmulanswer(3) = "None of the above"
strmulanswer(4) = "Get out of a condition loop"
strmulanswer(5) = "None of the above"
strmulanswer(6) = "None of the above"
strmulanswer(7) = "FIELD() returns the type of a specified field"
strmulanswer(8) = "EXIT"
strmulanswer(9) = "SET RELATION"
strmulanswer(10) = "None of the above"
End Sub

