VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCPapers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Dim StrQuestions(10) As String * 5000
 Dim stroptions(10) As String
 Dim stranswers(10) As String
 Dim strcolumnA(10) As String
 Dim strcolumnB(10) As String

'Variable declaration for the multiple choice
 Dim strmultiple(10) As String
 Dim strmuloption(40) As String
 Dim strmulanswer(10) As String

'Enum Category
' Fill_Questions = 0
' Fill_Options = 1
' Fill_Answers = 2
' Matching_ColumnA = 3
' Matching_ColumnB = 4
' Matching_Answers = 5
' Questions = 6
' TF_Questions = 7
' TF_Answers = 8
' MC_Questions = 9
' MC_Options_1 = 10
' MC_Options_2 = 11
' MC_Options_3 = 12
' MC_Options_4 = 13
' MC_Answer = 14
' End Enum

'Enum Paper
' Paper1 = 1
' Paper2 = 2
' Paper3 = 3
' Paper4 = 4
' Paper5 = 5
' Paper6 = 6
' Paper7 = 7
' Paper8 = 8
' Paper9 = 9
' Paper10 = 10
' Paper11 = 11
' Paper12 = 12
' Paper13 = 13
' Paper14 = 14
' Paper15 = 15
'End Enum

Function PC(PaperType As String, PCPaper As Paper, mCat As Category, Number As Integer)
'ptype is the paper type
'PCPaper is the number of paper as paper1,paper2 ...
'mcat is the category
'number is the number of questions, or options or answers to be taken

 StrPaperType = PaperType
' StrPaper = Paper
If PaperType = "Fill in the Blanks" Then
   If PCPaper = Paper1 Then Fill_in_the_Blanks_Paper1
   If PCPaper = Paper2 Then Fill_in_the_Blanks_Paper2
   If PCPaper = Paper3 Then Fill_in_the_Blanks_Paper3
   If PCPaper = Paper4 Then Fill_in_the_Blanks_Paper4
   If PCPaper = Paper5 Then Fill_in_the_Blanks_Paper5
   If PCPaper = Paper6 Then Fill_in_the_Blanks_Paper6
   If PCPaper = Paper7 Then Fill_in_the_Blanks_Paper7
   If PCPaper = Paper8 Then Fill_in_the_Blanks_Paper8
   If PCPaper = Paper9 Then Fill_in_the_Blanks_Paper9
   If PCPaper = Paper10 Then Fill_in_the_Blanks_Paper10
   If PCPaper = Paper11 Then Fill_in_the_Blanks_Paper11
   If PCPaper = Paper12 Then Fill_in_the_Blanks_Paper12

ElseIf PaperType = "Matching Columns" Then
       If PCPaper = Paper1 Then Matching_Columns_Paper1
       If PCPaper = Paper2 Then Matching_Columns_Paper2
       If PCPaper = Paper3 Then Matching_Columns_Paper3
       If PCPaper = Paper4 Then Matching_Columns_Paper4
       If PCPaper = Paper5 Then Matching_Columns_Paper5
       If PCPaper = Paper6 Then Matching_Columns_Paper6
       If PCPaper = Paper7 Then Matching_Columns_Paper7
       If PCPaper = Paper8 Then Matching_Columns_Paper8
       If PCPaper = Paper9 Then Matching_Columns_Paper9
       If PCPaper = Paper10 Then Matching_Columns_Paper10
       If PCPaper = Paper11 Then Matching_Columns_Paper11
       If PCPaper = Paper12 Then Matching_Columns_Paper12
       
ElseIf PaperType = "True Or False" Then
       If PCPaper = Paper1 Then TrueorFalse_Paper1
       If PCPaper = Paper2 Then TrueorFalse_Paper2
       If PCPaper = Paper3 Then TrueorFalse_Paper3
       If PCPaper = Paper4 Then TrueorFalse_Paper4
       If PCPaper = Paper5 Then TrueorFalse_Paper5
       If PCPaper = Paper6 Then TrueorFalse_Paper6
       If PCPaper = Paper7 Then TrueorFalse_Paper7
       If PCPaper = Paper8 Then TrueorFalse_Paper8
       If PCPaper = Paper9 Then TrueorFalse_Paper9
       If PCPaper = Paper10 Then TrueorFalse_Paper10
       If PCPaper = Paper11 Then TrueorFalse_Paper11
       If PCPaper = Paper12 Then TrueorFalse_Paper12
       If PCPaper = Paper13 Then TrueorFalse_Paper13
       If PCPaper = Paper14 Then TrueorFalse_Paper14
       If PCPaper = Paper15 Then TrueorFalse_Paper15
       If PCPaper = Paper16 Then TrueorFalse_Paper16
       If PCPaper = Paper17 Then TrueorFalse_Paper17
       If PCPaper = Paper18 Then TrueorFalse_Paper18
       If PCPaper = Paper19 Then TrueorFalse_Paper19

ElseIf PaperType = "Descriptive Questions" Then
       If PCPaper = Paper1 Then questions_paper1
       If PCPaper = Paper2 Then questions_paper2
       If PCPaper = Paper3 Then questions_paper3
       If PCPaper = Paper4 Then questions_paper4
       If PCPaper = Paper5 Then questions_paper5
       If PCPaper = Paper6 Then questions_paper6
       If PCPaper = Paper7 Then questions_paper7
       If PCPaper = Paper8 Then questions_paper8
       If PCPaper = Paper9 Then questions_paper9
       If PCPaper = Paper10 Then questions_paper10
       If PCPaper = Paper11 Then questions_paper11
       If PCPaper = Paper12 Then questions_paper12
       If PCPaper = Paper13 Then questions_paper13
       If PCPaper = Paper14 Then questions_paper14
       If PCPaper = Paper15 Then questions_paper15
       
ElseIf PaperType = "Multiple Choice" Then
       If PCPaper = Paper1 Then multiplechoice_Paper1
       If PCPaper = Paper2 Then multiplechoice_Paper2
       If PCPaper = Paper3 Then multiplechoice_Paper3
       If PCPaper = Paper4 Then multiplechoice_Paper4
       If PCPaper = Paper5 Then multiplechoice_Paper5
       If PCPaper = Paper6 Then multiplechoice_Paper6
       If PCPaper = Paper7 Then multiplechoice_Paper7
       If PCPaper = Paper8 Then multiplechoice_Paper8
       If PCPaper = Paper9 Then multiplechoice_Paper9
       If PCPaper = Paper10 Then multiplechoice_Paper10
       If PCPaper = Paper11 Then multiplechoice_Paper11
       If PCPaper = Paper12 Then multiplechoice_Paper12
              
End If

If mCat = Fill_Questions Then
   PC = StrQuestions(Number)
ElseIf mCat = Fill_Options Then
       PC = stroptions(Number)
ElseIf mCat = Fill_Answers Then
       PC = stranswers(Number)

ElseIf mCat = Matching_ColumnA Then
       PC = strcolumnA(Number)
ElseIf mCat = Matching_ColumnB Then
       PC = strcolumnB(Number)
ElseIf mCat = Matching_Answers Then
       PC = stranswers(Number)
       
ElseIf mCat = Questions Then
       PC = StrQuestions(Number)
       
ElseIf mCat = TF_Questions Then
       PC = StrQuestions(Number)
ElseIf mCat = TF_Answers Then
       PC = stranswers(Number)
       
ElseIf mCat = MC_Questions Then
       PC = strmultiple(Number)
       
ElseIf mCat = MC_Options_1 Then
       PC = strmuloption(Number)
       
ElseIf mCat = MC_Options_2 Then
       PC = strmuloption(Number)
       
ElseIf mCat = MC_Options_3 Then
       PC = strmuloption(Number)
       
ElseIf mCat = MC_Options_4 Then
       PC = strmuloption(Number)
       
ElseIf mCat = MC_Answer Then
       PC = strmulanswer(Number)
End If

End Function

Private Sub Fill_in_the_Blanks_Paper1()
StrQuestions(1) = "The dash command transfers files from one directory to another"
StrQuestions(2) = "To delete all files in a directory, including subdirectories, the dash command can be used."
StrQuestions(3) = "The dash command can be used to increase the storage space in disk."
StrQuestions(4) = "An used diskette has been formatted with FORMAT. The dash command can be used to increase the storage space in a disk."
StrQuestions(5) = "The name of the text Editor included in DOS is dash."
StrQuestions(6) = "The dash command can be used to correct the file system."
StrQuestions(7) = "To display the help information about the COPY command, we can use dash."
StrQuestions(8) = "dash sends the directory listing to printer."
StrQuestions(9) = "The dash command recovers readable information from bad or defective disks."
StrQuestions(10) = "The dash command is used to change the volume label of a diskette."

stroptions(1) = "RECOVER"
stroptions(2) = "HELP COPY"
stroptions(3) = "LABEL"
stroptions(4) = "EDIT"
stroptions(5) = "DIR > PRN"
stroptions(6) = "UNFORMAT"
stroptions(7) = "MOVE"
stroptions(8) = "SCANDISK"
stroptions(9) = "DBLSPACE"
stroptions(10) = "DELTREE"

stranswers(1) = "MOVE"
stranswers(2) = "DELTREE"
stranswers(3) = "DBLSPACE"
stranswers(4) = "UNFORMAT"
stranswers(5) = "EDIT"
stranswers(6) = "SCANDISK"
stranswers(7) = "HELP COPY"
stranswers(8) = "DIR > PRN"
stranswers(9) = "UNFORMAT"
stranswers(10) = "LABEL"

End Sub

Private Sub Fill_in_the_Blanks_Paper2()
StrQuestions(1) = "dash is used to check disk surface for errors."
StrQuestions(2) = "For placing the prompt in the first line on screen the DOS command is dash."
StrQuestions(3) = "The command to select a new directory is dash"
StrQuestions(4) = "The command to delete all files having extension name " & Chr(26) & ".DOC" & Chr(26) & " from the current directory is dash"
StrQuestions(5) = "COPY C:\SAMPLE \NAME.DOC A:\NEW copies files from the dash sample directory in drive C to diskette in drive A."
StrQuestions(6) = "The dash command is better than COPY C:\NEW\*.* A:."
StrQuestions(7) = "The storage capacity of a double sided high density 5.25 diskette is about dash"
StrQuestions(8) = "The storage capacity of a double sided high density 3.5 diskette is about dash."
StrQuestions(9) = "The current drive is C. The command to delete all files having the extension name " & Chr(26) & ".PRG" & Chr(26) & " from the root directory of drive A is dash"
StrQuestions(10) = "The command to create a directory in a floppy diskette is dash"

stroptions(1) = "1.44 MB"
stroptions(2) = "DEL *.DOC/ERASE *.DOC"
stroptions(3) = "XCOPY C:\NEW\*.* A:"
stroptions(4) = "MD/ MKDIR"
stroptions(5) = "DEL A:\*.PRG or ERASE A:\*.PRG"
stroptions(6) = "1.2 MB"
stroptions(7) = "SCANDISK"
stroptions(8) = "SAMPLE"
stroptions(9) = "CLS"
stroptions(10) = "CD/CHDIR"

stranswers(1) = "SCANDISK"
stranswers(2) = "CLS"
stranswers(3) = "CD/CHDIR"
stranswers(4) = "DEL A:\*.PRG or ERASE A:\*.PRG"
stranswers(5) = "SAMPLE"
stranswers(6) = "XCOPY C:\NEW\*.* A:"
stranswers(7) = "1.2 MB"
stranswers(8) = "1.44 MB"
stranswers(9) = "DEL A:\*.PRG or ERASE A:\*.PRG"
stranswers(10) = "MD/ MKDIR"

End Sub

Private Sub Fill_in_the_Blanks_Paper3()
StrQuestions(1) = "To change the prompt character(s) on the screen, we can use the dash command."
StrQuestions(2) = "To get back an erased file, the dash command is used."
StrQuestions(3) = "A DSHD Floppy Diskette has maximum track density of dash TPI."
StrQuestions(4) = "The concentric circles in a floppy are called dash."
StrQuestions(5) = "Bootstrap loader is stored in the dash of the floppy."
StrQuestions(6) = "Partition table is found in dash."
StrQuestions(7) = "By pressing the dash one can execute the commands in the CONFIG.SYS file step by step."
StrQuestions(8) = "By pressing the dash key, one can bypass the execution of CONFIG.SYS and AUTOEXEC.BAT during booting."
StrQuestions(9) = "The name of the text editor included in DOS is dash. "
StrQuestions(10) = "dash is used to check a disk surface for errors."

stroptions(1) = "First Sector"
stroptions(2) = "F8 key"
stroptions(3) = "F5"
stroptions(4) = "PROMPT"
stroptions(5) = "EDIT"
stroptions(6) = "Hard Disks"
stroptions(7) = "UNDELETE"
stroptions(8) = "96"
stroptions(9) = "Tracks"
stroptions(10) = "SCANDISK"

stranswers(1) = "PROMPT"
stranswers(2) = "UNDELETE"
stranswers(3) = "96"
stranswers(4) = "Tracks"
stranswers(5) = "First Sector"
stranswers(6) = "Hard Disks"
stranswers(7) = "F8 key"
stranswers(8) = "F5"
stranswers(9) = "EDIT"
stranswers(10) = "SCANDISK"

End Sub

Private Sub Fill_in_the_Blanks_Paper4()
StrQuestions(1) = "The commands to open, close and print a worksheet are available under the dash menu."
StrQuestions(2) = "In Excel, four types of documents can be created. These are: worksheets, charts, macro sheets and dash"
StrQuestions(3) = "To delete a worksheet document select dash from the file menu."
StrQuestions(4) = "A worksheet can now protected with a dash"
StrQuestions(5) = "To select a cell, dash mouse on the desired cell."
StrQuestions(6) = "To select an entire row, click mouse on the dash"
StrQuestions(7) = "Formula bar includes an enter box and a/a dash box."
StrQuestions(8) = "If Excel cannot completely display the result of a numeric expression, it automatically converts the results to the dash format."
StrQuestions(9) = "If you enter 12 + 24 in a cell, Excel will display dash in the cell."
StrQuestions(10) = "The address of the last cell in a worksheet is dash."

stroptions(1) = "12+24"
stroptions(2) = "Delete"
stroptions(3) = "IV16384"
stroptions(4) = "Password"
stroptions(5) = "Scientific"
stroptions(6) = "File"
stroptions(7) = "Workbooks"
stroptions(8) = "Cancel"
stroptions(9) = "Click"
stroptions(10) = "Row heading"

stranswers(1) = "File"
stranswers(2) = "Workbooks"
stranswers(3) = "Delete"
stranswers(4) = "Password"
stranswers(5) = "click"
stranswers(6) = "Row Heading"
stranswers(7) = "Cancel"
stranswers(8) = "Scientific"
stranswers(9) = "12+24"
stranswers(10) = "IV16384"

End Sub

Private Sub Fill_in_the_Blanks_Paper5()
StrQuestions(1) = "A worksheet contains dash columns."
StrQuestions(2) = "If the result of a formula cannot be displayed completely in a cell, it displays dash"
StrQuestions(3) = "To stop a macro that is running, press dash key."
StrQuestions(4) = "To create a button for a macro, the dash toolbar is used."
StrQuestions(5) = "Besides creating a macro by recording, it is also possible to create it by dash"
StrQuestions(6) = "To run a macro that runs automatically when you open worksheet its name should start with dash"
StrQuestions(7) = "The dash function returns the net present value of an investment based on a series of cash flows and a discount rate."
StrQuestions(8) = "To display the current date, use the dash function."
StrQuestions(9) = "The dash function returns the periodic payment for an annuity."
StrQuestions(10) = "The dash function returns the day of the month from a date serial number."

stroptions(1) = "NPV ()"
stroptions(2) = "DAY ()"
stroptions(3) = "Writing"
stroptions(4) = "PMT ()"
stroptions(5) = "NOW () or TODAY ()"
stroptions(6) = "Customize"
stroptions(7) = "######"
stroptions(8) = "Esc"
stroptions(9) = "Auto_Open"
stroptions(10) = "256"

stranswers(1) = "256"
stranswers(2) = "######"
stranswers(3) = "Esc"
stranswers(4) = "Customize"
stranswers(5) = "Writing"
stranswers(6) = "Auto-Open"
stranswers(7) = "NPV()"
stranswers(8) = "NOW() or TODAY()"
stranswers(9) = "PMT()"
stranswers(10) = "DAY()"

End Sub

Private Sub Fill_in_the_Blanks_Paper6()
StrQuestions(1) = "dash enables you to control a wide range of parameters that determines he way the WINDOWS environment appears."
StrQuestions(2) = "The CPU must be at least dash to run WINDOWS in enhanced mode."
StrQuestions(3) = "dash file contains the ssetting that control WINDOWS at an operating level."
StrQuestions(4) = "dash files define the system resources for WINDOWS"
StrQuestions(5) = "The dash indicates the active document in WINDOWS"
StrQuestions(6) = "Notepad is a dash"
StrQuestions(7) = "PIF stands for dash"
StrQuestions(8) = "dash ensures one application uses the services of another."
StrQuestions(9) = "An applicatiion receiving data is scalled dash"
StrQuestions(10) = "The application that sources the data is called dash."

stroptions(1) = "OLE"
stroptions(2) = "Title Bar"
stroptions(3) = "Server applcation"
stroptions(4) = "SYSTEM.Ini"
stroptions(5) = "Client application"
stroptions(6) = "Program Information File"
stroptions(7) = "Control Panel"
stroptions(8) = "Text editor"
stroptions(9) = "80386 & Above"
stroptions(10) = "WIN.INI"

stranswers(1) = "Control Panel"
stranswers(2) = "80286 & Above"
stranswers(3) = "Win.ini"
stranswers(4) = "System.ini"
stranswers(5) = "Title Bar"
stranswers(6) = "Text Editor"
stranswers(7) = "Program Information File"
stranswers(8) = "OLE"
stranswers(9) = "Client Application"
stranswers(10) = "Server Application"

End Sub
Private Sub Fill_in_the_Blanks_Paper7()
StrQuestions(1) = "To select a single word in a document, dash the mouse on the word."
StrQuestions(2) = "Before you can apply a font on a particluar text, you need to dash that text."
StrQuestions(3) = "Character size is measured in dash"
StrQuestions(4) = "The 12 point text is dash than the 10 point text."
StrQuestions(5) = "Times Roman is the name of a dash"
StrQuestions(6) = "When WORD automatically starts a new page, this page break iscalled the dash page break."
StrQuestions(7) = "When a user manually inserts a page break at a particluar position in the socument, this page break is called the dash page break."
StrQuestions(8) = "To insert a page break in adocument, use the dash option from the insert menu."
StrQuestions(9) = "dash reduces the rugged appearance of unjustified text."
StrQuestions(10) = "The characer that fills the space up to the next tab stop position is called the dash character."

stroptions(1) = "Break"
stroptions(2) = "Leader"
stroptions(3) = "Double click"
stroptions(4) = "Font"
stroptions(5) = "Hyphenation"
stroptions(6) = "Larger"
stroptions(7) = "Point"
stroptions(8) = "Hard"
stroptions(9) = "Select"
stroptions(10) = ""

stranswers(1) = "Double Click"
stranswers(2) = "Select"
stranswers(3) = "Point"
stranswers(4) = "Larger"
stranswers(5) = "Font"
stranswers(6) = "Soft"
stranswers(7) = "Hard"
stranswers(8) = "Break"
stranswers(9) = "Hyphenation"
stranswers(10) = "Leader"

End Sub

Private Sub Fill_in_the_Blanks_Paper8()
StrQuestions(1) = "Mail Merge involves merging in of the dash with a data source."
StrQuestions(2) = "A table can have dash number of rows."
StrQuestions(3) = "dash cuts the selected text and graphics and puts it is clipboard"
StrQuestions(4) = "dash copies the selected text and grapics into clipboard."
StrQuestions(5) = "dash pastes the clipboard contents into a document."
StrQuestions(6) = "To cancel the last editing, we can use the dash button."
StrQuestions(7) = "The dash mark stores all formatting for the text and graphics in the paragraph."
StrQuestions(8) = "To copy the selected text with mouse, hold down the dash key, point to the selected text and drag the insertion point to the new location."
StrQuestions(9) = "To select a complete line, click in the dash bar to the left of the line."
StrQuestions(10) = "To zoom the document use the Zoom option from the dash menu."

stroptions(1) = "Paragraph"
stroptions(2) = "View"
stroptions(3) = "Main document"
stroptions(4) = "Undo"
stroptions(5) = "Ctrl C"
stroptions(6) = "Selection"
stroptions(7) = "Ctrl V"
stroptions(8) = "Any"
stroptions(9) = "Ctrl"
stroptions(10) = "Ctrl X"

stranswers(1) = "Main Document"
stranswers(2) = "any"
stranswers(3) = "Ctrl-X"
stranswers(4) = "Ctrl-C"
stranswers(5) = "Ctrl-V"
stranswers(6) = "Undo"
stranswers(7) = "Paragraph"
stranswers(8) = "Ctrl"
stranswers(9) = "Selection"
stranswers(10) = "View"

End Sub

Private Sub Fill_in_the_Blanks_Paper9()
StrQuestions(1) = "dash memory is volatile memory."
StrQuestions(2) = "A prerequisite to dash is data collection."
StrQuestions(3) = "The computer keyboard layout is called dash layout."
StrQuestions(4) = "dash unit can be considered both as a secondary storage device and as an I\O device."
StrQuestions(5) = "The term CAM stands for dash."
StrQuestions(6) = "In dash language programming address  is symbolic not absolute."
StrQuestions(7) = "The object program is executed at dash ."
StrQuestions(8) = "With an dash the source program needs to tbe interpreted every tim it is to run."
StrQuestions(9) = "COBOL is a suitable language, for dash applications"
StrQuestions(10) = "Connecting different comoputers in an organised manner is called dash."

stroptions(1) = "Run time"
stroptions(2) = "Networking"
stroptions(3) = "Interpretor"
stroptions(4) = "Magnetic tape"
stroptions(5) = "Assembly"
stroptions(6) = "Data processing"
stroptions(7) = "Business"
stroptions(8) = "QWERTY"
stroptions(9) = "RAM"
stroptions(10) = "Computer AIded Manufacturing"

stranswers(1) = "RAM"
stranswers(2) = "Data Processing"
stranswers(3) = "Qwerty"
stranswers(4) = "magnetic Tape"
stranswers(5) = "Computer Aided Manufacturing"
stranswers(6) = "Assembly"
stranswers(7) = "Run Time"
stranswers(8) = "Interpreter"
stranswers(9) = "Business"
stranswers(10) = "Networking"

End Sub

Private Sub Fill_in_the_Blanks_Paper10()
StrQuestions(1) = "A popular multi user Operating System is dash"
StrQuestions(2) = "Hexadecimal numbers are dash bits wide."
StrQuestions(3) = "The term BIOS stands for dash"
StrQuestions(4) = "One billionth of one second is dash"
StrQuestions(5) = "If " & Chr(26) & "A" & Chr(26) & " is for decimal 10, then " & Chr(26) & "F" & Chr(26) & " is for dash"
StrQuestions(6) = "The term BPS stands for dash"
StrQuestions(7) = "Software designed for a spedividapplication,such as pat calculation , processing ofexamination reslt etc. are known as dash"
StrQuestions(8) = "The dash is used to create, modify and store a text file."
StrQuestions(9) = "In a dash system, the responce time is crucially important."
StrQuestions(10) = "An assembler translates dash language programs into machine language."

stroptions(1) = "Decimal"
stroptions(2) = "Assembly"
stroptions(3) = "Four"
stroptions(4) = "Text Editor"
stroptions(5) = "One nano Second"
stroptions(6) = "Real time system"
stroptions(7) = "Bits per second"
stroptions(8) = "Basic I/O system"
stroptions(9) = "Application software"
stroptions(10) = "Unix"

stranswers(1) = "UNIX"
stranswers(2) = "Four"
stranswers(3) = "Basic I/O System"
stranswers(4) = "One Nano Sec."
stranswers(5) = "Decimal"
stranswers(6) = "Bits per Second"
stranswers(7) = "Application Software"
stranswers(8) = "Text Editor"
stranswers(9) = "Real Time"
stranswers(10) = "Assembly"

End Sub

Private Sub Fill_in_the_Blanks_Paper11()
StrQuestions(1) = "There exists a dash correspondence  between assembly language instructions and machine language instructions."
StrQuestions(2) = "A dash memory is faster than main memory."
StrQuestions(3) = "Assembly langage instructions use dash addresses."
StrQuestions(4) = "The dash printes are used for DTP applications"
StrQuestions(5) = "The dash analysis phase of the compilation process groups the characters of the source code into tokens."
StrQuestions(6) = "The dash module, ofthe Operating  System maintains directories, which contain information about files present inthe system."
StrQuestions(7) = "The language that the computer understands without requiring any translation is called dash"
StrQuestions(8) = "The term used to denote an error in a computer program is dash"
StrQuestions(9) = "A source  program contains instructions written in a dash language like BASIC or COBOL."
StrQuestions(10) = "One of the main functions of the Operating Systems is dash."

stroptions(1) = "Machine langage"
stroptions(2) = "Memory management"
stroptions(3) = "Laser"
stroptions(4) = "High level"
stroptions(5) = "One to One"
stroptions(6) = "Bug"
stroptions(7) = "File management"
stroptions(8) = "Symbolic"
stroptions(9) = "Lexical"
stroptions(10) = "Cache"

stranswers(1) = "one to one"
stranswers(2) = "cache"
stranswers(3) = "symbolic"
stranswers(4) = "Laser"
stranswers(5) = "Lexical"
stranswers(6) = "File management"
stranswers(7) = "machine Language"
stranswers(8) = "Bug"
stranswers(9) = "High Level"
stranswers(10) = "Memory management"

End Sub

Private Sub Fill_in_the_Blanks_Paper12()
StrQuestions(1) = "dash are used for weather forecasting applications."
StrQuestions(2) = "A small busines which cannot afford to permanently hire computer professonals can still make use of a computer by using a dash."
StrQuestions(3) = "We can instruct the Operating System to carry out a sequence of tasks using dash"
StrQuestions(4) = "Applications program generation is a feature of dash"
StrQuestions(5) = "Machine language codes are dash dependent"
StrQuestions(6) = "dash is the output of a compiler."
StrQuestions(7) = "dash is the interleaved exectition of two or more  different and independent programs bu the same computer."
StrQuestions(8) = "There exiss a dash correspondence between high level statements and machine language statements. "
StrQuestions(9) = "If there is a circular reference in worksheet, it displays dash in the status line."
StrQuestions(10) = "If AUTOEXEC.BAT is for MS-DOs then dash is for WINDOWS"

stroptions(1) = "4GL"
stroptions(2) = "One to many"
stroptions(3) = "Hardware"
stroptions(4) = "Multitasking"
stroptions(5) = "Circular"
stroptions(6) = "Startup"
stroptions(7) = "JCL"
stroptions(8) = "Custom designed pacage"
stroptions(9) = "Object program"
stroptions(10) = "Super computers"

stranswers(1) = "Super Computers"
stranswers(2) = "Custom Designed Package"
stranswers(3) = "JCL"
stranswers(4) = "4GL"
stranswers(5) = "Hardware"
stranswers(6) = "Object Program"
stranswers(7) = "Multitasking"
stranswers(8) = "One to Many"
stranswers(9) = "Circular"
stranswers(10) = "Startup"

End Sub


Private Sub Matching_Columns_Paper1()
strcolumnA(1) = "XCOPY"
strcolumnA(2) = "CLS"
strcolumnA(3) = "VOL"
strcolumnA(4) = "DATE"
strcolumnA(5) = "ERASE"
strcolumnA(6) = "MSBACKUP"
strcolumnA(7) = "MSAV"
strcolumnA(8) = "SCANDISK"
strcolumnA(9) = "TREE"
strcolumnA(10) = "DISKCOPY"

strcolumnB(1) = "Better than CHKDSK"
strcolumnB(2) = "Directory structure"
strcolumnB(3) = "Anti Virus utility"
strcolumnB(4) = "Backup and restore utility"
strcolumnB(5) = "Makes a copy of diskette"
strcolumnB(6) = "Clears the screen"
strcolumnB(7) = "Displays date and accepts a new date"
strcolumnB(8) = "DEL"
strcolumnB(9) = "Faster than COPY"
strcolumnB(10) = "Internal command"

stranswers(1) = "Faster than COPY"
stranswers(2) = "Clears the screen"
stranswers(3) = "Internal command"
stranswers(4) = "Displays date and accepts a new date"
stranswers(5) = "DEL"
stranswers(6) = "Backup and restore utility"
stranswers(7) = "Anti Virus utility"
stranswers(8) = "Better than CHKDSK"
stranswers(9) = "Directory structure"
stranswers(10) = "Makes a copy of diskette"

End Sub
Private Sub Matching_Columns_Paper2()
strcolumnA(1) = "REN"
strcolumnA(2) = "MD"
strcolumnA(3) = "VER"
strcolumnA(4) = "a:\>"
strcolumnA(5) = "Primary filename"
strcolumnA(6) = "?"
strcolumnA(7) = "/W"
strcolumnA(8) = "/P"
strcolumnA(9) = "Attrib"
strcolumnA(10) = "MEMMAKER"

strcolumnB(1) = "Does not display file size"
strcolumnB(2) = "DOS version number"
strcolumnB(3) = "Makes file read only."
strcolumnB(4) = "Maximum 8 characters"
strcolumnB(5) = "RENAME"
strcolumnB(6) = "Pauses after displaying one page"
strcolumnB(7) = "Improves memory usage"
strcolumnB(8) = "MKDIR"
strcolumnB(9) = "Root directory"
strcolumnB(10) = "Wildcard character"

stranswers(1) = "RENAME"
stranswers(2) = "MKDIR"
stranswers(3) = "DOS version number"
stranswers(4) = "Root directory"
stranswers(5) = "Maximum 8 characters"
stranswers(6) = "Wildcard character"
stranswers(7) = "Does not display file size"
stranswers(8) = "Pauses after displaying one page"
stranswers(9) = "Attrib"
stranswers(10) = "Improves memory usage"

End Sub

Private Sub Matching_Columns_Paper3()
strcolumnA(1) = "SYS"
strcolumnA(2) = "FILES"
strcolumnA(3) = "Sectors"
strcolumnA(4) = "Clusters"
strcolumnA(5) = "Hard disk"
strcolumnA(6) = "Lost clusters"
strcolumnA(7) = "CHOICE command"
strcolumnA(8) = "Directory entry"

strcolumnB(1) = "FILE0001.CHK"
strcolumnB(2) = "Partition table"
strcolumnB(3) = "AUTOEXEC.BAT"
strcolumnB(4) = "Name of a file"
strcolumnB(5) = "Transfers DOS"
strcolumnB(6) = "Logic dimension of track"
strcolumnB(7) = "Minimum unit of disk storage"
strcolumnB(8) = "Used in CONFIG.SYS"

stranswers(1) = "Transfers DOS"
stranswers(2) = "Used in CONFIG.SYS"
stranswers(3) = "Logic dimension of track"
stranswers(4) = "Minimum unit of disk storage"
stranswers(5) = "Partition table"
stranswers(6) = "FILE0001.CHK"
stranswers(7) = "AUTOEXEC.BAT"
stranswers(8) = "Name of a file"

End Sub
Private Sub Matching_Columns_Paper4()
strcolumnA(1) = "Formula"
strcolumnA(2) = "Legend"
strcolumnA(3) = "Macro"
strcolumnA(4) = "Home"
strcolumnA(5) = "F4"
strcolumnA(6) = "MOD()"
strcolumnA(7) = "+D$4"
strcolumnA(8) = "IV16384"
strcolumnA(9) = "NOW()"
strcolumnA(10) = "Tab"

strcolumnB(1) = "Changes addressing mode"
strcolumnB(2) = "Current date/time"
strcolumnB(3) = "Remainder"
strcolumnB(4) = "Moves to the next column"
strcolumnB(5) = "Mixed addressing"
strcolumnB(6) = "Begins with ="
strcolumnB(7) = "Chart"
strcolumnB(8) = "Moves to the first column"
strcolumnB(9) = "Address of the last cell"
strcolumnB(10) = "Automate a task"

stranswers(1) = "Begins with ="
stranswers(2) = "Chart"
stranswers(3) = "Automate a task"
stranswers(4) = "Moves to the first column"
stranswers(5) = "Changes addressing mode"
stranswers(6) = "Remainder"
stranswers(7) = "Mixed addressing"
stranswers(8) = "Address of the last cell"
stranswers(9) = "Current date/time"
stranswers(10) = "Moves to the next column"

End Sub

Private Sub Matching_Columns_Paper5()
strcolumnA(1) = "TODAY()"
strcolumnA(2) = "256"
strcolumnA(3) = "Autofill"
strcolumnA(4) = "MIN()"
strcolumnA(5) = "PMT()"
strcolumnA(6) = "COUNT()"
strcolumnA(7) = "Data extract"
strcolumnA(8) = "ROUND()"
strcolumnA(9) = "PV()"
strcolumnA(10) = "Data sort"

strcolumnB(1) = "Equated monthly instalment"
strcolumnB(2) = "Rounds off a number"
strcolumnB(3) = "Returns the minimum of numbers"
strcolumnB(4) = "Arranges data in a range in order"
strcolumnB(5) = "Number of columns in worksheet"
strcolumnB(6) = "Present value of an investment"
strcolumnB(7) = "Current date"
strcolumnB(8) = "Number of non blank cells"
strcolumnB(9) = "Creates a series"
strcolumnB(10) = "Copies matching records from the database"

stranswers(1) = "Current date"
stranswers(2) = "Number of columns in worksheet"
stranswers(3) = "Creates a series"
stranswers(4) = "Returns the minimum of numbers"
stranswers(5) = "Equated monthly instalment"
stranswers(6) = "Number of non blank cells"
stranswers(7) = "Copies matching records from the database"
stranswers(8) = "Rounds off a number"
stranswers(9) = "Present value of an investment"
stranswers(10) = "Arranges data in a range in order"

End Sub

Private Sub Matching_Columns_Paper6()
strcolumnA(1) = "To quit WINDOWS"
strcolumnA(2) = "To switch tasks"
strcolumnA(3) = "GANESH.BMP"
strcolumnA(4) = "WINDOWS system settings"
strcolumnA(5) = "Control Menu"
strcolumnA(6) = "WINDOWS with program manager"
strcolumnA(7) = "Default PIF"
strcolumnA(8) = "Moving from one option to another"
strcolumnA(9) = "Closing an active group"
strcolumnA(10) = "PROMT SCREEN key"

strcolumnB(1) = "DOS application"
strcolumnB(2) = "TAB key"
strcolumnB(3) = "Top left corner"
strcolumnB(4) = "CTRL + F4"
strcolumnB(5) = "Copying the image of a screen into clipboard"
strcolumnB(6) = "Alt + Tab"
strcolumnB(7) = "Group WINDOW"
strcolumnB(8) = "Bitmap file"
strcolumnB(9) = "Alt + F4"
strcolumnB(10) = "SYSYEM.INI"

stranswers(1) = "Alt + F4"
stranswers(2) = "Alt + Tab"
stranswers(3) = "Bitmap file"
stranswers(4) = "SYSYEM.INI"
stranswers(5) = "Top left corner"
stranswers(6) = "Group WINDOW"
stranswers(7) = "DOS application"
stranswers(8) = "TAB key"
stranswers(9) = "CTRL + F4"
stranswers(10) = "Copying the image of a screen into clipboard"

End Sub

Private Sub Matching_Columns_Paper7()
strcolumnA(1) = "WIN/2"
strcolumnA(2) = "WIN/3"
strcolumnA(3) = "WINDOWS help"
strcolumnA(4) = "Menu item followed by triangular button"
strcolumnA(5) = "Menu item that has ellipses following the word or pharase"
strcolumnA(6) = "DDE"

strcolumnB(1) = "Transfers data or instruction between WINDOWS applications"
strcolumnB(2) = "Cascading Menu item"
strcolumnB(3) = "Dialog box"
strcolumnB(4) = "Enhanced mode of WINDOWS"
strcolumnB(5) = "F1 key"
strcolumnB(6) = "Standard mode of WINDOWS"

stranswers(1) = "Standard mode of WINDOWS"
stranswers(2) = "Enhanced mode of WINDOWS"
stranswers(3) = "F1 key"
stranswers(4) = "Cascading Menu item"
stranswers(5) = "Dialog box"
stranswers(6) = "Transfers data or instruction between WINDOWS applications"
End Sub
Private Sub Matching_Columns_Paper8()
strcolumnA(1) = "Paragraph formatting"
strcolumnA(2) = "1.5"
strcolumnA(3) = "^V"
strcolumnA(4) = "^I"
strcolumnA(5) = "^B"
strcolumnA(6) = "Autocorect"
strcolumnA(7) = "Paragraph mark"
strcolumnA(8) = "Double click"
strcolumnA(9) = "Point"
strcolumnA(10) = "^U"

strcolumnB(1) = "Italics"
strcolumnB(2) = "Inserts new text"
strcolumnB(3) = "Stores all formatting information"
strcolumnB(4) = "Underline"
strcolumnB(5) = "Paste text "
strcolumnB(6) = "Selects a word"
strcolumnB(7) = "Character size"
strcolumnB(8) = "Boldface"
strcolumnB(9) = "Keep with next"
strcolumnB(10) = "Line spacing"

stranswers(1) = "Keep with next"
stranswers(2) = "Line spacing"
stranswers(3) = "Paste text "
stranswers(4) = "Italics"
stranswers(5) = "Boldface"
stranswers(6) = "Inserts new text"
stranswers(7) = "Stores all formatting information"
stranswers(8) = "Selects a word"
stranswers(9) = "Character size"
stranswers(10) = "Underline"

End Sub

Private Sub Matching_Columns_Paper9()
strcolumnA(1) = "^X"
strcolumnA(2) = "^C"
strcolumnA(3) = "Header"
strcolumnA(4) = "Leader character"
strcolumnA(5) = "Decimal tab"
strcolumnA(6) = "Main document"
strcolumnA(7) = "Print priview"
strcolumnA(8) = "Frame"
strcolumnA(9) = "Courier"
strcolumnA(10) = "Deletes selected text"

strcolumnB(1) = "WYSIWIG"
strcolumnB(2) = "Displayed in top margin"
strcolumnB(3) = "Backspace"
strcolumnB(4) = "Contains graphics"
strcolumnB(5) = "Font"
strcolumnB(6) = "Copies text in clipboard"
strcolumnB(7) = "MailMerge"
strcolumnB(8) = "Cuts selected text"
strcolumnB(9) = "Fills up to the next tab"
strcolumnB(10) = "Aligns at decimal point"

stranswers(1) = "Cuts selected text"
stranswers(2) = "Copies text in clipboard"
stranswers(3) = "Displayed in top margin"
stranswers(4) = "Fills up to the next tab"
stranswers(5) = "Aligns at decimal point"
stranswers(6) = "MailMerge"
stranswers(7) = "WYSIWIG"
stranswers(8) = "Contains graphics"
stranswers(9) = "Font"
stranswers(10) = "Backspace"

End Sub

Private Sub Matching_Columns_Paper10()
strcolumnA(1) = "Machine language"
strcolumnA(2) = "Algorithm"
strcolumnA(3) = "Firmware"
strcolumnA(4) = "QWERTY"
strcolumnA(5) = "DTP"
strcolumnA(6) = "One byte"
strcolumnA(7) = "Bug"
strcolumnA(8) = "Oneword"
strcolumnA(9) = "LAN"
strcolumnA(10) = "One MB"

strcolumnB(1) = "Eight bits"
strcolumnB(2) = "1024 KB"
strcolumnB(3) = "Computer keyboard"
strcolumnB(4) = "Error in a computer program"
strcolumnB(5) = "Ethernet network"
strcolumnB(6) = "Sixteen bits"
strcolumnB(7) = "Microcode stored in ROM"
strcolumnB(8) = "What you see is what you get"
strcolumnB(9) = "A sequence of instructions which, when followed properly solves a problem."
strcolumnB(10) = "The language which communicates with the computer using only the binary digits 1 and 0"

stranswers(1) = "The language which communicates with the computer using only the binary digits 1 and 0"
stranswers(2) = "A sequence of instructions which, when followed properly solves a problem."
stranswers(3) = "Microcode stored in ROM"
stranswers(4) = "Computer keyboard"
stranswers(5) = "What you see is what you get"
stranswers(6) = "Eight bits"
stranswers(7) = "Error in a computer program"
stranswers(8) = "Sixteen bits"
stranswers(9) = "Ethernet network"
stranswers(10) = "1024 KB"

End Sub

Private Sub Matching_Columns_Paper11()
strcolumnA(1) = "Megahertz"
strcolumnA(2) = "Operating System"
strcolumnA(3) = "Microprocessor"
strcolumnA(4) = "MOUSE"
strcolumnA(5) = "Compiler"
strcolumnA(6) = "Assembly"
strcolumnA(7) = "COBOL"
strcolumnA(8) = "word processor"
strcolumnA(9) = "Line editor"
strcolumnA(10) = "Multi window editor"

strcolumnB(1) = "An editor where the basic unit of operation is a line."
strcolumnB(2) = "A translator which transforms a high level language into a particular computer's machine language."
strcolumnB(3) = "Language for business applications"
strcolumnB(4) = "The editor with the help of which a user can edit as eparate file in each window"
strcolumnB(5) = "A collection of programs that controls the execution of jobs on the computer"
strcolumnB(6) = "A program that helps to add, delete or corect information in a document "
strcolumnB(7) = "An input device"
strcolumnB(8) = "Measure of Micropocessor speed"
strcolumnB(9) = "A unit of the computer system that interprets  instructions and executes them"
strcolumnB(10) = "Language for real time and system programming "

stranswers(1) = "Measure of Micropocessor speed"
stranswers(2) = "A collection of programs that controls the execution of jobs on the computer"
stranswers(3) = "A unit ofthe computer system that interprets  instructions and executes them"
stranswers(4) = "An input device"
stranswers(5) = "A translator which transforms a high level language into a particular computer's machine language."
stranswers(6) = "A unit of the computer system that interprets  instructions and executes them"
stranswers(7) = "Language for business applications"
stranswers(8) = "A program that helps to add, delete or corect information in a document "
stranswers(9) = "An editor where the basic unit of operation is a line."
stranswers(10) = "The editor with the help of which a user can edit as eparate file in each window"
End Sub
Private Sub Matching_Columns_Paper12()
strcolumnA(1) = "Peripherals"
strcolumnA(2) = "Bootstrap loader"
strcolumnA(3) = "SPOOLER"
strcolumnA(4) = "Compiler"
strcolumnA(5) = "Word processor"
strcolumnA(6) = "Portability"
strcolumnA(7) = "Text editor"
strcolumnA(8) = "OS"
strcolumnA(9) = "Editor"
strcolumnA(10) = "UNIX"

strcolumnB(1) = "Provides layered user interface"
strcolumnB(2) = "A method  to get round the bottleneck of slow speed I/O devices in multi programming environment"
strcolumnB(3) = "A program to add, delete or correct information in a document"
strcolumnB(4) = "Add, delete & correct text information in a file"
strcolumnB(5) = "Capability for converting a software to run on different types of machines"
strcolumnB(6) = "The program which loads the Operating System."
strcolumnB(7) = "Utility"
strcolumnB(8) = "Multi user OS"
strcolumnB(9) = "Tool for software development"
strcolumnB(10) = "The various input/output devices and auxilliary storage units, of a computer system."

stranswers(1) = "The various input/output devices and auxilliary storage units, of a computer system."
stranswers(2) = "The program which loads the Operating System."
stranswers(3) = "A method  to get round the bottleneck of slow speed I/O devices in multi programming environment"
stranswers(4) = "Tool for software development"
stranswers(5) = "A program to add, delete or correct information in a document"
stranswers(6) = "Capability for converting a software to run on different types of machines"
stranswers(7) = "Add, delete & correct text information in a file"
stranswers(8) = "Provides layered user interface"
stranswers(9) = "Utility"
stranswers(10) = "Multi user OS"

End Sub

Private Sub questions_paper1()
StrQuestions(1) = "a. Briefly explain how you use COPYCON to create a AUTOEXEC.BAT file.    b. List the various types of diskettes that can be used with a PC along with their storage capacity."
StrQuestions(2) = "Explain the use of the DIR commandto display a list of files  a. One page at a time   b. All files including files in sudirectories   c. All files with extension name .DOC   d. All files without displaying the file size for each file."
StrQuestions(3) = "Write down comands to delete  a. All files with the extension '.BAK' from drive A.    b. All files fro the directory C:\CHECK c. The file CHAPTER1.TXT from the root directory in dirve A.    d. All files that are 5 characters long and begin with 'F', (r.g, FCHAP, FILE1.DOC, FILE2TXT,etc.)"
StrQuestions(4) = "a. What is the difference between a file and a directory?    b. What is a hidden file?   c. How cab you find out the list of hidden files in a directory?    d. Can you make a directory hidden?"
StrQuestions(5) = "Assume that the following files ar present in the current directory: FOR1.TXT    FORX.EXE    COM.BIN COMA    FOR.DAT COM.TXT COM2.TXT    COMB.PRG    FOR2.DAT    A.BIN       COM2.EXE    COMB.PRG    Write the DOS comand to do the following tasks :    a. List all filenames that start with 'F' .  b. List all filenames having the extension '.TXT'.  c. List all finenames that contain 3 letters.   d. List all filenames that begin with 'C' and are 4 characters wide, e.g. COM2.TXT, COM2.EXE, COMA, etc."
StrQuestions(6) = "What do yu understand by the 'ROOT Directory' ?"
StrQuestions(7) = " a. What is the use of the DEFRAG command?   b. Explain the DELTREE command in brief."
StrQuestions(8) = "a. What do you nderstand by FAT? Is it applicable to floppy diskettes?  b. What is a boot record and where is it stored?"
StrQuestions(9) = "Explain with the help of an example , how you transfer the output of one command to another is DOS?"
StrQuestions(10) = "a. How do you redirect output of a command to another device say a printer or to another device say a printer or to a file?  b. Briefly explain the use of MEMMAKER."
End Sub
Private Sub questions_paper2()
StrQuestions(1) = "Compare CHKDSK with SCANDISK."
StrQuestions(2) = "Explain the  XCOPY command and the various options (switches) that can be used woth it."
StrQuestions(3) = "a. Explain the difference between XCOPY *.* A: and COPY *.* A: commands. b. Explain RENAME (REN) and MKDIR(MD) commands in brief."
StrQuestions(4) = "Explain the FORMAT comand and its options"
StrQuestions(5) = "What is the difference between internal and external DOS commands?"
StrQuestions(6) = "What is the significance of CONFIG.SYS and AUTOEXEC.BAT files in DOS? Briefly describe the use of FILES, BUFFERS, COUNTRY, DEVICE, SHELL, LASTDRIVE in CONFIG.SYS."
StrQuestions(7) = "In a PC, the root directory of drive C is selected. The FORMAT command is available in the DOS drectoruy under the root directory i ndrive C. Write down commands to format :    a. A disk in drive A    b. A disk in drive A and to also transfer DOS.  c. Only one side of a disk in drive A.  d. A disk in drive B."
StrQuestions(8) = "Write down command(s) to do/perform the following tasks: a. Display the current time b. Display the current date c. Change the current date  d. Display the directory structure  e. To define C:\PERSONAL and C:\DOS as the path for commands    f. To compare the contents of two diskettes g. Print the directory structure without displaying it on the screen"
StrQuestions(9) = "Explain the following commands in brief  a. PROMPT   b. TYPE c. PRINT    d. PATH"
StrQuestions(10) = "Describe the features of MSBACKUP"
End Sub
Private Sub questions_paper3()
StrQuestions(1) = "Write down the Booting sequence from power p to the display of the prompt"
StrQuestions(2) = "Explain what is File Allocation Table & the need of FAT?"
StrQuestions(3) = "Cite, with example , generation of a multiple confiiguration in the CONFIG.SYS file?"
StrQuestions(4) = "A CONFIG.SYS file contains the following commands.       [menu]  menuitem=one    menuitem=two    [common]    dos=high    buffers=20  [one]    files=40   device=c:\dos\himem.sys device=c:\dos\mouse.sys    [two]    files=20    device=c:\dos\mouse.sys  When you bot with the above mentioned CONFIG.SYS what will be the result?"
StrQuestions(5) = "What does the CHOICE command inthe AUTOEXEC.BAT file do? Explain the various options ofthis command?"
StrQuestions(6) = "Name the various batch commands that are supported in DOS?"
StrQuestions(7) = "a. How do you abort a batch file execution?  b. Is it posible to execute abatc file from within another batch file? If yes, then how? if No, why?"
StrQuestions(8) = "Explain briefly how one cab pass parameters to a batch file?"
End Sub
Private Sub questions_paper4()
StrQuestions(1) = "explain the following in brief" & Chr(13) & _
    "a.Menu bar" & Chr(13) & _
    "b.formula bar" & Chr(13) & _
    "c.Status bar" & Chr(13) & _
    "d.Active window"
StrQuestions(2) = "a. You enter 45+12 in a cell. The worksheet does not display 57 in the cell. Instead, it shows 45+12. What is the reason or this? How will you ensure that you get the correct result in the cell?" & Chr(13) & _
    "b. Explain the following functions inbrief:" & Chr(13) & _
    "SUM()" & Chr(13) & _
    "ABS()"
StrQuestions(3) = "How do you do the following in a worksheet?" & Chr(13) & _
    "a. Change the width of a column" & Chr(13) & _
    "b. Change the width ofa range of columns" & Chr(13) & _
    "c. Change the height of a row" & Chr(13) & _
    "d. Display text in the vertical orientation in a cell" & Chr(13) & _
    "e. Display the text in a cell in a bigger size characters" & Chr(13) & _
    "f. Enclose a cell in a hollow box" & Chr(13) & _
    "g. Enclose a cell in a shadow box."
StrQuestions(4) = "When you enter the in a cell, it is left aligned. How would you right align the text in a range? What isthe procedure to ensure that the text that you enter in a cell is automatically right alighned?"
StrQuestions(5) = "Write down command(s) that can be used to perform the following operations:" & Chr(13) & _
    "a. Open an existing worksheet" & Chr(13) & _
    "b. Delete the current column" & Chr(13) & _
    "c. Insert acolumn before column C" & Chr(13) & _
    "d. Erase a range" & Chr(13) & _
    "e. Erase the current worksheet"
StrQuestions(6) = "Consider the following worksheet:" & Chr(13) & _
    "A       B           C       D       E       F" & Chr(13) & _
"1" & Chr(13) & _
"2           sales statement" & Chr(13) & _
"3   ITem        1996            1997" & Chr(13) & _
"4   Chair 13000" & Chr(13) & _
"5   Table 5000" & Chr(13) & _
"6   Stool 3400" & Chr(13) & _
"7.  Bench       5600" & Chr(13) & _
"8" & Chr(13) & _
    "The worksheet conatins the salesfigures for 1996. The projected sales for 1997 is 10% more than that for 1996." & Chr(13) & _
    "a. Write down the formula for the sales of table for 1997 in cell C4." & Chr(13) & _
    "b. Copy this formula to other cells in column C." & Chr(13) & _
    "c. How would you formt range C4:C7 to display Rs infornt of figures in this range?" & Chr(13) & _
    "d. How would you increase the width of column C to 15?"
StrQuestions(7) = "The following worksheet contains the lis to items supplied by ABC Ltd. ti XYZ Ltd." & Chr(13) & _
    "A       B       C   D       E       F" & Chr(13) & _
"1           ABC Ltd" & Chr(13) & _
"2   Item        Quantity        Rate        Amount  Sales Tax" & Chr(13) & _
"3   Floppy              50  25.00" & Chr(13) & _
"4   Pen     400         8.00" & Chr(13) & _
"5   Pencil              300 1.50" & Chr(13) & _
"6   Note pad                250 10.00" & Chr(13) & _
"7   _________________________________________" & Chr(13) & _
"8   Total" & Chr(13) & _
"9   Sales Tax Rate          0.05" & Chr(13) & _
    "a. Write a formula in cell D3 to calculate the total amount for floppies. (amount = Quantity x Rate)" & Chr(13) & _
    "b. How would you copy formula used in cell D3 for other products?" & Chr(13) & _
    "c. The sales tax rate for all items is entered in cell C9. Write down a formula in cell E3 to calculate the sales tax for floppies." & Chr(13) & _
    "d. It is required to copy the formula used in cell E3 to other cells in column E. how would you do this? What would be the formula in cell E6?"
StrQuestions(8) = "Explain the various addressing modes supported in Excel that can be used in the formulas"
StrQuestions(9) = "Explain the following functins in brief." & Chr(13) & _
    "a. MAX()" & Chr(13) & _
    "b. LOG()" & Chr(13) & _
    "c. LN()"
StrQuestions(10) = "Explain the following functions with the help of an example:" & Chr(13) & _
    "a. PMT()" & Chr(13) & _
    "b. NPV()" & Chr(13) & _
    "c. RATE()"
End Sub
Private Sub questions_paper5()
StrQuestions(1) = "The following data isentered in a worksheet:" & Chr(13) & _
    "A       B       C   D   E       F" & Chr(13) & _
"1   Name        mARKS" & Chr(13) & _
"2   Rohit 35" & Chr(13) & _
"3   rajan 67" & Chr(13) & _
"4   Suman 45" & Chr(13) & _
"5   Swamy 65" & Chr(13) & _
"6   Babu 45" & Chr(13) & _
"7   Kumar 58" & Chr(13) & _
"8   Neena 76" & Chr(13) & _
"9" & Chr(13) & _
"10" & Chr(13) & _
    "a. How would you arrange the data in the range A2:B8 in the ascending order of names?" & Chr(13) & _
    "b. Write down a function to display the total number of students in cell A10 and the average marks in cell b10."
StrQuestions(2) = "a. In a worksheet, cell K12 has a value. A formula is to be entered in cell K15, such that if the value in cell K12 is mor than 300, the valuein cell K15 would be 1.33 times the value of cell K12. Otherwise the value in cell K15 would be 1.5 times the value in cell K12. Explain the formula that you would use to achieve this." & Chr(13) & _
    "b. Explain the following functions in brief:" & Chr(13) & _
    "RAND()      INT()"
StrQuestions(3) = "a. List down the various ways that you can use to enter a specific date in a cell." & Chr(13) & _
    "b. What is the advantage of entering a date as a date serial numer instead of entering the date as atext entry?"
StrQuestions(4) = "a. Explain the function of the select special command in the formula menu." & Chr(13) & _
    "b. If the status bar displays Circular:A15, what does it mean? How will you remove it?"
StrQuestions(5) = "Explain, how you create a chart using the chart wizard. Ater you have created a hart, how do yu modify it? List down the various types of vharts that you can create in Excel."
StrQuestions(6) = "You have created a chart based on the data in a worksheet. The chart isdisplayed in the same worksheet." & Chr(13) & _
    "a. How will you change the size of the chart?" & Chr(13) & _
    "b. If the data that was used to create the cart is changed what happens to the chart?" & Chr(13) & _
    "c. It is required to include another data series in the chart. How will you implement this?"
StrQuestions(7) = "Explain in brief the followin with reference to a chart:" & Chr(13) & _
    "a.Legend" & Chr(13) & _
    "b.Area chart" & Chr(13) & _
    "c.Data series"
StrQuestions(8) = "How do you create a pie chart? How do you explode a particular slice of the pie chart?"
StrQuestions(9) = " It is required to enter a series - 1,3,5,7,...,51 in column . A starting from row 1. How would you enter this series using the minimum number of steps? How will you enter the names of months ( January through Decemebr)in column B starting from cell B1 using the minimum number of steps?"
StrQuestions(10) = "a. A worksheet contains some data. It is required to move all data from the range A1:B6 to D1:E6. How will you accpmplish this?" & Chr(13) & _
    "b. How will ou copy data in th range A1:B6 to F1:G6?"
End Sub
Private Sub questions_paper6()
StrQuestions(1) = "List down the varrious formats that you can apply on the numeric data entered in a cell"
StrQuestions(2) = "It is required to print a specific part to the worksheet. How will you define the print range, setup the page and print the range? Can you save the print output in a file?"
StrQuestions(3) = "a.What is the significance of the header in the worksheet? b. How do you define header for a worksheet c. How will you insert the current date and page numbers in the header?"
StrQuestions(4) = " a.What is a range? b. How are cell ranges named? What are the advantages to assigning a name to a cell or a range of cells? Can you rename or delete a range name?   5.  a. How will you copy some data from a worksheet to another worksheet?    b. Explain with the help of an example how a formula can refer to a cell in some other worksheet."
StrQuestions(5) = "What is a database in a worksheet? Describe the various operations that can be performed on a database."
End Sub
Private Sub questions_paper7()
StrQuestions(1) = "What is startup group? Write down the steps to include an application icon into startup group."
StrQuestions(2) = "Write down two ways to copy a file fro one group to another?"
StrQuestions(3) = "What is a PIF and how can a PIF be edited?"
StrQuestions(4) = "What changes are required in AUTOEXEC.BAT and CONFIG.SYS in order to run WINDOWS 3.1 ?"
StrQuestions(5) = "What is the use of associate command in the file option of the file manager?"
StrQuestions(6) = "Differentiate between an applcation icon and a program icon?"
StrQuestions(7) = "What is a user interface? What are special features of GUI?"
StrQuestions(8) = "What is the difference between an application icon and a program icon?"
StrQuestions(9) = "Explain OLE."
StrQuestions(10) = "What is clipboard & how is it useful in WINDOWS based applictions?"
End Sub
Private Sub questions_paper8()
StrQuestions(1) = "Differentiate between CUI and GUI."
StrQuestions(2) = "How do you change the settings of the mouse in WINDOWS? Is it possible to swap the right and left buttons of the mouse?"
StrQuestions(3) = "What is a compound document and explain how it can be created."
StrQuestions(4) = "Explain briefly two ways to run an application in WINDOWS 3.1."
StrQuestions(5) = "What does one mean be Cascding WINDOWS and Rolling Windows?"
StrQuestions(6) = "Explain briefly the various types of menu items."
StrQuestions(7) = "Explain what is a control menu. What are the options in control menu?"
StrQuestions(8) = "What is a screen saver and explain its utility?"
StrQuestions(9) = "What is true type font?"
StrQuestions(10) = "What is a BITMAP?"
End Sub
Private Sub questions_paper9()
StrQuestions(1) = "How do you copy a portion of a text file in to another file using WRITE?"
StrQuestions(2) = "List any three applications in MAIN group and explain its use?"
End Sub
Private Sub questions_paper10()
StrQuestions(1) = "a. How do you move a paragraph from the end of a page to the beginning of the page?  b. You are editing a document . It is required to ave the modified document n a new file. How will accomplish this?  c. You have modified a document.; Now you decide to save the modified version. What will you do?"
StrQuestions(2) = "a.A document contains five paragraphs. It is required to delete the third paragraph. How will you do this?   b. How will you quicjly select a sentence? How will you copy a sentence from one part ofthe document to another?"
StrQuestions(3) = "What do you understand by the following?  a. Font b. Font size    c. Font style"
StrQuestions(4) = "a. You are editing a Word for WINDOWS document file. It is required to remove all text formating from the document and to use this document with some DOS based word processor. How will you do this?  b. It is rquired to transfer some formatted text from a document ABC.DOC to XYZ.DOC. How will you accomplish this?"
StrQuestions(5) = "List down the various types of tabs supported in WORD? How do you define tabs for a group of paragraphs?"
StrQuestions(6) = "Explain the following terms in brief: a. To leader character  b. Default tab stop c. Decimal tab  d. Right aligned tab"
StrQuestions(7) = "Consider the following paragraph:    ' MailMerge: MailMerge is a facility to prepare personalised letters by combining a master documentfile with a data file. It can also be used to print envelops and mailing labels. How do you create this type of paragraph in WORD?"
StrQuestions(8) = "Explain the follosing in brief:  a. Enclosing a pragraph in a box    b. Shadow box   c. Shading"
StrQuestions(9) = "Explain the term 'Style sheet'. What is its application in a WORD document?"
StrQuestions(10) = "With reference to formatting of a paragraph, explain the following in brief:  a. Keep with next   b. Page break before    c. Keep lines together"
End Sub
Private Sub questions_paper11()
StrQuestions(1) = " a. Briefly expain the use of formattin toolbar in WORD.   b. What is the finction of the status bar in WORD? List down the information that is displayed in the status bar."
StrQuestions(2) = "a. It is required to print page numbers at the bottom of the page. How will you achieve this?  b. How do you control the position of page number printing?  c. Is it possible to print page nuumbers at the top of the page?   d. How will you hide page number on the first page of the document?"
StrQuestions(3) = "a. What do you understand by footer? b. How do you create a footer for a document? c. How do you define the position of the footer on a page?"
StrQuestions(4) = "a.What do you undrstand by sections in a document    b. What is the need to divide a document ino sections?  c. How do you divide a docuemnt into sections?"
StrQuestions(5) = "Explain the following terms:  a. Footnotes    b. Endnotes c. Paragraph mark"
StrQuestions(6) = "a. Explain the terms. How do you superscript and subscript text in WORD?   b. Briefly explain how you will search for a particluar word in the docuemnt and replace it with a new word."
StrQuestions(7) = "a. What do you understand by the Outline View? How does it help to organise a document?  b. How do you select the Outline mode?  c. How do you promote and demote a paragraph?"
StrQuestions(8) = "How will you insert graphics, say a paintbrush picture, in a document? How will you resize and move the picture to a nw page in the same document?"
StrQuestions(9) = "Describe the AutoCorect feature in brief."
StrQuestions(10) = "Describe the AutoText feature in brief"
End Sub
Private Sub questions_paper12()
StrQuestions(1) = "What is the difference between AutoCorrect and AutoText?"
StrQuestions(2) = "Explain the featurs of the spelling check facility in WORD. Hoe do you add new words in dictionaries?"
StrQuestions(3) = "Explain the grammar check facility. How will you disable a particular grammar rule?"
StrQuestions(4) = "Explain the MailMerge facility in WORD."
StrQuestions(5) = "What do you understand by hyphenation? How do you use this facility in a WORD document?"
StrQuestions(6) = "How do you define the page size and margins for a document? Briefly describe the following terms:    a. Left margin  b. Right margin c. Gutter margin"
StrQuestions(7) = "Briefly describe the various options available in the print dialog box."
StrQuestions(8) = "You have created a main document and data souce to print form leters. Now it is required to print mailing labels. How will you accomplish this? How will you print address labels directly on the envelopes?"
StrQuestions(9) = "What are templates and wizards? How ar these helpful in WORD?"
End Sub
Private Sub questions_paper13()
StrQuestions(1) = "Explain the different generations o compuers and their characteristics."
StrQuestions(2) = "Why is it said that computers operate in GIGO mode?"
StrQuestions(3) = "What do you understand  by ther term 'PAPER-Less Ofice'?"
StrQuestions(4) = "Explain how office autmation tools enhance the productivity in a typical ofice environment."
StrQuestions(5) = "Explain thefnction of an Operating System in a computer system."
StrQuestions(6) = "a. What is a batch procesing Operating System?   b. What are its advantages?"
StrQuestions(7) = "What are the different factors involved in any data processing operation? Explain with examples."
StrQuestions(8) = "Distinguish between system software and application software?"
StrQuestions(9) = "Distinguish between multi programming and time sharin Operating System."
StrQuestions(10) = "What is a Fourth Generation language? How does it differ from a conventional High Level Language?"
End Sub
Private Sub questions_paper14()
StrQuestions(1) = "What are the difference between Assembly language and Machine language."
StrQuestions(2) = "How does a High level laguage differ from an Assembly language?"
StrQuestions(3) = "Draw and explain the block diagram of a digital computer sstem. Match the diagram with the different parts of a PC system"
StrQuestions(4) = "Distiguish between Virtual memory and Real memory"
StrQuestions(5) = "If a computer's memory can store one Kilo Bits(1KB), how many nibbles, Bytes, words, and doule words can it store?"
StrQuestions(6) = "Distinguish between Compilers and Interpretors."
StrQuestions(7) = "What are the functions of the two passes of an asembler?"
StrQuestions(8) = "Why are binary digits used to code data to be stored in a compuer?"
StrQuestions(9) = "What is ASCII code? Explain its use?"
StrQuestions(10) = "What do you understand by ISCII code(s)?"
End Sub
Private Sub questions_paper15()
StrQuestions(1) = "Explain  how graphical information such as graphs can be input/output to/from acomputer?"
StrQuestions(2) = "What are the differences between Line editors and Screen editors?"
StrQuestions(3) = "What do you mean by SPOOLING?"
StrQuestions(4) = "Illustrate how the files in a computer system may be organised in an hierarchical structure."
StrQuestions(5) = "List the various characteristics of High Level Languages"
StrQuestions(6) = "Explain the meaning of the following :  a. Hardware is a one time expense, whereas Software is a continuing expense  b. Computer Viruses behave in the same manner as its bioloogical counterpart.   c. 'Prevention is better than cure' - in the context of Computer Viruses."
StrQuestions(7) = "Give three factors that make High Level anguages more advantageous than machine level languages."
StrQuestions(8) = "What is a bootstrap program?"
StrQuestions(9) = "Describe the advantages and disadvantages of timesharing Operating System."
StrQuestions(10) = "Distinguish between single user and multi user Operating Systems with examples."
End Sub
Private Sub questions_paper16()
StrQuestions(1) = "Write a short note about macros."
StrQuestions(2) = "What so you understand by the TSR programs?"
StrQuestions(3) = "What do you inderstand by the Source program and the Object program?"
StrQuestions(4) = "Describe the role of the Operating System supervisor in Operating Systems."
StrQuestions(5) = "Explain the kernel of the Operating System."
StrQuestions(6) = "What is meant by linking? Why it is required?"
StrQuestions(7) = "What is Multi- programming?"
StrQuestions(8) = "Define 'Virtual Memory'."
StrQuestions(9) = "What are the similarities and differences between a cache memory and a Virtual memory."
StrQuestions(10) = "Why is an Operating system required for a computer?"
End Sub
Private Sub questions_paper17()
StrQuestions(1) = "What is the purpose of ROM in a PC System?"
StrQuestions(2) = "What ae the main parameters that are used to classify computers?"
StrQuestions(3) = "Define a Pixel."
StrQuestions(4) = "What do you understand by a 'Virus attack'?"
StrQuestions(5) = "What are the precautions that use to keep the Computer(s) 'Virus-free'?"
StrQuestions(6) = "What do you understand by 'Software Piracy'?"
StrQuestions(7) = "Write down the steps required to make a 'Virus' infected system to a 'Virus-free' system."
StrQuestions(8) = "What is 'Networking'?"
StrQuestions(9) = "Write short notes onthe following:  a. Computer Aided Design    b. Computer Aided Instuction    c. MIS System   d. LAN  e. Network Topology f. Firmware and liveware    g. Computer viruses h. Cache mamory i. 4GL  j. Multi Media systems  k. Internet l. GUI  m. Boot viruses n. Translators  o. Client Server systems    p. Electronic Mail  q. World Wide Web   r. Desk Top Publishing  s. Applications of Super Computers  t. Note book computers  u. Anti-virus software  v. Indian language fonts    w. Paper - Less office  x. Office automation tools  y. UNIX Operating System    z. Laser printers versus Ink jet printers"
End Sub

Private Sub TrueorFalse_Paper1()
StrQuestions(1) = "MSAV is an anti virus utility"
StrQuestions(2) = "FAT is used in hard disk orbit not in floppy diskette."
StrQuestions(3) = "The COPY command can be used to combine three files."
StrQuestions(4) = "The command' COPY ABC.TXT + XYZ.DOC NEW.DOC' is invalid."
StrQuestions(5) = "The MOVE command can be used to rename adirectory."
StrQuestions(6) = "The PRINT command transfers the file(s) to be printed to a print queue."
StrQuestions(7) = "'SORT > LIST.TXT' is an invalid command."
StrQuestions(8) = "DIR | MORE is avalid command."
StrQuestions(9) = "The 'FORMAT command cannot format a hard disk."
StrQuestions(10) = "FORMAT can also transfer the DOS system after formatting a floppy diskette."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"

End Sub

Private Sub TrueorFalse_Paper2()
StrQuestions(1) = "DIR and DIRECTORY are similar commands."
StrQuestions(2) = "The diskette volume name cannot be changed."
StrQuestions(3) = "CHKDSK/F is used to fix the file system on the disk."
StrQuestions(4) = "MEM and CLS are internal DOS commands."
StrQuestions(5) = "REN cannot be used with wild card characters."
StrQuestions(6) = "It is possible to display the system  date as command prompt."
StrQuestions(7) = "CLEAR a DOS command which clears all information on the screen"
StrQuestions(8) = "MKDIR or Md is used to create a directory on the hard disk."
StrQuestions(9) = "It is not possible to use the DISKCOPY operation on a PC with a single floppy disk drive."
StrQuestions(10) = "In a PC, drive A is a 1.2 MB floppy disk drive and drive B is a 1.44 MB floppy disk drive. We can use DISKCOPY A: B: to copy contents of 1.2 MB floppy diskette into a 1.44 MB diskette."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper3()
StrQuestions(1) = "The TREE command can list only the file structre but not the list of the files."
StrQuestions(2) = "FAT keeps track of the names of the files in disk."
StrQuestions(3) = "CONFIG.SYS file is execited after AUTOEXEC>BAT file while bootinh."
StrQuestions(4) = "The command DEVICE = C:\DOS\HIMEM.SYS is placed in AUTOEXEC.BAT file to enable the usage of extended memory."
StrQuestions(5) = "By pressing CTRL - D one can abort the execution of AUTOEXEC.BAT. dile."
StrQuestions(6) = "Boot record on a disk contains the complete Operating System."
StrQuestions(7) = "Directory entry n a diskette follows the FAT."
StrQuestions(8) = "Boot record on a diskette contains information regarding the type of diskette."
StrQuestions(9) = "A high density diskette cannot be formatted as a low density diskette."
StrQuestions(10) = "A DSHD diskette cannot be read by a double sided dou le density drive."
stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"

End Sub

Private Sub TrueorFalse_Paper4()
StrQuestions(1) = "Tracks form a part of sectors in a floppy diskette."
StrQuestions(2) = "Hard disk can have more than two heads."
StrQuestions(3) = "Only one copy of FAT is stored in the floppy diskette."
StrQuestions(4) = "Index holes in a floppy are meant for indexing files."
StrQuestions(5) = "One can store only a limited number of files in the root directory of a DSDD floppy even though enough space is available."
StrQuestions(6) = "It is not possible to bypass the execition of CONFIG.SYS and AUTOEXEC.BAT file while booting."
StrQuestions(7) = "It is possible to execute the commands in the CONFIG.SYS file step by step."
StrQuestions(8) = "ONLY .EXE and .COM files can be executed in DOS."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
End Sub

Private Sub TrueorFalse_Paper5()
StrQuestions(1) = "An object such as a paintbrush picture and word document can be inserted in a worksheet."
StrQuestions(2) = "If you enter12 + XYZ in a cell, Excel displays a error message."
StrQuestions(3) = "If we enter '13/5/96' or '13-MAy-96' in a cell, Excel gives an error message."
StrQuestions(4) = "+F4+f5 and =f4 +f5 are similar foumulas."
StrQuestions(5) = "Both, NOW() and TODAY() can be used to display the current time."
StrQuestions(6) = "To display the current time, the TIME() function can be used."
StrQuestions(7) = "Macros that are stored in a macro sheet work with all worksheets."
StrQuestions(8) = "Once a macro is running , it is possible to stop it."
StrQuestions(9) = "To run a macro,select macro from the run menu."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper6()
StrQuestions(1) = "Macros stored in the global macro sheet can be used with any worksheet."
StrQuestions(2) = "A macro cab ne assigned to a button."
StrQuestions(3) = "The ROUND() function can be used to round off a number to hundreds."
StrQuestions(4) = "The RAND() function returns a random nmber between 0 and 10."
StrQuestions(5) = "Depending on the task being performed the shape of the mouse pointer changes automatically."
StrQuestions(6) = "More than one window can be open on the screen at one time . However we can work only in one window at a time "
StrQuestions(7) = "Workbooks are used in DTP."
StrQuestions(8) = "If one forgets to put the closing parenthesis in formula and presses enter, Excel gives an error message."
StrQuestions(9) = "To check spellings of the text in a worksheet, we need to take the help of a word processor."
StrQuestions(10) = "It is impossible to save a worksheet in such a way that any one can open it and view it; but only those knowing the password  can modify it."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper7()
StrQuestions(1) = "To create a new worksheet, select worksheet from the file menu."
StrQuestions(2) = "It is not possile to select two non adjacent ranges at atime."
StrQuestions(3) = "The select special commane, in the formula menu, is used to select cells with specific types of cell contents."
StrQuestions(4) = "You have creates a new worksheet . If you enter a number in a cell, Excel automatically assigns the number a built in format bxased on the typed number."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
End Sub

Private Sub TrueorFalse_Paper8()
StrQuestions(1) = "A document having object links will have biger size than an object embedded in it."
StrQuestions(2) = "WINDOWS 3.1 runs in two modes."
StrQuestions(3) = "WINDOWs 3.1 is an Operating System."
StrQuestions(4) = "It is necessary to load MS - DOS first before running WINDOWS."
StrQuestions(5) = "Multiple programs can be executed simultaneously in WINDOWS."
StrQuestions(6) = "WINDOWS provides the user a good GUI."
StrQuestions(7) = "Without a mouse one cannot work with WINDOWS."
StrQuestions(8) = "A group  can have more than one application program."
StrQuestions(9) = "Minimising WINDOWS is same as closing WINDOWS."
StrQuestions(10) = "It is possible to chang the speed of mouse movement across WINDOWS ?"

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper9()
StrQuestions(1) = "It is possible to swap the right and left button of a mouse in WINDOWS"
StrQuestions(2) = "All WINDOWS applications can be executed in WINDOWS or DOS mode."
StrQuestions(3) = "It is possible to have individual strartup groups for different users."
StrQuestions(4) = "One can use file ,manager to run an application."
StrQuestions(5) = "It is not possible to move a file from one directory to another using file manager."
StrQuestions(6) = "Within a group, one can have at most 10 applications."
StrQuestions(7) = "It is necessary that all application files in a group should resides in the same directory."
StrQuestions(8) = "It is not possible to have more than one application in a startup group."
StrQuestions(9) = "Applications can be moved from one group to another using mouse dragging."
StrQuestions(10) = "It is possible to run a DOS application as an icon."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper10()
StrQuestions(1) = "It is not possible to five command line parameters for a DOS app;ication running in WINDOWS."
StrQuestions(2) = "The clipboard contents are treates as one unit."
StrQuestions(3) = "It is not possible to paste a portion of the cilpboard content."
StrQuestions(4) = "Embedding an object into a document, the object becomes part of the document."
StrQuestions(5) = "When you link an object, the actual data remains separate from the client data."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
End Sub

Private Sub TrueorFalse_Paper11()
StrQuestions(1) = "If you double click mouse on a Word, it selects that word."
StrQuestions(2) = "AutoText can be used to insert grapics in the document."
StrQuestions(3) = "WORD includes a number of AutoCorrect entries. We can use the entries but cannot modify them."
StrQuestions(4) = "Bold, Italics and bold italics are available for all fonts."
StrQuestions(5) = "^B selects the italics style for the selected text."
StrQuestions(6) = "^U selects the italics style for the selected text."
StrQuestions(7) = "^I selects the underline style for the selected text."
StrQuestions(8) = "A 24 point chanacter isabout 1 inch high."
StrQuestions(9) = "The current date cannot be inserted in the header."
StrQuestions(10) = "The current date can be inserted in the footer."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper12()
StrQuestions(1) = "WORD prints header in the top margin."
StrQuestions(2) = "If the header is too large to fit in the top margin, WORD adjusts the top margin."
StrQuestions(3) = "The view menu is used to create header and footer."
StrQuestions(4) = "The same header is necesarily printed on all even pages of the document."
StrQuestions(5) = "Even and odd pages can have different footers."
StrQuestions(6) = "All sections in a documents necessarily use the same margin."
StrQuestions(7) = "A new section can be forced to begin the next odd page. "
StrQuestions(8) = "WORD can enclose a paragrpah in a normal as well as shadow box."
StrQuestions(9) = "While using MailMerge it is possible to preview and print the resultant merged letters selectively."
StrQuestions(10) = "Hyphenation can affect page breads."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"

End Sub

Private Sub TrueorFalse_Paper13()
StrQuestions(1) = "Top and bottom margins may change after creating header and footer."
StrQuestions(2) = "Multiple columns cannot be viewed in the pane layout view."
StrQuestions(3) = "Page preview shows the page in WYSIWYG mode."
StrQuestions(4) = "In case of multiple columns, a vertical line can be draw between columns."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
End Sub

Private Sub TrueorFalse_Paper14()
StrQuestions(1) = "No matter however simple or complex a computer may be, it can only do what it is told to do'."
StrQuestions(2) = "An algorithm is nothing but a sequence of instructions telling what to do and how."
StrQuestions(3) = "Assembly language instructions are written in binary digits."
StrQuestions(4) = "Software maintenance means careful presernvation of the floppy diskettes containing the programs and maintain them free from virus."
StrQuestions(5) = "A real time system is inherently on line."
StrQuestions(6) = "An algorithm should terminate after a finite number of steps."
StrQuestions(7) = "Railway Ticket R eservation System is not an on line processing system.,"
StrQuestions(8) = "Clock speed is one of the considerations to determine the power of a microcomputer."
StrQuestions(9) = "ASCII codes are 10 bit binary codes."
StrQuestions(10) = "4GL is a non procedural language."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"

End Sub

Private Sub TrueorFalse_Paper15()
StrQuestions(1) = "An electronic calculator is really a small computer."
StrQuestions(2) = "A compter is intelligent enough to understand an ambiguity in an algorithm."
StrQuestions(3) = "A program, written in a machine language of a particular computer, can be executed in a different computer."
StrQuestions(4) = "The use of machine language requires some knowledge of different registers in the CPU."
StrQuestions(5) = "Assembly language allows us to use symbolic names for operands."
StrQuestions(6) = " Satellites can be used for communicating with different computers."
StrQuestions(7) = "A fully compiled or assembled program, that is ready to be loaded into the computer, is known as the object program."
StrQuestions(8) = "A system software package that converts a HIgh Level laguage program to machine language program is called an Assembler."
StrQuestions(9) = "A software used to convert source progtam instructions, to machine language instructions is called a language prosessor."
StrQuestions(10) = "An object program is created by a compiler, from a sorurce program at run time."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper16()
StrQuestions(1) = "An interpretor creates  an object program from a source program."
StrQuestions(2) = "SQL is a database sub language."
StrQuestions(3) = "Programs available in CD - ROMs are always 'Virus' free."
StrQuestions(4) = "Computer software is normally classified as system software and application software."
StrQuestions(5) = "System software is intended to satisfy you particular business or personal processing requirements."
StrQuestions(6) = "Compiler translates pseudocode into macime language."
StrQuestions(7) = "DTP packages can also perform word processing functions."
StrQuestions(8) = "ASCII code is a standard code used or textual data transmission."
StrQuestions(9) = "Firmware programs reside in the ROM portin of a computer."
StrQuestions(10) = "'Viruses' are not computer programs."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper17()
StrQuestions(1) = "Systems software direct the proper fuctioning of a computer."
StrQuestions(2) = "A time sharing system is essentially a single user system."
StrQuestions(3) = "UNIX is a multiuser, time sharing system."
StrQuestions(4) = "A linker can combine object modules from multiple files."
StrQuestions(5) = "Editor is a software used to interactively review and modify text matrial."
StrQuestions(6) = "Line editor gives more facilities to the user than screen editors."
StrQuestions(7) = "Machine language is machine dependent while assembly languages are machine independent."
StrQuestions(8) = "Computer can perform almost any task, provided it can be reduced to series of logical computational steps."
StrQuestions(9) = "The mnemonic code does not vary from computer to computer."
StrQuestions(10) = "The input to the syntax analyser is the output of the lexical analyser."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper18()
StrQuestions(1) = "All commands of DOS or Word Processing software can be entered in uper case or lower case letters or a mix of both."
StrQuestions(2) = "Bootstrap program is an example of an application software."
StrQuestions(3) = "Logical errors in a program can be identified by the compiler."
StrQuestions(4) = "Buffering is necessary because the speed of some printers is very high compared with the processing speed of the computer."
StrQuestions(5) = "Aseembly language is called low level language because the CPU directly understands it."
StrQuestions(6) = "The utility service programs control the overall management of a computer system."
StrQuestions(7) = "Text editor is one type of utility program."
StrQuestions(8) = "Translations through an interpreter is faster than a compiler."
StrQuestions(9) = "Pascal supports dynamic memory management."
StrQuestions(10) = "Pseducode should be written in a speciic language."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper19()
StrQuestions(1) = "Multi programming and multitasking are synonymous of each other."
StrQuestions(2) = "Througput is defined as the time required to execute a user's job by the computer system."
StrQuestions(3) = "Job - Control Language statements are instuctions to the Operating System."
StrQuestions(4) = "The machine languages of all computers are identical."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
End Sub


Private Sub multiplechoice_Paper1()
strmultiple(1) = "Address of the first cell in the worksheet is :"
strmuloption(1) = "1"
strmuloption(2) = "A"
strmuloption(3) = "A1"
strmuloption(4) = "1A"

strmultiple(2) = "The address of the cell at 27th column and 30th row will be :"
strmuloption(5) = "27,30"
strmuloption(6) = "AB27"
strmuloption(7) = "AD27"
strmuloption(8) = "AA30"

strmultiple(3) = "The formula incell A2 is = +B2+C3. On copying this formula to cell C2, the formula in cell C2 will be :"
strmuloption(9) = "=+B2+C3"
strmuloption(10) = "=+D2+C3"
strmuloption(11) = "=+D2+E3"
strmuloption(12) = "=+B2+E3"

strmultiple(4) = "The formula in cell A1 is =+$B$3*6. On copying this foumula to cell A6, the formula in cell A6 will be :"
strmuloption(13) = "=-$B$3*6     "
strmuloption(14) = "=+$B$9*6"
strmuloption(15) = "=+$H$9*6"
strmuloption(16) = "=+$B$3*36"

strmultiple(5) = "In a blank workwheet, if you enter profit and loss account in cell C1, then :"
strmuloption(17) = "Only Profit will be displayed"
strmuloption(18) = "Error message will appear"
strmuloption(19) = "Profit and Loss Account will be displayed"
strmuloption(20) = "None of the above"

strmultiple(6) = "Which one of the following can be entered as a number in a cell ?"
strmuloption(21) = "1,300.00"
strmuloption(22) = "1.1e+2"
strmuloption(23) = "(5000.00)"
strmuloption(24) = "All of the above"

strmultiple(7) = "Which of the following cannot be used to enter a date?"
strmuloption(25) = "3/4/97"
strmuloption(26) = "4-Mar-97"
strmuloption(27) = "97/04/12"
strmuloption(28) = "Mar-97"

strmultiple(8) = "Which of the following is not a financial function?"
strmuloption(29) = "FV()"
strmuloption(30) = "NPV()"
strmuloption(31) = "SUM()"
strmuloption(32) = "PMT()"

strmultiple(9) = "Which of the following is not a math function?"
strmuloption(33) = "ABS()"
strmuloption(34) = "COUNT()"
strmuloption(35) = "SIGN()"
strmuloption(36) = "INT()"

strmultiple(10) = "Which of the following is not a statisticalfunction ?"
strmuloption(37) = "MIN()"
strmuloption(38) = "COUNT()"
strmuloption(39) = "PI()"
strmuloption(40) = "All of the above"

strmulanswer(1) = "A1"
strmulanswer(2) = "AA30"
strmulanswer(3) = "=+D2+E3"
strmulanswer(4) = "=+$B$3*6"
strmulanswer(5) = "Profit and Loss Account will be displayed"
strmulanswer(6) = "All of the above"
strmulanswer(7) = "97/04/12"
strmulanswer(8) = "SUM()"
strmulanswer(9) = "COUNT()"
strmulanswer(10) = "PI()"
End Sub

Private Sub multiplechoice_Paper2()
strmultiple(1) = "Which of the following is not a database function?"
strmuloption(1) = "DCOUNT()"
strmuloption(2) = "DMIN()"
strmuloption(3) = "DISC()"
strmuloption(4) = "DAVERAGE()"

strmultiple(2) = "Which of the following is not a text function?"
strmuloption(5) = "CHAR()"
strmuloption(6) = "LEN()"
strmuloption(7) = "VVAR()"
strmuloption(8) = "ABS()"

strmultiple(3) = "Which of thefollowing is not a date function ?"
strmuloption(9) = "DAY()"
strmuloption(10) = "DAYS360()"
strmuloption(11) = "TODAY()"
strmuloption(12) = "LEFT()"

strmultiple(4) = "Mark the most appropriate statement."
strmuloption(13) = "RAND() returns a random number."
strmuloption(14) = "RAND() returns a random number between 0 and 10."
strmuloption(15) = "RAND() returns a random number between 0 and 1."
strmuloption(16) = "RAND() returns a random number between -1 and 1."

strmultiple(5) = "Mark the most appropriate statemnt."
strmuloption(17) = "PV() is used to calculate the parent value."
strmuloption(18) = "PV() is used to calculate the present value of an investment."
strmuloption(19) = "PV() is a financial function."
strmuloption(20) = "None of the above."

strmultiple(6) = "Mark the most appropriate statment."
strmuloption(21) = "NOW() can be used to diaplay the current time."
strmuloption(22) = "NOW() can be used to diaplay the current date."
strmuloption(23) = "NOW() can be used to diaplay the current date as  well as the curent time in the same cell."
strmuloption(24) = "None of the above."

strmultiple(7) = "MOD(-3,2) entered in a cell displays."
strmuloption(25) = "-1"
strmuloption(26) = "1"
strmuloption(27) = "-1.5"
strmuloption(28) = "0"

strmultiple(8) = "The maximum number of arguments that can be used with MIN() are :"
strmuloption(29) = "2        "
strmuloption(30) = "3"
strmuloption(31) = "30"
strmuloption(32) = "No limit"

strmultiple(9) = "The addressing mode(s) that can be used in a formula is/ are :"
strmuloption(33) = "Only absolute"
strmuloption(34) = "Only absulute andrelative"
strmuloption(35) = "Absolute, relative and mixed"
strmuloption(36) = "Not applicable"

strmultiple(10) = "The functon ROUND()1345.45,-2) returns:"
strmuloption(37) = "1343"
strmuloption(38) = "-1343"
strmuloption(39) = "-1300"
strmuloption(40) = "1300"

strmulanswer(1) = "DISC()"
strmulanswer(2) = "ABS()"
strmulanswer(3) = "LEFT()"
strmulanswer(4) = "RAND() returns a random number between 0 and 10."
strmulanswer(5) = "PV() is used to calculate the present value of an investment."
strmulanswer(6) = "NOW() can be used to diaplay the current date as  well as the curent time in the same cell."
strmulanswer(7) = "-1"
strmulanswer(8) = "30"
strmulanswer(9) = "Absolute, relative and mixed"
strmulanswer(10) = "1300"
End Sub

Private Sub multiplechoice_Paper3()
strmultiple(1) = "Which one of the following isinvalid?"
strmuloption(1) = "=SQRT(SUM(C1:C5))"
strmuloption(2) = "=SQRT(ROUND(A3,-3)"
strmuloption(3) = "=SQRT(36)"
strmuloption(4) = "None of the above"

strmultiple(2) = "Assuming that cell A1 contains 1456.5823, which one of the following is incorrect?"
strmuloption(5) = "=ROUND(A1,0)"
strmuloption(6) = "=ROUND(A1,2)"
strmuloption(7) = "=ROUND(A1,-2)"
strmuloption(8) = "None of the above"

strmultiple(3) = "Which one of the following statements is incorrect?"
strmuloption(9) = "A macro can be used to erase a range."
strmuloption(10) = " A macro can be used to put the current time im a cell."
strmuloption(11) = " A macro can be used with more than one worksheet."
strmuloption(12) = "None of the above"

strmultiple(4) = "If PROFIT and SAMPLE are names of ranges in a Worksheet, which one of the following is invalid?"
strmuloption(13) = "@SUM(A1:G3,25,27)"
strmuloption(14) = "@SUM(PROFIT,12,15)"
strmuloption(15) = "=SUM(PROFIT,$SAMPLE,24)"
strmuloption(16) = "None of the above"

strmultiple(5) = "In a new WOrksheet, cell A1 contains 5, cell A2 contains 7 and cell A4 contains 9. If cell B1 contains = COUNT(A1:A4), the following result will be displayed in cell B1:"
strmuloption(17) = "9"
strmuloption(18) = "3"
strmuloption(19) = "4"
strmuloption(20) = "5"

strmulanswer(1) = "None of the above"
strmulanswer(2) = "None of the above"
strmulanswer(3) = "None of the above"
strmulanswer(4) = "3"
End Sub

Private Sub multiplechoice_Paper4()
strmultiple(1) = "WINDOWS is popular becase of its -"
strmuloption(1) = "GUI features"
strmuloption(2) = "Multitasking capacity"
strmuloption(3) = "Desktop Technology"
strmuloption(4) = "Being inexpensive"

strmultiple(2) = "WINDOWS IS -"
strmuloption(5) = "an Operating system"
strmuloption(6) = "a Operating environment"
strmuloption(7) = "a Character user interface"
strmuloption(8) = "None of the above"

strmultiple(3) = "Paint brush belongs to -"
strmuloption(9) = "Main group"
strmuloption(10) = "Application group"
strmuloption(11) = "Accessories group"
strmuloption(12) = "Startup group"

strmultiple(4) = "WINDOWS  settings are recorded in -"
strmuloption(13) = "COMMAND.INI"
strmuloption(14) = "WIN.INI"
strmuloption(15) = "WINDOWS.INI"
strmuloption(16) = "GROUPS.INI"

strmultiple(5) = "All systems settings in WINDOWS are stored in -"
strmuloption(17) = "CONTROL.INI"
strmuloption(18) = "MAIN.INI"
strmuloption(19) = "SYSTEM.INI"
strmuloption(20) = "SETTING.INI"

strmultiple(6) = "PIF editor belongs to which group?"
strmuloption(21) = "Main"
strmuloption(22) = "Accessories"
strmuloption(23) = "Games"
strmuloption(24) = "None of the above"

strmultiple(7) = "WINDOWS can work in which of the following modes?"
strmuloption(25) = "Extended"
strmuloption(26) = "Expanded"
strmuloption(27) = "Enhanced & Standard"
strmuloption(28) = "Compressed"

strmultiple(8) = "Through a PIF file one can -"
strmuloption(29) = "Load a DOS based application"
strmuloption(30) = "Load a WINDOWS based "
strmuloption(31) = "Load a CONFIG.SYS file"
strmuloption(32) = "None of the above"

strmultiple(9) = "-default.pif is found in -"
strmuloption(33) = "WINDOWS directory"
strmuloption(34) = "DOS directory"
strmuloption(35) = "Root directory"
strmuloption(36) = "None of the above"

strmultiple(10) = "Filename with extension as .VXD represents-"
strmuloption(37) = "Normal application file"
strmuloption(38) = "Text file"
strmuloption(39) = "Device drivers"
strmuloption(40) = "Picture file"

strmulanswer(1) = "GUI features"
strmulanswer(2) = "a Operating environment"
strmulanswer(3) = "Accessories group"
strmulanswer(4) = "WIN.INI"
strmulanswer(5) = "SYSTEM.INI"
strmulanswer(6) = "Main"
strmulanswer(7) = "Enhanced & Standard"
strmulanswer(8) = "Load a DOS based application"
strmulanswer(9) = "WINDOWS directory"
strmulanswer(10) = "Device drivers"
End Sub

Private Sub multiplechoice_Paper5()
strmultiple(1) = "Which one does not belong to the Main group?"
strmuloption(1) = "Control Panel"
strmuloption(2) = "File Manager"
strmuloption(3) = "Character Map"
strmuloption(4) = "PIF Manager"

strmultiple(2) = "PIF editor is a -"
strmuloption(5) = "DOS based application"
strmuloption(6) = "WINDOWS based application"
strmuloption(7) = "Device driver"
strmuloption(8) = "NOne of the above."

strmultiple(3) = "A compound document can receeve objects from -"
strmuloption(9) = "Only one server application"
strmuloption(10) = "At most two server applications"
strmuloption(11) = "Any number of server applications"
strmuloption(12) = "None of the above"

strmultiple(4) = "The command startup= MYGROUP, will make the applications in the MYGROUP to run as soon as WINDOWS is started. Where should tis command be written?"
strmuloption(13) = "WIN.INI"
strmuloption(14) = "SYSTEM.INI"
strmuloption(15) = "PROGRAM.INI"
strmuloption(16) = "MYGROUP.INI"

strmultiple(5) = "WIN/2 starts WINDOWS in -"
strmuloption(17) = "Enhanced mode "
strmuloption(18) = "Standard mode"
strmuloption(19) = "Invalid command"
strmuloption(20) = "None of the above"

strmultiple(6) = "RTF stands for"
strmuloption(21) = "Real Time Fonts"
strmuloption(22) = "Real Time FIles"
strmuloption(23) = "Rich Text Format"
strmuloption(24) = "Rich Text Fonts"

strmultiple(7) = "Combo Box, in a dialog box, is -"
strmuloption(25) = "A combination of list box and check box"
strmuloption(26) = "A list box"
strmuloption(27) = "An option button"
strmuloption(28) = "A combination o text box a list box"

strmulanswer(1) = "Character Map"
strmulanswer(2) = "WINDOWS based application"
strmulanswer(3) = "Any number of server applications"
strmulanswer(4) = "PROGRAM.INI"
strmulanswer(5) = "Standard mode"
strmulanswer(6) = "Rich Text Format"
strmulanswer(7) = "A combination o text box a list box"
End Sub

Private Sub multiplechoice_Paper6()
strmultiple(1) = "Which of the following is incorrect ?"
strmuloption(1) = "A paragraph can be formatted such that it never breaks into two pages."
strmuloption(2) = "Two patagraphs can be formatted in such a way that these are always on the same page."
strmuloption(3) = "A paragraph can be formatted in such a way that it is always at the top of the page."
strmuloption(4) = "None of these"

strmultiple(2) = "Mark the most appropriate statement."
strmuloption(5) = "A paragraph can be enclosed in a box."
strmuloption(6) = "A paragraph can be enclosed in a box and the background shade of the box can also be controlled."
strmuloption(7) = "A paragraph can be enclosed in a shadow box."
strmuloption(8) = "All of the above"

strmultiple(3) = "Mark the most appropriate statement."
strmuloption(9) = "MailMerge cab ne used to print the form letters."
strmuloption(10) = "MailMerge cab ne used to print the form letters a d mailing labels."
strmuloption(11) = "MailMerge cab ne used to print the form letters and mailing envelops."
strmuloption(12) = "All of the above."

strmultiple(4) = "Creating form letters using MailMErge involves"
strmuloption(13) = "Creating the main document"
strmuloption(14) = "Creating the main document and data source"
strmuloption(15) = "Inserting merge fields in th main document and merging he main document with the data source."
strmuloption(16) = "Both (b) and (c)"

strmultiple(5) = "Graphics is inseted in -"
strmuloption(17) = "Frame"
strmuloption(18) = "Box"
strmuloption(19) = "Page"
strmuloption(20) = "None of the above"

strmultiple(6) = "Frame can include"
strmuloption(21) = "Graphics"
strmuloption(22) = "Tables and Graphics"
strmuloption(23) = "Text and Graphics"
strmuloption(24) = "All of the above"

strmultiple(7) = "Select the most appropriate statement:"
strmuloption(25) = "Frame can be resized"
strmuloption(26) = " Frame can include any type of text ad border and shade can be added to frame."
strmuloption(27) = "(b) and (c)"
strmuloption(28) = "Frame is automatically repeated on every page."

strmultiple(8) = "All formatting for the paragraph is stored -"
strmuloption(29) = "In the page setup"
strmuloption(30) = "In the paragraph mark."
strmuloption(31) = "In the style sheet"
strmuloption(32) = "None of the above"

strmultiple(9) = "To paste text from the clipboard in the document being edited, press the key(s) :"
strmuloption(33) = "Ctrl X   "
strmuloption(34) = "Ctrl V"
strmuloption(35) = "Ctrl A"
strmuloption(36) = "Ins"

strmultiple(10) = "To select a sentence, click anywhere on the wentence while holding the following key depressed -"
strmuloption(37) = "Shift"
strmuloption(38) = "Ctrl"
strmuloption(39) = "Alt"
strmuloption(40) = "Esc"

strmulanswer(1) = "None of these"
strmulanswer(2) = "All of the above"
strmulanswer(3) = "All of the above"
strmulanswer(4) = "Both (b) and (c)"
strmulanswer(5) = "Frame"
strmulanswer(6) = "All of the above"
strmulanswer(7) = "(b) and (c)"
strmulanswer(8) = "In the paragraph mark."
strmulanswer(9) = "Ctrl V"
strmulanswer(10) = "Ctrl"
End Sub

Private Sub multiplechoice_Paper7()
strmultiple(1) = "To delete the selected sentence we can press the following key:"
strmuloption(1) = "Del"
strmuloption(2) = "Backspace"
strmuloption(3) = "Both (a) & (b)"
strmuloption(4) = "None of the above"

strmultiple(2) = "Select the most appropriate statement."
strmuloption(5) = "Auto text is the standard feature of all word processors."
strmuloption(6) = "AutoText automatically inserts the text and/or graphics on typing the AutoText entry and pressing the spacebar."
strmuloption(7) = "AutoText is used to insert commonly used text and graphics in a document."
strmuloption(8) = "None of the above."

strmultiple(3) = "The document can be zoomed up to:"
strmuloption(9) = "100%"
strmuloption(10) = "150%"
strmuloption(11) = "200%"
strmuloption(12) = "300%"

strmultiple(4) = "Select the most appropriate statement."
strmuloption(13) = "AutoText and AutoCorrect are similar"
strmuloption(14) = "AutoCorrect automatically inserts the text and/or graphics on typing the AutoCorrect entry."
strmuloption(15) = "AutoCorrect automatically inserts the text and/or graphics on typing the AutoCorect entry and pressing spacebar."
strmuloption(16) = "None of the above."

strmultiple(5) = "Which one of the following statements is true."
strmuloption(17) = "A Word Document can have a maximum of 4 columns."
strmuloption(18) = "Multiple columns cannot be viewed in the normal view mode."
strmuloption(19) = "In case of multiple columns in a document, all columns are of the same width."
strmuloption(20) = "None of the above"

strmultiple(6) = "Mark the incorrect statement."
strmuloption(21) = "Hyphenation can be done automatically or manually."
strmuloption(22) = "The hyphenation facility is available in the Insert menu."
strmuloption(23) = "Hyphenation can affect page breaks."
strmuloption(24) = "None of the above."

strmultiple(7) = "Which one of the following statements is true."
strmuloption(25) = "To copy the selected text with mouse, hold down the shift key and drag the insertion point to the new location."
strmuloption(26) = "If you delete a paparagraph mark, Word applies the fromatting of the previous patagraph to the current patagraph."
strmuloption(27) = "Selected text can be moved to a new location in the document by dragging."
strmuloption(28) = "None of the above."

strmultiple(8) = "Which one of the following statemens is true?"
strmuloption(29) = "Line spacing can be set to 1.5"
strmuloption(30) = "A specified spacing can be left before and after any paragraph."
strmuloption(31) = "Both (a) and (b)"
strmuloption(32) = "None of the above."

strmultiple(9) = "Which one of the following statements is false?"
strmuloption(33) = "A specified spacing can be left after a paragraph."
strmuloption(34) = "A paragraph can use the linespacing of 3.5."
strmuloption(35) = "A border can be drawn around a paragraph. However, two paragraphs cannot be included with in a single box."
strmuloption(36) = "None of the above."

strmultiple(10) = "After a table has been created, which of the following operations cannot be performed?"
strmuloption(37) = "Insert rows in a table"
strmuloption(38) = "Delete and insert columns"
strmuloption(39) = "Split the table into two tables"
strmuloption(40) = "None of the above."

strmulanswer(1) = "Both (a) & (b)"
strmulanswer(2) = "AutoText is used to insert commonly used text and graphics in a document."
strmulanswer(3) = "200%"
strmulanswer(4) = "AutoCorrect automatically inserts the text and/or graphics on typing the AutoCorect entry and pressing spacebar."
strmulanswer(5) = "Multiple columns cannot be viewed in the normal view mode."
strmulanswer(6) = "The hyphenation facility is available in the Insert menu."
strmulanswer(7) = "Selected text can be moved to a new location in the document by dragging."
strmulanswer(8) = "Both (a) and (b)"
strmulanswer(9) = "A border can be drawn around a paragraph. However, two paragraphs cannot be included with in a single box."
strmulanswer(10) = "None of the above."
End Sub

Private Sub multiplechoice_Paper8()
strmultiple(1) = "A sequence of instructions, in a computer language, to get the desired result is known as"
strmuloption(1) = "an algorithm"
strmuloption(2) = "a decision table"
strmuloption(3) = "a program"
strmuloption(4) = "none of the above"

strmultiple(2) = "The first commercially available computer was the "
strmuloption(5) = "ENIAC"
strmuloption(6) = "ABACUS"
strmuloption(7) = "UNIAC - I"
strmuloption(8) = "MARK - I"

strmultiple(3) = "Which type of errors are flagged by compilers?"
strmuloption(9) = "logical errors"
strmuloption(10) = "run time errors"
strmuloption(11) = "syntax errors"
strmuloption(12) = "none of the above"

strmultiple(4) = "An algorithm is best described as"
strmuloption(13) = "a computer language"
strmuloption(14) = "a step by step procedure for solving a problem."
strmuloption(15) = "aa branch of mathematics"
strmuloption(16) = "none of the above"

strmultiple(5) = "The abbreviation CAD stands for"
strmuloption(17) = "Computr And Data"
strmuloption(18) = "Computer Aided Design"
strmuloption(19) = "Commonly Available Data"
strmuloption(20) = "Computer Aided Drawing"

strmultiple(6) = "Bug means"
strmuloption(21) = "A logical error in a program."
strmuloption(22) = "A difficulat syntax error in a program."
strmuloption(23) = "Documenting programs using an efficient documentation tool."
strmuloption(24) = "None of these"

strmultiple(7) = "The part of a machine level instruction, which tells the central processor what has to be done, is"
strmuloption(25) = "an operation code"
strmuloption(26) = "an address"
strmuloption(27) = "an operand"
strmuloption(28) = "none of the above"

strmultiple(8) = "Indicate which of the following is not true about 4GL?"
strmuloption(29) = "4GL do not support a high level screen interaction"
strmuloption(30) = "many data base management system pachages support 4GL's."
strmuloption(31) = "a 4GL isa software tool which is written, possibly in some third generation language."
strmuloption(32) = "none of the above"

strmultiple(9) = "Indicate which of the follwing best describes the term SOFTWARE"
strmuloption(33) = "Operating System programs only"
strmuloption(34) = "Application programs only"
strmuloption(35) = "Both (a) and (b)"
strmuloption(36) = "none of the above"

strmultiple(10) = "A language translator is best described  as"
strmuloption(37) = "an application software"
strmuloption(38) = "a system software"
strmuloption(39) = "a hardware component"
strmuloption(40) = "none of the above"

strmulanswer(1) = "a program"
strmulanswer(2) = "UNIAC - I"
strmulanswer(3) = "syntax errors"
strmulanswer(4) = "a step by step procedure for solving a problem."
strmulanswer(5) = "Computer Aided Design"
strmulanswer(6) = "A logical error in a program."
strmulanswer(7) = "an operation code"
strmulanswer(8) = "4GL do not support a high level screen interaction"
strmulanswer(9) = "Both (a) and (b)"
strmulanswer(10) = "a system software"
End Sub

Private Sub multiplechoice_Paper9()
strmultiple(1) = "Indicate which of the following in not true about an interpreter"
strmuloption(1) = "Intepretor generates an object program from the source program"
strmuloption(2) = "Interpretor is a kind oftranslator."
strmuloption(3) = "Interpretor analysis each source statement every time it is to be executed."
strmuloption(4) = "None of the above"

strmultiple(2) = "The errors that can be pointed out by the compiler are"
strmuloption(5) = "syntax errors"
strmuloption(6) = "semantic errors"
strmuloption(7) = "logical errors"
strmuloption(8) = "None of the above"

strmultiple(3) = "Which of the following statements is not true?"
strmuloption(9) = "FORTRAN is a high level language."
strmuloption(10) = "A FORTRAN program written for the IBM- PC is totally different from a FORTRAN program written for the execution of SUN machine."
strmuloption(11) = "FORTRAN is extensively used to write programs for performing scientific computations."
strmuloption(12) = "None of the above"

strmultiple(4) = "'C' is"
strmuloption(13) = "an assembly language"
strmuloption(14) = "a 3GL"
strmuloption(15) = "a machine language"
strmuloption(16) = "none of the above"

strmultiple(5) = "A program that converts a high level language program to a set of instuctions that can run on a computer is called as"
strmuloption(17) = "compiler"
strmuloption(18) = "debugger"
strmuloption(19) = "editor."
strmuloption(20) = "none of the above"

strmultiple(6) = "The function of a kernel in an Operating System is to"
strmuloption(21) = "interact with the user"
strmuloption(22) = "interact with the hardware devices"
strmuloption(23) = "interpret the JCL commands"
strmuloption(24) = "None of the above"

strmultiple(7) = "Indicate which of the following is necessar to work on a compter"
strmuloption(25) = "compiler"
strmuloption(26) = "Operating System"
strmuloption(27) = "assembler"
strmuloption(28) = "none of the above"

strmultiple(8) = "Indicate which of the following is true about PSEDUCODE?"
strmuloption(29) = "is a machine language"
strmuloption(30) = "is an assembly language"
strmuloption(31) = "is a high level language"
strmuloption(32) = "is none of the above"

strmultiple(9) = "The language used to describe the user's requrements to a computer's Operating System is"
strmuloption(33) = "JCL"
strmuloption(34) = "COBOL"
strmuloption(35) = "Assembly language"
strmuloption(36) = "None of the above"

strmultiple(10) = "Indicate which of the following is not true ?"
strmuloption(37) = "In general. 4Gl's blong tothe category of non procedural languages."
strmuloption(38) = "4GLs do not support a high level of screen interaction."
strmuloption(39) = "A 4GL is a software that is itself written in some other computer language."
strmuloption(40) = "Many 4GLs work on a database"

strmulanswer(1) = "Intepretor generates an object program from the source program"
strmulanswer(2) = "syntax errors"
strmulanswer(3) = "A FORTRAN program written for the IBM- PC is totally different from a FORTRAN program written for the execution of SUN machine."
strmulanswer(4) = "a 3GL"
strmulanswer(5) = "compiler"
strmulanswer(6) = "interact with the hardware devices"
strmulanswer(7) = "Operating System"
strmulanswer(8) = "is none of the above"
strmulanswer(9) = "COBOL"
strmulanswer(10) = "4GLs do not support a high level of screen interaction."
End Sub

Private Sub multiplechoice_Paper10()
strmultiple(1) = "JCL stands for"
strmuloption(1) = "Jump Control Logic"
strmuloption(2) = "Just Command Language"
strmuloption(3) = "Job Control Language"
strmuloption(4) = "None of the above"

strmultiple(2) = "The abbreviation EFT stands for"
strmuloption(5) = "Electronic forms Typing"
strmuloption(6) = "Electronic File Transfer"
strmuloption(7) = "Electronic Funds Transfer"
strmuloption(8) = "Electronin Flow Table"

strmultiple(3) = "Dot Matrix printer speed is specified in trms of"
strmuloption(9) = "BPS"
strmuloption(10) = "CPS"
strmuloption(11) = "BPI"
strmuloption(12) = "CPI"

strmultiple(4) = "Laser printer resolution is specified in terms of "
strmuloption(13) = "CPI"
strmuloption(14) = "API"
strmuloption(15) = "DPI"
strmuloption(16) = "LSI"

strmultiple(5) = "The stored Program concept was proposed by"
strmuloption(17) = "Abacus"
strmuloption(18) = "John Von Neumann"
strmuloption(19) = "Pascal"
strmuloption(20) = "Von Leibniz"

strmultiple(6) = "Which of the following allws users to continue to operate computers while printing is in progress?"
strmuloption(21) = "Spooler"
strmuloption(22) = "OS Supervisor"
strmuloption(23) = "Job Control Program"
strmuloption(24) = "None of the above"

strmultiple(7) = "Choose the 'Odd One Out' of the following"
strmuloption(25) = "COBOL"
strmuloption(26) = "FORTRAN"
strmuloption(27) = "PASCAL"
strmuloption(28) = "ASSEMBLER"

strmultiple(8) = "A computer cannot 'Boot' if it does not have the"
strmuloption(29) = "compiler"
strmuloption(30) = "loader"
strmuloption(31) = "Operating system"
strmuloption(32) = "assembler"

strmultiple(9) = "Indicate which, of the following is a phase of a compilation process?"
strmuloption(33) = "Lexical"
strmuloption(34) = "Code generation"
strmuloption(35) = "Both of the above"
strmuloption(36) = "None of the above"

strmultiple(10) = "Connecting different computers in an organized manner, within an office can be termed as"
strmuloption(37) = "ANN"
strmuloption(38) = "WAN"
strmuloption(39) = "MAN"
strmuloption(40) = "LAN"

strmulanswer(1) = "Job Control Language"
strmulanswer(2) = "Electronic Funds Transfer"
strmulanswer(3) = "CPS"
strmulanswer(4) = "DPI"
strmulanswer(5) = "John Von Neumann"
strmulanswer(6) = "Spooler"
strmulanswer(7) = "ASSEMBLER"
strmulanswer(8) = "Operating system"
strmulanswer(9) = "Both of the above"
strmulanswer(10) = "LAN"
End Sub

Private Sub multiplechoice_Paper11()
strmultiple(1) = "Which one of the following is not an input device?"
strmuloption(1) = "Microphone"
strmuloption(2) = "Mouse"
strmuloption(3) = "Scanner"
strmuloption(4) = "None of the above"

strmultiple(2) = "Operating System is"
strmuloption(5) = "a collection of hardware components"
strmuloption(6) = "a collection of input-output devices."
strmuloption(7) = "a collection of software routines"
strmuloption(8) = "None of the above"

strmultiple(3) = "Operating System"
strmuloption(9) = "Liks a progrm with the sunroutines it refers"
strmuloption(10) = "Provides a layered, user friendly, interface"
strmuloption(11) = "Enables theprogrammer to draw a flowchart."
strmuloption(12) = "None of the above"

strmultiple(4) = "A characteristics of an on line, real time system is"
strmuloption(13) = "more than one CPU"
strmuloption(14) = "No delay in processing"
strmuloption(15) = "Off line batch processing"
strmuloption(16) = "None of the above"

strmultiple(5) = "Line printerspeed is specified in terms of"
strmuloption(17) = "COM"
strmuloption(18) = "CPM"
strmuloption(19) = "DPM"
strmuloption(20) = "LPM"

strmultiple(6) = "Job control sttements are used to"
strmuloption(21) = "read the input from the slow speed card reader to the high speed magnetic disk."
strmuloption(22) = "specify to the Operating System the beginning and end of a hob in a batch"
strmuloption(23) = "allocate the CPU to a job"
strmuloption(24) = "none of the above"

strmultiple(7) = "Optical storage technique is used in"
strmuloption(25) = "Floppy drive"
strmuloption(26) = "Hard disk drive"
strmuloption(27) = "Magnetic disk drive"
strmuloption(28) = "CD ROM drive"

strmultiple(8) = "The technique, which stores a program on disk and then transfers the program into Main memory as and when they are needed is known as"
strmuloption(29) = "Spooling"
strmuloption(30) = "Swapping"
strmuloption(31) = "thrashing"
strmuloption(32) = "None of the above"

strmultiple(9) = "Batch processing means"
strmuloption(33) = "a batch of computer programs"
strmuloption(34) = "entering a batch ofrecords"
strmuloption(35) = "processing data in periodic intervals oftime"
strmuloption(36) = "none of these"

strmultiple(10) = "Education through computers is refereed to as"
strmuloption(37) = "CAD"
strmuloption(38) = "CAM"
strmuloption(39) = "CAI"
strmuloption(40) = "None of these"

strmulanswer(1) = "None of the above"
strmulanswer(2) = "a collection of software routines"
strmulanswer(3) = "Provides a layered, user friendly, interface"
strmulanswer(4) = "No delay in processing"
strmulanswer(5) = "LPM"
strmulanswer(6) = "specify to the Operating System the beginning and end of a hob in a batch"
strmulanswer(7) = "CD ROM drive"
strmulanswer(8) = "Swapping"
strmulanswer(9) = "a batch of computer programs"
strmulanswer(10) = "CAI"
End Sub

Private Sub multiplechoice_Paper12()
strmultiple(1) = "A database is"
strmuloption(1) = "A set of condensed Data highilghtins specific information"
strmuloption(2) = "A vast amount odata stored in a grou o integer Files"
strmuloption(3) = "An arrrangement of data in a particluar order"
strmuloption(4) = "Basic data for processing an application"

strmultiple(2) = "The name given to a sequence of instructions in a computer language to get the desired result is"
strmuloption(5) = "a program"
strmuloption(6) = "an algorithm"
strmuloption(7) = "a pseudocode"
strmuloption(8) = "a decision table"

strmultiple(3) = "Thelanguage understood by a computer without translation is called"
strmuloption(9) = "command language"
strmuloption(10) = "high level language"
strmuloption(11) = "assembly language"
strmuloption(12) = "machine language"

strmultiple(4) = "Program to convert a hih level compuer program into a set ofinstructions that will rin on a machine is called a"
strmuloption(13) = "compiler"
strmuloption(14) = "debugger"
strmuloption(15) = "linker"
strmuloption(16) = "editor"

strmultiple(5) = "1 Kbits is"
strmuloption(17) = "1000 bits"
strmuloption(18) = "1024 bits"
strmuloption(19) = "512 bits"
strmuloption(20) = "none of these"

strmultiple(6) = "A sequence o steps writen in natural language to solve a given problem is called"
strmuloption(21) = "an algorithm"
strmuloption(22) = "a program"
strmuloption(23) = "binary code"
strmuloption(24) = "none ofthe above"

strmultiple(7) = "Which of the following describes the functions ofa program?"
strmuloption(25) = "Specifying the locations of variables in the memory "
strmuloption(26) = "Specifying constants"
strmuloption(27) = "Specifying the various operations to be performed by the computer"
strmuloption(28) = "All of the above"

strmultiple(8) = "Which one of the following is not an Operating Systems?"
strmuloption(29) = "MS-DOS"
strmuloption(30) = "PC-DOS"
strmuloption(31) = "CP/M"
strmuloption(32) = "Z-80"

strmultiple(9) = "The mapping from assembly language instructions into machine language instructions is"
strmuloption(33) = "many - one"
strmuloption(34) = "one - many"
strmuloption(35) = "one - one"
strmuloption(36) = "many - many"

strmultiple(10) = " A compiler is"
strmuloption(37) = "machineindependent and OS dependent"
strmuloption(38) = "machine dependent and OS dependent"
strmuloption(39) = "machine dependent and OS independent"
strmuloption(40) = "machine independent and Os dependent"

strmulanswer(1) = "An arrrangement of data in a particluar order"
strmulanswer(2) = "a program"
strmulanswer(3) = "machine language"
strmulanswer(4) = "compiler"
strmulanswer(5) = "1024 bits"
strmulanswer(6) = "an algorithm"
strmulanswer(7) = "Specifying the various operations to b performed by the computer"
strmulanswer(8) = "Z-80"
strmulanswer(9) = "one - one"
strmulanswer(10) = "machine dependent and OS independent"
End Sub


