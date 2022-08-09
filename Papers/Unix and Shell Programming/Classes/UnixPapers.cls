VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnixPapers"
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
 
Public Function Unix(PaperType As String, UnixPaper As Paper, mCat As Category, Number As Integer)
'ptype is the paper type
'Unixpaper is the number of paper as paper1,paper2 ...
'mcat is the category
'number is the number of questions, or options or answers to be taken

StrPaperType = PaperType
' StrPaper = Paper
If PaperType = "Fill in the Blanks" Then
   If UnixPaper = Paper1 Then Fill_in_the_Blanks_Paper1
   If UnixPaper = Paper2 Then Fill_in_the_Blanks_Paper2
   If UnixPaper = Paper3 Then Fill_in_the_Blanks_Paper3
   If UnixPaper = Paper4 Then Fill_in_the_Blanks_Paper4
   If UnixPaper = Paper5 Then Fill_in_the_Blanks_Paper5
   If UnixPaper = Paper6 Then Fill_in_the_Blanks_Paper6
   If UnixPaper = Paper7 Then Fill_in_the_Blanks_Paper7
   If UnixPaper = Paper8 Then Fill_in_the_Blanks_Paper8
   If UnixPaper = Paper9 Then Fill_in_the_Blanks_Paper9
   If UnixPaper = Paper10 Then Fill_in_the_Blanks_Paper10
   If UnixPaper = Paper11 Then Fill_in_the_Blanks_Paper11
   
ElseIf PaperType = "Matching Columns" Then
       If UnixPaper = Paper1 Then Matching_Columns_Paper1
       If UnixPaper = Paper2 Then Matching_Columns_Paper2
       If UnixPaper = Paper3 Then Matching_Columns_Paper3
       If UnixPaper = Paper4 Then Matching_Columns_Paper4
       If UnixPaper = Paper5 Then Matching_Columns_Paper5
       If UnixPaper = Paper6 Then Matching_Columns_Paper6
       If UnixPaper = Paper7 Then Matching_Columns_Paper7
       If UnixPaper = Paper8 Then Matching_Columns_Paper8
       If UnixPaper = Paper9 Then Matching_Columns_Paper9
       If UnixPaper = Paper10 Then Matching_Columns_Paper10
       
ElseIf PaperType = "True Or False" Then
       If UnixPaper = Paper1 Then TrueorFalse_Paper1
       If UnixPaper = Paper2 Then TrueorFalse_Paper2
       If UnixPaper = Paper3 Then TrueorFalse_Paper3
       If UnixPaper = Paper4 Then TrueorFalse_Paper4
       If UnixPaper = Paper5 Then TrueorFalse_Paper5
       If UnixPaper = Paper6 Then TrueorFalse_Paper6
       If UnixPaper = Paper7 Then TrueorFalse_Paper7
       If UnixPaper = Paper8 Then TrueorFalse_Paper8
       If UnixPaper = Paper9 Then TrueorFalse_Paper9
       If UnixPaper = Paper10 Then TrueorFalse_Paper10
       
ElseIf PaperType = "Descriptive Questions" Then
       If UnixPaper = Paper1 Then questions_paper1
       If UnixPaper = Paper2 Then questions_paper2
       If UnixPaper = Paper3 Then questions_paper3
       If UnixPaper = Paper4 Then questions_paper4
       If UnixPaper = Paper5 Then questions_paper5
       If UnixPaper = Paper6 Then questions_paper6
       
ElseIf PaperType = "Multiple Choice" Then
       If UnixPaper = Paper1 Then multiplechoice_Paper1
       If UnixPaper = Paper2 Then multiplechoice_Paper2
       If UnixPaper = Paper3 Then multiplechoice_Paper3
       If UnixPaper = Paper4 Then multiplechoice_Paper4
       If UnixPaper = Paper5 Then multiplechoice_Paper5
       If UnixPaper = Paper6 Then multiplechoice_Paper6
       If UnixPaper = Paper7 Then multiplechoice_Paper7
       If UnixPaper = Paper8 Then multiplechoice_Paper8
       If UnixPaper = Paper9 Then multiplechoice_Paper9
       If UnixPaper = Paper10 Then multiplechoice_Paper10
End If

If mCat = Fill_Questions Then
   Unix = StrQuestions(Number)
ElseIf mCat = Fill_Options Then
       Unix = stroptions(Number)
ElseIf mCat = Fill_Answers Then
       Unix = stranswers(Number)
       
ElseIf mCat = Matching_ColumnA Then
       Unix = strcolumnA(Number)
ElseIf mCat = Matching_ColumnB Then
       Unix = strcolumnB(Number)
ElseIf mCat = Matching_Answers Then
       Unix = stranswers(Number)
       
ElseIf mCat = Questions Then
       Unix = StrQuestions(Number)
       
ElseIf mCat = TF_Questions Then
       Unix = StrQuestions(Number)
ElseIf mCat = TF_Answers Then
       Unix = stranswers(Number)
       
ElseIf mCat = MC_Questions Then
       Unix = strmultiple(Number)
       
ElseIf mCat = MC_Options_1 Then
       Unix = strmuloption(Number)
       
ElseIf mCat = MC_Options_2 Then
       Unix = strmuloption(Number)
       
ElseIf mCat = MC_Options_3 Then
       Unix = strmuloption(Number)
       
ElseIf mCat = MC_Options_4 Then
       Unix = strmuloption(Number)
       
ElseIf mCat = MC_Answer Then
       Unix = strmulanswer(Number)
       
End If

End Function

Private Sub Fill_in_the_Blanks_Paper1()
StrQuestions(1) = "The shell acts as the dash interpreter."
StrQuestions(2) = "The kernel processes have dash priority."
StrQuestions(3) = "The user passwords are kept in the dash file."
StrQuestions(4) = "There are dash standard files in Unix."
StrQuestions(5) = "The file descriptor of stdin is dash ."
StrQuestions(6) = "UNIX has a dash file organization."
StrQuestions(7) = "Most of the UNIX code is written in the high level language dash."
StrQuestions(8) = "The device information can be obtained from the dash directory."
StrQuestions(9) = "The command to remove a directory containing files is dash."
StrQuestions(10) = "Two files can be merged together using the dash command at the command line."

stroptions(1) = "01"
stroptions(2) = "/dev"
stroptions(3) = "cat"
stroptions(4) = "Hierarchical"
stroptions(5) = "/etc/passwd"
stroptions(6) = "C"
stroptions(7) = "Higher"
stroptions(8) = "rm - r"
stroptions(9) = "Three"
stroptions(10) = "Command"

stranswers(1) = "Command"
stranswers(2) = "Higher"
stranswers(3) = "/etc/passwd"
stranswers(4) = "Three"
stranswers(5) = "01"
stranswers(6) = "Hierarchical"
stranswers(7) = "C"
stranswers(8) = "/dev"
stranswers(9) = "rm - r"
stranswers(10) = "cat"

End Sub

Private Sub Fill_in_the_Blanks_Paper2()
StrQuestions(1) = "When a file is created the default access permission is dash."
StrQuestions(2) = "The date is displayed by the dash command."
StrQuestions(3) = "When the date command is executed it runs as a child of the dash process."
StrQuestions(4) = "The PID of a process is displayed by the dash command."
StrQuestions(5) = "The ls -I today command will display the dash number of the file today."
StrQuestions(6) = "The copy of a file can be obtained by the dash command."
StrQuestions(7) = "You can check whether a device is character oriented or block oriented, looking at the first column of the output obtained from the dash /dev command."
StrQuestions(8) = "A file is displayed with line numbers using the dash command."
StrQuestions(9) = "We can redirect the output of a command only if the command is a dash."
StrQuestions(10) = "The default standard error file is dash."

stroptions(1) = "Filter"
stroptions(2) = "Shell"
stroptions(3) = "nu"
stroptions(4) = "Cp"
stroptions(5) = "Ps"
stroptions(6) = "Read write"
stroptions(7) = "stderr"
stroptions(8) = "Ls -l"
stroptions(9) = "inode"
stroptions(10) = "Date"

stranswers(1) = "Read write"
stranswers(2) = "Date"
stranswers(3) = "Shell"
stranswers(4) = "Ps"
stranswers(5) = "inode"
stranswers(6) = "Cp"
stranswers(7) = "Ls -l"
stranswers(8) = "nu"
stranswers(9) = "Filter"
stranswers(10) = "stderr"

End Sub


Private Sub Fill_in_the_Blanks_Paper3()
StrQuestions(1) = "cat this file - will concatentate this file with the inputs from the dash file."
StrQuestions(2) = "The output of one command can be directly fed in to the input of another command using a dash."
StrQuestions(3) = "A shell script can be directly run if it has the dash permission."
StrQuestions(4) = "The command sh creates dash as a child process."
StrQuestions(5) = "The pattern dash will match dog and dig, but no of the word"
StrQuestions(6) = "The user created shell variables must start with a dash."
StrQuestions(7) = "The statement echo 2 + 5  will display dash."
StrQuestions(8) = "PATH is a dash shell variable."
StrQuestions(9) = "The shell variable that defines internal field separator is dash."
StrQuestions(10) = "You can change the shell prompt by reassigning the dash variable."

stroptions(1) = "System"
stroptions(2) = "Alphabet"
stroptions(3) = "A shell"
stroptions(4) = "2+5"
stroptions(5) = "Execute"
stroptions(6) = "PS1"
stroptions(7) = "Pipe"
stroptions(8) = "IFB"
stroptions(9) = "d?g"
stroptions(10) = "Keyboard"

stranswers(1) = "Keyboard"
stranswers(2) = "Pipe"
stranswers(3) = "Execute"
stranswers(4) = "A shell"
stranswers(5) = "d?g"
stranswers(6) = "Alphabet"
stranswers(7) = "2+5"
stranswers(8) = "System"
stranswers(9) = "IFB"
stranswers(10) = "PS1"

End Sub

Private Sub Fill_in_the_Blanks_Paper4()
StrQuestions(1) = "The scope of a shell variable can be made global to the subshells by the dash command."
StrQuestions(2) = "The value of the variable number will be displayed by typing echo dash."
StrQuestions(3) = "When you login in the system, the system automatically looks for dash file in your home directory."
StrQuestions(4) = "The startup instructions for a user’s account can be kept in the dash file."
StrQuestions(5) = "In a shell script $0 stands for the dash"
StrQuestions(6) = "The positional parameter $# stands for the dash."
StrQuestions(7) = "The dash command will tell you the number of characters in a file."
StrQuestions(8) = "Command substation is done by dash."
StrQuestions(9) = "The dash command will tell you only the number of lines in a file."
StrQuestions(10) = "The dash command will tell you only the number of words in a file."

stroptions(1) = "Command"
stroptions(2) = ".profile"
stroptions(3) = "wc -1"
stroptions(4) = "wc -w"
stroptions(5) = "Backquote"
stroptions(6) = "$variable name"
stroptions(7) = "The number of arguments"
stroptions(8) = "Export"
stroptions(9) = "wc -c"
stroptions(10) = ".login"

stranswers(1) = "Export"
stranswers(2) = "$variable name"
stranswers(3) = ".profile"
stranswers(4) = ".login"
stranswers(5) = "Command"
stranswers(6) = "The number of arguments"
stranswers(7) = "wc -c"
stranswers(8) = "Backquote"
stranswers(9) = "wc -1"
stranswers(10) = "wc -w"

End Sub

Private Sub Fill_in_the_Blanks_Paper5()
StrQuestions(1) = "$* stands for all the arguments except the dash."
StrQuestions(2) = "The status of executing of a command is retained in the dash shell variable."
StrQuestions(3) = "To assign the value of an expression to a variable you have to use the dash construct."
StrQuestions(4) = "\ is in general a/an dash character."
StrQuestions(5) = "* is in general a/an dash character."
StrQuestions(6) = "Several commands put in a line have to be separated by dash."
StrQuestions(7) = "To run the program junk in the background you have to type dash."
StrQuestions(8) = "Commands can be grouped together by writing dash."
StrQuestions(9) = "The output of a program can be appended to a file by using dash."
StrQuestions(10) = "If command 1 is to be executed only if command 2 fails to execute, you write them as dash."

stroptions(1) = "shellscripts"
stroptions(2) = "wild card"
stroptions(3) = "junk&"
stroptions(4) = "2 | | 1"
stroptions(5) = "escape"
stroptions(6) = ">>"
stroptions(7) = "$?"
stroptions(8) = ";"
stroptions(9) = "expr"
stroptions(10) = "Command"

stranswers(1) = "Command"
stranswers(2) = "$?"
stranswers(3) = "expr"
stranswers(4) = "escape"
stranswers(5) = "wild card"
stranswers(6) = ";"
stranswers(7) = "junk&"
stranswers(8) = "shellscripts"
stranswers(9) = ">>"
stranswers(10) = "2 | | 1"

End Sub

Private Sub Fill_in_the_Blanks_Paper6()
StrQuestions(1) = "echo \* will print dash."
StrQuestions(2) = "To print a \ we will have to type dash."
StrQuestions(3) = "Pair of dash negates all the special properties of the enclosed characters."
StrQuestions(4) = "The special effect of $ is not negated by a pair of dash."
StrQuestions(5) = "dash denotes a comment in a shell script."
StrQuestions(6) = "The effect of $@ is the same as that of dash except when enclosed in double quotes."
StrQuestions(7) = "dash denotes ado nothing command."
StrQuestions(8) = "Positional parameters can be shifted using the dash command."
StrQuestions(9) = "Deterministic loops can be achieved in shell scripts using the dash construct."
StrQuestions(10) = "dash pattern filename prints the lines having the following patterns in the specified filename."

stroptions(1) = "Single quotes"
stroptions(2) = "Shift"
stroptions(3) = "Double quotes"
stroptions(4) = "grep"
stroptions(5) = "\ \"
stroptions(6) = ":"
stroptions(7) = "#"
stroptions(8) = "*"
stroptions(9) = "for"
stroptions(10) = "$*"

stranswers(1) = "*"
stranswers(2) = "\ \"
stranswers(3) = "Single quotes"
stranswers(4) = "Double quotes"
stranswers(5) = "#"
stranswers(6) = "$*"
stranswers(7) = ":"
stranswers(8) = "Shift"
stranswers(9) = "for"
stranswers(10) = "grep"

End Sub

Private Sub Fill_in_the_Blanks_Paper7()
StrQuestions(1) = "All occurances ofth pattern old can be replaced by Old using the dash filter."
StrQuestions(2) = "test dash filename will be rue ifthefile isreadable."
StrQuestions(3) = "The system call open returns dash if thefilecan be opened."
StrQuestions(4) = "The read system call returns dash read."
StrQuestions(5) = "The file descriptor for the stdout is dash."
StrQuestions(6) = "dash filter can arrange the lines of the file in alphabetical sequence."
StrQuestions(7) = "To get the number of lines in a file we type dash."
StrQuestions(8) = "Specifying no action in awk command implies that dash is that default action."
StrQuestions(9) = "dash is a non interactive stream editor."
StrQuestions(10) = "We can split a file vertically using the dash command."

stroptions(1) = "sort"
stroptions(2) = "O"
stroptions(3) = "wc -1 filename"
stroptions(4) = "-r"
stroptions(5) = "print"
stroptions(6) = "cut"
stroptions(7) = "The number of bytes"
stroptions(8) = "Tr"
stroptions(9) = "Sed"
stroptions(10) = "O2"

stranswers(1) = "Tr"
stranswers(2) = "-r"
stranswers(3) = "O"
stranswers(4) = "The number of bytes"
stranswers(5) = "O2"
stranswers(6) = "sort"
stranswers(7) = "wc -1 filename"
stranswers(8) = "print"
stranswers(9) = "Sed"
stranswers(10) = "cut"

End Sub

Private Sub Fill_in_the_Blanks_Paper8()
StrQuestions(1) = "New processes can be created using the dash system calls."
StrQuestions(2) = "The dash command allows you to see which users are currently logged in the system."
StrQuestions(3) = "You can delay a command from being executed by 5 seconds using dash."
StrQuestions(4) = "The awk command can be run from afile using the dash option."
StrQuestions(5) = "The write system call writes into a file from the dash."
StrQuestions(6) = "The dash command allows a process to run in the background even after the parent terminates."
StrQuestions(7) = "Two linked files share the same dash."
StrQuestions(8) = "patters$ when used with grep, will match the pattern that dash with pattern."
StrQuestions(9) = "^pattern when usedwith grep, will match the pattern that dash with pattern."
StrQuestions(10) = "dash lists non-printable characters as octal values."

stroptions(1) = "buffer"
stroptions(2) = "ends"
stroptions(3) = "nohup"
stroptions(4) = "ls -b"
stroptions(5) = "sleep 5"
stroptions(6) = "begins"
stroptions(7) = "who"
stroptions(8) = "inode"
stroptions(9) = "-F "
stroptions(10) = "fork and /exec"

stranswers(1) = "fork and /exec"
stranswers(2) = "who"
stranswers(3) = "sleep 5"
stranswers(4) = "-F "
stranswers(5) = "buffer"
stranswers(6) = "nohup"
stranswers(7) = "inode"
stranswers(8) = "ends"
stranswers(9) = "begins"
stranswers(10) = "ls -b"

End Sub

Private Sub Fill_in_the_Blanks_Paper9()
StrQuestions(1) = "dash lists the inode number of the file."
StrQuestions(2) = "dash lists additional info such as permissions, links and size."
StrQuestions(3) = "dash displays a symbol after the files in a directory that distinguishes between directories, executable files and symbolically linked files."
StrQuestions(4) = "Mot of configuration files are present in dash directory."
StrQuestions(5) = "dash can be used to change the default access permissions of file."
StrQuestions(6) = "Working conditions of all users can bechanged by changing dash file."
StrQuestions(7) = "An unsuccessful command returns a / an dash value as exit code."
StrQuestions(8) = "A file abnormally terminated while being edited, using vi, can be recovered using dash command."
StrQuestions(9) = "To create textat the beginning of a new line below the current line, vi has a commanddash."
StrQuestions(10) = "To return to command mode of vi, from editin mode, one has to press dash key."

stroptions(1) = "Non zero"
stroptions(2) = "O"
stroptions(3) = "Ls -l"
stroptions(4) = "<esc>"
stroptions(5) = "ls -F"
stroptions(6) = "vi - filename"
stroptions(7) = "unmask"
stroptions(8) = "ls -i"
stroptions(9) = "/etc /profile"
stroptions(10) = "/etc"

stranswers(1) = "ls -i"
stranswers(2) = "Ls -l"
stranswers(3) = "ls -F"
stranswers(4) = "/etc"
stranswers(5) = "unmask"
stranswers(6) = "vi - filename"
stranswers(7) = "Non zero"
stranswers(8) = "/etc /profile"
stranswers(9) = "O"
stranswers(10) = "<esc>"

End Sub

Private Sub Fill_in_the_Blanks_Paper10()
StrQuestions(1) = "To inspect file with read only option, we have to give dash command."
StrQuestions(2) = "To define a printer as default printer, dash command has to be given."
StrQuestions(3) = "dash permission allows a user to search the directory for a filename."
StrQuestions(4) = "Most commonly used commands ar present in dash directory."
StrQuestions(5) = "A single command which list all files in a directory distinguishing between a directory, executable and symbolic link files is dash."
StrQuestions(6) = "The technique in which UNIX system serves all users in a polled manner, providing some fixed time to each one, is called dash."
StrQuestions(7) = "Two main functions of shell in UNIX are command interpretor and dash."
StrQuestions(8) = "An option of tar, which writes the fileat the end of current fiels in archive, is dash."
StrQuestions(9) = "The dash command determines about space being used by files & directories."
StrQuestions(10) = "To print a simple text file we use dash."

stroptions(1) = "/usr/bin-> /bin"
stroptions(2) = "du"
stroptions(3) = "execute"
stroptions(4) = "Timesharing"
stroptions(5) = "Lpstat -d printer_name"
stroptions(6) = "Programming language"
stroptions(7) = "lp filename"
stroptions(8) = "view filename"
stroptions(9) = "r"
stroptions(10) = "ls -F"

stranswers(1) = "view filename"
stranswers(2) = "Lpstat -d printer_name"
stranswers(3) = "execute"
stranswers(4) = "/usr/bin-> /bin"
stranswers(5) = "ls -F"
stranswers(6) = "Timesharing"
stranswers(7) = "Programming language"
stranswers(8) = "r "
stranswers(9) = "du"
stranswers(10) = "lp filename"

End Sub

Private Sub Fill_in_the_Blanks_Paper11()
StrQuestions(1) = "In vi, to change a word in command mode, one has to type dash."
StrQuestions(2) = "A successful command returns dash as exit code."
StrQuestions(3) = "By giving dash in the end of command, it is executed in background."
StrQuestions(4) = "chmod dash f1, will make file f1 executable by all users."
StrQuestions(5) = "When a file system is mounted on a directory, the files prevvioysly present in that directory dash accessible."
StrQuestions(6) = "Encrypted password of user is stored in dash file."
StrQuestions(7) = "For matching delimeter we have to use dash, in vi command mode."
StrQuestions(8) = "To redirect the current date without time , in a file m we have to give dash."
StrQuestions(9) = "echo * will display dash in current directory."
StrQuestions(10) = "dash is a set of programs which controls all hardware accessibility of a machine."

stroptions(1) = "&"
stroptions(2) = "%"
stroptions(3) = "date | cut -c1-10-> filename"
stroptions(4) = "zero"
stroptions(5) = "/etc/shadow"
stroptions(6) = "all filename"
stroptions(7) = "+x"
stroptions(8) = "cw"
stroptions(9) = "operating system"
stroptions(10) = "are not"

stranswers(1) = "cw"
stranswers(2) = "zero"
stranswers(3) = "&"
stranswers(4) = "+x"
stranswers(5) = "are not"
stranswers(6) = "/etc/shadow"
stranswers(7) = "%"
stranswers(8) = "date | cut -c1-10-> filename"
stranswers(9) = "all filename"
stranswers(10) = "operating system"

End Sub

Private Sub Matching_Columns_Paper1()
strcolumnA(1) = "The ls command with filename that lists all files that end with the letter a is"
strcolumnA(2) = "The ls command with ilename that lists all files that have 3 characters in their name is"
strcolumnA(3) = "The ls command with filename that lists all files hae 3 or more characters in their name is"
strcolumnA(4) = "The command that will remove a directory and any files in sundirectories in it is"
strcolumnA(5) = "The ls command and ilename that lists all files that end with an even number is"
strcolumnA(6) = "The ls command and filename that lists all files having the character sequence 123 somewhere in the filename is"
strcolumnA(7) = "One of the ways to delete the file –a is"
strcolumnA(8) = "The command to hard link two files in a directory is"
strcolumnA(9) = "The command to symbolically link two files in a directory is"
strcolumnA(10) = "The command option that displays any non printable characters in a filename is"

strcolumnB(1) = "ls *[02468]"
strcolumnB(2) = "ls *a"
strcolumnB(3) = "ls ???"
strcolumnB(4) = "ls ???*"
strcolumnB(5) = "ln f1 f2"
strcolumnB(6) = "ln –s f1 f2"
strcolumnB(7) = "ls *123*"
strcolumnB(8) = "rm ---- -a"
strcolumnB(9) = "ls –q"
strcolumnB(10) = "rm –r dirname"

stranswers(1) = "ls *a"
stranswers(2) = "ls ???"
stranswers(3) = "ls ???*"
stranswers(4) = "rm –r dirname"
stranswers(5) = "ls *[02468]"
stranswers(6) = "ls *123*"
stranswers(7) = "rm ---- -a"
stranswers(8) = "ln f1 f2"
stranswers(9) = "ln –s f1 f2"
stranswers(10) = "ls –q"

End Sub

Private Sub Matching_Columns_Paper2()
strcolumnA(1) = "Suspend execution of a running process."
strcolumnA(2) = "Resume running a suspended process."
strcolumnA(3) = "Suspend the output (to screen) of a process."
strcolumnA(4) = "Resume sending the output of a process to screen."
strcolumnA(5) = "Send (signal) an ijterrupt request to a process."
strcolumnA(6) = "Send (signal) an end of file signal"
strcolumnA(7) = "Remove last word from command line input buffer."
strcolumnA(8) = "Delete or clear the current command line input buffer."
strcolumnA(9) = "Display the contents of a file, pausing between screens."
strcolumnA(10) = "Remove duplicate adjacent lines from a text file."

strcolumnB(1) = "fg"
strcolumnB(2) = "[ctrl] c"
strcolumnB(3) = "[ctrl] u"
strcolumnB(4) = "[ctrl] s"
strcolumnB(5) = "[ctrl] q"
strcolumnB(6) = "more less page"
strcolumnB(7) = "uniq"
strcolumnB(8) = "[ctrl] z"
strcolumnB(9) = "[ctrl] w"
strcolumnB(10) = "[ctrl] d"

stranswers(1) = "[ctrl] z"
stranswers(2) = "fg"
stranswers(3) = "[ctrl] s"
stranswers(4) = "[ctrl] q"
stranswers(5) = "[ctrl] c"
stranswers(6) = "[ctrl] d"
stranswers(7) = "[ctrl] w"
stranswers(8) = "[ctrl] u"
stranswers(9) = "more less page"
stranswers(10) = "[ctrl] z"
End Sub

Private Sub Matching_Columns_Paper3()
strcolumnA(1) = "List processes running."
strcolumnA(2) = "Show if a user is logged on, their real name and mail status"
strcolumnA(3) = "Deletes a directory only if it is empty"
strcolumnA(4) = "Login to another computer for specific purpose of copying files between the two systems."
strcolumnA(5) = "Remove all files in current directory."
strcolumnA(6) = "Force shell script to run as part of the current shell rather than invoking a subshell. Use the cshell version."
strcolumnA(7) = "Show the current directory that you are in.(NOT ls)/"
strcolumnA(8) = "Tell the system to continue  running a process, even if the user logs out."
strcolumnA(9) = "Lists non printable characers as ?."
strcolumnA(10) = "Lists thefiles along with their i-node numbers."

strcolumnB(1) = "rmdir"
strcolumnB(2) = "lls –i"
strcolumnB(3) = "finger"
strcolumnB(4) = "ls –q"
strcolumnB(5) = "ftp"
strcolumnB(6) = "rm –r"
strcolumnB(7) = "nohup"
strcolumnB(8) = "ps"
strcolumnB(9) = "source"
strcolumnB(10) = "pwd or echo $cwd"

stranswers(1) = "ps"
stranswers(2) = "finger"
stranswers(3) = "rmdir"
stranswers(4) = "ftp"
stranswers(5) = "rm –r"
stranswers(6) = "source"
stranswers(7) = "pwd or echo $cwd"
stranswers(8) = "nohup"
stranswers(9) = "ls –q"
stranswers(10) = "lls –i"

End Sub

Private Sub Matching_Columns_Paper4()
strcolumnA(1) = "What communication command allows a user to communicate with another user logged on by splitting the screen and providing two way communication?"
strcolumnA(2) = "What communication command provides communication to another user logged on by writing to the bottom of their terminal."
strcolumnA(3) = "What command controls (turns on or off) the ability of another user to perform either of the commands above?"
strcolumnA(4) = "What command will remove a directory and ant files nd subdirectories in it?"
strcolumnA(5) = "What is the optio of a file manipulation command(cp rm mv) that will prompt the userfor confirmation before acting?"
strcolumnA(6) = "What command and optio lists the current processes running for a specified user?"
strcolumnA(7) = "What command and option lists all current processes running?"
strcolumnA(8) = "What command can destroy a process, given its id?"
strcolumnA(9) = "What conrol character signals an end of file?"
strcolumnA(10) = "Give a command sequence that executes cmdl and if it issuccessful, cmd2 is also executed."

strcolumnB(1) = "rm –r dirname"
strcolumnB(2) = "talk"
strcolumnB(3) = "[ctrl] d"
strcolumnB(4) = "cmd1 && cmd2"
strcolumnB(5) = "mesg n or mesg y"
strcolumnB(6) = "write"
strcolumnB(7) = " ps –u userid"
strcolumnB(8) = "rm –I fn"
strcolumnB(9) = "kill"
strcolumnB(10) = "ps –e"

stranswers(1) = "talk"
stranswers(2) = "write"
stranswers(3) = "mesg n or mesg y"
stranswers(4) = "rm –r dirname"
stranswers(5) = "rm –I fn"
stranswers(6) = "rm –I fn"
stranswers(7) = "ps –e"
stranswers(8) = "kill"
stranswers(9) = "[ctrl] d"
stranswers(10) = "cmd1 && cmd2"
End Sub
Private Sub Matching_Columns_Paper5()
strcolumnA(1) = "awk command script lnes are composed of two parts"
strcolumnA(2) = "give the expression that tells awk to perform a series of commands before reading the first line of text file."
strcolumnA(3) = "Give the expression that tells awk to perform a series of commands after reading the last line of text file."
strcolumnA(4) = "What is the awk riable that contains the current count of text lies read?"
strcolumnA(5) = "Gie the expression that tests the 3rd field in a line for the word gopher."
strcolumnA(6) = "Give the action that adds the 2nd and 4th fields togethero and prints them. Include the {} around the statement"
strcolumnA(7) = "Give the action that prints the number of the fields in a line and the content of the last field. Include the {} around the statement."
strcolumnA(8) = "Write a sed command statement that changes the 3rd occurance of cow in a line to bull."
strcolumnA(9) = "What variable returns the number of arguments on the command line?"
strcolumnA(10) = "What variable returns all of the arguments on the command line as individual separate strings?"

strcolumnB(1) = "NR"
strcolumnB(2) = "Exp [cmd]"
strcolumnB(3) = "$3 ~ /gopher/{action}"
strcolumnB(4) = "END{actions}"
strcolumnB(5) = "{print $2 + $ 4}"
strcolumnB(6) = "$#"
strcolumnB(7) = "sed ‘s/cow/bull/3’ fn"
strcolumnB(8) = "{print NF, $NF}"
strcolumnB(9) = "$@"
strcolumnB(10) = "BEGIN{actions}"

stranswers(1) = "Exp [cmd]"
stranswers(2) = "BEGIN{actions}"
stranswers(3) = "END{actions}"
stranswers(4) = "NR"
stranswers(5) = "$3 ~ /gopher/{action}"
stranswers(6) = "{print $2 + $ 4}"
stranswers(7) = "{print NF, $NF}"
stranswers(8) = "sed ‘s/cow/bull/3’ fn"
stranswers(9) = "$#"
stranswers(10) = "BEGIN{actions"
End Sub

Private Sub Matching_Columns_Paper6()
strcolumnA(1) = "To change the prompt string we have to define "
strcolumnA(2) = "The same file which runs automatically for all users before login."
strcolumnA(3) = "Command that reports status of LP request Scheduler."
strcolumnA(4) = "Command that displays all status information for a printer."
strcolumnA(5) = "Command that displays status of output request for users."
strcolumnA(6) = "To give message to all logged in Users simultaneously, we use"
strcolumnA(7) = "Name of file is stored in "
strcolumnA(8) = "Access permissions of file name is stored in"
strcolumnA(9) = "To see disk space occupied by a directory, the command is"
strcolumnA(10) = "To copy contents of directory to a floppy"

strcolumnB(1) = "lpstat –r"
strcolumnB(2) = "PS1"
strcolumnB(3) = "/etc/profile"
strcolumnB(4) = "Lpstat –t"
strcolumnB(5) = "Wall"
strcolumnB(6) = "Dd"
strcolumnB(7) = "Lpstat –u"
strcolumnB(8) = "Du"
strcolumnB(9) = "Directory of file"
strcolumnB(10) = "Inode of file"

stranswers(1) = "PS1"
stranswers(2) = "/etc/profile"
stranswers(3) = "lpstat –r"
stranswers(4) = "Lpstat –t"
stranswers(5) = "Lpstat –u"
stranswers(6) = "Wall"
stranswers(7) = "Directory of file"
stranswers(8) = "Inode of file"
stranswers(9) = "Du"
stranswers(10) = "Dd"

End Sub

Private Sub Matching_Columns_Paper7()
strcolumnA(1) = "In vi from command mode, to append after the current cursor position, we use"
strcolumnA(2) = "In vi from command mode, to append at theend of current line, we use"
strcolumnA(3) = "In vi from command mode, to create a line below the current line, we use"
strcolumnA(4) = "In I from command mode, to create a line above the current line, we use"
strcolumnA(5) = "In ed, to see the specified line, we use"
strcolumnA(6) = "In ed, to change contents of line , we use"
strcolumnA(7) = "In ed to enter text below the specified line in buffer,m we use"
strcolumnA(8) = "In ed, to tneter text before thespecified lin in buffer we use"
strcolumnA(9) = "To delete a line in vi we use"
strcolumnA(10) = "To delete a line in ed we use"

strcolumnB(1) = "dd"
strcolumnB(2) = "a"
strcolumnB(3) = "A"
strcolumnB(4) = "I"
strcolumnB(5) = "D"
strcolumnB(6) = "a"
strcolumnB(7) = "P"
strcolumnB(8) = "O"
strcolumnB(9) = "s"
strcolumnB(10) = "o"

stranswers(1) = "a"
stranswers(2) = "A"
stranswers(3) = "o"
stranswers(4) = "O"
stranswers(5) = "P"
stranswers(6) = "o"
stranswers(7) = "s"
stranswers(8) = "I"
stranswers(9) = "dd"
stranswers(10) = "D"

End Sub


Private Sub Matching_Columns_Paper8()
strcolumnA(1) = "Command to create the poster is"
strcolumnA(2) = "To give a message to any user at the time of logging, we have to modify file"
strcolumnA(3) = "If a directory does not have this permission you can not locate a file in it."
strcolumnA(4) = "If a directory does not have not permission you can not locate contents of a  file in it"
strcolumnA(5) = "If a directory does not have this permission you can not chane contents of a file in it"
strcolumnA(6) = "The command used to do any pending I/O operations "
strcolumnA(7) = "The file which gets executedforevery user before it logs in"
strcolumnA(8) = "Command which gives all differences between two files"
strcolumnA(9) = "Command which gives location of 1st difference between two files"
strcolumnA(10) = "Command which can print lines of 1st file matching with 2nd file"

strcolumnB(1) = "/etc/motd"
strcolumnB(2) = "banner"
strcolumnB(3) = "display"
strcolumnB(4) = "execute"
strcolumnB(5) = "sync"
strcolumnB(6) = "/etc/profile"
strcolumnB(7) = "write"
strcolumnB(8) = "read"
strcolumnB(9) = "diff"
strcolumnB(10) = "comm."

stranswers(1) = "banner"
stranswers(2) = "/etc/motd"
stranswers(3) = "execute"
stranswers(4) = "read"
stranswers(5) = "write"
stranswers(6) = "sync"
stranswers(7) = "/etc/profile"
stranswers(8) = "diff"
stranswers(9) = "display"
stranswers(10) = "comm"

End Sub
Private Sub Matching_Columns_Paper9()
strcolumnA(1) = "Command which changes the priority of a process"
strcolumnA(2) = "Command used to rename file"
strcolumnA(3) = "Command to know about type of file"
strcolumnA(4) = "To take out a column from each line of a file we use "
strcolumnA(5) = "To add contents of two files one after the other"
strcolumnA(6) = "To add contents of two files horizontally i.e. 1st to 1st oine etc."
strcolumnA(7) = "To perform arithmetic operation on terminal we can use"
strcolumnA(8) = "Command used for data conversion and deals with raw data"
strcolumnA(9) = "Command used to see list jobs ready to run at  a  specified time."
strcolumnA(10) = "Command used to run a command at a specified time."

strcolumnB(1) = "dd"
strcolumnB(2) = "cat"
strcolumnB(3) = "nice"
strcolumnB(4) = "cut"
strcolumnB(5) = "bc"
strcolumnB(6) = "at"
strcolumnB(7) = "atq"
strcolumnB(8) = "file"
strcolumnB(9) = "mv"
strcolumnB(10) = "paste"

stranswers(1) = "nice"
stranswers(2) = "mv"
stranswers(3) = "file"
stranswers(4) = "cut"
stranswers(5) = "cat"
stranswers(6) = "paste"
stranswers(7) = "bc"
stranswers(8) = "dd"
stranswers(9) = "atq"
stranswers(10) = "at"

End Sub

Private Sub Matching_Columns_Paper10()
strcolumnA(1) = "Option of cpio used for copying data from tape to disk."
strcolumnA(2) = "Option of cpio used for coyingg data from disk to tape."
strcolumnA(3) = "Command used to see the groupid & userid of yourself"
strcolumnA(4) = "Command used to see in system"
strcolumnA(5) = "Command used to see names of all groups in system"
strcolumnA(6) = "Command to see the values of defaults set in different files of /etc/default."
strcolumnA(7) = "Command by which you can see very large files efficiently as compared to ed"
strcolumnA(8) = "Command to find the checksum and blck count of a file."
strcolumnA(9) = "Command available to compress a file"
strcolumnA(10) = "Command which gives information about users login  on machine with latest user 1st"

strcolumnB(1) = "dispgid"
strcolumnB(2) = "defadm"
strcolumnB(3) = "–i"
strcolumnB(4) = "pack"
strcolumnB(5) = "id"
strcolumnB(6) = "last"
strcolumnB(7) = "–o"
strcolumnB(8) = "bfs"
strcolumnB(9) = "sum"
strcolumnB(10) = "dispuid"

stranswers(1) = "–i"
stranswers(2) = "–o"
stranswers(3) = "id"
stranswers(4) = "dispuid"
stranswers(5) = "dispgid"
stranswers(6) = "defadm"
stranswers(7) = "bfs"
stranswers(8) = "sum"
stranswers(9) = "pack"
stranswers(10) = "last"

End Sub

Private Sub TrueorFalse_Paper1()
StrQuestions(1) = "File management is the task ofhe shell in UNIX."
StrQuestions(2) = "The deice derivers are a part of the kernel in UNIX."
StrQuestions(3) = "Real time operating systems must be batch processing system."
StrQuestions(4) = "A person who will write the kernel of an O.S. must hae knowledge about te hardware."
StrQuestions(5) = "A user is willing to delete a file from the disk. He would directly interactwith he kernel of the O.S."
StrQuestions(6) = "A directory is also a file in UNIX."
StrQuestions(7) = "The absolute path name of a file must always sart with the root."
StrQuestions(8) = "The directory directly stores information about the physical location of the file."
StrQuestions(9) = "The directory directly stores information about the file protection."
StrQuestions(10) = "The ls –l <f-name> command shows us whether f-name is afile or a directory."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper2()
StrQuestions(1) = "We can delete a directory containing only one file, using the rmdir command."
StrQuestions(2) = "The rm –r command delees the filrs and not thedirectory."
StrQuestions(3) = "The ls –l command displays the files in a directory sorted rowwise."
StrQuestions(4) = "A file having protection 466 cannot be read by anybody other than the owner."
StrQuestions(5) = "Only the owner of a file and the superuser can change the access permissions of a file."
StrQuestions(6) = "A file permission can be changed by the chmod command."
StrQuestions(7) = "chmod u-x gives he file execute permission."
StrQuestions(8) = "The command  pwd shows the name of the  current directory."
StrQuestions(9) = "It is possible to link a file of one file system to aother."
StrQuestions(10) = "The advantage of the hierarchical file system is that it consumes less space."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper3()
StrQuestions(1) = "UNIX makes a distinction between the text files and the binary files."
StrQuestions(2) = "The devices are considered as ordinary files in UNIX."
StrQuestions(3) = "The devices are considered as ordinary files in UNIX."
StrQuestions(4) = "Typical blocks sizes are 512/1024 bytes."
StrQuestions(5) = "The command r map* removes all files that begin with ap from al ldirectories."
StrQuestions(6) = "The command rm * removes all files from the present directories."
StrQuestions(7) = "The command mv file1 and file2 renames file1 as file2."
StrQuestions(8) = "the command mv file1 file2 keeps a duplicate copy file1 as file2."
StrQuestions(9) = "We can view a file page by page through the cat command"
StrQuestions(10) = "The more command appends new lines to a file."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper4()
StrQuestions(1) = "You can  start the vusual editor using the ed command."
StrQuestions(2) = "The vi editor is a line editor."
StrQuestions(3) = "In a I editor, you can set the length of the line and automatically insert line feed by the wrap command."
StrQuestions(4) = "In a vi editor you can work in either in the editor mode or in the command mode."
StrQuestions(5) = "Pressing O inserts a line at the top of the current line in vi."
StrQuestions(6) = "Pressing y deletes a character in vi."
StrQuestions(7) = "A file edited through vi can be saved by pressing ZZ."
StrQuestions(8) = "All files edited through vi are stored as ASCII files. "
StrQuestions(9) = "Pressing dd deletes a line in vi."
StrQuestions(10) = "It is possible to switch from a vi editor and the ed editor by pressing Escape."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper5()
StrQuestions(1) = "It is possible to switch from a vi editor and the ed editor by pressing: when in the command mode."
StrQuestions(2) = "An i-node is associated with every file."
StrQuestions(3) = "An i-node is a special file."
StrQuestions(4) = "The device files are stored in the /etc directory."
StrQuestions(5) = "The devices files can be seen through the ls –l command."
StrQuestions(6) = "A terminal is a character oriented file."
StrQuestions(7) = "A disk is a character oriented file."
StrQuestions(8) = "Deice files can be either character oriented or block oriented."
StrQuestions(9) = "We can  execute a deice file directly."
StrQuestions(10) = "An i-node keeps track of the physical location of a file."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper6()
StrQuestions(1) = "A physical file is stored in continuous locations."
StrQuestions(2) = "The i-nodes reside in the disk."
StrQuestions(3) = "The maximum size a file can have depends on the i-node size."
StrQuestions(4) = "The i-node number corresponding to afile cannot be seen by the user."
StrQuestions(5) = "A filter is a program that can take inputs from the standard input and produces the output on the standard output."
StrQuestions(6) = "We can write the output of a program on a file using redirection."
StrQuestions(7) = "We can append the contents of a file to another using >>."
StrQuestions(8) = "cat is a filter."
StrQuestions(9) = "tee is a filter."
StrQuestions(10) = "more is a filter."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper7()
StrQuestions(1) = "sort is afilter."
StrQuestions(2) = "The test command is used to test a condition of an in statement."
StrQuestions(3) = "The test command can be applied over strings and files but not on numeric values."
StrQuestions(4) = "UNIX files assume all data to be consisting of streams of bytes."
StrQuestions(5) = "The df command displays the systems in a disk."
StrQuestions(6) = "A file system consists of two things – the data files and the i-nodes."
StrQuestions(7) = "The system calls are the system library functions."
StrQuestions(8) = "The system calls can be used by the users directly."
StrQuestions(9) = "UNIX commands use the system calls for their operation."
StrQuestions(10) = "Using the access system call we can know whether a file is read protected."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper8()
StrQuestions(1) = "printf is a system library function and not a system call."
StrQuestions(2) = "UNIX adopts a priority based scheduling strategy."
StrQuestions(3) = "UNIX process priorites cannot be changed by the user."
StrQuestions(4) = "The password of a user is kept in the /etc/passwd file."
StrQuestions(5) = "The superuser can view the password of user in the /etc/paswd file."
StrQuestions(6) = "The password of a user can be changed using the passwd command."
StrQuestions(7) = "The process scheduling is done by the kernel in UNIX."
StrQuestions(8) = "The memory management is done by the UNIX shell."
StrQuestions(9) = "When a program is to be compiled the user will hae to make a kernel call."
StrQuestions(10) = "It is not possible to see how many users are currently working in the system."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper9()
StrQuestions(1) = "A process can be killed by the kill command."
StrQuestions(2) = "A user can know the process identification of a process using the ps command."
StrQuestions(3) = "The kill -9 command cannot be ignored."
StrQuestions(4) = "The UNIX system call signal is used to handle interrupts."
StrQuestions(5) = "We can search for a pattern in a file using grep command."
StrQuestions(6) = "The tr command is essential  a filter."
StrQuestions(7) = "We can replace one character by another fo all lines in a file using the tr command."
StrQuestions(8) = "It is possible extract the first 4 lines of a file using the head command."
StrQuestions(9) = "The awk command can take decisions."
StrQuestions(10) = "The parameters of awk command can also be fed through a file."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper10()
StrQuestions(1) = "UNIX is a multitasking system."
StrQuestions(2) = "More than one processes can be active in UNIX een if there is only one processor."
StrQuestions(3) = "A background process runs as a child process of the shell."
StrQuestions(4) = "If the parent process is killed then, always the child process is also terminated."
StrQuestions(5) = "A process can be killed by pressing Control+C keys."
StrQuestions(6) = "The shell in UNIX starts immediately when the power of the system is put on."
StrQuestions(7) = "In shell programming p1|| p2 denotes that the p2 is run only if p1 fails."
StrQuestions(8) = "$1 is a shell ariable which has got a specific meaning."
StrQuestions(9) = "In he statement name = 12,1 and 2 elements of a string and not the number 12."
StrQuestions(10) = "The output of the statement name = expr2 + 3 is 4."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub


Private Sub questions_paper1()
StrQuestions(1) = "Justify  or modify thefollowing statements with reasons." & Chr(13) & _
                  "a.  File management is the task of the shell in UNIX." & Chr(13) & _
                  "b.  The deice dreiers are part of the kernel in UNIX" & Chr(13) & _
                  "c.  Real time operating systems must be batch processing systems." & Chr(13) & _
                  "d.  A person who will write the kernel of an O.S. must have a knowledge  about hardware." & Chr(13) & _
                  "e.  A user is willing to delete a file from the disk; he would directly interactwith the kernel of the O.S"
StrQuestions(2) = "a. What are the major components of an O.S?" & Chr(13) & _
                  "b. Which of the following instructions should be executed in the supervisory(kernel) mode and which should be eexecuted in the user mode?" & Chr(13) & _
                  "i)  Setting the alue of a timer" & Chr(13) & _
                  "ii) Reading a clock" & Chr(13) & _
                  "iii)    Clearing memory" & Chr(13) & _
                  "iv) Switcing from user mode to supervisory mode"
StrQuestions(3) = "a.  What ae the major features of the UNIX operating system that has made it so popular?" & Chr(13) & _
                  "b.  where is a file stored in UNIX?" & Chr(13) & _
                  "c.  Can the directory be also calleda file? If not – why? If so – what are its contents?" & Chr(13) & _
                  "d.  What is the difference between absolute and relatie path names of an UNIX file?"
StrQuestions(4) = "a.  What is a filter? What is its advantage?" & Chr(13) & _
                  "b.  How can you use “cat” to create a file “story” consisting of keyboard nput sandwiched between two files file1 and file2?" & Chr(13) & _
                  "c.  How can you know whether a particular file name is a file or a directory?" & Chr(13) & _
                  "d.  A file has got a protection 756(octal). What protection does it really have?"
StrQuestions(5) = "a.  What is the difference between a process and a program?" & Chr(13) & _
                  "b.  Name two events that can cause a process to lose control of the processor?" & Chr(13) & _
                  "c.  what is meant by the statement that a process is sequential?" & Chr(13) & _
                  "d.  When can we surely say that concurrent processes will not result in time dependent errors?"
StrQuestions(6) = "a.  Diffrentiate between multiprogramming and multiprocessing." & Chr(13) & _
                  "b.  What is meant by time sharing sytem? Why is it required? Draw  the process sate diagam ofa time sharing systm & explain it."
StrQuestions(7) = "a.  Suppose you are running the command rm <filename> from the shell prompt. Explain how the shell will execute the command rm." & Chr(13) & _
                  "b.  How can you run a process in the background?"
StrQuestions(8) = "a.  What is the function of the exec system call. What is the difference between ‘exec’ in a code and calling a subrouting code?" & Chr(13) & _
                  "b.  Which process performs the exec system call, the parent or the child or the copy of the parent? How are thet identified in a program?"
StrQuestions(9) = "a.  What does the parent do in UNIX, while the child is executing?" & Chr(13) & _
                  "b.  What is the function of the wait system call with respect to the UNIX process creation?"
StrQuestions(10) = "a.  What is the function of signal system call?" & Chr(13) & _
                   "b.  How can you ensure that a process will be killed?" & Chr(13) & _
                   "c.  What is the difference between cret() and open() system calls?"
End Sub

Private Sub questions_paper2()
StrQuestions(1) = "a.  What are the different queuwes maintained by an operting systm for the program executed by the CPU?" & Chr(13) & _
                  "b.  when a process is scheduled by a short term scheduler, from where is the program executed by the CPU?"
StrQuestions(2) = "a.  Who brings in a process from the disk to the Main Memory?" & Chr(13) & _
                  "b.  What is the advantage of the First Come First Served policy of process scheduling?" & Chr(13) & _
                  "c.  what is the advantage of the shortest job first scheduling?"
StrQuestions(3) = "a.  What is meant by staration? How is the problem of staration alleviated?" & Chr(13) & _
                  "b.  What isth relationship between the priorities of kernel processes and user processes? Can a user modify the priorities of kernel processes?"
StrQuestions(4) = "a.  through which command can you find the i-node number assigned to a file? What other information do you get through the same command?" & Chr(13) & _
                  "b.  If you link a file F1 with another existing filr F2 and the file F2 has an i-node number assigne as 43215, what will be the i-node number assigned to F1?" & Chr(13) & _
                  "c.  What would be your answer ifthe file F1 in abovequestion was already existing?"
StrQuestions(5) = "a.  What would happen to the i-noes if a file is copied into another?" & Chr(13) & _
                  "b.  How can you check what file systems your system has?"
StrQuestions(6) = "a.  What is the utput of the command ls –l/dev?" & Chr(13) & _
                  "b.  Explain clearly what do you mean by mounting a file system on another   "
StrQuestions(7) = "a.  Present the overall structure of the UNIX operting system and the function of each component." & Chr(13) & _
                  "b.  What do you mean by the statement “Thte shell essentially acts as a command interpreter?"
StrQuestions(8) = "You have to store a number of files in UNIX file system. The files are:" & Chr(13) & _
                  "a.  average.bas (a BASIC program to compute the average of a set of integers)" & Chr(13) & _
                  "b.  tom.let( a leter to your friend Tom)" & Chr(13) & _
                  "c.  sup.let (a letter to your immediate superior" & Chr(13) & _
                  "d.  salary(a program that computes  the salary of your staff.)" & Chr(13) & _
                  "e.  appl.1 ( you are tying for a better job. This is your first application.)" & Chr(13) & _
                  "f.  answer.1(This is the first reply you got from the company FORProg pt. Ltd.)" & Chr(13) & _
                  "g.  mywife.let (this is a letter to your wife stating your experience of th first day in the job." & Chr(13) & _
                  "h.  apple( this is your application to a company XProg Pt. Ltd.)" & Chr(13) & _
                  "i.  itax.c (A C program to compute  income tax)" & Chr(13) & _
                  "Explain your designs with reasons."
StrQuestions(9) = "a.  Write a single command line statement (using redirection, pipes and tees) that will accept lines typed through the keyboard and copy them in sorted order in a file “newsortfile” and will also count the number of lines and store the count in another file “count_of_lines”?" & Chr(13) & _
                  "b.  How are the UNIx files protected? Through which command can you check whether a file is read protected or not?"
StrQuestions(10) = "a.  State two ways by which you can change the protection of a file to such that owner will hae read and write permissions, the group will hae only read permission and others will have no access?" & Chr(13) & _
                   "b.  You want to remoe files ( from the current directory) whose names start with xy. Some of these files can be write protedtec, but you do not know which ones are such. What are the steps hat you will adopt to remove the desired files? Gie the command sequences with the necessary explanation."
End Sub

Private Sub questions_paper3()
StrQuestions(1) = "a.  What do your mean by a multitasking system? What is its difference with a multiprocessing system" & Chr(13) & _
                  "b.  What is meant by a multiprogramming system? What is its difference with a multiprocessing system?"
StrQuestions(2) = "a.  What is meant by interrupt? With the help of a flowchart show what happens when an interrupt occurs?" & Chr(13) & _
                  "b.  Explain clearly th role an interrupt plays in a –" & Chr(13) & _
                  "i)  Multiprogrammed system" & Chr(13) & _
                  "ii) Time shared system"
StrQuestions(3) = "a.  What is meant by atomiiyty of the semaphore operations?" & Chr(13) & _
                  "b.  How is process created in UNIX? Explain in brief?"
StrQuestions(4) = "a.  What is meant by atomicity of the semaphore operations?" & Chr(13) & _
                  "b.  How is process created in UNIX? Explain in brief."
StrQuestions(5) = "a.  Explain how a user is logged on in UNIX. What are the processes inoled? What happens when a user logs out but the system is not powered off?" & Chr(13) & _
                  "b.  What happens and when , to the parent process, when the foreground process terminates?"
StrQuestions(6) = "a.  What happens to the background process when the foreground process terminates?" & Chr(13) & _
                  "b.  How many parameters does the exec system call require? State their data types data types and functions."
StrQuestions(7) = "Write a c program segment which will execute the cp program as a child process. Clearly identify how many arguments are necessary and why?"
StrQuestions(8) = "a.  Explain clearly the difference between system calls and the system library functions?" & Chr(13) & _
                  "b.  Is it possible to use the sigbal system call to display a particular message when a particular key (for example K) is pressed? If not, if so, why and how?"
StrQuestions(9) = "a.  What system call would you use to read a byte at a time from the standard input? Write the system call along with the proper arguments?" & Chr(13) & _
                  "b.  You want to ensure that, while a program is running, if the user presses a particular key (for example K), it will be ignored. How can you achieve this?"
StrQuestions(10) = "For BOURNE shell -" & Chr(13) & _
                   "a.  What variable returns the buber of arguments on the command line?" & Chr(13) & _
                   "b.  What variable returns all of the arguments on the commnd line as individual separate strings?" & Chr(13) & _
                   "c.  What variable returns the success/failure status of last command inoked?" & Chr(13) & _
                   "d.  What variable returns the process id of the current process ( the wcript process id)" & Chr(13) & _
                   "e.  What variable returns the command gien in the command line?"
End Sub

Private Sub questions_paper4()
StrQuestions(1) = "a.  Write a simple Borune for structure that displays all of the arguments on the command line" & Chr(13) & _
                  "b.  Write a simple Bourne for structure that lists each of the files in the current directory."
StrQuestions(2) = "a.  Instead of alias, write a function called pswhich inokes the command ps –ef. (single line command)" & Chr(13) & _
                  "b.  In bourne shell, assign a local variable called var1,the string “hi there”." & Chr(13) & _
                  "c.  Make that variable a global or an environmental ariable"
StrQuestions(3) = "Write a Bourne case atement that performs the following code:" & Chr(13) & _
                  "NOTE: Think of the string expression in the case like a filename." & Chr(13) & _
                  "if file is named all" & Chr(13) & _
                  "Delete it" & Chr(13) & _
                  "else if file contains the letter a or A in it anywhere" & Chr(13) & _
                  "List it" & Chr(13) & _
                  "else if file filename contains only 2 characters" & Chr(13) & _
                  "copy itto temp" & Chr(13) & _
                  "Else" & Chr(13) & _
                  "Print the; message; unknown; "
StrQuestions(4) = "a.  Gie the command to setup an alias called cd that changes to the specified directory, resets the prompt to include the path name and the history numbering . Remember he single quoes gets he whole string nto he alias. The backslash makes sure that the ! history specifier is not ealuated immediately but is put into the alias. The  * history option specifies all of the argument on th command line b used toregenerate the command." & Chr(13) & _
                  "b.  Give the command hat lists all the current aliases." & Chr(13) & _
                  "c.  Give the command that remoes cd from he alias list." & Chr(13) & _
                  "d.  If he following command was issued, what woule appear in the alias list for the listing of that alias? Alias myhome echo ‘$home’"
StrQuestions(5) = "a.  Gie the set command that sets up the your login so that the last 20 commands from he previous login are available." & Chr(13) & _
                  "b.  Gie the command line command that execues the previoys command." & Chr(13) & _
                  "c.  The first line of a shell script determines the type of shell that the script runs under. For best documentation, there is a way of explicitly naming the desired shell." & Chr(13) & _
                  "Write the line of Bourne script."
StrQuestions(6) = "When shell scripts are normally executed, they run under a subshell." & Chr(13) & _
                  "a.  What command causes a script to execute in the current C shellrather that in a child shell? Answer by showing an example using myscript as the script to be run." & Chr(13) & _
                  "b.  What commane causes a script to execute in the current Bourne shell rather that in a child shell? Answer by showing an example using myscript as the script to be run. Both the C and Bourne shell allow the user to eclare local variables in their hell and scripts." & Chr(13) & _
                  "c.  Give the commad to set a local ariable in a C shell environment." & Chr(13) & _
                  "d.  Gie the command to set a local variable in the Bourne shell environment."
StrQuestions(7) = "a.  Both the C and Bourne shell allow the user to make global ariable which can be passed sub shells. Gie the commnd to make variable in the C shell globalk. Gie the command to make variable in the Bourne shell global." & Chr(13) & _
                  "b.  Both the C and Bourne shell allow the user to remove local variables previoysly declared." & Chr(13) & _
                  "Gie the command REMOE a variable in the C shell." & Chr(13) & _
                  "Give the command REMOVE a variable in the Bourne shell." & Chr(13) & _
                  "For the following, you are to write some shell script program flow structures. If    specified examples are not requested, use the following terms where needed." & Chr(13) & _
                  "Logical exp     - a logical true or false expression" & Chr(13) & _
                  "String          - a literal string" & Chr(13) & _
                  "Strvar          - a variable that can contain a string" & Chr(13) & _
                  "Numeric     - a numeric value" & Chr(13) & _
                  "Numvar      - a variable than can contain a numeric value" & Chr(13) & _
                  "Cmd             - any of the available Unix commands"
StrQuestions(8) = "Both the C and Bourne shell allow the use of the logic structure" & Chr(13) & _
                  "IF (test condition)" & Chr(13) & _
                  "Action" & Chr(13) & _
                  "ELSEIF (test condition)" & Chr(13) & _
                  "Action" & Chr(13) & _
                  "Else" & Chr(13) & _
                  "Action" & Chr(13) & _
                  "End If" & Chr(13) & _
                  "Code the aboe logic as a C shell if structure. Use the aboe terms appropriate to the if command. Include all required syntax." & Chr(13) & _
                  "Code the above sample logic as a Bourne shell if structure. Use the above terms appropriate to the if command. Include all required suntax."
StrQuestions(9) = "Both C and Bourne shell allow the us of the while structure which can perform the following basic logic." & Chr(13) & _
                  "WHILE (test expression) DO actions" & Chr(13) & _
                  "ENDWHILE" & Chr(13) & _
                  "Code the above sample logic as a C shell while structure. Use the above terms appropriate for while command. Code the above sample logic as a Bourne shell while structure." & Chr(13) & _
                  "Use the above terms appropriate for while command."
StrQuestions(10) = "Both C shell & Bourne shell allow the user to enter information while the wcript is running." & Chr(13) & _
                   "For C shell, show the command for reading input from the user into the variable name_var."
End Sub

Private Sub questions_paper5()
StrQuestions(1) = "Write single command lines using suitable filters to perform the following :" & Chr(13) & _
                  "1.  List all subdirectory names." & Chr(13) & _
                  "2.  List files which other can read and write." & Chr(13) & _
                  "3.  List all the ‘c’ source file and sizes (bytes)."
StrQuestions(2) = "write a shell script “compute” which allows you to perform some numerical operations between the data. The typical functions that may be performed by compute are shown below:" & Chr(13) & _
                  "$   compute 5 + 3" & Chr(13) & _
                  "8" & Chr(13) & _
                  "$   compute 3 – 3" & Chr(13) & _
                  "0" & Chr(13) & _
                  "$   compute 2 \*3" & Chr(13) & _
                  "6" & Chr(13) & _
                  "$   compute 3 % 2" & Chr(13) & _
                  "1" & Chr(13) & _
                  "$   compute 3 | 5" & Chr(13) & _
                  "unknown operator"
StrQuestions(3) = "a.  Write a shell script which will display “ You are welcome to UNIXSystem “ by using for loop." & Chr(13) & _
                  "b.  Write a script which will accept  the name and display on the screen. This will continue until user presses enter without  giing any character. Write his script by using while loop only. Don’t use test command."
StrQuestions(4) = "a.  write a shell script which will redirect the output of the date command without the time into a file." & Chr(13) & _
                  "b.  use a for loop to move a list of files in the current directory to another directory. How can you move all your files to another directories?"
StrQuestions(5) = "a.  Write a shell script which will test at particular time interals whether any core file has been created or not in your current directory." & Chr(13) & _
                  "b.  Write a shell script to display the number of numbers that are divisible by 3 within a particular range. Also print the numbers."
StrQuestions(6) = "a.  Write a shell scrit that will accept a string either thorugh the command line or from script while runnn. If you press <Enter>, without gigin a single character, it will take you to the ‘$’ prompt. If you accept string then the script will replace space with a ‘-‘. There should be not more than one -, as the delimeter between the words." & Chr(13) & _
                  "b.  write a set of shell commands that will keep on giving the user name and terminal number when he or she logged in."
StrQuestions(7) = "a.  What is the awk variable that contains the current count of text lnes read?" & Chr(13) & _
                  "b.  Give the expression that tests the 3rd field in a line for the word gopher $3 ~ /gopher/ {action}" & Chr(13) & _
                  "c.  Give the acton that adds the 2n and 4th fields together and prints them."
StrQuestions(8) = "a.  Give the action that prints the number of the fields in a line and the content of the last field. Include the {}around the statement." & Chr(13) & _
                  "b.  Give the data file:" & Chr(13) & _
                  "    quiz 8" & Chr(13) & _
                  "    homework 79" & Chr(13) & _
                  "    quiz 9" & Chr(13) & _
                  "    quiz 4" & Chr(13) & _
                  "    homework 84" & Chr(13) & _
                  "    quiz 10" & Chr(13) & _
                  "    quiz 5" & Chr(13) & _
                  "    quiz 9" & Chr(13) & _
                  "    homework 74" & Chr(13) & _
                  "    test        143" & Chr(13) & _
                  "Write an awk statement that accumulates total for each category listed in column 1." & Chr(13) & _
                  "c.  Write an awk statement that prints out the categories  and the associated totals."
StrQuestions(9) = "a.  Write a shell script that displays how many users there are in each group." & Chr(13) & _
                  "b.  Write a utility shell script that will remove all the files that you have specified in the command line. If you do not specify the name of the path, then script will assume the current directory."
StrQuestions(10) = "Write a shell script which will print the calendar accepting the month name and year as command line argument."
End Sub

Private Sub questions_paper6()
StrQuestions(1) = "Write a shell script that will moe all the files havine size greater than 1K to a directory called large."
StrQuestions(2) = "a.  A Unix file system has 1024 byte block size, with 32 bit address. The inode has 10 direct disk addresses, one single indirect and one double indirect address. What is the maximum file size in this system?" & Chr(13) & _
                  "b.  what important information is stored in super block of a file system?"
StrQuestions(3) = "a.  Write a shell script with while loop to execute endlessle (till terminated by user) a loop which displays contents of current directory, diskspace status, sleep for 10 secs and displays the users currently logged in, on the screen." & Chr(13) & _
                  "b.  What are system calls? How are they different freom Unix executable and library functions?"
StrQuestions(4) = "a.  What are the main functions performed by system administrator in UNIX?" & Chr(13) & _
                  "b.  Give ariouys security features available in UNIX operationg system."
StrQuestions(5) = "a.  Explain three important options each of TAR and CPIO." & Chr(13) & _
                  "b.  Explain varioys steps executed in fsck command." & Chr(13) & _
                  "c.  What is fragmentation? discuss various factors  wgucg deternube tge kicj suze if fuke system?"
StrQuestions(6) = "a.  What is demand paged memory management? What are its advantages and drawbacks?" & Chr(13) & _
                  "b.  What are varioys commands available in UNIX for taking backup?"
End Sub


Private Sub multiplechoice_Paper1()
strmultiple(1) = "The objective of a time sharing operating system is to"
strmuloption(1) = "minimize the net execution time"
strmuloption(2) = "maximize processor utilization"
strmuloption(3) = "minimize the user response time"
strmuloption(4) = "none of the above"

strmultiple(2) = "The objective of a multiprogramming operating system is to"
strmuloption(5) = "Minimize the net execution time"
strmuloption(6) = "Maximize processor utilization"
strmuloption(7) = "Minimize the user response time"
strmuloption(8) = "None of the above"

strmultiple(3) = "UNIX is an example of –"
strmuloption(9) = "batch processing operating system"
strmuloption(10) = "time sharing operating system"
strmuloption(11) = "real time sharing system"
strmuloption(12) = "none of the above"

strmultiple(4) = "In UNIX terminology, a process is a "
strmuloption(13) = "program in execution"
strmuloption(14) = "subprogram which can be called from other programs"
strmuloption(15) = "sequence of commands which are to be given in order to run a program"
strmuloption(16) = "none of the above"

strmultiple(5) = "An operating system is a software which"
strmuloption(17) = "helps the operators who are maintaining a computer system"
strmuloption(18) = "helps other programs run"
strmuloption(19) = "helps in resource management"
strmuloption(20) = "is run to recover lost data when power failure occurs."

strmultiple(6) = "The function of the fork system call is to –"
strmuloption(21) = "Split a process into several parts"
strmuloption(22) = "Create another copy of the process which are then executed concurrently"
strmuloption(23) = "Change the standard output to a specified device"
strmuloption(24) = "None of the above"

strmultiple(7) = "The UNIX command “cat abc def>xyz”"
strmuloption(25) = "Merges the two files avc and def and stored it is xyz"
strmuloption(26) = "Finds the categories of the files abc and def and stores it in xyz"
strmuloption(27) = "Tests whether the variables and def are greater than xyz"
strmuloption(28) = "None of the above"

strmultiple(8) = "The UNIX command “cp ch? Book”"
strmuloption(29) = "Copies all files staring with ch to the directory book"
strmuloption(30) = "Copies all files with three character names and starting with ch to the directory book"
strmuloption(31) = "Compares whether a file starting with ch exists in directory book"
strmuloption(32) = "None of the above"

strmultiple(9) = "The binary equivalent of 63 is"
strmuloption(33) = "0110 0011"
strmuloption(34) = "0111 1110"
strmuloption(35) = "0111 0001"
strmuloption(36) = "none of the above"

strmultiple(10) = "The maximum and minimum numbers (unsigned) that can be stored in a 8-bit word is"
strmuloption(37) = "0 and 99"
strmuloption(38) = "0 and 255"
strmuloption(39) = "-128 and 127"
strmuloption(40) = "0 and 127"

strmulanswer(1) = "minimize the user response time"
strmulanswer(2) = "Maximize processor utilization"
strmulanswer(3) = "time sharing operating system"
strmulanswer(4) = "program in execution"
strmulanswer(5) = "helps other programs run"
strmulanswer(6) = "Create another copy of the process which are then executed concurrently"
strmulanswer(7) = "Merges the two files avc and def and stored it is xyz"
strmulanswer(8) = "Copies all files with three character names and starting with ch to the directory book"
strmulanswer(9) = "none of the above"
strmulanswer(10) = "0 and 255"
End Sub

Private Sub multiplechoice_Paper2()
strmultiple(1) = "A byte is defined as"
strmuloption(1) = "A group of 16 bits"
strmuloption(2) = "A group of 4 bits"
strmuloption(3) = "One fourth of the machine word size"
strmuloption(4) = "None of the above"

strmultiple(2) = "Almost all computers stored characters strings"
strmuloption(5) = "As a sequence of ASCII coded binary strings"
strmuloption(6) = "In BCD code"
strmuloption(7) = "As a sequence of integers indicating the position of the characters in the alphabetical sequence"
strmuloption(8) = "None of the above"

strmultiple(3) = "The number of bits used to represent a character is"
strmuloption(9) = "4   "
strmuloption(10) = "8"
strmuloption(11) = "16"
strmuloption(12) = "none of the above"

strmultiple(4) = "In UNIX, the command to delete all files starting with sg and ending with hj is"
strmuloption(13) = "rm sg*hj"
strmuloption(14) = "rm sg?hj"
strmuloption(15) = "rm –prefix sg –suffix hj"
strmuloption(16) = "none of the above"

strmultiple(5) = "In UNIX the command to redirect the output of a program (say,sort) to a file called sort.out is"
strmuloption(17) = "sort sort.out"
strmuloption(18) = "sort -> sort.out"
strmuloption(19) = "sort >sort.out"
strmuloption(20) = "none of the above"

strmultiple(6) = "The instructions of a program which is currently being executed is sorted in"
strmuloption(21) = "Secondary memory"
strmuloption(22) = "Main memory"
strmuloption(23) = "Read only memory"
strmuloption(24) = "None of the above"

strmultiple(7) = "The programming language FORTRAN is mainly used for "
strmuloption(25) = "Data processing applications"
strmuloption(26) = "Scientific applications"
strmuloption(27) = "Real time applications"
strmuloption(28) = "None of the above"

strmultiple(8) = "In the programming language C,"
strmuloption(29) = "Variable names are case sensitive"
strmuloption(30) = "There are no reserved words which are not allowed as variable names"
strmuloption(31) = "Blanks can be given within a variable name"
strmuloption(32) = "None of the above"

strmultiple(9) = "The function of the UNIX system calls are to"
strmuloption(33) = "Request for service(s) offered by the operating system kernel"
strmuloption(34) = "A mechanism using which one program can call another"
strmuloption(35) = "A signal generated by the operating system signaling an error"
strmuloption(36) = "None of the above"

strmultiple(10) = "The UNIX command  “ls | more “ displays a list of "
strmuloption(37) = "Files in the current directory one line at a time"
strmuloption(38) = "Files in the current directory one screen at a time"
strmuloption(39) = "All the files in the current directory and then waits for the next command"
strmuloption(40) = "None of the above"

strmulanswer(1) = "None of the above"
strmulanswer(2) = "As a sequence of ASCII coded binary strings"
strmulanswer(3) = "8"
strmulanswer(4) = "rm sg*hj"
strmulanswer(5) = "sort >sort.out"
strmulanswer(6) = "Main memory"
strmulanswer(7) = "Scientific applications"
strmulanswer(8) = "Variable names are case sensitive"
strmulanswer(9) = "Request for service(s) offered by the operating system kernel"
strmulanswer(10) = "Files in the current directory one screen at a time"
End Sub


Private Sub multiplechoice_Paper3()
strmultiple(1) = "The UNIX command “cat xyz | grep computer”"
strmuloption(1) = "Adds the string computer to end of the file xyz"
strmuloption(2) = "Displays all lines in file xyz containing the string computer"
strmuloption(3) = "Counts number of times the string computer appears in file xyz"
strmuloption(4) = "None of the above"

strmultiple(2) = "The UNIX command “a.out \&” runs the program a.out"
strmuloption(5) = "With highest priority"
strmuloption(6) = "Only when no other process is running on the system"
strmuloption(7) = "In the background"
strmuloption(8) = "None of the above"

strmultiple(3) = "In a typical operating system, shell is a program which"
strmuloption(9) = "Reads a command line from the standard input and interprets it according to affixed set of rules"
strmuloption(10) = "Is a framework into which functions or utilities can be added when required"
strmuloption(11) = "Helps in the efficient running of a program"
strmuloption(12) = "None of the above"

strmultiple(4) = "The user startup file for Bourne shell is stored in the home directory with the name"
strmuloption(13) = "Cshrc"
strmuloption(14) = ".init"
strmuloption(15) = ".login"
strmuloption(16) = ".startup"

strmultiple(5) = "Suppose that the current directory is /home/user/xyz/prog. A possible command to change to a directory /home/user/abc/letters is"
strmuloption(17) = "Cd ../../abc/letters"
strmuloption(18) = "Cd ././abc/letters"
strmuloption(19) = "Cd ./home/user/abc/letters"
strmuloption(20) = ""

strmultiple(6) = "The UNIX command “rm –r project” will"
strmuloption(21) = "Delete the file project from the current directory"
strmuloption(22) = "Delete all the files from the directory project"
strmuloption(23) = "Recursively delete the directory project and all its subdirectories"
strmuloption(24) = "None of the above"

strmultiple(7) = "The UNIX command “rm chap*/xyz” will delete all files "
strmuloption(25) = "Under directory chap*/xyz"
strmuloption(26) = "With name xyz from subdirectories whose names begin with chap"
strmuloption(27) = "With names starting with chap and ending with xyz"
strmuloption(28) = "None of the above"

strmultiple(8) = "The UNIX command “cc test.c” will compile a c program called test.c and"
strmuloption(29) = "Create the executable file as test.exe"
strmuloption(30) = "Create the executable file as a.out"
strmuloption(31) = "Execute it"
strmuloption(32) = "None of the above"

strmultiple(9) = "The UNIX command “cc –o cat dog.c” will compile the program dog.c and"
strmuloption(33) = "Link it with the library cat"
strmuloption(34) = "Create the executable as cat"
strmuloption(35) = "Set the output file name as cat"
strmuloption(36) = "None of the above"

strmultiple(10) = "The include directive in C is used to"
strmuloption(37) = "Link other object modules to the current program at the time of compilation"
strmuloption(38) = "Include header files in a program"
strmuloption(39) = "Specify a list of files from where data will be read at the time of execution"
strmuloption(40) = "None of the above"

strmulanswer(1) = "Displays all lines in file xyz containing the string computer"
strmulanswer(2) = "In the background"
strmulanswer(3) = "Reads a command line from the standard input and interprets it according to affixed set of rules"
strmulanswer(4) = "Cshrc"
strmulanswer(5) = "Cd ../../abc/letters"
strmulanswer(6) = "Recursively delete the directory project and all its subdirectories"
strmulanswer(7) = "With name xyz from subdirectories whose names begin with chap"
strmulanswer(8) = "Create the executable file as a.out"
strmulanswer(9) = "Create the executable as cat"
strmulanswer(10) = "Include header files in a program"
End Sub


Private Sub multiplechoice_Paper4()
strmultiple(1) = "In C, the statement out = fopen (“test.dat””,r)"
strmuloption(1) = "Opens an existing file called test.dat in read mode, and assigns the file pointer to out"
strmuloption(2) = "Creates a new file called test.dat into which data will be read in, and assigns the file pointer to out"
strmuloption(3) = "Reads the contents of the file test.dat and stores it in out "
strmuloption(4) = "None of the above"

strmultiple(2) = "The “chmod” command in UNIX"
strmuloption(5) = "Changes the current execution status from user mode the kernel mode"
strmuloption(6) = "Makes a file hidden so that it cannot be seen using the ls command"
strmuloption(7) = "Changes the access permissions of a file or directory"
strmuloption(8) = "None of the above"

strmultiple(3) = "The “chown” command in UNIX changes the "
strmuloption(9) = "Home directory of a user"
strmuloption(10) = "Access permissions of a file or directory"
strmuloption(11) = "Priority of a process"
strmuloption(12) = "None of the above"

strmultiple(4) = "The available disk space can be determined under UNIX using the command"
strmuloption(13) = "Dir"
strmuloption(14) = "Df"
strmuloption(15) = "Du"
strmuloption(16) = "File"

strmultiple(5) = "In UNIX, mounting a file system means"
strmuloption(17) = "Copying all the files from one file system to another"
strmuloption(18) = "Moving all the files from one file system to another"
strmuloption(19) = "Loading a file system from backup medium like tape"
strmuloption(20) = "Providing a link to the file system to be mounted so that it appears as a local subdirectory"

strmultiple(6) = "The correct ordering of the storage devices main memory, tape and disk in increasing order of access times is"
strmuloption(21) = "Main memory tape, disk"
strmuloption(22) = "Disk, main memory, tape"
strmuloption(23) = "Disk,tape, main memory"
strmuloption(24) = "Main memory, disk, tape"

strmultiple(7) = "The typical capacity of hard disk in a personal computer is "
strmuloption(25) = "32 Megabytes "
strmuloption(26) = "1 Gigabytes"
strmuloption(27) = "100 Gigabytes"
strmuloption(28) = "none of the above"

strmultiple(8) = "The typical storage capacity of a 3.5 inch floppy disk is"
strmuloption(29) = "360 kilobytes"
strmuloption(30) = "1.2 megabytes"
strmuloption(31) = "1.44 megabytes"
strmuloption(32) = "none of the above"

strmultiple(9) = "In terms of print quality"
strmuloption(33) = "Laser printer is better than ink jet printer, which is better than dot matrix printer"
strmuloption(34) = "Laser printer is better than dot matrix printer, which is better than ink jet printer"
strmuloption(35) = "Ink jet printer is better than laser printer, which is better than dot matrix printer"
strmuloption(36) = "None of the above"

strmultiple(10) = "Assuming that the variables A,B,C are declared as float, the associative law (A*B) * C "
strmuloption(37) = "Is always true"
strmuloption(38) = "Is not always true due to round off errors in the partial product"
strmuloption(39) = "Is not always true due to finite representation of numbers in a machine"
strmuloption(40) = "None of the above"

strmulanswer(1) = "Opens an existing file called test.dat in read mode, and assigns the file pointer to out"
strmulanswer(2) = "Changes the access permissions of a file or directory"
strmulanswer(3) = "None of the above"
strmulanswer(4) = "Df"
strmulanswer(5) = "Providing a link to the file system to be mounted so that it appears as a local subdirectory"
strmulanswer(6) = "Main memory, disk, tape"
strmulanswer(7) = "1 Gigabytes"
strmulanswer(8) = "1.44 megabytes"
strmulanswer(9) = "Laser printer is better than ink jet printer, which is better than dot matrix printer"
strmulanswer(10) = "Is not always true due to finite representation of numbers in a machine"
End Sub

Private Sub multiplechoice_Paper5()
strmultiple(1) = "The MSDOS operating system is a "
strmuloption(1) = "Time sharing operating system"
strmuloption(2) = "Single user operating system"
strmuloption(3) = "Distributed operating system"
strmuloption(4) = "None of the above"

strmultiple(2) = "An error in the program logic can be detected"
strmuloption(5) = "By the compiler"
strmuloption(6) = "By the operating system"
strmuloption(7) = "During execution of the program"
strmuloption(8) = "None of the above"

strmultiple(3) = "The shutdown procedure must be run in a UNIX system before switching off the power because"
strmuloption(9) = "Otherwise the file system may become inconsistent"
strmuloption(10) = "All users must be informed that the system is being switched off"
strmuloption(11) = "The power to the disk and display units must be shut off first"
strmuloption(12) = "None of the above"

strmultiple(4) = "What command would you use to change your password in a Unix system?"
strmuloption(13) = "Pass"
strmuloption(14) = "Chpass"
strmuloption(15) = "Passwd"
strmuloption(16) = "Set pass"

strmultiple(5) = "What command is used to find out which keys on your terminal serve the erase, kill and interrupt functions?"
strmuloption(17) = "Sty"
strmuloption(18) = "Term"
strmuloption(19) = "Print try"
strmuloption(20) = "None of the above"

strmultiple(6) = "The normal way to stop(interrupt) a process is by pressing the key"
strmuloption(21) = "CTRL/Z"
strmuloption(22) = "CTRL/D"
strmuloption(23) = "CTRL/S"
strmuloption(24) = "CTRL/C"

strmultiple(7) = "What are the three types of files in the UNIX system?"
strmuloption(25) = "Regular files, directories, device files"
strmuloption(26) = "Regular disk files, special files, device files"
strmuloption(27) = "Tape files, directories, special files"
strmuloption(28) = "None of the above"

strmultiple(8) = "Which Unix command provides the type of information contained in a particular file?"
strmuloption(29) = "More filename"
strmuloption(30) = "Type filename"
strmuloption(31) = "File filename"
strmuloption(32) = "None of the above"

strmultiple(9) = "In Unix, full pathname originates from"
strmuloption(33) = "Your home directory"
strmuloption(34) = "The root directory"
strmuloption(35) = "Your working directory"
strmuloption(36) = "None of the above"

strmultiple(10) = "Which of the following commands will give the user read, write and execute access and the world no access in the file myfile.exe?"
strmuloption(37) = "Chmod 660 myfile.exe"
strmuloption(38) = "Chmod 770 myfile.exe"
strmuloption(39) = "Chmod 600 myfile.exe"
strmuloption(40) = "None of the above"

strmulanswer(1) = "Single user operating system"
strmulanswer(2) = "During execution of the program"
strmulanswer(3) = "All users must be informed that the system is being switched off"
strmulanswer(4) = "Pass"
strmulanswer(5) = "Print try"
strmulanswer(6) = "CTRL/Z"
strmulanswer(7) = "None of the above"
strmulanswer(8) = "More filename"
strmulanswer(9) = "Your working directory"
strmulanswer(10) = "Chmod 770 myfile.exe"
End Sub

Private Sub multiplechoice_Paper6()
strmultiple(1) = "Which of the following is not a basic data type in C?"
strmuloption(1) = "Integer"
strmuloption(2) = "Complex"
strmuloption(3) = "Character"
strmuloption(4) = "None of the above"

strmultiple(2) = "Which of the following statements are true of symbolic links?"
strmuloption(5) = "Can be made to a directory"
strmuloption(6) = "The file can be accessed through the symbolic link ever after the original file has been deleted"
strmuloption(7) = "The creation of each symbolic link causes the link count to be incremented"
strmuloption(8) = "None of the above"

strmultiple(3) = "Which of the following commands would you use to copy files from a remote UNIX system to your local Unix system?"
strmuloption(9) = "Cp"
strmuloption(10) = "Finger"
strmuloption(11) = "telnet"
strmuloption(12) = "rcp"

strmultiple(4) = "Which Unix command would append a file called test to the end of a file called output?"
strmuloption(13) = "Cat output >> test"
strmuloption(14) = "Append output >> test"
strmuloption(15) = "Write output >> test"
strmuloption(16) = "Cat test >> output"

strmultiple(5) = "What does the redirection symbol “$<$” do?"
strmuloption(17) = "Redirects standard output"
strmuloption(18) = "Redirects standard input"
strmuloption(19) = "Adds output to the end of a file"
strmuloption(20) = "None of the above"

strmultiple(6) = "Several Unix commands can given on the same line by  separating them with the character"
strmuloption(21) = "“;”"
strmuloption(22) = "“,”"
strmuloption(23) = "“.”"
strmuloption(24) = "None of the above"

strmultiple(7) = "The file /usr/lib/crontab is a handy place to put files that are to be run"
strmuloption(25) = "At the highest execution priority"
strmuloption(26) = "Each time a user logs out of the system"
strmuloption(27) = "Each time a user logs into the system"
strmuloption(28) = "Periodically by the system at a specific time and date"

strmultiple(8) = "The file /etc/motd contains a message which is to be displayed on the screen "
strmuloption(29) = "Whenever the number of users logged on exceeds a certain threshold"
strmuloption(30) = "Each time a user logs out of the system"
strmuloption(31) = "Each time a user logs into the system"
strmuloption(32) = "Periodically by the system at a specific time and date"

strmultiple(9) = "The primary purpose of the pr command is to"
strmuloption(33) = "Print files to the default printer"
strmuloption(34) = "Format print pages for standard output"
strmuloption(35) = "Provide a list offprint requests in the printer queue"
strmuloption(36) = "None of the above"

strmultiple(10) = "Which command is send to remove a print job from the printer queue?"
strmuloption(37) = "lprm"
strmuloption(38) = "lpq"
strmuloption(39) = "lpstat"
strmuloption(40) = "none of the above"

strmulanswer(1) = "None of the above"
strmulanswer(2) = "The file can be accessed through the symbolic link ever after the original file has been deleted"
strmulanswer(3) = "telnet"
strmulanswer(4) = "Cat test >> output"
strmulanswer(5) = "None of the above"
strmulanswer(6) = "“,”"
strmulanswer(7) = "At the highest execution priority"
strmulanswer(8) = "Periodically by the system at a specific time and date"
strmulanswer(9) = "None of the above"
strmulanswer(10) = "lpq"
End Sub

Private Sub multiplechoice_Paper7()
strmultiple(1) = "The command to take a backup of files and directories on magnetic tape is "
strmuloption(1) = "Tar –xv"
strmuloption(2) = "Tar –tv"
strmuloption(3) = "Tar –cv"
strmuloption(4) = "None of the above"

strmultiple(2) = "The command to restore files and directories from magnetic tape is"
strmuloption(5) = "Tar –xv"
strmuloption(6) = "Tar –tv"
strmuloption(7) = "Tar –cv"
strmuloption(8) = "None of the above"

strmultiple(3) = "The command to list files and directories stored in a magnetic tape is "
strmuloption(9) = "Tar –xv"
strmuloption(10) = "Tar –tv"
strmuloption(11) = "Tar –cv"
strmuloption(12) = "None of the above"

strmultiple(4) = "The hidden files in Unix can be listed using the command "
strmuloption(13) = "Ls –lt"
strmuloption(14) = "Ls –lh"
strmuloption(15) = "Ls –lt"
strmuloption(16) = "None of the above"

strmultiple(5) = "The hidden files in Unix "
strmuloption(17) = "Have to special status bit associates with the file descriptor"
strmuloption(18) = "Have names starting with a dot"
strmuloption(19) = "Can be accessed only by the operating system kernel"
strmuloption(20) = "None of the above"

strmultiple(6) = "The grep command in Unix is used to "
strmuloption(21) = "Sort the records of a given file"
strmuloption(22) = "Search a file for the presence of a given string"
strmuloption(23) = "Globally replace a string by another"
strmuloption(24) = "None of the above"

strmultiple(7) = "The following is not a Unix shell"
strmuloption(25) = "Bourne shell"
strmuloption(26) = "C shell"
strmuloption(27) = "Korn shell"
strmuloption(28) = "Perl shell"

strmultiple(8) = "The Unix command who  |  wc  -l prints"
strmuloption(29) = "A list of all users in the system"
strmuloption(30) = "A list off all users presently logged into the system"
strmuloption(31) = "The number of users presently logged into the system"
strmuloption(32) = "None of the above"

strmultiple(9) = "The priority of a process can be increased using the "
strmuloption(33) = "Pr command"
strmuloption(34) = "Renice command"
strmuloption(35) = "Proc command"
strmuloption(36) = "None of the above"

strmultiple(10) = "The priority of a process can be increased only by"
strmuloption(37) = "The user who has started the process"
strmuloption(38) = "Any user presently logged into the system"
strmuloption(39) = "The super user"
strmuloption(40) = "None of the above"

strmulanswer(1) = "Tar –xv"
strmulanswer(2) = "Tar –cv"
strmulanswer(3) = "Tar –xv"
strmulanswer(4) = "Ls –lt"
strmulanswer(5) = "None of the above"
strmulanswer(6) = "Search a file for the presence of a given string"
strmulanswer(7) = "C shell"
strmulanswer(8) = "None of the above"
strmulanswer(9) = "Proc command"
strmulanswer(10) = "Any user presently logged into the system"
End Sub

Private Sub multiplechoice_Paper8()
strmultiple(1) = "The default shell, which is to be used when a user logs in, is defined in the"
strmuloption(1) = "/etc/ passwd file"
strmuloption(2) = "/etc/shell file"
strmuloption(3) = "startup file in the user home directory"
strmuloption(4) = "none of the above"

strmultiple(2) = "A write permission to  a directory allows a user to"
strmuloption(5) = "Edit and create files located in the directory"
strmuloption(6) = "Edit but not create files located in the directory"
strmuloption(7) = "Search a director for the presence of a file"
strmuloption(8) = "None of the above"

strmultiple(3) = "The execute permission to a directory allows a user to"
strmuloption(9) = "Execute any file located in the directory"
strmuloption(10) = "Edit any file located in he directory"
strmuloption(11) = "Search a directory for the presence of a file"
strmuloption(12) = "None of the above"

strmultiple(4) = "Information about a file is stored in"
strmuloption(13) = "A system wide file table"
strmuloption(14) = "The inode associated with the file"
strmuloption(15) = "A file allocation table within the kernel"
strmuloption(16) = "None of the above"

strmultiple(5) = "The alias command can be used in Unix to"
strmuloption(17) = "Assign the same name to two files"
strmuloption(18) = "Change the name of a file"
strmuloption(19) = "Produce a listing of the login ids of all users in the system"
strmuloption(20) = "None of the above"

strmultiple(6) = "The vi program available under Unix can be used to"
strmuloption(21) = "Crate plain ASCII files"
strmuloption(22) = "Create formatted documents with different font sizes"
strmuloption(23) = "Open a virtual terminal"
strmuloption(24) = "None of the above"

strmultiple(7) = "The Unix system calls are used to"
strmuloption(25) = "Inform the system administrator about an abnormal situation"
strmuloption(26) = "Obtain some service from the kernel"
strmuloption(27) = "Execute a Unix command from shell script"
strmuloption(28) = "None of the above"

strmultiple(8) = "The dup system call in Unix is used to"
strmuloption(29) = "Duplicate the contents of file in another file"
strmuloption(30) = "Return the descriptor of an already open file"
strmuloption(31) = "Check for duplicate filenames"
strmuloption(32) = "None of the above"

strmultiple(9) = "The command used to copy file from remote machine to local machine , in fpp is"
strmuloption(33) = "Get"
strmuloption(34) = "Copy"
strmuloption(35) = "Rep"
strmuloption(36) = "Bring"

strmultiple(10) = "In Unix the free data blocks are arrange as"
strmuloption(37) = "Linked list"
strmuloption(38) = "Array"
strmuloption(39) = "Sequentially "
strmuloption(40) = "None of the above"


strmulanswer(1) = "startup file in the user home directory"
strmulanswer(2) = "Edit and create files located in the directory"
strmulanswer(3) = "Execute any file located in the directory"
strmulanswer(4) = "The inode associated with the file"
strmulanswer(5) = "Produce a listing of the login ids of all users in the system"
strmulanswer(6) = "None of the above"
strmulanswer(7) = "Inform the system administrator about an abnormal situation"
strmulanswer(8) = "Return the descriptor of an already open file"
strmulanswer(9) = "Get"
strmulanswer(10) = "Linked list"
End Sub

Private Sub multiplechoice_Paper9()
strmultiple(1) = "The comm. – 12 f1 f2 prints lines"
strmuloption(1) = "In file f1 only"
strmuloption(2) = "Common in both files"
strmuloption(3) = "In file f2 only"
strmuloption(4) = "Not common in files."

strmultiple(2) = "The excessive movement of page back & forth between memory and disk is called"
strmuloption(5) = "Fragmentation"
strmuloption(6) = "Page fault"
strmuloption(7) = "Thrashing"
strmuloption(8) = "Flashing"

strmultiple(3) = "To find out the time taken by a command to execute, we use"
strmuloption(9) = "Texec"
strmuloption(10) = "Time"
strmuloption(11) = "Date"
strmuloption(12) = "Extime"

strmultiple(4) = "To convert Unix OS text file to dos format the command is"
strmuloption(13) = "Unixtodos"
strmuloption(14) = "Unix_to_dos"
strmuloption(15) = "Unixdos"
strmuloption(16) = "Udos"

strmultiple(5) = "Command that can initialize the terminal"
strmuloption(17) = "Tput"
strmuloption(18) = "Tini"
strmuloption(19) = "Tse"
strmuloption(20) = "ttype"

strmultiple(6) = "To find printable strings in n object file, we have to five "
strmuloption(21) = "Fstring"
strmuloption(22) = "Objstring"
strmuloption(23) = "Strings"
strmuloption(24) = "Strconf"

strmultiple(7) = "To see 1st 3 lines of a file, we give"
strmuloption(25) = "Head -3 f1"
strmuloption(26) = "Tail -3 f1"
strmuloption(27) = "Tail +3 f1"
strmuloption(28) = "Head +3 f1"

strmultiple(8) = "Command which informs the user automatically about arrival of new mail"
strmuloption(29) = "Mailcheck"
strmuloption(30) = "Notify"
strmuloption(31) = "Rmail"
strmuloption(32) = "Mailx"

strmultiple(9) = "Command used in shell to read a line of data from terminals "
strmuloption(33) = "Rline"
strmuloption(34) = "line"
strmuloption(35) = "lread"
strmuloption(36) = "none of the above"

strmultiple(10) = "The prime factor of numbers can be obtained using command"
strmuloption(37) = "Prime"
strmuloption(38) = "Prmfact"
strmuloption(39) = "Factor"
strmuloption(40) = "Factorize"

strmulanswer(1) = "Common in both files"
strmulanswer(2) = "Thrashing"
strmulanswer(3) = "Time"
strmulanswer(4) = "Unixdos"
strmulanswer(5) = "Tput"
strmulanswer(6) = "Strings"
strmulanswer(7) = "Head -3 f1"
strmulanswer(8) = "Notify"
strmulanswer(9) = "line"
strmulanswer(10) = "Factor"
End Sub

Private Sub multiplechoice_Paper10()
strmultiple(1) = "The specific command provided to make a literal copy of file system"
strmuloption(1) = "Tar"
strmuloption(2) = "Dd"
strmuloption(3) = "Volcopy"
strmuloption(4) = "Facp"

strmultiple(2) = "c1 || c2 means"
strmuloption(5) = "c2 is executed if c1 fails"
strmuloption(6) = "c2 executed if c1 is successful"
strmuloption(7) = "c1 and c2 both are executed"
strmuloption(8) = "none of the above"

strmultiple(3) = "A big file can be broken into smaller files using"
strmuloption(9) = "Break"
strmuloption(10) = "Split"
strmuloption(11) = "Divide"
strmuloption(12) = "Bifurcate"

strmultiple(4) = "For storing the current terminal session in a file"
strmuloption(13) = "Screen"
strmuloption(14) = "Termstor"
strmuloption(15) = "Session"
strmuloption(16) = "Script"

strmultiple(5) = "The command run toe remove inconsistencies in a file system"
strmuloption(17) = "Fsclear"
strmuloption(18) = "Fsck"
strmuloption(19) = "Clearfs"
strmuloption(20) = "Mkfs"

strmultiple(6) = "The default sh in Unix OS is"
strmuloption(21) = "Jsh"
strmuloption(22) = "Csh"
strmuloption(23) = "Bourne shell"
strmuloption(24) = "Ksh"

strmultiple(7) = "The problem which takes user input interprets it an takes necessary actions is"
strmuloption(25) = "Kernel"
strmuloption(26) = "System calls"
strmuloption(27) = "Sh"
strmuloption(28) = "Scheduler"

strmultiple(8) = "To change the terminal type we have to give"
strmuloption(29) = "TERM = vt100; export TERM"
strmuloption(30) = "TERM ==vt100; export TERM"
strmuloption(31) = "term = vt100;export term"
strmuloption(32) = "term  = vt100;export term"

strmultiple(9) = "On changing the permissions of a file following field is modified"
strmuloption(33) = "File modify time"
strmuloption(34) = "Incode modification time"
strmuloption(35) = "File access time"
strmuloption(36) = "None of the above"

strmultiple(10) = "Its main hob is to allocate CPU to processes"
strmuloption(37) = "Login"
strmuloption(38) = "Shell"
strmuloption(39) = "Getty"
strmuloption(40) = "Scheduler"


strmulanswer(1) = "Volcopy"
strmulanswer(2) = "c2 is executed if c1 fails"
strmulanswer(3) = "Split"
strmulanswer(4) = "Script"
strmulanswer(5) = "Fsck"
strmulanswer(6) = "Bourne shell"
strmulanswer(7) = "Sh"
strmulanswer(8) = "TERM = vt100; export TERM"
strmulanswer(9) = "Incode modification time"
strmulanswer(10) = "Scheduler"
End Sub
