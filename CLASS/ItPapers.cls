VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItPapers"
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
'
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
 
Function IT(PaperType As String, ItPaper As Paper, mCat As Category, Number As Integer)
'ptype is the paper type
'itpaper is the number of paper as paper1,paper2 ...
'mcat is the category
'number is the number of questions, or options or answers to be taken

 StrPaperType = PaperType
' StrPaper = Paper
If PaperType = "Fill in the Blanks" Then
   If ItPaper = Paper1 Then Fill_in_the_Blanks_Paper1
   If ItPaper = Paper2 Then Fill_in_the_Blanks_Paper2
   If ItPaper = Paper3 Then Fill_in_the_Blanks_Paper3
   If ItPaper = Paper4 Then Fill_in_the_Blanks_Paper4
   If ItPaper = Paper5 Then Fill_in_the_Blanks_Paper5
   If ItPaper = Paper6 Then Fill_in_the_Blanks_Paper6
   If ItPaper = Paper7 Then Fill_in_the_Blanks_Paper7
   If ItPaper = Paper8 Then Fill_in_the_Blanks_Paper8
   If ItPaper = Paper9 Then Fill_in_the_Blanks_Paper9
   If ItPaper = Paper10 Then Fill_in_the_Blanks_Paper10
   If ItPaper = Paper11 Then Fill_in_the_Blanks_Paper11
   If ItPaper = Paper12 Then Fill_in_the_Blanks_Paper12
   If ItPaper = Paper13 Then Fill_in_the_Blanks_Paper13
   If ItPaper = Paper14 Then Fill_in_the_Blanks_Paper14
   If ItPaper = Paper15 Then Fill_in_the_Blanks_Paper15
   
ElseIf PaperType = "Matching Columns" Then
       If ItPaper = Paper1 Then Matching_Columns_Paper1
       If ItPaper = Paper2 Then Matching_Columns_Paper2
       If ItPaper = Paper3 Then Matching_Columns_Paper3
       If ItPaper = Paper4 Then Matching_Columns_Paper4
       If ItPaper = Paper5 Then Matching_Columns_Paper5
       If ItPaper = Paper6 Then Matching_Columns_Paper6
       If ItPaper = Paper7 Then Matching_Columns_Paper7
       If ItPaper = Paper8 Then Matching_Columns_Paper8
       If ItPaper = Paper9 Then Matching_Columns_Paper9
       If ItPaper = Paper10 Then Matching_Columns_Paper10
       If ItPaper = Paper11 Then Matching_Columns_Paper11
       If ItPaper = Paper12 Then Matching_Columns_Paper12
       If ItPaper = Paper13 Then Matching_Columns_Paper13
       If ItPaper = Paper14 Then Matching_Columns_Paper14
       If ItPaper = Paper15 Then Matching_Columns_Paper15
       
ElseIf PaperType = "True Or False" Then
       If ItPaper = Paper1 Then TrueorFalse_Paper1
       If ItPaper = Paper2 Then TrueorFalse_Paper2
       If ItPaper = Paper3 Then TrueorFalse_Paper3
       If ItPaper = Paper4 Then TrueorFalse_Paper4
       If ItPaper = Paper5 Then TrueorFalse_Paper5
       If ItPaper = Paper6 Then TrueorFalse_Paper6
       If ItPaper = Paper7 Then TrueorFalse_Paper7
       If ItPaper = Paper8 Then TrueorFalse_Paper8
       If ItPaper = Paper9 Then TrueorFalse_Paper9
       If ItPaper = Paper10 Then TrueorFalse_Paper10
       If ItPaper = Paper11 Then TrueorFalse_Paper11
       If ItPaper = Paper12 Then TrueorFalse_Paper12
       If ItPaper = Paper13 Then TrueorFalse_Paper13
       If ItPaper = Paper14 Then TrueorFalse_Paper14
       If ItPaper = Paper15 Then TrueorFalse_Paper15
       
ElseIf PaperType = "Descriptive Questions" Then
       If ItPaper = Paper1 Then questions_paper1
       If ItPaper = Paper2 Then questions_paper2
       If ItPaper = Paper3 Then questions_paper3
       If ItPaper = Paper4 Then questions_paper4
       If ItPaper = Paper5 Then questions_paper5
       If ItPaper = Paper6 Then questions_paper6
       If ItPaper = Paper7 Then questions_paper7
       If ItPaper = Paper8 Then questions_paper8
       If ItPaper = Paper9 Then questions_paper9
       If ItPaper = Paper10 Then questions_paper10
       If ItPaper = Paper11 Then questions_paper11
       If ItPaper = Paper12 Then questions_paper12
       If ItPaper = Paper13 Then questions_paper13
       If ItPaper = Paper14 Then questions_paper14
       If ItPaper = Paper15 Then questions_paper15

ElseIf PaperType = "Multiple Choice" Then
       If ItPaper = Paper1 Then multiplechoice_Paper1
       If ItPaper = Paper2 Then multiplechoice_Paper2
       If ItPaper = Paper3 Then multiplechoice_Paper3
       If ItPaper = Paper4 Then multiplechoice_Paper4
       If ItPaper = Paper5 Then multiplechoice_Paper5
       If ItPaper = Paper6 Then multiplechoice_Paper6
       If ItPaper = Paper7 Then multiplechoice_Paper7
       If ItPaper = Paper8 Then multiplechoice_Paper8
       If ItPaper = Paper9 Then multiplechoice_Paper9
       If ItPaper = Paper10 Then multiplechoice_Paper10
       If ItPaper = Paper11 Then multiplechoice_Paper11
       If ItPaper = Paper12 Then multiplechoice_Paper12
       If ItPaper = Paper13 Then multiplechoice_Paper13
       If ItPaper = Paper14 Then multiplechoice_Paper14
       If ItPaper = Paper15 Then multiplechoice_Paper15
End If

If mCat = Fill_Questions Then
   IT = StrQuestions(Number)
ElseIf mCat = Fill_Options Then
       IT = stroptions(Number)
ElseIf mCat = Fill_Answers Then
       IT = stranswers(Number)
       
ElseIf mCat = Matching_ColumnA Then
       IT = strcolumnA(Number)
ElseIf mCat = Matching_ColumnB Then
       IT = strcolumnB(Number)
ElseIf mCat = Matching_Answers Then
       IT = stranswers(Number)
       
ElseIf mCat = Questions Then
       IT = StrQuestions(Number)
       
ElseIf mCat = TF_Questions Then
       IT = StrQuestions(Number)
       
ElseIf mCat = TF_Answers Then
       IT = stranswers(Number)
       
ElseIf mCat = MC_Questions Then
       IT = strmultiple(Number)
       
ElseIf mCat = MC_Options_1 Then
       IT = strmuloption(Number)
       
ElseIf mCat = MC_Options_2 Then
       IT = strmuloption(Number)
       
ElseIf mCat = MC_Options_3 Then
       IT = strmuloption(Number)
       
ElseIf mCat = MC_Options_4 Then
       IT = strmuloption(Number)
       
ElseIf mCat = MC_Answer Then
       IT = strmulanswer(Number)
       
End If

End Function

Private Sub Fill_in_the_Blanks_Paper1()
StrQuestions(1) = "Information is dash data."
StrQuestions(2) = "dash are raw facts."
StrQuestions(3) = "Highly dash decisions can be made by computer programs."
StrQuestions(4) = "Information needed by managers should be timely and dash."
StrQuestions(5) = "Middle level managers use dash information."
StrQuestions(6) = "Balance sheets, profits and loss accounts and future sales forecast form an integral part of the dash information."
StrQuestions(7) = "Each attribute of an entity set has a dash of values."
StrQuestions(8) = "A dash is a group of related data."
StrQuestions(9) = "Floppy disk is an example of dash access of storage."
StrQuestions(10) = "The density of recording of Winchester disk is dash than floppy disk."

stroptions(1) = "Direct "
stroptions(2) = "Range"
stroptions(3) = "Record"
stroptions(4) = "Tactical"
stroptions(5) = "Structured"
stroptions(6) = "Complete"
stroptions(7) = "Corporate"
stroptions(8) = "Data"
stroptions(9) = "Processed"
stroptions(10) = "Higher"

stranswers(1) = "Processed"
stranswers(2) = "Data"
stranswers(3) = "Structured"
stranswers(4) = "Complete"
stranswers(5) = "Tactical"
stranswers(6) = "Corporate"
stranswers(7) = "Range"
stranswers(8) = "Record"
stranswers(9) = "Direct"
stranswers(10) = "Higher"

End Sub

Private Sub Fill_in_the_Blanks_Paper2()
StrQuestions(1) = "Direct or Random Access files are most suitable for dash inquiry processing."
StrQuestions(2) = "Real Time processing is an on line processing with dash response time."
StrQuestions(3) = "In dash processing, processing is performed in different locations without any connections between them."
StrQuestions(4) = "dash is also accomplished in special purpose computers with array processors."
StrQuestions(5) = "dash means organizing similar items of data into groups of similar type."
StrQuestions(6) = "In dash testing the system, thoroughly tested modules are brought together and tested as a whole."
StrQuestions(7) = "Putting the group of words in middle of left and right margins is known as dash."
StrQuestions(8) = "Mouse is graphics dash input device."
StrQuestions(9) = "The backbone of E - mail system is dash network."
StrQuestions(10) = "A dash service bureau is an organization which provides data processing anytime sharing service to its customers."

stroptions(1) = "testing"
stroptions(2) = "Multiprocessing"
stroptions(3) = "Classifying"
stroptions(4) = "Centering"
stroptions(5) = "Input"
stroptions(6) = "Very Fast"
stroptions(7) = "Decentralized"
stroptions(8) = "On-Line"
stroptions(9) = "Service"
stroptions(10) = "Communication"

stranswers(1) = "On-Line"
stranswers(2) = "Very Fast"
stranswers(3) = "Decentralized"
stranswers(4) = "Multiprocessing"
stranswers(5) = "Classifying"
stranswers(6) = "testing"
stranswers(7) = "Centering"
stranswers(8) = "Input"
stranswers(9) = "Communication"
stranswers(10) = "Service"

End Sub
Private Sub Fill_in_the_Blanks_Paper3()
StrQuestions(1) = "dash is general purpose machine which manipulate raw facts according to a set of instructions."
StrQuestions(2) = "A computer can beaded to operate continuously dash taking any rest."
StrQuestions(3) = "Final results of processing, before releasing, to output unit are held in dash unit."
StrQuestions(4) = "Input interface converts the instructions and data in computer dash acceptable form."
StrQuestions(5) = "dash converts computer produced results into human acceptable form."
StrQuestions(6) = "Before utilizing the resource, the Operating System must know its dash."
StrQuestions(7) = "MS - DOS is classified as dash user Operating System."
StrQuestions(8) = "An Operating System provides its users a convenient interface in a dash view."
StrQuestions(9) = "A window utility located in Accessories is represented by dash."
StrQuestions(10) = "dash performs basic functions of Operating Systems. "

stroptions(1) = "Without"
stroptions(2) = "Storage"
stroptions(3) = "Acceptable"
stroptions(4) = "Output Interface"
stroptions(5) = "Status"
stroptions(6) = "Top Down"
stroptions(7) = "Computer"
stroptions(8) = "Kernel"
stroptions(9) = "Single"
stroptions(10) = "An icon"

stranswers(1) = "Computer"
stranswers(2) = "Without"
stranswers(3) = "Storage"
stranswers(4) = "Acceptable"
stranswers(5) = "Output Interface"
stranswers(6) = "Status"
stranswers(7) = "Single"
stranswers(8) = "Top Down"
stranswers(9) = "An icon"
stranswers(10) = "Kernel"

End Sub
Private Sub Fill_in_the_Blanks_Paper4()
StrQuestions(1) = "Time Sharing Operating Systems were evolved to give rapid dash to every user."
StrQuestions(2) = "Paging is used to implement dash memory."
StrQuestions(3) = "A printer is dash device."
StrQuestions(4) = "dash access, as file operation, is commonly provided where that makes sense."
StrQuestions(5) = "If an executable program is not found in an active directory, DOS searches the directories provided by dash command."
StrQuestions(6) = "In DOS, root directory may contain only dash number of files and subdirectory entries."
StrQuestions(7) = "In dos the command dash may copy all files in source directory and its subdirectories."
StrQuestions(8) = "Attrib command in DOS with r is used to dash files."
StrQuestions(9) = "REN is a dash command."
StrQuestions(10) = "Just deleted file can be dash in dos."

stroptions(1) = "virtual"
stroptions(2) = "Turn around"
stroptions(3) = "Path"
stroptions(4) = "Limited"
stroptions(5) = "Recovered"
stroptions(6) = "Character"
stroptions(7) = "Random"
stroptions(8) = "XCopy"
stroptions(9) = "Protect"
stroptions(10) = "Internal"

stranswers(1) = "Turn around"
stranswers(2) = "virtual"
stranswers(3) = "Character"
stranswers(4) = "Random"
stranswers(5) = "Path"
stranswers(6) = "Limited"
stranswers(7) = "XCopy"
stranswers(8) = "Protect"
stranswers(9) = "Internal"
stranswers(10) = "Recovered"

End Sub

Private Sub Fill_in_the_Blanks_Paper5()
StrQuestions(1) = "File permission of current directory can be known by using ls command with dash option."
StrQuestions(2) = "To know the home directory of the user in UNIX we may give dash command followed by path."
StrQuestions(3) = "[ls directory] will not list files if [directory] does not have dash permission for the user."
StrQuestions(4) = "[cd directory] will not change to [directory] if the current directory does not have dash permission o the user."
StrQuestions(5) = "dash means commands, to be performed are displayed on screen."
StrQuestions(6) = "dash is that portion of memory which WINDOWS uses to store a temporary copied image or text."
StrQuestions(7) = "dash editor creates and alters file used by WINDOWS to run an- WINDOW applications."
StrQuestions(8) = "WINDOWS is dash over DOS environment."
StrQuestions(9) = "All icons are located on dash work area in group icon."
StrQuestions(10) = "dash icons are inactive WINDOWS whose applications are still loaded in memory."

stroptions(1) = "Execute"
stroptions(2) = "GUI"
stroptions(3) = "Loaded"
stroptions(4) = "Clipboard"
stroptions(5) = "d"
stroptions(6) = "cd"
stroptions(7) = "Read"
stroptions(8) = "PIF"
stroptions(9) = "Shell"
stroptions(10) = "Program"

stranswers(1) = "d"
stranswers(2) = "cd"
stranswers(3) = "Read"
stranswers(4) = "Execute"
stranswers(5) = "GUI"
stranswers(6) = "Clipboard"
stranswers(7) = "PIF"
stranswers(8) = "Shell"
stranswers(9) = "Program"
stranswers(10) = "Loaded"

End Sub

Private Sub Fill_in_the_Blanks_Paper6()
StrQuestions(1) = "dash consists of heading for all down menus."
StrQuestions(2) = "Dialog box that appears on a pull down menu uses items that have dash."
StrQuestions(3) = "The speed of communication between terminals and computer is dash in time sharing system."
StrQuestions(4) = "dash communication is required between computers in case of LAN."
StrQuestions(5) = "In LAN, computers can be connected in dash way."
StrQuestions(6) = "dash of files is one of the common activities in LAN."
StrQuestions(7) = "In case of computer networks, the rate of data transfer between computers is dependent of __________ of public communication lines."
StrQuestions(8) = "To avoid dash incompatibility with different computers standard organizations provide rules and conventions."
StrQuestions(9) = "dash ISO is a world wide group associated with the development of international standards for data communication."
StrQuestions(10) = "Upon 9600 bps data transmission speed can be achieved by using dash copper wires within shorter distance of 100 Meters."

stroptions(1) = "Slow"
stroptions(2) = "Fast"
stroptions(3) = "More than one"
stroptions(4) = "Menu bar"
stroptions(5) = "Ellipses"
stroptions(6) = "ISO"
stroptions(7) = "Twisted"
stroptions(8) = "Sharing"
stroptions(9) = "Speed"
stroptions(10) = "Incompatibility"

stranswers(1) = "Menu bar"
stranswers(2) = "Ellipses"
stranswers(3) = "Slow"
stranswers(4) = "Fast"
stranswers(5) = "More than one"
stranswers(6) = "Sharing"
stranswers(7) = "Speed"
stranswers(8) = "Incompatibility"
stranswers(9) = "ISO"
stranswers(10) = "Twisted"

End Sub

Private Sub Fill_in_the_Blanks_Paper7()
StrQuestions(1) = "Data transmission in coaxial cable can be dash than twisted pair of wire."
StrQuestions(2) = "Microwave transmission is dash method of communication."
StrQuestions(3) = "dash are used to strengthen weak signals in microwave transmission."
StrQuestions(4) = "A communication satellite is a dash relay station at 36000 km above the Equator."
StrQuestions(5) = "In dash medium, converted electrical signals, in for of light are used."
StrQuestions(6) = "Communication between one computer and another within 100 ft. may be carried out using __________ communication channel."
StrQuestions(7) = "Conversion of analog signal to digital signal is known as dash"
StrQuestions(8) = "Bandwidth indicates the dash of a channel."
StrQuestions(9) = "dash of a transmission channel is expressed in Bauds."
StrQuestions(10) = "Computer generated data is dash."

stroptions(1) = "Line of sight"
stroptions(2) = "Repeaters"
stroptions(3) = "Optical fiber"
stroptions(4) = "Faster"
stroptions(5) = "Signaling rate"
stroptions(6) = "Digital"
stroptions(7) = "Microwave"
stroptions(8) = "Parallel"
stroptions(9) = "Demodulation "
stroptions(10) = "Capacity"

stranswers(1) = "Faster"
stranswers(2) = "Line of sight"
stranswers(3) = "Repeaters"
stranswers(4) = "Microwave"
stranswers(5) = "Optical fiber"
stranswers(6) = "Parallel"
stranswers(7) = "Demodulation"
stranswers(8) = "Capacity"
stranswers(9) = "Signaling rate"
stranswers(10) = "Digital"

End Sub

Private Sub Fill_in_the_Blanks_Paper8()
StrQuestions(1) = "Telephone lines in India carry dash"
StrQuestions(2) = "Conversion of digital signal to analog signal is known as dash"
StrQuestions(3) = "Data are transferred over distances more than hundred feet using dash communication."
StrQuestions(4) = "The process of modulation & demodulation is carried out by dash."
StrQuestions(5) = "dash is a special type of modem used in portable terminals to communicate with distant CPU."
StrQuestions(6) = "PC/AT uses dash bit data bus containing direction."
StrQuestions(7) = "dash is dash which switches direction."
StrQuestions(8) = "Echo of character through dash is called Local Echo."
StrQuestions(9) = "Synchronous transfer is characterized by dash data transmission."
StrQuestions(10) = "In dash transmission, ach data character is preceded and followed by one or more bit signals as start & stop of character."

stroptions(1) = "Modulation"
stroptions(2) = "Serial"
stroptions(3) = "Software"
stroptions(4) = "Acoustic Coupler"
stroptions(5) = "Modem"
stroptions(6) = "16"
stroptions(7) = "Half Duplex"
stroptions(8) = "Continues"
stroptions(9) = "Asynchronous"
stroptions(10) = "Analog"

stranswers(1) = "Analog"
stranswers(2) = "Modulation"
stranswers(3) = "Serial"
stranswers(4) = "Modem"
stranswers(5) = "Acoustic Coupler"
stranswers(6) = "16"
stranswers(7) = "Half Duplex"
stranswers(8) = "Software"
stranswers(9) = "Continues"
stranswers(10) = "Asynchronous"

End Sub
Private Sub Fill_in_the_Blanks_Paper9()
StrQuestions(1) = "dash bits frame each transmitted byte in asynchronous transmission."
StrQuestions(2) = "Synchronous data are transmitted in dash oriented protocol which recognizes bytes of data as specific character."
StrQuestions(3) = "Baud rate is determined by communication hardware and can be selected by communication dash."
StrQuestions(4) = "Asynchronous transfer is characterized by dash data transmission."
StrQuestions(5) = "The purpose of synchronization into recognize dash pattern assent by the transmitting station."
StrQuestions(6) = "dash is a technique used to distinguish flag but pattern in an information field."
StrQuestions(7) = "Several low speed communication lines are brought together as dash channels in a multiplexer."
StrQuestions(8) = "The best example of dash is the way we receive various stations on a radio set."
StrQuestions(9) = "Commands in menu bar affect dash subordinate Window."
StrQuestions(10) = "dash oriented protocol recognizes bytes of data as specific character."

stroptions(1) = "Software"
stroptions(2) = "Discontinuous"
stroptions(3) = "Bit"
stroptions(4) = "High speed"
stroptions(5) = "Bit stuffing"
stroptions(6) = "Start and stop"
stroptions(7) = "Block"
stroptions(8) = "Character"
stroptions(9) = "FDM "
stroptions(10) = "Active"

stranswers(1) = "Start and stop"
stranswers(2) = "Block"
stranswers(3) = "Software"
stranswers(4) = "Discontinuous"
stranswers(5) = "Bit"
stranswers(6) = "Bit stuffing"
stranswers(7) = "High speed"
stranswers(8) = "FDM"
stranswers(9) = "Active"
stranswers(10) = "Character"

End Sub
Private Sub Fill_in_the_Blanks_Paper10()
StrQuestions(1) = "The channel capacity is fully utilized in TDM by dash number of data streams of different users."
StrQuestions(2) = "Gateways are communication dash"
StrQuestions(3) = "Database software packages dash operate differently."
StrQuestions(4) = "In base band network systems dash signal conversion device is required."
StrQuestions(5) = "In broadband LAN the channel width varies to the desired dash."
StrQuestions(6) = "dash topologies are used in token ring."
StrQuestions(7) = "Arcnet is base band, dash passing network."
StrQuestions(8) = "Distributed bus networks are well suited for dash."
StrQuestions(9) = "In India ISDN provides dash voice and one data circuits"
StrQuestions(10) = "dash processing concept is an extension of shared device processing."

stroptions(1) = "Server"
stroptions(2) = "Will not"
stroptions(3) = "Star and ring"
stroptions(4) = "No"
stroptions(5) = "Interleaving"
stroptions(6) = "Two"
stroptions(7) = "Client/Server"
stroptions(8) = "Data rate"
stroptions(9) = "Token"
stroptions(10) = "Linear"

stranswers(1) = "Interleaving"
stranswers(2) = "Server"
stranswers(3) = "will not"
stranswers(4) = "no"
stranswers(5) = "Data Rate"
stranswers(6) = "Star and Ring"
stranswers(7) = "Token"
stranswers(8) = "Linear"
stranswers(9) = "Two"
stranswers(10) = "Client/Server"

End Sub
Private Sub Fill_in_the_Blanks_Paper11()
StrQuestions(1) = "Internet host computer uses dash to talk to the rest of internet."
StrQuestions(2) = "dash are designed to process compels scientific applications."
StrQuestions(3) = "Entire main memory of supercomputers consists of high speed dash memory."
StrQuestions(4) = "Most popular desktop machines are made by dash."
StrQuestions(5) = "Power PC 601 is a dash owed by Apple Macintosh machine."
StrQuestions(6) = "Workstations normally use dash processor."
StrQuestions(7) = "AIX (IBM), solaries(SUN) and  HP-UNIX(HP) are the derivations of dash"
StrQuestions(8) = "In Client server computing several specialized services are provided by dash"
StrQuestions(9) = "A supercomputer's main storage is given in dash bytes."
StrQuestions(10) = "Parallel processing is achieved in uniprocessor using dash."

stroptions(1) = "Cache"
stroptions(2) = "IBM"
stroptions(3) = "Processor"
stroptions(4) = "Server"
stroptions(5) = "RISC"
stroptions(6) = "UNIX"
stroptions(7) = "Giga"
stroptions(8) = "Supercomputer"
stroptions(9) = "TCP/IP"
stroptions(10) = "Pipelining"

stranswers(1) = "TCP/IP"
stranswers(2) = "Supercomputer"
stranswers(3) = "Cache"
stranswers(4) = "IBM"
stranswers(5) = "Processor"
stranswers(6) = "RISC"
stranswers(7) = "UNIX"
stranswers(8) = "Server"
stranswers(9) = "Giga"
stranswers(10) = "Pipelining"

End Sub

Private Sub Fill_in_the_Blanks_Paper12()
StrQuestions(1) = "In assembly language dash are used to represent operation code."
StrQuestions(2) = "The input to dash is called source program."
StrQuestions(3) = "Assembly language programs are dash from one machine to another machine of other model."
StrQuestions(4) = "dash is automatically stored in memory by hard water of a computer when the computer is switched on."
StrQuestions(5) = "In high level language dash correct statement need not be semantically meaningful."
StrQuestions(6) = "Translation of one high level language statement to many machine language statements is dash translation."
StrQuestions(7) = "dash translates whole program into equivalent machine language program before starting execution."
StrQuestions(8) = "An interpreter translates dash instruction at a time into machine language & executes it."
StrQuestions(9) = "A high level language is used to specify dash."
StrQuestions(10) = "Execution of instructions takes place at dash."

stroptions(1) = "One to many"
stroptions(2) = "Not Portable"
stroptions(3) = "PL"
stroptions(4) = "Compiler"
stroptions(5) = "Syntactically"
stroptions(6) = "One"
stroptions(7) = "How to do"
stroptions(8) = "Mnemonic"
stroptions(9) = "Assembler"
stroptions(10) = "ALU"

stranswers(1) = "Mnemonic"
stranswers(2) = "Assembler"
stranswers(3) = "Not Portable"
stranswers(4) = "PL"
stranswers(5) = "Syntactically"
stranswers(6) = "One to many"
stranswers(7) = "Compiler"
stranswers(8) = "One"
stranswers(9) = "How to do"
stranswers(10) = "ALU"

End Sub

Private Sub Fill_in_the_Blanks_Paper13()
StrQuestions(1) = "A control unit acts as dash for all components of a computer."
StrQuestions(2) = "The CPU is dash of a computer system."
StrQuestions(3) = "A general purpose dash holds data and intermediate result during the execution of a program."
StrQuestions(4) = "A stack registers one of the dash purpose registers."
StrQuestions(5) = "The function of the memory is dash information."
StrQuestions(6) = "The dash memory is directly addressed by the CPU."
StrQuestions(7) = "The hard disk is the example of dash memory."
StrQuestions(8) = "Symbols 0 and 1 represent dash states of components of a computer."
StrQuestions(9) = "To store a printable character dash byte is used."
StrQuestions(10) = "A machine with 16 - bit word will be dash than 32 - bit word machine."

stroptions(1) = "Register"
stroptions(2) = "Binary"
stroptions(3) = "One"
stroptions(4) = "To store "
stroptions(5) = "Main"
stroptions(6) = "Secondary"
stroptions(7) = "Special"
stroptions(8) = "Slower"
stroptions(9) = "Central Nervous System"
stroptions(10) = "Brain"

stranswers(1) = "Central Nervous System"
stranswers(2) = "Brain"
stranswers(3) = "Register"
stranswers(4) = "Special"
stranswers(5) = "To store"
stranswers(6) = "Main"
stranswers(7) = "Secondary"
stranswers(8) = "Binary"
stranswers(9) = "One"
stranswers(10) = "Slower"

End Sub

Private Sub Fill_in_the_Blanks_Paper14()
StrQuestions(1) = "Semi conductor memories are of two types’ dash and RAM."
StrQuestions(2) = "RAM memory is dash."
StrQuestions(3) = "Data and instructions are entered into a computer through dash."
StrQuestions(4) = "Virus is dash surviving by capturing interrupt."
StrQuestions(5) = "A smart terminal is more dash than a dumb interrupt"
StrQuestions(6) = "Magnetic ink character is one type of dash"
StrQuestions(7) = "dash detects alphanumeric characters printed on paper."
StrQuestions(8) = "dash printers are classified as non impact printers."
StrQuestions(9) = "Multiple copies can be had through dash printers."
StrQuestions(10) = "Some plotters employ dash in place of ink pens."

stroptions(1) = "Expensive"
stroptions(2) = "Input"
stroptions(3) = "TSR"
stroptions(4) = "Impact"
stroptions(5) = "Ink Jet"
stroptions(6) = "Scanner"
stroptions(7) = "OCR"
stroptions(8) = "Laser"
stroptions(9) = "ROM"
stroptions(10) = "Volatile"

stranswers(1) = "ROM"
stranswers(2) = "Volatile"
stranswers(3) = "Input"
stranswers(4) = "TSR"
stranswers(5) = "Expensive"
stranswers(6) = "Scanner"
stranswers(7) = "OCR"
stranswers(8) = "Laser"
stranswers(9) = "Impact"
stranswers(10) = "Ink Jet"

End Sub

Private Sub Fill_in_the_Blanks_Paper15()
StrQuestions(1) = "The data transfer rate for magnetic tape is the product of tape dash and tape speed."
StrQuestions(2) = "Indivisible concentric circles on disk are called dash."
StrQuestions(3) = "Tracks are further subdivided into dash."
StrQuestions(4) = "CD-ROM is read-through use of dash"
StrQuestions(5) = "Both data integrity and system dash are headed to create an oersib's right to privacy."
StrQuestions(6) = "Netware is dash"
StrQuestions(7) = "Netware services are totally dash of the network media."
StrQuestions(8) = "DOS command CD A: will show he pathname of dash if drive."
StrQuestions(9) = "Primary function of the dash is to execute the program."
StrQuestions(10) = "WINDOWS is a dash under DOS."

stroptions(1) = "Sectors"
stroptions(2) = "Security"
stroptions(3) = "Active directory"
stroptions(4) = "CPU"
stroptions(5) = "Photo diode"
stroptions(6) = "NOS"
stroptions(7) = "Independent"
stroptions(8) = "Environment"
stroptions(9) = "Density"
stroptions(10) = "Tracks"

stranswers(1) = "Density"
stranswers(2) = "Tracks"
stranswers(3) = "Sectors"
stranswers(4) = "Photo diode"
stranswers(5) = "Security"
stranswers(6) = "NOS"
stranswers(7) = "Independent"
stranswers(8) = "Active directory"
stranswers(9) = "CPU"
stranswers(10) = "Environment"

End Sub

Private Sub Matching_Columns_Paper1()
strcolumnA(1) = "Processed data"
strcolumnA(2) = "Raw Facts"
strcolumnA(3) = "Time critical"
strcolumnA(4) = "Brevity"
strcolumnA(5) = "Strategic information"
strcolumnA(6) = "Variation of currency exchange rate"
strcolumnA(7) = "Smallest pirce of meaningful information of a file"
strcolumnA(8) = "Collection of logically related records"
strcolumnA(9) = "Sequential storage device"
strcolumnA(10) = "Empty portion of the tape between two records"

strcolumnB(1) = "Data"
strcolumnB(2) = "On line system"
strcolumnB(3) = "Qualities of information"
strcolumnB(4) = "Magnetic tape"
strcolumnB(5) = "IRG"
strcolumnB(6) = "Data item"
strcolumnB(7) = "Top level manager"
strcolumnB(8) = "File"
strcolumnB(9) = "Information"
strcolumnB(10) = "International information"

stranswers(1) = "Information"
stranswers(2) = "Data"
stranswers(3) = "On line system"
stranswers(4) = "Qualities of information"
stranswers(5) = "International information"
stranswers(6) = "Top level manager"
stranswers(7) = "IRG"
stranswers(8) = "File"
stranswers(9) = "Magnetic tape"
stranswers(10) = "Data item"

End Sub

Private Sub Matching_Columns_Paper2()
strcolumnA(1) = "Arrangement of records in file"
strcolumnA(2) = "Addition and deletion rate of records in file"
strcolumnA(3) = "Manipulation of facts"
strcolumnA(4) = "Processing has fixed time constraints"
strcolumnA(5) = "System of computers connected together by communication networks"
strcolumnA(6) = "Orgaizing similar itmes of data into groups"
strcolumnA(7) = "Prototype model"
strcolumnA(8) = "Breaking of system down into smaller components"
strcolumnA(9) = "Modification of existing system"
strcolumnA(10) = "To write the word on the next line automatically when end of righr margin is reached"

strcolumnB(1) = "Problem definition stage"
strcolumnB(2) = "File organization"
strcolumnB(3) = "Word wrap"
strcolumnB(4) = "File volatility"
strcolumnB(5) = "System maintenance"
strcolumnB(6) = "Distributed processing"
strcolumnB(7) = "Classifying"
strcolumnB(8) = "Real Time processing"
strcolumnB(9) = "Top down methodology"
strcolumnB(10) = "Data processing"

stranswers(1) = "File organization"
stranswers(2) = "File volatility"
stranswers(3) = "Data processing"
stranswers(4) = "Real Time processing"
stranswers(5) = "Distributed processing"
stranswers(6) = "Classifying"
stranswers(7) = "Problem definition stage"
stranswers(8) = "Top down methodology"
stranswers(9) = "System maintenance"
stranswers(10) = "Word wrap"

End Sub

Private Sub Matching_Columns_Paper3()
strcolumnA(1) = "Expanding shorter lines by adding a few spaces between word in the line"
strcolumnA(2) = "Creation and processing of picture images"
strcolumnA(3) = "Freelance plus"
strcolumnA(4) = "Transmission of message, memos over communicaions neworks"
strcolumnA(5) = "Communications of printed pages bewee remote locations"
strcolumnA(6) = "Integration of word processing, E mail , graphics and desktop publishing"
strcolumnA(7) = "Set of instructions performing a particular task"
strcolumnA(8) = "Electronic data processing machine meant for individual use"
strcolumnA(9) = "Processes information of physical nature"
strcolumnA(10) = "Measurment of speed of computers"

strcolumnB(1) = "Graphical package"
strcolumnB(2) = "E - mail"
strcolumnB(3) = "Graphics"
strcolumnB(4) = "Office automation"
strcolumnB(5) = "Personal computer"
strcolumnB(6) = "MIPS"
strcolumnB(7) = "Program"
strcolumnB(8) = "Analog computer"
strcolumnB(9) = "Justification"
strcolumnB(10) = "Fax"

stranswers(1) = "Justification"
stranswers(2) = "Graphics"
stranswers(3) = "Graphical package"
stranswers(4) = "E - mail"
stranswers(5) = "Fax"
stranswers(6) = "Office automation"
stranswers(7) = "Program"
stranswers(8) = "Personal computer"
stranswers(9) = "Analog computer"
stranswers(10) = "MIPS"

End Sub

Private Sub Matching_Columns_Paper4()
strcolumnA(1) = "Transmission of data into binary codes"
strcolumnA(2) = "Linking computer with external environment"
strcolumnA(3) = "Allocation of resources"
strcolumnA(4) = "Single user Operating System"
strcolumnA(5) = "The part of Operating system respossible for loading the module when needed"
strcolumnA(6) = "Preventing execution of program forever"
strcolumnA(7) = "Transforming programs from main memory to disk storage and back"
strcolumnA(8) = "Device dependent software"
strcolumnA(9) = "Used to recognize file internally UNIX"
strcolumnA(10) = "Connection between one's directory and shared folder of another directory"

strcolumnB(1) = "Output unit"
strcolumnB(2) = "MSDOS"
strcolumnB(3) = "Input interface"
strcolumnB(4) = "Kernel"
strcolumnB(5) = "CPU protection"
strcolumnB(6) = "Swapping"
strcolumnB(7) = "i- node"
strcolumnB(8) = "Link"
strcolumnB(9) = "Device driver"
strcolumnB(10) = "Operating System"

stranswers(1) = "Input interface"
stranswers(2) = "Output unit"
stranswers(3) = "Link"
stranswers(4) = "MSDOS"
stranswers(5) = "Kernel"
stranswers(6) = "CPU protection"
stranswers(7) = "Swapping"
stranswers(8) = "Operating System"
stranswers(9) = "i- node"
stranswers(10) = "Device driver"

End Sub

Private Sub Matching_Columns_Paper5()
strcolumnA(1) = "CNTL+ALT+DEL"
strcolumnA(2) = "UNIX"
strcolumnA(3) = "Protection of file in DOS"
strcolumnA(4) = "Files not listed under DIR command"
strcolumnA(5) = "One kind of protection in UNIX (not file access permission)"
strcolumnA(6) = "Regulating use offile among users in UNIX"
strcolumnA(7) = "Can open any file in UNIX"
strcolumnA(8) = "Command to become superuser in UNIX"
strcolumnA(9) = "Operator interacts with computer"
strcolumnA(10) = "Unique identification of record"

strcolumnB(1) = "Case sensitive"
strcolumnB(2) = "Super user"
strcolumnB(3) = "Hidden files"
strcolumnB(4) = "Attrib "
strcolumnB(5) = "Console"
strcolumnB(6) = "Primary key"
strcolumnB(7) = "File acces permission"
strcolumnB(8) = "Login password"
strcolumnB(9) = "Su"
strcolumnB(10) = "WARM Booting  "

stranswers(1) = "WARM Booting"
stranswers(2) = "Case sensitive"
stranswers(3) = "Attrib"
stranswers(4) = "Hidden files"
stranswers(5) = "Login password"
stranswers(6) = "File access permission"
stranswers(7) = "Super user"
stranswers(8) = "Su"
stranswers(9) = "Console"
stranswers(10) = "Primary key"

End Sub

Private Sub Matching_Columns_Paper6()
strcolumnA(1) = "A program/utility under WINDOWS"
strcolumnA(2) = "Exchange of images or data between appliations"
strcolumnA(3) = "More automatic than clipboard"
strcolumnA(4) = "Runs DOS commands concerning files and directories"
strcolumnA(5) = "A communication program to connect computers over telephone lines"
strcolumnA(6) = "Tool to move from one Window to another "
strcolumnA(7) = "Waiting to start"
strcolumnA(8) = "Open area wher text, graph etc. is typed ordrawn"
strcolumnA(9) = "Confirmation of changes in dialog box and its execution begins with processing"
strcolumnA(10) = "Geographical restriction in placing inteconnected computers"

strcolumnB(1) = "Clipboard"
strcolumnB(2) = "Tasklist"
strcolumnB(3) = "Ok"
strcolumnB(4) = "LAN"
strcolumnB(5) = "Dynamic data excahange "
strcolumnB(6) = "Terminal"
strcolumnB(7) = "File program"
strcolumnB(8) = "Unloaded icon"
strcolumnB(9) = "Icon"
strcolumnB(10) = "Work area"

stranswers(1) = "Icon"
stranswers(2) = "Clipboard"
stranswers(3) = "Dynamic data exchange"
stranswers(4) = "File program"
stranswers(5) = "Terminal"
stranswers(6) = "Tasklist"
stranswers(7) = "Unloaded icon"
stranswers(8) = "Work area"
stranswers(9) = "Ok"
stranswers(10) = "LAN"

End Sub

Private Sub Matching_Columns_Paper7()
strcolumnA(1) = "The way computers are connected"
strcolumnA(2) = "Traffic between computers in LAN"
strcolumnA(3) = "Computer network spread in large geographical area"
strcolumnA(4) = "Communication lines in computer networks"
strcolumnA(5) = "Developer of data commuication standards"
strcolumnA(6) = "An EIA standard for connection of peripheral devices with PCs"
strcolumnA(7) = "Used in connecting terminals to computers at short distance"
strcolumnA(8) = "Group of specially wrapped and insulated wire lines "
strcolumnA(9) = "Line of sight method of communication"
strcolumnA(10) = "Speed of satellite matches earth's rotation speed"

strcolumnB(1) = "Short burst"
strcolumnB(2) = "Twisted pair"
strcolumnB(3) = "DOT circuits"
strcolumnB(4) = "ISO"
strcolumnB(5) = "Coaxial cable"
strcolumnB(6) = "Geo synchronous orbit"
strcolumnB(7) = "RS-232C"
strcolumnB(8) = "Microwave transmission"
strcolumnB(9) = "Public network"
strcolumnB(10) = "Topology"

stranswers(1) = "Topology"
stranswers(2) = "Short burst"
stranswers(3) = "Public network"
stranswers(4) = "DOT circuits"
stranswers(5) = "ISO"
stranswers(6) = "RS-232C"
stranswers(7) = "Twisted pair"
stranswers(8) = "Coaxial cable"
stranswers(9) = "Microwave transmission"
stranswers(10) = "Geo synchronous orbit"

End Sub

Private Sub Matching_Columns_Paper8()
strcolumnA(1) = "Data transmission through optical fiber"
strcolumnA(2) = "Transmission capacity of communication channel"
strcolumnA(3) = "Transmission speed of communication channel"
strcolumnA(4) = "Modulation/Demodulation "
strcolumnA(5) = "Parallel data lines"
strcolumnA(6) = "Bi directioinal data bus. Rules and coventions for synchronization"
strcolumnA(7) = "Rules and conventions for synchronization"
strcolumnA(8) = "Low voltage signal on data lines"
strcolumnA(9) = "A character oriented protocol"
strcolumnA(10) = "Data is SDLC/HDLC"

strcolumnB(1) = "Bandwidth"
strcolumnB(2) = "Baud rate"
strcolumnB(3) = "Modem"
strcolumnB(4) = "Data link control"
strcolumnB(5) = "Marking state on line"
strcolumnB(6) = "BSC(Bisync)"
strcolumnB(7) = "Half Duplex"
strcolumnB(8) = "Information field"
strcolumnB(9) = "Total internal reflection"
strcolumnB(10) = "Data bus"

stranswers(1) = "Total internal reflection"
stranswers(2) = "Bandwidth"
stranswers(3) = "Baud rate"
stranswers(4) = "Modem"
stranswers(5) = "Data bus"
stranswers(6) = "Half Duplex"
stranswers(7) = "Data link control"
stranswers(8) = "Marking state on line"
stranswers(9) = "BSC(Bisync)"
stranswers(10) = "Information field"

End Sub

Private Sub Matching_Columns_Paper9()
strcolumnA(1) = "Flag bits to signal start of data"
strcolumnA(2) = "Transmission of multiple signals over single channel"
strcolumnA(3) = "Splitting available bandwidth into smaller disjoint bandwidth"
strcolumnA(4) = "Allocating a time slot to each user of channel "
strcolumnA(5) = "Any device to offer service to LAN users"
strcolumnA(6) = "Set of rules for orderly transfer of data among network users"
strcolumnA(7) = "Manner in which digits 1's or 0's are trasmitted on netowrk's data path"
strcolumnA(8) = "The way networks are physically connected together"
strcolumnA(9) = "Elementary asynchronous protocol"
strcolumnA(10) = "Providing security and privacy in distributed dat aprocessing systems"

strcolumnB(1) = "Encryption"
strcolumnB(2) = "LAN signalling"
strcolumnB(3) = "ASCII"
strcolumnB(4) = "Server"
strcolumnB(5) = "Protocol"
strcolumnB(6) = "TDM"
strcolumnB(7) = "Topology"
strcolumnB(8) = "FDM"
strcolumnB(9) = "Multiplexing"
strcolumnB(10) = "HDLC/SDLC"

stranswers(1) = "HDLC/SDLC"
stranswers(2) = "Multiplexing"
stranswers(3) = "FDM"
stranswers(4) = "TDM"
stranswers(5) = "Server"
stranswers(6) = "Protocol"
stranswers(7) = "LAN signaling"
stranswers(8) = "Topology"
stranswers(9) = "ASCII"
stranswers(10) = "Encryption"

End Sub

Private Sub Matching_Columns_Paper10()
strcolumnA(1) = "Client/server processing enviromnent"
strcolumnA(2) = "Cooperative processing"
strcolumnA(3) = "Internet protocol"
strcolumnA(4) = "Access to internet in India"
strcolumnA(5) = "Cycle time 4 nano seconds or less"
strcolumnA(6) = "PARAM"
strcolumnA(7) = "Motorola Microprocessor"
strcolumnA(8) = "The Standard Operating System of Work Station"
strcolumnA(9) = "Suitable for CAD and and Simulation of Complex Systems"
strcolumnA(10) = " GFLOPS"

strcolumnB(1) = "Client/server"
strcolumnB(2) = "Supercomputer"
strcolumnB(3) = "Apple's PC"
strcolumnB(4) = "UNIX"
strcolumnB(5) = "Supercomputer performance"
strcolumnB(6) = "Workstation"
strcolumnB(7) = "VSNL's GIAS network"
strcolumnB(8) = "Indian supercomputer"
strcolumnB(9) = "TCP/IP"
strcolumnB(10) = "Sybase"

stranswers(1) = "Sybase"
stranswers(2) = "Client/server"
stranswers(3) = "TCP/IP"
stranswers(4) = "VSNL's GIAS network"
stranswers(5) = "Supercomputer"
stranswers(6) = "Indian supercomputer"
stranswers(7) = "Apple's PC"
stranswers(8) = "UNIX"
stranswers(9) = "Workstation"
stranswers(10) = "Supercomputer performance"

End Sub

Private Sub Matching_Columns_Paper11()
strcolumnA(1) = "Concentration on the problem specification "
strcolumnA(2) = "Picture transmission"
strcolumnA(3) = "DOS command to list diretories and their files"
strcolumnA(4) = "Combining voice, text and picture"
strcolumnA(5) = "Page printer"
strcolumnA(6) = "Any computer which can access server"
strcolumnA(7) = "Exchange of data between two networks"
strcolumnA(8) = "To find and fix errors"
strcolumnA(9) = "Continuous variable sugbals"
strcolumnA(10) = "Small high speed memory."

strcolumnB(1) = "Fax"
strcolumnB(2) = "TREE/F"
strcolumnB(3) = "Non procedural language"
strcolumnB(4) = "Client"
strcolumnB(5) = "Bridge"
strcolumnB(6) = "Multimedia"
strcolumnB(7) = "Laser printer"
strcolumnB(8) = "Cache"
strcolumnB(9) = "Debug"
strcolumnB(10) = "Analog computer"

stranswers(1) = "Non procedural language"
stranswers(2) = "Fax"
stranswers(3) = "TREE/F"
stranswers(4) = "Multimedia"
stranswers(5) = "Laser printer"
stranswers(6) = "Client"
stranswers(7) = "Bridge"
stranswers(8) = "Debug"
stranswers(9) = "Analog computer"
stranswers(10) = "Cache"

End Sub
Private Sub Matching_Columns_Paper12()
strcolumnA(1) = "Concurrently performing several computations"
strcolumnA(2) = "Processor component in multiprocessing"
strcolumnA(3) = "Assembly language"
strcolumnA(4) = "Output of assembler"
strcolumnA(5) = "Conversion of mnemonic operation code to absolute numeric operation code"
strcolumnA(6) = "Reading program into memory"
strcolumnA(7) = "Converting assembly level program into machine lnaguage"
strcolumnA(8) = "Assembly language statement translation to machine language"
strcolumnA(9) = "Translator of higg level language program"
strcolumnA(10) = "Statement by statement translation & execution"

strcolumnB(1) = "One to one"
strcolumnB(2) = "Parallel processing"
strcolumnB(3) = "Interpreter"
strcolumnB(4) = "Compiler"
strcolumnB(5) = "Loader"
strcolumnB(6) = "Transputer"
strcolumnB(7) = "Assembler"
strcolumnB(8) = "Translator"
strcolumnB(9) = "Object program"
strcolumnB(10) = "mnemonics"

stranswers(1) = "Parallel processing"
stranswers(2) = "Transputer"
stranswers(3) = "mnemonics"
stranswers(4) = "Object program"
stranswers(5) = "Assembler"
stranswers(6) = "Loader"
stranswers(7) = "Translator"
stranswers(8) = "One to one"
stranswers(9) = "Compiler"
stranswers(10) = "Interpreter"

End Sub
Private Sub Matching_Columns_Paper13()
strcolumnA(1) = "End to end data transfer"
strcolumnA(2) = "Program to edit, view and create text"
strcolumnA(3) = "Terminal with keyboard and printer"
strcolumnA(4) = "Sharing a channel with many data sources"
strcolumnA(5) = "Billion"
strcolumnA(6) = "Unwanted and meaningless words in memory/storage"
strcolumnA(7) = "Information processing and retrieval"
strcolumnA(8) = "Wide area network"
strcolumnA(9) = "Set of interconnected computers"
strcolumnA(10) = "Low cost magnetic storage"

strcolumnB(1) = "Word processor"
strcolumnB(2) = "Database"
strcolumnB(3) = "Multiplexing"
strcolumnB(4) = "Teleprinter"
strcolumnB(5) = "Network"
strcolumnB(6) = "Floppy disk"
strcolumnB(7) = "Circuit switching"
strcolumnB(8) = "Giga"
strcolumnB(9) = "Garbage"
strcolumnB(10) = "Satellite link"

stranswers(1) = "Circuit switching"
stranswers(2) = "Word processor"
stranswers(3) = "Teleprinter"
stranswers(4) = "Multiplexing"
stranswers(5) = "Giga"
stranswers(6) = "Garbage"
stranswers(7) = "Database"
stranswers(8) = "Satellite link"
stranswers(9) = "Network"
stranswers(10) = "Floppy disk"

End Sub

Private Sub Matching_Columns_Paper14()
strcolumnA(1) = "Macro instructions"
strcolumnA(2) = "Responsible for decision making in a computer"
strcolumnA(3) = "Central nervous system for all components of a computer"
strcolumnA(4) = "Over all responsible for execution of program"
strcolumnA(5) = "CPU on chip"
strcolumnA(6) = "Special memory unit within CPU"
strcolumnA(7) = "Address of nexrt instruction to be executed"
strcolumnA(8) = "Main memory"
strcolumnA(9) = "Mass storage memory"
strcolumnA(10) = "Binary digits"

strcolumnB(1) = "ALU"
strcolumnB(2) = "Control unit"
strcolumnB(3) = "CPU"
strcolumnB(4) = "Secondary memory"
strcolumnB(5) = "Registers"
strcolumnB(6) = "Microprocesssors"
strcolumnB(7) = "Macro programming"
strcolumnB(8) = "Bits"
strcolumnB(9) = "Program register"
strcolumnB(10) = "RAM"

stranswers(1) = "Macro programming"
stranswers(2) = "ALU"
stranswers(3) = "Control unit"
stranswers(4) = "CPU"
stranswers(5) = "Microprocessors"
stranswers(6) = "Registers"
stranswers(7) = "Program register"
stranswers(8) = "RAM"
stranswers(9) = "Secondary memory"
stranswers(10) = "Bits"
End Sub

Private Sub Matching_Columns_Paper15()
strcolumnA(1) = "A 4 bit portion of byte"
strcolumnA(2) = "Basic information unit of computer"
strcolumnA(3) = "RAM and ROM"
strcolumnA(4) = "I/O devices and secondary storage unit"
strcolumnA(5) = "Terminal having inbuilt processing capabitlity "
strcolumnA(6) = "Mouse"
strcolumnA(7) = "Non impact printer"
strcolumnA(8) = "Drums are NOT used in"
strcolumnA(9) = "Inter record gap"
strcolumnA(10) = "Rotating at variable speed"

strcolumnB(1) = "Word"
strcolumnB(2) = "Semi conductor memory"
strcolumnB(3) = "Microgrip plotter"
strcolumnB(4) = "Pointing device"
strcolumnB(5) = "Magnetic tape"
strcolumnB(6) = "Peripherals"
strcolumnB(7) = "CD - ROM"
strcolumnB(8) = "Chemical technology"
strcolumnB(9) = "A smart terminal"
strcolumnB(10) = "Nibble"


stranswers(1) = "Nibble"
stranswers(2) = "Word"
stranswers(3) = "Semi conductor memory"
stranswers(4) = "Peripherals"
stranswers(5) = "A smart terminal"
stranswers(6) = "Pointing device"
stranswers(7) = "Chemical technology"
stranswers(8) = "Microgrip plotter"
stranswers(9) = "Magnetic tape"
stranswers(10) = "CD - ROM"

End Sub

Private Sub questions_paper1()
StrQuestions(1) = "Justify this statement." & Chr(4) & "Data are Raw materials and Information is the Product " & Chr(4)
StrQuestions(2) = "A. What are the values of information to the receivers?" & Chr(13) & _
                  "B. How do they help receivers in the decision making process?"
StrQuestions(3) = "a. State the various categories of information n a business organization." & Chr(13) & _
                  "b. How is each of these categories of information utilized in different kinds of planning?"
StrQuestions(4) = "a. What levels of information are required for large, medium and small sized companies and departments?" & Chr(13) & _
                  "b. What are the factors on which these levels depend?"
StrQuestions(5) = "How are logical data physically stored and retrieved?"
StrQuestions(6) = "Compare the storage and retrieval operations of data performed but manual and computer systems."
StrQuestions(7) = "a. What are the operations involved in a data processing activity" & Chr(13) & _
                  "b. How is relevance of information maintained?"
StrQuestions(8) = "What are the advantages of organizing data into files and how do they help in accessing information?"
StrQuestions(9) = "a. Explain the following types of data processing:" & Chr(13) & _
                  "i. Commercial application" & Chr(13) & _
                  "ii. Process control application" & Chr(13) & _
                  "b. What are their similarities and dissimilarities?"
StrQuestions(10) = "a. Why are mechanical aids needed for data processing in a business environment?" & Chr(13) & _
                   "b. What are its immediate outcomes?"
End Sub

Private Sub questions_paper2()
StrQuestions(1) = "A. Why do you think electrical and electronic machines are inevitable in a large business organization?" & Chr(13) & _
                  "B. What are their effects?"
StrQuestions(2) = "Explain the following data processing methods:" & Chr(13) & _
                  "a.Transaction processing" & Chr(13) & _
                  "b.Batch processing" & Chr(13) & _
                  "c. On line processing" & Chr(13) & _
                  "d.Word processing"
StrQuestions(3) = "a. What are word processors?" & Chr(13) & _
                  "b. What are the functions of word processor?"
StrQuestions(4) = "A. What are the different forms of reports?" & Chr(13) & _
                  "B. What are their relative merits and demerits?"
StrQuestions(5) = "What are the advantages of graphic reports over tabular reports?"
StrQuestions(6) = "a. Explain briefly the operations performed by an office automation system." & Chr(13) & _
                  "b. Name and justify the applications suitable for each of them."
StrQuestions(7) = "a. What is analog and digital processing of data?" & Chr(13) & _
                  "b. Name and justify the applications suitable for each of them."
StrQuestions(8) = "Draw the block diagram of a digital computer and explain the functions of each of constituent blocks."
StrQuestions(9) = "a. What are the functions of a CPU?" & Chr(13) & _
                  "b. How are they performed by the CPU?"
StrQuestions(10) = "How is the arithmetic and logical operations performed?"
End Sub

Private Sub questions_paper3()
StrQuestions(1) = "Explain the functions of the following" & Chr(13) & _
    "a. Program Counter" & Chr(13) & _
    "b. Memory address register" & Chr(13) & _
    "c. Memory data register" & Chr(13) & _
    "d. Instruction register" & Chr(13) & _
    "e. Stack register" & Chr(13) & _
    "f. Index register" & Chr(13) & _
    "g. Base address registers" & Chr(13) & _
    "h. Flag register" & Chr(13) & _
    "i. General purpose registers" & Chr(13) & _
    "j. Accumulator"
StrQuestions(2) = "How are alpha numeric characters stored in RAM?"
StrQuestions(3) = "a. What do you understand by the capacity of a memory?" & Chr(13) & _
                  "b. What are the factors on which the capacity of a memory unit depends?"
StrQuestions(4) = "What determines the bit length of a computer?"
StrQuestions(5) = "a. Accessing methods" & Chr(13) & _
                  "b. Address decoding" & Chr(13) & _
                  "c. Volatility and non volatility" & Chr(13) & _
                  "d. Bipolar & MOS memories" & Chr(13) & _
                  "e. ROMs and RAMs" & Chr(13) & _
                  "f. ROMs , PROMs, EPROMs And EEPROMs" & Chr(13) & _
                  "g. State and dynamic RAM's" & Chr(13) & _
                  "h. DRO and NDRO""& chr(13) & "
StrQuestions(6) = "a. What do you understand by memory word length?" & Chr(13) & _
                  "b. What is the value of the maximum unsigned number stored in a memory word?"
StrQuestions(7) = "What is the total capacity of memory in bits, when the lengths of MAR ad MDR are m and n bits respectively?"
StrQuestions(8) = "How is an instruction executed by a computer?"
StrQuestions(9) = "Explain briefly the following" & Chr(13) & _
                  "a. Clock cycles" & Chr(13) & _
                  "b. Machine cycles" & Chr(13) & _
                  "c. Instruction fetch cycles" & Chr(13) & _
                  "d. Instruction execution cycles" & Chr(13) & _
                  "e. Instruction cycles" & Chr(13)
StrQuestions(10) = "a. How are instructions executed sequentially in a computer?" & Chr(13) & _
                   "b. How can the sequential execution be modified?"
End Sub

Private Sub questions_paper4()
StrQuestions(1) = "a. What are conditional and unconditional transfers of control flow?" & Chr(13) & _
                  "b. How are they realized?"
StrQuestions(2) = "a. What is a multiprocessing system?" & Chr(13) & _
                  "b. Why is it required and how does it operate?"
StrQuestions(3) = "a. What are input devices?" & Chr(13) & _
                  "b. What is the importance of an input device?" & Chr(13) & _
                  "c. What are the characteristics of a good input device?" & Chr(13) & _
                  "d. What are the various types of inputs?"
StrQuestions(4) = "a. How are data entered from source documents into a computer?" & Chr(13) & _
                  "b. Explain the functions of off line and on line data entry operations; state their relative merits and demerits."
StrQuestions(5) = "Why are the following input devices used and how do they work?" & Chr(13) & _
                  "a.MICR" & Chr(13) & _
                  "b.OCR" & Chr(13) & _
                  "c.OMR"
StrQuestions(6) = "What are scanners and how do they work?"
StrQuestions(7) = "a. Explain the working principles of the following:" & Chr(13) & _
                  "a.Joy sticks" & Chr(13) & _
                  "b.Mouse" & Chr(13) & _
                  "c. Light pens"
StrQuestions(8) = "a. Is it possible to input voice into computers?" & Chr(13) & _
                  "b. How does a voice input device work?" & Chr(13) & _
                  "c. What functions can be performed by a voice input system?"
StrQuestions(9) = "a. What are multilingual applications of computer?" & Chr(13) & _
                  "b. Why do we require them in India?" & Chr(13) & _
                  "c. How do they work?"
StrQuestions(10) = "Name some of the output devices used with computer systems. Mention their applications"
End Sub

Private Sub questions_paper5()
StrQuestions(1) = "What do you understand by hard copy and soft copy output devices? Give examples."
StrQuestions(2) = "What do you understand by impact and non impact printers? Give examples."
StrQuestions(3) = "Compare the following printers:" & Chr(13) & _
                  "a. DMPs" & Chr(13) & _
                  "b. Line Printers" & Chr(13) & _
                  "c. Inkjet Printers" & Chr(13) & _
                  "d. Laser Printers"
StrQuestions(4) = "Compare the following terminals:" & Chr(13) & _
                  "a. Graphic and non graphic" & Chr(13) & _
                  "b. Monochrome and color."
StrQuestions(5) = "a. What are the functions of plotters?" & Chr(13) & _
                  "b. Why are they user?" & Chr(13) & _
                  "c. How do they work?"
StrQuestions(6) = "Explain the working principle of a multipen plotter."
StrQuestions(7) = "a. What is a micro film?" & Chr(13) & _
                  "b. Why is it needed and how does it work?"
StrQuestions(8) = "a. What is a voice synthesizer?" & Chr(13) & _
                  "b. Why is it needed and how it work?"
StrQuestions(9) = "a. What is spooling?" & Chr(13) & _
                  "b. Why is it needed and how does it work?   "
StrQuestions(10) = "a. How is information stored & retrieved on & from magnetic tape?" & Chr(13) & _
                   "b. What are tracks and recording density?"
End Sub

Private Sub questions_paper6()
StrQuestions(1) = "a. What are inter record gaps?" & Chr(13) & _
                  "b. Why are they required?"
StrQuestions(2) = "Explain the operation of a magnetic ape system. Estimate the average access time of a record from a magnetic tape."
StrQuestions(3) = "Define and explain the following with respect to magnetic disk units:" & Chr(13) & _
                  "a. Tracks" & Chr(13) & _
                  "b. Sectors" & Chr(13) & _
                  "c. Surfaces" & Chr(13) & _
                  "d. Cylinders" & Chr(13) & _
                  "e. Recording density" & Chr(13) & _
                  "f. Capacity" & Chr(13) & _
                  "g. Rotational speed and data transfer rate" & Chr(13) & _
                  "h. Seek Time" & Chr(13) & _
                  "i. Latency" & Chr(13) & _
                  "g. Access Time"
StrQuestions(4) = "What do you understand by the following?" & Chr(13) & _
                  "a. Fixed head and moving head disks" & Chr(13) & _
                  "b. Floppy and Winchester disks."
StrQuestions(5) = "a. How are records addressed in a disk system?" & Chr(13) & _
                  "b. Why are disks called DASD?"
StrQuestions(6) = "For records stored on magnetic disks explain by and how it is possible to access/add/delete records in random fashion?"
StrQuestions(7) = "a. Explain the operating principles of an optical memory?" & Chr(13) & _
                  "b. Why is it also called video disks or CD ROM?"
StrQuestions(8) = "a. Define software." & Chr(13) & _
                  "B.What is the roles of software in computers?"
StrQuestions(9) = "With the help of suitable examples explain briefly the following categories of software:" & Chr(13) & _
                  "a. System software" & Chr(13) & _
                  "b. Utilities" & Chr(13) & _
                  "c. Application software" & Chr(13) & _
                  "d. Packages"
StrQuestions(10) = "Explain the function o the following system software:" & Chr(13) & _
                   "a. Operating Systems" & Chr(13) & _
                   "b. Assemblers" & Chr(13) & _
                   "c. Compilers" & Chr(13) & _
                   "d. Interpreters" & Chr(13) & _
                   "e. Simulators" & Chr(13) & _
                   "f. Emulators" & Chr(13) & _
                   "Name some utility software underplay their functions."
End Sub

Private Sub questions_paper7()
StrQuestions(1) = "a. What is a disk defragmenter?" & Chr(13) & _
                  "b. Why is it needed and how does it work?"
StrQuestions(2) = "Briefly describe the functioning of virus scanners & removers."
StrQuestions(3) = "What is licensed software and how are they updated?"
StrQuestions(4) = "a. How can computers be classified?" & Chr(13) & _
                  "b. What are the characteristics used for such classification?"
StrQuestions(5) = "Compare the performance of the following classes of computers:" & Chr(13) & _
                  "a.Super computers" & Chr(13) & _
                  "b.Mainframe computers" & Chr(13) & _
                  "c.Midi computers" & Chr(13) & _
                  "d.Mini computers" & Chr(13) & _
                  "e.Micro computers" & Chr(13) & _
                  "f.Personal computers" & Chr(13) & _
                  "g.Work stations"
StrQuestions(6) = "Explain and compare CISC and RISC."
StrQuestions(7) = "Explain the functions of an Operating System."
StrQuestions(8) = "How does a processor manager perform job scheduling in the following environments?" & Chr(13) & _
                  "a.Single user" & Chr(13) & _
                  "b.Multi programming" & Chr(13) & _
                  "c.Time sharing"
StrQuestions(9) = "Explain and compare the following" & Chr(13) & _
                  "a.Batch processing" & Chr(13) & _
                  "b.Multiprogramming" & Chr(13) & _
                  "c.Time sharing" & Chr(13) & _
                  "d.Multiprogramming"
StrQuestions(10) = "What do you understand but priorities and protections in a multiprogramming environment?"
End Sub

Private Sub questions_paper8()
StrQuestions(1) = "a. What are the functions of a device manager?" & Chr(13) & _
                  "b. What are sharable and non sharable devices?"
StrQuestions(2) = "What is virtual storage and how is it implemented?"
StrQuestions(3) = "In a virtual memory system, explain the following:" & Chr(13) & _
                  "a.Demand Paging" & Chr(13) & _
                  "b. Locality of references" & Chr(13) & _
                  "c. Memory allocations" & Chr(13) & _
                  "d. Segments" & Chr(13) & _
                  "e. Paging" & Chr(13) & _
                  "f. Paged segments" & Chr(13) & _
                  "g. Pre emptive and non preemptive" & Chr(13) & _
                  "h. Strategies for preemptive allocations." & Chr(13) & _
                  "i. FIFO and LRU replacement policies"
StrQuestions(4) = "a. What are the functions of an information manager?" & Chr(13) & _
                  "b. What are the different structures of files?" & Chr(13) & _
                  "c. What are their merits and demerits?"
StrQuestions(5) = "What are the functions of a command interpreter?"
StrQuestions(6) = "Describe the log n a log off processes used in UNIX and / or network Operating Systems. How re the user accounts created in UNIX and / or Network."
StrQuestions(7) = "Name the commands used for basic file manipulation in DOS/Unix/Network."
StrQuestions(8) = "Explain the directory structures interrelated commands for creating, deleting & changing directories in DOS/Unix/Network."
StrQuestions(9) = "How are three file securities included in UNIX systems?"
StrQuestions(10) = "a. What are shells in UNIX?" & Chr(13) & _
                   "b. Describe facilities offered by shells."
End Sub

Private Sub questions_paper9()
StrQuestions(1) = "What are GUI's and what are their functions?"
StrQuestions(2) = "a. Explain the features of WINDOWS." & Chr(13) & _
                  "b. How is it different from DOS?" & Chr(13) & _
                  "c. How is the mouse used in WINDOWS"
StrQuestions(3) = "a. Describe computer communication." & Chr(13) & _
                  "b. How is data communicated over long distances?"
StrQuestions(4) = "a. What is meant but the band with of a channel?" & Chr(13) & _
                  "b. Arrange the following channels in ascending order of band width." & Chr(13) & _
                  "i]   Twisted pair" & Chr(13) & _
                  "ii] Coaxial pair" & Chr(13) & _
                  "iii] Optical fiber" & Chr(13) & _
                  "c. What are the advantages of having higher band width?" & Chr(13) & _
                  "d. What are the relative merits and demerits of these channels?"
StrQuestions(5) = "Establish the relationship between baud rate and band width."
StrQuestions(6) = "Explain the following modes of transmission in a communication environment" & Chr(13) & _
                  "a. Simplex" & Chr(13) & _
                  "b. Duplex" & Chr(13) & _
                  "c. Half Duplex"
StrQuestions(7) = "Justify the statement “For transmission of data over distant, serial communication is preferred to parallel communication"""
StrQuestions(8) = "a. Differentiate between asynchronous and synchronous communications" & Chr(13) & _
                  "b. What are their relative merits and demerits?"
StrQuestions(9) = "a. Draw and explain the format of the data stream in an asynchronous ode of data transmission." & Chr(13) & _
                  "b. What do you mean by a synchronization character and why is it required?" & Chr(13) & _
                  "c. What are double synchronization characters and why are they used?"
StrQuestions(10) = "a. State reasons for preferring analog transmission over digital." & Chr(13) & _
                   "b. What role does a modem play in communications?" & Chr(13) & _
                   "c. What is its principle of operation?"
End Sub

Private Sub questions_paper10()
StrQuestions(1) = "a. Draw schematic diagram depicting computer to computer communication using modems" & Chr(13) & _
                  "b. Is this communication mode duplex ? Justify."
StrQuestions(2) = "a. Define multiplexing of a communication channel." & Chr(13) & _
                  "b. Explain the different types of multiplexing?"
StrQuestions(3) = "What is the need for and advantages of, computer networks?"
StrQuestions(4) = "Differentiate between LAN and WAN highlighting advantages and disadvantages"
StrQuestions(5) = "Enumerate and compare the different topologies used in a LAN."
StrQuestions(6) = "Comment on the following Statement." & Chr(4) & "Distributed processing is no different from a network." & Chr(4)
StrQuestions(7) = "a. Tabulate the important features of real time and on line systems." & Chr(13) & _
                  "b. Give examples of three systems in each category:"
StrQuestions(8) = "Explain the role plated by gateways and bridges in networking."
StrQuestions(9) = Chr(4) & " Internet has led to Globalization in Information Technology." & Chr(4) & "Comment in about 200 words."
StrQuestions(10) = Chr(4) & " E mail has expanded the horizons of Commerce." & Chr(4) & "Comment in about 200 words mentioning also some unsatisfactory features."
End Sub

Private Sub questions_paper11()
StrQuestions(1) = "WWW is said to lead to an open Society? What are its important features and how do you get to open a home page?"
StrQuestions(2) = "Explain briefly the following:" & Chr(13) & _
                  "a.Gopher" & Chr(13) & _
                  "b.Mosaic" & Chr(13) & _
                  "c.Java" & Chr(13) & _
                  "d. TCP/IP"
StrQuestions(3) = "Explain how machine independent cross platform programming is performed on a network?"
StrQuestions(4) = "a. Enumerate the important characteristics of super and main frame computers?" & Chr(13) & _
                  "b. How are their performances measured?" & Chr(13) & _
                  "c. Name three of each type."
StrQuestions(5) = "Compare and contrast the main features of a work station and a personal computer."
StrQuestions(6) = "Write the specifications of a currently available entry level personal computer system. What add on are possible?"
StrQuestions(7) = Chr(4) & " Client Server architecture is NOT just another computer network." & Chr(13) & "Analyze bringing out benefits accruing to management."
StrQuestions(8) = "What are the various parts of the following and how are they written & stored in main memory?" & Chr(13) & _
                  "a. A machine language instruction." & Chr(13) & _
                  "b. An assembly language instruction."
StrQuestions(9) = "Explain the following" & Chr(13) & _
                  "a.Mnemonics" & Chr(13) & _
                  "b.op -codes" & Chr(13) & _
                  "c.Pseudo op - codes" & Chr(13) & _
                  "d.Symbolic addresses"
StrQuestions(10) = "a. What are assembler directives?" & Chr(13) & _
                   "b. How are they interpreted by the machine?"
End Sub

Private Sub questions_paper12()
StrQuestions(1) = "a. What is a two pass assembler?" & Chr(13) & _
                  "b. How does it convert an assume language program to the machine language?"
StrQuestions(2) = "a. Give example of high level languages?" & Chr(13) & _
                  "b. Why do you require so many different languages?" & Chr(13) & _
                  "c. What are their special features?"
StrQuestions(3) = "Explain the following:" & Chr(13) & _
                  "a.Machine independence" & Chr(13) & _
                  "b.Portability" & Chr(13) & _
                  "c. Problem oriented languages"
StrQuestions(4) = "Explain how a program written in a high level language is converted to a machine executable form."
StrQuestions(5) = "Com is compilers and interpreters."
StrQuestions(6) = "Explain the compilation process."
StrQuestions(7) = "Explain the functional purpose of linkers and loaders."
StrQuestions(8) = "List the special features and give examples of third generation languages."
StrQuestions(9) = "What do you understand by the following and give examples to justify your answer?" & Chr(13) & _
                  "a. Scientific programming languages" & Chr(13) & _
                  "b. Business oriented programming languages."
StrQuestions(10) = "Explain the characteristics with suitable examples, procedural and non procedural languages."
End Sub

Private Sub questions_paper13()
StrQuestions(1) = "What are important measures necessary to ensure security of computer based information systems?"
StrQuestions(2) = "Explain the following and why they are required" & Chr(13) & _
                  "a.Authorization" & Chr(13) & _
                  "b.Authentication" & Chr(13) & _
                  "c.Password" & Chr(13) & _
                  "d.File Permission"
StrQuestions(3) = "a. What do you understand by privacy and protection of users?" & Chr(13) & _
                  "b. How can these be assured?"
StrQuestions(4) = "Explain the reasons for introducing the following and how they should be implemented:" & Chr(13) & _
                  "a. electronic access control" & Chr(13) & _
                  "b.Cryptography"
StrQuestions(5) = "a. Give examples of pervasive software." & Chr(13) & _
                  "b. How do they affect the system?"
StrQuestions(6) = "a. What are viruses?" & Chr(13) & _
                  "b. How can the system be protected from viruses?"
StrQuestions(7) = "Write notes on the following:" & Chr(13) & _
                  "a. Intellectual property rights." & Chr(13) & _
                  "b. Legitimate software."
StrQuestions(8) = "Explain the following applications of computers highlighting their special features:" & Chr(13) & _
                  "a.Weather forecasting" & Chr(13) & _
                  "b.Remote sensing"
StrQuestions(9) = "a. How can the computer be used for collecting data for providing information for planning activities?" & Chr(13) & _
                   "b. Give example of an Indian system."
StrQuestions(10) = "a. What is a voice input device?" & Chr(13) & _
                   "b. Why is it required and how does it work?"
End Sub

Private Sub questions_paper14()
StrQuestions(1) = "a. What is ISCII?" & Chr(13) & _
                  "b. How can they be entered, displayed and printed?"
StrQuestions(2) = "a. How can the key board functions be interpreted to enter multilingual character set?" & Chr(13) & _
                  "b. How can he multilingual character set be displayed on the screens?" & Chr(13) & _
                  "c. How can the multilingual characters be printed?"
StrQuestions(3) = "Explain the purposes of the following on going IT projects in India:" & Chr(13) & _
                  "a. NICNET   b.ERNET,    c. INDONET"
StrQuestions(4) = "a. What is PARAM ?" & Chr(13) & _
                  "b. Where was it developed?" & Chr(13) & _
                  "c. What are the spatial features of this system?" & Chr(13) & _
                  "d. How does it differ from the other contemporary systems?"
StrQuestions(5) = "What are the different types of errors that are normally encountered during testing and debugging of programmes? How can they be eliminated?"
StrQuestions(6) = "How are computers classified based on: (a) Principle employed (b) Flinn's classification"
StrQuestions(7) = "How is an OCR device is different from MICR technology?"
StrQuestions(8) = "What is the difference between hardware and software?"
StrQuestions(9) = "Illustrate asynchronous and synchronous data transmission modes. Which of the above is used in high speed data transmission and why?"
StrQuestions(10) = "Mention three factors that makes high level languages ore advantageous than machine level languages."
End Sub

Private Sub questions_paper15()
StrQuestions(1) = "Discuss the data communication on a satellite link. Why is earth to satellite transmission frequency different from satellite to earth retransmission frequency?"
StrQuestions(2) = "What is meant by linking? Why is it required?"
StrQuestions(3) = "A computer needs a floppy disk or hard disk program to boot. Why is not this program stored in the ROM?"
StrQuestions(4) = "What is the difference between a graphical printer and a plotter? Compare in terms of speed, accuracy or quality of output."
StrQuestions(5) = "compare fixed head and moving head magnetic disks."
StrQuestions(6) = "What causes head crashes? How can Winchester technology reduce the read crash problem?"
StrQuestions(7) = "How can data stored on magnetic disks be accesses? What determines the time required to access the needed data?"
StrQuestions(8) = "What is data communication? What components are typically found in a simple data communication system?"
StrQuestions(9) = "Explain difference between LAN's and WAN's?"
StrQuestions(10) = "What are the advantages of using optical fiber as compared to copper wires for data communication? discuss the order of bit rates supported by optical fibers."
End Sub

Private Sub TrueorFalse_Paper1()
StrQuestions(1) = "Data and information are the same."
StrQuestions(2) = "Information is obtained after manipulating raw facts."
StrQuestions(3) = "A structured decision is not routine."
StrQuestions(4) = "Managers need information which is critical to their performance and success."
StrQuestions(5) = "All managers in a big organization use the same type of information."
StrQuestions(6) = "Printers can also be classified on how they print."
StrQuestions(7) = "Entities may have similar attributes."
StrQuestions(8) = "Data items in records are also known as fields."
StrQuestions(9) = "In Sequential type of storage devices, any record can be accessed at time taken to write or read any record will be same."
StrQuestions(10) = "Winchester disks come in various sizes."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper2()
StrQuestions(1) = "Sequential files are suited for on line inquiry processing."
StrQuestions(2) = "A dot matrix printer can print any shape of character through software."
StrQuestions(3) = "On line processing and real tine processing are same."
StrQuestions(4) = "Decentralized processing and destructed processing are same."
StrQuestions(5) = "Classifying and sorting data are same."
StrQuestions(6) = "Prototyping mode is created at the problem definition stage of system development."
StrQuestions(7) = "Drawings produced by pen plotters are precise and good quality."
StrQuestions(8) = "A lighter or darker i.e. a more solid impression on printed page, by a word processor package is possible."
StrQuestions(9) = "A word processor package may have the facility for caching the spelling of words in a document."
StrQuestions(10) = "Lower the resolution o a display screen, the more readable the images in accuracy of detail."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"

End Sub
Private Sub TrueorFalse_Paper3()
StrQuestions(1) = "E-mail reaches its destination instantly."
StrQuestions(2) = "FAX reaches its destination instantly."
StrQuestions(3) = "A computer is an electronic data processing machine used by several persons at a time."
StrQuestions(4) = "Now a days only digital computes are used."
StrQuestions(5) = "The measure of speed is also given in terms of speed of the microprocessor."
StrQuestions(6) = "A microprocessor based computer with a higher clock frequency."
StrQuestions(7) = "Intermediate results of processing are held at a storage unit."
StrQuestions(8) = "An input interface is responsible for conversion of instruction and data to computer accessible form."
StrQuestions(9) = "The job of an output nit is just the reverse of the job of an input unit."
StrQuestions(10) = "Allocation of the processor to the job is one of the functions of an Operating System."

stranswers(1) = "False"
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
Private Sub TrueorFalse_Paper4()
StrQuestions(1) = "A single user Operating system does not have a memory management module."
StrQuestions(2) = "To start an accessory after positioning the pointer double clicks the left mouse button."
StrQuestions(3) = "The kernel performs the basic functions of an Operating System."
StrQuestions(4) = "In a single user Operating system, no protection is required."
StrQuestions(5) = "A time sharing Operating System is an interactive multiprogrammed system with job scheduling and several users"
StrQuestions(6) = "A program, of size greater than the size of the main memory cannot be run."
StrQuestions(7) = "A clock is a character device."
StrQuestions(8) = "The UNIX Operating System is better at device independence than MS SOS."
StrQuestions(9) = "In UNIX, the shared file concept is possible."
StrQuestions(10) = "Netware requires DOS to run."

stranswers(1) = "False"
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
Private Sub TrueorFalse_Paper5()
StrQuestions(1) = "In DOS, the file to be typed if not found in active directory, is searched in directories give by the PATH command."
StrQuestions(2) = "Netware services depend on client/server protocols."
StrQuestions(3) = "PATH NAME in DOS can contains any a number of characters."
StrQuestions(4) = "In DOS< a subdirectory may contain any number of files and subdirectories."
StrQuestions(5) = "In DOS, XCOPY command with /M option the source disk cannot be write protected."
StrQuestions(6) = "Backup of files in DOS use archive bits."
StrQuestions(7) = "DIR command will list all files in the directory."
StrQuestions(8) = "Renaming a file means moving a file to another directory."
StrQuestions(9) = "COPY command can be used to create a new file."
StrQuestions(10) = "CD command of DOS and UNIX behave in an identical manner."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"

End Sub
Private Sub TrueorFalse_Paper6()
StrQuestions(1) = "UNIX does not have internal command."
StrQuestions(2) = "UNIX is newer than MS-DOS."
StrQuestions(3) = "Many UNIX filenames cannot be used in DOS environment."
StrQuestions(4) = "Execute permission o a directory means search is permitted in the directory for a file."
StrQuestions(5) = "In UNIX, a directory name can have an extension."
StrQuestions(6) = "Several Windows can be opened at one time."
StrQuestions(7) = "It is not possible to save keystrokes or mouse movements."
StrQuestions(8) = "Windows control panel offers a complete set of utilities to control operation of the computer and WINDOWS."
StrQuestions(9) = "Windows can overlaps one another."
StrQuestions(10) = "Network support is not possible for WINDOWS."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"

End Sub
Private Sub TrueorFalse_Paper7()
StrQuestions(1) = "In 386 enhanced mode, one program in a Window may be affected by the erratic behavior of another program."
StrQuestions(2) = "WINDOWS can be run on a 8086/88 machine in restricted way."
StrQuestions(3) = "No two dialog boxes in same application look exactly same."
StrQuestions(4) = "It is not possible to close a dialog box without making changes."
StrQuestions(5) = "In a time sharing system, the speed of communication to a computer, when a user is entering program and data, is very fast."
StrQuestions(6) = "In LAN,  each computer can fulfill a function."
StrQuestions(7) = "Sharing of executable files in LAN, need not result in failure of the whole network."
StrQuestions(8) = "Non operability of a computer, in LAN need not result in failure of the whole network."
StrQuestions(9) = "There is no difference between a LAN and a computer network."
StrQuestions(10) = "Normally, the transfer rate of data in a computer network is not under he control of the owner of a computer."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"

End Sub
Private Sub TrueorFalse_Paper8()
StrQuestions(1) = "The size of a word depends on the architecture of the computer."
StrQuestions(2) = "IEEE is concerned with the mechanical standards of microcomputer systems."
StrQuestions(3) = "Typical speed of digital signal transmission, over telephone lines is 1200 Bps."
StrQuestions(4) = "Co axial cables are capable of transmitting digital signals at a very high rate of 10 Mbps."
StrQuestions(5) = "Electromagnetic waves at microwave frequencies cannot bend round hills."
StrQuestions(6) = "Microwave transmission is possible for long distances with out repeaters etc."
StrQuestions(7) = "Received signals are amplified by a satellite."
StrQuestions(8) = "The transmitting frequency and receiving frequency signals are the same in satellite communication."
StrQuestions(9) = "Data transmission using optical fiber is based on the principle of total internal selection."
StrQuestions(10) = "Parallel communication channels are normally used for communication between two computers situated at distant places."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"

End Sub
Private Sub TrueorFalse_Paper9()
StrQuestions(1) = "Higher bandwidth allows transmission of more data."
StrQuestions(2) = "Data transmission rate expressed in bits per second and baud is same."
StrQuestions(3) = "Digital data as such, can be sent only on analog facility."
StrQuestions(4) = "Two modems are required when analog signal facility is used between two digital devices."
StrQuestions(5) = "Modem may reduce the effects of noise and distortion."
StrQuestions(6) = "The PCAT uses 8 bit data bus."
StrQuestions(7) = "For long distance data communication using telephone facility, we use serial data transfer."
StrQuestions(8) = "Serial data transfer moves data faster than parallel data transfer."
StrQuestions(9) = "Full duplex allows data to flow in both directions but not simultaneously."
StrQuestions(10) = "Echo of characters through software is a substitute for switching the full duplex modem to half duplex."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "True"

End Sub
Private Sub TrueorFalse_Paper10()
StrQuestions(1) = "The information added or deleted by communication hardware for synchronization is transparent to the user."
StrQuestions(2) = "Binary synchronous communication requires sending and receiving devices be synchronized  before transmission of data."
StrQuestions(3) = "Asynchronous data contains only one stop bit."
StrQuestions(4) = "Synchronous information is always provided within the block."
StrQuestions(5) = "Synchronous method allows the user to transmit data that is now made up of 8 bits."
StrQuestions(6) = "Deceasing duration of bit signal results in increased baud rate."
StrQuestions(7) = "Synchronous communication protocols are more complex than asynchronous communication protocols."
StrQuestions(8) = "Storage capacity of magnetic tape is almost unlimited."
StrQuestions(9) = "In HDLC/SDLC, break in data transmission of information field results in error."
StrQuestions(10) = "A multiplexer brings several high speed channels together."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub
Private Sub TrueorFalse_Paper11()
StrQuestions(1) = "In FDM, the transmission signals must be analog."
StrQuestions(2) = "TDM can only be used with analog signal."
StrQuestions(3) = "CD ROM are produced on amass a scale."
StrQuestions(4) = "LAN is a special case of a multi-user system."
StrQuestions(5) = "Servers are always PCs."
StrQuestions(6) = "Communications methodologies of a LAN are transparent to user applications."
StrQuestions(7) = "The NIC may contain firmware in ROM."
StrQuestions(8) = "Signaling in LAN is the responsibility of NIC."
StrQuestions(9) = "Base band in LAN is the responsibility of NIC."
StrQuestions(10) = "Arc net is associated with IEEE standard 802.3 CSMA/CD."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"

End Sub
Private Sub TrueorFalse_Paper12()
StrQuestions(1) = "Arc net is well suited for multi floor installation"
StrQuestions(2) = "ISDN allows transmission or reception of telegraph , telex, music and facsimile over the same pair of wires."
StrQuestions(3) = "Client/ server architecture is a special form of distributed processing."
StrQuestions(4) = "Internet is owned by the United States of America."
StrQuestions(5) = "To maximize the speed of computation, each address location, in the supercomputer holds 64 bits of information."
StrQuestions(6) = "A supercomputer has the computing capacity of less than 40000 IBM Personal Computers"
StrQuestions(7) = "Workstations are extensively used for numeric and graphics applications."
StrQuestions(8) = "Workstations are extensively used for numeric and graphics applications."
StrQuestions(9) = "Most workstations do not have a built in capacity to connect to a LAN."
StrQuestions(10) = "IBM PC or IBM compatible PCs are most popular computers."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"

End Sub
Private Sub TrueorFalse_Paper13()
StrQuestions(1) = "A communication network , connecting computers in a distributed system is slow."
StrQuestions(2) = "CD - ROM used concentric tracks to store data."
StrQuestions(3) = "Parallel processing cannot be achieved through a single processor."
StrQuestions(4) = "It is easier to write a program in machine language rather than in assembly Language."
StrQuestions(5) = "Assembly language programmers must know the logical structure of the computer."
StrQuestions(6) = "All hardware features are available to an assembly language programmer."
StrQuestions(7) = "An assembler is a machine language program."
StrQuestions(8) = "All features of the computer are available to a high level language programmer."
StrQuestions(9) = "High level languages are independent of the structure of a specific computer."
StrQuestions(10) = "An interpreter requires large memory in the computer."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"

End Sub
Private Sub TrueorFalse_Paper14()
StrQuestions(1) = "A compiler is a more complex program than an interpreter ."
StrQuestions(2) = "An interpreted program runs faster than a compiled program."
StrQuestions(3) = "Processing is done in the primary storage unit."
StrQuestions(4) = "Fetching of an instruction for execution is done by the decisions."
StrQuestions(5) = "Privacy protection depends on Judicial and Legislative decisions."
StrQuestions(6) = "Accumulator is available to an assembly language program executed."
StrQuestions(7) = "Program counter holds the address of next instruction to he executed."
StrQuestions(8) = "Memory is used to store data, programs and results."
StrQuestions(9) = "Semiconductor RAMs are mostly used as main memory."
StrQuestions(10) = "CPU can read information directly from secondary memory."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"

End Sub
Private Sub TrueorFalse_Paper15()
StrQuestions(1) = "Secondary memory is slower than main memory"
StrQuestions(2) = "Bits is the abbreviation of binary digits"
StrQuestions(3) = "A byte usually consists of 8 bits."
StrQuestions(4) = "Light pens and joysticks are both pointing devices"
StrQuestions(5) = "ROM is also used to store trigonometric functions."
StrQuestions(6) = "Even though the power supply is on, the RAM may loose its stored information."
StrQuestions(7) = "Multiprocessor systems and networks are the same."
StrQuestions(8) = "A smart terminal has in built processing capability."
StrQuestions(9) = "Only executable files can be infected by virus."
StrQuestions(10) = "One measure to protect computers, in the organization, from virus, is putting a ban on floppies being brought in by users without observing standard operating procedures covering such a contingency an penalizing non observance."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub multiplechoice_Paper1()
strmultiple(1) = "Information is "
strmuloption(1) = "Same as data."
strmuloption(2) = "Summarization process of data."
strmuloption(3) = "Textual form of data."
strmuloption(4) = "None of these"

strmultiple(2) = "Data are"
strmuloption(5) = "Raw facts and figures"
strmuloption(6) = "Information"
strmuloption(7) = "Electronic representation of facts"
strmuloption(8) = "None of these"

strmultiple(3) = "A structured decision"
strmuloption(9) = "ls made after completion of a preplanned series of steps."
strmuloption(10) = "Depends on information not known in advance."
strmuloption(11) = "ls not routine."
strmuloption(12) = "None of these"

strmultiple(4) = "Managers performing strategic and technical role need"
strmuloption(13) = "Internal information only"
strmuloption(14) = "External information only"
strmuloption(15) = "Both, internal and external information"
strmuloption(16) = "None of these"

strmultiple(5) = "Top level managers use"
strmuloption(17) = "Strategic information"
strmuloption(18) = "Tactical information"
strmuloption(19) = "Operational information"
strmuloption(20) = "None of these"

strmultiple(6) = "Corporate information"
strmuloption(21) = "Brings healthy relationship between employees and employers."
strmuloption(22) = "Contains personal particulars."
strmuloption(23) = "Means information derived from weekly meetings."
strmuloption(24) = "None of these"

strmultiple(7) = "Entities, attributes and relationships are associated with"
strmuloption(25) = "Logical concepts of data"
strmuloption(26) = "Physical concepts of data."
strmuloption(27) = "Persons of organizations"
strmuloption(28) = "None of these"

strmultiple(8) = "Storage devices may be "
strmuloption(29) = "Sequential type only"
strmuloption(30) = "Direct access type only"
strmuloption(31) = "Sequential type or direct access type"
strmuloption(32) = "None of these"

strmultiple(9) = "Blocking factor is "
strmuloption(33) = "The number of records grouped in a block."
strmuloption(34) = "The number of blocks in a file."
strmuloption(35) = "Ratio between number of records accessed and total number of records"
strmuloption(36) = "None of these"

strmultiple(10) = "The number of read/write heads in a magnetic hard disk is equal to"
strmuloption(37) = "total number of disk surfaces minus two."
strmuloption(38) = "Total number of Disk surfaces."
strmuloption(39) = "One"
strmuloption(40) = "None of these"

strmulanswer(1) = "Summarization process of data."
strmulanswer(2) = "Raw facts and figures"
strmulanswer(3) = "ls made after completion of a preplanned series of steps."
strmulanswer(4) = "Both, internal and external information"
strmulanswer(5) = "Strategic information"
strmulanswer(6) = "Brings healthy relationship between employees and employers."
strmulanswer(7) = "Logical concepts of data"
strmulanswer(8) = "Sequential type or direct access type"
strmulanswer(9) = "The number of records grouped in a block."
strmulanswer(10) = "total number of disk surfaces minus two."
End Sub

Private Sub multiplechoice_Paper2()
strmultiple(1) = "File volatility means"
strmuloption(1) = "Highly active file."
strmuloption(2) = "Rate at which record are added and deleted from files."
strmuloption(3) = "Number of updates to the records in file."
strmuloption(4) = "None of these."

strmultiple(2) = "Data processing cycle consists of "
strmuloption(5) = "Input cycle and output cycle"
strmuloption(6) = "Output cycle and processing cycle."
strmuloption(7) = "Input cycle, output cycle an processing cycle."
strmuloption(8) = "None of these"

strmultiple(3) = "Transaction processing is called on line processing if."
strmuloption(9) = "Transaction is processed as soon as it enters in the system."
strmuloption(10) = "Transaction is processes in batches  "
strmuloption(11) = "Transaction is processed in a running system"
strmuloption(12) = "None of these"

strmultiple(4) = "A Distributed data processing system"
strmuloption(13) = "Attempts to capture advantage of both centralized and decentralized processing"
strmuloption(14) = "Does not allow greater flexibilities"
strmuloption(15) = "Allow access to data."
strmuloption(16) = "None of these"

strmultiple(5) = "Multiprocessing means"
strmuloption(17) = "Simultaneous processing with two or more processes."
strmuloption(18) = "Simultaneous running of two or more programs."
strmuloption(19) = "Running of two or more processes."
strmuloption(20) = "None of these"

strmultiple(6) = "In data processing, classifying, calculating, sorting and summarizing data are connected with"
strmuloption(21) = "Capturing of input data"
strmuloption(22) = "Manipulating of data"
strmuloption(23) = "Managing output results"
strmuloption(24) = "None of these"

strmultiple(7) = "Prototyping means"
strmuloption(25) = "Creating, developing and refining a working model of the final operational system."
strmuloption(26) = "Testing the compute system"
strmuloption(27) = "Designing the compute system"
strmuloption(28) = "None of these"

strmultiple(8) = "A Word Processor is"
strmuloption(29) = "A Program which processes the word e.g. giving frequency of words, sorting of works etc."
strmuloption(30) = "CPU of microcomputer."
strmuloption(31) = "A program which provides the facility for creating, viewing editing and manipulating test before printing."
strmuloption(32) = "None of these"

strmultiple(9) = "Which of the following is a graphical package?"
strmuloption(33) = "CorelDraw"
strmuloption(34) = "Lotus 1-2-3"
strmuloption(35) = "WordStar"
strmuloption(36) = "None of these"

strmultiple(10) = "E-mail is"
strmuloption(37) = "Mail concerning electronics devices"
strmuloption(38) = "Transaction of letters, messages and memos over a communications network."
strmuloption(39) = "Transaction of messages within a computer."
strmuloption(40) = "None of these"

strmulanswer(1) = "Rate at which record are added and deleted from files."
strmulanswer(2) = "Input cycle, output cycle an processing cycle."
strmulanswer(3) = "Transaction is processed as soon as it enters in the system."
strmulanswer(4) = "Attempts to capture advantage of both centralized and decentralized processing"
strmulanswer(5) = "Simultaneous processing with two or more processes."
strmulanswer(6) = "Manipulating of data"
strmulanswer(7) = "Creating, developing and refining a working model of the final operational system."
strmulanswer(8) = "A program which provides the facility for creating, viewing editing and manipulating test before printing."
strmulanswer(9) = "CorelDraw"
strmulanswer(10) = "Transaction of letters, messages and memos over a communications network."
End Sub

Private Sub multiplechoice_Paper3()
strmultiple(1) = "The computer as a machine and all other equipment associated with it, are termed as"
strmuloption(1) = "Hardware"
strmuloption(2) = "Software"
strmuloption(3) = "Firmware"
strmuloption(4) = "None of these"

strmultiple(2) = "Which of the following devices can be used to directly input printed text?"
strmuloption(5) = "OCR"
strmuloption(6) = "OMR"
strmuloption(7) = "MICR"
strmuloption(8) = "None of these"

strmultiple(3) = "Computer clock speed"
strmuloption(9) = "Is the speed of the inner clock of a computer."
strmuloption(10) = "Gives the frequency at which pulses are generated and processed."
strmuloption(11) = "Is speed of internal buses of computer."
strmuloption(12) = "None of these"

strmultiple(4) = "Holding of all data to be processed and instructions required for processing is one of the functions of"
strmuloption(13) = "CPU"
strmuloption(14) = "Storage unit"
strmuloption(15) = "Output unit"
strmuloption(16) = "None of these"

strmultiple(5) = "CPU reads the information from secondary memory"
strmuloption(17) = "Directly"
strmuloption(18) = "First information is transferred to main memory and from there, the CPU reads"
strmuloption(19) = "Through registers"
strmuloption(20) = "None of these"

strmultiple(6) = "Conversion of computer generated results not human acceptable form is a function of "
strmuloption(21) = "CPU"
strmuloption(22) = "Input interface"
strmuloption(23) = "Output interface"
strmuloption(24) = "None of these"

strmultiple(7) = "The functions of an Operating System as a resource manager may be viewed as"
strmuloption(25) = "Top down view"
strmuloption(26) = "Bottom up view"
strmuloption(27) = "Both a and b above"
strmuloption(28) = "None of these"

strmultiple(8) = "The single user Operating System does not have"
strmuloption(29) = "Memory management module"
strmuloption(30) = "CPU scheduling module"
strmuloption(31) = "Device management module"
strmuloption(32) = "None of these"

strmultiple(9) = "The part of an Operating System which permanently resides in main memory is"
strmuloption(33) = "Job scheduling"
strmuloption(34) = "Kernel"
strmuloption(35) = "Translator"
strmuloption(36) = "none of these"

strmultiple(10) = "The protection feature of an Operating System, is required in"
strmuloption(37) = "Single user system only"
strmuloption(38) = "Multi user system only"
strmuloption(39) = "Both a and b above"
strmuloption(40) = "None of these"

strmulanswer(1) = "a. Hardware"
strmulanswer(2) = "MICR"
strmulanswer(3) = "Gives the frequency at which pulses are generated and processed."
strmulanswer(4) = "Storage unit"
strmulanswer(5) = "First information is transferred to main memory and from there, the CPU reads"
strmulanswer(6) = "Output interface"
strmulanswer(7) = "Bottom up view"
strmulanswer(8) = "CPU scheduling module"
strmulanswer(9) = "Kernel"
strmulanswer(10) = "Both a and b above"
End Sub

Private Sub multiplechoice_Paper4()
strmultiple(1) = "A Time sharing system is "
strmuloption(1) = "multiprogrammed batch system"
strmuloption(2) = "Single user system used by several users"
strmuloption(3) = "Interactive multiprogrammed system with several users"
strmuloption(4) = "None of these"

strmultiple(2) = "Virtual memory is"
strmuloption(5) = "Part of main memory only used for swapping."
strmuloption(6) = "A technique to allow a program of size more than the size of the main memory to run"
strmuloption(7) = "Part of secondary storage used in program execution"
strmuloption(8) = "None of these"

strmultiple(3) = "Block device"
strmuloption(9) = "Delivers or accepts a stream of characters without any structure."
strmuloption(10) = "Stores information in fixed sized blocks"
strmuloption(11) = "Is an output device to print information in blocks"
strmuloption(12) = "None of these"

strmultiple(4) = "Device independence is associated with"
strmuloption(13) = "File Systems"
strmuloption(14) = "Device drivers"
strmuloption(15) = "Virtual systems"
strmuloption(16) = "None of these"

strmultiple(5) = "Shared file needs a file system with"
strmuloption(17) = "Single directory structure"
strmuloption(18) = "Tree structure"
strmuloption(19) = "Directed cyclic graph"
strmuloption(20) = "None of these"

strmultiple(6) = "A program X, not in active directory in DOS< cannot be executed because"
strmuloption(21) = "An executable program must be in active directory"
strmuloption(22) = "Directory containing X must be specified through PATH command "
strmuloption(23) = "It does not have executable permission"
strmuloption(24) = "None of these"

strmultiple(7) = "Netware supports"
strmuloption(25) = "IBM PC compatible servers only"
strmuloption(26) = "APPLE Macintosh servers only"
strmuloption(27) = "Both a and b above"
strmuloption(28) = "None of these"

strmultiple(8) = "To protect computers from virus we may use"
strmuloption(29) = "SW "
strmuloption(30) = "HW"
strmuloption(31) = "Both HW and SW "
strmuloption(32) = "None of these"

strmultiple(9) = "In DOS, cd command without any operand would"
strmuloption(33) = "Result in error (insufficient operand)"
strmuloption(34) = "Show the pathname of current active directory"
strmuloption(35) = "Show the active drive"
strmuloption(36) = "None of these"

strmultiple(10) = "In DOS, A>C, will take you to"
strmuloption(37) = "Drive C with root as active drive"
strmuloption(38) = "Drive C with its previous active directory"
strmuloption(39) = "Drive C with active directory A."
strmuloption(40) = "None of these"

strmulanswer(1) = "Interactive multiprogrammed system with several users"
strmulanswer(2) = "A technique to allow a program of size more than the size of the main memory to run"
strmulanswer(3) = "Stores information in fixed sized blocks"
strmulanswer(4) = "Virtual systems"
strmulanswer(5) = "Single directory structure"
strmulanswer(6) = "Directory containing X must be specified through PATH command "
strmulanswer(7) = "Both a and b above"
strmulanswer(8) = "Both HW and SW "
strmulanswer(9) = "Show the active drive"
strmulanswer(10) = "Drive C with root as active drive"
End Sub

Private Sub multiplechoice_Paper5()
strmultiple(1) = "C > DIR without any argument in MS-DOS will"
strmuloption(1) = "List files and subdirectories of root directory of C drive"
strmuloption(2) = "Drive C with its previous active directory."
strmuloption(3) = "Result in error"
strmuloption(4) = "None of these"

strmultiple(2) = "C > PRINT without any argument in MS - DOS will"
strmuloption(5) = "Print all files of active directory of drive C"
strmuloption(6) = "Result in error"
strmuloption(7) = "Print the names of files which are being printed and are in queue for printing."
strmuloption(8) = "None of these"

strmultiple(3) = "Concatenating binary files and text files together through COPY command, in DOS"
strmuloption(9) = "Is same as concatenating text files"
strmuloption(10) = "Is same as concatenating binary files."
strmuloption(11) = "Requires identification of binary files and text files through appropriate options."
strmuloption(12) = "None of these"

strmultiple(4) = "Group of files may be referenced by the DOS command"
strmuloption(13) = "Listing each file of the group"
strmuloption(14) = "Using wild cards"
strmuloption(15) = "By no means"
strmuloption(16) = "None of these"

strmultiple(5) = "Total memory of system can be known using"
strmuloption(17) = "MEM commands"
strmuloption(18) = "DIR commands"
strmuloption(19) = "Warm boot"
strmuloption(20) = "None of these"

strmultiple(6) = "Rename command is used to"
strmuloption(21) = "Name the file to another directory."
strmuloption(22) = "Move the file to same directory on another drive"
strmuloption(23) = "Change the name of the file in directory"
strmuloption(24) = "None of these"

strmultiple(7) = "In UNIX, $A:ls will be taken as"
strmuloption(25) = Chr(4) & "ls" & Chr(4) & "command to be read from drive A and executed"
strmuloption(26) = Chr(4) & "a:ls" & Chr(4) & "command to be executed"
strmuloption(27) = "Error, as directory A containing ls is not specified"
strmuloption(28) = "None of these"

strmultiple(8) = "In UNIX, cat a.s.v."
strmuloption(29) = "Will list the contents of the file" & Chr(4) & "a.s.v." & Chr(4)
strmuloption(30) = "Will display" & Chr(4) & "a.s.v." & Chr(4) & "invalid filename."
strmuloption(31) = "Will list the contents of file" & Chr(4) & "a." & Chr(4)
strmuloption(32) = "None of these"

strmultiple(9) = "In UNIX, cd command without argument would"
strmuloption(33) = "Result in error(insufficient operands)"
strmuloption(34) = "Show the pathname is current directory."
strmuloption(35) = "Change the current directory to the home directory"
strmuloption(36) = "None of these"

strmultiple(10) = "You can become superuser if you know"
strmuloption(37) = "Password of the current user"
strmuloption(38) = "Root password"
strmuloption(39) = "Name of the system administrator"
strmuloption(40) = "None of these"

strmulanswer(1) = "Drive C with its previous active directory."
strmulanswer(2) = "Result in error"
strmulanswer(3) = "Requires identification of binary files and text files through appropriate options."
strmulanswer(4) = "Using wild cards"
strmulanswer(5) = "MEM commands"
strmulanswer(6) = "Change the name of the file in directory"
strmulanswer(7) = "Error, as directory A containing ls is not specified"
strmulanswer(8) = "Will list the contents of the file" & Chr(4) & "a.s.v." & Chr(4)
strmulanswer(9) = "Change the current directory to the home directory"
strmulanswer(10) = "Root password"
End Sub

Private Sub multiplechoice_Paper6()
strmultiple(1) = "In UNIX, a file cannot be deleted if"
strmuloption(1) = "The file does not have write permission for the user."
strmuloption(2) = "Even if the file has write permission for the user but the patent directory does not have write permission for the user."
strmuloption(3) = "The file does not have permission of any user (i.e. owner, group, others)."
strmuloption(4) = "None of these"

strmultiple(2) = "Which one is not the software written for WINDOWS environment i.e. WINDOWS application?"
strmuloption(5) = "ALDUS PageMaker"
strmuloption(6) = "MICROSOFT EXCEL"
strmuloption(7) = "Micrograph Designer"
strmuloption(8) = "None of these"

strmultiple(3) = "WINDOWS is "
strmuloption(9) = "An operating System"
strmuloption(10) = "An environment under DOS."
strmuloption(11) = "An environment under UNIX"
strmuloption(12) = "None of these"

strmultiple(4) = "Which IS not desktop accessory?"
strmuloption(13) = "Terminal"
strmuloption(14) = "Calendar"
strmuloption(15) = "Clipboard"
strmuloption(16) = "None of these"

strmultiple(5) = "Crashing of one program may not affect the other program in WINDOWS running at"
strmuloption(17) = "Real mode"
strmuloption(18) = "Standard mode"
strmuloption(19) = "386 enhanced mode"
strmuloption(20) = "None of these"

strmultiple(6) = "A new printer caN be added by printer option is"
strmuloption(21) = "WINDOW control panel"
strmuloption(22) = "File manager"
strmuloption(23) = "Dynamic data exchange"
strmuloption(24) = "None of these"

strmultiple(7) = "For running WNDOWS minimum disk free space required is"
strmuloption(25) = "8 MB"
strmuloption(26) = "10 MB"
strmuloption(27) = "16 MB"
strmuloption(28) = "none of these"

strmultiple(8) = "In inactive WINDOWS which of the following is not true?"
strmuloption(29) = "Application is opened bit not running unless 386 enhanced modes is used"
strmuloption(30) = "Loaded applications are available for immediate use"
strmuloption(31) = "There are loaded icons"
strmuloption(32) = "None of these"

strmultiple(9) = "Commands in menu bar of program manager, affect one subroutine of WINDOWS which is "
strmuloption(33) = "Inactive"
strmuloption(34) = "Active"
strmuloption(35) = "Just below the top Window"
strmuloption(36) = "When data and program are entered manually through terminals which are connected to a computer, there is"

strmultiple(10) = "When data and program are entered manually through terminals which are connected to a computer, there is"
strmuloption(37) = "Not heavy demands on communications line"
strmuloption(38) = "Very heavy demands on communications line "
strmuloption(39) = "Very heavy or very less demands depending upon use of the terminals"
strmuloption(40) = "None of these"

strmulanswer(1) = "The file does not have permission of any user (i.e. owner, group, others)."
strmulanswer(2) = "None of these"
strmulanswer(3) = "An operating System"
strmulanswer(4) = "Terminal"
strmulanswer(5) = "Standard mode"
strmulanswer(6) = "WINDOW control panel"
strmulanswer(7) = "8 MB"
strmulanswer(8) = "None of these"
strmulanswer(9) = "Just below the top Window"
strmulanswer(10) = "None of these"
End Sub

Private Sub multiplechoice_Paper7()
strmultiple(1) = "In LAN communicating between different machines is"
strmuloption(1) = "Slow"
strmuloption(2) = "FaSt"
strmuloption(3) = "Slow or fast depending upon traffic"
strmuloption(4) = "None of these"

strmultiple(2) = "Break down of a computer in LAN will result in "
strmuloption(5) = "Total failure of network"
strmuloption(6) = "Only non availability of that non operatable computers. "
strmuloption(7) = "Total failure of network or non availability depending on its topology"
strmuloption(8) = "None of these"

strmultiple(3) = "Network formed between computers spread in large geographical area is called"
strmuloption(9) = "LAN"
strmuloption(10) = "Computer Network"
strmuloption(11) = "Public network"
strmuloption(12) = "One of these"

strmultiple(4) = "Rules and combinations, followed in design and fabrication of devices for external communication between PC's is framed but"
strmuloption(13) = "CSI"
strmuloption(14) = "AT&T"
strmuloption(15) = "EIA"
strmuloption(16) = "None of these"

strmultiple(5) = "Electrical characteristics of connections between PC's and external peripherals are governed but"
strmuloption(17) = "Intel 8086 standard"
strmuloption(18) = "RS-232 standard"
strmuloption(19) = "ISO standard "
strmuloption(20) = "None of these"

strmultiple(6) = "To avoid noise interference wire pairs used in communication are -"
strmuloption(21) = "Kept parallel"
strmuloption(22) = "Twisted"
strmuloption(23) = "Kept in cover"
strmuloption(24) = "None of these"

strmultiple(7) = "To provide higher noise immunity in data communication we select - "
strmuloption(25) = "Co axial cables"
strmuloption(26) = "Pair of wires"
strmuloption(27) = "Parallel transmission "
strmuloption(28) = "None of these"

strmultiple(8) = "For data communication cable lying is not required when using"
strmuloption(29) = "Microwave transmission"
strmuloption(30) = "Co axial cable"
strmuloption(31) = "Optical fiber"
strmuloption(32) = "None of these"

strmultiple(9) = "Faster data communication is possible when using"
strmuloption(33) = "CO axial cable"
strmuloption(34) = "Microwave transmission"
strmuloption(35) = "Optical fiber"
strmuloption(36) = "None of these"

strmultiple(10) = "Computes send data to peripherals with the help of "
strmuloption(37) = "Serial I/O interface"
strmuloption(38) = "Parallel I/O interface"
strmuloption(39) = "Modem"
strmuloption(40) = "None of these"

strmulanswer(1) = "Slow or fast depending upon traffic"
strmulanswer(2) = "Total failure of network or non availability depending on its topology"
strmulanswer(3) = "Public network"
strmulanswer(4) = "CSI"
strmulanswer(5) = "ISO standard "
strmulanswer(6) = "Kept in cover"
strmulanswer(7) = "Parallel transmission "
strmulanswer(8) = "Optical fiber"
strmulanswer(9) = "Microwave transmission"
strmulanswer(10) = "Parallel I/O interface"
End Sub

Private Sub multiplechoice_Paper8()
strmultiple(1) = "Bandwidth is"
strmuloption(1) = "The difference between the lowest and the highest frequency transmitted"
strmuloption(2) = "Union of distortion"
strmuloption(3) = "Transmission speed of channel"
strmuloption(4) = "None of these"

strmultiple(2) = "Baud rate is "
strmuloption(5) = "The diference between the lowest and the highest frequency transmitted"
strmuloption(6) = "Transmission capacity"
strmuloption(7) = "Transmission speed of channel"
strmuloption(8) = "None of these"

strmultiple(3) = "Modem is used to "
strmuloption(9) = "Convert digital signal to analog signal only"
strmuloption(10) = "Convert analog signal to digital signal only"
strmuloption(11) = "Convert digital signal to analog signal and vice versa "
strmuloption(12) = "None of these"

strmultiple(4) = "A smart modem is one which"
strmuloption(13) = "Can vary transmission rate"
strmuloption(14) = "Contains acoustic coupler"
strmuloption(15) = "Contains microprocessor"
strmuloption(16) = "None of these"

strmultiple(5) = "Data bus is used in"
strmuloption(17) = "Parallel data transfer"
strmuloption(18) = "Serial data transfer"
strmuloption(19) = "Computers using analog signal"
strmuloption(20) = "none of these"

strmultiple(6) = "Unidirectional datanys is called"
strmuloption(21) = "Simplex configuration"
strmuloption(22) = "Half duplex"
strmuloption(23) = "Full duplex"
strmuloption(24) = "None of these"

strmultiple(7) = "Local echo means"
strmuloption(25) = "When in duplex mode, mismatch exists between both modems and produces echo of character"
strmuloption(26) = "Display of character through software"
strmuloption(27) = "Display of character in half duplex"
strmuloption(28) = "None of these"

strmultiple(8) = "Synchronization between transmitter and receiver"
strmuloption(29) = "Automatically takes place"
strmuloption(30) = "Is modem generated"
strmuloption(31) = "Is governed by rules & conventions"
strmuloption(32) = "None of these"

strmultiple(9) = "Asynchronous method of data link control is used in"
strmuloption(33) = "TTY protocol"
strmuloption(34) = "BSE"
strmuloption(35) = "High-level data link control"
strmuloption(36) = "None of these"

strmultiple(10) = "Communication protocol is "
strmuloption(37) = "Hardware and Software standard for transmitting data between peripherals and computers"
strmuloption(38) = "Rules to synchronous transmitting and receiving devices"
strmuloption(39) = "Rules to set clock contained in hardware at sending or receiving station."
strmuloption(40) = "None of these"

strmulanswer(1) = "The difference between the lowest and the highest frequency transmitted"
strmulanswer(2) = "Transmission speed of channel"
strmulanswer(3) = "Convert digital signal to analog signal and vice versa "
strmulanswer(4) = "Can vary transmission rate"
strmulanswer(5) = "Parallel data transfer"
strmulanswer(6) = "Half duplex"
strmulanswer(7) = "When in duplex mode, mismatch exists between both modems and produces echo of character"
strmulanswer(8) = "Is governed by rules & conventions"
strmulanswer(9) = "High-level data link control"
strmulanswer(10) = "Rules to set clock contained in hardware at sending or receiving station."
End Sub

Private Sub multiplechoice_Paper9()
strmultiple(1) = "A Synchronous protocol can be"
strmuloption(1) = "Character oriented null"
strmuloption(2) = "Bit oriented only"
strmuloption(3) = "Either a or b above"
strmuloption(4) = "None of these"

strmultiple(2) = "Data transfer rate in serial communication related to duration of voltage signal used to represent the binary digit, is given as"
strmuloption(5) = "Bits per second"
strmuloption(6) = "Baud"
strmuloption(7) = "Bytes per second"
strmuloption(8) = "None of these"

strmultiple(3) = "In binary synchronous communication, start of data is signalled by"
strmuloption(9) = "Synchronous idle characters"
strmuloption(10) = "Flag bits"
strmuloption(11) = "Synchronous clock"
strmuloption(12) = "None of these"

strmultiple(4) = "Which of the following is true?"
strmuloption(13) = "Pauses in data transmission in SDLC/HDLC "
strmuloption(14) = "SDLC/HDLC has control character equivalent to SYN character of BSC"
strmuloption(15) = "SDLC/HDLC requires continuous transmission of information field"
strmuloption(16) = "None of these"

strmultiple(5) = "A multiplexer"
strmuloption(17) = "Boosts the strength of signal"
strmuloption(18) = "Brings several low speed channels together and transforms them into one single high speed channel"
strmuloption(19) = "Is a device to eliminate noise from signal"
strmuloption(20) = "None of these"

strmultiple(6) = "A concentrator is"
strmuloption(21) = "Another name for a multiplexer"
strmuloption(22) = "A device to strengthen the signal"
strmuloption(23) = "An intelligent multiplexer"
strmuloption(24) = "None of these"

strmultiple(7) = "Which is not true for LAN?"
strmuloption(25) = "LAN are public networks"
strmuloption(26) = "LANs are designed to work in a restricted geographical area"
strmuloption(27) = "LANs operate at higher speed than WANs "
strmuloption(28) = "None of these"

strmultiple(8) = "A server in LAN is a"
strmuloption(29) = "Device offering service to user"
strmuloption(30) = "NIC"
strmuloption(31) = "Workstation with huge secondary storage"
strmuloption(32) = "None of these"

strmultiple(9) = "Select" & Chr(4) & "Odd man out" & Chr(4) & "in the following"
strmuloption(33) = "Ethernet"
strmuloption(34) = "Arc net"
strmuloption(35) = "IBM token ring LAN"
strmuloption(36) = "None of these"

strmultiple(10) = "One of the functions of an NIC is"
strmuloption(37) = "to receive and transmit data at a speed and in the format required by network media"
strmuloption(38) = "to provide sharing of files "
strmuloption(39) = "to provide sharing of memory"
strmuloption(40) = "None of these"

strmulanswer(1) = "Either a or b above"
strmulanswer(2) = "Baud"
strmulanswer(3) = "Flag bits"
strmulanswer(4) = "SDLC/HDLC requires continuous transmission of information field"
strmulanswer(5) = "Brings several low speed channels together and transforms them into one single high speed channel"
strmulanswer(6) = "None of these"
strmulanswer(7) = "LANs are designed to work in a restricted geographical area"
strmulanswer(8) = "Device offering service to user"
strmulanswer(9) = "None of these"
strmulanswer(10) = "to receive and transmit data at a speed and in the format required by network media"
End Sub

Private Sub multiplechoice_Paper10()
strmultiple(1) = "The topology of an Ethernet system is "
strmuloption(1) = "Ring "
strmuloption(2) = "Linear bus with CSMa/CD"
strmuloption(3) = "Star"
strmuloption(4) = "None of these"

strmultiple(2) = "CSMA/CD is popularly known as"
strmuloption(5) = "Star LAN"
strmuloption(6) = "Active Repeaters "
strmuloption(7) = "Multiplexer"
strmuloption(8) = "None of these"

strmultiple(3) = "Medium access units are referred to as"
strmuloption(9) = "NIC"
strmuloption(10) = "Active Repeaters "
strmuloption(11) = "Multiplexer"
strmuloption(12) = "None of these"

strmultiple(4) = "Integrated Service Digital NETWORK (isdn) LINES Provide "
strmuloption(13) = "Only voice circuits"
strmuloption(14) = "Only data circuits"
strmuloption(15) = "Both voice and data circuits"
strmuloption(16) = "None of these"

strmultiple(5) = "For talking to Internet host computer, the computer uses"
strmuloption(17) = "TCP/IP "
strmuloption(18) = "Emulation software"
strmuloption(19) = "IPX"
strmuloption(20) = "None of these"

strmultiple(6) = "Which of the following is not network system?"
strmuloption(21) = "Netware "
strmuloption(22) = "VINES"
strmuloption(23) = "UNIX"
strmuloption(24) = "None of these"

strmultiple(7) = "Scalability modular architecture, diversity, simplicity are a few features of"
strmuloption(25) = "Operating systems of main frames"
strmuloption(26) = "Network Operating System"
strmuloption(27) = "Operating Systems of PC's"
strmuloption(28) = "None of these"

strmultiple(8) = "The gigantic work of simulating he airflow around an entire aircraft can only be done by using fast"
strmuloption(29) = "Microcomputers"
strmuloption(30) = "Minicomputers"
strmuloption(31) = "Supercomputers"
strmuloption(32) = "None of these"

strmultiple(9) = "One of the following is the supercomputer"
strmuloption(33) = "X-MP"
strmuloption(34) = "S/370"
strmuloption(35) = "NCR 9300"
strmuloption(36) = "None of these"

strmultiple(10) = " Desktop machines have"
strmuloption(37) = "4 - 8 MB main memory"
strmuloption(38) = "8 - 16 MB memory"
strmuloption(39) = "16 - 32 MB memory"
strmuloption(40) = "None of these"

strmulanswer(1) = "Star"
strmulanswer(2) = "Star LAN"
strmulanswer(3) = "None of these"
strmulanswer(4) = "Both voice and data circuits"
strmulanswer(5) = "TCP/IP "
strmulanswer(6) = "VINES"
strmulanswer(7) = "Operating Systems of PC's"
strmulanswer(8) = "Supercomputers"
strmulanswer(9) = "None of these"
strmulanswer(10) = "4 - 8 MB main memory"
End Sub

Private Sub multiplechoice_Paper11()
strmultiple(1) = "The frequency range used in satellite comunications is of the order of "
strmuloption(1) = "KHz      "
strmuloption(2) = "MHz"
strmuloption(3) = "Ghz"
strmuloption(4) = "None of these"

strmultiple(2) = "Which of the following admirable capabilities is provided in most of workstations?"
strmuloption(5) = "Numeric processing"
strmuloption(6) = "Graphics"
strmuloption(7) = "Text manipulation"
strmuloption(8) = "None of these"

strmultiple(3) = "A common use of distributed computers is "
strmuloption(9) = "Scientific computing"
strmuloption(10) = "Client server computing"
strmuloption(11) = "Business oriented computing"
strmuloption(12) = "None of these"

strmultiple(4) = "At present a supercomputer may have a maximum of "
strmuloption(13) = "2.2 GFLOPS and 2 GB main memory"
strmuloption(14) = "4.5 GFLOPS and 4 GB main memory"
strmuloption(15) = "35.2 GFLOPS and 32 GB main memory"
strmuloption(16) = "None of these"

strmultiple(5) = "Assembly language is "
strmuloption(17) = "Machine independent"
strmuloption(18) = "Machine dependent"
strmuloption(19) = "High level language"
strmuloption(20) = "None of these"

strmultiple(6) = "An assembly language program must be translated into "
strmuloption(21) = "Machine language"
strmuloption(22) = "High level language"
strmuloption(23) = "Machine independent language"
strmuloption(24) = "None of these"

strmultiple(7) = "An assembler converts source program into"
strmuloption(25) = "Executable program"
strmuloption(26) = "Object program"
strmuloption(27) = "Link program"
strmuloption(28) = "None of these"

strmultiple(8) = "A Bootstrap loading is a procedure"
strmuloption(29) = "used to test memory internally"
strmuloption(30) = "In which first few instructions load the rest of loader"
strmuloption(31) = "To read assembler from memory"
strmuloption(32) = "None of these"

strmultiple(9) = "In high level language, a statement is translated to"
strmuloption(33) = "One machine language statement "
strmuloption(34) = "Many machine language statements"
strmuloption(35) = "Several high level language statements"
strmuloption(36) = "Nome of these"

strmultiple(10) = "Translator which takes one statement of high level language at a time and translates it into machine language statements, which are immediately executed is known as"
strmuloption(37) = "Compiler"
strmuloption(38) = "Interpreter"
strmuloption(39) = "Assembler"
strmuloption(40) = "None of these"

strmulanswer(1) = "Ghz"
strmulanswer(2) = "Numeric processing"
strmulanswer(3) = "Client server computing"
strmulanswer(4) = "35.2 GFLOPS and 32 GB main memory"
strmulanswer(5) = "Machine dependent"
strmulanswer(6) = "Machine language"
strmulanswer(7) = "Object program"
strmulanswer(8) = "In which first few instructions load the rest of loader"
strmulanswer(9) = "Several high level language statements"
strmulanswer(10) = "Interpreter"
End Sub

Private Sub multiplechoice_Paper12()
strmultiple(1) = "Normally a high level language is"
strmuloption(1) = "Procedural oriented"
strmuloption(2) = "Non procedural oriented"
strmuloption(3) = "Object oriented"
strmuloption(4) = "None of these"

strmultiple(2) = "C is a "
strmuloption(5) = "Low level language"
strmuloption(6) = "Non procedural language"
strmuloption(7) = "Procedural language"
strmuloption(8) = "None of these"

strmultiple(3) = "Which of the following is not a high level language program?"
strmuloption(9) = "C"
strmuloption(10) = "C++"
strmuloption(11) = "ADA"
strmuloption(12) = "None of these"

strmultiple(4) = "Add, Subtract, Divide, Multiply and Logic operations are performed by"
strmuloption(13) = "Registers"
strmuloption(14) = "Control unit"
strmuloption(15) = "ALU "
strmuloption(16) = "None of these"

strmultiple(5) = "Interpretation of instructions in a computer is done by "
strmuloption(17) = "Registers"
strmuloption(18) = "Control unit"
strmuloption(19) = "ALU "
strmuloption(20) = "None of these"

strmultiple(6) = "A CPU consists of "
strmuloption(21) = "ALU"
strmuloption(22) = "ALU,and control unit"
strmuloption(23) = "ALU, control unit and registers"
strmuloption(24) = "None of these"

strmultiple(7) = "AND, OR, NOT Logical operations are performed in"
strmuloption(25) = "ALU"
strmuloption(26) = "Control unit"
strmuloption(27) = "Registers"
strmuloption(28) = "None of these"

strmultiple(8) = "An accumulator is used to hold"
strmuloption(29) = "Operand only before execution of instructions"
strmuloption(30) = "Result only after execution of instructions"
strmuloption(31) = "Operand and results in execution of instructions"
strmuloption(32) = "None of these"

strmultiple(9) = "A CPU holds the instruction is"
strmuloption(33) = "Program counter"
strmuloption(34) = "Index register"
strmuloption(35) = "Instruction register"
strmuloption(36) = "None of these"

strmultiple(10) = "Address of best instruction to be executed is available in"
strmuloption(37) = "Program counter"
strmuloption(38) = "Index register"
strmuloption(39) = "Instruction register"
strmuloption(40) = "None of these"

strmulanswer(1) = "Object oriented"
strmulanswer(2) = "Low level language"
strmulanswer(3) = "C"
strmulanswer(4) = "Registers"
strmulanswer(5) = "Control unit"
strmulanswer(6) = "ALU, control unit and registers"
strmulanswer(7) = "ALU"
strmulanswer(8) = "Operand and results in execution of instructions"
strmulanswer(9) = "Instruction register"
strmulanswer(10) = "Index register"
End Sub

Private Sub multiplechoice_Paper13()
strmultiple(1) = "Main memory stores"
strmuloption(1) = "Data only"
strmuloption(2) = "Program only"
strmuloption(3) = "Data, program and results or any other kind of information"
strmuloption(4) = "none of these"

strmultiple(2) = "Access time of RAM is of the order of"
strmuloption(5) = "80 - 100 micro second"
strmuloption(6) = "80 - 100 nano second"
strmuloption(7) = "80 - 100 milli second"
strmuloption(8) = "None of these"

strmultiple(3) = "Conversion of instruction and data into computer acceptable form is the function of"
strmuloption(9) = "CPU"
strmuloption(10) = "Input interface"
strmuloption(11) = "Output interface"
strmuloption(12) = "None of these"

strmultiple(4) = "Main memory"
strmuloption(13) = "Volatile"
strmuloption(14) = "Non Volatile"
strmuloption(15) = "Volatile as well as non volatile"
strmuloption(16) = "None of these"

strmultiple(5) = "Access time of a typical hard disk may be"
strmuloption(17) = "20 Micro second"
strmuloption(18) = "20 Nano second"
strmuloption(19) = "20 Milli second"
strmuloption(20) = "None of these"

strmultiple(6) = "Bits (as used in computers) is an abbreviation of"
strmuloption(21) = "Binary Digits"
strmuloption(22) = "Birla institute of Technology"
strmuloption(23) = "British Institute of Technology"
strmuloption(24) = "None of these"

strmultiple(7) = "RAM is"
strmuloption(25) = "Read Only Memory"
strmuloption(26) = "Write Only memory"
strmuloption(27) = "Read/Write Memory"
strmuloption(28) = "None of these"

strmultiple(8) = "A Multiprocessor is one with"
strmuloption(29) = "One CPU and several channels"
strmuloption(30) = "Several CPU's"
strmuloption(31) = "One CPU executing several processes"
strmuloption(32) = "None of these"

strmultiple(9) = "A smart terminal is one which"
strmuloption(33) = "Looks very attractive"
strmuloption(34) = "Has inbuilt processing capability"
strmuloption(35) = "Must convert input on CRT screen"
strmuloption(36) = "None of these"

strmultiple(10) = "A light pen"
strmuloption(37) = "Is a pointing device"
strmuloption(38) = "Is an optical scanner"
strmuloption(39) = "Uses magnetic ink"
strmuloption(40) = "None of these"

strmulanswer(1) = "Data, program and results or any other kind of information"
strmulanswer(2) = "80 - 100 micro second"
strmulanswer(3) = "CPU"
strmulanswer(4) = "Volatile as well as non volatile"
strmulanswer(5) = "20 Nano second"
strmulanswer(6) = "Binary Digits"
strmulanswer(7) = "Read/Write Memory"
strmulanswer(8) = "Several CPU's"
strmulanswer(9) = "Has inbuilt processing capability"
strmulanswer(10) = "Is a pointing device"
End Sub

Private Sub multiplechoice_Paper14()
strmultiple(1) = "The objective type of questions of the DOEACC 'O' level examination can be evaluated through"
strmuloption(1) = "Optical Mark Reader"
strmuloption(2) = "Optical Bar Code Reader"
strmuloption(3) = "Magnetic Character Reader"
strmuloption(4) = "None of these"

strmultiple(2) = "An impact printer uses"
strmuloption(5) = "Ink Jet technology"
strmuloption(6) = "Thermal technology"
strmuloption(7) = "Electro Mechanical mechanism"
strmuloption(8) = "None of these"

strmultiple(3) = "A Laser printer uses"
strmuloption(9) = "Raster scan"
strmuloption(10) = "Camera lens"
strmuloption(11) = "Heat sensitive paper"
strmuloption(12) = "None of these"

strmultiple(4) = "A plotter is"
strmuloption(13) = "An input device to produce good quality graphics"
strmuloption(14) = "An output device to produce drawings and graphics"
strmuloption(15) = "A fast output device using camera lenses"
strmuloption(16) = "None of these"

strmultiple(5) = "Magnetic tape is a "
strmuloption(17) = "A sequential access device"
strmuloption(18) = "A direct access device"
strmuloption(19) = "A sequential as well as a random access device"
strmuloption(20) = "None of these"

strmultiple(6) = "A Hard disk having n disk plates will have"
strmuloption(21) = "N tracks per cylinder "
strmuloption(22) = "N - 2 tracks per cylinder "
strmuloption(23) = "2n - 2 tracks per cylinder"
strmuloption(24) = "none of these"

strmultiple(7) = "In a hard disk, the track is"
strmuloption(25) = "The outermost"
strmuloption(26) = "The innermost"
strmuloption(27) = "Totally machine dependent"
strmuloption(28) = "None of these"

strmultiple(8) = "A soft sectored disk is associated with"
strmuloption(29) = "Hard disks"
strmuloption(30) = "Floppy disks"
strmuloption(31) = "Both hard disks and floppy disks"
strmuloption(32) = "None of these"

strmultiple(9) = "A CD-ROM is"
strmuloption(33) = "An optical ROM"
strmuloption(34) = "A magnetic ROM"
strmuloption(35) = "An erasable ROM"
strmuloption(36) = "None of these"

strmultiple(10) = "Data are stored serially in CD - ROM by using"
strmuloption(37) = "Concentric tracks "
strmuloption(38) = "Spiral tracks "
strmuloption(39) = "Concentric or spiral tracks"
strmuloption(40) = "None of these"

strmulanswer(1) = "Magnetic Character Reader"
strmulanswer(2) = "Thermal technology"
strmulanswer(3) = "Heat sensitive paper"
strmulanswer(4) = "An input device to produce good quality graphics"
strmulanswer(5) = "A sequential access device"
strmulanswer(6) = "2n - 2 tracks per cylinder"
strmulanswer(7) = "The outermost"
strmulanswer(8) = "Both hard disks and floppy disks"
strmulanswer(9) = "An optical ROM"
strmulanswer(10) = "Spiral tracks "
End Sub

Private Sub multiplechoice_Paper15()
strmultiple(1) = "Netware supports"
strmuloption(1) = "IBM PC compatible clients only."
strmuloption(2) = "APPLE Macintosh clients only"
strmuloption(3) = "Both a and b above"
strmuloption(4) = "None of these"

strmultiple(2) = "Virus survives by capturing"
strmuloption(5) = "Executable programs"
strmuloption(6) = "Interrupts"
strmuloption(7) = "COM programs"
strmuloption(8) = "None of these"

strmultiple(3) = "Which of the following is true for GUI?"
strmuloption(9) = "Making use of computer easier for non professionals"
strmuloption(10) = "They use icons and menus which users can select with mouse"
strmuloption(11) = "Both (a) and (b)"
strmuloption(12) = "None of these"

strmultiple(4) = "A transaction file is a type of"
strmuloption(13) = "Master file"
strmuloption(14) = "Data file"
strmuloption(15) = "Access method"
strmuloption(16) = "None of these"

strmultiple(5) = "Supercomputers are primarily useful for"
strmuloption(17) = "Input output intensive processing "
strmuloption(18) = "Mathematical intensive scientific applications"
strmuloption(19) = "Data retrieval operations"
strmuloption(20) = "None of these"

strmultiple(6) = "Data transfer rate in modems is measured in"
strmuloption(21) = "Bits per second "
strmuloption(22) = "Bits per minute"
strmuloption(23) = "HZ"
strmuloption(24) = "None of these"

strmultiple(7) = "E- mail"
strmuloption(25) = "Cannot address many users"
strmuloption(26) = "Does not provide protection given to first class mail"
strmuloption(27) = "Always uses bridge to send messages in different networks"
strmuloption(28) = "None of these"

strmultiple(8) = "The number used to identify a storage location is called"
strmuloption(29) = "A Bit   "
strmuloption(30) = "A Byte"
strmuloption(31) = "An address"
strmuloption(32) = "None of these"

strmultiple(9) = "A computer system is made up of"
strmuloption(33) = "Hardware only"
strmuloption(34) = "Software only"
strmuloption(35) = "Hardware and Software"
strmuloption(36) = "None of these"

strmultiple(10) = "A workstation has"
strmuloption(37) = "8 - 16 MB main memory"
strmuloption(38) = "16 - 32 MB main memory"
strmuloption(39) = "32 - 64 MB main memory"
strmuloption(40) = "None of these"


strmulanswer(1) = "Both a and b above"
strmulanswer(2) = "Executable programs"
strmulanswer(3) = "Both (a) and (b)"
strmulanswer(4) = "Access method"
strmulanswer(5) = "Mathematical intensive scientific applications"
strmulanswer(6) = "Bits per second "
strmulanswer(7) = "Always uses bridge to send messages in different networks"
strmulanswer(8) = "An address"
strmulanswer(9) = "Hardware and Software"
strmulanswer(10) = "32 - 64 MB main memory"
End Sub


