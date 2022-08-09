VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SADPapers"
Attribute VB_GlobalNameSpace = True
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
 
Function SAD(PaperType As String, ItPaper As Paper, mCat As Category, Number As Integer)
'ptype is the paper type
'itpaper is the number of paper as paper1,paper2 ...
'mcat is the category
'number is the number of questions, or options or answers to be taken

 StrPaperType = PaperType
' StrPaper = Paper
If PaperType = "Fill in the Blanks" Then
   If SADPaper = Paper1 Then Fill_in_the_Blanks_Paper1
   If SADPaper = Paper2 Then Fill_in_the_Blanks_Paper2
   If SADPaper = Paper3 Then Fill_in_the_Blanks_Paper3
   If SADPaper = Paper4 Then Fill_in_the_Blanks_Paper4
   If SADPaper = Paper5 Then Fill_in_the_Blanks_Paper5
   If SADPaper = Paper6 Then Fill_in_the_Blanks_Paper6
   If SADPaper = Paper7 Then Fill_in_the_Blanks_Paper7
   If SADPaper = Paper8 Then Fill_in_the_Blanks_Paper8
   If SADPaper = Paper9 Then Fill_in_the_Blanks_Paper9
   If SADPaper = Paper10 Then Fill_in_the_Blanks_Paper10
   If SADPaper = Paper11 Then Fill_in_the_Blanks_Paper11
   If SADPaper = Paper12 Then Fill_in_the_Blanks_Paper12
   If SADPaper = Paper13 Then Fill_in_the_Blanks_Paper13
   If SADPaper = Paper14 Then Fill_in_the_Blanks_Paper14
   If SADPaper = Paper15 Then Fill_in_the_Blanks_Paper15
   
ElseIf PaperType = "Matching Columns" Then
       If SADPaper = Paper1 Then Matching_Columns_Paper1
       If SADPaper = Paper2 Then Matching_Columns_Paper2
       If SADPaper = Paper3 Then Matching_Columns_Paper3
       If SADPaper = Paper4 Then Matching_Columns_Paper4
       If SADPaper = Paper5 Then Matching_Columns_Paper5
       If SADPaper = Paper6 Then Matching_Columns_Paper6
       If SADPaper = Paper7 Then Matching_Columns_Paper7
       If SADPaper = Paper8 Then Matching_Columns_Paper8
       If SADPaper = Paper9 Then Matching_Columns_Paper9
       If SADPaper = Paper10 Then Matching_Columns_Paper10
       If SADPaper = Paper11 Then Matching_Columns_Paper11
       If SADPaper = Paper12 Then Matching_Columns_Paper12
       If SADPaper = Paper13 Then Matching_Columns_Paper13
       If SADPaper = Paper14 Then Matching_Columns_Paper14
       If SADPaper = Paper15 Then Matching_Columns_Paper15
       
ElseIf PaperType = "True Or False" Then
       If SADPaper = Paper1 Then TrueorFalse_Paper1
       If SADPaper = Paper2 Then TrueorFalse_Paper2
       If SADPaper = Paper3 Then TrueorFalse_Paper3
       If SADPaper = Paper4 Then TrueorFalse_Paper4
       If SADPaper = Paper5 Then TrueorFalse_Paper5
       If SADPaper = Paper6 Then TrueorFalse_Paper6
       If SADPaper = Paper7 Then TrueorFalse_Paper7
       If SADPaper = Paper8 Then TrueorFalse_Paper8
       If SADPaper = Paper9 Then TrueorFalse_Paper9
       If SADPaper = Paper10 Then TrueorFalse_Paper10
       If SADPaper = Paper11 Then TrueorFalse_Paper11
       If SADPaper = Paper12 Then TrueorFalse_Paper12
       If SADPaper = Paper13 Then TrueorFalse_Paper13
       If SADPaper = Paper14 Then TrueorFalse_Paper14
       If SADPaper = Paper15 Then TrueorFalse_Paper15
       
ElseIf PaperType = "Descriptive Questions" Then
       If SADPaper = Paper1 Then questions_paper1
       If SADPaper = Paper2 Then questions_paper2
       If SADPaper = Paper3 Then questions_paper3
       If SADPaper = Paper4 Then questions_paper4
End If

If mCat = Fill_Questions Then
   SAD = StrQuestions(Number)
ElseIf mCat = Fill_Options Then
       SAD = stroptions(Number)
ElseIf mCat = Fill_Answers Then
       SAD = stranswers(Number)
       
ElseIf mCat = Matching_ColumnA Then
       SAD = strcolumnA(Number)
ElseIf mCat = Matching_ColumnB Then
       SAD = strcolumnB(Number)
ElseIf mCat = Matching_Answers Then
       SAD = stranswers(Number)
       
ElseIf mCat = Questions Then
       SAD = StrQuestions(Number)
       
ElseIf mCat = TF_Questions Then
       SAD = StrQuestions(Number)
ElseIf mCat = TF_Answers Then
       SAD = stranswers(Number)
End If

End Function

Private Sub Fill_in_the_Blanks_Paper1()
StrQuestions(1) = "The three basic subsystems (business functions) of any industrial organization are marketing, manufacturing and dash."
StrQuestions(2) = "If the return on proposed system alternative falls below the acceptable level for capital projects, the dash factors are critical to the final decision."
StrQuestions(3) = "dash will ensure accuracy of inputs to a computer."
StrQuestions(4) = "dash testing must precede unit testing and system testing."
StrQuestions(5) = "Data dash describes every data element comprehensively."
StrQuestions(6) = "dash is the use of model in an attempt to identify and/or reflect the behavior of a real person, process of system."
StrQuestions(7) = "In dash processing, the data are gathered for a period and collected into a group before they are entered into a computer and processed."
StrQuestions(8) = "A dash is the primary tool used in structured system development to graphically depict systems."
StrQuestions(9) = "dash charts are an effective means for relating input data to files and output reports."
StrQuestions(10) = "dash is one of the methods used to capture the data for analysis of any organization's functioning."

stroptions(1) = "Module"
stroptions(2) = "Flow Chart"
stroptions(3) = "Interview"
stroptions(4) = "Simulation"
stroptions(5) = "System flow"
stroptions(6) = "Batch"
stroptions(7) = "Scanner"
stroptions(8) = "Dictionary"
stroptions(9) = "Finance"
stroptions(10) = "Intangible"

stranswers(1) = "Finance"
stranswers(2) = "Intangible"
stranswers(3) = "Scanner"
stranswers(4) = "Module"
stranswers(5) = "Dictionary"
stranswers(6) = "Simulation"
stranswers(7) = "Batch"
stranswers(8) = "Flow Chart"
stranswers(9) = "System flow"
stranswers(10) = "Interview"
End Sub

Private Sub Fill_in_the_Blanks_Paper2()
StrQuestions(1) = "The trend continues towards more dash for management information system."
StrQuestions(2) = "dash is the method of processing current data by both the old and the new system."
StrQuestions(3) = "dash is a computer science branch that is involved with using computers to solve problems that appear to require human imagination or intelligence."
StrQuestions(4) = "dash is the hardware/software boundary that permits communication between people and computers."
StrQuestions(5) = "dash charts are an effective means for relating input data to files and output reports."
StrQuestions(6) = "Communications packages allow people to send electronic dash to specified users along a network of computers and workstations."
StrQuestions(7) = "A dash questionnaire is a tool determining managerial weaknesses when evaluating internal administrative control."
StrQuestions(8) = "Gantt chart is a dash tool for managing projects."
StrQuestions(9) = "A principle of good input system design dictates that a common dash be established in order to keep manual tasks to a minimum"
StrQuestions(10) = "The longest sequence of events in a project is called the dash path."

stroptions(1) = "management audit"
stroptions(2) = "Mail"
stroptions(3) = "Data representation"
stroptions(4) = "Critical"
stroptions(5) = "Planning"
stroptions(6) = "Parallel run"
stroptions(7) = "Grid"
stroptions(8) = "Interface"
stroptions(9) = "AI"
stroptions(10) = "Distributed processing"

stranswers(1) = "Distributed processing"
stranswers(2) = "Parallel run"
stranswers(3) = "Planning"
stranswers(4) = "Interface"
stranswers(5) = "Grid"
stranswers(6) = "Mail"
stranswers(7) = "management audit"
stranswers(8) = "Planning"
stranswers(9) = "Data representation"
stranswers(10) = "Critical"

End Sub
Private Sub Fill_in_the_Blanks_Paper3()
StrQuestions(1) = "In converting from an old system to a new one,. four methods to cut over exist; Parallel Systems, Direct cut over Phase in and dash."
StrQuestions(2) = "Menu driven mode is a dash strategy to interface with the system in an interactive mode."
StrQuestions(3) = "Open systems interact with the dash."
StrQuestions(4) = "The total amount of processing that a computer can complete over a fixed period of time is called dash"
StrQuestions(5) = "The inclusion of the dash factor in output design can increase the chances of a successful system design."
StrQuestions(6) = "Eliminating or keeping dash of information to a minimum is an important consideration when designing new reports."
StrQuestions(7) = "dash bit coding scheme is used to representing a byte in data transmission."
StrQuestions(8) = "Pseudocode is a tool to define dash."
StrQuestions(9) = "A LAN is nucleated around a dash."
StrQuestions(10) = "dash refers to the capability of the system to changing environmental factors."

stroptions(1) = "ASCII"
stroptions(2) = "File Server"
stroptions(3) = "Flexibility"
stroptions(4) = "Human"
stroptions(5) = "Duplication"
stroptions(6) = "Program logic"
stroptions(7) = "Environment"
stroptions(8) = "Pilot Approach"
stroptions(9) = "Throughput"
stroptions(10) = "Dialogue"

stranswers(1) = "Pilot Approach"
stranswers(2) = "Dialogue"
stranswers(3) = "Environment"
stranswers(4) = "Throughput"
stranswers(5) = "Human"
stranswers(6) = "Duplication"
stranswers(7) = "ASCII"
stranswers(8) = "Program logic"
stranswers(9) = "File Server"
stranswers(10) = "Flexibility"
End Sub
Private Sub Fill_in_the_Blanks_Paper4()
StrQuestions(1) = "Virtual reality is a computer generated interactive dash dimensional environment, in which a person is immersed."
StrQuestions(2) = "Data Base management packages may set up an dash to control the sequence in which records are presented in a report."
StrQuestions(3) = "The modular concept is applicable to systems design as well as dash"
StrQuestions(4) = "Key to address conversion is frequently found in dash organization method."
StrQuestions(5) = "Brainstorming and dash may be used to elicit information regarding he user's requirements."
StrQuestions(6) = "The grouping together of two or more logical records for working on magnetic tape is called dash."
StrQuestions(7) = "A portion of a program that performs a specific program function is called a dash."
StrQuestions(8) = "A Gantt chart is fundamentally a dash chart."
StrQuestions(9) = "Chaining and dash are frequently used techniques in respect of storage of records in files."
StrQuestions(10) = "The next major step before system design and after feasibility study is dash"

stroptions(1) = "Sub routine"
stroptions(2) = "Linking"
stroptions(3) = "Analysis activity"
stroptions(4) = "Bar"
stroptions(5) = "Direct file"
stroptions(6) = "Three"
stroptions(7) = "Blocking"
stroptions(8) = "Computer programming"
stroptions(9) = "Interviews"
stroptions(10) = "Index"

stranswers(1) = "Three"
stranswers(2) = "Index"
stranswers(3) = "Computer programming"
stranswers(4) = "Direct file"
stranswers(5) = "Interviews"
stranswers(6) = "Blocking"
stranswers(7) = "Sub routine"
stranswers(8) = "Bar"
stranswers(9) = "Linking"
stranswers(10) = "Analysis activity"
End Sub
Private Sub Fill_in_the_Blanks_Paper5()
StrQuestions(1) = "The components that make up any system are known as dash."
StrQuestions(2) = "An external dash document is an external output document all or part of which enters as input to the system."
StrQuestions(3) = "Tangible benefits by their very nature, require dash evaluation."
StrQuestions(4) = "A dash interface is a convenient way of user interaction when it is difficult for the user to remember all the options available and for typing in complicated commands."
StrQuestions(5) = "If a Data Dictionary is not included in the Systems Analysis and Design of software protect, then dash cannot be carried out well."
StrQuestions(6) = "One major advantage of systems developed using DBMS, as compared to systems using conventional file system is in responding to dash."
StrQuestions(7) = "A Structured Decision is not dash."
StrQuestions(8) = "In considering the total cost associated with desired information, cost of data collection, data input and computer processing costs are dash"
StrQuestions(9) = "dash allocation is a tactical decision."
StrQuestions(10) = "When there are not enough transactions with the main computer system, remote job entry is done best under dash mode."

stroptions(1) = "Recurring"
stroptions(2) = "Ad Hoc enquiries"
stroptions(3) = "off line"
stroptions(4) = "Budget"
stroptions(5) = "Subjective"
stroptions(6) = ""
stroptions(7) = "Data consistency checks"
stroptions(8) = "System boundary"
stroptions(9) = "Menu"
stroptions(10) = "Menu"
stroptions(11) = "on line"
stroptions(12) = "Real time"

stranswers(1) = "System boundary"
stranswers(2) = "Menu"
stranswers(3) = "Subjective"
stranswers(4) = "Menu"
stranswers(5) = "Data consistency checks"
stranswers(6) = "Ad Hoc enquiries"
stranswers(7) = ""
stranswers(8) = "Recurring"
stranswers(9) = "Budget"
stranswers(10) = "off line"
End Sub
Private Sub Fill_in_the_Blanks_Paper6()
StrQuestions(1) = "The members of the Executive Steering Committee members are , generally dash managers."
StrQuestions(2) = "Ergonomics is concerned with the dash of the operator."
StrQuestions(3) = "The fields within a logical record on which the files sorted are called dash"
StrQuestions(4) = "dash is set a tool of SASD."
StrQuestions(5) = "dash is a part of the SDLC."
StrQuestions(6) = "If the dash analysis phase of a software development project is not done properly then the users real needs are not met."
StrQuestions(7) = "A major component of modularization is dash"
StrQuestions(8) = "'Scheduling of delivery' is an example of dash system."
StrQuestions(9) = "The type of code, which involves the assignment of a consecutive and unique number to each item, is the dash code."
StrQuestions(10) = "Reviews and walkthroughs are carried out at dash stages of the SDLC."

stroptions(1) = "Bench marking"
stroptions(2) = "Decision support"
stroptions(3) = "Almost all"
stroptions(4) = "Numeric"
stroptions(5) = "Cohesion"
stroptions(6) = "Sort keys"
stroptions(7) = "Requirement"
stroptions(8) = "Top level"
stroptions(9) = "HIPO"
stroptions(10) = "Comfort"

stranswers(1) = "Top level"
stranswers(2) = "Comfort"
stranswers(3) = "Sort keys"
stranswers(4) = "HIPO"
stranswers(5) = "Bench marking"
stranswers(6) = "Requirement"
stranswers(7) = "Cohesion"
stranswers(8) = "Decision support"
stranswers(9) = "Numeric"
stranswers(10) = "Almost all"
End Sub
Private Sub Fill_in_the_Blanks_Paper7()
StrQuestions(1) = "MODEM is used in dash systems."
StrQuestions(2) = "dash is a set of programs that manage the resources of a computer system whereas application programs deliver management requirements of information."
StrQuestions(3) = "As the management level goes up the hierarchy, information becomes more and more dash."
StrQuestions(4) = "Audit, reliability, Security are integral to dash."
StrQuestions(5) = "Managers performing strategic function need dash information."
StrQuestions(6) = "The detailed investigation of the present system is referred to as dash"
StrQuestions(7) = "Whether a proposed system can provide right information for the organization’s personnel falls under the study of dash feasibility."
StrQuestions(8) = "Conventional file system applications are dash"
StrQuestions(9) = "Reason for normalization is to simplify dash"
StrQuestions(10) = "System quality requires that unless user needs and software requirements specifications are reviewed, dash should not be initiated."

stroptions(1) = "External"
stroptions(2) = "Operational"
stroptions(3) = "System design"
stroptions(4) = "Inflexible"
stroptions(5) = "Updates"
stroptions(6) = "SDLC"
stroptions(7) = "WAN"
stroptions(8) = "Summarized"
stroptions(9) = "System analysis"
stroptions(10) = "System software"

stranswers(1) = "WAN"
stranswers(2) = "System software"
stranswers(3) = "Summarized"
stranswers(4) = "SDLC"
stranswers(5) = "External"
stranswers(6) = "System analysis"
stranswers(7) = "Operational"
stranswers(8) = "Inflexible"
stranswers(9) = "Updates"
stranswers(10) = "System design"
End Sub
Private Sub Fill_in_the_Blanks_Paper8()
StrQuestions(1) = "An indexed file essentially is a dash file."
StrQuestions(2) = "dash are difficult to answer in conventional file system applications."
StrQuestions(3) = "'Removal of functional dependency from a relation' falls under the category of dash of Normalization."
StrQuestions(4) = "A Data Flow Diagram is referred to during the design of input because there is dash relation between the input items and processes"
StrQuestions(5) = "Reason for normalization is to dash data."
StrQuestions(6) = "System quality requires that inspection should be carried out at pre specified dash"
StrQuestions(7) = "In the implementation phase of SASD dash is included"
StrQuestions(8) = "dash is the creating, developing and refining a working model of the final Operational system."
StrQuestions(9) = "dash cannot be used for data entry."
StrQuestions(10) = "Stub is net with in dash of modules"

stroptions(1) = "Sizing"
stroptions(2) = "Structure"
stroptions(3) = "Testing"
stroptions(4) = "COM"
stroptions(5) = "Prototyping"
stroptions(6) = "A one to one"
stroptions(7) = "Ad-hoc queries"
stroptions(8) = "Milestones"
stroptions(9) = "2nd form"
stroptions(10) = "Sequential"

stranswers(1) = "Sequential"
stranswers(2) = "Ad-hoc queries"
stranswers(3) = "2nd form"
stranswers(4) = "A one to one"
stranswers(5) = "Structure"
stranswers(6) = "Milestones"
stranswers(7) = "Sizing"
stranswers(8) = "Prototyping"
stranswers(9) = "COM"
stranswers(10) = "Testing"
End Sub

Private Sub Fill_in_the_Blanks_Paper9()
StrQuestions(1) = "Reason for normalization is to simplify dash"
StrQuestions(2) = "A blocking factor of 1/3 is dash 1/4."
StrQuestions(3) = "dash ensures simultaneous updates of all data structures to keep the status current"
StrQuestions(4) = "System quality  requires that a proper dash should be prepared and followed"
StrQuestions(5) = "dash messages may be stored and forwarded by voicemail system"
StrQuestions(6) = "Cost of error correction is dash at requirement analysis stage"
StrQuestions(7) = "dash is used in end user training before implementation."
StrQuestions(8) = "In the implementation phase of SASD, dash is included."
StrQuestions(9) = "The main advantage of normalized relations in relational DBMS is that they do not suffer from anomalies during dash operations."
StrQuestions(10) = "Digital signature is used in dash."

stroptions(1) = "worse than"
stroptions(2) = "spoken"
stroptions(3) = "CAI"
stroptions(4) = "data communication"
stroptions(5) = "Parallel run"
stroptions(6) = "Delete and update"
stroptions(7) = "Test plan"
stroptions(8) = "Query retrieval"
stroptions(9) = "Least"
stroptions(10) = "DBMS"
stroptions(10) = "Better than"
stroptions(10) = "Most"

stranswers(1) = "Query retrieval"
stranswers(2) = "worse than"
stranswers(3) = "DBMS"
stranswers(4) = "Test plan"
stranswers(5) = "spoken"
stranswers(6) = "Least"
stranswers(7) = "CAI"
stranswers(8) = "Parallel run"
stranswers(9) = "Delete and update"
stranswers(10) = "data communication"
End Sub
Private Sub Fill_in_the_Blanks_Paper10()
StrQuestions(1) = "To produce training manual for your organization you would use dash"
StrQuestions(2) = "In the implementation phase of SASD dash is included"
StrQuestions(3) = "A dash is made after completion of preplanned series of steps"
StrQuestions(4) = "Top level managers use dash"
StrQuestions(5) = "dash process identifies redundant data that may exist in the logical data structure."
StrQuestions(6) = "The generic solution to a problem is termed dash."
StrQuestions(7) = "dash aids in Quality assurance"
StrQuestions(8) = "dash facilitates conditions to be related to actions"
StrQuestions(9) = "dash detects transposition errors"
StrQuestions(10) = "In a DFD dash refers to source or destination of data."

stroptions(1) = "Data Store"
stroptions(2) = "Normalizations"
stroptions(3) = "Check Digit"
stroptions(4) = "Peer Review"
stroptions(5) = "Decision table"
stroptions(6) = "Structured decision"
stroptions(7) = "DTP"
stroptions(8) = "An algorithm"
stroptions(9) = "Specification freezes"
stroptions(10) = "Strategic information"

stranswers(1) = "DTP"
stranswers(2) = "Specification freezes"
stranswers(3) = "Structured decision"
stranswers(4) = "Strategic information"
stranswers(5) = "Normalizations"
stranswers(6) = "An algorithm"
stranswers(7) = "Peer Review"
stranswers(8) = "Decision table"
stranswers(9) = "Check Digit"
stranswers(10) = "Data Store"
End Sub
Private Sub Fill_in_the_Blanks_Paper11()
StrQuestions(1) = "dash testing must precede unit testing and system testing."
StrQuestions(2) = "Data dictionary describes every data element and dash"
StrQuestions(3) = "Brainstorming and dash may be used to elicit information regarding the user's requirements."
StrQuestions(4) = "A dash is the primary tool used in structured system development to graphically depict systems."
StrQuestions(5) = "dash is the method of processing current data by both the old and the new system."
StrQuestions(6) = "Programmable dash are incorporated into programs in a management information system."
StrQuestions(7) = "Eliminating or keeping dash of information to a minimum is an important consideration when designing new reports."
StrQuestions(8) = "dash charts are an effective means for relating input data to file and output reports."
StrQuestions(9) = "A management audit questionnaire is  a tool determining managerial dash when evaluating internal administrative control."
StrQuestions(10) = "Questionnaire and dash methods are used to capture the data for analysis of an Organization’s functioning."

stroptions(1) = "Control "
stroptions(2) = "Feedback"
stroptions(3) = "Delphi method"
stroptions(4) = "Data dictionary"
stroptions(5) = "Module"
stroptions(6) = "Boundary"
stroptions(7) = "Parallel running"
stroptions(8) = "Black box"
stroptions(9) = "Data structure"
stroptions(10) = "Data Flow Diagram"
stroptions(11) = "System flow"
stroptions(12) = "Decisions"
stroptions(13) = "Coding"
stroptions(14) = "Weaknesses"
stroptions(15) = "Interview"
stroptions(16) = "Process"
stroptions(17) = "Duplication"
stroptions(18) = "Pie"

stranswers(1) = "Module"
stranswers(2) = "Data structure"
stranswers(3) = "Interview"
stranswers(4) = "Data Flow Diagram"
stranswers(5) = "Parallel running"
stranswers(6) = "Coding"
stranswers(7) = "Pie"
stranswers(8) = "System flow"
stranswers(9) = "Weaknesses"
stranswers(10) = "Delphi method"
End Sub
Private Sub Fill_in_the_Blanks_Paper12()
StrQuestions(1) = "A DBMSD uses dash command to access data."
StrQuestions(2) = "dash is information on how well a system is performing."
StrQuestions(3) = "dash provides a list of each data element used in the system."
StrQuestions(4) = "dash operation must be performed on a transaction file before it can be used in a sequential updating run."
StrQuestions(5) = "dash is an integrated collection of data, with minimum redundancy, so that different applications can use the required data."
StrQuestions(6) = "dash maintains databases, establishes security controls and performs backup p and recovery."
StrQuestions(7) = "A system in which transaction is processed as and when they are entered, is dash application."
StrQuestions(8) = "Where one has to process all the records in a file, the best file organization is dash."
StrQuestions(9) = "dash is essential to endure that the developer and the customer have the same perception of the system."
StrQuestions(10) = "dash is used to connect processes to each other and the arrowhead indicates direction of data transfer."

stroptions(1) = "Sorting"
stroptions(2) = "A review"
stroptions(3) = "UNIX"
stroptions(4) = "ERD"
stroptions(5) = "Sequential"
stroptions(6) = "An on-line"
stroptions(7) = "Feedback"
stroptions(8) = "DML"
stroptions(9) = "Data dictionary"
stroptions(10) = "DBMS"

stranswers(1) = "DML"
stranswers(2) = "Feedback"
stranswers(3) = "Data dictionary"
stranswers(4) = "Sorting"
stranswers(5) = "DBMS"
stranswers(6) = "UNIX"
stranswers(7) = "An on-line"
stranswers(8) = "Sequential"
stranswers(9) = "A review"
stranswers(10) = "ERD"
End Sub

Private Sub Fill_in_the_Blanks_Paper13()
StrQuestions(1) = "If a code i the input file had to be validated b checking with the master file, the best file organization for the master file is dash"
StrQuestions(2) = "System design is made up of data design and dash design."
StrQuestions(3) = "dash is a unique string that must be supplied by the user of a computer before being granted access."
StrQuestions(4) = "dash is a method of interlinking different departments using computers located in the same building."
StrQuestions(5) = "A flow chart is dash representation of the processing logic."
StrQuestions(6) = "The dash grouping of connected data elements is called a file."
StrQuestions(7) = "To calculate dash database size, volumes of all transactions must be known."
StrQuestions(8) = "dash operation of the transaction file is necessary to update a sequential master file."
StrQuestions(9) = "System boundary separates the system from its dash"
StrQuestions(10) = "dash is a hierarchical diagram showing the relationship between various program modules."

stroptions(1) = "Structure chart"
stroptions(2) = "Logical"
stroptions(3) = "Environment"
stroptions(4) = "Password"
stroptions(5) = "Fetch"
stroptions(6) = "Multi-user"
stroptions(7) = "Process"
stroptions(8) = "LAN"
stroptions(9) = "Indexed sequential"
stroptions(10) = "A graphical"

stranswers(1) = "Indexed sequential"
stranswers(2) = "Process"
stranswers(3) = "Password"
stranswers(4) = "LAN"
stranswers(5) = "A graphical"
stranswers(6) = "Logical"
stranswers(7) = "Multi-user"
stranswers(8) = "Fetch"
stranswers(9) = "Environment"
stranswers(10) = "Structure chart"
End Sub

Private Sub Fill_in_the_Blanks_Paper14()
StrQuestions(1) = "User training is implicit in the dash stage of SDLC."
StrQuestions(2) = "Changing hardware and software procedures, without having to write application programs, is dash "
StrQuestions(3) = "The concept, that data is a common resource, is used in dash  systems."
StrQuestions(4) = "A dash  is a bounded, named and continuous set of program statements."
StrQuestions(5) = "The fundamental Data legal, recorded in a data dictionary is the Tracing, of any input record or process, performed on a system back to its original source, is dash "
StrQuestions(6) = "Tracing, of any input record or process, performed on a system back to its original source, is dash "
StrQuestions(7) = "A Tracing, of any input record or process, performed on a system back to its original source, is dash  enables one to organize and document, a system that is to be implemented, without any programming structured chart."
StrQuestions(8) = "When before programming a systems analyst presents the proposed system to a designated team , answers questions and takes note of omissions and additions, the processes is called Tracing, of any input record or process, performed on a system back to its original source, is dash "
StrQuestions(9) = "Sub systems cannot function in dash "
StrQuestions(10) = "The dash  phase of SDLC will face problems like non-cooperation resistance to relocation and changes n the organization tree"

stroptions(1) = "Conversion"
stroptions(2) = "Structured walkthrough"
stroptions(3) = "An audit trail"
stroptions(4) = "Database"
stroptions(5) = "Implementation"
stroptions(6) = "Isolation"
stroptions(7) = "Data independence"
stroptions(8) = "Structured chart"
stroptions(9) = "Data element"
stroptions(10) = "Module sub routing"

stranswers(1) = "Implementation"
stranswers(2) = "Data independence"
stranswers(3) = "Database"
stranswers(4) = "Module sub routing"
stranswers(5) = "Data element"
stranswers(6) = "An audit trail"
stranswers(7) = "Structured chart"
stranswers(8) = "Structured walkthrough"
stranswers(9) = "Isolation"
stranswers(10) = "Conversion"
End Sub

Private Sub Fill_in_the_Blanks_Paper15()
StrQuestions(1) = "Two examples of reports produced for management planning are balance sheet and dash "
StrQuestions(2) = "Programs with long executing times should have dash facility in built in them"
StrQuestions(3) = "Manuals, forms, and other descriptive information that portray the use and/or operation of a system, are called dash  of the system"
StrQuestions(4) = "dash  processing shares resources between computer users"
StrQuestions(5) = "Under the dash  concept the system is defined in terms of inputs and outputs, rather than in terms of how the system effects a transformation."
StrQuestions(6) = "The most commonly used dialogue strategy for an on line conversation with a computer is dash "
StrQuestions(7) = "dash  is the proc3ss of using software in a live environment in order to fine errors."
StrQuestions(8) = "dash  is the key to structured analysis and design"
StrQuestions(9) = "dash  is a primary tool used in the Structured System Development, to graphically depict the system."
StrQuestions(10) = "Structured English is often called dash  because it mimics programming code."

stroptions(1) = "Validation"
stroptions(2) = "The DFD"
stroptions(3) = "P&L statement"
stroptions(4) = "Pseudo code"
stroptions(5) = "documentation"
stroptions(6) = "Hierarchical"
stroptions(7) = "Re start"
stroptions(8) = "Menu driven"
stroptions(9) = "Black box"
stroptions(10) = "Distributed"

stranswers(1) = "P&L statement"
stranswers(2) = "Re start"
stranswers(3) = "documentation"
stranswers(4) = "Distributed"
stranswers(5) = "Black box"
stranswers(6) = "Menu driven"
stranswers(7) = "Validation"
stranswers(8) = "Hierarchical"
stranswers(9) = "The DFD"
stranswers(10) = "Pseudo code"
End Sub

Private Sub Fill_in_the_Blanks_Paper16()
StrQuestions(1) = "dash are essential to assure that the developer and the customer have the same perception of the system."
StrQuestions(2) = "Defining the problem and determining the new system objectives are part of the dash  phase of SDLC."
StrQuestions(3) = "dash  produces a step-by-step model of the proposed system"
StrQuestions(4) = "ISO 9000 concerns dash  set up in an organization."
StrQuestions(5) = "An advantage of parallel conversion at system implementation is dash "
StrQuestions(6) = "Computer based dash  can be through or round the computer"
StrQuestions(7) = "An Information System that gathers, stores and analyses data on the products of a competitor,  is known as dash system."
StrQuestions(8) = "In the testing phase of the SDLC, the first activity is dash "
StrQuestions(9) = "In most computer systems access control is exercised by means of some form of dash "
StrQuestions(10) = "Changes made periodically to a system after its implementation is called system dash "

stroptions(1) = "Maintenance"
stroptions(2) = "System specs"
stroptions(3) = "Authentication"
stroptions(4) = "System analysis"
stroptions(5) = "Module testing"
stroptions(6) = "Structured System Analysis"
stroptions(7) = "A marketing intelligence"
stroptions(8) = "Quality Assurance systems"
stroptions(9) = "Audit"
stroptions(10) = "Safety"

stranswers(1) = "System specs"
stranswers(2) = "System analysis"
stranswers(3) = "Structured System Analysis"
stranswers(4) = "Quality Assurance systems"
stranswers(5) = "Safety"
stranswers(6) = "Audit"
stranswers(7) = "A marketing intelligence"
stranswers(8) = "Module testing"
stranswers(9) = "Authentication"
stranswers(10) = "Maintenance"
End Sub
Private Sub Matching_Columns_Paper1()
strcolumnA(1) = "structured walkthrough"
strcolumnA(2) = "work sampling"
strcolumnA(3) = "combination check"
strcolumnA(4) = "return on investment"
strcolumnA(5) = "condition stub"
strcolumnA(6) = "prototyping"
strcolumnA(7) = "GUI"
strcolumnA(8) = "Critical path"
strcolumnA(9) = "Flow chart"
strcolumnA(10) = "Tuple"

strcolumnB(1) = "feasibility study"
strcolumnB(2) = "input control"
strcolumnB(3) = "freezing user requirements"
strcolumnB(4) = "facilitating use of application software"
strcolumnB(5) = "relation"
strcolumnB(6) = "program logic"
strcolumnB(7) = "fact finding"
strcolumnB(8) = "peer review"
strcolumnB(9) = "project schedule"
strcolumnB(10) = "decision table"

stranswers(1) = "peer review"
stranswers(2) = "fact finding"
stranswers(3) = "input control"
stranswers(4) = "feasibility study"
stranswers(5) = "decision table"
stranswers(6) = "freezing user requirements"
stranswers(7) = "facilitating use of application software"
stranswers(8) = "project schedule"
stranswers(9) = "program logic"
stranswers(10) = "relation"
End Sub
Private Sub Matching_Columns_Paper2()
strcolumnA(1) = "password"
strcolumnA(2) = "payback period"
strcolumnA(3) = "DDL"
strcolumnA(4) = "display layout chart"
strcolumnA(5) = "system log"
strcolumnA(6) = "state transition diagram"
strcolumnA(7) = "structure chart"
strcolumnA(8) = "magnetic ink character"
strcolumnA(9) = "relative file"
strcolumnA(10) = "primary key"

strcolumnB(1) = "DBMS"
strcolumnB(2) = "real time system design"
strcolumnB(3) = "key to address transformation"
strcolumnB(4) = "a file of all updates"
strcolumnB(5) = "output design"
strcolumnB(6) = "access control"
strcolumnB(7) = "profitability analysis"
strcolumnB(8) = "modularization"
strcolumnB(9) = "cheque processing"
strcolumnB(10) = "file indexing"

stranswers(1) = "access control"
stranswers(2) = "profitability analysis"
stranswers(3) = "DBMS"
stranswers(4) = "output design"
stranswers(5) = "a file of all updates"
stranswers(6) = "real time system design"
stranswers(7) = "modularization"
stranswers(8) = "cheque processing"
stranswers(9) = "key to address transformation"
stranswers(10) = "file indexing"

End Sub

Private Sub Matching_Columns_Paper3()
strcolumnA(1) = "the historical  nature of MIS reports and need to deal with unstructured problem situation prompted the introduction of"
strcolumnA(2) = "daily information needed to operate the business is known as"
strcolumnA(3) = "the measure in determining the number of transactions that the computer must accept, process and return within a certain time period is known as"
strcolumnA(4) = "system which isolates it from environment is known as "
strcolumnA(5) = "information that relates to long range planning policies is known as"
strcolumnA(6) = "in SDLC, the stage which refers to the technical specifications for input, output, file and processing that will be applied in implementing the candidate system is known as"
strcolumnA(7) = "the process of converting a new system design into operating is known as the stage of formation"
strcolumnA(8) = "a structured repository of data about data is known as"
strcolumnA(9) = "the most widely used form of verification methods is some type of"
strcolumnA(10) = "DFD that modeled a whole system by one process is called    "

strcolumnB(1) = "documentation"
strcolumnB(2) = "data dictionary"
strcolumnB(3) = "implementation"
strcolumnB(4) = "evaluation"
strcolumnB(5) = "closed system"
strcolumnB(6) = "prototyping"
strcolumnB(7) = "operational information"
strcolumnB(8) = "strategic information"
strcolumnB(9) = "system life cycle"
strcolumnB(10) = "feedback"

stranswers(1) = "system life cycle"
stranswers(2) = "operational information"
stranswers(3) = "evaluation"
stranswers(4) = "closed system"
stranswers(5) = "strategic information"
stranswers(6) = "documentation"
stranswers(7) = "implementation"
stranswers(8) = "data dictionary"
stranswers(9) = "feedback"
stranswers(10) = "prototyping"
End Sub

Private Sub Matching_Columns_Paper4()
strcolumnA(1) = "smart card"
strcolumnA(2) = "knowledge base"
strcolumnA(3) = "alpha test"
strcolumnA(4) = "validation"
strcolumnA(5) = "CASE"
strcolumnA(6) = "action statement"
strcolumnA(7) = "sizing"
strcolumnA(8) = "tuples"
strcolumnA(9) = "bench mark"
strcolumnA(10) = "topology"

strcolumnB(1) = "record occurrences"
strcolumnB(2) = "hardware selection"
strcolumnB(3) = "table"
strcolumnB(4) = "storage assessment"
strcolumnB(5) = "SASD tool"
strcolumnB(6) = "expert system"
strcolumnB(7) = "updating prelude"
strcolumnB(8) = "communications"
strcolumnB(9) = "programming"
strcolumnB(10) = "transaction processing"

stranswers(1) = "transaction processing"
stranswers(2) = "expert system"
stranswers(3) = "programming"
stranswers(4) = "updating prelude"
stranswers(5) = "SASD tool"
stranswers(6) = "table"
stranswers(7) = "storage assessment"
stranswers(8) = "record occurrences"
stranswers(9) = "hardware selection"
stranswers(10) = "communications"
End Sub

Private Sub Matching_Columns_Paper5()
strcolumnA(1) = "feedback"
strcolumnA(2) = "strategic planning"
strcolumnA(3) = "feasibility study"
strcolumnA(4) = "through put"
strcolumnA(5) = "user training"
strcolumnA(6) = "graphical seeking"
strcolumnA(7) = "batch processing"
strcolumnA(8) = "edit report"
strcolumnA(9) = "what - if analysis"
strcolumnA(10) = "computer crime"

strcolumnB(1) = "cost benefit analysis"
strcolumnB(2) = "system performance"
strcolumnB(3) = "control"
strcolumnB(4) = "unstructured decisions"
strcolumnB(5) = "inventory report generation"
strcolumnB(6) = "executive information system"
strcolumnB(7) = "system implementation"
strcolumnB(8) = "hacking"
strcolumnB(9) = "errors during transactions"
strcolumnB(10) = "screen design"

stranswers(1) = "control"
stranswers(2) = "executive information system"
stranswers(3) = "cost benefit analysis"
stranswers(4) = "system performance"
stranswers(5) = "system implementation"
stranswers(6) = "screen design"
stranswers(7) = "inventory report generation"
stranswers(8) = "errors during transactions"
stranswers(9) = "unstructured decisions"
stranswers(10) = "hacking"
End Sub

Private Sub Matching_Columns_Paper6()
strcolumnA(1) = "television"
strcolumnA(2) = "monkey"
strcolumnA(3) = "a linear program"
strcolumnA(4) = "an organization chart"
strcolumnA(5) = "strategy"
strcolumnA(6) = "a refrigerator with a thermostat"
strcolumnA(7) = "a budgeting system with no variance analysis"
strcolumnA(8) = "EOQ"
strcolumnA(9) = "ERD"
strcolumnA(10) = "CSF"

strcolumnB(1) = "normative model"
strcolumnB(2) = "data model"
strcolumnB(3) = "a key area for organizational success"
strcolumnB(4) = "inventory control"
strcolumnB(5) = "closed loop system"
strcolumnB(6) = "desired future course of an organization"
strcolumnB(7) = "adaptive system"
strcolumnB(8) = "descriptive model"
strcolumnB(9) = "man made system"
strcolumnB(10) = "open loop system"

stranswers(1) = "man made system"
stranswers(2) = "adaptive system"
stranswers(3) = "normative model"
stranswers(4) = "descriptive model"
stranswers(5) = "desired future course of an organization"
stranswers(6) = "closed loop system"
stranswers(7) = "open loop system"
stranswers(8) = "inventory control"
stranswers(9) = "data model"
stranswers(10) = "a key area for organizational success"
End Sub

Private Sub Matching_Columns_Paper7()
strcolumnA(1) = "PERT"
strcolumnA(2) = "multimedia"
strcolumnA(3) = "breakeven analysis"
strcolumnA(4) = "efficiency"
strcolumnA(5) = "economy"
strcolumnA(6) = "check digit"
strcolumnA(7) = "structured walkthrough"
strcolumnA(8) = "CASE tools"
strcolumnA(9) = "quality assurance"
strcolumnA(10) = "plug compatible"

strcolumnB(1) = "common method for checking transposition errors"
strcolumnB(2) = "analysis where cost of new system equals cost of old"
strcolumnB(3) = "system developed is too complex to operate"
strcolumnB(4) = "cost overruns are very high"
strcolumnB(5) = "automation of entire system development process"
strcolumnB(6) = "review of a system or its software by persons involved in development team"
strcolumnB(7) = "frequently, analysts for reason of cost, consider an add-on, for a particular make of computer, that is not manufactured by original vendor"
strcolumnB(8) = "system specific testing"
strcolumnB(9) = "system project management"
strcolumnB(10) = "implementation training"

stranswers(1) = "system project management"
stranswers(2) = "implementation training"
stranswers(3) = "analysis where cost of new system equals cost of old"
stranswers(4) = "system developed is too complex to operate"
stranswers(5) = "cost overruns are very high"
stranswers(6) = "common method for checking transposition errors"
stranswers(7) = "review of a system or its software by persons involved in development team"
stranswers(8) = "automation of entire system development process"
stranswers(9) = "system specific testing"
stranswers(10) = "frequently, analysts for reason of cost, consider an add-on, for a particular make of computer, that is not manufactured by original vendor"
End Sub

Private Sub Matching_Columns_Paper8()
strcolumnA(1) = "inventory procurement"
strcolumnA(2) = "stores receipt"
strcolumnA(3) = "algorithm"
strcolumnA(4) = "system analyst is"
strcolumnA(5) = "common pool of shared data which is linked enterprise wide"
strcolumnA(6) = "pseudo code precedes"
strcolumnA(7) = "response time crucial"
strcolumnA(8) = "normalization simplifies"
strcolumnA(9) = "data encryption essential"
strcolumnA(10) = "data dictionary"

strcolumnB(1) = "data about data"
strcolumnB(2) = "purchase order"
strcolumnB(3) = "coding"
strcolumnB(4) = "EFT"
strcolumnB(5) = "agent of change"
strcolumnB(6) = "goods receipt not"
strcolumnB(7) = "generic solution to problems"
strcolumnB(8) = "data base"
strcolumnB(9) = "data management"
strcolumnB(10) = "real time systems"

stranswers(1) = "purchase order"
stranswers(2) = "goods receipt not"
stranswers(3) = "generic solution to problems"
stranswers(4) = "agent of change"
stranswers(5) = "data base"
stranswers(6) = "coding"
stranswers(7) = "real time systems"
stranswers(8) = "data management"
stranswers(9) = "EFT"
stranswers(10) = "data about data"
End Sub

Private Sub Matching_Columns_Paper9()
strcolumnA(1) = "4GL"
strcolumnA(2) = "procedure oriented"
strcolumnA(3) = "a set of tables"
strcolumnA(4) = "hash total"
strcolumnA(5) = "LAN"
strcolumnA(6) = "real time systems"
strcolumnA(7) = "cost benefit analysis"
strcolumnA(8) = "audit trail"
strcolumnA(9) = "data store"
strcolumnA(10) = "analysis task"

strcolumnB(1) = "an economic justification for a system"
strcolumnB(2) = "software requirements specs produced at culmination of"
strcolumnB(3) = "a repository of data, arrow indicating net inputs/ outputs to the store"
strcolumnB(4) = "languages most commonly used to develop information processing applications are"
strcolumnB(5) = "general class of languages designed to improve the efficient of application development processes are called "
strcolumnB(6) = "method which allows tracing of transactions from time of recording up to report generation"
strcolumnB(7) = "to ensure data integrity DBMS when implemented maintain"
strcolumnB(8) = "security feature"
strcolumnB(9) = "airline ticketing is"
strcolumnB(10) = "file server is part of"

stranswers(1) = "general class of languages designed to improve the efficient of application development processes are called "
stranswers(2) = "languages most commonly used to develop information processing applications are"
stranswers(3) = "to ensure data integrity DBMS when implemented maintain"
stranswers(4) = "security feature"
stranswers(5) = "file server is part of"
stranswers(6) = "airline ticketing is"
stranswers(7) = "an economic justification for a system"
stranswers(8) = "method which allows tracing of transactions from time of recording up to report generation"
stranswers(9) = "a repository of data, arrow indicating net inputs/ outputs to the store"
stranswers(10) = "software requirements specs produced at culmination of"
End Sub

Private Sub Matching_Columns_Paper10()
strcolumnA(1) = "transaction"
strcolumnA(2) = "range check"
strcolumnA(3) = "scrolling"
strcolumnA(4) = "source document"
strcolumnA(5) = "pull down menu"
strcolumnA(6) = "structure chart"
strcolumnA(7) = "JAVA"
strcolumnA(8) = "presentation chart"
strcolumnA(9) = "bench marking"
strcolumnA(10) = "check digit"

strcolumnB(1) = "selection of computer, hardware and software"
strcolumnB(2) = "selection of an option by user"
strcolumnB(3) = "hierarchical relationship between modules of a system"
strcolumnB(4) = "reasonable limits of input"
strcolumnB(5) = "business event"
strcolumnB(6) = "connections between input files, output files & programs"
strcolumnB(7) = "module eleven technique"
strcolumnB(8) = "vertical movement of display"
strcolumnB(9) = "internet"
strcolumnB(10) = "form containing input data"

stranswers(1) = "business event"
stranswers(2) = "reasonable limits of input"
stranswers(3) = "vertical movement of display"
stranswers(4) = "form containing input data"
stranswers(5) = "selection of an option by user"
stranswers(6) = "hierarchical relationship between modules of a system"
stranswers(7) = "internet"
stranswers(8) = "connections between input files, output files & programs"
stranswers(9) = "selection of computer, hardware and software"
stranswers(10) = "module eleven technique"

End Sub

Private Sub Matching_Columns_Paper11()
strcolumnA(1) = "on line"
strcolumnA(2) = "feasibility study"
strcolumnA(3) = "interview"
strcolumnA(4) = "distribute processing"
strcolumnA(5) = "TOES analysis"
strcolumnA(6) = "structured design"
strcolumnA(7) = "grid chart"
strcolumnA(8) = "cost benefit analysis"
strcolumnA(9) = "turnaround time"
strcolumnA(10) = "modularity"

strcolumnB(1) = "performance indicator"
strcolumnB(2) = "systems analysis"
strcolumnB(3) = "modularity"
strcolumnB(4) = "structured design"
strcolumnB(5) = "tangible/intangible review"
strcolumnB(6) = "data gathering method"
strcolumnB(7) = "cost benefit analysis"
strcolumnB(8) = "SAM missile"
strcolumnB(9) = "feasibility report"
strcolumnB(10) = "linking computers at local sites"

stranswers(1) = "SAM missile"
stranswers(2) = "cost benefit analysis"
stranswers(3) = "data gathering method"
stranswers(4) = "linking computers at local sites"
stranswers(5) = "feasibility report"
stranswers(6) = "modularity"
stranswers(7) = "systems analysis"
stranswers(8) = "tangible/intangible review"
stranswers(9) = "performance indicator"
stranswers(10) = "structured design"
End Sub

Private Sub Matching_Columns_Paper12()
strcolumnA(1) = "prototype"
strcolumnA(2) = "DFD"
strcolumnA(3) = "data dictionary"
strcolumnA(4) = "data store"
strcolumnA(5) = "data flow"
strcolumnA(6) = "turnaround document"
strcolumnA(7) = "CASE"
strcolumnA(8) = "interfaces"
strcolumnA(9) = "DSS"
strcolumnA(10) = " OLTP"

strcolumnB(1) = "system boundary"
strcolumnB(2) = "theatre reservation"
strcolumnB(3) = "strategic management"
strcolumnB(4) = "output later used s input"
strcolumnB(5) = "tool for system development"
strcolumnB(6) = "input to system design"
strcolumnB(7) = "model"
strcolumnB(8) = "DFD"
strcolumnB(9) = "arrow"
strcolumnB(10) = "repository of data about data"

stranswers(1) = "model"
stranswers(2) = "input to system design"
stranswers(3) = "repository of data about data"
stranswers(4) = "DFD"
stranswers(5) = "arrow"
stranswers(6) = "output later used s input"
stranswers(7) = "tool for system development"
stranswers(8) = "system boundary"
stranswers(9) = "strategic management"
stranswers(10) = "system boundary"
End Sub

Private Sub Matching_Columns_Paper13()
strcolumnA(1) = "design phase"
strcolumnA(2) = "walkthrough"
strcolumnA(3) = "beta test"
strcolumnA(4) = "audit trail"
strcolumnA(5) = "lease"
strcolumnA(6) = "stub testing"
strcolumnA(7) = "check point restart"
strcolumnA(8) = "digital signature"
strcolumnA(9) = "UPS"
strcolumnA(10) = "IPR"

strcolumnB(1) = "proving module"
strcolumnB(2) = "communication security"
strcolumnB(3) = "piracy"
strcolumnB(4) = "process recovery"
strcolumnB(5) = "alternative hardware induction"
strcolumnB(6) = "power failure"
strcolumnB(7) = "peer review"
strcolumnB(8) = "SDLC"
strcolumnB(9) = "tracking data item"
strcolumnB(10) = "proving an IS"

stranswers(1) = "SDLC"
stranswers(2) = "peer review"
stranswers(3) = "proving an IS"
stranswers(4) = "tracking data item"
stranswers(5) = "alternative hardware induction"
stranswers(6) = "proving module"
stranswers(7) = "process recovery"
stranswers(8) = "communication security"
stranswers(9) = "power failure"
stranswers(10) = "piracy"
End Sub

Private Sub Matching_Columns_Paper14()
strcolumnA(1) = "MMI"
strcolumnA(2) = "CAI"
strcolumnA(3) = "MICR"
strcolumnA(4) = "WAN"
strcolumnA(5) = "BIS"
strcolumnA(6) = "ATM"
strcolumnA(7) = "CTD"
strcolumnA(8) = "GANTT"
strcolumnA(9) = "OLTP"
strcolumnA(10) = "LCD"

strcolumnB(1) = "passenger reservation"
strcolumnB(2) = "security"
strcolumnB(3) = "real time system"
strcolumnB(4) = "project management"
strcolumnB(5) = "self learning"
strcolumnB(6) = "palm tops"
strcolumnB(7) = "cheque clearing"
strcolumnB(8) = "joystick"
strcolumnB(9) = "ISI mark"
strcolumnB(10) = "NICNET"

stranswers(1) = "cheque clearing"
stranswers(2) = "self learning"
stranswers(3) = "NICNET"
stranswers(4) = "cheque clearing"
stranswers(5) = "ISI mark"
stranswers(6) = "real time system"
stranswers(7) = "security"
stranswers(8) = "project management"
stranswers(9) = "passenger reservation"
stranswers(10) = "palm tops"
End Sub

Private Sub Matching_Columns_Paper15()
strcolumnA(1) = "mail box"
strcolumnA(2) = "barcode"
strcolumnA(3) = "OCR"
strcolumnA(4) = "DSS"
strcolumnA(5) = "MTBF"
strcolumnA(6) = "MTR"
strcolumnA(7) = "prithvi"
strcolumnA(8) = "hashing"
strcolumnA(9) = "MAN"
strcolumnA(10) = "EFT"

strcolumnB(1) = "MTNL Delhi"
strcolumnB(2) = "real time"
strcolumnB(3) = "system maintenance"
strcolumnB(4) = "key transformation"
strcolumnB(5) = "banking"
strcolumnB(6) = "email"
strcolumnB(7) = "reliability"
strcolumnB(8) = "library"
strcolumnB(9) = "CEO"
strcolumnB(10) = "examination evaluation"

stranswers(1) = "email"
stranswers(2) = "library"
stranswers(3) = "examination evaluation"
stranswers(4) = "CEO"
stranswers(5) = "reliability"
stranswers(6) = "system maintenance"
stranswers(7) = "real time"
stranswers(8) = "key transformation"
stranswers(9) = "MTNL Delhi"
stranswers(10) = "banking"
End Sub

Private Sub questions_paper1()
StrQuestions(1) = "In a two year post graduate course of a business school, there are three types of students in the first year." & Chr(13) & _
                  "A. Normal students;" & Chr(13) & _
                  "b. Those who failed the first year in the previous year and are asked to repeat the first year" & Chr(13) & _
                  "c. Students from other universities, who spend first yea in this school on a special permission" & Chr(13) & _
                  "A student of first year is required to undergo 15 different courses. In each course, a student is awarded a letter grade, A,B,C,D or False. A is the best possible grade, B is the next best, and so on, and False the worst." & Chr(13) & _
                  "At the end of the first year, the head of the program has to convey to each student, except the outside students, the school's decision regarding the promotion of the student to second year" & Chr(13) & _
                  "A normal student gets promoted if the/she gets not ore than 3 grades which are D or worse. in case the student cannot be promoted, then he/she will be asked to repeat the entire first year, if he/she got and A in the least one course of B's in at least two courses." & Chr(13) & _
                  "Otherwise, the student has to withdraw from the program. In the case of a student who has repeated the first year, the promotion criteria are the same as above, except that the possibility of repeating does not exist. Either he / she gets promoted or has to withdraw from the program." & Chr(13) & _
                  "In the case of outside students, the  grades are simply reported to the student's parent university, and no other decision is needed." & Chr(13) & _
                  "a. Draw a decision table that describes the above rules from the point of view of the head of the postgraduate program. Minimize the number of rows and columns in the table" & Chr(13) & _
                  "b. Describe the above rules also through a pseudo-code (or Structured English)"
StrQuestions(2) = "a. What are Structured Analysis and Data structures?" & Chr(13) & _
                  "b. What is the relationship that Data Structures have with Data Elements, Data Flows and Data stores?"
StrQuestions(3) = "A disk has 15 recording surfaces , on each surface there are 1000 tracks. Each track is divided into 60 sectors and each sector has 640 bytes of storage." & Chr(13) & _
                  "a. How many cylinders does the disk have?" & Chr(13) & _
                  "b. How many tracks are there percylinde?" & Chr(13) & _
                  "c. What’s the storage capacity of the disk?"
StrQuestions(4) = "a. What is the purpose behind Normalization?" & Chr(13) & _
                  "b. You have obtained, after analyzing a system the following Unnormalised Relation -R:" & Chr(13) & _
                  "R(Emp_no#,Emp_name,STore_branch,Department_no,Item_no,Item_description,Sales_price)" & Chr(13) & _
                  "Convert this relation into 1st, 2nd and 3rd normalized form"
End Sub
Private Sub questions_paper2()
StrQuestions(1) = "Fairdeal corpn. manufactures domestic applications and markets these throughout the country. Its product range includes refrigerators, gas and electric cooking ranges, washing machines, dryers and toasters. The company currently sells 100000 appliances (on an average) each year. The company requires its customers to complete and send in a standard form to activate the guarantee for each appliance that it sells." & Chr(13) & _
                  "From these forms the company intends to develop a direct mailing list for promotion. The marketing department of the company feels that people who are satisfied with the products ,quite likely, will buy another one. They are also interested in analyzing the patterns, if any, in which people buy appliances. for example," & Chr(13) & _
                  "A. Which geographic area do the specific products sell well?" & Chr(13) & _
                  "B. Where do the high provide modules sell well?" & Chr(13) & _
                  "C. What  is the possibility that someone who has bought a washing machine will buy a dryer within next 12 - 15 - 18 months ? etc." & Chr(13) & _
                  "Design the standard guarantee input to the system."
StrQuestions(2) = "ABC manufacturing company employed XYZ consultants to develop a package that can find for them, during each planning horizon the optimal product mix; that is which products are to be produced and in what quantity. " & Chr(13) & _
                  "The consultants formulated the optical product mix problem as a linear program and the package supplied by them solves this linear program. it is implemented on a PC. The package is men driven. The only menu options offered by the package are" & Chr(13) & _
                  "A. Create data file for LP;" & Chr(13) & _
                  "B. Run LP on the data file" & Chr(13) & _
                  "c.Quit" & Chr(13) & _
                  "In the A option the user is asked through various prompts the number of variables and constraints, and to input data on objective function and constraints. Once the entire data is given it is saved with a name know only to the system. if a data file is already existing, it will be replaced with the file newly input." & Chr(13) & _
                  "In the B option the system access the data file and solves the corresponding LP and exhibits the optimal solution obtained by the system or error messages, if any, on the screen." & Chr(13) & _
                  "a. Can the above package be called a Decision Support system (DSS)? If yes justify; if NO why not?" & Chr(13) & _
                  "b. Suggest some extensions to the package that would in your opinion qualify it to be called a true DSS."
StrQuestions(3) = "A project management package finds the earliest possible and latest permissible start and finish times of activities in a project. the input for the package is set of activities that constituent the project their estimated durations and precedence relations. In its computations, the package uses the standard forward pass and backward pass computation technique that is found in most standard text books on PERT/CPM. " & Chr(13) & _
                  "The package offers a menu for management of input and output data and for generating the results. But for this there is no other interaction between the user and the package while it is running. Can this package be called an expert system? If YES justify; if NO why not?"
StrQuestions(4) = "XYZ supermarket located in a pajot city of India, has three checkout counters. In order to improve the check out process, a PC has been set up at each of these counters along with a printer." & Chr(13) & _
                  "Each item in the super market is packed and each packet carries on it, a price tag on which are printed its item( or packet) code and price." & Chr(13) & _
                  "Each customer picks out the desired packets from the shelves, places them in a standard shopping basket supplied by the supermarket and carries the basket to one of the checkout counters. The counter operator takes out each of the packets from the basket, keys n the twelve digit item code." & Chr(13) & _
                  "Once all the packets are dealt with the operator presses a key that causes the computer to compute and show the total amount of the bill. The attendant reads out the amount to the customer and collects the cast. The amount given by the customer is keyed in; then the computer displays the amount to be returned to the customer" & Chr(13) & _
                  "(in case the amount given by the customer is more than the amount of the bill). The operator returns the balance, if any o he customer , and presses key o initiate the data entry for he next customer. A his point he entire bill gets printed on the printer in duplicate the original is handed over to the customer and the duplicate is preserved by the operator. A copy of the bill is also stored on the disk automatically." & Chr(13) & _
                  "The above system has been operating successfully for the last six months." & Chr(13) & _
                  "Answer he following questions related to the above system:" & Chr(13) & _
                  "a. Identify the data files and their fields used in generating the customer bill." & Chr(13) & _
                  "b. A study has shown that the slowest step in the above process is the data entry by the counter operator. Suggest some ways o make the data entry step faster." & Chr(13) & _
                  "c. The proprietor of the supermarket is wondering whether he could derive any information useful to him in his purchase decisions, by incorporating additional features in the above system. Suggest one extension to meet his objective, and design a weekly report that can be generated for use by the proprietor in this purchase planning."
StrQuestions(5) = "With respect to purchasing and inventory control systems explain any three of the following:" & Chr(13) & _
                  "A. Why do retail outlets carry inventory" & Chr(13) & _
                  "B. Inventory carrying cost" & Chr(13) & _
                  "C. Procurement lead time" & Chr(13) & _
                  "D. Bill of material"
StrQuestions(6) = "With respect to manufacturing systems, explain the following :" & Chr(13) & _
                  "a.Job shop" & Chr(13) & _
                  "b.flow shop" & Chr(13) & _
                  "c.production scheduling" & Chr(13) & _
                  "D. Work in process inventory" & Chr(13) & _
                  "E. Master production schedule" & Chr(13) & _
                  "False.Project "
StrQuestions(7) = "ABC electronics manufactures a single consumer product and is facing a lot of competition/. The owner of ABC has realized that there is very little difference between the quality and features of his own product and those of the competitors. There is also no room for either changing the price of increasing the quality." & Chr(13) & _
                  "The only way of competing in the market is by capturing a large market share, through aggressive and effective advertising and through improved after sales service. The owner has recently takes a lot of steps to implement the above ides and to improve his market share buy about 20% in the next one year. " & Chr(13) & _
                  "First he has increased the dealership network to offer better service, and second he has appointed a new MIS manager. Until now there has been no computerization at ABC. The MIS manager worked hard for a week and prepared a plan for new computer applications in which" & Chr(13) & _
                  "- the first priority is given to a system that reports on the daily work in process inventory status. The owner of the company will be provided with a computer terminal through which he can find how much material in at what stage of manufacture, and how much is ready to be dispatched to the market;" & Chr(13) & _
                  "- at second priority is the application that reports on employee productivity." & Chr(13) & _
                  "The MIs manager has estimated that for both the systems be mature, it would take about two to three years. He also submitted a detailed wok schedule and feasibility plan for these systems. The major reason for choosing the above two applications was that most of the manufacturing shops use these systems and the MIs manager successfully implemented hem in a firm with which he had worked earlier." & Chr(13) & _
                  "Was the MIs manager right in his choice of applications for computerization? Why or why not? If your answer is no, suggest which hare as should have been chosen for computerization by the MIS manager; yes justify."
StrQuestions(8) = "A sales analysis system for a marketing company is designed to have the following modules:" & Chr(13) & _
                  "A. Master file creation and amendment module: Creates and edits the customer, product, and salesman master files; also does data validation" & Chr(13) & _
                  "B. Transaction file creation and amendment module: creates and edits the file of sales transactions of a quarter; also dos data validation" & Chr(13) & _
                  "C. Sales summary module: creates a file that consists of a summary of the sales transaction file; the summary shows the salesman wise, product wise and customer wise sales during the quarter." & Chr(13) & _
                  "D. Sales budgeting module: creates and edits a file of quarterly sales targets routine, customer wise and sales mansion. Uses the master files created in modules A& B." & Chr(13) & _
                  "E. Sales analysis module: compares the budgeted and actual sales productive, customer wise and salesman wise for a quarter. Gives summary statistics such as the mean sales, standard deviation, frequency distribution etc. Stores the results in a disk file. Uses the master files, the sales budget file, and sales summary file." & Chr(13) & _
                  "a. Propose a sequence for testing the above modular system while it is being developed." & Chr(13) & _
                  "b. The above system is expected to the operated entirely buy the staff of the marketing department. These staff are not computer professionals. The systems developers propose to supply them with the following documentation." & Chr(13) & _
                  "A. Source listings of each of the modules in the hard copy form, along with a pseudo code explaining the logic of each module;" & Chr(13) & _
                  "B. A listing of the object code of each module;" & Chr(13) & _
                  "C. A chapter on the file structures used and disk space and memory requirements" & Chr(13) & _
                 "Is the above documentation relevant to the marketing department? Suggest modifications of your own, for what is to be included in the documentation to be given to the marketing department."
StrQuestions(9) = "Describe the files that would be necessary for each of the following computer systems. Your answer for each of the following computer systems. Your answer should include:" & Chr(13) & _
                   "A. name of file, a selection of typical contents of a record, the record key and type of file .e/.g." & Chr(13) & _
                   "Customer file" & Chr(13) & _
                   "Customer account number - Record key, Name and Address" & Chr(13) & _
                   "Credit limit" & Chr(13) & _
                   "Master file" & Chr(13) & _
                   "B. The method of working or any special requirements should also be stated, e.g. real-time updating" & Chr(13) & _
                   "a. A mail order system which receives orders from agents, checks which items are in stock produces a file of ordered items in stock and them prints invoice (by referring to a file of agents names and addresses)" & Chr(13) & _
                   "b. A police system for recording details of stolen cars and answering queries from on duty police officers" & Chr(13) & _
                   "c. A computer dating agency which matches applicants requirements with previously stored data relating to other people." & Chr(13) & _
                   "d. A large shipping company which has a system to locate ships and keeps details of cargo, destination insurance etc."
End Sub
Private Sub questions_paper3()
StrQuestions(1) = "Explain the following" & Chr(13) & _
                  "A.Project" & Chr(13) & _
                  "b.Project scheduling" & Chr(13) & _
                  "C.Critical Path" & Chr(13) & _
                  "d.Milestones" & Chr(13) & _
                  "E.Checkpoints" & Chr(13) & _
                  "False. Project review"
StrQuestions(2) = "It is estimated that a project takes 15 weeks for its completion. There are five tasks A-E in the project. The schedule calls for:" & Chr(13) & _
                  "A. Task B to begin when task A is completed" & Chr(13) & _
                  "B. Task C overlaps the end of task B and the beginning of task D, by one week each;" & Chr(13) & _
                  "C. Task D will begin at week 7 and will end at week 11;" & Chr(13) & _
                  "D. At some points in the schedule, work n more than on task will be undertaken at the same time;" & Chr(13) & _
                  "E. Task E is to begin at week 8 and end at week 15."
StrQuestions(3) = "List and describe in not less than 500 words the profile of at least four typical applications of DSS in a typical: medium sized and large business organizations."
StrQuestions(4) = "A. What is strategic planning? Relate strategic planning to management control and operational control." & Chr(13) & _
                  "b. Why are strategic plans and management control integrated within a DSS framework?"
StrQuestions(5) = "Discuss the limitations of Management Information System. Do you think they have been overcome with the help of DSS? Comment"
StrQuestions(6) = "The office of the department of computer science organizes and handles seminars/workshops" & Chr(13) & _
                  "When a person wanton to register for a seminar, he /she sends a request" & Chr(13) & _
                  "The office clerk notes the person's name, address, and telephone number and organization name on a registration form. Copy of the form is placed in the master file, an acknowledgment letter prepared and sent to the registrant." & Chr(13) & _
                  "Another copy is sent to the billing clerk who creates and invoice record for the registrant in another central file and sends a copy of the invoice to the registrants." & Chr(13) & _
                  "One month before the seminar the clerk prepares nametags and seminar material for each participant." & Chr(13) & _
                  "The registrants are expected to pay in advance of the seminar." & Chr(13) & _
                  "When a cheque comes in the clerk verifies the registration against the registration record and enters the information that the payment has been received." & Chr(13) & _
                  "In case registrants do not send the advance payment, their namesake listed on a separate sheet. This defaulter sheet is given to the seminar coordinator who takes the appropriate action, which may be:" & Chr(13) & _
                  "A. A person is allowed to attend the seminar, if he agrees o make the payment on the first day before the seminar starts or" & Chr(13) & _
                  "B. Brings in a sanction letter from the head of his organization" & Chr(13) & _
                  "All registration records are maintained in a master file and used as a mailing list on an ongoing basis." & Chr(13) & _
                  "a. Design data flow diagram for the above registration procedure for the seminar" & Chr(13) & _
                  "b. List major files along with specifications of the key fields" & Chr(13) & _
                  "c. List major reports that should be generated by the system"
StrQuestions(7) = "Bhayankar state public service commission (BSPSC) has a new CEO, Alec clever, who is keen to get its operations computerised." & Chr(13) & _
                  "The functional parameters of BXPSC are:" & Chr(13) & _
                  "- Job vacancies in the state can only be filled though it;" & Chr(13) & _
                  "- Caters to an estimated 360 categories of posts;" & Chr(13) & _
                  "- Whereas an interview is an essential requirement for recruitment there is a dichotomy in preliminary procedures depending on the category -" & Chr(13) & _
                  "* In some there is only an examination;" & Chr(13) & _
                  "* In others there is an examination followed by a screening phase at BSPSC" & Chr(13) & _
                  "- Question papers are set by experts selected from a panel;" & Chr(13) & _
                  "- There is also a panel of evaluators to mark answer books; setter can also be an evaluator and vice versa;" & Chr(13) & _
                  "- Considering the importance of the function there is a special panel of experts to conduct interview indicating areas overspecialization;" & Chr(13) & _
                  "- For serving on any of these functions an honorarium is paid;" & Chr(13) & _
                  "- Pests are advertised thorough gazettes and newspapers;" & Chr(13) & _
                  "- Applications have to be accompanied by fee that differs with the post;" & Chr(13) & _
                  "- Reservation of posts exists for most categories as per Govt directive;" & Chr(13) & _
                  "- Results must be polished with 6 weeks of the examination in selected newspapers; an explanation for non-compliance has to be sent to the governor." & Chr(13) & _
                  "a. Assuming that there is only a manual system how would you plan for entering the computerized era? Support your answer with a network diagram" & Chr(13) & _
                  "b. Identify three major database lies and in each mention six fields as also where you will get the source and subsequent updating;" & Chr(13) & _
                  "c. would a good RDBMS meet the needs of BSPSC? if NOT justify; if so, how would you plan it?"
End Sub
Private Sub questions_paper4()
StrQuestions(1) = "This is the age of acronyms and abbreviations. Out of the following briefly in not more than 100 words describe SIX:" & Chr(13) & _
                  "a. ERP  b. CCITT    c. OSI  d. PARAM    e. OLAP False. PPP  g. HTML"
StrQuestions(2) = "Answer in not more than 500 words three of the following:" & Chr(13) & _
                  "a. assuring security on communication channels" & Chr(13) & _
                  "b. gaining access to internet;" & Chr(13) & _
                  "c. role of management in MIS development" & Chr(13) & _
                  "d. Virtual reality role in training"
StrQuestions(3) = "This is the age of acronyms and abbreviations. Out of the following briefly in not more than 100 words describe SIX:" & Chr(13) & _
                  "a. EMC  b. FCC  c. IEEE d. DVD  e. CHAP False. DOVM g. HTML"
StrQuestions(4) = "This is the aloof acronyms and abbreviations. Out of the following briefly in not more than 100 words describe SIX:" & Chr(13) & _
                  "a. GIST b. EFT  c. VRML d. IFIP e. EDI  False. PAP  g. NC"
StrQuestions(5) = "This is the age of acronyms and abbreviations. Out of the following briefly in not more than 100 words, describe SIX:" & Chr(13) & _
                  "a. NII  b. ISP  c. MAC  d. URM  e. TCP/IP   False. VSNL g. DNS"
StrQuestions(6) = "This is the age of acronyms and abbreviations. Out of the following briefly in not more than 1`00 words describe SIX:" & Chr(13) & _
                  "a. IRC  b. GII  c. NASSCOM  d. CDOT e. CDAC False. GUI  g. IRG"
End Sub

Private Sub TrueorFalse_Paper1()
StrQuestions(1) = "The major objective of a structured interview is to allow the interviewee to discuss his/her role in the organization structure the way he/she pleases, without any interruption from or control by the interviewer."
StrQuestions(2) = "A bar menu is a type of interface in which the user provides the require input interactively by answering a series of questions appearing on the control bar area of the screen."
StrQuestions(3) = "Data entry is a step in the input preparation during which data from a source usually a source document is converted to data in a machine-readable form"
StrQuestions(4) = "A sequential access file can reside only on a magnetic tape."
StrQuestions(5) = "A logical data flow diagram should record explicitly for each concerned data file whether it is of sequential access type."
StrQuestions(6) = "A data dictionary is used for spelling checks in a word processor"
StrQuestions(7) = "The term controlled redundancy rears to the practice of allowing the storage of the same data in more than replace subject to some rules so that such duplicating, if any is known to the system"
StrQuestions(8) = "In a real time system, the system monitors its environment and then either controls devices that change the environment or helps humans effect the change."
StrQuestions(9) = "A logical model of system is concerned solely with internal consistency of its description."
StrQuestions(10) = "Selection of equipment and software in a software development project precedes the requirements analysis phase."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper2()
StrQuestions(1) = "In system design hat must be done , how to do it, who must do it, and when to do it, must be decided in advance."
StrQuestions(2) = "A decision table is a pictorial representation of data flow."
StrQuestions(3) = "The primary functions of PERT charts are for planning and controlling complex system projects."
StrQuestions(4) = "The basic sub systems of any service organization are Marketing, Manufacturing and Finance"
StrQuestions(5) = "Top management involvement in system project should be minimal"
StrQuestions(6) = "A procedure is a series of logical steps by which a job is accomplished."
StrQuestions(7) = "Simplicity of a well designed system refers to its capability to change in response to new environmental conditions."
StrQuestions(8) = "An example of an internal report is an inventory summary."
StrQuestions(9) = "The current trend in MIS is toward reducing data file size."
StrQuestions(10) = "Log of computer time is a part of securit7y controls"

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper3()
StrQuestions(1) = "The batch  processing approach allow data to be entered as they occur or information to be retrieved relatively quickly."
StrQuestions(2) = "Generally there is no logical starting point for system design."
StrQuestions(3) = "It is necessary to specify output requirements before determining the content of the output."
StrQuestions(4) = "Variable information is defined as that information which does not change for each line item in a report."
StrQuestions(5) = "Data life cycle and SDLC are NOT synonymous."
StrQuestions(6) = "Context Diagrams are applicable only to the top level."
StrQuestions(7) = "HIPO diagrams are used to develop packages."
StrQuestions(8) = "VTOC shows a hierarchy of documents"
StrQuestions(9) = "Audit trails insure against loss of data."
StrQuestions(10) = "ERD reflects he business requirements of proposed system.   "

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub
Private Sub TrueorFalse_Paper4()
StrQuestions(1) = "Validation errors are detected by check digits."
StrQuestions(2) = "A output of 2500 keystrokes per hour, is good standard for Data Entry"
StrQuestions(3) = "Open systems are not safe."
StrQuestions(4) = "Aliases are included in data dictionaries."
StrQuestions(5) = "Prototyping model is created at the problem definition stage of system development"
StrQuestions(6) = "In LAN, each computer can fulfil a function."
StrQuestions(7) = "Managers need information, which is critical to heir performance and success."
StrQuestions(8) = "All managers In a big organization use the same type of information."
StrQuestions(9) = "Entities may have similar attributes."
StrQuestions(10) = "In Sequential type of storage devices, any record can be accessed and time taken to write or read any record will be same."

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
Private Sub TrueorFalse_Paper5()
StrQuestions(1) = "Sequential files are suited for on line inquire processing."
StrQuestions(2) = "On line processing and real time processing are same."
StrQuestions(3) = "Decentralized processing and distributed processing are same"
StrQuestions(4) = "Classifying and sorting data are same."
StrQuestions(5) = "A program of size grater than the size of the main memory cannot be run."
StrQuestions(6) = "A Structured decision is not routing"
StrQuestions(7) = "A business transaction is a representation of reality in symbols"
StrQuestions(8) = "An MIS manager deals with strategic planning of MIS."
StrQuestions(9) = "Information means the same as data."
StrQuestions(10) = "Systems approach involves studying each element of system deeply and assuming that the interactions between elements weak"

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub
Private Sub TrueorFalse_Paper6()
StrQuestions(1) = "External data is needed in strategic planning."
StrQuestions(2) = "Balance sheet generation is necessary component of every personnel management system."
StrQuestions(3) = "Check digit is a technique used by financial systems to automatically generate voucher numbers"
StrQuestions(4) = "A decision support system need not necessarily use computer graphics"
StrQuestions(5) = "One of the reasons for resistance to introduction of computer based MIS in an organization is that it would require the staff to learn to  do some thing  differently"
StrQuestions(6) = "User involvement is highly avoidable n the development of decision support systems"
StrQuestions(7) = "For the development of a knowledge base in a medical diagnosis system, the main source of production rules are the patients."
StrQuestions(8) = "A file server in a LAN must be a supercomputer."
StrQuestions(9) = "A relational data base system stores all its files in the form of sequential files."
StrQuestions(10) = "Any DBMS that uses the CODASYL network model requires necessarily a computer network for its implementation."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub
Private Sub TrueorFalse_Paper7()
StrQuestions(1) = "A decision support system need not necessarily use computer graphics."
StrQuestions(2) = "Decision support systems support structured decisions."
StrQuestions(3) = "Each data flow in the DFD has a corresponding entry in the Data Dictionary."
StrQuestions(4) = "Module coupling refers to he relationship among elements within a module"
StrQuestions(5) = "Quality assurance is involved in every phase of the SDLC"
StrQuestions(6) = "A prototype is essentially a complete information system but without the database"
StrQuestions(7) = "A Management control system is a component of management information system."
StrQuestions(8) = "Management information system is sub set or component of management system."
StrQuestions(9) = "Transaction processing and real time systems are the same."
StrQuestions(10) = "Turn around documents are used for control purposes.    "

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
Private Sub TrueorFalse_Paper8()
StrQuestions(1) = "Critical success factors are required to design an MIS."
StrQuestions(2) = "When there are not enough transactions with the main computer system, remote job entry is done under off-line mode "
StrQuestions(3) = "The detailed study/investigation of the present system is frequently referred to as system planning."
StrQuestions(4) = "The approach to system output by using the principle of management by perception is considered to be forward looking."
StrQuestions(5) = "The next major step before system design and after feasibility study is programming"
StrQuestions(6) = "'Schedule of delivery' is an example of a DSS"
StrQuestions(7) = "The procurement process for Hardware starts after system analysis."
StrQuestions(8) = "Tangible benefits by their very nature."
StrQuestions(9) = "In a typical data processing context, whether master files are updated to procure desired output is known as transaction processing."
StrQuestions(10) = "The major focus of system methods and procedures is to relate input at to final output that is meaningful to the potential users of the system."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub
Private Sub TrueorFalse_Paper9()
StrQuestions(1) = "The components that make up any system is known as its boundary"
StrQuestions(2) = "The members of the executive steering committee members are generally drawn from the systems division"
StrQuestions(3) = "Whether a proposed system can provide right information for the organization’s personnel falls under the study of economic feasibility."
StrQuestions(4) = "'Removal of functional dependency from a relation' falls under the category of Third form of Normalization"
StrQuestions(5) = "The term Stub is met with in module testing"
StrQuestions(6) = "CASE is not used in the SDLC"
StrQuestions(7) = "Bench marking is part of SDLC"
StrQuestions(8) = "Audit ability is not part of SDLC"
StrQuestions(9) = "Multi media can be harnesses for end user training in the implementation phase of the SDLC  "
StrQuestions(10) = "Sizing is a part of the programming phase of SDLC."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub
Private Sub TrueorFalse_Paper10()
StrQuestions(1) = "Normalization simplifies updating"
StrQuestions(2) = "Walkthrough is not necessary in the requirement analysis stage of the SDLC"
StrQuestions(3) = "Error correction is least troublesome at the requirement analysis stage."
StrQuestions(4) = "Design  specifications must include a PERT network."
StrQuestions(5) = "Program flow charts are essential tools for data collecting."
StrQuestions(6) = "Structured English is an essential fool for systems analysis"
StrQuestions(7) = "Maintainability is one of the factors that determine system quality"
StrQuestions(8) = "Developing end user information needs  is a vital factor in System design"
StrQuestions(9) = "Non involvement of end users is a major factor that leads to the failure of MIS."
StrQuestions(10) = "To implement a MIS a computer is indispensable"

stranswers(1) = "True"
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
Private Sub TrueorFalse_Paper11()
StrQuestions(1) = "There is NO control representation in a DFD except on one situation."
StrQuestions(2) = "One can easily represent data structure in flow charting."
StrQuestions(3) = "Decision trees describe decision s while avoiding difficulties in communications."
StrQuestions(4) = "Physical DFDs are implementation dependent."
StrQuestions(5) = "A data structure is a stet of data items that are related to one another."
StrQuestions(6) = "Sometimes structure chart may be derived from DFD using transaction analysis."
StrQuestions(7) = "Coupling between modules should be maximum for better maintainability of the software."
StrQuestions(8) = "A parallel run False necessity involves tow different terminals accessing a common database"
StrQuestions(9) = "The loss of data integrity means the data has become inconsistent"
StrQuestions(10) = "In a typical university system the entities students and courses represent the relationship one to many."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub
Private Sub TrueorFalse_Paper12()
StrQuestions(1) = "Feedback is essential for system modifications"
StrQuestions(2) = "Computers are a MUST for systems analysis"
StrQuestions(3) = "Data validation is not necessary for on line data input"
StrQuestions(4) = "Transposition errors are NOT shown up by check digits"
StrQuestions(5) = "LANs provide connectivity between homogenous systems only"
StrQuestions(6) = "Strategic Information Systems are used to monitor day to data operations of an organization"
StrQuestions(7) = "During the design phase of the SDLC, input design is governed by the volume and frequency of data."
StrQuestions(8) = "In a DFD, a Data store refers to source or destination of data."
StrQuestions(9) = "Flow in information in a good organization is always vertical."
StrQuestions(10) = "4GL programs always execute faster then Third Generation Languages."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub
Private Sub TrueorFalse_Paper13()
StrQuestions(1) = "Unit test is done to test a module comprising of a few programs."
StrQuestions(2) = "In on line applications the use of a check digit has not relevance since one can cross verify the input by accessing the master file."
StrQuestions(3) = "DFDs are drawn only as pat of the system design documentation and it has NO relevance at the stage of requirements analysis."
StrQuestions(4) = "Programs are written in Assembly level languages for applications where response time become very critical."
StrQuestions(5) = "In Decision tables, for each rule, thee can be more than one condition being true"
StrQuestions(6) = "Structured English must be used in the user manual."
StrQuestions(7) = "A database administrator is the one who designs the database for an application"
StrQuestions(8) = "Batch totals are not implemented while designing on line applications"
StrQuestions(9) = "Back up procedures are implemented only when the application is an on line application"
StrQuestions(10) = "Normalization technique is useful only for relational databases"

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub
Private Sub TrueorFalse_Paper14()
StrQuestions(1) = "The logical schema of a system that uses a database is identical to the entity relationship diagram of the system."
StrQuestions(2) = "A normalized set of relations eliminates numerical inaccuracies in arithmetical computations."
StrQuestions(3) = "An information system is a closed system."
StrQuestions(4) = "Output of echo phase of the SDLC must be verified and validated before starting with the next phase."
StrQuestions(5) = "A technically sound and feasible system may not be taken up for development"
StrQuestions(6) = "Each at flow in the DFD has a corresponding entry in the data dictionary."
StrQuestions(7) = "Physical design must precede the logical design of the sytem"
StrQuestions(8) = "Once a new system is implemented , NO subsequent evaluation of its performance in necessary."
StrQuestions(9) = "A data dictionary issued for spelling checks in word processor"
StrQuestions(10) = "Site preparation is usually done before system design"

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub
Private Sub TrueorFalse_Paper15()
StrQuestions(1) = "User interaction is required only during the system study and analysis phases of SDLC."
StrQuestions(2) = "Benchmarking is useful in selecting computer systems"
StrQuestions(3) = "Configuration should be calculated for Peak loads and NOT just for average loads"
StrQuestions(4) = "Vendor's track record is not an important factor in computer selection."
StrQuestions(5) = "Feasibility study should always be carried out before embarking on systems analysis and design"
StrQuestions(6) = "A Technically feasible solution recommended by a study would always prove operationally feasible."
StrQuestions(7) = "A cost benefit analysis calculates the Break even point"
StrQuestions(8) = "System, implementation should begin only after system acceptance"
StrQuestions(9) = "Parallel runs should be used for system acceptance."
StrQuestions(10) = "If various modules have been tested thoroughly, the need for total system testing is obviated"

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub
Private Sub TrueorFalse_Paper16()
StrQuestions(1) = "Decision trees and decision tables perform the same function"
StrQuestions(2) = "Entity relationship diagrams are used to design files"
StrQuestions(3) = "A very skilled programmer would deliver the best results during the systems study phase"
StrQuestions(4) = "Starting from scratch and redesigning a compute system is the best strategy for over system deficiencies."
StrQuestions(5) = "System in effectiveness can result due to reasons other than program bugs."
StrQuestions(6) = "Indexed files are faster than sequential files in all situations."
StrQuestions(7) = "Indexed files can have more than new index key."
StrQuestions(8) = "Relative files cannot be created on tape drives."
StrQuestions(9) = "Streamer tape is NOT as good for back up as Floppy disk of a few data files."
StrQuestions(10) = "Program specification should be prepared after the system is accepted and implemented"

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub
Private Sub TrueorFalse_Paper17()
StrQuestions(1) = "The system review phase is carried out periodically even after the system is successfully implemented"
StrQuestions(2) = "Key to address hashing always results in a n unique address in direct files"
StrQuestions(3) = "System requirements keep changing over the system life cycle and need to be incorporated periodically through well-planned methods"
StrQuestions(4) = "Information needs depend on the responsibilities that managers are required to discharge"
StrQuestions(5) = "Information needs are not dependent up to the strategic goals and targets of the organisation"
StrQuestions(6) = "An efficient and well managed organization need not have a well designed MIS."
StrQuestions(7) = "Synonyms, resulting out of hashing speed up processing of relative files."
StrQuestions(8) = "It would be most desirable to cut over after feasibility study since all problems re identified."
StrQuestions(9) = "Since all uses belong to the organisation and familiar with its working training after system acceptance can be done away with"
StrQuestions(10) = "A well designed and tested system does not typically require maintenance"

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
Private Sub TrueorFalse_Paper18()
StrQuestions(1) = "Systems  overview can be prescribed by DFDs"
StrQuestions(2) = "DFDs incorporate decision trees."
StrQuestions(3) = "Grid charts and decision tables used n the system life cycle are similar"
StrQuestions(4) = "Cost comparison is not pertinent to the management evaluation of proposed computer hardware; they MUST accept the deniability report"
StrQuestions(5) = "Data security is not a major consideration when renting a computer from a reputed vendor"
StrQuestions(6) = "The computer's ability to perform complex tasks under adverse conditions is termed reliability"
StrQuestions(7) = "Data dictionaries can be used or detecting errors."
StrQuestions(8) = "A failure is the occurrence of software error; whereas an error occurs whenever the system does not produce the expected results."
StrQuestions(9) = "An outside consultant employed by an organization in the feasibility study MUST condors with all decision reached."
StrQuestions(10) = "Alpha testing is the verification process a system goes trough for the first time."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub multiplechoice_Paper1()
strmultiple(1) = "The approach to system output by using the 'principle of management by perception' is considered to be -"
strmuloption(1) = "backward looking"
strmuloption(2) = "forward looking"
strmuloption(3) = "historical approach"
strmuloption(4) = "none of the above"

strmultiple(2) = "The major focus of system methods and procedures is to related input data to final output that is meaningful to the -"
strmuloption(5) = "users"
strmuloption(6) = "programmers"
strmuloption(7) = "designers"
strmuloption(8) = "none of the above"

strmultiple(3) = "The components that make up any system is known as system -"
strmuloption(9) = "boundary"
strmuloption(10) = "environment"
strmuloption(11) = "description"
strmuloption(12) = "none of the above"

strmultiple(4) = "The members of the Executive Steering Committee are, generally managers at -"
strmuloption(13) = "operational level"
strmuloption(14) = "middle level"
strmuloption(15) = "top level"
strmuloption(16) = "combination of all"

strmultiple(5) = "ERP is used in talking of -"
strmuloption(17) = "operating systems"
strmuloption(18) = "security passes"
strmuloption(19) = "information environment"
strmuloption(20) = "none of these"

strmultiple(6) = "A management information system is designed to give information about the -"
strmuloption(21) = "past"
strmuloption(22) = " present"
strmuloption(23) = "future"
strmuloption(24) = "all above"

strmultiple(7) = "Which of he following is a Tactical decision?"
strmuloption(25) = "Diversification"
strmuloption(26) = "Workshop location"
strmuloption(27) = "Budget allocation"
strmuloption(28) = "None of the above"

strmultiple(8) = "Which of the following distinguish the level of Management decision-making?"
strmuloption(29) = "Time frames"
strmuloption(30) = "Type"
strmuloption(31) = "Frequency"
strmuloption(32) = "All of above"

strmultiple(9) = "Select the most appropriate statement of the following"
strmuloption(33) = "Operational managers make unstructured decisions"
strmuloption(34) = "Middle level managers undertake long range planning"
strmuloption(35) = "As the management level goes up the hierarchy , Information becomes more and more summarized"
strmuloption(36) = "None of these"

strmultiple(10) = "Management is linked to information by -"
strmuloption(37) = "decisions"
strmuloption(38) = "data"
strmuloption(39) = "both a) and b)"
strmuloption(40) = "neither a) and b)"

strmulanswer(1) = "forward looking"
strmulanswer(2) = "users"
strmulanswer(3) = "boundary"
strmulanswer(4) = "top level"
strmulanswer(5) = "information environment"
strmulanswer(6) = "future"
strmulanswer(7) = "None of the above"
strmulanswer(8) = "All of above"
strmulanswer(9) = "As the management level goes up the hierarchy , Information becomes more and more summarized"
strmulanswer(10) = "neither a) and b)"

End Sub

Private Sub multiplechoice_Paper2()
strmultiple(1) = "An external turnaround document is an output document -"
strmuloption(1) = "that is printed at an external agency"
strmuloption(2) = "all or part of which enters as input to the system"
strmuloption(3) = "that gets circulated between several external agencies."
strmuloption(4) = "None of the above"

strmultiple(2) = "A major principle of modularization is:"
strmuloption(5) = "The cohesion of each module should be low and coupling between modules should be strong"
strmuloption(6) = "The number of modules should be as low as possible"
strmuloption(7) = "The number of modules should be as high as possible"
strmuloption(8) = "Each module should have a high degree of cohesion whereas the coupling between should be low"

strmultiple(3) = "When here are not enough transactions with the main computer system the remote job entry is done under:"
strmuloption(9) = "on line mode"
strmuloption(10) = " Real time mode"
strmuloption(11) = "off line mode"
strmuloption(12) = "None of the above"

strmultiple(4) = "Schedule of delivery is an example of a "
strmuloption(13) = "Data processing system"
strmuloption(14) = "Transaction based system"
strmuloption(15) = "Decision support system"
strmuloption(16) = "None of the above"

strmultiple(5) = "The trend in data processing systems is to"
strmuloption(17) = "Distributed processing"
strmuloption(18) = "remote processing"
strmuloption(19) = "Real time processing"
strmuloption(20) = "None of the above"

strmultiple(6) = "A typical data processing context where master files are updated to produce desired output is known as "
strmuloption(21) = "Validation checking"
strmuloption(22) = "Transaction processing"
strmuloption(23) = "Normalization"
strmuloption(24) = "None of the above"

strmultiple(7) = "Which of the following are tools of SASD?"
strmuloption(25) = "HIPO"
strmuloption(26) = "CASE"
strmuloption(27) = "DFD"
strmuloption(28) = "None of the above"

strmultiple(8) = "Which of the following is part of the SDLC?"
strmuloption(29) = "Requirement analysis"
strmuloption(30) = "Program specification"
strmuloption(31) = "Bench marking"
strmuloption(32) = "None of the above"

strmultiple(9) = "Which of the following is not part of the SDLC?"
strmuloption(33) = "Audit"
strmuloption(34) = "Reliability"
strmuloption(35) = "Security"
strmuloption(36) = "None of the above"

strmultiple(10) = "Reviews and walkthroughs are carried out at the stage of :"
strmuloption(37) = "Requirement analysis"
strmuloption(38) = "System design"
strmuloption(39) = "Coding"
strmuloption(40) = "All of the above"

strmulanswer(1) = "all or part of which enters as input to the system"
strmulanswer(2) = "Each module should have a high degree of cohesion whereas the coupling between should be low"
strmulanswer(3) = "off line mode"
strmulanswer(4) = "Decision support system"
strmulanswer(5) = "Distributed processing"
strmulanswer(6) = "Transaction processing"
strmulanswer(7) = "None of the above"
strmulanswer(8) = "None of the above"
strmulanswer(9) = "None of the above"
strmulanswer(10) = "All of the above"

End Sub

Private Sub multiplechoice_Paper3()
strmultiple(1) = "If the requirements analysis phase of a software development project is not done properly then the -"
strmuloption(1) = "resulting system would be delivered before time"
strmuloption(2) = "output reports would be indecipherable"
strmuloption(3) = "system might fail to address the real needs of users."
strmuloption(4) = "All of the above"

strmultiple(2) = "The detailed study/investigation of the present system is frequently referred to as -"
strmuloption(5) = "System planning"
strmuloption(6) = "System analysis"
strmuloption(7) = "Feasibility study"
strmuloption(8) = "None of the above"

strmultiple(3) = "Tangible benefits by their very nature require-"
strmuloption(9) = "Subjective evaluation"
strmuloption(10) = "Quantifiable evaluation"
strmuloption(11) = "Feasible evaluation"
strmuloption(12) = "None of the above"

strmultiple(4) = "Whether a proposed system can provide right information for the organization’s personnel, falls under the study of -"
strmuloption(13) = "Economic feasibility"
strmuloption(14) = "Operational feasibility"
strmuloption(15) = "Technical feasibility"
strmuloption(16) = "All of the above"

strmultiple(5) = "In considering the total cost associated with desired information, cost of data collection data input and computer processing costs are -"
strmuloption(17) = "one time"
strmuloption(18) = "temporary"
strmuloption(19) = "permanent"
strmuloption(20) = "recurring"

strmultiple(6) = "The next major step before system design and after feasibility stud is -"
strmuloption(21) = "Analysis activity"
strmuloption(22) = "Equipment selection activity"
strmuloption(23) = "Implementation activity"
strmuloption(24) = "None of the above"

strmultiple(7) = "If a data dictionary is not included in the systems analysis and design of a software project, then the following tasks cannot be carried out well:"
strmuloption(25) = "over all cost benefit analysis of the project"
strmuloption(26) = "Cross referencing and consistency checks of various data"
strmuloption(27) = "Bench marking of equipment to be purchased"
strmuloption(28) = "None of the above"

strmultiple(8) = "A major advantage of systems developed using DBMS as compared to systems using conventional file system is the following :"
strmuloption(29) = "The DBMS involves simultaneous updates of all data structures to keep the status current"
strmuloption(30) = "Conventional file system applications are inflexible"
strmuloption(31) = "Ad - hoc queries are not possible in conventional file system applications."
strmuloption(32) = "All of the above"

strmultiple(9) = "'Removal of functional dependency from a relation' falls under the category of -"
strmuloption(33) = "1st form of Normalization"
strmuloption(34) = "2nd form of Normalization"
strmuloption(35) = "3rd form of Normalization"
strmuloption(36) = "None of the above"

strmultiple(10) = "Reason for normalization is to-"
strmuloption(37) = "structure data"
strmuloption(38) = "simplify maintenance . E.g. updates"
strmuloption(39) = "simplify retrieval . E.g. query"
strmuloption(40) = "All of the above"

strmulanswer(1) = "system might fail to address the real needs of users."
strmulanswer(2) = "System analysis"
strmulanswer(3) = "Quantifiable evaluation"
strmulanswer(4) = "Operational feasibility"
strmulanswer(5) = "recurring"
strmulanswer(6) = "Analysis activity"
strmulanswer(7) = "Cross referencing and consistency checks of various data"
strmulanswer(8) = "All of the above"
strmulanswer(9) = "2nd form of Normalization"
strmulanswer(10) = "All of the above"

End Sub

Private Sub multiplechoice_Paper4()
strmultiple(1) = "Why is Data Flow Diagram referred to during the design of input/ because -"
strmuloption(1) = "there is a one to one relation between the input items and processes"
strmuloption(2) = "a DFD shows the storage methods for data files"
strmuloption(3) = "input items are some of the data flows on the data flow diagram"
strmuloption(4) = "All of the above"

strmultiple(2) = "To ensure system quality -"
strmuloption(5) = "unless user needs and software requirements specifications are reviewed system design should not be initiated"
strmuloption(6) = "inspection should be carried out a t pre specified milestones"
strmuloption(7) = "a proper test plan should be prepared and followed"
strmuloption(8) = "All of the above"

strmultiple(3) = "Stub is met with tin the context of -"
strmuloption(9) = "Data communication"
strmuloption(10) = "Testing of modules"
strmuloption(11) = "Random access"
strmuloption(12) = "None of the above"

strmultiple(4) = "Which of the following is not used in end user training before implementation?"
strmuloption(13) = "CAI"
strmuloption(14) = "Multimedia"
strmuloption(15) = "User manual"
strmuloption(16) = "None of the above"

strmultiple(5) = "In the implementation phase of SASD, following are included:"
strmuloption(17) = "Parallel run"
strmuloption(18) = "Sizing"
strmuloption(19) = "Specification freezes"
strmuloption(20) = "All of the above"

strmultiple(6) = "To produce a training manual for your organization, you would be"
strmuloption(21) = "Ventura"
strmuloption(22) = "DTP"
strmuloption(23) = "PageMaker"
strmuloption(24) = "All of the above"

strmultiple(7) = "Which of the following cannot be used for data entry?"
strmuloption(25) = "Com"
strmuloption(26) = "CDROM"
strmuloption(27) = "MICR"
strmuloption(28) = "None of these"

strmultiple(8) = "Spoken messages may be stored and forwarded by -"
strmuloption(29) = "Voicemail system"
strmuloption(30) = "Videoconferencing"
strmuloption(31) = "E-mail"
strmuloption(32) = "Expert system"

strmultiple(9) = "Chaining & linking are frequently used techniques in respect of storage of records in -"
strmuloption(33) = "tables"
strmuloption(34) = "files"
strmuloption(35) = "access"
strmuloption(36) = "None of the above"

strmultiple(10) = "When are many interfacing a convenient way of user interaction?"
strmuloption(37) = "for data processing in a restaurant"
strmuloption(38) = "When mouse is used as the main input device"
strmuloption(39) = "When data has to be imported from a spreadsheet"
strmuloption(40) = "When it is difficult  for the user to remember all the options available and for typing in complicated commands"

strmulanswer(1) = "there is a one to one relation between the input items and processes"
strmulanswer(2) = "All of the above"
strmulanswer(3) = "Testing of modules"
strmulanswer(4) = "None of the above"
strmulanswer(5) = "All of the above"
strmulanswer(6) = "All of the above"
strmulanswer(7) = "Com"
strmulanswer(8) = "Voicemail system"
strmulanswer(9) = "files"
strmulanswer(10) = "When it is difficult  for the user to remember all the options available and for typing in complicated commands"

End Sub

Private Sub multiplechoice_Paper5()
strmultiple(1) = "The main advantage of normalized relations in relational DBMS is that they -"
strmuloption(1) = "is highly secure"
strmuloption(2) = "do not suffer from anomalies during delete and update operations"
strmuloption(3) = "occupy minimal storage"
strmuloption(4) = "All of the above"

strmultiple(2) = "Which one of the following statements is incorrect?"
strmuloption(5) = "A report format can show vertical lines in the report"
strmuloption(6) = "A report format file can show any object within a box."
strmuloption(7) = "Current data can be shown in page header as well as page footer."
strmuloption(8) = "None of the above"

strmultiple(3) = "A portfolio management program is -"
strmuloption(9) = "never included in other financial management package"
strmuloption(10) = "another name for a personal financial management package"
strmuloption(11) = "used to track the fluctuating value of investment holdings"
strmuloption(12) = "normally used to track personal budget items and write cheques to pay bills"

strmultiple(4) = "Statement of changes in financial position or statement of sources and issue funds is called -"
strmuloption(13) = "Cash flow statement"
strmuloption(14) = "funds flow statement"
strmuloption(15) = "Journal"
strmuloption(16) = "None of the above"

strmultiple(5) = "The funds flow statement does not present the -"
strmuloption(17) = "estimation of working capital requirement"
strmuloption(18) = "uses of working capital"
strmuloption(19) = "sources of working capital"
strmuloption(20) = "net change in working capital"

strmultiple(6) = "A company to tide over a temporary cash flow problem would undertake the following -"
strmuloption(21) = "increase credit limits"
strmuloption(22) = "increase credit period"
strmuloption(23) = "make prompt creditors payments"
strmuloption(24) = "none of the above"

strmultiple(7) = "A business accounting system includes -"
strmuloption(25) = "Order entry"
strmuloption(26) = "Sales analysis"
strmuloption(27) = "Shipping/ Billing"
strmuloption(28) = "All of the above"

strmultiple(8) = "The transactions relating to all temporary and permanent accounts is maintained in the -"
strmuloption(29) = "Cash journal"
strmuloption(30) = "General ledger"
strmuloption(31) = "Accounts payable component"
strmuloption(32) = "All of the above"

strmultiple(9) = "Statement of changes in financial position or statement of sources and issue funds is called -"
strmuloption(33) = "Cash flow statement"
strmuloption(34) = "Funds flow statement"
strmuloption(35) = "Journal"
strmuloption(36) = "None of the above"

strmultiple(10) = "Books of account for a proprietary concern - "
strmuloption(37) = "may have subsidiary ledgers"
strmuloption(38) = "may not have subsidiary ledgers"
strmuloption(39) = "Both a) and b)"
strmuloption(40) = "neither a) not b)."

strmulanswer(1) = "do not suffer from anomalies during delete and update operations"
strmulanswer(2) = "None of the above"
strmulanswer(3) = "another name for a personal financial management package"
strmulanswer(4) = "Journal"
strmulanswer(5) = "estimation of working capital requirement"
strmulanswer(6) = "none of the above"
strmulanswer(7) = "All of the above"
strmulanswer(8) = "General ledger"
strmulanswer(9) = "Funds flow statement"
strmulanswer(10) = "Both a) and b)"
End Sub

Private Sub multiplechoice_Paper6()
strmultiple(1) = "A voucher when posted in the books -"
strmuloption(1) = "may not affect the total for all debits"
strmuloption(2) = "may affect the total of all debits"
strmuloption(3) = "neither a) nor b)"
strmuloption(4) = "Both a) and b)"

strmultiple(2) = "Reserves and surpluses-"
strmuloption(5) = "constitute the liabilities side of the balance sheet"
strmuloption(6) = "constitute the assets side of he balance sheet"
strmuloption(7) = "do not affect the balance sheet"
strmuloption(8) = "All the above"

strmultiple(3) = "Schedules to Annual Accounts -"
strmuloption(9) = "have to be prepared to comply with Statuary requirements of the law"
strmuloption(10) = "show detailed information regarding the working of the company"
strmuloption(11) = "Both the above"
strmuloption(12) = "Neither of the above"

strmultiple(4) = "The Balance Sheet of a company reflects -"
strmuloption(13) = "its status at any given point of time"
strmuloption(14) = "its performance during the course of the year"
strmuloption(15) = "Neither a) nor b)"
strmuloption(16) = "Both a) and b)"

strmultiple(5) = "Whenever we withdraw funds from the bank to meet our day-to-day expenses we credit -"
strmuloption(17) = "the bank account and debit contra A/C"
strmuloption(18) = "contra account and debit cast book"
strmuloption(19) = "Neither a) and b)"
strmuloption(20) = "Both a) and b)"

strmultiple(6) = "A Journal voucher -"
strmuloption(21) = "does not affect the bank or cash book balance"
strmuloption(22) = "affects the bank and cash book balance"
strmuloption(23) = "Both a) and b)"
strmuloption(24) = "Neither a) nor b)"

strmultiple(7) = "The balances of contra account -"
strmuloption(25) = "need not always be zero"
strmuloption(26) = "should always be zero."
strmuloption(27) = "Neither a) nor b)"
strmuloption(28) = "Both a) and b)"

strmultiple(8) = "A journal voucher is -"
strmuloption(29) = "also passed to incorporate corrections/amendments to the books of account"
strmuloption(30) = "not meant to incorporate corrections/amendments to the books of account"
strmuloption(31) = "Both a) and b)"
strmuloption(32) = "None of the above"

strmultiple(9) = "After each year we -"
strmuloption(33) = "need not start with the P & L statement afresh"
strmuloption(34) = "MUST start with the P & L statement afresh"
strmuloption(35) = "Both are correct"
strmuloption(36) = "None of the above"

strmultiple(10) = "After every year we merge the P & L of operations with the -"
strmuloption(37) = "cash book"
strmuloption(38) = "balance sheet"
strmuloption(39) = "bank a/c"
strmuloption(40) = "All of the above"

strmulanswer(1) = "Both a) and b)"
strmulanswer(2) = "constitute the liabilities side of the balance sheet"
strmulanswer(3) = "Both the above"
strmulanswer(4) = "its status at any given point of time"
strmulanswer(5) = "Both a) and b)"
strmulanswer(6) = "does not affect the bank or cash book balance"
strmulanswer(7) = "should always be zero."
strmulanswer(8) = "also passed to incorporate corrections/amendments to the books of account"
strmulanswer(9) = "MUST start with the P & L statement afresh"
strmulanswer(10) = "balance sheet"

End Sub

Private Sub multiplechoice_Paper7()
strmultiple(1) = "The procurement process ends to an outside vendor with the preparation of -"
strmuloption(1) = "Customer order"
strmuloption(2) = "Sales order"
strmuloption(3) = "Purchase order"
strmuloption(4) = "None of above"

strmultiple(2) = "In a supermarket which of the following, will be found in use in its computerized system?"
strmuloption(5) = "Scanner"
strmuloption(6) = "Bar coder"
strmuloption(7) = "POS terminal"
strmuloption(8) = "All of the above"

strmultiple(3) = "In a bus Passenger reservation system which of the following would be the primary key?"
strmuloption(9) = "Route"
strmuloption(10) = "Driver"
strmuloption(11) = "The bus"
strmuloption(12) = "None of the above"

strmultiple(4) = "A new drug is being tested on Guinea pigs and a computer keeps tag on amount and dates. Which of the following would the computer is least competent to do?"
strmuloption(13) = "Keep record of each guinea pig, the data and drug dosage"
strmuloption(14) = "Determine the total dosage individually and collectively"
strmuloption(15) = "Decide whether the drug is safe for use on humans"
strmuloption(16) = "Record weight changes of each guinea pig"

strmultiple(5) = "BOM -"
strmuloption(17) = "is used in Finance systems"
strmuloption(18) = "lists all the items required in fabricating an item"
strmuloption(19) = "is used for sales analysis and performance"
strmuloption(20) = "None of the above"

strmultiple(6) = "WIP -"
strmuloption(21) = "is a term used in the wire rope industry"
strmuloption(22) = "represents issuance of stock from the stores"
strmuloption(23) = "is used in production planning"
strmuloption(24) = "None of the above"

strmultiple(7) = "CAM is used in which industry?"
strmuloption(25) = "Automobile"
strmuloption(26) = "Helicopter"
strmuloption(27) = "Three wheeler"
strmuloption(28) = "All the above"

strmultiple(8) = "Bill of Materials File would typically be -"
strmuloption(29) = "master file"
strmuloption(30) = "transactions file"
strmuloption(31) = "both a master as well as a transactions file"
strmuloption(32) = "none of the above"

strmultiple(9) = "How does the check digit method capture mistakes that occur while keying in a code, which incorporates a check digit?"
strmuloption(33) = "The method checks the number of digits keyed in with the maximum permissible number of digits in a code"
strmuloption(34) = " The method applies a hashing algorithm to the code entered and if the result is not a single digit then it includes the code entered in erroneous"
strmuloption(35) = "The checking is done completely through hardware"
strmuloption(36) = "Using the method adopted for generating the check digit while the code was originally designed the method computes the check digit again for the non-check digit part of the entered key. If the check digit computed thus is not equal to the check digit entered, then it indicates and error condition"

strmultiple(10) = "Back up refers to"
strmuloption(37) = "delays in production of outputs"
strmuloption(38) = "accumulation of programs in memory"
strmuloption(39) = "job scheduling failures"
strmuloption(40) = "None of the above"

strmulanswer(1) = "None of above"
strmulanswer(2) = "All of the above"
strmulanswer(3) = "None of the above"
strmulanswer(4) = "Decide whether the drug is safe for use on humans"
strmulanswer(5) = "lists all the items required in fabricating an item"
strmulanswer(6) = "None of the above"
strmulanswer(7) = "All the above"
strmulanswer(8) = "master file"
strmulanswer(9) = "Using the method adopted for generating the check digit while the code was originally designed the method computes the check digit again for the non-check digit part of the entered key. If the check digit computed thus is not equal to the check digit entered, then it indicates and error condition"
strmulanswer(10) = "None of the above"
End Sub

Private Sub multiplechoice_Paper8()
strmultiple(1) = "Cost of error correction is least a which stage?"
strmuloption(1) = "Requirement analysis"
strmuloption(2) = "Design"
strmuloption(3) = "Development"
strmuloption(4) = "Implementation"

strmultiple(2) = "Digital signature is used in -"
strmuloption(5) = "Floppy disk marking"
strmuloption(6) = "Programming"
strmuloption(7) = "Data Communication"
strmuloption(8) = "None of the above"

strmultiple(3) = "The four icons used in building DFDs are -"
strmuloption(9) = "Flow, process source/destination, store"
strmuloption(10) = "Flow, source, store, process"
strmuloption(11) = "Flow, process, source, store"
strmuloption(12) = "Source, process, destination , store"

strmultiple(4) = "Which of the following will endure accuracy of inputs to a computer?"
strmuloption(13) = "Scanner"
strmuloption(14) = "OCR"
strmuloption(15) = "Check digits"
strmuloption(16) = "All of them"

strmultiple(5) = "The type of code which involves the assignment of a consecutive and unique number to each item is -"
strmuloption(17) = "sequence"
strmuloption(18) = "group"
strmuloption(19) = "numeric"
strmuloption(20) = "None of the above"

strmultiple(6) = "A Structured decision -"
strmuloption(21) = "is made after completion of a preplanned series of steps."
strmuloption(22) = "depends on information not known in advance"
strmuloption(23) = "is not routine"
strmuloption(24) = "None of these"

strmultiple(7) = "Managers performing Strategic and Technical role need -"
strmuloption(25) = "Internal Information ONLY"
strmuloption(26) = "External Information ONLY"
strmuloption(27) = "Both , Internal and External Information"
strmuloption(28) = "None of these"

strmultiple(8) = "Top Level managers are"
strmuloption(29) = "Strategic information"
strmuloption(30) = "Tactical information"
strmuloption(31) = "Operational information"
strmuloption(32) = "None of these"

strmultiple(9) = "Corporate Information"
strmuloption(33) = "brings healthy relationship between employees and employers"
strmuloption(34) = "contains personal particulars"
strmuloption(35) = "means Information derived from weekly meetings"
strmuloption(36) = "None of these"

strmultiple(10) = "Entities, Attributes and Relationships are associated with -"
strmuloption(37) = "logical concepts of Data"
strmuloption(38) = "physical concepts of data"
strmuloption(39) = "persons of an Organization"
strmuloption(40) = "None of these"

strmulanswer(1) = "Requirement analysis"
strmulanswer(2) = "Data Communication"
strmulanswer(3) = "Flow, process source/destination, store"
strmulanswer(4) = "All of them"
strmulanswer(5) = "numeric"
strmulanswer(6) = "is made after completion of a preplanned series of steps."
strmulanswer(7) = "Both , Internal and External Information"
strmulanswer(8) = "Strategic information"
strmulanswer(9) = "None of these"
strmulanswer(10) = "logical concepts of Data"
End Sub

Private Sub multiplechoice_Paper9()
strmultiple(1) = "Transaction processing is called On line if transaction -"
strmuloption(1) = "is processed as soon as it enters the System"
strmuloption(2) = "is processed in discrete lots"
strmuloption(3) = "is processed in a running system"
strmuloption(4) = "None of these"

strmultiple(2) = "A Distributed data processing system -"
strmuloption(5) = "attempts to capture advantage of bot centralized and decentralized processing"
strmuloption(6) = "does NOT allow greater flexibilities"
strmuloption(7) = "provides slow access to data"
strmuloption(8) = "None of these"

strmultiple(3) = "In Data Processing, Classifying, Calculating, Sorting and Summarizing Data are connected with -"
strmuloption(9) = "capturing of input data"
strmuloption(10) = " manipulating of data"
strmuloption(11) = "managing output results"
strmuloption(12) = "None of these"

strmultiple(4) = "Prototyping means -"
strmuloption(13) = "creating, developing and refining a working model of the final operational system"
strmuloption(14) = "testing the Computer system"
strmuloption(15) = "designing the computer system"
strmuloption(16) = "None of these"

strmultiple(5) = "Storage devices may be -"
strmuloption(17) = "sequential type ONLY"
strmuloption(18) = "direct access type ONLY"
strmuloption(19) = "either a or b"
strmuloption(20) = "None of these"

strmultiple(6) = "Blocking factor is the -"
strmuloption(21) = "number of records grouped in a block"
strmuloption(22) = "number of blocks in a file"
strmuloption(23) = "ratio between numbers of records accessed and total number of records"
strmuloption(24) = "None of these"

strmultiple(7) = "The first step in the SDLC is -"
strmuloption(25) = "preliminary investigation and analysis"
strmuloption(26) = "system design"
strmuloption(27) = "database design"
strmuloption(28) = "None of these"

strmultiple(8) = "An aid to System Design should primarily -"
strmuloption(29) = "help in documentation"
strmuloption(30) = "generate code"
strmuloption(31) = "help analyze both data and activities"
strmuloption(32) = "None of these"

strmultiple(9) = "Peer Review is an aid to -"
strmuloption(33) = "programming"
strmuloption(34) = "quality assurance"
strmuloption(35) = "feasibility"
strmuloption(36) = "None of these"

strmultiple(10) = "HIPO charts facilitate -"
strmuloption(37) = "analysis of modules"
strmuloption(38) = "top down design"
strmuloption(39) = "organizational structure"
strmuloption(40) = "None of these"

strmulanswer(1) = "is processed as soon as it enters the System"
strmulanswer(2) = "attempts to capture advantage of bot centralized and decentralized processing"
strmulanswer(3) = "capturing of input data"
strmulanswer(4) = "creating, developing and refining a working model of the final operational system"
strmulanswer(5) = "either a or b"
strmulanswer(6) = "ratio between numbers of records accessed and total number of records"
strmulanswer(7) = "preliminary investigation and analysis"
strmulanswer(8) = "help in documentation"
strmulanswer(9) = "quality assurance"
strmulanswer(10) = "top down design"
End Sub

Private Sub multiplechoice_Paper10()
strmultiple(1) = "Decision tables link conditions and -"
strmuloption(1) = "tables"
strmuloption(2) = "programs"
strmuloption(3) = "actions"
strmuloption(4) = "None of these"

strmultiple(2) = "Prototyping aims at -"
strmuloption(5) = "end user understanding and approval"
strmuloption(6) = "program logic"
strmuloption(7) = "planning of data flow organization"
strmuloption(8) = "None of these"

strmultiple(3) = "In a Passenger Reservation System which of the following is the most critical?"
strmuloption(9) = "Ease of programming"
strmuloption(10) = "Response time"
strmuloption(11) = "GUI"
strmuloption(12) = "None of these"

strmultiple(4) = "A software design description document only includes -"
strmuloption(13) = "Data dictionary"
strmuloption(14) = "ER diagram"
strmuloption(15) = "DFD"
strmuloption(16) = "a, b and c"

strmultiple(5) = "Pseudo code is -"
strmuloption(17) = "false logic"
strmuloption(18) = "programming aid"
strmuloption(19) = "both a and b"
strmuloption(20) = "neither a nor b"

strmultiple(6) = "A data dictionary -"
strmuloption(21) = "gives the meaning of the data names used in the system "
strmuloption(22) = "defines the data types"
strmuloption(23) = "defines all data elements and structures used in DFD"
strmuloption(24) = "none of the above"

strmultiple(7) = "Backup and recovery procedures are necessary to -"
strmuloption(25) = "reorganize the disk"
strmuloption(26) = "control the DBA"
strmuloption(27) = "handle contingencies like files getting corrupt or becoming irretrievable"
strmuloption(28) = "none of these"

strmultiple(8) = "Which of the following is an Exception report?"
strmuloption(29) = "value of stocks in hand"
strmuloption(30) = "monthly salary sheet"
strmuloption(31) = "machine replacement report"
strmuloption(32) = "none of these"

strmultiple(9) = "File conversion is part of -"
strmuloption(33) = "system cut over"
strmuloption(34) = "system design"
strmuloption(35) = "day-to-day activity to DP"
strmuloption(36) = "none of these"

strmultiple(10) = "Of the following, which is not a function of RDBMS?"
strmuloption(37) = "Audit trail"
strmuloption(38) = "Security"
strmuloption(39) = "File and data integrity"
strmuloption(40) = "none of these"

strmulanswer(1) = "actions"
strmulanswer(2) = "end user understanding and approval"
strmulanswer(3) = "Response time"
strmulanswer(4) = "a, b and c"
strmulanswer(5) = "programming aid"
strmulanswer(6) = "none of the above"
strmulanswer(7) = "handle contingencies like files getting corrupt or becoming irretrievable"
strmulanswer(8) = "machine replacement report"
strmulanswer(9) = "system cut over"
strmulanswer(10) = "none of these"
End Sub

Private Sub multiplechoice_Paper11()
strmultiple(1) = "An important function of top management is to -"
strmuloption(1) = "negotiate with employee unions"
strmuloption(2) = "train employees"
strmuloption(3) = "do spell check"
strmuloption(4) = "define organizational objectives"

strmultiple(2) = "A model is used to analyze a complete system because it -"
strmuloption(5) = "is usually computer based"
strmuloption(6) = "is closer to reality than the original system"
strmuloption(7) = "contains only the essential features needed and hence easy to manipulate"
strmuloption(8) = "is aesthetically appealing"

strmultiple(3) = "Structured analysis and design uses -"
strmuloption(9) = "documentation produced on word processors"
strmuloption(10) = "trained programmers only in all phases"
strmuloption(11) = "prototypes generated using object oriented methods"
strmuloption(12) = "diagrams like DFDs"

strmultiple(4) = "A DBMS is used mainly to -"
strmuloption(13) = "separate management from work"
strmuloption(14) = "promotes decentralized and end user computing"
strmuloption(15) = "facilitate transmission of files across continents"
strmuloption(16) = "eliminate data redundancy"

strmultiple(5) = "Modular design of a system means -"
strmuloption(17) = "deliver a complex system to he customer one piece at a time"
strmuloption(18) = "divide the whole system into a number of units each of which is quite cohesive within itself and is not too dependent on others and to design and develop the units separately"
strmuloption(19) = "using models at each phase of systems life cycle"
strmuloption(20) = "using sophisticated subroutine libraries available in shareware"

strmultiple(6) = "DSS role is -"
strmuloption(21) = "trial and error search for solutions"
strmuloption(22) = "planning"
strmuloption(23) = "analyzing alternatives"
strmuloption(24) = "a, b and c above"

strmultiple(7) = "Design specifications do not normally include -"
strmuloption(25) = "Output requirements"
strmuloption(26) = "input and storage requirements"
strmuloption(27) = "control provisions"
strmuloption(28) = "Blueprints showing the layout of hardware"

strmultiple(8) = "The sequence of steps following in a system study is :"
strmuloption(29) = "problem definition, system design, system analysis, Programming and Implementation"
strmuloption(30) = "Problem definition, System Analysis, Programming and Implementation"
strmuloption(31) = "Systems analysis, System  Design, and Systems Implementations"
strmuloption(32) = "Problem Definitions, Systems Analysis, System Design, Programming and Implementation"

strmultiple(9) = "Which of the following is not a tool of data collection?"
strmuloption(33) = "on site observations"
strmuloption(34) = "program flowcharts"
strmuloption(35) = "interviews"
strmuloption(36) = "Questionnaires"

strmultiple(10) = "which of the following tools is not used during systems analysis?"
strmuloption(37) = "data flow diagram"
strmuloption(38) = "Structured English"
strmuloption(39) = "Structure chart"
strmuloption(40) = "Decision table"

strmulanswer(1) = "define organizational objectives"
strmulanswer(2) = "contains only the essential features needed and hence easy to manipulate"
strmulanswer(3) = "diagrams like DFDs"
strmulanswer(4) = "eliminate data redundancy"
strmulanswer(5) = "divide the whole system into a number of units each of which is quite cohesive within itself and is not too dependent on others and to design and develop the units separately"
strmulanswer(6) = "a, b and c above"
strmulanswer(7) = "Blueprints showing the layout of hardware"
strmulanswer(8) = "Problem Definitions, Systems Analysis, System Design, Programming and Implementation"
strmulanswer(9) = "program flowcharts"
strmulanswer(10) = "Structured English"
End Sub

Private Sub multiplechoice_Paper12()
strmultiple(1) = "Data dictionary keeps details of the contents of -"
strmuloption(1) = "data flows"
strmuloption(2) = "data stores"
strmuloption(3) = "both the above"
strmuloption(4) = "neither of the two"

strmultiple(2) = "System quality relates to its -"
strmuloption(5) = "reliability"
strmuloption(6) = "efficiency"
strmuloption(7) = "maintainability"
strmuloption(8) = "all the three above"

strmultiple(3) = "A distinguishing feature of Strategic Information Systems is -"
strmuloption(9) = "end user computing"
strmuloption(10) = "transaction processing"
strmuloption(11) = "information for Top management"
strmuloption(12) = "none of the above"

strmultiple(4) = "Most important aspects, of systems design, focus on-"
strmuloption(13) = "economic & technical feasibility"
strmuloption(14) = "operational feasibility"
strmuloption(15) = "information for Top management"
strmuloption(16) = "All of the above"

strmultiple(5) = "The most important reason for failure of MIS is :"
strmuloption(17) = "use of improper tools for design "
strmuloption(18) = "non involvement of end users"
strmuloption(19) = "improper specification"
strmuloption(20) = "none of the above"

strmultiple(6) = "To implement a MIS -"
strmuloption(21) = "a Computer is inescapable"
strmuloption(22) = "E mail is required"
strmuloption(23) = "proper systems and procedures must be in place"
strmuloption(24) = "none of the above"

strmultiple(7) = "An exception report pertaining to an edit run lists -"
strmuloption(25) = "all input entries whether they contain errors or not "
strmuloption(26) = "correct entries only"
strmuloption(27) = "inputs with errors only"
strmuloption(28) = "none of these"

strmultiple(8) = "ABC analysis in Inventory Control is performed by comparing -"
strmuloption(29) = "items used against items unused. infrequently used"
strmuloption(30) = "value of items consumed with their quantities"
strmuloption(31) = "value of items held in stock with their quantities"
strmuloption(32) = "value of items held in stock with the number of items"

strmultiple(9) = "System specification is mainly concerned with -"
strmuloption(33) = "how the functions are performed"
strmuloption(34) = "what functions are to be performed"
strmuloption(35) = "both a & b"
strmuloption(36) = "none of these"

strmultiple(10) = "In the analysis and presentation of logic, which of the following techniques will ensure that all the combinations of conditions have been considered?"
strmuloption(37) = "HIPO chart"
strmuloption(38) = "Decision table"
strmuloption(39) = "Pseudocode"
strmuloption(40) = "DFD"

strmulanswer(1) = "both the above"
strmulanswer(2) = "all the three above"
strmulanswer(3) = "information for Top management"
strmulanswer(4) = "All of the above"
strmulanswer(5) = "non involvement of end users"
strmulanswer(6) = "proper systems and procedures must be in place"
strmulanswer(7) = "inputs with errors only"
strmulanswer(8) = "value of items consumed with their quantities"
strmulanswer(9) = "both a & b"
strmulanswer(10) = "Decision table"
End Sub

Private Sub multiplechoice_Paper13()
strmultiple(1) = "The physical design of a new system must consider all of the following except -"
strmuloption(1) = "input, output , processing"
strmuloption(2) = "management objectives"
strmuloption(3) = "control and back up"
strmuloption(4) = "conversion methods"

strmultiple(2) = "A structure chart-"
strmuloption(5) = "shows module inter relationships in a hierarchical manner"
strmuloption(6) = "describes the internal structure of a program in a graphical manner"
strmuloption(7) = "is a graphical representation of structured English"
strmuloption(8) = "depicts data structures in the form of a chart"

strmultiple(3) = "To determine, whether all the data in a batch run have been processed, we should use a -"
strmuloption(9) = "reasonable check"
strmuloption(10) = "limit check"
strmuloption(11) = "record count"
strmuloption(12) = "consistency check."

strmultiple(4) = "A computer center is expressing its inability to make changes to an existing system,. This is most probably due to -"
strmuloption(13) = "non availability of user manual"
strmuloption(14) = "inability to trace the programmer who developed the system"
strmuloption(15) = "non availability of source code"
strmuloption(16) = "all these"

strmultiple(5) = "The phase of system development associated with creation of test data is -"
strmuloption(17) = "Systems analysis"
strmuloption(18) = "Physical design"
strmuloption(19) = "systems acceptance"
strmuloption(20) = "logical design"

strmultiple(6) = "The provident fund(PF) contribution rates for employees were hard coded (made pat of the program) in a payroll system. Whenever rate changes occurred, it meant program modification and recompilation. Which of the following is most appropriate?"
strmuloption(21) = "it would have been better to design PF rates to be input from the console;"
strmuloption(22) = "it would have been more efficient to design PF rates as Batch inputs"
strmuloption(23) = "Hard coding has helped in minimizing input errors and info aster execution time"
strmuloption(24) = "data that changes rarely can be hard coded"

strmultiple(7) = "A prototype -"
strmuloption(25) = "is a working system which will be changed through several iterations"
strmuloption(26) = "helps to crystallize essential design requirements"
strmuloption(27) = "will provide preliminary information on the workability of the concept"
strmuloption(28) = "all these three"

strmultiple(8) = "A parallel run involves -"
strmuloption(29) = "firing two different applications from different terminals accessing a common database"
strmuloption(30) = "compiling programs with two different languages"
strmuloption(31) = "the concurrent operation the existing and newly developed systems"
strmuloption(32) = "none of these"

strmultiple(9) = "A data dictionary is -"
strmuloption(33) = "a comprehensive integration collection of data about data"
strmuloption(34) = "a dictionary or algorithm for obtaining the address of logical records on a storage device"
strmuloption(35) = "the relationship among the different fields of data in a secondary storage"
strmuloption(36) = "none of these"

strmultiple(10) = "A menu driven software-"
strmuloption(37) = "helps the user reduce errors in data entry"
strmuloption(38) = "software developed for planning meals at a restaurant"
strmuloption(39) = "contains options for the user to indicate his or her choice "
strmuloption(40) = "all these three "

strmulanswer(1) = "conversion methods"
strmulanswer(2) = "shows module inter relationships in a hierarchical manner"
strmulanswer(3) = "record count"
strmulanswer(4) = "non availability of source code"
strmulanswer(5) = "Not AVailable"
strmulanswer(6) = "data that changes rarely can be hard coded"
strmulanswer(7) = "all these three"
strmulanswer(8) = "the concurrent operation the existing and newly developed systems"
strmulanswer(9) = "a comprehensive integration collection of data about data"
strmulanswer(10) = "contains options for the user to indicate his or her choice "
End Sub

Private Sub multiplechoice_Paper14()
strmultiple(1) = "All the following are examples of physical security except "
strmuloption(1) = "limiting access to the computer complex"
strmuloption(2) = "using passwords to access a system"
strmuloption(3) = "CCTV"
strmuloption(4) = "storing tapes and disks not in use in a separate building"

strmultiple(2) = "Loss of Data integrity implies that data is -"
strmuloption(5) = "not suitable for running in an integrated environment"
strmuloption(6) = "inconsistent"
strmuloption(7) = "repeated"
strmuloption(8) = "outdated"

strmultiple(3) = "In a DBMS two record types and their relationship are called -"
strmuloption(9) = "schema"
strmuloption(10) = " segment"
strmuloption(11) = "set"
strmuloption(12) = "database record"

strmultiple(4) = "On line data entry is preferable to batch data entry because -"
strmuloption(13) = "it ensures better accuracy of input"
strmuloption(14) = "it is faster"
strmuloption(15) = "it reduces human efforts in checking control totals, checklist etc."
strmuloption(16) = "All three above"

strmultiple(5) = "In a hospital database system, the data relationship between the entity [Patient] and the entity [ward] would be -"
strmuloption(17) = "one to one"
strmuloption(18) = "one to many"
strmuloption(19) = "many to one "
strmuloption(20) = "many to many"

strmultiple(6) = "An Automatic Teller Machine(ATM) represents -"
strmuloption(21) = "an on line banking system"
strmuloption(22) = "a dedicated PC based system"
strmuloption(23) = "a combination of batch cum on line system"
strmuloption(24) = "none of the above"

strmultiple(7) = "Which of the following fact finding techniques is most useful in collecting quantitative data?"
strmuloption(25) = "Interviews"
strmuloption(26) = "old records"
strmuloption(27) = "questionnaires"
strmuloption(28) = "All these"

strmultiple(8) = "The primary function of a LAN is -"
strmuloption(29) = "permitting access to certain files by every user"
strmuloption(30) = "extending electronic massaging facility"
strmuloption(31) = "optimizing resource sharing"
strmuloption(32) = "all three above"

strmultiple(9) = "Structured English is -"
strmuloption(33) = "grammatically correct description of the system"
strmuloption(34) = "a tool for the systems analyst to document the processing procedure using action oriented verbs and terms defined in the data dictionary"
strmuloption(35) = "a COBOL statement in the PROCEDURE DIVISION"
strmuloption(36) = "none of the above"

strmulanswer(1) = "using passwords to access a system"
strmulanswer(2) = "inconsistent"
strmulanswer(3) = "schema"
strmulanswer(4) = "it is faster"
strmulanswer(5) = "many to many"
strmulanswer(6) = "an on line banking system"
strmulanswer(7) = "questionnaires"
strmulanswer(8) = "all three above"
strmulanswer(9) = "a tool for the systems analyst to document the processing procedure using action oriented verbs and terms defined in the data dictionary"

End Sub

Private Sub multiplechoice_Paper15()
strmultiple(1) = "Steering committee-"
strmuloption(1) = "sanctions budget to the EDP department"
strmuloption(2) = "is a forum for interaction between users and EDP personnel"
strmuloption(3) = "guides the EDP department on what to do"
strmuloption(4) = "none of these"

strmultiple(2) = "Structured systems Development Cycle approach includes the following characteristics:"
strmuloption(5) = "Top down"
strmuloption(6) = "Data flow"
strmuloption(7) = "Data structure"
strmuloption(8) = "None of the above"

strmultiple(3) = "Which of the following is NOT part of MIS?"
strmuloption(9) = "exception report for middle management"
strmuloption(10) = "summary report for top management"
strmuloption(11) = "action report for the line management"
strmuloption(12) = "payroll for workers"

strmultiple(4) = "Data integrity stands for -"
strmuloption(13) = "thoroughly validated data"
strmuloption(14) = "data stored in different files showing consistency"
strmuloption(15) = "data fields integrated to provide summary"
strmuloption(16) = "none of the above"

strmultiple(5) = "Which of the following is not a data security measure?"
strmuloption(17) = "using password for access"
strmuloption(18) = "using retention periods for all files"
strmuloption(19) = "allowing only EDP personnel to access any file"
strmuloption(20) = "none of these"

strmultiple(6) = "Which one of the following is not a Top management function?"
strmuloption(21) = "planning"
strmuloption(22) = "organizing"
strmuloption(23) = "decision making"
strmuloption(24) = "Day to day operations"

strmultiple(7) = "The feasibility report of an existing system highlights -"
strmuloption(25) = "dataflow and processing"
strmuloption(26) = "information needs"
strmuloption(27) = "problem areas"
strmuloption(28) = "all three above"

strmultiple(8) = "Prototype is a -"
strmuloption(29) = "minimize of the existing system"
strmuloption(30) = "minimodel of the proposed system"
strmuloption(31) = "working model of the existing system"
strmuloption(32) = "none of the above"

strmultiple(9) = "Which is the most critical phase of the SCLS?"
strmuloption(33) = "feasibility study"
strmuloption(34) = "systems analysis"
strmuloption(35) = "systems design"
strmuloption(36) = "all these"

strmultiple(10) = "With regard to DFDs which of the following is incorrect?"
strmuloption(37) = "show the sequence of steps"
strmuloption(38) = "are different from flowcharts"
strmuloption(39) = "are system analyst's tools"
strmuloption(40) = "none of these"

strmulanswer(1) = "guides the EDP department on what to do"
strmulanswer(2) = "None of the above"
strmulanswer(3) = "payroll for workers"
strmulanswer(4) = "none of the above"
strmulanswer(5) = "none of these"
strmulanswer(6) = "Day to day operations"
strmulanswer(7) = "problem areas"
strmulanswer(8) = "minimodel of the proposed system"
strmulanswer(9) = "all these"
strmulanswer(10) = "none of these"

End Sub

Private Sub multiplechoice_Paper16()
strmultiple(1) = "Which of the following is NOT an element in the physical DFD?"
strmuloption(1) = "Internal / External entity"
strmuloption(2) = "Data flows"
strmuloption(3) = "Processes"
strmuloption(4) = "Flowchart"

strmultiple(2) = "Which of the following defines the rules of DFD?"
strmuloption(5) = "All data flows must begin and / or end at a process"
strmuloption(6) = "All data flow must contain data"
strmuloption(7) = "If two or more documents travel together , each has to be represented as a separate data flows"
strmuloption(8) = "Only processes can connect to Data Stores"

strmultiple(3) = "The traditional systems flow chart is connected with -"
strmuloption(9) = "flow of information"
strmuloption(10) = "showing, graphically, operations, files equipments etc with the system"
strmuloption(11) = "introducing the various systems in the organization"
strmuloption(12) = "none of the above"

strmultiple(4) = "A Data dictionary is a list containing details of -"
strmuloption(13) = "data elements input to the system"
strmuloption(14) = "data elements output by the system"
strmuloption(15) = "description of processes"
strmuloption(16) = "all these"

strmultiple(5) = "Which of the following statements attributed to Prototyping is NOT true"
strmuloption(17) = "always gives faster performance than the actual system"
strmuloption(18) = "encourages an active end user participation"
strmuloption(19) = "it is an active model that end users an see feel and experience"
strmuloption(20) = "can accelerate several processes of systems design"

strmultiple(6) = "Hashing procedure is used in -"
strmuloption(21) = "Random files"
strmuloption(22) = "sequential files"
strmuloption(23) = "indexed sequential files"
strmuloption(24) = "none of these"

strmultiple(7) = "A DDL- "
strmuloption(25) = "establishes relationships, fields and record types"
strmuloption(26) = "helps maintaining data in the database"
strmuloption(27) = "creates data base"
strmuloption(28) = "none of these"

strmultiple(8) = "A Post implementation review looks into -"
strmuloption(29) = "goals achieved vis a vis targets"
strmuloption(30) = "user friendliness of the systems"
strmuloption(31) = "realization of cost benefits"
strmuloption(32) = "all these"

strmultiple(9) = "To achieve good response time from OTLP, critical requirements are -"
strmuloption(33) = "faster disk transfer"
strmuloption(34) = "efficient ID controller"
strmuloption(35) = "efficient programming"
strmuloption(36) = "all these"

strmultiple(10) = "Which of the following is not an input specification for a Batch system?"
strmuloption(37) = "source document layout"
strmuloption(38) = "prototype reports"
strmuloption(39) = "Expanded data dictionary"
strmuloption(40) = "input record layout"

strmulanswer(1) = "Internal / External entity"
strmulanswer(2) = "If two or more documents travel together , each has to be represented as a separate data flows"
strmulanswer(3) = "introducing the various systems in the organization"
strmulanswer(4) = "all these"
strmulanswer(5) = "always gives faster performance than the actual system"
strmulanswer(6) = "Random files"
strmulanswer(7) = "establishes relationships, fields and record types"
strmulanswer(8) = "all these"
strmulanswer(9) = "all these"
strmulanswer(10) = "Expanded data dictionary"

End Sub

Private Sub multiplechoice_Paper17()
strmultiple(1) = "Which of the following is NOT part of SDLC ?"
strmuloption(1) = "feasibility study"
strmuloption(2) = "design specs"
strmuloption(3) = "unit testing"
strmuloption(4) = "none of these"

strmultiple(2) = "On line data entry is most suitable in the case of -"
strmuloption(5) = "entering monthly journal entries"
strmuloption(6) = "payroll master updating"
strmuloption(7) = "processing payment of cheque in a bank"
strmuloption(8) = "none of these"

strmultiple(3) = "Mistakes, made in he requirements determination stage show up in -"
strmuloption(9) = "system design"
strmuloption(10) = "implementation"
strmuloption(11) = "system development and testing"
strmuloption(12) = "none of these"

strmultiple(4) = "While designing a report layout which, of the following should NOT be done?"
strmuloption(13) = "report should be read left to right and top to bottom"
strmuloption(14) = "the most important items should be the earliest to find "
strmuloption(15) = "use abbreviations"
strmuloption(16) = "none of these"

strmultiple(5) = "In development of an application system, which accesses data under a DBMS, the user views the data base as a -"
strmuloption(17) = "group of files"
strmuloption(18) = "logical structure"
strmuloption(19) = "random storage structure"
strmuloption(20) = "none of these"

strmultiple(6) = "A data dictionary -"
strmuloption(21) = "is a list of all data elements included in a set of DFDs describing a system"
strmuloption(22) = "is a dictionary containing details of data"
strmuloption(23) = "talks about meanings of termiology used in databases"
strmuloption(24) = "is none of these"

strmultiple(7) = "Of the following which one does not relate to logical design of the system ?"
strmuloption(25) = "input specifications"
strmuloption(26) = "output specs"
strmuloption(27) = "file maintenance specs"
strmuloption(28) = "none of these"

strmultiple(8) = "System prototyping helps the designer in -"
strmuloption(29) = "making the programmers understand how the system will be"
strmuloption(30) = "giving a demo of the software to the Systems Manager he or she reports to"
strmuloption(31) = "communicating to the user, quickly how the system"
strmuloption(32) = "none of these"

strmultiple(9) = "A Zero level DFD describes- "
strmuloption(33) = "overview of processes, inputs and outputs"
strmuloption(34) = "the fully blown up system design"
strmuloption(35) = "that the System Design cannot be split further"
strmuloption(36) = "none of these"

strmultiple(10) = "A Transaction processing application is one -"
strmuloption(37) = "where data transactions can be processed"
strmuloption(38) = "developed for those that transact business"
strmuloption(39) = "driven by the transaction entered by the user"
strmuloption(40) = "none of these"

strmulanswer(1) = "none of these"
strmulanswer(2) = "processing payment of cheque in a bank"
strmulanswer(3) = "system development and testing"
strmulanswer(4) = "none of these"
strmulanswer(5) = "random storage structure"
strmulanswer(6) = "is a dictionary containing details of data"
strmulanswer(7) = "file maintenance specs"
strmulanswer(8) = "communicating to the user, quickly how the system"
strmulanswer(9) = "overview of processes, inputs and outputs"
strmulanswer(10) = "driven by the transaction entered by the user"

End Sub

Private Sub multiplechoice_Paper18()
strmultiple(1) = "The basic objective of systems analysis is to -"
strmuloption(1) = "understand computer hardware by opening the system unit"
strmuloption(2) = "Tran managers in mathematical analysis"
strmuloption(3) = "run simulating programs"
strmuloption(4) = "understands a complex system and modifies it in some way"

strmultiple(2) = "One of the main differences between the responsibilities of a system analyst and a programmer is -"
strmuloption(5) = "programmers use lower level languages whereas systems analysts use higher-level languages"
strmuloption(6) = "programmers deal with computer programs whereas a systems analyst deals with many more things such as identifying opportunities for computerization, designing  new procedures selection of equipment etc."
strmuloption(7) = "systems analyst is concerned with accounting function"
strmuloption(8) = "programmer's responsibility is to train System Analysts in PC usage"

strmultiple(3) = "Systems supports means -"
strmuloption(9) = "furniture specially designed for holding computer systems"
strmuloption(10) = "battery set that supports computes when the power is off"
strmuloption(11) = "a general fund maintained by each organization towards computer maintenance"
strmuloption(12) = "activity of making modifications and offering the help as and when needed after a system is made operational"

strmultiple(4) = "An exception report contains data of the following type -"
strmuloption(13) = "unfiltered list of all customer documents"
strmuloption(14) = "data, which violates some, specified condition, value or standard"
strmuloption(15) = "a list of all transactions already predicted to happen"
strmuloption(16) = "list of customer orders with order quantities greater than one"

strmultiple(5) = "An unstructured problem means a/an -"
strmuloption(17) = "problem in which the objectives, constraints or relationships between variables and solution procedures cannot be clearly specified"
strmuloption(18) = "problem not related to civil engineering"
strmuloption(19) = "imaginary problem"
strmuloption(20) = "problem with well specified objectives and well specified relationships between variables but does not have an efficient solution procedure"

strmultiple(6) = "The activity that does not form a phase of the SDLC is -"
strmuloption(21) = "designing the system"
strmuloption(22) = "study of curent system"
strmuloption(23) = "formulating a strategy for the organization of which the system under discussion is a part"
strmuloption(24) = "define user needs"

strmultiple(7) = "A Prototype means a/an -"
strmuloption(25) = "kind of printer font that is preferred by system analysts"
strmuloption(26) = "type of generic system from which several other systems can be generated"
strmuloption(27) = "working model of the real system"
strmuloption(28) = "automated version of a manual system"

strmultiple(8) = "A data model represents -"
strmuloption(29) = "all the types of data and their associations that are relevant to the system"
strmuloption(30) = "a sample set of data"
strmuloption(31) = "all facts related to the existing system"
strmuloption(32) = "all the processes of the system along with the sequence in which they are executed."

strmultiple(9) = "A data dictionary consists of -"
strmuloption(33) = "all transactions that have been entered in the system"
strmuloption(34) = "all words which can be referred to during the spellchecker by a word processor"
strmuloption(35) = "data about the files and their contents and about the processes used by the system"
strmuloption(36) = "an indexed sequential file containing the frequency of access of each data items in the system"

strmultiple(10) = "A decision table specifies -"
strmuloption(37) = "a list of decisions taken by the top management during the current planning horizon"
strmuloption(38) = "in the form of a table, the actions to be taken under different conditions for a given process"
strmuloption(39) = "a tree showing the causes and effects in a decision process"
strmuloption(40) = "the consequence of all possible decisions in tabular form"

strmulanswer(1) = "understands a complex system and modifies it in some way"
strmulanswer(2) = "programmers deal with computer programs whereas a systems analyst deals with many more things such as identifying opportunities for computerization, designing  new procedures selection of equipment etc."
strmulanswer(3) = "activity of making modifications and offering the help as and when needed after a system is made operational"
strmulanswer(4) = "data, which violates some, specified condition, value or standard"
strmulanswer(5) = "problem in which the objectives, constraints or relationships between variables and solution procedures cannot be clearly specified"
strmulanswer(6) = "formulating a strategy for the organization of which the system under discussion is a part"
strmulanswer(7) = "automated version of a manual system"
strmulanswer(8) = "all the types of data and their associations that are relevant to the system"
strmulanswer(9) = "data about the files and their contents and about the processes used by the system"
strmulanswer(10) = "in the form of a table, the actions to be taken under different conditions for a given process"
End Sub

Private Sub multiplechoice_Paper19()
strmultiple(1) = "Back up procedure helps in restoring -"
strmuloption(1) = "data files whenever there is a system crash"
strmuloption(2) = "both application and system software wherever there is a Disk corruption. failure"
strmuloption(3) = "both a and b"
strmuloption(4) = "none of these"

strmultiple(2) = "A DFD is -"
strmuloption(5) = "mainly used at the system specification stage"
strmuloption(6) = "the primary output of the systems design phase"
strmuloption(7) = "the modern version of a flow chart"
strmuloption(8) = "none of these"

strmultiple(3) = "Which of the following appropriately explains the desirable characteristic of a good systems design"
strmuloption(9) = "modular approach"
strmuloption(10) = "proper documentation"
strmuloption(11) = "neither a not b"
strmuloption(12) = "both a and b"

strmultiple(4) = "One of the following is not a part of a DFD -"
strmuloption(13) = "bubble (process)"
strmuloption(14) = "arrow (flow)"
strmuloption(15) = "disk storage"
strmuloption(16) = "store"

strmultiple(5) = "System specs are used to -  "
strmuloption(17) = "describe system flow"
strmuloption(18) = "avoid ambuity"
strmuloption(19) = "give accurate picture of system"
strmuloption(20) = "present all these"

strmultiple(6) = "Which of the following is not a tool for application prototyping?"
strmuloption(21) = "third generation languages"
strmuloption(22) = "report generators"
strmuloption(23) = "screen generators"
strmuloption(24) = "application generators"

strmultiple(7) = "Which of the following would not be a major deliverable of the structured systems analysis phase?"
strmuloption(25) = "DFD"
strmuloption(26) = "data dictionaries"
strmuloption(27) = "entity relationship diagrams"
strmuloption(28) = "prototype model"

strmultiple(8) = "Structured programming involves -"
strmuloption(29) = "functional modularization"
strmuloption(30) = "localization of errors"
strmuloption(31) = "decentralized programming"
strmuloption(32) = "all of these"

strmultiple(9) = "A System design is said to be functionally modular if -"
strmuloption(33) = "each module performs a specific function and can be developed relatively independently by programmers"
strmuloption(34) = "system is able to handle all functions of the application"
strmuloption(35) = "system is developed using structured programming through COBOL or PASCAL"
strmuloption(36) = "System makes extensive use of function keys for maximum user assistance"

strmultiple(10) = "Controls, in processing must be established n a system in order to -"
strmuloption(37) = "prohibit tampering with information by unauthorized persons"
strmuloption(38) = "verify that all data has bee processed"
strmuloption(39) = "block or trap faulty data from entry into processing"
strmuloption(40) = "all these three"

strmulanswer(1) = "both a and b"
strmulanswer(2) = "mainly used at the system specification stage"
strmulanswer(3) = "modular approach"
strmulanswer(4) = "disk storage"
strmulanswer(5) = "present all these"
strmulanswer(6) = "third generation languages"
strmulanswer(7) = "prototype model"
strmulanswer(8) = "functional modularization"
strmulanswer(9) = "each module performs a specific function and can be developed relatively independently by programmers"
strmulanswer(10) = "all these three"
End Sub

Private Sub multiplechoice_Paper20()
strmultiple(1) = "The term used to refer to the checking of outputs of a computer with corresponding input documents is called -"
strmuloption(1) = "auditing through the computer"
strmuloption(2) = "auditing around the computer"
strmuloption(3) = " process control"
strmuloption(4) = "beta test"

strmultiple(2) = "System implementation phase of SDLC entails -"
strmuloption(5) = "System checkouts"
strmuloption(6) = "pilot run"
strmuloption(7) = "parallel runs"
strmuloption(8) = "all the three"

strmultiple(3) = "Ideally, information system objectives should be stated in the form of -"
strmuloption(9) = "increased product sales"
strmuloption(10) = "reduced marketing cost"
strmuloption(11) = "increased sales of product X by Y% in the next quarter"
strmuloption(12) = "none of the above"

strmultiple(4) = "All of the following are considered usual steps in systems development except -"
strmuloption(13) = "purchase of equipment"
strmuloption(14) = "program development"
strmuloption(15) = "analysis of old system"
strmuloption(16) = "program testing"

strmultiple(5) = "The approach used in top down analysis and design is to -"
strmuloption(17) = "identify the top-level functions by combining many smaller components into a single entity"
strmuloption(18) = "identifies a top-level function and then creates a hierarchy of lower level modules and components"
strmuloption(19) = "prepare flowcharts after programming this been completed"
strmuloption(20) = "none of these"

strmultiple(6) = "Cost benefit analysis"
strmuloption(21) = "compares the costs, with the benefits of introducing a computer based system"
strmuloption(22) = "estimates hardware and software costs"
strmuloption(23) = "evaluates the tangible and intangible factors"
strmuloption(24) = "all these three"

strmultiple(7) = "Which criterion of a real time system makes it different form an interactive system?"
strmuloption(25) = "Data gathering from internal environment"
strmuloption(26) = "Response time limited by strict time constraints"
strmuloption(27) = "information transfer to external world"
strmuloption(28) = "none of these"

strmultiple(8) = "A self-checks digit is useful in detecting -"
strmuloption(29) = "transcription errors only"
strmuloption(30) = "transposition errors only"
strmuloption(31) = "random errors only"
strmuloption(32) = "all the three"

strmultiple(9) = "Which of the following is NOT considered a tool to be used in the system design phase?"
strmuloption(33) = "DFD"
strmuloption(34) = "Decision table"
strmuloption(35) = "Systems flowchart"
strmuloption(36) = "Pie chart"

strmultiple(10) = "A site license enables a software package to be used -"
strmuloption(37) = "only on one system"
strmuloption(38) = "over the enterprise spread over  all metropolitan cities"
strmuloption(39) = "at one locating up to a specified number of computers"
strmuloption(40) = "none of the above"

strmulanswer(1) = "auditing around the computer"
strmulanswer(2) = "all the three"
strmulanswer(3) = "increased sales of product X by Y% in the next quarter"
strmulanswer(4) = "purchase of equipment"
strmulanswer(5) = "identifies a top-level function and then creates a hierarchy of lower level modules and components"
strmulanswer(6) = "all these three"
strmulanswer(7) = "Response time limited by strict time constraints"
strmulanswer(8) = "all the three"
strmulanswer(9) = "Pie chart"
strmulanswer(10) = "at one locating up to a specified number of computers"
End Sub
