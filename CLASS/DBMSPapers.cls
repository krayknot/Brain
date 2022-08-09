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
 

Public Function DBMS(PaperType As String, DBMSPaper As Paper, mCat As Category, Number As Integer)
'ptype is the paper type
'DBMSpaper is the number of paper as paper1,paper2 ...
'mcat is the category
'number is the number of questions, or options or answers to be taken

StrPaperType = PaperType
' StrPaper = Paper
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
         
ElseIf PaperType = "Matching Columns" Then
       If DBMSPaper = Paper1 Then Matching_Columns_Paper1
       If DBMSPaper = Paper2 Then Matching_Columns_Paper2
       If DBMSPaper = Paper3 Then Matching_Columns_Paper3
       If DBMSPaper = Paper4 Then Matching_Columns_Paper4
       If DBMSPaper = Paper5 Then Matching_Columns_Paper5
       If DBMSPaper = Paper6 Then Matching_Columns_Paper6
       If DBMSPaper = Paper7 Then Matching_Columns_Paper7
       If DBMSPaper = Paper8 Then Matching_Columns_Paper8
       If DBMSPaper = Paper9 Then Matching_Columns_Paper9
       If DBMSPaper = Paper10 Then Matching_Columns_Paper10
              
ElseIf PaperType = "True Or False" Then
       If DBMSPaper = Paper1 Then TrueorFalse_Paper1
       If DBMSPaper = Paper2 Then TrueorFalse_Paper2
       If DBMSPaper = Paper3 Then TrueorFalse_Paper3
       If DBMSPaper = Paper4 Then TrueorFalse_Paper4
       If DBMSPaper = Paper5 Then TrueorFalse_Paper5
       If DBMSPaper = Paper6 Then TrueorFalse_Paper6
       If DBMSPaper = Paper7 Then TrueorFalse_Paper7
       If DBMSPaper = Paper8 Then TrueorFalse_Paper8
       If DBMSPaper = Paper9 Then TrueorFalse_Paper9
       If DBMSPaper = Paper10 Then TrueorFalse_Paper10
              
ElseIf PaperType = "Descriptive Questions" Then
       If DBMSPaper = Paper1 Then questions_paper1
       If DBMSPaper = Paper2 Then questions_paper2
       If DBMSPaper = Paper3 Then questions_paper3
       
ElseIf PaperType = "Multiple Choice" Then
       If DBMSPaper = Paper1 Then multiplechoice_Paper1
       If DBMSPaper = Paper2 Then multiplechoice_Paper2
       If DBMSPaper = Paper3 Then multiplechoice_Paper3
       If DBMSPaper = Paper4 Then multiplechoice_Paper4
       If DBMSPaper = Paper5 Then multiplechoice_Paper5
       If DBMSPaper = Paper6 Then multiplechoice_Paper6
       If DBMSPaper = Paper7 Then multiplechoice_Paper7
       If DBMSPaper = Paper8 Then multiplechoice_Paper8
       If DBMSPaper = Paper9 Then multiplechoice_Paper9
       If DBMSPaper = Paper10 Then multiplechoice_Paper10
End If

If mCat = Fill_Questions Then
   DBMS = StrQuestions(Number)
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
       
ElseIf mCat = Questions Then
       DBMS = StrQuestions(Number)
       
ElseIf mCat = TF_Questions Then
       DBMS = StrQuestions(Number)
ElseIf mCat = TF_Answers Then
       DBMS = stranswers(Number)
       
ElseIf mCat = MC_Questions Then
       DBMS = strmultiple(Number)
       
ElseIf mCat = MC_Options_1 Then
       DBMS = strmuloption(Number)
       
ElseIf mCat = MC_Options_2 Then
       DBMS = strmuloption(Number)
       
ElseIf mCat = MC_Options_3 Then
       DBMS = strmuloption(Number)
       
ElseIf mCat = MC_Options_4 Then
       DBMS = strmuloption(Number)
       
ElseIf mCat = MC_Answer Then
       DBMS = strmulanswer(Number)
End If

End Function

Private Sub Fill_in_the_Blanks_Paper1()
StrQuestions(1) = "SQL queries are based on dash operation."
StrQuestions(2) = "The dash of a given relation is the set of tuples appearing in that relation at any given instant."
StrQuestions(3) = "In the relational data model, the columns of a table are called dash."
StrQuestions(4) = "A multi level index leads to a dash."
StrQuestions(5) = "dash is a specification of the portion off data which is oriented to the needs of one or more application programs."
StrQuestions(6) = "SET dash TO is used to set up an open index file as the master index, without closing other index files."
StrQuestions(7) = "The primary key index does not allow dash data in a field."
StrQuestions(8) = "dash is updated before a transaction commits."
StrQuestions(9) = "In a lock based schedule dash can occur."
StrQuestions(10) = "Read lock of a data item by a transaction prevents other transactions to dash that data."

stroptions(1) = "Attributes"
stroptions(2) = "B-Tree"
stroptions(3) = "Selection"
stroptions(4) = "Extension"
stroptions(5) = "Subschema"
stroptions(6) = "Modify"
stroptions(7) = "Log"
stroptions(8) = "Order"
stroptions(9) = "Deadlock"
stroptions(10) = "Duplicate"

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
Private Sub Fill_in_the_Blanks_Paper2()
StrQuestions(1) = "The maximum height of a B+ tree of order m with n key values is dash"
StrQuestions(2) = "To create an index sequential file dash storage device is used."
StrQuestions(3) = "A function dependency is a relationship between dash"
StrQuestions(4) = "Recovery of data by first removing invalid changes and then reapplying valid ones is known as dash."
StrQuestions(5) = "A View of the data base that appears to an application program is known as dash."
StrQuestions(6) = "The language used for data modification is known as dash."
StrQuestions(7) = "dash is a collection of records having one to many relationships."
StrQuestions(8) = "dash of a relation X, produces a relation Y such that Y contains only selected attributes of X."
StrQuestions(9) = "In a relational database dash are not  allowed to have null values."
StrQuestions(10) = "Complex entities can be constructed using dash relationship."

stroptions(1) = "Attributes"
stroptions(2) = "Direct access"
stroptions(3) = "Subschema"
stroptions(4) = "Aggregation"
stroptions(5) = "DML"
stroptions(6) = "Projection"
stroptions(7) = "Logm/2 + (n/2) + 1"
stroptions(8) = "Keys"
stroptions(9) = "Network model"
stroptions(10) = "Undo – redo"

stranswers(1) = "Logm/2 + (n/2) + 1"
stranswers(2) = "Direct access"
stranswers(3) = "Attributes"
stranswers(4) = "Undo - redo"
stranswers(5) = "Subschema"
stranswers(6) = "DML"
stranswers(7) = "Network model"
stranswers(8) = "Projection"
stranswers(9) = "Keys"
stranswers(10) = "Aggregation"

End Sub

Private Sub Fill_in_the_Blanks_Paper3()
StrQuestions(1) = "By default FoxPro takes the dash key index."
StrQuestions(2) = "Graphics objects can be stored in the dash type of field."
StrQuestions(3) = "once marked for deletion, records are completely deleted from the table by the dash command."
StrQuestions(4) = "Views combine the qualities of tables and dash."
StrQuestions(5) = "dash relationships need not be recreated each time the tables are used."
StrQuestions(6) = "Visual FoxPro fieldnames can be at most dash characters long."
StrQuestions(7) = "Temporary relationships created with dash command control the relationship between record pointers in tables."
StrQuestions(8) = "dash organizes all the objects created into a project."
StrQuestions(9) = "The small black rectangle in the lower left corner of the Browse window is known as the dash dash."
StrQuestions(10) = "Multiple dash can be manipulated together in Visual FoxPro."

stroptions(1) = "Persistent"
stroptions(2) = "SET RELATION "
stroptions(3) = "Project manager"
stroptions(4) = "Regular"
stroptions(5) = "Project manager"
stroptions(6) = "General"
stroptions(7) = "PACK"
stroptions(8) = "Window splitter"
stroptions(9) = "Forms"
stroptions(10) = "Queries "

stranswers(1) = "Regular"
stranswers(2) = "General"
stranswers(3) = "PACK"
stranswers(4) = "Queries"
stranswers(5) = "Persistent"
stranswers(6) = "Project manager"
stranswers(7) = "SET RELATION "
stranswers(8) = "Project manager"
stranswers(9) = "Window splitter"
stranswers(10) = "Forms"

End Sub

Private Sub Fill_in_the_Blanks_Paper4()
StrQuestions(1) = "For each DO CASE statement, the program must use a correspondent dash command."
StrQuestions(2) = "SET ECHO ON opens the dash window."
StrQuestions(3) = "To disable the debug and trace windows, use the command dash."
StrQuestions(4) = "When an alternate file is defined with the SET ALTERNATE TO command, it is necessity to se the dash command to initiate storage of screen activity in the file"
StrQuestions(5) = "When data is entered with the @ …GET command, use the dash clause with @... GET for data validation."
StrQuestions(6) = "To remove the field headings from the output of the LIST command, use the FoxPro command dash."
StrQuestions(7) = "The universal quantifier in relational calculus can be implemented using dash SQL command."
StrQuestions(8) = "The list of tuples selected using the SQL SELECT FROM command may contain dash entries."
StrQuestions(9) = "SQL based RDBMS queries collets tuples having same value of attributes in the dash clause in one cluster."
StrQuestions(10) = "The embedded SQL SELECT command cannot reference relations into which tuples are being dash."

stroptions(1) = "VALID"
stroptions(2) = "SET DEBUG OFF"
stroptions(3) = "SET HEADING OFF"
stroptions(4) = "SET ALTERNATE ON"
stroptions(5) = "END CASE"
stroptions(6) = "Trace"
stroptions(7) = "NOT EXISTS"
stroptions(8) = "GROUP BY"
stroptions(9) = "Duplicate"
stroptions(10) = "Inserted"

stranswers(1) = "END CASE"
stranswers(2) = "Trace"
stranswers(3) = "SET DEBUG OFF"
stranswers(4) = "SET ALTERNATE ON"
stranswers(5) = "VALID"
stranswers(6) = "SET HEADING OFF"
stranswers(7) = "NOT EXISTS"
stranswers(8) = "Duplicate"
stranswers(9) = "GROUP BY"
stranswers(10) = "Inserted"

End Sub

Private Sub Fill_in_the_Blanks_Paper5()
StrQuestions(1) = "DBTG set type represents dash relationships among two record types."
StrQuestions(2) = "In CODASYL terminology, a member record can exist on its own without being a member of a set occurrence if its retention option is dash."
StrQuestions(3) = "Along with the GROUP BY clause, one can use the dash clause to define a condition on the group of tuples associated with each value of the grouping attributes."
StrQuestions(4) = "The primary key of a relation is a minimal dash."
StrQuestions(5) = "Insertion of data into a B-Tree may cause leaf and internal nodes to dash."
StrQuestions(6) = "In a relational database a referential integrity constraint is specified with the help of dash."
StrQuestions(7) = "The complementation rule holds for dash dependencies."
StrQuestions(8) = "A dash fragment of a relation is a subset of the tuples of that relation."
StrQuestions(9) = "Although the two phase commit protocol is resilient to site and link failures it may introduce dash"
StrQuestions(10) = "If a relation is in dash  it is also in 3NF."

stroptions(1) = "BCNF"
stroptions(2) = "Superkey"
stroptions(3) = "OPTIONAL"
stroptions(4) = "1 : N"
stroptions(5) = "HAVING"
stroptions(6) = "Split"
stroptions(7) = "Foreign key"
stroptions(8) = "Multivalued"
stroptions(9) = "Horizontal"
stroptions(10) = "Blocking"

stranswers(1) = "1 : N"
stranswers(2) = "OPTIONAL"
stranswers(3) = "HAVING"
stranswers(4) = "Superkey"
stranswers(5) = "Split"
stranswers(6) = "Foreign key"
stranswers(7) = "Multivalued"
stranswers(8) = "Horizontal"
stranswers(9) = "Blocking"
stranswers(10) = "BCNF"
End Sub

Private Sub Fill_in_the_Blanks_Paper6()
StrQuestions(1) = "The network address of the dash process, to which an Oracle client connects, using SQL *Net middleware is provided by the listener process."
StrQuestions(2) = "An Oracle server handles clients requests on a first - come first- serve basis, with the help of the shared dash server processes."
StrQuestions(3) = "Data can be loaded into an Oracle database from text files with the help of dash utility."
StrQuestions(4) = "A dash user is the owner off an Oracle database's dictionary tables."
StrQuestions(5) = "In order to start or stop an Oracle database one should connect as an dash user."
StrQuestions(6) = "The password of an Oracle user can be changed using dash command."
StrQuestions(7) = "The size and growth characteristics of Oracle tables can be specified using the dash clause with CREATE TABLE command."
StrQuestions(8) = "Oracle system allocates space for an object in clusters of data blocks called dash."
StrQuestions(9) = "When a users transaction changes data in the database. Oracle maingains 'undo' information in a dash segment."
StrQuestions(10) = "In Oracle, the transaction log for each database requires at least dash log groups to efficiently handle changes made by ongoing  transactions."

stroptions(1) = "Rollback"
stroptions(2) = "SYS"
stroptions(3) = "INTERNAL"
stroptions(4) = "Front-end"
stroptions(5) = "SQL*Loader"
stroptions(6) = "Dispatcher"
stroptions(7) = "ALTER USER"
stroptions(8) = "STORAGE"
stroptions(9) = "Extensts"
stroptions(10) = "Two"

stranswers(1) = "Dispatcher"
stranswers(2) = "Front-end"
stranswers(3) = "SQL*Loader"
stranswers(4) = "SYS"
stranswers(5) = "INTERNAL"
stranswers(6) = "ALTER USER"
stranswers(7) = "STORAGE"
stranswers(8) = "Extensts"
stranswers(9) = "Rollback"
stranswers(10) = "Two"
End Sub

Private Sub Fill_in_the_Blanks_Paper7()
StrQuestions(1) = "Oracle permits uses to update through a view when the view is created with dash option."
StrQuestions(2) = "The information about the columns involved in the integrity constaints, defined for an Oracle database, can be obtained by executing an SQL select command on the dash data dictionary view."
StrQuestions(3) = "The information about he object privileges that have been granted, are maintained by Oracle in the dash data dictionary view."
StrQuestions(4) = "The ALTER ROLE command is used to modify the dash of a role."
StrQuestions(5) = "The insert intensive OLTP applications can use the dash command to generate unique primary key for the new row."
StrQuestions(6) = "In an Oracle database server, the system global area(SGA) contains l9brary, dash and buffer caches."
StrQuestions(7) = "An Oracle server backs up transaction log groups, when that become full using the dash process."
StrQuestions(8) = "To add data files to a table space, with the help of SQL * DBA, the user must have dash TABLESPACE privilege.  "
StrQuestions(9) = "The processing loads, of different users of an Oracle system, can be monitored using the system I/O monitor of dash."
StrQuestions(10) = "The command ROLLBACK TO SAVEPOINT is used to dash rollback a transaction."

stroptions(1) = "CREATE SEQUENCE"
stroptions(2) = "Dictionary"
stroptions(3) = "CHECK"
stroptions(4) = "USER_CONS_COLUMNS"
stroptions(5) = "Password"
stroptions(6) = "TAB_PRIVS_MADE"
stroptions(7) = "Archiver"
stroptions(8) = "ALTER"
stroptions(9) = "SQL *DBA"
stroptions(10) = "Partially"

stranswers(1) = "CHECK"
stranswers(2) = "USER_CONS_COLUMNS"
stranswers(3) = "TAB_PRIVS_MADE"
stranswers(4) = "Password"
stranswers(5) = "CREATE SEQUENCE"
stranswers(6) = "Dictionary"
stranswers(7) = "Archiver"
stranswers(8) = "ALTER"
stranswers(9) = "SQL *DBA"
stranswers(10) = "Partially"

End Sub

Private Sub Fill_in_the_Blanks_Paper8()
StrQuestions(1) = "A dash is a stored procedure that Oracle automatically fires under appropriate conditions."
StrQuestions(2) = "The binary image frames, of size 1024 × 1024 pixels (8 bit), can be stored in an Oracle database as a data of type dash."
StrQuestions(3) = "when one column of a table refers to the values in another column of the same table, it introduces a dash integrity constraint."
StrQuestions(4) = "In order to improve concurrency , Oracle supports dash locking."
StrQuestions(5) = "When a transaction aborts, Oracle reads dash segment assigned to the transaction and restores the affected rows to their previous state."
StrQuestions(6) = "A user, with the dash system privilege can modify the database by adding new files to it."
StrQuestions(7) = "Oracle treats a dash as a collection of related privileges that a database administrator can grant collectively to the users."
StrQuestions(8) = "The transaction log in Oracle is also called dash log."
StrQuestions(9) = "In an Oracle server , the library ad dictionary caches make up the dash."
StrQuestions(10) = "In Oracle, execution of SQL statements can be controlled using a suitable dash enclosing it within a comment."

stroptions(1) = "Self referential"
stroptions(2) = "Role"
stroptions(3) = "Rollback"
stroptions(4) = "ALTER DATABASE"
stroptions(5) = "LONG RAW"
stroptions(6) = "Row –level"
stroptions(7) = "Trigger"
stroptions(8) = "Redo"
stroptions(9) = "Shared-pool"
stroptions(10) = "Hint"

stranswers(1) = "Trigger"
stranswers(2) = "LONG RAW"
stranswers(3) = "Self referential"
stranswers(4) = "Row -level"
stranswers(5) = "Rollback"
stranswers(6) = "ALTER DATABASE"
stranswers(7) = "Role"
stranswers(8) = "Redo"
stranswers(9) = "Shared-pool"
stranswers(10) = "Hint"

End Sub

Private Sub Fill_in_the_Blanks_Paper9()
StrQuestions(1) = "GRANT is a  dash command."
StrQuestions(2) = "In Oracle, trigger body cannot include dash statement."
StrQuestions(3) = "The body of a trigger in Oracle can contain dash statement."
StrQuestions(4) = "Oracle stores procedures, functions and packages in the dash tablespace."
StrQuestions(5) = "A table space can contain tuples of dash relations created using CREATE TABLE commands."
StrQuestions(6) = "In an Oracle PL/SQL program which includes a query that returns multiple rows from the database, the query should be declared using dash command."
StrQuestions(7) = "An Oracle database contains a publicly accessible dictionary table called DUAL which can be used to return the result of a SQL dash."
StrQuestions(8) = "In an Oracle server the database writer(DBWR) is a dash process."
StrQuestions(9) = "In an Oracle server, the background processes communicate among themselves using dash."
StrQuestions(10) = "The CHECK clause in CREATE TABLE command can be used to enforce dash."

stroptions(1) = "ALTER"
stroptions(2) = "UPDATE"
stroptions(3) = "Domain dependenct"
stroptions(4) = "System"
stroptions(5) = "Different"
stroptions(6) = "CURSOR"
stroptions(7) = "Client processes"
stroptions(8) = "Function"
stroptions(9) = "DCL"
stroptions(10) = "Background"

stranswers(1) = "DCL"
stranswers(2) = "ALTER"
stranswers(3) = "UPDATE"
stranswers(4) = "System"
stranswers(5) = "Different"
stranswers(6) = "CURSOR"
stranswers(7) = "Function"
stranswers(8) = "Background"
stranswers(9) = "Client processes"
stranswers(10) = "Domain dependenct"

End Sub
Private Sub Matching_Columns_Paper1()
strcolumnA(1) = "Key"
strcolumnA(2) = "Centralized resource shared by users"
strcolumnA(3) = "Location mode"
strcolumnA(4) = "LUPDATE"
strcolumnA(5) = "Join"
strcolumnA(6) = "Concurrent processing"
strcolumnA(7) = "Physical model"
strcolumnA(8) = "Authentication"
strcolumnA(9) = "Distributed databases"
strcolumnA(10) = "View"

strcolumnB(1) = "Data security"
strcolumnB(2) = "Minimal set of attributes that uniquely identifies a row"
strcolumnB(3) = "External Model"
strcolumnB(4) = "Storage Organization "
strcolumnB(5) = "Location transparency"
strcolumnB(6) = "Network model"
strcolumnB(7) = "Last update date of active database"
strcolumnB(8) = "Deadlock"
strcolumnB(9) = "Database"
strcolumnB(10) = "Combines tables"

stranswers(1) = "Minimal set of attributes that uniquely identifies a row"
stranswers(2) = "Database"
stranswers(3) = "Network model"
stranswers(4) = "Last update date of active database"
stranswers(5) = "Combines tables"
stranswers(6) = "Deadlock"
stranswers(7) = "Storage Organization "
stranswers(8) = "Data security"
stranswers(9) = "Location transparency"
stranswers(10) = "External Model"

End Sub
Private Sub Matching_Columns_Paper2()
strcolumnA(1) = "Entity integrity"
strcolumnA(2) = "Referential constraints"
strcolumnA(3) = ".FMT"
strcolumnA(4) = "SELECT"
strcolumnA(5) = "Networks "
strcolumnA(6) = "Rows"
strcolumnA(7) = "Triggers"
strcolumnA(8) = "Virtual tables"
strcolumnA(9) = "Storage Structure"
strcolumnA(10) = "Two-Phase commit"

strcolumnB(1) = "Format files"
strcolumnB(2) = "Foreign keys"
strcolumnB(3) = "Primary key values"
strcolumnB(4) = "B+- Trees"
strcolumnB(5) = "SQL"
strcolumnB(6) = "Concurrency control"
strcolumnB(7) = "Integrity checks"
strcolumnB(8) = "Tuples"
strcolumnB(9) = "User views"
strcolumnB(10) = "Sets"

stranswers(1) = "Primary key values"
stranswers(2) = "Foreign keys"
stranswers(3) = "Format files"
stranswers(4) = "SQL"
stranswers(5) = "Sets"
stranswers(6) = "Tuples"
stranswers(7) = "Integrity checks"
stranswers(8) = "User views"
stranswers(9) = "B+- Trees"
stranswers(10) = "Concurrency control"
End Sub
Private Sub Matching_Columns_Paper3()
strcolumnA(1) = "QBE"
strcolumnA(2) = "Owner "
strcolumnA(3) = "Null"
strcolumnA(4) = "Join"
strcolumnA(5) = "Hashing"
strcolumnA(6) = "Lock"
strcolumnA(7) = "Transaction"
strcolumnA(8) = "Recovery"
strcolumnA(9) = "Master index"
strcolumnA(10) = "Set date ANSI"

strcolumnB(1) = "Exclusive"
strcolumnB(2) = "ISAM"
strcolumnB(3) = "Synonyms"
strcolumnB(4) = "Shadow paging"
strcolumnB(5) = "Attribute"
strcolumnB(6) = "Set"
strcolumnB(7) = "YY.MM.DD"
strcolumnB(8) = "Domain calculus"
strcolumnB(9) = "Commit"
strcolumnB(10) = "Dangling tuples"

stranswers(1) = "Domain calculus"
stranswers(2) = "Set"
stranswers(3) = "Attribute"
stranswers(4) = "Dangling tuples"
stranswers(5) = "Synonyms"
stranswers(6) = "Exclusive"
stranswers(7) = "Commit"
stranswers(8) = "Shadow paging"
stranswers(9) = "ISAM"
stranswers(10) = "YY.MM.DD"
End Sub

Private Sub Matching_Columns_Paper4()
strcolumnA(1) = "Specialization"
strcolumnA(2) = "External Model"
strcolumnA(3) = "Distributed Concurrent Control"
strcolumnA(4) = "Display tuples"
strcolumnA(5) = "Overflow management"
strcolumnA(6) = "Data dependencies"
strcolumnA(7) = "Query expression"
strcolumnA(8) = "Collision resolution"
strcolumnA(9) = "Conflict equivalence"
strcolumnA(10) = "Network data model"

strcolumnB(1) = "Linear probing"
strcolumnB(2) = "Cover"
strcolumnB(3) = "Precedence graph"
strcolumnB(4) = "Inheritance"
strcolumnB(5) = "ORDER BY "
strcolumnB(6) = "PSB"
strcolumnB(7) = "Biased locking "
strcolumnB(8) = "Hyper graph"
strcolumnB(9) = "Singular set"
strcolumnB(10) = "Push Through"

stranswers(1) = "Inheritance"
stranswers(2) = "PSB"
stranswers(3) = "Biased locking"
stranswers(4) = "ORDER BY "
stranswers(5) = "Push Through"
stranswers(6) = "Cover"
stranswers(7) = "Hyper graph"
stranswers(8) = "Linear probing"
stranswers(9) = "Precedence graph"
stranswers(10) = "Singular set"
End Sub
Private Sub Matching_Columns_Paper5()
strcolumnA(1) = "View equivalence"
strcolumnA(2) = "Update anomaly"
strcolumnA(3) = "Aggregate function"
strcolumnA(4) = "Tuples substitution"
strcolumnA(5) = "Foreign key"
strcolumnA(6) = "Latency"
strcolumnA(7) = "Integrity checking"
strcolumnA(8) = "Dynamic hashing"
strcolumnA(9) = "Schema specification"
strcolumnA(10) = "Adequate decomposition"

strcolumnB(1) = "Direct access storage device"
strcolumnB(2) = "Trigger"
strcolumnB(3) = "Bucket splitting"
strcolumnB(4) = "Lossless join"
strcolumnB(5) = "First Normal Form"
strcolumnB(6) = "Join strategy"
strcolumnB(7) = "Polygraph"
strcolumnB(8) = "CREATE TABLE"
strcolumnB(9) = "Referential integrity"
strcolumnB(10) = "SUM"

stranswers(1) = "Polygraph"
stranswers(2) = "First Normal Form"
stranswers(3) = "SUM"
stranswers(4) = "Join strategy"
stranswers(5) = "Referential integrity"
stranswers(6) = "Direct access storage device"
stranswers(7) = "Trigger"
stranswers(8) = "Bucket splitting"
stranswers(9) = "CREATE TABLE"
stranswers(10) = "Lossless join"

End Sub
Private Sub Matching_Columns_Paper6()
strcolumnA(1) = "Universal quantifier"
strcolumnA(2) = "SQL string matching operator"
strcolumnA(3) = "Set operation"
strcolumnA(4) = "Weak entity set"
strcolumnA(5) = "Mapping constraint"
strcolumnA(6) = "Data dependency"
strcolumnA(7) = "Condition box"
strcolumnA(8) = "Buffer replacement strategy"
strcolumnA(9) = "Transaction state"
strcolumnA(10) = "Deadlock Detection"

strcolumnB(1) = "Active"
strcolumnB(2) = "LRU"
strcolumnB(3) = "NOT EXISTS"
strcolumnB(4) = "LIKE"
strcolumnB(5) = "1 : N"
strcolumnB(6) = "MVD"
strcolumnB(7) = "Doubly outlined box"
strcolumnB(8) = "MINUS"
strcolumnB(9) = "QBE"
strcolumnB(10) = "WFG"

stranswers(1) = "NOT EXISTS"
stranswers(2) = "LIKE"
stranswers(3) = "MINUS"
stranswers(4) = "Doubly outlined box"
stranswers(5) = "1 : N"
stranswers(6) = "MVD"
stranswers(7) = "QBE"
stranswers(8) = "LRU"
stranswers(9) = "Active"
stranswers(10) = "WFG"
End Sub

Private Sub Matching_Columns_Paper7()
strcolumnA(1) = "Sequential retrieval"
strcolumnA(2) = "Bachman diagram"
strcolumnA(3) = "Nested lop approach"
strcolumnA(4) = "Retention option"
strcolumnA(5) = "Reducer operation"
strcolumnA(6) = "Indexed sequential file"
strcolumnA(7) = "Currency indicator"
strcolumnA(8) = "Starvation"
strcolumnA(9) = "Union rule"
strcolumnA(10) = "Universal instance assumption"

strcolumnB(1) = "Selection"
strcolumnB(2) = "Distributed free space"
strcolumnB(3) = "Null values"
strcolumnB(4) = "Set type representation"
strcolumnB(5) = "Get Next"
strcolumnB(6) = "FIXED"
strcolumnB(7) = "Join strategy"
strcolumnB(8) = "Live lock"
strcolumnB(9) = "Current of run unit"
strcolumnB(10) = "Functional dependency"

stranswers(1) = "Get Next"
stranswers(2) = "Set type representation"
stranswers(3) = "Join strategy"
stranswers(4) = "FIXED"
stranswers(5) = "Selection"
stranswers(6) = "Distributed free space"
stranswers(7) = "Current of run unit"
stranswers(8) = "Live lock"
stranswers(9) = "Functional dependency"
stranswers(10) = "Null values"

End Sub

Private Sub Matching_Columns_Paper8()
strcolumnA(1) = "Insertion option"
strcolumnA(2) = "Set selection by specifying a field of the owner record"
strcolumnA(3) = "Deferred splitting"
strcolumnA(4) = "Resilient to site failure"
strcolumnA(5) = "Vertical Fragmentation"
strcolumnA(6) = "Two key cryptography"
strcolumnA(7) = "Complexity of sort merge join strategy"
strcolumnA(8) = "Hierarchical pointer"
strcolumnA(9) = "Connection graph"
strcolumnA(10) = "Duplicate removal"

strcolumnB(1) = "MANUAL"
strcolumnB(2) = "SET SELECTION IS BY VALUE OF"
strcolumnB(3) = "Representation of a relational query expression"
strcolumnB(4) = "Two phase commit protocol"
strcolumnB(5) = "Projection"
strcolumnB(6) = "O(nlogon)"
strcolumnB(7) = "SELECT DISTINCT FROM"
strcolumnB(8) = "HDAM Storage organization"
strcolumnB(9) = "Public key cryptosystem"
strcolumnB(10) = "Dynamic hashing"

stranswers(1) = "MANUAL"
stranswers(2) = "SET SELECTION IS BY VALUE OF"
stranswers(3) = "Dynamic hashing"
stranswers(4) = "Two phase commit protocol"
stranswers(5) = "Projection"
stranswers(6) = "Public key cryptosystem"
stranswers(7) = "O(nlogon)"
stranswers(8) = "HDAM Storage organization"
stranswers(9) = "Representation of a relational query expression"
stranswers(10) = "SELECT DISTINCT FROM"
End Sub
Private Sub Matching_Columns_Paper9()
strcolumnA(1) = "Chile/twin pointers"
strcolumnA(2) = "System default"
strcolumnA(3) = "Mandatory"
strcolumnA(4) = "Procedural DML"
strcolumnA(5) = "Complementation rule"
strcolumnA(6) = "Project Join mapping"
strcolumnA(7) = "Segment search argument"
strcolumnA(8) = "Acyclic hyper graph"
strcolumnA(9) = "Schema specification storage"
strcolumnA(10) = "Blocking problem"

strcolumnB(1) = "Set retention option"
strcolumnB(2) = "Set ordering option"
strcolumnB(3) = "Relational calculus"
strcolumnB(4) = "IMS DL/1"
strcolumnB(5) = "HIDAM storage organization"
strcolumnB(6) = "Lossless join"
strcolumnB(7) = "MVD axiom"
strcolumnB(8) = "Full reducer"
strcolumnB(9) = "Data dictionary"
strcolumnB(10) = "Two phase commit protocol"

stranswers(1) = "HIDAM storage organization"
stranswers(2) = "Set ordering option"
stranswers(3) = "Set retention option"
stranswers(4) = "Relational calculus"
stranswers(5) = "MVD axiom"
stranswers(6) = "Lossless join"
stranswers(7) = "IMS DL/1"
stranswers(8) = "Full reducer"
stranswers(9) = "Data dictionary"
stranswers(10) = "Two phase commit protocol"


End Sub

Private Sub Matching_Columns_Paper10()
strcolumnA(1) = "Commutative operation"
strcolumnA(2) = "Conjunctive queries"
strcolumnA(3) = "FIND"
strcolumnA(4) = "CONNECT"
strcolumnA(5) = "Deadlock free protocol"
strcolumnA(6) = "Dependency among nonprime attributes"
strcolumnA(7) = "B+- tree"
strcolumnA(8) = "Cylinder index"
strcolumnA(9) = "Ellipses "
strcolumnA(10) = "Fourth Normal Form"

strcolumnB(1) = "2NF"
strcolumnB(2) = "Natural join"
strcolumnB(3) = "Sets current of run unit"
strcolumnB(4) = "Inserts member records into set instances"
strcolumnB(5) = "Safe domain calculus expressions"
strcolumnB(6) = "Leaf and non leaf nodes are of same size"
strcolumnB(7) = "Represent attributes"
strcolumnB(8) = "Multivalued dependence"
strcolumnB(9) = "Tree protocol"
strcolumnB(10) = "Indexed sequential file organization"

stranswers(1) = "Natural join"
stranswers(2) = "Safe domain calculus expressions"
stranswers(3) = "Sets current of run unit"
stranswers(4) = "Inserts member records into set instances"
stranswers(5) = "Tree protocol"
stranswers(6) = "2NF"
stranswers(7) = "Leaf and non leaf nodes are of same size"
stranswers(8) = "Indexed sequential file organization"
stranswers(9) = "Represent attributes"
stranswers(10) = "Multivalued dependence"
End Sub

Private Sub TrueorFalse_Paper1()
StrQuestions(1) = "Join operation can be used to implement a selection operation."
StrQuestions(2) = "The intersection of two super keys is necessarily a key."
StrQuestions(3) = "In SQL, attributes of different relations can have same name."
StrQuestions(4) = "In an SQL relation duplicates may be present."
StrQuestions(5) = "CREATE TABLE command is SQL can be used o create a virtual relation."
StrQuestions(6) = "Views defined using grouping and aggregate functions is updatable."
StrQuestions(7) = "In a nested SQL query, the result of the outer query is passed to the inner query."
StrQuestions(8) = "The universal quantifier in relational calculus can be implemented using the NOT EXISTS construct in SQL."
StrQuestions(9) = "The relational algebra is a declarative language."
StrQuestions(10) = "A weak entity set has a primary key."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper2()
StrQuestions(1) = ".    Every weak entity set can be converted in to strong entity set by adding appropriate attributes."
StrQuestions(2) = "Let r be a relation over a scheme R and A eð ðR, but A is not subset of R. Then the following equality holds."
StrQuestions(3) = "For any tow relations q and r Q || r = q || (q || r)"
StrQuestions(4) = "Any equijoin can be expressed using natural join and renaming operators."
StrQuestions(5) = "The transitive closure of any binary relation can be computer using relational algebra."
StrQuestions(6) = "A candidate key is a subset of a super key."
StrQuestions(7) = "Insertion into an indexed sequential file may require storage of record in the overflow area."
StrQuestions(8) = "An indexed sequential file is not suitable for applications which require batch access."
StrQuestions(9) = "A good hash function should distribute the records uniformly among the available buckets."
StrQuestions(10) = "    In a B-tree non leaf nodes are smaller than the leaf nodes."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper3()
StrQuestions(1) = "Starting from the current of the run unit, FIND ANY command locates the next record that satisfies the search criterion."
StrQuestions(2) = "If the retention option of a set type is MANDATORY, then a member record of a set occurrence can be moved to another set occurrence using the RECONNECT command."
StrQuestions(3) = "The current of a set type will always be the current of the run unit."
StrQuestions(4) = "A DBTG set represents man to many relationship among record types."
StrQuestions(5) = "A singular set can have only one set occurrence."
StrQuestions(6) = "The current of the run unit is the most recently accessed record of the database."
StrQuestions(7) = "The ERASE command can always be used to delete member records of any set occurrence."
StrQuestions(8) = "Many to many relationship between two entities can be represented  by two DBTG sets."
StrQuestions(9) = "A set type is called recursive set if the same record type is both owner and member of the set."
StrQuestions(10) = "FIND FIRST command is used to locate the first owner record of a set type."

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

Private Sub TrueorFalse_Paper4()
StrQuestions(1) = "In IMS, the schema of a hierarchical database is defined using DBDGEN control statements."
StrQuestions(2) = "A program specification block (PSB) defines the external view of an IMS hierarchical database."
StrQuestions(3) = "The Get Next command in DL/I can only retrieve the children records of particular segment occurrence."
StrQuestions(4) = "In the IMS storage organization a hierarchical pointer attached to a segment points to the next segment occurrence under the same parent."
StrQuestions(5) = "The segment search argument with a GET Unique command must specify a unique search condition for each level off the hierarchy till the target segment."
StrQuestions(6) = "Armstrong's axioms for FDs are sound but not complete."
StrQuestions(7) = "If an FD f can be derived from a set of FDs F, then f can be derived from F without using the reflexivity axiom."
StrQuestions(8) = "Given a relation scheme R with a set of FDs F. Then R has a unique third normal form lossless join decomposition that preserves the dependencies."
StrQuestions(9) = "Given a set of FDs for a relation scheme R, then a key of R can be determined using an algorithm having quadratic time complexity."
StrQuestions(10) = "The FDs XY -> Z and Z -> X, imply Z -> Y."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper5()
StrQuestions(1) = "The MVVD X->-> Y implies the FD X -> Y."
StrQuestions(2) = "The functional dependency X -> Y implies the multivalued dependency X ->-> Y."
StrQuestions(3) = "If a relation R(XYZ) satisfies the FD X -> Y. then R has a lossless join decomposition {R1(XY),R2(YZ)}."
StrQuestions(4) = "Suppose that a relation scheme R with a set of functional dependencies F has a BCNF decomposition (R1,R2,…,Rk) that preserve the dependencies. Then this decomposition would always be lossless join with respect to F."
StrQuestions(5) = "Given a relation scheme R with a set of FDs F. Then it is always possible to find a decomposition of R in BCNF schemes which preserves the given dependencies."
StrQuestions(6) = "For given instances of relations in a join expression, a full reducer may exist even when the join graph has a cycle."
StrQuestions(7) = "If two sets of functional dependencies F and G are equivalent, then either F is a subset of G or G is a subset of F."
StrQuestions(8) = "The schedules generated by a scheduler following strict two phase locking protocol would be free of deadlock."
StrQuestions(9) = "Even when transactions are allowed to write into the database only after they commit, undo operation would be required during recovery."
StrQuestions(10) = "If two transactions are view equivalent then that would also be conflict equivalent."

stranswers(1) = "False"
stranswers(2) = "True"
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
StrQuestions(1) = "If two transactions are view equivalent, then they would also be conflict equivalent."
StrQuestions(2) = "A timestamp based scheduler cannot produce live lock."
StrQuestions(3) = "Even if transactions in a schedule do not follow two phase locking, they may still be serializable."
StrQuestions(4) = "Basic timestamp based scheduler may produce cascading abort of transactions."
StrQuestions(5) = "Tree locking protocol does not ensure deadlock free schedule."
StrQuestions(6) = "In an RDBMS a user who has been granted access to a data appearing in a view will automatically have the privilege to permit any other user to have same access right."
StrQuestions(7) = "In a relational database, the foreign key of relation would be the primary key of an another relation."
StrQuestions(8) = "An indexed sequential file with push through approach for overflow management requires only one overflow pointer per block of primary records."
StrQuestions(9) = "An ODBC transaction can span multiple databases."
StrQuestions(10) = "A database of a particular DBMS on some given machine can serve as an ODBC data source if it is known to the ODBC driver manager."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper7()
StrQuestions(1) = "An ODBC transaction can be nested."
StrQuestions(2) = "An ODBC API is common for all types of relational databases."
StrQuestions(3) = "A trigger can be used to enforce that a relational database satisfies an FD."
StrQuestions(4) = "The semi join of a relation r with another relation s,(r | s) produces a relation which is a subset of s."
StrQuestions(5) = "Let r and s be two relations over two disjoint schemes R and S respectively. Then ( r × s) / s = r."
StrQuestions(6) = "A two phase schedule where transactions are allowed to convert locks before releasing then, would still be serializable."
StrQuestions(7) = "The functional dependencies X -> Y and YZ -> W imply the functional dependent XZ -> W."
StrQuestions(8) = "The transitive closure of a binary relation can be computer using relational calculus."
StrQuestions(9) = "Every two phased schedule can be produced by a timestamp ordered scheduler."
StrQuestions(10) = "The decomposition axiom of functional dependencies can be derived from the reflexivity, augmentation and transitivity rules."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper8()
StrQuestions(1) = "Boyce – Codd Normal form and Third normal form are identical ."
StrQuestions(2) = "Deleting a record from an inverted file is more difficult than deleting a record from a multilink file."
StrQuestions(3) = "A relation R is in Second Normal Form if and only if every non key attribute is non transitively dependent on the primary key."
StrQuestions(4) = "The subschema must contain all the descriptions of certain data that are in the scheme."
StrQuestions(5) = "The details regarding the users of data is included in the data dictionary."
StrQuestions(6) = "If the size of the two databases are 10 kb and 20 kb respectively the size of a database created using the JOIN command will be less than or equal 30 kb."
StrQuestions(7) = "Two database files can be combined to create a report file."
StrQuestions(8) = "CODASYL DBTG model is an example of hierarchical data model."
StrQuestions(9) = "To update a file using another file, both the files must be sorted on the same key."
StrQuestions(10) = "INDEX command can be used to index only selected number of records."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper9()
StrQuestions(1) = "The deletion of records can be prohibited by the BROWSE command."
StrQuestions(2) = "If records are deleted from the database with PACK, the6 can be undeleted by the RECALL command."
StrQuestions(3) = "DISPLAY NEXT 6 TO PRINTER FOR PIN = “721302” FIELDS CITY, PIN is invalid."
StrQuestions(4) = "SET MEMOWIDTH can be used to change the width of memo window."
StrQuestions(5) = "When a database file that has a structural compound index file is opened, the first index tag becomes the controlling index tag."
StrQuestions(6) = "FOXPRO system automatically opens the compound index file along with the associated database file."
StrQuestions(7) = "USE SUPPLIER IN 4 opens SUPPLIER.dbf in work area 4 is different files but cannot make modifications."
StrQuestions(8) = "When linked files are browsed, one can view fields from different files but cannot make modifications."
StrQuestions(9) = "JOIN creates a new database file by joining two or more database files."
StrQuestions(10) = "    Only two files can be linked by a single SET RELATION command."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper10()
StrQuestions(1) = "A portion of a custom screen can be enclosed in a box."
StrQuestions(2) = "While designing a screen file, any object on the screen can be selected by moving the cursor on that object and pressing the spacebar."
StrQuestions(3) = "While specifying the field list with LIST, one cannot alter the field order from the one used in the database file."
StrQuestions(4) = "You may edit records of an index file with BROWSE command."
StrQuestions(5) = "A file cannot be indexed on a memo field."
StrQuestions(6) = "Structure of a target file should be defined before using SORT command."
StrQuestions(7) = "SET CONFIRM can be used to warn a user before overwriting a file."
StrQuestions(8) = " + and -  are mathematical as well as string operators."
StrQuestions(9) = "? DTOC (“12/01/96”) + “TUESDAY” is valid."
StrQuestions(10) = "    SCATTER TO A moves data from database record to array "

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub multiplechoice_Paper1()
strmultiple(1) = "The level of data abstraction which describes how the data is actually stored is"
strmuloption(1) = "Physical level"
strmuloption(2) = "Conceptual level"
strmuloption(3) = "Storage level"
strmuloption(4) = "File level"

strmultiple(2) = "The level of abstraction describes only part of the entire data base iscalled-"
strmuloption(5) = "Conceptual level"
strmuloption(6) = "View level"
strmuloption(7) = "Procedural level"
strmuloption(8) = "None of the above"

strmultiple(3) = "Object based data models are used in describing the abstraction of the following level(s)-"
strmuloption(9) = "Conceptual and View"
strmuloption(10) = "Only Physical"
strmuloption(11) = "Physical and Conceptual"
strmuloption(12) = "Only Conceptual"

strmultiple(4) = "An association of several entities in a Entity-Relation model is called -"
strmuloption(13) = "Tuple"
strmuloption(14) = "Record"
strmuloption(15) = "Relationship"
strmuloption(16) = "Field"

strmultiple(5) = "The overall logical structure of a database can be expressed graphically by -"
strmuloption(17) = "Data Flow Chart"
strmuloption(18) = "Flow Chart"
strmuloption(19) = "Directed graph"
strmuloption(20) = "Entity Relationship diagram"

strmultiple(6) = "In an object oriented model, one object can access data of another object by passing -"
strmuloption(21) = "Instance variable"
strmuloption(22) = "Variable"
strmuloption(23) = "Message"
strmuloption(24) = "Function"

strmultiple(7) = "The collection of information stored in a database at a particular moment is -"
strmuloption(25) = "View"
strmuloption(26) = "Instance"
strmuloption(27) = "Scheme"
strmuloption(28) = "None of these"

strmultiple(8) = "The overall design of a database is called -"
strmuloption(29) = "Scheme of the database"
strmuloption(30) = "The screen of database"
strmuloption(31) = "Structure of the database"
strmuloption(32) = "View of the database"

strmultiple(9) = "The concept of data independence is similar to the concept of -"
strmuloption(33) = "Data type"
strmuloption(34) = "Abstract data type"
strmuloption(35) = "Consolidation"
strmuloption(36) = "Isolation"

strmultiple(10) = " The tables generated on compilation of data definition language(DDL) are stored in -"
strmuloption(37) = "Data abstraction"
strmuloption(38) = "Data dictionary"
strmuloption(39) = "Data file"
strmuloption(40) = "Meta data"

strmulanswer(1) = "Physical level"
strmulanswer(2) = "View level"
strmulanswer(3) = "Conceptual and View"
strmulanswer(4) = "Relationship"
strmulanswer(5) = "Entity Relationship diagram"
strmulanswer(6) = "Message"
strmulanswer(7) = "Instance"
strmulanswer(8) = "Scheme of the database"
strmulanswer(9) = "Abstract data type"
strmulanswer(10) = "Data dictionary"

End Sub

Private Sub multiplechoice_Paper2()
strmultiple(1) = "The statement requesting the retrieval of information is usually in the form of -"
strmuloption(1) = "Report"
strmuloption(2) = "Query"
strmuloption(3) = "Data Entry screen"
strmuloption(4) = "Record"

strmultiple(2) = "The part of data manipulation language (DML) that involves information retrieval is called -"
strmuloption(5) = "Query Language"
strmuloption(6) = "Report Generation"
strmuloption(7) = "Data definition language"
strmuloption(8) = "Meta language"

strmultiple(3) = "This is NOT one of the functions of a database administrator -"
strmuloption(9) = "Granting of authorization for data access "
strmuloption(10) = "Scheme definition"
strmuloption(11) = "Interaction with file manager"
strmuloption(12) = "Storage structure and access method definition"

strmultiple(4) = "  Data about data is normally termed as -"
strmuloption(13) = "Meta data"
strmuloption(14) = "Data dictionary"
strmuloption(15) = "Data base"
strmuloption(16) = "None of these"

strmultiple(5) = "  If D1,D2,…Dn are domains in a relational model then the relation is a table which is a subset of -"
strmuloption(17) = "{D1,D2…Dn}"
strmuloption(18) = "D1 × D2 × … Dn"
strmuloption(19) = "D1 U D2 U … Dn"
strmuloption(20) = "Maximum {D1, D2 … Dn}"

strmultiple(6) = "  A variable in programming language corresponds to the following in Entity - Relation model -"
strmuloption(21) = "Instance"
strmuloption(22) = "Record"
strmuloption(23) = "Entity"
strmuloption(24) = "Key"

strmultiple(7) = "  The number of entities to which another entity can be associated via a relationship set is expressed as -"
strmuloption(25) = "Entity"
strmuloption(26) = "Cardinality"
strmuloption(27) = "Scheme"
strmuloption(28) = "Attributes"

strmultiple(8) = "  The mapping cardinality for a binary relationship between entity sets A and B can be -"
strmuloption(29) = "One-to-one only"
strmuloption(30) = "One-to-many only"
strmuloption(31) = "Many-to-one only"
strmuloption(32) = "One-to-one, one-to-many, many-to-one or many-to-many"

strmultiple(9) = "  If entity x is existence -dependent on entity y then x is said to be -"
strmuloption(33) = "Subordinate entity"
strmuloption(34) = "Dominant entity"
strmuloption(35) = "Primary entity"
strmuloption(36) = "Secondary entity"

strmultiple(10) = " One limitation of the Entity - relation model is that it cannot"
strmuloption(37) = "use generalization"
strmuloption(38) = "express relationships among relationships"
strmuloption(39) = "Use one single primary key."
strmuloption(40) = "None of the above"

strmulanswer(1) = "Query"
strmulanswer(2) = "Query Language"
strmulanswer(3) = "Storage structure and access method definition"
strmulanswer(4) = "Meta data"
strmulanswer(5) = "D1 × D2 × … Dn"
strmulanswer(6) = "Entity"
strmulanswer(7) = "Cardinality"
strmulanswer(8) = "One-to-one, one-to-many, many-to-one or many-to-many"
strmulanswer(9) = "Subordinate entity"
strmulanswer(10) = "express relationships among relationships"

End Sub

Private Sub multiplechoice_Paper3()
strmultiple(1) = "  Every weak entity set can be converted to a strong entity set by -"
strmuloption(1) = "Simply adding appropriate attributes"
strmuloption(2) = "Using generalization"
strmuloption(3) = "Using aggregation"
strmuloption(4) = "Repeating the entity set several times"

strmultiple(2) = "  An Entity - Relation diagram having same entity set appearing several times is -"
strmuloption(5) = "Good practice"
strmuloption(6) = "Bad practice"
strmuloption(7) = "Effective use of entities"
strmuloption(8) = "No effect on efficiency  of design"

strmultiple(3) = "  A relational model which allows non atomic domains is -"
strmuloption(9) = "Nested relational data model"
strmuloption(10) = "Non atomic data model"
strmuloption(11) = "Hierarchical data model"
strmuloption(12) = "Such model can not exists"

strmultiple(4) = "  A database which conforms to an Entity - Relationship diagram can be represented by a -"
strmuloption(13) = "Collection of tables"
strmuloption(14) = "A tuple"
strmuloption(15) = "Relations"
strmuloption(16) = "None of the above"

strmultiple(5) = "  If E1 and E2 are relational algebra expressions, then which of the following is NOT a relational algebra expression"
strmuloption(17) = "E1 U E2"
strmuloption(18) = "E1 / E2"
strmuloption(19) = "E1 - E2"
strmuloption(20) = "E1 × E2"

strmultiple(6) = "  A modification to the database expressed in terms of a view must be translated to the -"
strmuloption(21) = "actual relations in the conceptual model off the database"
strmuloption(22) = "queries in the actual database"
strmuloption(23) = "relations of all the views of that database"
strmuloption(24) = "Need not be translated and the view of database accommodates the modifications"

strmultiple(7) = "  Tuple variables in SQL are defined in -"
strmuloption(25) = "From clause"
strmuloption(26) = "Select clause"
strmuloption(27) = "Where clause"
strmuloption(28) = "All of these"

strmultiple(8) = "  Which one of these is an aggregate function is Sql?"
strmuloption(29) = "avg"
strmuloption(30) = "select"
strmuloption(31) = "ordered by"
strmuloption(32) = "None of these"

strmultiple(9) = "  Quel is a query language for the following database:"
strmuloption(33) = "DBASEIV"
strmuloption(34) = "ORACLE"
strmuloption(35) = "INGRESS"
strmuloption(36) = "DBII"


strmulanswer(1) = "Simply adding appropriate attributes"
strmulanswer(2) = "Bad practice"
strmulanswer(3) = "Nested relational data model"
strmulanswer(4) = "Collection of tables"
strmulanswer(5) = "E1 / E2"
strmulanswer(6) = "actual relations in the conceptual model off the database"
strmulanswer(7) = "From clause"
strmulanswer(8) = "avg"
strmulanswer(9) = "INGRESS"

End Sub

Private Sub multiplechoice_Paper4()
strmultiple(1) = "  Every Boyce-Codd Normal Form(BCNF) is in -"
strmuloption(1) = "First Normal Form (1NF)"
strmuloption(2) = "Second Normal Form (2NF)"
strmuloption(3) = "Third Normal Form(3NF)"
strmuloption(4) = "None of these"

strmultiple(2) = "  Functional dependencies are a generalization of -"
strmuloption(5) = "Key dependencies"
strmuloption(6) = "Relation dependencies"
strmuloption(7) = "Database dependencies"
strmuloption(8) = "None of these"

strmultiple(3) = "  Every Boyce -Codd Normal Form (BCNF) decomposition is -"
strmuloption(9) = "Dependency preserving"
strmuloption(10) = "Not dependence preserving"
strmuloption(11) = "Need be dependency preserving"
strmuloption(12) = "None of these"

strmultiple(4) = "  Which normal form is most desirable?"
strmuloption(13) = "Domain Key normal Form(DKNF)"
strmuloption(14) = "Boyce-Codd Normal Form (BCNF)"
strmuloption(15) = "Third Normal Form(3NF)"
strmuloption(16) = "Forth Normal Form (4NF)"

strmultiple(5) = "  The main disadvantage of Indexed Sequential File is -"
strmuloption(17) = "It consumes lot of space"
strmuloption(18) = "Performance degrades with the increase in file size"
strmuloption(19) = "Accessing is very difficult."
strmuloption(20) = "None of the above"

strmultiple(6) = "  The redundant storage of search key values in B+ trees is eliminated in -"
strmuloption(21) = "B-Trees"
strmuloption(22) = "Indexed sequential file"
strmuloption(23) = "AVL trees"
strmuloption(24) = "Binary trees"

strmultiple(7) = "  the strategy for processing a query is improved by -"
strmuloption(25) = "Query evaluation"
strmuloption(26) = "Query optimization"
strmuloption(27) = "Decomposition"
strmuloption(28) = "None of the above"

strmultiple(8) = "  The translation of a query into its internal form is usually done using -"
strmuloption(29) = "Query processing"
strmuloption(30) = "Relational algebra"
strmuloption(31) = "Database scheme"
strmuloption(32) = "The structure of data base"

strmultiple(9) = "  The syntax of the users query is verified by the -"
strmuloption(33) = "Parser"
strmuloption(34) = "The database manager"
strmuloption(35) = "Query optimization"
strmuloption(36) = "Database administrator"

strmultiple(10) = " A join of three or more relations can be processed in parallel by constructing a -"
strmuloption(37) = "Join-decomposition"
strmuloption(38) = "Pipeline"
strmuloption(39) = "Parallel query"
strmuloption(40) = "Merge -join"


strmulanswer(1) = "Third Normal Form(3NF)"
strmulanswer(2) = "Key dependencies"
strmulanswer(3) = "Need be dependency preserving"
strmulanswer(4) = "Domain Key normal Form(DKNF)"
strmulanswer(5) = "Performance degrades with the increase in file size"
strmulanswer(6) = "B-Trees"
strmulanswer(7) = "Query evaluation"
strmulanswer(8) = "Relational algebra"
strmulanswer(9) = "Parser"
strmulanswer(10) = "Pipeline"
End Sub

Private Sub multiplechoice_Paper5()
strmultiple(1) = "  Roll back of transactions is normally used to -"
strmuloption(1) = "Recover from the transaction failure"
strmuloption(2) = "Update the transaction"
strmuloption(3) = "Retrieve old records"
strmuloption(4) = "Restore the record"

strmultiple(2) = "  Cascading Roll-backs can be avoided by using -"
strmuloption(5) = "Validation protocol"
strmuloption(6) = "Time stamped based protocol"
strmuloption(7) = "Transaction processing"
strmuloption(8) = "None of the above"

strmultiple(3) = "  If there exists a set of transactions such that every transaction in the set is waiting for another transaction in the set, then the system is called -"
strmuloption(9) = "Transaction failure"
strmuloption(10) = "Transaction Processing"
strmuloption(11) = "Dead-lock"
strmuloption(12) = "Two phase lock"

strmultiple(4) = "  To ensure that the system will never enter a deadlock, we can use -"
strmuloption(13) = "Deadlock prevention protocol"
strmuloption(14) = "Boyce codd Normal Form"
strmuloption(15) = "Time - stamped transaction"
strmuloption(16) = "Validation protocol"

strmultiple(5) = "  Locking the data before the beginning of transaction execution in order to prevent deadlock may lead to -"
strmuloption(17) = "Program abortion"
strmuloption(18) = "Starvation"
strmuloption(19) = "Delay"
strmuloption(20) = "Poor performance"

strmultiple(6) = "  A deadlock prevention scheme which avoid starvation is -"
strmuloption(21) = "Locking of data by transactions before execution"
strmuloption(22) = "A deadlock prevention scheme using time stamps"
strmuloption(23) = "Deadlock detection and prevention scheme."
strmuloption(24) = "All of the above"

strmultiple(7) = "  A directed graph which represent the deadlock is called -"
strmuloption(25) = "Cyclic graph"
strmuloption(26) = "Deadlock graph"
strmuloption(27) = "Wait-for graph"
strmuloption(28) = "Deadlock detection graph"

strmultiple(8) = "  A deadlock exists in the system if and only if the wait for graph -"
strmuloption(29) = "Contains a cycle "
strmuloption(30) = "There is a path from first node to last node"
strmuloption(31) = "There is no path from first node to last node"
strmuloption(32) = "Contains a non - reachable node from the first node."

strmultiple(9) = "  In order to reduce the overhead imposed on transaction processing by the O.S. one can use -"
strmuloption(33) = "Transaction monitor"
strmuloption(34) = "Parallel processing"
strmuloption(35) = "Concurrency in transactions"
strmuloption(36) = "High performance transactions"

strmultiple(10) = " When the detection algorithm detects a deadlock, the recovery is normally accomplished by -"
strmuloption(37) = "Roll back of transactions"
strmuloption(38) = "Consistency checking"
strmuloption(39) = "Locking of data"
strmuloption(40) = "Calling back the same transaction"


strmulanswer(1) = "Recover from the transaction failure"
strmulanswer(2) = "Validation protocol"
strmulanswer(3) = "Dead-lock"
strmulanswer(4) = "Deadlock prevention protocol"
strmulanswer(5) = "Starvation"
strmulanswer(6) = "Locking of data by transactions before execution"
strmulanswer(7) = "Wait-for graph"
strmulanswer(8) = "Contains a cycle"
strmulanswer(9) = "Transaction monitor"
strmulanswer(10) = "Roll back of transactions"

End Sub

Private Sub multiplechoice_Paper6()
strmultiple(1) = "  The operation of a relation X, produces Y such that T contains only selected attributes of S. Such an operation is -"
strmuloption(1) = "Projection"
strmuloption(2) = "Intersection"
strmuloption(3) = "Union"
strmuloption(4) = "Difference"

strmultiple(2) = "  Complex entities can be constructed using the following operation -"
strmuloption(5) = "Sum"
strmuloption(6) = "Union"
strmuloption(7) = "Aggregation"
strmuloption(8) = "Collection"

strmultiple(3) = "  The maximum height of B+ tree of order n with k key values is -"
strmuloption(9) = "(n+k) / 2"
strmuloption(10) = "log nk"
strmuloption(11) = "log n/2 (k/2+1)"
strmuloption(12) = "log k/2n"

strmultiple(4) = "  The following entities / attributes in a relational database should not have null values:"
strmuloption(13) = "Keys"
strmuloption(14) = "Relations"
strmuloption(15) = "Variables"
strmuloption(16) = "All of these"

strmultiple(5) = "  The following relationships need not be created each time the tables are used:"
strmuloption(17) = "Functional dependent "
strmuloption(18) = "Persistent"
strmuloption(19) = "Intersections"
strmuloption(20) = "Loss- less decomposition"

strmultiple(6) = "  Insertion of data / record into a B-tree may cause leaf and internal nodes to -"
strmuloption(21) = "Split"
strmuloption(22) = "Be deleted"
strmuloption(23) = "Maintain the same order"
strmuloption(24) = "Any one of these"

strmultiple(7) = "  A view of a database that appears to an application program is known as -"
strmuloption(25) = "Report"
strmuloption(26) = "Structure"
strmuloption(27) = "Subschema"
strmuloption(28) = "Scheme"

strmultiple(8) = "  The universal quantifier in relational calculus can be implemented using the following SQL command -"
strmuloption(29) = "SET RELATION TO"
strmuloption(30) = "NOT EXISTS"
strmuloption(31) = "SELECT"
strmuloption(32) = "None of these"

strmultiple(9) = "  This may be regarded as meta data -"
strmuloption(33) = "E-R diagram"
strmuloption(34) = "Table"
strmuloption(35) = "Data dictionary"
strmuloption(36) = "View of a database"

strmultiple(10) = " In a relational database a referential integrity constraint is specified with the help of -"
strmuloption(37) = "Primary key"
strmuloption(38) = "Foreign key"
strmuloption(39) = "Secondary key"
strmuloption(40) = "Consistency checking"

strmulanswer(1) = "Projection"
strmulanswer(2) = "Aggregation"
strmulanswer(3) = "log n/2 (k/2+1)"
strmulanswer(4) = "Keys"
strmulanswer(5) = "Persistent"
strmulanswer(6) = "Split"
strmulanswer(7) = "Subschema"
strmulanswer(8) = "NOT EXISTS"
strmulanswer(9) = "Data dictionary"
strmulanswer(10) = "Foreign key"

End Sub

Private Sub multiplechoice_Paper7()
strmultiple(1) = "  The two phase commit protocol is resilient to -"
strmuloption(1) = "Deadlock"
strmuloption(2) = "Site and link failures"
strmuloption(3) = "Dependencies"
strmuloption(4) = "Starving"

strmultiple(2) = "  The to phase commit protocol may introduce -"
strmuloption(5) = "Deadlock"
strmuloption(6) = "Starving"
strmuloption(7) = "Blocking"
strmuloption(8) = "None of the above"

strmultiple(3) = "  A block that is not permitted to be written back to disk is called -"
strmuloption(9) = "Dead code"
strmuloption(10) = "Read only"
strmuloption(11) = "Pinned"
strmuloption(12) = "Zapped"

strmultiple(4) = "  In a relational data model the columns of a table are called -"
strmuloption(13) = "Fields"
strmuloption(14) = "Attributes"
strmuloption(15) = "Records"
strmuloption(16) = "Entities"

strmultiple(5) = "  A multi level indexing may lead to"
strmuloption(17) = "Deadlock"
strmuloption(18) = "B-tree"
strmuloption(19) = "Binary tree"
strmuloption(20) = "Transaction failure"

strmultiple(6) = "  The primary key indexing technique does not allow -"
strmuloption(21) = "Duplicate data in a field"
strmuloption(22) = "Multiple attributes"
strmuloption(23) = "Sets of relations"
strmuloption(24) = "Many to many relation"

strmultiple(7) = "  Before a transaction commits the following is updated:"
strmuloption(25) = "Log"
strmuloption(26) = "The query processor"
strmuloption(27) = "The database"
strmuloption(28) = "Database manager"

strmultiple(8) = "  All the objects created into a project are organized by the -"
strmuloption(29) = "Project manager"
strmuloption(30) = "Database manager"
strmuloption(31) = "Database administrator"
strmuloption(32) = "Sorting"

strmultiple(9) = "  The data model in which data is represented by collection of records and relationship among data are represented by links is called - "
strmuloption(33) = "Hierarchical data model"
strmuloption(34) = "Relational data model"
strmuloption(35) = "Network data model"
strmuloption(36) = "Logical data model"

strmultiple(10) = " It is a scheme representing the design of a network database:"
strmuloption(37) = "Entity-Relation diagram"
strmuloption(38) = "Data structure diagram"
strmuloption(39) = "Wait for graph"
strmuloption(40) = "Network of records/fields"


strmulanswer(1) = "Site and link failures"
strmulanswer(2) = "Blocking"
strmulanswer(3) = "Pinned"
strmulanswer(4) = "Attributes"
strmulanswer(5) = "B-tree"
strmulanswer(6) = "Duplicate data in a field"
strmulanswer(7) = "Log"
strmulanswer(8) = "Project manager"
strmulanswer(9) = "Network data model"
strmulanswer(10) = "Data structure diagram"


End Sub

Private Sub multiplechoice_Paper8()
strmultiple(1) = "  A data structure diagram of a network model is analogous:"
strmuloption(1) = "Entity Relation diagram"
strmuloption(2) = "Relational table"
strmuloption(3) = "B-tree representation"
strmuloption(4) = "None of these"

strmultiple(2) = "  The first network database is-"
strmuloption(5) = "Ingress"
strmuloption(6) = "Database I"
strmuloption(7) = "DBTG CODASYL"
strmuloption(8) = "SAP"

strmultiple(3) = "  In COSASYL many to many links are-"
strmuloption(9) = "Obtained from data structure diagram"
strmuloption(10) = "Obtained by transformation of E-R diagram to data structure diagram"
strmuloption(11) = "Converted to binary relations."
strmuloption(12) = "Not allowed"

strmultiple(4) = "  CODASYL allows a member record to exist on its own without being member of a set occurrence, if its retention is -"
strmuloption(13) = "Allowed"
strmuloption(14) = "Optional"
strmuloption(15) = "Persistent"
strmuloption(16) = "Dependent"

strmultiple(5) = "  For finding the duplicate records in a network database one can use the following DML command:"
strmuloption(17) = "FIND DUPLICATE"
strmuloption(18) = "AUTOMATIC"
strmuloption(19) = "SELECTION"
strmuloption(20) = "READ"

strmultiple(6) = "  In a Database Task Group(DBTG) model, the representation of weak entities in the E-R model is achieved by -"
strmuloption(21) = "DBTG sets"
strmuloption(22) = "Repeating groups"
strmuloption(23) = "Transformation "
strmuloption(24) = "None of these"

strmultiple(7) = "  In a network model, the presence of the following create certain problems,"
strmuloption(25) = "Cycles"
strmuloption(26) = "Duplication of data"
strmuloption(27) = "Weak entities"
strmuloption(28) = "Many to many relation"

strmultiple(8) = "  In Query by example the query is formed from the relations displayed by filling in the following:"
strmuloption(29) = "SQL form"
strmuloption(30) = "Templates"
strmuloption(31) = "DML"
strmuloption(32) = "DDL statements"

strmultiple(9) = "  The data manipulation language of the DBTG model consists of a number of commands that are embedded in -"
strmuloption(33) = "SQL"
strmuloption(34) = "Command language"
strmuloption(35) = "DDL"
strmuloption(36) = "Host language"

strmultiple(10) = " Which one of the following is an example for Network database?"
strmuloption(37) = "IDMS"
strmuloption(38) = "Oracle"
strmuloption(39) = "Ingress"
strmuloption(40) = "DBMS"

strmulanswer(1) = "Entity Relation diagram"
strmulanswer(2) = "DBTG CODASYL"
strmulanswer(3) = "Not allowed"
strmulanswer(4) = "Optional"
strmulanswer(5) = "FIND DUPLICATE"
strmulanswer(6) = "Repeating groups"
strmulanswer(7) = "Cycles"
strmulanswer(8) = "Templates"
strmulanswer(9) = "Host language"
strmulanswer(10) = "IDMS"

End Sub

Private Sub multiplechoice_Paper9()
strmultiple(1) = "  The two most frequently used DBTG commands are -"
strmuloption(1) = "Find and get"
strmuloption(2) = "Selection and from"
strmuloption(3) = "Retrieve from"
strmuloption(4) = "None of these"

strmultiple(2) = "  In a hierarchical database, a hashing function is used to locate the -"
strmuloption(5) = "Collisions"
strmuloption(6) = "Root"
strmuloption(7) = "Primary key"
strmuloption(8) = "Duplicate records"

strmultiple(3) = "  The scheme for a hierarchical database is-"
strmuloption(9) = "Tree structure diagram"
strmuloption(10) = "E-R diagram"
strmuloption(11) = "A graph of links between records"
strmuloption(12) = "None of these"

strmultiple(4) = "  In a hierarchical database if a relation includes a descriptive attribute , the transformation from an E-R diagram to tree structure is more complicated because:"
strmuloption(13) = "the description can not be quantified"
strmuloption(14) = "the link can not contain any data value"
strmuloption(15) = "The hierarchical nature can not be specified"
strmuloption(16) = "This leads to be ambiguous"

strmultiple(5) = "  In a hierarchical sequential access method(HSAM) the following operation is not permitted:"
strmuloption(17) = "Deletion"
strmuloption(18) = "Insertion"
strmuloption(19) = "Indexing"
strmuloption(20) = "Sorting"

strmultiple(6) = "  Which one of the following is an example for hierarchical database?"
strmuloption(21) = "Oracle"
strmuloption(22) = "Ingress"
strmuloption(23) = "DBII"
strmuloption(24) = "SYSTEM2000"

strmultiple(7) = "  In a hierarchical database, modeling of many to many relations is achieved by -"
strmuloption(25) = "Record replication"
strmuloption(26) = "DML"
strmuloption(27) = "DDL"
strmuloption(28) = "Preorder threads"

strmultiple(8) = "  In most of the implementations of hierarchical databases, parent to child links are stored with the help of -"
strmuloption(29) = "Pointers"
strmuloption(30) = "Reference variables"
strmuloption(31) = "Virtual records"
strmuloption(32) = "Preorder threads"

strmultiple(9) = "  In hierarchical database, a virtual record contains -"
strmuloption(33) = "the redundant data of a particular data (duplication of a particular record)"
strmuloption(34) = "Contains no data but contains logical pointer to a particular record"
strmuloption(35) = "Preorder threads"
strmuloption(36) = "Currency pointers"

strmultiple(10) = " Indexed sequential file is created using the following storage device:"
strmuloption(37) = "Direct access"
strmuloption(38) = "Random access"
strmuloption(39) = "Sequential access"
strmuloption(40) = "Primary"

strmulanswer(1) = "Find and get"
strmulanswer(2) = "Root"
strmulanswer(3) = "Tree structure diagram"
strmulanswer(4) = "the link can not contain any data value"
strmulanswer(5) = "Insertion"
strmulanswer(6) = "SYSTEM2000"
strmulanswer(7) = "Record replication"
strmulanswer(8) = "Preorder threads"
strmulanswer(9) = "Contains no data but contains logical pointer to a particular record"
strmulanswer(10) = "Direct access"

End Sub

Private Sub multiplechoice_Paper10()
strmultiple(1) = "An association among several entities is called -"
strmuloption(1) = "Table"
strmuloption(2) = "Relationship"
strmuloption(3) = "Data"
strmuloption(4) = "None of these"

strmultiple(2) = "The principle means of identifying entities within an entity set is-"
strmuloption(5) = "Tuple"
strmuloption(6) = "Primary key"
strmuloption(7) = "Pointer"
strmuloption(8) = "Record"

strmultiple(3) = "A language that enables users to access or manipulate data :"
strmuloption(9) = "Data manipulation language (DML)"
strmuloption(10) = "Data base language"
strmuloption(11) = "SQL"
strmuloption(12) = "All of these"

strmultiple(4) = "A language that is used to specify the data base scheme through a set ofdefinitions is called -"
strmuloption(13) = "SQL"
strmuloption(14) = "Data definition language(DDL)"
strmuloption(15) = "E-R diagram"
strmuloption(16) = "Fourth generation language(4GL)"

strmultiple(5) = "A program module in a DBMS which provides interface between low level data stored in the data base and the application programs & queries submitted to he system is -"
strmuloption(17) = "Query processor"
strmuloption(18) = "Database manager"
strmuloption(19) = "Database creator"
strmuloption(20) = "DBMS"

strmultiple(6) = "A collection of conceptual tools for describing data, data relationships, data semantics and data constraints is called:"
strmuloption(21) = "Data model"
strmuloption(22) = "E-R model"
strmuloption(23) = "Scheme"
strmuloption(24) = "All of these"

strmultiple(7) = "A hierarchical data model is -"
strmuloption(25) = "Totally ordered"
strmuloption(26) = "Partially ordered"
strmuloption(27) = "Not ordered"
strmuloption(28) = "Concept of ordering does not apply"

strmultiple(8) = "Any relational calculus expression can-"
strmuloption(29) = "Be expressed using relational algebra"
strmuloption(30) = "Not always be expressed using relational algebra"
strmuloption(31) = "Be decomposed"
strmuloption(32) = "None of these"

strmultiple(9) = "The multivalued  transitivity rule is -"
strmuloption(33) = "Sound"
strmuloption(34) = "Complete"
strmuloption(35) = "Not sound but complete"
strmuloption(36) = "Soundness does not apply"

strmultiple(10) = "SQL is relationally -"
strmuloption(37) = "Complete language"
strmuloption(38) = "Incomplete language"
strmuloption(39) = "Can not handle certain relations"
strmuloption(40) = "Sound language"


strmulanswer(1) = "Relationship"
strmulanswer(2) = "Primary key"
strmulanswer(3) = "Data manipulation language (DML)"
strmulanswer(4) = "Data definition language(DDL)"
strmulanswer(5) = "Database manager"
strmulanswer(6) = "Data model"
strmulanswer(7) = "Totally ordered"
strmulanswer(8) = "Be expressed using relational algebra"
strmulanswer(9) = "Sound"
strmulanswer(10) = "Complete language"

End Sub


Private Sub questions_paper1()
StrQuestions(1) = "(a) What is meant by a sound an complete set of inference axioms? When is a set of axioms said to be independent?" & Chr(13) & _
"(b)  Why is BCNF a more desirable normal form than lower order normal forms? Give an example of a relational scheme that is in 3NF but not is BCNF." & Chr(13) & _
"(c) Identify three basic goals to be satisfied while designing relational database schemes. Is there any reason to design a database scheme that is in 3NF but not in no other higher normal form?" & Chr(13) & _
"(d) What is meant by live lock? In a concurrency control scheme how can live lock be avoided?" & Chr(13) & _
"(e) When are two schedules said to be equivalent? How can equivalence of schedules be detected?" & Chr(13) & _
"(f) Identify possible types of failures that may occur in a distributed database environment. Which items of your list are applicable also to a centralized system?"
StrQuestions(2) = "a)  What is referential integrity?" & Chr(13) & _
"b) Differentiate clearly between privacy, security, correctness and integrity of a database." & Chr(13) & _
"c) How can we avoid deadlocks?" & Chr(13) & _
"d) State Armstrong's axioms for implication of function dependencies." & Chr(13) & _
"e) Let R(A,B,C,) be a relation scheme with the functional dependencies F = {A -> B, C-> , C}. What is the highest normal form of F?" & Chr(13) & _
"f) What do you mean by soundness of a set of axioms?"
StrQuestions(3) = "a)  What is the difference between a key and a super key? Is the " & Chr(13) & _
    "Intersection of     two super keys necessarily a key? Prove or disprove." & Chr(13) & _
"b) When are two sets of functional dependencies equivalent? How can we determine their equivalence?" & Chr(13) & _
"c) What is the difference between universal instance assumption and universal scheme assumption?" & Chr(13) & _
"d) What do you mean by the schedule of a set of transactions? What is a serializable schedule?" & Chr(13) & _
"e) What is meant by cascading abort of transactions? Givve sufficient conditions to avoid cascading abort of transactions?" & Chr(13) & _
"f) In a distributed database system what do you mean by the following terms: local autonomy, distribution transparency, and fragmentation transpiration and replication transparency."
StrQuestions(4) = "a)  When is a relation scheme in third normal form? Can there be any " & Chr(13) & _
    "Update anomaly in 3NF relation? If so, give an example." & Chr(13) & _
"b) What is meant by the canonical cover of set of functional dependencies? Is the canonical cover unique? If not, give a counterexample." & Chr(13) & _
"c) What is meant by the dependency preservation property of decomposition? Why is it not always possible to arrive at a BCNF decomposition that preserves the dependencies?" & Chr(13) & _
"d)  How does a query tree represent a relational algebra expression? Give the rules which allow selection operation to move as far down the tree as possible." & Chr(13) & _
"e)  When is a schedule said to be serializable? Why isit important? In a serializable schedule can a deadlock occur? If so, give an example." & Chr(13) & _
"f)  In a distributed database how and why are relations fragmented? How can a relation be composed from its fragments? How can a query be expressed in terms of fragments?"
StrQuestions(5) = "a)  Construct a B tree for the following set of key values " & Chr(13) & _
    "(3,9,11,13,26,35,41,50,52) under the assumption that the number" & Chr(13) & _
    "of search key vvalues that fit in one node is 3." & Chr(13) & _
    "b)  Show the steps involved in the following tasks:" & Chr(13) & _
        "(i) Find records with search key values of 35" & Chr(13) & _
        "(ii)    Insert 10" & Chr(13) & _
"(iii)   Delete 41" & Chr(13) & _
"(iv)    Insert 43"
StrQuestions(6) = "a)  Outline an algorithm for insertion of a record in a B+tree file." & Chr(13) & _
"b)  i)  A PART file with part # as key field includes records with the following part # values :" & Chr(13) & _
"23,65,37,60,46,92,48,71,56,59,18,21,10,74,15,20,69,75" & Chr(13) & _
"Suppose the search field values are inserted in the given order in a B+ tree of order p = 4, Show how the tree will expand and what the final tree would look like." & Chr(13) & _
"ii) Show the steps involved in the following queries:" & Chr(13) & _
    "* retrieve the record with part #= 21." & Chr(13) & _
    "* delete the record with part # = 10." & Chr(13) & _
    "* insert a part record with part # = 43." & Chr(13) & _
"c)  Suppose that for the part file in (b), a secondary index of an no ordering field 'Price' is to be maintained. Describe a suitable procedure for implementing such an index and outline a procedure for finding all parts having a given price."
StrQuestions(7) = "a)  Suppose that the PART file in 6(b) is to be stored as a hashed file " & Chr(13) & _
        "to support direct access. Assume that each bucket can hold three" & Chr(13) & _
        "records and there are 8 buckets." & Chr(13) & _
"i)  Briefly describe any three types of hash functions that may be used for this file. With appropriate justification, select a hash function and for the key values (Part #) as given in 6(b), estimate the maximum number of collisions." & Chr(13) & _
"ii) With the hash function thus selected, outline algorithms for retrieval and insertion of records from this file. Clearly explain collision resolution and overflow management strategies used." & Chr(13) & _
"b) Instead of allocating a fixed number of buckets to the join file, suppose the file is allowed to expand and shrink dynamically. Describe a suitable modification of the basic hashing technique to support dynamic file extension."
StrQuestions(8) = "a) Let F be a set of functional dependencies over R. Find a bound on the size of F+ in functional dependencies, in terms of number of attributes in R." & Chr(13) & _
    "b) Why is it necessary ton normalize relation schemes ? Illustrate with examples the updating problems that may arise in relation schemes which are in second and third normal forms. give an example of a relation in third normal form that has some prime attributes transitively dependent upon a key." & Chr(13) & _
    "c) While synthesizing relation schemes, when would the synthesized scheme adequately represent the input specification? Give the conditions to be satisfied and jstify their requirement."
StrQuestions(9) = "a) Use the definition of multivalued dependency to prove that the following axioms are sound." & Chr(13) & _
        "(i) the complementation rule." & Chr(13) & _
        "(ii) the multivalued transitivity rule." & Chr(13) & _
    "b) Prove that if a relation r over R(X,Y,Z) satisfies the multivalued dependent X ->-> Y, hen r decomposes lossless onto the relation schemes XY and YZ. Is the converse also true?"
StrQuestions(10) = "a) Let F be a set of FDs over R. Let X,Y, and Z be subset of R with Z = R = XY. Let XPLUS be the closure of X under F. Show that when Y and Z are not subsets of XPLUS , there exists a relation r over R that satisfies F and does not satisfy the multivalued dependency X->-> Y." & Chr(13) & _
    "b)  (i) Given a set of functional dependencies F over a set of attributes U. Design an algorithm that synthesizes relation schemes in 3NF and preserves the dependencies." & Chr(13) & _
        "(ii) Does your algorithm ensure lossless join of the synthesized schemes? If not suggest suitable steps to ensure lossless join."
End Sub

Private Sub questions_paper2()
StrQuestions(1) = "a) Given a set F of functions dependencies over U. Show that there exists a relation r over U(called Armstrong relation), such that for each FD f oiver Y, r satisfieds f , if and only if f is in the closure of F." & Chr(13) & _
    "b) Using reflexivity, augmentation and transitivity axioms, show that the set of DFDs F = {AB ->, A -> D, CD -> EG} implies the FD AB -> G. From this set of FDs synthesize relation schemes in BCNF. Does this decomposition have lossless join?" & Chr(13) & _
    "c) Prove or disprove the soundness of the following rules." & Chr(13) & _
    "(i) If X -> YZ, then X -> Y." & Chr(13) & _
    "(ii) If XW -> Y and XY -> Z, then X-> (Z - W)."
StrQuestions(2) = "a) What is meant by the dependency basis of a set of attributes with respect to a set of data dependencies?" & Chr(13) & _
    "b) Given a set of multivalued dependencies M<  over a relation scheme R and a subset of attributes X of R, design an algorithm for computing the dependency basis of X with respect to M." & Chr(13) & _
    "c) Determine the dependency basis of the attribute C of the relation R (CTHRSG) with data dependencies C ->-> HR and CS -> G."
StrQuestions(3) = "a) Consider a relation scheme R with a set of functional dependencies F. Show that F implies the MVD X ->-> Y, if and only if F in[lies the FD X-> Y, or F implies the FD X -> Z, where Z = R - XY." & Chr(13) & _
    "b) Let r be a relation over a scheme R, and suppose ht X,Y and Z are subsets of R, such that Z = R- XY. Prove that the relation r satisfied the multivalued dependency (MVD) X ->-> Y if r has lossless joint decomposition over schemes R1(SY) and R2 (XZ)." & Chr(13) & _
    "c) Show that Fourth Normal Form (4NF) implies 3NF."
StrQuestions(4) = "a) Let R be a relation scheme defined o n three attributes A, B and C. Determine sets of functional dependencies F such that (R,F) is in third normal form." & Chr(13) & _
    "b) Let R (A,B,C,D,Ef) be a relation scheme with functional dependencies F = {AB -> C, BE -> F}. determine whether F implies AC -> F." & Chr(13) & _
    "c) Show that the following inference action is not a single one:" & Chr(13) & _
        "If X ->-> Y and XY ->, Z then X->Z"
StrQuestions(5) = "a) Define lossless jong decomposition. Determine whether p(AB,CD) is a lossless join decomposition of R(ABCD) with" & Chr(13) & _
    "F = {A -> B, C-<>D}." & Chr(13) & _
    "b) Given a set of FDS F = {AB ->CD,D->E,EG->H,AG->BH,DH->A}. Are the following decompositions lossless with respect to F?" & Chr(13) & _
        "i) p ={R1 = (ABCD), R2(BDE), R3(AGH)}." & Chr(13) & _
        "ii) p = {R1(ACD),R2(BCEG),R3(EGH)}." & Chr(13) & _
    "c) Given a relational database schema R with a set of data dependencies D. While synthesizing normalized relation schemas, wh6y is it desirable that the final decomposition has lossless join?"
StrQuestions(6) = "a) Given a relational database consisting of the following relations about the students in a university." & Chr(13) & _
    "STUDENT(ROLL_NO,STUDENT_NAME,aDDRESS,AGE) DEPARTMENT(DEPT_NAME,COURSE#,TEACHER)TIME_TABLE(COUTSE#,SUBJECT_NAMESEMESTER,TIME,ROOM_NO)PERFORMANCE(ROLL_NO,COURSE#GRADE) Express the following query in SQL Get the names of students who have secured 'A' grade in a course offered by the Computer Science Department in the Autumn Semester." & Chr(13) & _
    "b) Consider a database having following relations. EMPS(EMP_NO,NAME, SLARY, DEPT_NO) DEPTS(DEPT_NO,DNAME,MANAGER)" & Chr(13) & _
        "(i) no employee earns more than Rs. 50000." & Chr(13) & _
        "(ii) No employee in DEPT_NO 50 earns more than 25000."
StrQuestions(7) = "a) Given a relational database consisting of following relations." & Chr(13) & _
    "EMPLOYEE(E_NO,EMP_NAME,AGE,SALARY)DEPARTMENT(DEPT_NAME,LOCATION,E_NO,MANAGER)PROJECTS(P_NO,PROJ_DETAILS,START_DATE,E-NO,PERIOD) Express the following query in SQL. Get the name of employees who are working in the testing department of a project for more than fifteen days(i.e. PERIOD > 15) and list the P_NO of such projects." & Chr(13) & _
    "b) i) Given a relational algebra expression involving selection, projection and join. Suggest an algorithm for constructing a table which can be treated as a conjunctive query of domain relational calculus that is equivalent to the given relational expression." & Chr(13) & _
       "ii) Grind a table which is equivalent to the query as given in the part(a) of this question"
StrQuestions(8) = "A) Show that SQL is a relationally complete language." & Chr(13) & _
    "b) What is trigger? Explain how triggers can be used to enforce integrity constraints in a database." & Chr(13) & _
    "c) Express the following integrity constraints using SQL." & Chr(13) & _
        "(i) In a relation EMPLOYEE (EMP_ID,EMP_NAME,DEPT,PERFORMANCE_RATING,SALARY),the salary of an employee can be updated only if his/her performance rating is greater than 0.6. Moreover, such an employee gets ten percent annual raise in salary." & Chr(13) & _
        "(ii) The relation R(X,Y,Z) satisfies the functional dependent X-> Y"
StrQuestions(9) = "Explain why transaction atomicity is one of the most important requirements for concurrency control. Describe how, in a multi-user environment, atomicity of read and write operations can be endured, especially for large data items i.e. when the granular with of data items is larger than a disk block."
StrQuestions(10) = "Describe the functional modules of a centralized database system responsible for concurrency control and recovery. Discuss how database read and write operations are implemented."
End Sub

Private Sub questions_paper3()
StrQuestions(1) = "Describe the functional modules of a centralized database system responsible for concurrency control and recovery. Discuss how database read and write operations are implemented."
StrQuestions(2) = "Consider a relational model of an airlines reservation system. Illustrate, with the help of suitable transactions operating on this database how a non serializable execution affects database consistency. In order to ensure consistency of the airlines reservation system what type of scheduler would you prefer? Justify your answer"
StrQuestions(3) = "a)  Give an algorithm to verify whether a schedule is conflict serializable" & Chr(13) & _
    "b)  Show that a static two phase locking schedule satisfies the conditions for dynamic two phase locking. Is he converse true?" & Chr(13) & _
    "c) What is phantom deadlock?" & Chr(13) & _
    "d) Suggest a suitable scheme for deadlock detection."
End Sub



