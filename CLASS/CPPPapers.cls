VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPPPapers"
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
 
Public Function CPP(PaperType As String, CPPPaper As Paper, mCat As Category, Number As Integer)
'ptype is the paper type
'CPPpaper is the number of paper as paper1,paper2 ...
'mcat is the category
'number is the number of questions, or options or answers to be taken

StrPaperType = PaperType
' StrPaper = Paper
If PaperType = "Fill in the Blanks" Then
   If CPPPaper = Paper1 Then Fill_in_the_Blanks_Paper1
   If CPPPaper = Paper2 Then Fill_in_the_Blanks_Paper2
   If CPPPaper = Paper3 Then Fill_in_the_Blanks_Paper3
   If CPPPaper = Paper4 Then Fill_in_the_Blanks_Paper4
   If CPPPaper = Paper5 Then Fill_in_the_Blanks_Paper5
   If CPPPaper = Paper6 Then Fill_in_the_Blanks_Paper6
   If CPPPaper = Paper7 Then Fill_in_the_Blanks_Paper7
   If CPPPaper = Paper8 Then Fill_in_the_Blanks_Paper8
   If CPPPaper = Paper9 Then Fill_in_the_Blanks_Paper9
   If CPPPaper = Paper10 Then Fill_in_the_Blanks_Paper10
   
ElseIf PaperType = "Matching Columns" Then
       If CPPPaper = Paper1 Then Matching_Columns_Paper1
       If CPPPaper = Paper2 Then Matching_Columns_Paper2
       If CPPPaper = Paper3 Then Matching_Columns_Paper3
       If CPPPaper = Paper4 Then Matching_Columns_Paper4
       If CPPPaper = Paper5 Then Matching_Columns_Paper5
       If CPPPaper = Paper6 Then Matching_Columns_Paper6
       If CPPPaper = Paper7 Then Matching_Columns_Paper7
       If CPPPaper = Paper8 Then Matching_Columns_Paper8
       If CPPPaper = Paper9 Then Matching_Columns_Paper9
       If CPPPaper = Paper10 Then Matching_Columns_Paper10
       
ElseIf PaperType = "True Or False" Then
       If CPPPaper = Paper1 Then TrueorFalse_Paper1
       If CPPPaper = Paper2 Then TrueorFalse_Paper2
       If CPPPaper = Paper3 Then TrueorFalse_Paper3
       If CPPPaper = Paper4 Then TrueorFalse_Paper4
       If CPPPaper = Paper5 Then TrueorFalse_Paper5
       If CPPPaper = Paper6 Then TrueorFalse_Paper6
       If CPPPaper = Paper7 Then TrueorFalse_Paper7
       If CPPPaper = Paper8 Then TrueorFalse_Paper8
       If CPPPaper = Paper9 Then TrueorFalse_Paper9
       If CPPPaper = Paper10 Then TrueorFalse_Paper10
       
ElseIf PaperType = "Descriptive Questions" Then
       If CPPPaper = Paper1 Then questions_paper1
       If CPPPaper = Paper2 Then questions_paper2
       If CPPPaper = Paper3 Then questions_paper3
       If CPPPaper = Paper4 Then questions_paper4
       If CPPPaper = Paper5 Then questions_paper5
       If CPPPaper = Paper6 Then questions_paper6
       If CPPPaper = Paper7 Then questions_paper7
       If CPPPaper = Paper8 Then questions_paper8
       If CPPPaper = Paper9 Then questions_paper9
       
ElseIf PaperType = "Multiple Choice" Then
       If CPPPaper = Paper1 Then multiplechoice_Paper1
       If CPPPaper = Paper2 Then multiplechoice_Paper2
       If CPPPaper = Paper3 Then multiplechoice_Paper3
       If CPPPaper = Paper4 Then multiplechoice_Paper4
       If CPPPaper = Paper5 Then multiplechoice_Paper5
       If CPPPaper = Paper6 Then multiplechoice_Paper6
       If CPPPaper = Paper7 Then multiplechoice_Paper7
       If CPPPaper = Paper8 Then multiplechoice_Paper8
       If CPPPaper = Paper9 Then multiplechoice_Paper9
       If CPPPaper = Paper10 Then multiplechoice_Paper10
             
End If

If mCat = Fill_Questions Then
   CPP = StrQuestions(Number)
ElseIf mCat = Fill_Options Then
       CPP = stroptions(Number)
ElseIf mCat = Fill_Answers Then
       CPP = stranswers(Number)
       
ElseIf mCat = Matching_ColumnA Then
       CPP = strcolumnA(Number)
ElseIf mCat = Matching_ColumnB Then
       CPP = strcolumnB(Number)
ElseIf mCat = Matching_Answers Then
       CPP = stranswers(Number)
       
ElseIf mCat = Questions Then
       CPP = StrQuestions(Number)
       
ElseIf mCat = TF_Questions Then
       CPP = StrQuestions(Number)
ElseIf mCat = TF_Answers Then
       CPP = stranswers(Number)
       
ElseIf mCat = MC_Questions Then
       CPP = strmultiple(Number)
       
ElseIf mCat = MC_Options_1 Then
       CPP = strmuloption(Number)
       
ElseIf mCat = MC_Options_2 Then
       CPP = strmuloption(Number)
       
ElseIf mCat = MC_Options_3 Then
       CPP = strmuloption(Number)
       
ElseIf mCat = MC_Options_4 Then
       CPP = strmuloption(Number)
       
ElseIf mCat = MC_Answer Then
       CPP = strmulanswer(Number)
       
End If

End Function

Private Sub Fill_in_the_Blanks_Paper1()
StrQuestions(1) = "The two major components of an object are __________ and functions that __________"
StrQuestions(2) = "In C++ a function contained within a class is called __________."
StrQuestions(3) = "Protecting data from access by unauthorized functions is called __________."
StrQuestions(4) = "The name of a destructor is same as __________."
StrQuestions(5) = "Bundling data and functions together is called __________."
StrQuestions(6) = "When a language has the capability to produce new data type, it is said to be __________."
StrQuestions(7) = "The ability of a function or operator to act in different ways on different data types is called __________."
StrQuestions(8) = "A normal C+= operator that acts in special ways on newly defined data types is said to be __________."
StrQuestions(9) = "A constructor which is used to declare and initialize an object from another object is __________."
StrQuestions(10) = "A function name must be followed by __________."

stroptions(1) = "encapsulation"
stroptions(2) = "Operator overloading"
stroptions(3) = "Its arguments"
stroptions(4) = "Extendibility"
stroptions(5) = "Polymorphism"
stroptions(6) = "Data hiding"
stroptions(7) = "Member function "
stroptions(8) = "Data, manipulate data"
stroptions(9) = "Copy constructor"
stroptions(10) = "The class"

stranswers(1) = "Data, manipulate data"
stranswers(2) = "Member function "
stranswers(3) = "Data hiding"
stranswers(4) = "The class"
stranswers(5) = "encapsulation"
stranswers(6) = "Extendibility"
stranswers(7) = "Polymorphism"
stranswers(8) = "Operator overloading"
stranswers(9) = "Data hiding"
stranswers(10) = "Its arguments"

End Sub

Private Sub Fill_in_the_Blanks_Paper2()
StrQuestions(1) = "A function body is delimited by dash."
StrQuestions(2) = "A C++ instruction that tells the computer to do something is called a dash."
StrQuestions(3) = "Accessing data type brings together a group of dash."
StrQuestions(4) = "An enumerated data type brings together a group of dash."
StrQuestions(5) = "The first three members of an enum type are normally represented by the values dash, dash and dash."
StrQuestions(6) = "Overload functions dash."
StrQuestions(7) = "In general, an inline function executes dash than a normal function , but requires dash memory."
StrQuestions(8) = "A default argument has a value that dash. Storage class is concerned with the dash and dash of a variable."
StrQuestions(9) = "Storage class is concerned with the dash and dash of a variable."
StrQuestions(10) = "The i/o function used to specify the required field size for displaying an output value is dash."

stroptions(1) = "class"
stroptions(2) = "0,1,2"
stroptions(3) = "{}"
stroptions(4) = "statement"
stroptions(5) = "slower, less"
stroptions(6) = "class, type"
stroptions(7) = "list items into integer values"
stroptions(8) = "have the same name/identifier"
stroptions(9) = "is defined in the prototype"
stroptions(10) = "width()"

stranswers(1) = "{}"
stranswers(2) = "statement"
stranswers(3) = "class"
stranswers(4) = "list items into integer values"
stranswers(5) = "0,1,2"
stranswers(6) = "have the same name/identifier"
stranswers(7) = "slower, less"
stranswers(8) = "is defined in the prototype"
stranswers(9) = "class, type"
stranswers(10) = "width()"

End Sub

Private Sub Fill_in_the_Blanks_Paper3()
StrQuestions(1) = "Register variables are internal auto variables that are to be stored in dash registers."
StrQuestions(2) = "In a class specifier, at or function designated private are accessible dash."
StrQuestions(3) = "The dot operator (or class member access operator) is a class member and dash."
StrQuestions(4) = "Member functions defined inside a class specifier are dash by default."
StrQuestions(5) = "A constructor name is the same as dash."
StrQuestions(6) = "Register variables are internal auto variables that are to be stored in dash registers."
StrQuestions(7) = "The only technical difference between structure and classes in C++ is that dash."
StrQuestions(8) = "Sending a message to an object is the same as dash ."
StrQuestions(9) = "All the elements in an array must be the dash data type."
StrQuestions(10) = "The elements of a ten element array are numbered from  dash to dash."

stroptions(1) = "is private"
stroptions(2) = "class"
stroptions(3) = "same"
stroptions(4) = "1,10"
stroptions(5) = "calling a member function to the object response"
stroptions(6) = "private"
stroptions(7) = "structure members are public by default whereas classes are private"
stroptions(8) = "only to the members of the class"
stroptions(9) = "hardware "
stroptions(10) = "width ()"

stranswers(1) = "width ()"
stranswers(2) = "only to the members of the class"
stranswers(3) = "is private"
stranswers(4) = "calling a member function to the object response"
stranswers(5) = "class"
stranswers(6) = "hardware "
stranswers(7) = "structure members are public by default whereas classes are private"
stranswers(8) = "calling a member function to the object response"
stranswers(9) = "same"
stranswers(10) = "1,10"

End Sub
Private Sub Fill_in_the_Blanks_Paper4()
StrQuestions(1) = "dash variables may be altered by some means unknown to the program."
StrQuestions(2) = "An array name, used in the source file, represents the dash of the array."
StrQuestions(3) = "A string in C++ is an dash of type dash ."
StrQuestions(4) = "An address is a dash , while a pointer is a dash."
StrQuestions(5) = "One way pointer are useful is to refer to a memory address that has no dash."
StrQuestions(6) = "An asterisk placed after a data type means dash."
StrQuestions(7) = "Escape sequences are used to denote characters that dash from the keyboard."
StrQuestions(8) = "Trigraphs can be used where dash are not a available on the key board."
StrQuestions(9) = "Of the three ways to pass arguments to functions, only passing by dash and passing by dash allow the function to modify the argument in the calling program."
StrQuestions(10) = "Using new may result in less dash memory than using an array."

stroptions(1) = "class"
stroptions(2) = "array, char"
stroptions(3) = "it is of pointer type"
stroptions(4) = "reference"
stroptions(5) = "free"
stroptions(6) = "value, reference"
stroptions(7) = "constant, var"
stroptions(8) = "volatile"
stroptions(9) = "cannot be directly"
stroptions(10) = "Special Symbols"

stranswers(1) = "volatile"
stranswers(2) = "class"
stranswers(3) = "array, char"
stranswers(4) = "constant, var"
stranswers(5) = "reference"
stranswers(6) = "it is of pointer type"
stranswers(7) = "cannot be directly"
stranswers(8) = "Special Symbols"
stranswers(9) = "value, reference"
stranswers(10) = "free"

End Sub

Private Sub Fill_in_the_Blanks_Paper5()
StrQuestions(1) = "The delete operator returns dash to the operating system."
StrQuestions(2) = "main() returns a value of type dash to the operating system."
StrQuestions(3) = "dash function reduces function call overload."
StrQuestions(4) = "If no appropriate handler is available, the dash function is invoked."
StrQuestions(5) = "dash functions have access to the private and protected members of a class."
StrQuestions(6) = "The member functions defined within the class definition are implicitly dash."
StrQuestions(7) = "A default constructor can have dash arguments."
StrQuestions(8) = "A static member function can access only dash from the class."
StrQuestions(9) = "The prototype of a function is dash before its definition."
StrQuestions(10) = "Exception handling is a method of dash control and information to an associated exception handler."

stroptions(1) = "inline"
stroptions(2) = "friend"
stroptions(3) = "void"
stroptions(4) = "free space"
stroptions(5) = "inline"
stroptions(6) = "functional declaration"
stroptions(7) = "transferring "
stroptions(8) = "default"
stroptions(9) = "unexpected()"
stroptions(10) = "static member data"

stranswers(1) = "free space"
stranswers(2) = "void"
stranswers(3) = "inline"
stranswers(4) = "unexpected()"
stranswers(5) = "friend"
stranswers(6) = "inline"
stranswers(7) = "default"
stranswers(8) = "static member data"
stranswers(9) = "functional declaration"
stranswers(10) = "friend"

End Sub

Private Sub Fill_in_the_Blanks_Paper6()
StrQuestions(1) = "A dash object has its storage assigned before the program starts up and is initialized only  once."
StrQuestions(2) = "In C++ when the compiler encounters a string constant, it stores in the programs dash."
StrQuestions(3) = "A dash parameter automatically receives the address of its corresponding argument."
StrQuestions(4) = "A reference parameter is similar to a dash parameter is Pascal."
StrQuestions(5) = "The mechanism that allows two related functions to share the same name is called dash."
StrQuestions(6) = "dash is a compile time operator that obtains size of a type or variable. "
StrQuestions(7) = "In C++, structures are called dash data types."
StrQuestions(8) = "dash is a pointer to the object that invokes a member."
StrQuestions(9) = "dash functions can be used to overload an operator."
StrQuestions(10) = "When a base class is inherited as public, public members of the base class become dash members of the derived "

stroptions(1) = "reference"
stroptions(2) = "member"
stroptions(3) = "VAR"
stroptions(4) = "Static"
stroptions(5) = "Function overloading"
stroptions(6) = "String table"
stroptions(7) = "Sizeof"
stroptions(8) = "Public"
stroptions(9) = "This"
stroptions(10) = "Aggregate"

stranswers(1) = "Static"
stranswers(2) = "String table"
stranswers(3) = "reference"
stranswers(4) = "VAR"
stranswers(5) = "Function overloading"
stranswers(6) = "Sizeof"
stranswers(7) = "Aggregate"
stranswers(8) = "This"
stranswers(9) = "member"
stranswers(10) = "Public"

End Sub

Private Sub Fill_in_the_Blanks_Paper7()
StrQuestions(1) = "The protected specifier allows to create a class member that is accessible within a class hierarchy , but is otherwise dash."
StrQuestions(2) = "When a virtual function is redefined in a derived class, it  is said to be dash."
StrQuestions(3) = "dash binds a function call a compile time."
StrQuestions(4) = "Every non persistent object has a dash."
StrQuestions(5) = "dash are free store operators."
StrQuestions(6) = "A dash expression is a basic expression from which other expressions are built."
StrQuestions(7) = "The only expression that takes three operators is knows as dash."
StrQuestions(8) = "A dash assignment combines a binary operator with simple assignment operator."
StrQuestions(9) = "The dash declaration specifies that the object can be used between separate translation units of a program."
StrQuestions(10) = "In C++, arrays of functions are dash."

stroptions(1) = "unique address"
stroptions(2) = "overridden"
stroptions(3) = "new and delete"
stroptions(4) = "early binding"
stroptions(5) = "private"
stroptions(6) = "primer"
stroptions(7) = "compound"
stroptions(8) = "not allowed"
stroptions(9) = "extern"
stroptions(10) = "conditional"

stranswers(1) = "private"
stranswers(2) = "overridden"
stranswers(3) = "early binding"
stranswers(4) = "unique address"
stranswers(5) = "new and delete"
stranswers(6) = "primer"
stranswers(7) = "compound"
stranswers(8) = "extern"
stranswers(9) = "not allowed"
stranswers(10) = "conditional"

End Sub

Private Sub Fill_in_the_Blanks_Paper8()
StrQuestions(1) = "The indirection operator is dash."
StrQuestions(2) = "Objects that exist between functions must be defined as dash or should have file scope, otherwise the object not longer exists when the function terminates."
StrQuestions(3) = "A function contained in the current program is dash function."
StrQuestions(4) = "The dash function is expanded during compile time to avoid invocation overhead."
StrQuestions(5) = "If the result returned by a function is not required, the dash can be used explicitly to specify this."
StrQuestions(6) = "A function dash is a mandatory statement that specifies the number and form of the parameters received by the function and return type."
StrQuestions(7) = "In C++, arrays  are passed to the functions by dash."
StrQuestions(8) = "The individual array elements are passed by dash."
StrQuestions(9) = "dash functions are used when syntax requires a function, but  no processing is required to the performed."
StrQuestions(10) = "dash function must be available at compile time."

stroptions(1) = "*"
stroptions(2) = "reference"
stroptions(3) = "internal"
stroptions(4) = "inline"
stroptions(5) = "Null"
stroptions(6) = "Static"
stroptions(7) = "Inline"
stroptions(8) = "Prototype"
stroptions(9) = "Value"
stroptions(10) = "Cast"

stranswers(1) = "*"
stranswers(2) = "Static"
stranswers(3) = "internal"
stranswers(4) = "inline"
stranswers(5) = "Cast"
stranswers(6) = "Prototype"
stranswers(7) = "reference"
stranswers(8) = "Value"
stranswers(9) = "Null"
stranswers(10) = "Inline"

End Sub

Private Sub Fill_in_the_Blanks_Paper9()
StrQuestions(1) = "The non class function that displays the contents of the data are in hexadecimalformat is dash."
StrQuestions(2) = "A structure may be dynamically created (at run time) using the dash operator."
StrQuestions(3) = "The structure pointer operator is dash."
StrQuestions(4) = "A union declaration with neither a name not an identifier is called dash union."
StrQuestions(5) = "A global anonymous union must be declared as dash."
StrQuestions(6) = "Them member names of an anonymous union must be dash within thescope."
StrQuestions(7) = "The dash pointer is normally used to indicate that the pointer does not contain an active entry."
StrQuestions(8) = "A C++ void pointer can be assigned only to another dash."
StrQuestions(9) = "One of the disadvantages of reference variable is the function invvokation dash that references are being used."
StrQuestions(10) = "Reference can not be changed but it an only be dash."

stroptions(1) = "void pointer"
stroptions(2) = "zero"
stroptions(3) = "does not indicate"
stroptions(4) = "static"
stroptions(5) = "->"
stroptions(6) = "unique"
stroptions(7) = "printf()"
stroptions(8) = "anonymous"
stroptions(9) = "new"
stroptions(10) = "initialized"

stranswers(1) = "printf()"
stranswers(2) = "new"
stranswers(3) = "->"
stranswers(4) = "anonymous"
stranswers(5) = "static"
stranswers(6) = "unique"
stranswers(7) = "zero"
stranswers(8) = "void pointer"
stranswers(9) = "does not indicate"
stranswers(10) = "initialized"

End Sub

Private Sub Fill_in_the_Blanks_Paper10()
StrQuestions(1) = "Classes may themselves be defined with in classes and these are called dash."
StrQuestions(2) = "A dash is a function that is invoked to perform initialization processing when an object is created."
StrQuestions(3) = "The name of the constructor is same as that of the dash it belongs"
StrQuestions(4) = "The dash constructor is a constructor that can be called without any arguments."
StrQuestions(5) = "A default constructor and copy constructor are generated if the are dash."
StrQuestions(6) = "A constructor can not be declared as dash."
StrQuestions(7) = "A constructor can not be explicitly called. One way to overcome this is to have the constructor dash which can then be explicitly called."
StrQuestions(8) = "Shallow copying is usually not sufficient when the object contains dash and it has destructor."
StrQuestions(9) = "dash specifies the processing to be performed before the actual constructor processing starts."
StrQuestions(10) = "dash constructor is invoked to convert an object of that argument type to the class type."

stroptions(1) = "static"
stroptions(2) = "pointers"
stroptions(3) = "class"
stroptions(4) = "conversion"
stroptions(5) = "nested classes"
stroptions(6) = "constructor initializer"
stroptions(7) = "defined"
stroptions(8) = "constructor initlalizer"
stroptions(9) = "invoke a member function"
stroptions(10) = "constructor"

stranswers(1) = "nested classes"
stranswers(2) = "constructor"
stranswers(3) = "class"
stranswers(4) = "constructor initializer"
stranswers(5) = "defined"
stranswers(6) = "static"
stranswers(7) = "invoke a member function"
stranswers(8) = "pointers"
stranswers(9) = "constructor initlalizer"
stranswers(10) = "conversion"

End Sub

Private Sub Matching_Columns_Paper1()
strcolumnA(1) = "class"
strcolumnA(2) = "object"
strcolumnA(3) = "method"
strcolumnA(4) = "derived class"
strcolumnA(5) = "abstract class"
strcolumnA(6) = "constructor"
strcolumnA(7) = "destructor"
strcolumnA(8) = "private"
strcolumnA(9) = "protected"
strcolumnA(10) = "public"

strcolumnB(1) = "a base class which has a member that is a pure virtual function."
strcolumnB(2) = "A definition of an object which contains members"
strcolumnB(3) = "And access privileges of these members"
strcolumnB(4) = "A function prototype within the declaration of a class."
strcolumnB(5) = "Inherits characteristics from one or more base class"
strcolumnB(6) = "Automatically initializes objects when they are created"
strcolumnB(7) = "Allows he encapsulation of functions within that class and derived classes only"
strcolumnB(8) = "Is used to  deallocate memory"
strcolumnB(9) = "Allows the encapsulation of functions within that class alone"
strcolumnB(10) = "Allows the accessibility of data elements and functions to derived classes and the remainder of the application"

stranswers(1) = "A definition of an object which contains members"
stranswers(2) = "And access privileges of these members"
stranswers(3) = "A function prototype within the declaration of a class."
stranswers(4) = "Inherits characteristics from one or more base class"
stranswers(5) = "a base class which has a member that is a pure virtual function."
stranswers(6) = "Automatically initializes objects when they are created"
stranswers(7) = "Is used to  deallocate memory"
stranswers(8) = "Allows the encapsulation of functions within that class alone"
stranswers(9) = "Allows he encapsulation of functions within that class and derived classes only"
stranswers(10) = "Allows the accessibility of data elements and functions to derived classes and the remainder of the application"

End Sub
Private Sub Matching_Columns_Paper2()
strcolumnA(1) = "int *ptr"
strcolumnA(2) = "char *const ptr"
strcolumnA(3) = "char const *ptr"
strcolumnA(4) = "prt = &A"
strcolumnA(5) = "ptr = A"
strcolumnA(6) = "cin"
strcolumnA(7) = "cout"
strcolumnA(8) = "cerr"
strcolumnA(9) = "clog"
strcolumnA(10) = "<<&>>"

strcolumnB(1) = "ptr is set to address of A"
strcolumnB(2) = "declare pointer to integer"
strcolumnB(3) = "ptr contains address of first element of array A(if A is defined as an array)"
strcolumnB(4) = "ptr is a pointer to a constant"
strcolumnB(5) = "ptr is a constant to a character"
strcolumnB(6) = "is the standard operator overloaded in header file iostream.h"
strcolumnB(7) = "is the standard error stream, unbuffered"
strcolumnB(8) = "is the standard input stream"
strcolumnB(9) = "is the standard output stream"
strcolumnB(10) = "is the standard error, buffered."

stranswers(1) = "declare pointer to integer"
stranswers(2) = "ptr is a pointer to a constant"
stranswers(3) = "ptr is a constant to a character"
stranswers(4) = "ptr is set to address of A"
stranswers(5) = "ptr contains address of first element of array A(if A is defined as an array)"
stranswers(6) = "is the standard input stream"
stranswers(7) = "is the standard output stream"
stranswers(8) = "is the standard error stream, unbuffered"
stranswers(9) = "is the standard operator overloaded in header file iostream.h"
stranswers(10) = "is the standard error stream, unbuffered"

End Sub
Private Sub Matching_Columns_Paper3()
strcolumnA(1) = "nline function"
strcolumnA(2) = "friend function"
strcolumnA(3) = "virtualfunction"
strcolumnA(4) = "member function"
strcolumnA(5) = "overload"
strcolumnA(6) = "new"
strcolumnA(7) = "delete"
strcolumnA(8) = "this"
strcolumnA(9) = "scope"
strcolumnA(10) = "const"

strcolumnB(1) = "a member within  a class which is  a function prototype within the declaration of a class"
strcolumnB(2) = "a member function or  method that results in a reduction of processing and function call overload of the program."
strcolumnB(3) = "Not a member of one class but has full access to the provate and protected members of that class"
strcolumnB(4) = "Allows functions that perform similar tasks but work with different data types of objects to have the same name"
strcolumnB(5) = "A group of functions which have function identical names but different inpementation"
strcolumnB(6) = "Is an invisible identifier which automatically points to the object for which it exists"
strcolumnB(7) = "Deallocates memory"
strcolumnB(8) = "Allocates memory dynamically"
strcolumnB(9) = "Indicates that value of the variable cannot be changed"
strcolumnB(10) = "Is related to only for a block or any locks within that block in which the variable appears"

stranswers(1) = "a member function or  method that results in a reduction of processing and function call overload of the program."
stranswers(2) = "Not a member of one class but has full access to the provate and protected members of that class"
stranswers(3) = "A group of functions which have function identical names but different inpementation"
stranswers(4) = "a member within  a class which is  a function prototype within the declaration of a class"
stranswers(5) = "Allows functions that perform similar tasks but work with different data types of objects to have the same name"
stranswers(6) = "Allocates memory dynamically"
stranswers(7) = "Deallocates memory"
stranswers(8) = "Is an invisible identifier which automatically points to the object for which it exists"
stranswers(9) = "Is related to only for a block or any locks within that block in which the variable appears"
stranswers(10) = "Indicates that value of the variable cannot be changed"

End Sub
Private Sub Matching_Columns_Paper4()
strcolumnA(1) = "Polymorphism"
strcolumnA(2) = "encapsulation"
strcolumnA(3) = "multiple"
strcolumnA(4) = "object oriented"
strcolumnA(5) = "operator"
strcolumnA(6) = "function name(add) matches"
strcolumnA(7) = "return value(type voide)"
strcolumnA(8) = "first argument s1 matches"
strcolumnA(9) = "second argument s2 match3s"
strcolumnA(10) = "second argument s2 matches"

strcolumnB(1) = "a feature which allows the meaning of most unary and binary operators to be changed"
strcolumnB(2) = "the feature by which one object inherits data and functions from or more base classes"
strcolumnB(3) = "is the property by which virtual inheritance functions are invoked using pointers to base classes"
strcolumnB(4) = "allows the collection of individual programming blocks of code into hierarchies of classes"
strcolumnB(5) = "the mechanism by which data  and overloading functions that are allowed to manipulate this data are bound together within an object definition"
strcolumnB(6) = "argument(s)"
strcolumnB(7) = "object of which matches operator is a member"
strcolumnB(8) = "operator(+)"
strcolumnB(9) = "return value (type string)"
strcolumnB(10) = "no match for this member"

stranswers(1) = "is the property by which virtual inheritance functions are invoked using pointers to base classes"
stranswers(2) = "the mechanism by which data  and overloading functions that are allowed to manipulate this data are bound together within an object definition"
stranswers(3) = "the feature by which one object inherits data and functions from or more base classes"
stranswers(4) = "allows the collection of individual programming blocks of code into hierarchies of classes"
stranswers(5) = "a feature which allows the meaning of most unary and binary operators to be changed"
stranswers(6) = "operator(+)"
stranswers(7) = "return value (type string)"
stranswers(8) = "argument(s)"
stranswers(9) = "no match for this member"
stranswers(10) = "object of which matches operator is a member"
End Sub

Private Sub Matching_Columns_Paper5()
strcolumnA(1) = "automatic"
strcolumnA(2) = "extern"
strcolumnA(3) = "static"
strcolumnA(4) = "register"
strcolumnA(5) = "instance"
strcolumnA(6) = "data hiding"
strcolumnA(7) = "inheritance"
strcolumnA(8) = "early binding"
strcolumnA(9) = "late binding"
strcolumnA(10) = "binding"

strcolumnB(1) = "the existence of predefined object"
strcolumnB(2) = "use of this reduces the size of program and improves the performance"
strcolumnB(3) = "storage class for variables which are by default"
strcolumnB(4) = "storage class fro which values persist even when the variables go out of scope"
strcolumnB(5) = "global can exist outside or inside functions"
strcolumnB(6) = "associated at the time of program compilation"
strcolumnB(7) = "the feature by which one class derives all data members and function definitions from one or more other class functions of two or more base classes"
strcolumnB(8) = "associated at the time of program run"
strcolumnB(9) = "occurrence of   event"
strcolumnB(10) = "data for members declared private"

stranswers(1) = "storage class for variables which are by default"
stranswers(2) = "global can exist outside or inside functions"
stranswers(3) = "storage class fro which values persist even when the variables go out of scope"
stranswers(4) = "use of this reduces the size of program and improves the performance"
stranswers(5) = "the existence of predefined object"
stranswers(6) = "data for members declared private"
stranswers(7) = "associated at the time of program run"
stranswers(8) = "associated at the time of program compilation"
stranswers(9) = "occurrence of   event"
stranswers(10) = "associated at the time of program run"
End Sub
Private Sub Matching_Columns_Paper6()
strcolumnA(1) = "functional declaration"
strcolumnA(2) = "the main advantage of orthogonality in a programming language ]"
strcolumnA(3) = "reference variable is nothing but"
strcolumnA(4) = "a memory location that is shared by two or more different variables"
strcolumnA(5) = "the scope resolution operator qualifies"
strcolumnA(6) = "a function which has to have same name as the class"
strcolumnA(7) = "declares a class type name prior to the definition of the class"
strcolumnA(8) = "operators can be overloaded using the function"
strcolumnA(9) = "a function contained in an external program is"
strcolumnA(10) = "a function that does not perform any processing and does not return any value"

strcolumnB(1) = "forward reference before its definition"
strcolumnB(2) = "null"
strcolumnB(3) = "a member name with class"
strcolumnB(4) = "external"
strcolumnB(5) = "alias"
strcolumnB(6) = "prototype"
strcolumnB(7) = "it enables compact code"
strcolumnB(8) = "union"
strcolumnB(9) = "constructor"
strcolumnB(10) = "operator"

stranswers(1) = "prototype"
stranswers(2) = "it enables compact code"
stranswers(3) = "alias"
stranswers(4) = "union"
stranswers(5) = "a member name with class"
stranswers(6) = "constructor"
stranswers(7) = "forward reference before its definition"
stranswers(8) = "operator"
stranswers(9) = "external"
stranswers(10) = "null"

End Sub
Private Sub Matching_Columns_Paper7()
strcolumnA(1) = "members of a class"
strcolumnA(2) = "a private member function can only be called"
strcolumnA(3) = "a member function which does not alter any data in the class"
strcolumnA(4) = "in C++ the equivalent of malloc() of C is"
strcolumnA(5) = "the syntax rules for overloaded operators are same as"
strcolumnA(6) = "the power of reusability of existing code in c++ comes from"
strcolumnA(7) = "the property that specifies the features of the base class are privately derived or publicly derived"
strcolumnA(8) = "if A is a base class for the derived class B; in turn a base class for the derived class c and C is a base class for the derived class d, then the chain ABCD is"
strcolumnA(9) = "If A is a base class for the derived class B an C is derived from a base class B & another base class K, then such inheritance is"
strcolumnA(10) = "The duplication of inherited members due to multiple paths can be avoided by making a common base class as"

strcolumnB(1) = "visibility mode"
strcolumnB(2) = "original operators"
strcolumnB(3) = "virtual base class"
strcolumnB(4) = "inheritance"
strcolumnB(5) = "inheritance path"
strcolumnB(6) = "are private by default"
strcolumnB(7) = "hybrid inheritance"
strcolumnB(8) = "by another member function B is of its class"
strcolumnB(9) = "new"
strcolumnB(10) = "constant member function"

stranswers(1) = "are private by default"
stranswers(2) = "by another member function B is of its class"
stranswers(3) = "constant member function"
stranswers(4) = "new"
stranswers(5) = "original operators"
stranswers(6) = "inheritance"
stranswers(7) = "visibility mode"
stranswers(8) = "inheritance path"
stranswers(9) = "hybrid inheritance"
stranswers(10) = "virtual base class"

End Sub
Private Sub Matching_Columns_Paper8()
strcolumnA(1) = "The interface that is device independent supplied bu the I/O system for accessing various devices is"
strcolumnA(2) = "the base class for I/O stream"
strcolumnA(3) = "the inheritance which allow to combine features of several existing classes"
strcolumnA(4) = "the stream which inherits properties from ios, istream and ostream"
strcolumnA(5) = "setting of various formatting flags can be done by"
strcolumnA(6) = "the set of functions which are used for setting the output formats are called"
strcolumnA(7) = "to access the manipulators the include file is "
strcolumnA(8) = "the basic file to be included in most of the programs"
strcolumnA(9) = "formatted flags are"
strcolumnA(10) = "the file stream that provides input operations"

strcolumnB(1) = "iostream.h"
strcolumnB(2) = "iostream"
strcolumnB(3) = "manipulators"
strcolumnB(4) = "ios"
strcolumnB(5) = "stream"
strcolumnB(6) = "ifstream"
strcolumnB(7) = "overloaded setf() function"
strcolumnB(8) = "enumerated type"
strcolumnB(9) = "ultiple inheritance"
strcolumnB(10) = "iomanip.h"

stranswers(1) = "stream"
stranswers(2) = "ios"
stranswers(3) = "stream"
stranswers(4) = "iostream"
stranswers(5) = "overloaded setf() function"
stranswers(6) = "manipulators"
stranswers(7) = "iomanip.h"
stranswers(8) = "iostream.h"
stranswers(9) = "enumerated type"
stranswers(10) = "ifstream"

End Sub
Private Sub Matching_Columns_Paper9()
strcolumnA(1) = "pointers to objects of base class are type compatible with pointers to"
strcolumnA(2) = "the correct format of the following function could be int f(int count int max = 10, int sum)"
strcolumnA(3) = "the correct format of the following function could be int f(int coung int max, intsum = count)"
strcolumnA(4) = "Simple assignment of an object to a variable produces"
strcolumnA(5) = "Hollow copy creates problems especially when different"
strcolumnA(6) = "a single pointer variable can be made to point to"
strcolumnA(7) = "the THIS pointer acts as an"
strcolumnA(8) = "the major application of 'this' pointer is"
strcolumnA(9) = "the concept of polymorphism is implemented using"
strcolumnA(10) = "the overloaded member functions are selected for invoking by matching arguments. This information is known to the system at"

strcolumnB(1) = "pointers are used"
strcolumnB(2) = "to return the object it points to"
strcolumnB(3) = "overloaded functions"
strcolumnB(4) = "int f(int count int sum int max = 10)"
strcolumnB(5) = "objects belonging to derived classes"
strcolumnB(6) = "hollow copy"
strcolumnB(7) = "int f(int count int max int sum =0)"
strcolumnB(8) = "compile time"
strcolumnB(9) = "implicit argument to all the member functions"
strcolumnB(10) = "objects of derived class"

stranswers(1) = "objects of derived class"
stranswers(2) = "int f(int count int sum int max = 10)"
stranswers(3) = "int f(int count int max int sum =0)"
stranswers(4) = "hollow copy"
stranswers(5) = "pointers are used"
stranswers(6) = "objects belonging to derived classes"
stranswers(7) = "implicit argument to all the member functions"
stranswers(8) = "to return the object it points to"
stranswers(9) = "overloaded functions"
stranswers(10) = "compile time"
End Sub
Private Sub Matching_Columns_Paper10()
strcolumnA(1) = "the availability of source code at compile time is called"
strcolumnA(2) = "the inclusion of code in the case o inline function is"
strcolumnA(3) = "the inclusion of code in case of macros is"
strcolumnA(4) = "late binding is also known as"
strcolumnA(5) = "the overloaded member functions are selected for invoking by matching arguments. This information is known to the system at compile time. This type of polymorphism is called"
strcolumnA(6) = "when a function is made virtual, C++ determines which function to use at run time based on"
strcolumnA(7) = "virtual functions can not be"
strcolumnA(8) = "a virtual function in the base class must be defined even though it may not be"
strcolumnA(9) = "a virtual constructor"
strcolumnA(10) = "in the case of virtual functions while a base pointer can point to any type of the derived object the reverse"

strcolumnB(1) = "early binding "
strcolumnB(2) = "late biding"
strcolumnB(3) = "compile time polymorphism"
strcolumnB(4) = "the type of object pointed to the base pointer rather than the type of the pointer"
strcolumnB(5) = "early binding or static binding"
strcolumnB(6) = "can not be defined"
strcolumnB(7) = "used"
strcolumnB(8) = "is not true"
strcolumnB(9) = "dynamic binding"
strcolumnB(10) = "static member functions"

stranswers(1) = "early binding or static binding"
stranswers(2) = "early binding "
stranswers(3) = "compile time polymorphism"
stranswers(4) = "dynamic binding"
stranswers(5) = "compile time polymorphism"
stranswers(6) = "the type of object pointed to the base pointer rather than the type of the pointer"
stranswers(7) = "static member functions"
stranswers(8) = "used"
stranswers(9) = "can not be defined"
stranswers(10) = "is not true"
End Sub

Private Sub questions_paper1()
StrQuestions(1) = "a)  What do you mean by object oriented programming? How does it differ" & Chr(13) & _
    "from the ordinary programming concepts?" & Chr(13) & _
    "b)  What do you mean by function prototype and function overloading?" & Chr(13) & _
    "c)  What is the function of scope, resolution operator? Discuss with an" & Chr(13) & _
             "example" & Chr(13) & _
    "d)  what is the advantage of using the keyword in line before faction? What" & Chr(13) & _
            "types of functions are best suited to be inline and what are not ?" & Chr(13) & _
"e)  What do you mean by a virtual function? What is the difference between a virtual function and an overloaded function?"
StrQuestions(2) = "a)  What is the difference between the following two declarations :" & Chr(13) & _
        "char *const ptr2;" & Chr(13) & _
        "char const *ptr2;" & Chr(13) & _
    "b)  What will be the output after compiling the following code segment?" & Chr(13) & _
        "#include <iostream.h>" & Chr(13) & _
        "Class A" & Chr(13) & _
            "{" & Chr(13) & _
                "public :" & Chr(13) & _
                "A() {cout << constructing A \n;}" & Chr(13) & _
                "-A() {cout << Destructing A \n;}" & Chr(13) & _
            "};" & Chr(13) & _
        "Class b" & Chr(13) & _
"{public : B() {cout << constructing B \n;} -B() {cout << destructing B\n;}" & Chr(13) & _
"};" & Chr(13) & _
        "class C: public A, pubic B" & Chr(13) & _
            "{public : C() {cout << constructing C\n;} -C() {cout << destruct            ting C\n;}" & Chr(13) & _
        "main();" & Chr(13) & _
        "{C obj; return O;" & Chr(13) & _
        "}"
StrQuestions(3) = "a)  Explain what protected means. When one class inherits" & Chr(13) & _
"another when are the classes constructor called? When are their destructors called?" & Chr(13) & _
    "b)  Write code in C++ to create a manipulator called setup() that sets the field width to 10, he precision to 4, and the fill character to 2." & Chr(13) & _
    "c)  Why does the insertion operator (<<) works for the different built in types? Write C++ code for overloaded + add operator which actually subtract the right element from the left (c =a+b means b is subtracted from  a and kept in c)" & Chr(13) & _
    "d)  What do you mean by virtual base  class?" & Chr(13) & _
    "e)  How do you display a variable, say int I, in octal format when default is decimal?"
StrQuestions(4) = "a)  What will the following program display?" & Chr(13) & _
    "#include <iostream.h>" & Chr(13) & _
    "main()" & Chr(13) & _
    "{cout.width(10);cout << Hello << \n;cout.fill('*');" & Chr(13) & _
        "cout.width(10);cout << Hello << \n;cout.setf(ios::left);" & Chr(13) & _
    "cout.width(10);cout << Hello << \n;return O;" & Chr(13) & _
    "}" & Chr(13) & _
"b)  Create a class that contains our name and roll number. Using new dynamically allocate an object of this class and put your name and roll number into these fields and within these objects. Finally release all dynamically allocated memory after display by using delete." & Chr(13) & _
"c)  Write C++ program to create an output file that adds your name and roll number to it, closes it and open it again as an input file and reads the information stored in it."
StrQuestions(5) = "Write a C++ code using modular programming ro show the number of days in a month."
StrQuestions(6) = "a)  Use the swap function to show that C++ functions can be called by value" & Chr(13) & _
       "by pointers and by reference." & Chr(13) & _
"b)  Compile the problem through separate files."
StrQuestions(7) = "Define a class batsman which will have the following information" & Chr(13) & _
"number of tests played" & Chr(13) & _
"number of first class matches;" & Chr(13) & _
"number of one day internationals;" & Chr(13) & _
"other data imaginatively entered; 9assuming you know cricket)" & Chr(13) & _
"The class must have input to update the statistics, and should be able to give various statistics when asked for."
StrQuestions(8) = "A bank has two kinds of accounts , savings and fixed deposit." & Chr(13) & _
"The interest rates for these are 4% and 11% respectively. Calculate the total amount(principal + interest) for these two schemes after a certain period of times. Show he use of enum and dynamic if possible. Calculate the actual on simple interest basis." & Chr(13) & _
"a.  Write the code for the following :" & Chr(13) & _
"Declare pointers to the storage space for the names (of 30 chars each) of 5 students. Also assign pointers for the marks obtained in three semesters for each of the students. Accept the name and marks from the user, and show the output in the following format:" & Chr(13) & _
"S.No Name" & Chr(13) & _
"SEM -1" & Chr(13) & _
"SEM -2" & Chr(13) & _
"SEM -3" & Chr(13) & _
"TOTAL -(100)" & Chr(13) & _
                           "(100)" & Chr(13) & _
                           "(100)" & Chr(13) & _
                           "(300)" & Chr(13) & _
"____________________________________________________________" & Chr(13) & _
"1.  S. GHOSH" & Chr(13) & _
"55" & Chr(13) & _
"65" & Chr(13) & _
"70" & Chr(13) & _
"190" & Chr(13) & _
"2.  …" & Chr(13) & _
"3.  …" & Chr(13) & _
"Lastly free the allocated space to he pointers (use NEW & DELETE for memory allocation and deallocation)."
StrQuestions(9) = "Write a code with default arguments to find the volume or area of a rectangular parallelepiped. Use this function with inline specification."
End Sub
Private Sub questions_paper2()
StrQuestions(1) = "Use the swap function buy reference and by pointers as arguments to build up an archival file. Use this archive to link your main program."
StrQuestions(2) = "Write small programs which illustrates function scope and files scope."
StrQuestions(3) = "Write a program to show how a global name can be mad local to the file"
StrQuestions(4) = "Declare a class of type Department with private memories processors, btech student, mtech student, and research scholar." & Chr(13) & _
    "(i) Declare a function output() as a friend function & show how it can have full access to private & protected members." & Chr(13) & _
    "(ii) If the declaration of output() is moved in private, what will be tne output of (i)?" & Chr(13) & _
"(iii)   Declare another class Deptoffice, and show how all members of department classcan become friends of Dept office via a single statement."
StrQuestions(5) = "Write C++ code by utilizing the concept of derived class for the following problem." & Chr(13) & _
    "In a computer system it is required to maintain three user groups called owner, user and others. The groups should have the following access capabilities for accessing a file. The other group can only read from the file. The user group as read and execute permission but is not allowed to delete or modify the file. Only the owner group has read, write and execute permissions on the file. Draw the inheritance diagram. The user of the program should first identify himself as belonging to a given group, and the can access any file as per the access restrictions mentioned. If he is not a bonafide group member for the usage of the file access to the file will be denied."
StrQuestions(6) = "Write a C++ code for a class coordinates which is used to represent the (x,y) coordinates on a two dimensional plane. Show how operator+() can be overloaded for the coordinates class by using a friend function. "
StrQuestions(7) = "Use display member functions both for a batsman class and bowler class. The display functions should be overloaded in both the classes. For the bowler class the member function display will be of the following forms giving the corresponding output." & Chr(13) & _
"display(Kapil,'W','A')" & Chr(13) & _
"The output should show wickets and bowling average for any bowler." & Chr(13) & _
"display ('A') or display('W');" & Chr(13) & _
"The output should show either the bowling average or the wickets taken respectively." & Chr(13) & _
"Similarly develop a display member for the batsman class."
StrQuestions(8) = "Develop link list class for the bowler class with the following implementations:" & Chr(13) & _
    "a.  insert a member at the beginning of the list." & Chr(13) & _
    "b.  append a new member at end of list." & Chr(13) & _
    "c.  remove all members" & Chr(13) & _
    "d.  check whether a gie bowler is existing in the list" & Chr(13) & _
    "e.  display the list" & Chr(13) & _
    "f.  concatenate two lists."
StrQuestions(9) = "Design a class company which can have the following data:" & Chr(13) & _
    "a.  either name of the company (pointer to char type) or code of the company(integer type);" & Chr(13) & _
    "b.  yearly turnover(int);" & Chr(13) & _
    "c.  type (Private company or Public company)" & Chr(13) & _
    "Now, call proper functions depending on company name or company code to show the information of the company. Create a function called reverse() that takes two parameters. The first parameter, called str, is a pointer to a string that will be reversed upon return from the function." & Chr(13) & _
"The second parameter is called count and it specified how many characters of str to reverse. Give count a default value that when present tells reserve () to reverse the entire string. Write a main function, input the count and print the input as well as output string."
End Sub

Private Sub questions_paper3()
StrQuestions(1) = "Use the addition operator to show an object of the batsman class when added with an integer number will increment the number of innings played by one, and increment the number of runs scored by the given number. Similarly, show the same effect for a bowler where the given integer number stands for the wickets taken."
StrQuestions(2) = "Implement he stack class as a template class in C++. Using this template class form the following stacks of students name of a class; stack of roll numbers of the above students; stack of marks obtained by the students."
StrQuestions(3) = "Write an interactive program in C++ which will run as follows. Enter afile name: TESTFILE Enter the option -" & Chr(13) & _
    "1.  Read-" & Chr(13) & _
    "2.  Write -" & Chr(13) & _
    "3.  Search -" & Chr(13) & _
"if you choose 1 then a file named TESTFILE1 will be opened and it can be read." & Chr(13) & _
"If you choose 2 then file named TESTFILE2 will be written or modified" & Chr(13) & _
"If you choose 3 then a file named TESTFLE3 will be opened and will ask for a number to search and show the result whether the entered number exists in TESTFILE3 or not.(Do not use stdio.h)"
StrQuestions(4) = "Write a program in C++ which have" & Chr(13) & _
    "a.  the overloaded + add which adds two arrays element wise;" & Chr(13) & _
    "b.  the overloaded - subtract which subtract two array element - wise, the right operand from the left." & Chr(13) & _
    "c.  the overloaded == returns true if each element of both the array(each operand) is the ame and returns false otherwise." & Chr(13) & _
    "How do you modify the above solution if it overloads the operators by using friend functions."
StrQuestions(5) = "Write a program for vectors handling complex classes with the following operations." & Chr(13) & _
"A.addition" & Chr(13) & _
"b.scaler multiplication" & Chr(13) & _
"c.input" & Chr(13) & _
"d.output "
StrQuestions(6) = "Use addition operator to show that in a student class when one data member (a string) added with another string(may be a data member) will concatenate to two strings separated by a blank and also display the message."
StrQuestions(7) = "Write a program to overload the *operation relative to a class called distance between two points having fixed coordinates."
StrQuestions(8) = "Create a string class which contains a string and its length. Use new delete for allocating the memory when a string object is created (initial string length is set to zero) deallocating the memory when it is destroyed."
StrQuestions(9) = "Create a class called stopwatch that emulates a stopwatch that keeps track of elapsed time. Use a constructor to initially set the elapsed time to zero. Provide two member function called start() and stop() that turn on and off the timer, respectively. Include a member function called show() that displays the elapsed time. Also have the destructor function automatically display the elapsed  time when a stopwatch object is destroyed."
StrQuestions(10) = "Create a class called calendar that is passed the current system time and date as a parameter to its constructor when it is created. Include a member function that displays this time and date on the screen.(Use the standard time and date functions found in the standard library to find and display the date.)"
End Sub

Private Sub questions_paper4()
StrQuestions(1) = "Create a class called room whose constructor function is passed here double values, which represents length, breadth and height of the room. Have the room class compute the volume of the air kept in the room and store the result in a double variable. Include a member functin called volume that displays the volume of each room object"
StrQuestions(2) = "Create a general class called Institution that describes certain characteristics about an ideal institution. This is inherited by two derived classes called degree and diploma. Each one defines some specific criteria for each of the two type/. Include member functions admission and examination for institution class which inherits both but project for degree only. Write C++ code for the member functions showing the inheritance of the class."
StrQuestions(3) = "Assume that Batsman is derived from the class called Cricket- Player which consists of the following information:" & Chr(13) & _
"(i) Player name" & Chr(13) & _
"(ii)    country" & Chr(13) & _
"(iii)   date of birth" & Chr(13) & _
"(iv)    date of debut" & Chr(13) & _
"(v) number of test played" & Chr(13) & _
"(vi)    number of one day international played" & Chr(13) & _
"The Batsman is now modified to contain the following information" & Chr(13) & _
"(i) Number of innings" & Chr(13) & _
"(ii)    Number of one day international played" & Chr(13) & _
"(iii)   Number of not outs in innings in one day international played" & Chr(13) & _
"(iv)    Number of not outs in innings in test" & Chr(13) & _
"(a) Draw the hierarchy diagram" & Chr(13) & _
"(b) Write member functions to display and update information of player for showing batsman biodata batting average etc."
StrQuestions(4) = "Suppose that it is required to generate a table of multiples of any given number. Write a C++ program,m that allows the user to enter a number, and then generate the table, formatting it into 5 columns of 10 lines(i.e. generate multiples of upto 50)"
StrQuestions(5) = "C++ operators such as >>, which reads input from the keyboard, converts a sequence of digits into a number. Write a program that emulates this behaviour. It should allow the user to type up to six digits, then display the resulting numbe as a type long integer, The individual digits should be read using getche(). For cnstructing the number, multiply the partial result by 10, and then add the new digit."
StrQuestions(6) = "Create a class called person having data member name, age working area and name of employer. Write a C++ program so that it creates an object of type person and sets the data in this object to the name entered by the user and keeps a pointer to the object in the pptr array  and add a member  function psort which sorts a group of person objects based on the alphabetic order of their names."
StrQuestions(7) = "Develop a menu driven C++ program having more that two operands the option as follows :" & Chr(13) & _
"1.  Additions" & Chr(13) & _
"2.  Subtraction" & Chr(13) & _
"3.  Multiplication" & Chr(13) & _
"4.  Division" & Chr(13) & _
"5.  Print" & Chr(13) & _
"6.  Do you want to continue? (Y/N)"
StrQuestions(8) = "Create a class of big cities BIGCITY of India, the data member of the class are name of the city, std code(say for Calcutta STD code is 033) etc. Write a member function which interactively ask the name and address, local phone number of residents and print in the following format." & Chr(13) & _
"1.  Name : S.P. Rama Rao" & Chr(13) & _
"2.  Address : 3/2 APC Road" & Chr(13) & _
"3.  Pin Code No: 700052" & Chr(13) & _
"4.  Phone No, : (033-4347270)"
StrQuestions(9) = "Crate a class called shapes for different geometric figures like circlem, sphere, cylinder, cone etc. Write a member function to find out the volume of each shape, it should take the argument of same type (Radius or height or both radio and height). Write a main function that reads the shape and its radius and/or height as it is and displays the result."
StrQuestions(10) = "Raising a number n to a power p is the same as multiplying n by itself p times. Write a function called power() that takes a double value for n and int value for p, and returns the results double value. Use a default argument of 3 for p, so that if this argument is omitted, the number will be cued. Write a main() function that reads different values of n and p and prints the result."
End Sub

Private Sub questions_paper5()
StrQuestions(1) = "Write function that takes two distinct value as arguments and returns the larger one. Include a main() program that accepts three distinct figures from the user, compares them and displays the larger. Enter a set of three distances and print the result."
StrQuestions(2) = "Create a class called time that has separate int data member for hours. Minutes and seconds. One constructor should initialize this data to 0 and another should initialize it to fixed values .Write a member function to display the time in the format 10:30:40. The final member function should add two objects of type time passed arguments."
StrQuestions(3) = "There is a toll collection mechanism at Vidyasagar Bridge in Calculatta, where every car passing across is expected to pay a toll of Rs. 5/- at a toll booth. Sometimes a car might speed through the booth without paying. The booth keeps track of the number of cars crossing the bridge, and the total money collected. " & Chr(13) & _
    "Model the too booth with a class tollbooth. Two data items are defined within the class, ncar of type unsigned int to keep track of the numbers of cars, and total of type double to keep track of money collected." & Chr(13) & _
"-   A constrictor initializes ncar and total to both zeros." & Chr(13) & _
"-   A member function good car increments ncar and adds 5 to total." & Chr(13) & _
"-   A member function bad car only increments ncar." & Chr(13) & _
"-   A member function display displays the two totals." & Chr(13) & _
    "Write a test program for above. The program should allow the user to push one key to count a good car, and another to count a bad car. Pressing ESCAPE key should cause the program to invoke display and exit."
StrQuestions(4) = "Write a function reverse that reverses a given character string. Write a program to invoke reverse. The program should read a character string from the user, reverse it and print out the result. It should also indicate whether the given string is a palindrome (i.e. one which reads the same backward or forward). The program should be able to handle character strings with embedded blanks."
StrQuestions(5) = "Create a class called employee that contains the employee number (type long), name (an array of char), and salary (type double). Include a member function called get data to get data from user for insertion into the object and another function put data to display the data. Write another function increment to increase the salaries of all employees by 20%."
StrQuestions(6) = "Consider a publishing company that markets both books and video cassette versions of certain subjects. Create a class publication that stores the title (string) and pricefloat of a publication. From this class derive two classes :" & Chr(13) & _
    "book,which adds a page count int;" & Chr(13) & _
    "cassette, which adds a playing time in minutes float." & Chr(13) & _
    "Each of these three classes should have a getdata function to get data from the usr at the keyword , and a putdata function to display its data."
StrQuestions(7) = "Consider a publishing company that markets both books and video cassette versions of certain subjects. Create a class publication that stores the title(string) and price float of a publication which adds a page count int;" & Chr(13) & _
"cassette, which adds a playing time in minutes float." & Chr(13) & _
"Add another base class sales that holds an array of three floats so that it can record the sales in Rupees of a particular publication for the last six months. Include a getdata() function to get three sales and a putdata() function to display the sales. The book and cassette classes can be derived from both publication and sales. An object of class book orcassette inputs and output sales along with its other book or cassette inputs and output sales along with its  data. Write a main() function to create a bool object and a cassette object and exercise their input/output capabilities."
StrQuestions(8) = "Write a program that reads a group of numbers from the keyword and places them into an array of type float. Once the numbers are stored in the array, the program should fine out the minimum, maximum and average of the numbers, and print out the result., Us pointer notation wherever possible."
StrQuestions(9) = "Write a program in C++ that emulates the DOS COPY command. The program should be invoked with two command line arguments, the source file and the destination file, as follows:" & Chr(13) & _
"C> ccopy srcfile.xxx dstfile.yyy" & Chr(13) & _
"In the program check that the usr has typed the correct number of command line arguments, and that the file specified can be opened. The program should signal an error if destination file already exists."
End Sub

Private Sub questions_paper6()
StrQuestions(1) = "a)  What is Dynamic binding?" & Chr(13) & _
        "b)  Illustrate polymorphism with at least two examples."
StrQuestions(2) = "a)  what is the value of 'Sat' in the following?" & Chr(13) & _
"Enum {sun,mon,tue,thu,Fri = 1, Sat}" & Chr(13) & _
"Justify your answer?" & Chr(13) & _
"b)  What are the advantages of reference variable over pointers."
StrQuestions(3) = "What are constructors and destructors? Explain it use by quoting examples."
StrQuestions(4) = "Distinguish between a Macro and Inline functions?"
StrQuestions(5) = "Illustrate the concept of data hiding in a class? Give one example."
StrQuestions(6) = "Define a class 'complex no' which has 2 real nos. as private variables one represents the real part and the other for complex part. Define a constructor to initialize the object."
StrQuestions(7) = "Write a function which takes two objects of the type Complex no and returns an object of the same type which is the sun of two complex numbers."
StrQuestions(8) = "Design a program to implement a stack data structure to store integer data types, defining suitable member functions (push and pp) and a constructor."
StrQuestions(9) = "Define a class 'MCA2_oops' which has the members: Entry_No, Name, Marks_Minor, Marks_Major, Total. Write a main program for handling 30 students and displaying their marks."
StrQuestions(10) = "Write a program to read Square Matrix of size from the keyboard. Write a function to find the race of a matrix."
End Sub

Private Sub questions_paper7()
StrQuestions(1) = "Write a program to define a stack as a class and write a program for evaluation of Postfix expression using stack."
StrQuestions(2) = "Define a class employee with variables Emp-code, Name, Designation, Basics, DA, HRA and total pay where Emp-code is the sl. No. and Name, Designation and Basic are obtained from Keyboard input, DA is 40% of Basic and HRA is 20% of Basics subjected to a minimum of Rs. 500/- Create an array of 10 objects of type employee. Write a program for accepting data of 10 employees and print the details in the following format" & Chr(13) & _
"Sl. No. Name    Designation Basic pay    DA  HRA Total-pay"
StrQuestions(3) = "a.  What is dynamic initialization of objects, what are the advantages of this and how is it achieved?" & Chr(13) & _
"b.  What are the limitations on constructors and destructors."
StrQuestions(4) = "Design a Queue data structure using a class Queue with suitable member functions, constructors and destructors. Define the member functions for insertion and deletion of elements from the queue."
StrQuestions(5) = "a.  Cam a 'friend' function be used for overloading the assignment operator? Justify your answer?" & Chr(13) & _
"c.  What is a conversion function and when is it used?"
StrQuestions(6) = "Create a class 'sqmatrix' of the type Square Matrix of real numbers and overload the operator '=' to find the sum of the two objects of the type sqmatrix."
StrQuestions(7) = "a.  A private member of a base class is inheritable. Is there any ay for the objects of a derived class to access the private members of base class without modifying the class? Justify your answer?" & Chr(13) & _
"b.  Describe how an object of a class that contains objects of other classes are created. Supplement your answer with an example."
StrQuestions(8) = "a.  What is a class inheritance? What are its advantages?" & Chr(13) & _
"b.  Design a class 'Employee' of a n organization with members Employe-no, Name, Designation, Basic Pay. Define a derived class Executing whose Basic pay is more than Rs. 10000."
StrQuestions(9) = "a.  Distinguish between an virtual class and an abstract class?" & Chr(13) & _
"b.  Class SET2 is derived from Class SET1 such that class SET2 does not contains any at members of its own. Does the class SET2 require constructors? Justify your answer." & Chr(13) & _
"c.  Distinguish between containership and inheritance."
End Sub

Private Sub questions_paper8()
StrQuestions(1) = "a.  What are pure virtual functions? What are the implications in making a virtual function 'Pure'?" & Chr(13) & _
"b.  What does this pointer point to and what is its use?" & Chr(13) & _
"c.  What is the difference between a virtual function and a overload function."
StrQuestions(2) = "Design manipulator and the corresponding C++ code to provide he following output specifications:" & Chr(13) & _
"i)  take a,b,c float values form the keyboard input." & Chr(13) & _
"ii) compute he value (-b + (b**2 - 4*a*c))/2*a" & Chr(13) & _
"iii)    repeat the above for 10 times"
StrQuestions(3) = "a.  Illustrate the concept of polymorphism with examples " & Chr(13) & _
"b.  How is polymorphism achieved at compile time?"
StrQuestions(4) = "a.  What is a template class? Where it is desirable to use template class? Give an example with C++ code"
StrQuestions(5) = "a.  What is a stream? Give the I/O steam class hierarchy?" & Chr(13) & _
"b. Illustrate the use of manipulators. Give two examples."
StrQuestions(6) = "What is the result of the following program segment?" & Chr(13) & _
"for(i=5;i<=50;i=i+5)" & Chr(13) & _
"{" & Chr(13) & _
 "cout.precision (6)" & Chr(13) & _
 "cout.Width (8)" & Chr(13) & _
 "cout<<I;" & Chr(13) & _
 "cout.Width (10)" & Chr(13) & _
 "cout<<i*i/2<<\n;" & Chr(13) & _
"}" & Chr(13) & _
"Also explain the result at the end of each statement."
StrQuestions(7) = "a.  Distinguish the functions of 'cin' and 'getline' which are used for input." & Chr(13) & _
"b.  Discuss the syntax for creating user defined manuipulators"
StrQuestions(8) = "Write a C++ code to read a text file and display the following information:" & Chr(13) & _
"a)  Number of lines" & Chr(13) & _
"b)  Number of words" & Chr(13) & _
"c)  Number of 3 letter words."
StrQuestions(9) = "Define a class matrix of size m×n (m>n) and define a derived class square matrix of size n. Write a program in C++ to computer the product of two matrices and trace of square matrix. The program should read the values form a data file."
StrQuestions(10) = "Our usual notation of writing a mathematical expression is called infix notation in which the binary operator appears n between two operands(e.g. a+b). If the operator appears after the operands then it is called postfix notation(e.g. ab+). Design an algorithm to convert an expression in Infix notation to postfix notation."
End Sub

Private Sub questions_paper9()
StrQuestions(1) = "Identify various classes required and design a C++ code for the problem in   Ques No. 80."
StrQuestions(2) = "Write a C++ programmed to evaluate the expressions in postfix notation."
StrQuestions(3) = "Define a class list element to denote the node of a linked list such that each object can have two members one is data and t ether is pointer(link) to another node in the list. Write a C++ code to read a string of n characters and store the characters of the string in the above data structure."
StrQuestions(4) = "Write a C++ code to split the list in the above question into two separate lists such that all characters in odd positions will be in one list and the characters in the even positions will be in another single list.    "
End Sub

Private Sub TrueorFalse_Paper1()
StrQuestions(1) = "Organizing a data wit methods that operate on the data and crating a new data type is called encapsulation."
StrQuestions(2) = "Classis similar to a variable."
StrQuestions(3) = "A nonmember function apart from friend cannot access nonpublic member of a class"
StrQuestions(4) = "By default, members cannot be inherited."
StrQuestions(5) = "Protected members cannot be inherited."
StrQuestions(6) = "Inline function specifier reduces the overheads associated with a normal function call."
StrQuestions(7) = "The new operator always returns a void pointer."
StrQuestions(8) = "'This' pointer has to be used while accessing data members in a member function."
StrQuestions(9) = "In C++, identifiers have to be declared at the beginning of the blocks."
StrQuestions(10) = "    In C++, an identifier must e initialized using constant expression."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"

End Sub

Private Sub TrueorFalse_Paper2()
StrQuestions(1) = " In C++, we cannot assign directly an integer value to enum."
StrQuestions(2) = " Size of operator applied any enum data type will give the size of an int in C++."
StrQuestions(3) = " Const identifiers have to be initialized at the time of declaration in C++."
StrQuestions(4) = " In C++ void pointer cannot be assigned to other pointer type without casting."
StrQuestions(5) = " 'new' operator always returns a void pointer."
StrQuestions(6) = " Data objects can be initialized when allocating memory using 'new'."
StrQuestions(7) = " Scope resolution operator has the highest precedence."
StrQuestions(8) = " Reference to an object behaves like a constant pointer. Inline function saves memory but reduces efficiency for small function call."
StrQuestions(9) = " Inline function saves memory but reduces efficiency for small function call."
StrQuestions(10) = "    We can make function inline using the keyword 'inline'."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"

End Sub

Private Sub TrueorFalse_Paper3()
StrQuestions(1) = " Inline function are exactly same as #define macro in C."
StrQuestions(2) = " void f(int x=1, int y = 0, int z) is valid function declaration with defaultarguments."
StrQuestions(3) = " A label can be invoked from any lock of the same function where label is defined."
StrQuestions(4) = " Global variables have file scope."
StrQuestions(5) = " A static variable declared inside a function exists till the execution of that function only."
StrQuestions(6) = " Private members of an structure can be accessed directly from the outside of the structure."
StrQuestions(7) = " One does require the inline keyword when implementing an inline function outside the class."
StrQuestions(8) = " The following syntax is valid.void inlineGram_Panchayat :: show_gram_Panchayat_info()"
StrQuestions(9) = " '[]' operator is a unary operator."
StrQuestions(10) = "    The members of class by default access control of class, structure & union have same access rules as that of a class"

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper4()
StrQuestions(1) = " Apart from the default access control of a class, structure & union have same access rules as that of a class."
StrQuestions(2) = " An object is an allocated space in memory."
StrQuestions(3) = " Constructor returns void type value."
StrQuestions(4) = " A class with no constructor can not initialized with another object of the same class."
StrQuestions(5) = " A copy constructor is used to copy an object member wise in another object of the same class."
StrQuestions(6) = " Member function cannot be called from within a constructor."
StrQuestions(7) = " A class can have aconst constructor."
StrQuestions(8) = " We cannot have the address of constructor."
StrQuestions(9) = " A constructor cannot be explicitly called."
StrQuestions(10) = "    General default constructor & copy constructor are private."

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

Private Sub TrueorFalse_Paper5()
StrQuestions(1) = " Object never get destroyed automatically even if it goes out of scope."
StrQuestions(2) = " A destructor can have arguments like constructor."
StrQuestions(3) = " A destructor can have return type."
StrQuestions(4) = " Objects get destroyed in the reverse order as they are created."
StrQuestions(5) = " 'this' is an implicit pointer."
StrQuestions(6) = " The value of 'this' pointer can be changed."
StrQuestions(7) = " The data member of an object can not be modified inside a constant member function."
StrQuestions(8) = " const int show_somedata() is the way of declaring constant member function."
StrQuestions(9) = " 'this' pointer can be casted."
StrQuestions(10) = "    An address of a member function can be derived by the statement.& func_name"

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper6()
StrQuestions(1) = " Assignment operator assigns one object member wise to another object."
StrQuestions(2) = " In case of nested class, enclosing class can directly access the private data member of nested class."
StrQuestions(3) = " To initialize an object of a class A(explicitly) , which is a data member of a class B, he construct of A has to be invoked in the constructor of class B."
StrQuestions(4) = " It is possible to initialize non class data member by initializing outside the constructor definition block."
StrQuestions(5) = " ios contains a pointer to streambuf."
StrQuestions(6) = " iostream is inherited from istream, ostream and ios class."
StrQuestions(7) = " cerr is an object of ostream_withassign."
StrQuestions(8) = " Static member function have file scope."
StrQuestions(9) = " Static data member occurs in only class scope."
StrQuestions(10) = "    Static data member can be declared as const too."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper7()
StrQuestions(1) = " The manipulator eds() inserts newline and flush."
StrQuestions(2) = " If a friend function is declared inside a class it can access all data members of the class."
StrQuestions(3) = " A friend function declared inside a class does not need scope resolution operator while defining it."
StrQuestions(4) = " A friend function has class scope."
StrQuestions(5) = " We can declare a member function of a class as friend function of some other class."
StrQuestions(6) = " A friend function declared inside he class has to be called like other member function call."
StrQuestions(7) = " Functions which have different return type but same number of arguments of same type in same order are called overload."
StrQuestions(8) = " Function overloading is possible only if signature is same."
StrQuestions(9) = " struct A::B /*…*/; is the syntax off derivation."
StrQuestions(10) = "    Members of a derived class can only be accessed by the objector pointer to the same class."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper8()
StrQuestions(1) = " A derived class can access the private members of its public base classdirectly."
StrQuestions(2) = " By default, a class derivation is public."
StrQuestions(3) = " Destruction of objects take place in reverse order of their construction."
StrQuestions(4) = " We cannot assign a base class pointer to derived class pointer."
StrQuestions(5) = " setf() member function reads current flag's value."
StrQuestions(6) = " A virtual function can be a static member function."
StrQuestions(7) = " A virtual function cannot be declared a friend In another class."
StrQuestions(8) = " We cannot have an object of but can have pointer to an abstract class."
StrQuestions(9) = " A class containing a virtual function is called abstract class."
StrQuestions(10) = "    A class containing even one pure virtual function is called abstract class."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"

End Sub

Private Sub TrueorFalse_Paper9()
StrQuestions(1) = " fstream class contains two pointers, one for input  and another for output."
StrQuestions(2) = " In C++ all identifiers must be declared before they can be used."
StrQuestions(3) = " In C, variables can be local to a function where as in C++, variables can bedeclared as local to a block."
StrQuestions(4) = " In C++, run time boundary checking on array indexing is possible."
StrQuestions(5) = " Reference are also pointers."
StrQuestions(6) = " In a function prototype, all parameters that take default values must appear to the right of those that  do not take default values."
StrQuestions(7) = " Any C structure is also valid in C++."
StrQuestions(8) = " Constructors can not be overloaded."
StrQuestions(9) = " A simple passing of objects using call by value may cause problems some times."
StrQuestions(10) = "    A function can not return an object."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper10()
StrQuestions(1) = " ++I and I++ always gives the same result."
StrQuestions(2) = " Friendship is not transitive."
StrQuestions(3) = " Friend functions can be inherited."
StrQuestions(4) = " A constructor can be virtual."
StrQuestions(5) = " unsetf() member function reads current format flag’s value."
StrQuestions(6) = " Operator(‘=’) can be used as both member and non member function."
StrQuestions(7) = " Binary operator(“+”) can be used as non member function."
StrQuestions(8) = " An overload binary operator() can be overloaded for ternary operation."
StrQuestions(9) = " :: (scope resolution operator) cannon be overloaded for other operation."
StrQuestions(10) = ".   An overload binary operator() should have two arguments always."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub multiplechoice_Paper1()
strmultiple(1) = "In C++, a function contained within a class is called –"
strmuloption(1) = "A member function"
strmuloption(2) = "An operator"
strmuloption(3) = "A class function"
strmuloption(4) = "A method"

strmultiple(2) = "Which of the following are good reasons to use an object oriented language ?"
strmuloption(5) = "You can define your data types"
strmuloption(6) = "Program statements are simpler than in procedural languages"
strmuloption(7) = "An OO program can be taught to correct its own errors"
strmuloption(8) = "Its easier to conceptualize and OO program"

strmultiple(3) = "When a language has the capacity to produce new data type, it is said to be –"
strmuloption(9) = "Reprehensible"
strmuloption(10) = "Encapsulated"
strmuloption(11) = "Overload"
strmuloption(12) = "Extensible"

strmultiple(4) = "A normal C++ operator that acts in a special ways on newly defined data types is said to be –"
strmuloption(13) = "Glorified"
strmuloption(14) = "Encapsulated"
strmuloption(15) = "Classified"
strmuloption(16) = "Overload"

strmultiple(5) = "main() returns a value of type –"
strmuloption(17) = "real"
strmuloption(18) = "character"
strmuloption(19) = "int"
strmuloption(20) = "null"

strmultiple(6) = "Sharing of common information are achieved by the concept of –"
strmuloption(21) = "Virtual copying"
strmuloption(22) = "Inheritance"
strmuloption(23) = "Encapsulation"
strmuloption(24) = "None of these"

strmultiple(7) = "memorizing the new terms used in C++ is –"
strmuloption(25) = "critically important"
strmuloption(26) = "something you can return to later"
strmuloption(27) = "the key to wealth and success"
strmuloption(28) = "completely irrelevant"

strmultiple(8) = "An expression"
strmuloption(29) = "Usually evaluates to a numerical value"
strmuloption(30) = "Indicates the emotional state of the program"
strmuloption(31) = "Always occurs outside a function"
strmuloption(32) = "may be a part of the statement"

strmultiple(9) = "In a for loop with a multi statement loop body, semicolons should appear following –"
strmuloption(33) = "The for statement itself"
strmuloption(34) = "The closing brace in a multi statement loop body"
strmuloption(35) = "Each statement within the loop body"
strmuloption(36) = "The test expression"

strmultiple(10) = " A variable defined within a block is visible"
strmuloption(37) = "From the point of definition onwards in the program"
strmuloption(38) = "From the point of definition onwards in the function"
strmuloption(39) = "From the point of definition onwards in the block"
strmuloption(40) = "Throughout the function"

strmulanswer(1) = "A member function"
strmulanswer(2) = "You can define your data types"
strmulanswer(3) = "Extensible"
strmulanswer(4) = "Overload"
strmulanswer(5) = "null"
strmulanswer(6) = "Inheritance"
strmulanswer(7) = "the key to wealth and success"
strmulanswer(8) = "may be a part of the statement"
strmulanswer(9) = "Each statement within the loop body"
strmulanswer(10) = "From the point of definition onwards in the block"
End Sub

Private Sub multiplechoice_Paper2()
strmultiple(1) = "  The library function exit() causes an exit from –"
strmuloption(1) = "The loop in which it occurs"
strmuloption(2) = "The block in which it occurs"
strmuloption(3) = "The function in which it occurs"
strmuloption(4) = "The program in which it occurs"

strmultiple(2) = "  The getch() library –"
strmuloption(5) = "Returns a character when any key is pressed"
strmuloption(6) = "Returns a character when ENTER is pressed"
strmuloption(7) = "Displays a character on the screen when a key is pressed"
strmuloption(8) = "Does not display a character on the screen"

strmultiple(3) = "  Binary overload operators are passed __________________ arguments"
strmuloption(9) = "Two"
strmuloption(10) = "One"
strmuloption(11) = "No "
strmuloption(12) = "None of the above"

strmultiple(4) = "  Unary overload operators are passed _________________ arguments"
strmuloption(13) = "Two"
strmuloption(14) = "One"
strmuloption(15) = "No"
strmuloption(16) = "None of the above"

strmultiple(5) = "  The && and || operators"
strmuloption(17) = "Compare two numeric values"
strmuloption(18) = "Combine two numeric values"
strmuloption(19) = "Compare two Boolean values"
strmuloption(20) = "Combine two Boolean values"

strmultiple(6) = "  The break statement cause an exit –"
strmuloption(21) = "Only from the innermost loop"
strmuloption(22) = "Only from the innermost switch"
strmuloption(23) = "Form all loops and switches"
strmuloption(24) = "From the innermost loop or switch"

strmultiple(7) = "  The goto statements causes control to go to –"
strmuloption(25) = "An operator"
strmuloption(26) = "A label"
strmuloption(27) = "A variable"
strmuloption(28) = "A function"

strmultiple(8) = "  When accessing a structure member, the identifier to the left of the dot operator is the name of –"
strmuloption(29) = "A structure member"
strmuloption(30) = "A structure tag"
strmuloption(31) = "A structure variable"
strmuloption(32) = "The keyword struct"

strmultiple(9) = "  cc __________________ option is used only to create object file"
strmuloption(33) = "–a"
strmuloption(34) = "–o"
strmuloption(35) = "–c"
strmuloption(36) = "none of these"

strmultiple(10) = " An enumerated data type brings together a group of –"
strmuloption(37) = "Items of different data type"
strmuloption(38) = "Related data variables"
strmuloption(39) = "Integers with user defined names"
strmuloption(40) = "Constant values"

strmulanswer(1) = "The program in which it occurs"
strmulanswer(2) = "Returns a character when any key is pressed"
strmulanswer(3) = "Two"
strmulanswer(4) = "One"
strmulanswer(5) = "Compare two Boolean values"
strmulanswer(6) = "Only from the innermost loop"
strmulanswer(7) = "A label"
strmulanswer(8) = "A structure tag"
strmulanswer(9) = "–o"
strmulanswer(10) = "Integers with user defined names"
End Sub

Private Sub multiplechoice_Paper3()
strmultiple(1) = "  A function argument is -"
strmuloption(1) = "A variable in the function that receives a value from the calling program."
strmuloption(2) = "A way that function resist accepting the calling program's values"
strmuloption(3) = "A value sent to the function by the calling program"
strmuloption(4) = "A value returned by the function to the calling program."

strmultiple(2) = "  :: is known as"
strmuloption(5) = "scope resolution operator"
strmuloption(6) = "global operator"
strmuloption(7) = "both (a) & (b)"
strmuloption(8) = "none of these"

strmultiple(3) = "  Which of the following can legitimately be passed to a function ?"
strmuloption(9) = "A constant"
strmuloption(10) = "A variable"
strmuloption(11) = "A structure"
strmuloption(12) = "A header file"

strmultiple(4) = "  When an argument is passed by reference"
strmuloption(13) = "A variable is created in the function to old the argument's value"
strmuloption(14) = "The function cannot access the argument's value"
strmuloption(15) = "A temporary variable is created in the align program to hold the argument's value"
strmuloption(16) = "The function accesses the argument's original value in the calling program"

strmultiple(5) = "  _________________ operator must have one class object"
strmuloption(17) = "+"
strmuloption(18) = "new"
strmuloption(19) = "all "
strmuloption(20) = "none of these"

strmultiple(6) = "  Overload functions -"
strmuloption(21) = "Are a group pf function with the same name"
strmuloption(22) = "All have the same number and types of arguments"
strmuloption(23) = "Make life simpler for programmer"
strmuloption(24) = "May fail unexpectedly due to stress"

strmultiple(7) = "  The default argument has a value that -"
strmuloption(25) = "May be supplied by the calling program"
strmuloption(26) = "May be supplied by the function"
strmuloption(27) = "Must have constant value"
strmuloption(28) = "Must have variable value"

strmultiple(8) = "  A static automatic variable is used to -"
strmuloption(29) = "Make a variable visible to several function"
strmuloption(30) = "Make variable visible to only one function"
strmuloption(31) = "Converse memory when a function is not executing"
strmuloption(32) = "Retain a value when function is not executing"

strmultiple(9) = "  In a class specifier, data or function designated private are accessible -"
strmuloption(33) = "To any function in the program"
strmuloption(34) = "Only if you know the password"
strmuloption(35) = "To member function of that class"
strmuloption(36) = "Only to public members of the class"

strmultiple(10) = " The dot operator (or class member access operator) connects thefollowing two entities (reading from left to right) -"
strmuloption(37) = "A class member and a class object"
strmuloption(38) = "A class object and a class"
strmuloption(39) = "A class and a member of that class"
strmuloption(40) = "A class object and a member of that class"

strmulanswer(1) = "A value sent to the function by the calling program"
strmulanswer(2) = "both (a) & (b)"
strmulanswer(3) = "A variable"
strmulanswer(4) = "A variable is created in the function to old the argument's value"
strmulanswer(5) = "new"
strmulanswer(6) = "Are a group pf function with the same name"
strmulanswer(7) = "Must have constant value"
strmulanswer(8) = "Converse memory when a function is not executing"
strmulanswer(9) = "To member function of that class"
strmulanswer(10) = "A class object and a class"
End Sub

Private Sub multiplechoice_Paper4()
strmultiple(1) = "  A member function can always access the data -"
strmuloption(1) = "In the object of which it is a member"
strmuloption(2) = "In the class of which it is a member"
strmuloption(3) = "In any object of the class of which it is a member"
strmuloption(4) = "In he public part of its class"

strmultiple(2) = "  Classes are useful because they -"
strmuloption(5) = "Are removed from memory when not in use"
strmuloption(6) = "Permit data to be hidden from other classes"
strmuloption(7) = "Bring together all aspects of an entity in one place"
strmuloption(8) = "Can closely model objects in the real world"

strmultiple(3) = "  Element double Array[7] is which element of the array ?"
strmuloption(9) = "The sixth"
strmuloption(10) = "The seventh"
strmuloption(11) = "The eight"
strmuloption(12) = "Impossible to tell"

strmultiple(4) = "  When a multidimensional array is accessed, each array index is -"
strmuloption(13) = "Separated by commas"
strmuloption(14) = "Surrounded by commas and surrounded by brackets"
strmuloption(15) = "Separated by commas and surrounded by brackets"
strmuloption(16) = "Surrounded by brackets"

strmultiple(5) = "  When an array name is passed to a function the function -"
strmuloption(17) = "Accesses exactly the same array as the calling program"
strmuloption(18) = "Accesses a copy of the array passed by the program"
strmuloption(19) = "Refers to the array using the same name as that used by the calling program"
strmuloption(20) = "Refers to the array using a different name from that used by the calling program"

strmultiple(6) = "  In a stack, data places on the stack first is -"
strmuloption(21) = "Not given an index number"
strmuloption(22) = "Given the index number 0"
strmuloption(23) = "The first data item to be removed"
strmuloption(24) = "The last data item to be removed"

strmultiple(7) = "  You can read input that consists of multiple lines of text using -"
strmuloption(25) = "The normal cout << combination"
strmuloption(26) = "The cin.get() function with one argument"
strmuloption(27) = "The cin.get() function with two argument"
strmuloption(28) = "The cin.get() function with three argument"

strmultiple(8) = "  Operator overloading is -"
strmuloption(29) = "Making C ++ operators work with bjects"
strmuloption(30) = "Giving C ++ operators more than that can handle"
strmuloption(31) = "Giving new meaning to existing C ++ operators"
strmuloption(32) = "Making new C++ operators"

strmultiple(9) = "  Private data members can be accessed -"
strmuloption(33) = "Only from the base can be accessed"
strmuloption(34) = "Both from he base class and from its derived classes"
strmuloption(35) = "From the class which is a friend of the base class"
strmuloption(36) = "None of the above is correct"

strmultiple(10) = " Protected data members can be accessed -"
strmuloption(37) = "Only from the base class itself"
strmuloption(38) = "Both from the base class and from its derived classes"
strmuloption(39) = "From the class which is a friend of the base class"
strmuloption(40) = "None of the above are correct"

strmulanswer(1) = "In the class of which it is a member"
strmulanswer(2) = "Permit data to be hidden from other classes"
strmulanswer(3) = "The eight"
strmulanswer(4) = "Separated by commas and surrounded by brackets"
strmulanswer(5) = "Accesses a copy of the array passed by the program"
strmulanswer(6) = "The last data item to be removed"
strmulanswer(7) = "The normal cout << combination"
strmulanswer(8) = "Giving new meaning to existing C ++ operators"
strmulanswer(9) = "Both from he base class and from its derived classes"
strmulanswer(10) = "Both from the base class and from its derived classes"
End Sub

Private Sub multiplechoice_Paper5()
strmultiple(1) = "  Public data members can be accessed -"
strmuloption(1) = "Only from the base class itself"
strmuloption(2) = "Both from the base class and from its derived classes"
strmuloption(3) = "From the class which is a friend of the base class"
strmuloption(4) = "None of the above are correct"

strmultiple(2) = "  Assume a class C with object obj1, obj2 and obj3. For the statement obj3 = obj1 - obj2 to work correctly, the overload operator must -"
strmuloption(5) = "Take two arguments"
strmuloption(6) = "Return a value"
strmuloption(7) = "Create a named temporary object"
strmuloption(8) = "Use the object of which the operator as an operand"

strmultiple(3) = "  When you overload an arithmetic assignment operator, the result -"
strmuloption(9) = "Goes in the object to the right of the operator"
strmuloption(10) = "Goes in the object to the left of the operator"
strmuloption(11) = "Goes in the object of which the operator is a member"
strmuloption(12) = "Must be returned"

strmultiple(4) = "  To convert from a user defined class to a basic type you would most likely use -"
strmuloption(13) = "A built in conversion function"
strmuloption(14) = "A one at fument constructor"
strmuloption(15) = "An overload = operator"
strmuloption(16) = "A conversion function that is member of the class"

strmultiple(5) = "  To convert from a basic type to a user defined class you would most likely use -"
strmuloption(17) = "A built in conversion function"
strmuloption(18) = "A one argument constructor"
strmuloption(19) = "An overload = operator"
strmuloption(20) = "Both (a) and (b)"

strmultiple(6) = "  New operator is used -"
strmuloption(21) = "To dynamically allocate storage"
strmuloption(22) = "To statically allocate storage"
strmuloption(23) = "To allocate storage for a new variable"
strmuloption(24) = "None of the above are correct"

strmultiple(7) = "  Delete operator is used -"
strmuloption(25) = "To allocate storage"
strmuloption(26) = "To deallocate storage"
strmuloption(27) = "To delete variable name"
strmuloption(28) = "None of the above is correct"

strmultiple(8) = "  Class members are _________________________ by default"
strmuloption(29) = "Private"
strmuloption(30) = "Protected"
strmuloption(31) = "Public"
strmuloption(32) = "None of the above is correct"

strmultiple(9) = "  Friend functions have access to the -"
strmuloption(33) = "Private and protected members"
strmuloption(34) = "Public members only"
strmuloption(35) = "Private members only"
strmuloption(36) = "None of the above is correct"

strmultiple(10) = " Inline functions _________________ call overload"
strmuloption(37) = "Increase"
strmuloption(38) = "Reduce"
strmuloption(39) = "None of the above is correct"
strmuloption(40) = "All of the above"

strmulanswer(1) = "From the class which is a friend of the base class"
strmulanswer(2) = "Create a named temporary object"
strmulanswer(3) = "Goes in the object to the left of the operator"
strmulanswer(4) = "A conversion function that is member of the class"
strmulanswer(5) = "Both (a) and (b)"
strmulanswer(6) = "To allocate storage for a new variable"
strmulanswer(7) = "To deallocate storage"
strmulanswer(8) = "Private"
strmulanswer(9) = "Public members only"
strmulanswer(10) = "Reduce"
End Sub

Private Sub multiplechoice_Paper6()
strmultiple(1) = "  The characteristic that data can be manipulated only through member functions that are part of the classis called -"
strmuloption(1) = "Encapsulation"
strmuloption(2) = "Data dependency"
strmuloption(3) = "Inheritance"
strmuloption(4) = "Instantiation"

strmultiple(2) = "  What is the value of green after the execution of the following function enum colors{red, blue, white = 5, yellow, green, pink}"
strmuloption(5) = "2"
strmuloption(6) = "5"
strmuloption(7) = "7"
strmuloption(8) = "1"

strmultiple(3) = "  In a class, only the member function can access data which is not accessible to outside. This feature is called "
strmuloption(9) = "Data security"
strmuloption(10) = "Data hiding"
strmuloption(11) = "Data manipulation"
strmuloption(12) = "Data definition"

strmultiple(4) = "  This is one of the reasons that C programs may not run with C++ compilers :"
strmuloption(13) = "In C++ nesting of functions is not allowed where as this is allowed in C"
strmuloption(14) = "Enum is allowed in C++ which is not there in C"
strmuloption(15) = "Data hiding and inheritance are not available in C"
strmuloption(16) = "C and C++ are two different languages"

strmultiple(5) = "  The scope resolution operator is -"
strmuloption(17) = "::"
strmuloption(18) = ";"
strmuloption(19) = "<<"
strmuloption(20) = "->"

strmultiple(6) = "  The signature of a function is -"
strmuloption(21) = "The return type"
strmuloption(22) = "The number and type of arguments"
strmuloption(23) = "The class of a function"
strmuloption(24) = "None of the above"

strmultiple(7) = "  Pick out the most appropriate statement -"
strmuloption(25) = "All variables must be declared before the are used"
strmuloption(26) = "Variables in C++ need not be declared and te type can be assigned dynamically"
strmuloption(27) = "Variables in C++ can be declared at the end of the program (before the main function terminates)"
strmuloption(28) = "Variables can not be used explicitly in C++"

strmultiple(8) = "  What is the value of Friday in the following -enum days {Monday, uesday,Wednesday = -1, Thursday,Friday,Saturday = 6, Sunday}"
strmuloption(29) = "1"
strmuloption(30) = "5"
strmuloption(31) = "0"
strmuloption(32) = "-3"

strmultiple(9) = "  Ortogonality in a programming language means -"
strmuloption(33) = "Ability to use recursive functions"
strmuloption(34) = "The ability of expressions to be used in other expressions."
strmuloption(35) = "Ability to convert recursive to non recursive procedures"
strmuloption(36) = "Use of functions rather than subroutines"

strmultiple(10) = " What is the output of the following program segment -for(i=1,j=0;i<10;i++)  j+ = I;cout << i << \n"
strmuloption(37) = "10"
strmuloption(38) = "20"
strmuloption(39) = "55"
strmuloption(40) = "there is an error in the program"

strmulanswer(1) = "Encapsulation"
strmulanswer(2) = "7"
strmulanswer(3) = "Data hiding"
strmulanswer(4) = "In C++ nesting of functions is not allowed where as this is allowed in C"
strmulanswer(5) = "::"
strmulanswer(6) = "The number and type of arguments"
strmulanswer(7) = "All variables must be declared before the are used"
strmulanswer(8) = "1"
strmulanswer(9) = "The ability of expressions to be usd in other expressions."
strmulanswer(10) = "55"

End Sub

Private Sub multiplechoice_Paper7()
strmultiple(1) = "  The string table in C++ holds the -"
strmuloption(1) = "String constants in your program"
strmuloption(2) = "Program statements in string form"
strmuloption(3) = "Variables whose type is of string"
strmuloption(4) = "None of the above"

strmultiple(2) = "  The declaration int **var1; shows that -"
strmuloption(5) = "Var1 can not be accessed"
strmuloption(6) = "Var1 is a pointer to a pointer of type int"
strmuloption(7) = "Var1 is a protected data type of integer"
strmuloption(8) = "This type declaration shows an error"

strmultiple(3) = "  In C++, the keyword auto can be used for -"
strmuloption(9) = "Automatic assignment of data to objects during instantiation"
strmuloption(10) = "Automatic all of a function"
strmuloption(11) = "To declare a local variable"
strmuloption(12) = "It is not a keyword in c++"

strmultiple(4) = "  By default, C ++ uses the following method of passing arguments -"
strmuloption(13) = "Call by reference"
strmuloption(14) = "Call by value"
strmuloption(15) = "Value by pointer"
strmuloption(16) = "None of the above"

strmultiple(5) = "  Pick out the most appropriate statement from the following -"
strmuloption(17) = "References are pointers"
strmuloption(18) = "Array of references can be created"
strmuloption(19) = "You can not reference a reference variable"
strmuloption(20) = "All of the above"

strmultiple(6) = "  Find out the error in the following statement -int myfunc1(float var1, int k, int I = 0, int j)"
strmuloption(21) = "The declaration int I = 0 should not be used"
strmuloption(22) = "This function can not return int and should be declared void"
strmuloption(23) = "Since the arable I has been given default value, all other variables also should be given default values"
strmuloption(24) = "Since the variable I has been given default value, variable j also should be gien default value"

strmultiple(7) = "  The ? can be used to replace -"
strmuloption(25) = "If-else statements"
strmuloption(26) = "Wild character"
strmuloption(27) = "No meaning in C++"
strmuloption(28) = "Returns the value"

strmultiple(8) = "  One of the ternary operators provided in C++ is-"
strmuloption(29) = "**"
strmuloption(30) = "::"
strmuloption(31) = "<<"
strmuloption(32) = "?"

strmultiple(9) = "  Dynamic memory can be allocated by the following declaration -"
strmuloption(33) = "New"
strmuloption(34) = "Volatile"
strmuloption(35) = "Static"
strmuloption(36) = "=="

strmultiple(10) = " The member of a structure can be directly accessed by -"
strmuloption(37) = "->"
strmuloption(38) = "dot operator"
strmuloption(39) = "::"
strmuloption(40) = ">>"

strmulanswer(1) = "String constants in your program"
strmulanswer(2) = "Var1 is a pointer to a pointer of type int"
strmulanswer(3) = "To declare a local variable"
strmulanswer(4) = "Call by value"
strmulanswer(5) = "You can not reference a reference variable"
strmulanswer(6) = "Since the variable I has been given default value, variable j also should be gien default value"
strmulanswer(7) = "If-else statements"
strmulanswer(8) = "?"
strmulanswer(9) = "New"
strmulanswer(10) = "dot operator"

End Sub

Private Sub multiplechoice_Paper8()
strmultiple(1) = "  The member of a structure can be accessed through a pointer by -"
strmuloption(1) = "->"
strmuloption(2) = "dot operator"
strmuloption(3) = "::"
strmuloption(4) = ">>"

strmultiple(2) = "  The members of a class can be made private by -"
strmuloption(5) = "Declaring them private"
strmuloption(6) = "By default they are private"
strmuloption(7) = "By declaring them in the beginning of the program immediately after man()"
strmuloption(8) = "They are always public"

strmultiple(3) = "  A function that is called automatically when an object is created is known as -"
strmuloption(9) = "Instantiation"
strmuloption(10) = "Function prototype"
strmuloption(11) = "Constructor"
strmuloption(12) = "Structure"

strmultiple(4) = "  A function that is called automatically when an object is destroyed is known as -"
strmuloption(13) = "Instantiation"
strmuloption(14) = "Function prototype"
strmuloption(15) = "Constructor"
strmuloption(16) = "Destructor"

strmultiple(5) = "  The difference between a class and a structure can be summarized as -"
strmuloption(17) = "Absolutely no relation between the two"
strmuloption(18) = "Class members are by default private and can become public any explicit declaration where as structure members are by default public"
strmuloption(19) = "Both are same in all respects"
strmuloption(20) = "Structure is a data type and class represents group of objects"

strmultiple(6) = "  It is possible to allow non member function access to private members of a class by declaring it as -"
strmuloption(21) = "Public "
strmuloption(22) = "Friend"
strmuloption(23) = "Private"
strmuloption(24) = "Not possible"

strmultiple(7) = "  The immediate termination of execution of a programmer can be made by calling -"
strmuloption(25) = "Exit ()"
strmuloption(26) = "Return"
strmuloption(27) = "Stop"
strmuloption(28) = "Abort"

strmultiple(8) = "  When one object initializes another object the following function is invoked -"
strmuloption(29) = "Copy constructor "
strmuloption(30) = "New"
strmuloption(31) = "Instantiation"
strmuloption(32) = "None of the above"

strmultiple(9) = "  The copy operation by a simple assignment causes in some cases-="
strmuloption(33) = "No effect"
strmuloption(34) = "Hollow copy"
strmuloption(35) = "Copy exactly the same as original"
strmuloption(36) = "Creates new"

strmultiple(10) = " A base class is inherited by -"
strmuloption(37) = "Derived class"
strmuloption(38) = "Inline function"
strmuloption(39) = "Constructor"
strmuloption(40) = "None of the above"

strmulanswer(1) = "->"
strmulanswer(2) = "By default they are private"
strmulanswer(3) = "Constructor"
strmulanswer(4) = "Destructor"
strmulanswer(5) = "Class members are by default private and can become public any explicit declaration where as structure members are by default public"
strmulanswer(6) = "Friend"
strmulanswer(7) = "Exit ()"
strmulanswer(8) = "Copy constructor "
strmulanswer(9) = "Hollow copy"
strmulanswer(10) = "Derived class"

End Sub

Private Sub multiplechoice_Paper9()
strmultiple(1) = "  By using protected one can crate class members that are private to their class but -"
strmuloption(1) = "That can not be inherited and accessed by a derived class"
strmuloption(2) = "That can still be inherited and accessed by a derived class"
strmuloption(3) = "That can be public"
strmuloption(4) = "None of the above"

strmultiple(2) = "  A class that includes a virtual function is known as -"
strmuloption(5) = "Virtual "
strmuloption(6) = "Polymorphic"
strmuloption(7) = "Derived"
strmuloption(8) = "Base"

strmultiple(3) = "  A class that inherits a base class containing virtual function is called -"
strmuloption(9) = "Virtual"
strmuloption(10) = "Polymorphic"
strmuloption(11) = "Derived"
strmuloption(12) = "Base"

strmultiple(4) = "  The version of a virtual function actually its called during run time is based solely upon the type of the object that is -"
strmuloption(13) = "Being instantiated and processed by te function"
strmuloption(14) = "In the derived class"
strmuloption(15) = "Beng pointed to by a cbase class pointer"
strmuloption(16) = "Being passed to the function"

strmultiple(5) = "  A pure virtual function is a virtual function that has -"
strmuloption(17) = "No definition in its base class"
strmuloption(18) = "A definition in its base class"
strmuloption(19) = "A definition in at least one derived class"
strmuloption(20) = "Definition in base class and at least one derived class"

strmultiple(6) = "  A class that contains at least one pure virtual function is called as -"
strmuloption(21) = "Pure class"
strmuloption(22) = "Abstract class"
strmuloption(23) = "Base class"
strmuloption(24) = "Derived class"

strmultiple(7) = "  The binding that binds a function call at run time is called -"
strmuloption(25) = "Early binding "
strmuloption(26) = "Run time binding"
strmuloption(27) = "Late binding"
strmuloption(28) = "Linking"

strmultiple(8) = "  One of the major disadvantage with late binding is -"
strmuloption(29) = "The source code should be made available at compile time "
strmuloption(30) = "The program runs slower"
strmuloption(31) = "Dynamic variables can not be used I the program"
strmuloption(32) = "Static variables can not be used"

strmultiple(9) = "  One of the important features of an abstract class is -"
strmuloption(33) = "It need not have any objects"
strmuloption(34) = "It should not have any members"
strmuloption(35) = "It should not have any members"
strmuloption(36) = "None of the above"

strmultiple(10) = " The advantage of declaring a virtual function as pure, is -"
strmuloption(37) = "You force derive any derived class to define its own implementation"
strmuloption(38) = "You can avoid making the class an abstract class"
strmuloption(39) = "Programs runs faster"
strmuloption(40) = "Early binding can be achieved"


strmulanswer(1) = "That can still be inherited and accessed by a derived class"
strmulanswer(2) = "Polymorphic"
strmulanswer(3) = "Polymorphic"
strmulanswer(4) = "Being pointed to by a cbase class pointer"
strmulanswer(5) = "No definition in its base class"
strmulanswer(6) = "Abstract class"
strmulanswer(7) = "Late binding"
strmulanswer(8) = "The program runs slower"
strmulanswer(9) = "It need not have any objects"
strmulanswer(10) = "You force derive any derived class to define its own implementation"

End Sub

Private Sub multiplechoice_Paper10()
strmultiple(1) = "  What is the output of the following statements ?cout.setf(ios:: oct , ios ::  basefield);cout << 30 << | n;"
strmuloption(1) = "29"
strmuloption(2) = "o36"
strmuloption(3) = "error"
strmuloption(4) = "08"

strmultiple(2) = "  The flag scientific in iostream indicates -"
strmuloption(5) = "Format floating point values in scientific form"
strmuloption(6) = "It allows scientific calculations in double precision"
strmuloption(7) = "Converts all numbers into float or real"
strmuloption(8) = "All of the above"

strmultiple(3) = "  In C++  iostream, the format flags can be reset by -    "
strmuloption(9) = "Resetf()"
strmuloption(10) = "Unsetf()"
strmuloption(11) = "Restiflags()"
strmuloption(12) = "Imaskf()"

strmultiple(4) = "  The standard stream classes can be -"
strmuloption(13) = "Extended"
strmuloption(14) = "Redefined"
strmuloption(15) = "Used in C programs also"
strmuloption(16) = "User defined"

strmultiple(5) = "  In C++, the stream base class is -"
strmuloption(17) = "Iostream"
strmuloption(18) = "Iofstream"
strmuloption(19) = "Ios"
strmuloption(20) = "Stdio"

strmultiple(6) = "  The template function declaration specifies-"
strmuloption(21) = "Template class"
strmuloption(22) = "A generic class"
strmuloption(23) = "Exception"
strmuloption(24) = "Identifier"

strmultiple(7) = "  In C++, the exception handler is invoke with a -"
strmuloption(25) = "Try block"
strmuloption(26) = "Throw expression"
strmuloption(27) = "Catch function"
strmuloption(28) = "Abort()"

strmultiple(8) = "  The exception is processed using -"
strmuloption(29) = "Unexpected ()"
strmuloption(30) = "Perform()"
strmuloption(31) = "Catch()"
strmuloption(32) = "Try()"

strmultiple(9) = "  If no appropriate handler is available, the function invoked to process the exception is -"
strmuloption(33) = "Unexpected()"
strmuloption(34) = "Exit()"
strmuloption(35) = "Abort()"
strmuloption(36) = "Stop()"

strmultiple(10) = "In exception handling, the control is passed to the first handler(catch) whose exception declaration signature matches -"
strmuloption(37) = "Try block"
strmuloption(38) = "Universal handler"
strmuloption(39) = "The thrown expression"
strmuloption(40) = "None of these"

strmulanswer(1) = "o36"
strmulanswer(2) = "It allows scientific calculations in double precision"
strmulanswer(3) = "Restiflags()"
strmulanswer(4) = "Extended"
strmulanswer(5) = "Ios"
strmulanswer(6) = "A generic class"
strmulanswer(7) = "Throw expression"
strmulanswer(8) = "Catch()"
strmulanswer(9) = "Unexpected()"
strmulanswer(10) = "The thrown expression"
End Sub




