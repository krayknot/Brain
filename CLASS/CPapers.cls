VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPapers"
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
 
Public Function C(PaperType As String, CPaper As Paper, mCat As Category, Number As Integer)
'ptype is the paper type
'Cpaper is the number of paper as paper1,paper2 ...
'mcat is the category
'number is the number of questions, or options or answers to be taken

  StrPaperType = PaperType
' StrPaper = Paper
If PaperType = "Fill in the Blanks" Then
   If CPaper = Paper1 Then Fill_in_the_Blanks_Paper1
   If CPaper = Paper2 Then Fill_in_the_Blanks_Paper2
   If CPaper = Paper3 Then Fill_in_the_Blanks_Paper3
   If CPaper = Paper4 Then Fill_in_the_Blanks_Paper4
   If CPaper = Paper5 Then Fill_in_the_Blanks_Paper5
   If CPaper = Paper6 Then Fill_in_the_Blanks_Paper6
   If CPaper = Paper7 Then Fill_in_the_Blanks_Paper7
   If CPaper = Paper8 Then Fill_in_the_Blanks_Paper8
   If CPaper = Paper9 Then Fill_in_the_Blanks_Paper9
   If CPaper = Paper10 Then Fill_in_the_Blanks_Paper10
   If CPaper = Paper11 Then Fill_in_the_Blanks_Paper11
      
ElseIf PaperType = "Matching Columns" Then
       If CPaper = Paper1 Then Matching_Columns_Paper1
       If CPaper = Paper2 Then Matching_Columns_Paper2
       If CPaper = Paper3 Then Matching_Columns_Paper3
       If CPaper = Paper4 Then Matching_Columns_Paper4
       If CPaper = Paper5 Then Matching_Columns_Paper5
       If CPaper = Paper6 Then Matching_Columns_Paper6
       If CPaper = Paper7 Then Matching_Columns_Paper7
       If CPaper = Paper8 Then Matching_Columns_Paper8
       If CPaper = Paper9 Then Matching_Columns_Paper9
       If CPaper = Paper10 Then Matching_Columns_Paper10
       If CPaper = Paper11 Then Matching_Columns_Paper11
              
ElseIf PaperType = "True Or False" Then
       If CPaper = Paper1 Then TrueorFalse_Paper1
       If CPaper = Paper2 Then TrueorFalse_Paper2
       If CPaper = Paper3 Then TrueorFalse_Paper3
       If CPaper = Paper4 Then TrueorFalse_Paper4
       If CPaper = Paper5 Then TrueorFalse_Paper5
       If CPaper = Paper6 Then TrueorFalse_Paper6
       If CPaper = Paper7 Then TrueorFalse_Paper7
       If CPaper = Paper8 Then TrueorFalse_Paper8
       If CPaper = Paper9 Then TrueorFalse_Paper9
       If CPaper = Paper10 Then TrueorFalse_Paper10
       If CPaper = Paper11 Then TrueorFalse_Paper11
       If CPaper = Paper12 Then TrueorFalse_Paper12
       If CPaper = Paper13 Then TrueorFalse_Paper13
       If CPaper = Paper14 Then TrueorFalse_Paper14
              
ElseIf PaperType = "Descriptive Questions" Then
       If CPaper = Paper1 Then questions_paper1
       If CPaper = Paper2 Then questions_paper2
       If CPaper = Paper3 Then questions_paper3
       If CPaper = Paper4 Then questions_paper4
       If CPaper = Paper5 Then questions_paper5
       If CPaper = Paper6 Then questions_paper6
       If CPaper = Paper7 Then questions_paper7
       If CPaper = Paper8 Then questions_paper8
       If CPaper = Paper9 Then questions_paper9
       If CPaper = Paper10 Then questions_paper10

ElseIf PaperType = "Multiple Choice" Then
       If CPaper = Paper1 Then multiplechoice_Paper1
       If CPaper = Paper2 Then multiplechoice_Paper2
       If CPaper = Paper3 Then multiplechoice_Paper3
       If CPaper = Paper4 Then multiplechoice_Paper4
       If CPaper = Paper5 Then multiplechoice_Paper5
       If CPaper = Paper6 Then multiplechoice_Paper6
       If CPaper = Paper7 Then multiplechoice_Paper7
       If CPaper = Paper8 Then multiplechoice_Paper8
       If CPaper = Paper9 Then multiplechoice_Paper9
       If CPaper = Paper10 Then multiplechoice_Paper10
       If CPaper = Paper11 Then multiplechoice_Paper11
       If CPaper = Paper12 Then multiplechoice_Paper12
       If CPaper = Paper13 Then multiplechoice_Paper13
       If CPaper = Paper14 Then multiplechoice_Paper14
       If CPaper = Paper15 Then multiplechoice_Paper15
End If

If mCat = Fill_Questions Then
   C = StrQuestions(Number)
ElseIf mCat = Fill_Options Then
       C = stroptions(Number)
ElseIf mCat = Fill_Answers Then
       C = stranswers(Number)
       
ElseIf mCat = Matching_ColumnA Then
       C = strcolumnA(Number)
ElseIf mCat = Matching_ColumnB Then
       C = strcolumnB(Number)
ElseIf mCat = Matching_Answers Then
       C = stranswers(Number)
       
ElseIf mCat = Questions Then
       C = StrQuestions(Number)
       
ElseIf mCat = TF_Questions Then
       C = StrQuestions(Number)
ElseIf mCat = TF_Answers Then
       C = stranswers(Number)
       
ElseIf mCat = MC_Questions Then
       C = strmultiple(Number)
       
ElseIf mCat = MC_Options_1 Then
       C = strmuloption(Number)
       
ElseIf mCat = MC_Options_2 Then
       C = strmuloption(Number)
       
ElseIf mCat = MC_Options_3 Then
       C = strmuloption(Number)
       
ElseIf mCat = MC_Options_4 Then
       C = strmuloption(Number)
       
ElseIf mCat = MC_Answer Then
       C = strmulanswer(Number)
End If

End Function

Private Sub Fill_in_the_Blanks_Paper1()
StrQuestions(1) = "Return statement return only dash expression."
StrQuestions(2) = "The else portion of the if-else statement is dash ."
StrQuestions(3) = "The dash statement is used to bypass the remainder of the current pass through a loop."
StrQuestions(4) = "The dash statement is used to alter normal sequence of execution by transferring control to some other part of program."
StrQuestions(5) = "An dash can be assigbed initial values by including appropriate expression within the variable declarations."
StrQuestions(6) = "An automatic variable does not retain dash once control is transferred out of its defining function."
StrQuestions(7) = "Scope of an automatic variable can be dash  than the entire function."
StrQuestions(8) = "An external variable declaration must begin with storage class specifierdash ."
StrQuestions(9) = "If the libe 'int sp;' occurs outside any function, it dash the external variable sp."
StrQuestions(10) = "If the libe' static char buffer[max];'appears in obe file of a program, then the variable name will not conflict with the same name in dash  of the same program."

stroptions(1) = "Other "
stroptions(2) = "Value"
stroptions(3) = "Extern"
stroptions(4) = "Automatic variable"
stroptions(5) = "Smaller"
stroptions(6) = "Continue"
stroptions(7) = "Defibes"
stroptions(8) = "Single-valued"
stroptions(9) = "Optional "
stroptions(10) = "goto"

stranswers(1) = "Single Valued"
stranswers(2) = "Optional"
stranswers(3) = "Continue"
stranswers(4) = "GoTo"
stranswers(5) = "Automatic Variable"
stranswers(6) = "Value"
stranswers(7) = "Smaller"
stranswers(8) = "Extern"
stranswers(9) = "Defibes"
stranswers(10) = "Other Files"

End Sub

Private Sub Fill_in_the_Blanks_Paper2()
StrQuestions(1) = "An external variable definition must not begin with storage class specifier dash ."
StrQuestions(2) = "A register variable is expected to be placed in machine  dash ."
StrQuestions(3) = "For dash  and static variables initializes must be constant expressions."
StrQuestions(4) = "When a function calls itself, each invocation gets afresh set of dash  variables, independent of the previous invocation."
StrQuestions(5) = "#define statement calls for simplest type of dash  substitution."
StrQuestions(6) = "The function sprintf places the result in dash  instead of on the standard output."
StrQuestions(7) = "printf(" & Chr(4) & "%.*s" & Chr(4) & ",max,s) can be used to print at most dash  characters from a string s."
StrQuestions(8) = "Array elements can be accessed dash ."
StrQuestions(9) = "Loop invariants are assertions that remain dash  before and after execution of loops."
StrQuestions(10) = "On executing f = !(i > 10), f will have values dash "


stroptions(1) = "Macro"
stroptions(2) = "Obe or zero"
stroptions(3) = "Randomly"
stroptions(4) = "External"
stroptions(5) = "Max"
stroptions(6) = "True"
stroptions(7) = "String"
stroptions(8) = "Automatic"
stroptions(9) = "Extern"
stroptions(10) = "Extern"

stranswers(1) = "Extern"
stranswers(2) = "Register"
stranswers(3) = "External"
stranswers(4) = "Automatic"
stranswers(5) = "Macro"
stranswers(6) = "String"
stranswers(7) = "max"
stranswers(8) = "Randomly"
stranswers(9) = "True"
stranswers(10) = "Obe or Zero"

End Sub

Private Sub Fill_in_the_Blanks_Paper3()
StrQuestions(1) = "Linked lists can be created dash ."
StrQuestions(2) = "The termination condition for defining a Fibonacci Sequence is dash ."
StrQuestions(3) = "The logical operator == checks for dash  of two values."
StrQuestions(4) = "In searching list of n items, sequential search will make dash  comparisons in the worst case."
StrQuestions(5) = "If an expression x = e1&&e2&&e3, then if e1 is false,e2 and e3 dash  be evaluated."
StrQuestions(6) = "All characters are represented internally as dash ."
StrQuestions(7) = "The operator i << 2 is equivalent to dash  i by 4."
StrQuestions(8) = "In the following fragment, to find the square root of real number m, what is te invariant for the repeat until loop?" & Chr(13) & _
    "{assert : m > 0. AND. sq1 = m/2}" & Chr(13) & _
    "sq2 = m/2;" & Chr(13) & _
    "{in variant dash }" & Chr(13) & _
    "do(" & Chr(13) & _
     "sq1 = sq2;" & Chr(13) & _
     "sq2 = (sq1 + m/sq1)/2;}" & Chr(13) & _
    "while fabs(sq1-sq2) < error;"
StrQuestions(9) = "The complexity of binary search algorithm is of the order of dash ."
StrQuestions(10) = "The statement " & Chr(13) & _
    "if (ch > = 'a' && ch <= 'z') return ('a' + ch - 'a' )" & Chr(4) & ", will return dash  for a given character between 'a' and 'z'."

stroptions(1) = "n"
stroptions(2) = "F(1)=F(2)=1"
stroptions(3) = "Integers"
stroptions(4) = "Capital"
stroptions(5) = "Multiply"
stroptions(6) = "logn"
stroptions(7) = "sq1 not equal to sq2"
stroptions(8) = "Equality"
stroptions(9) = "Will not"
stroptions(10) = "Dynamically"

stranswers(1) = "Dynamically"
stranswers(2) = "F(1) = F(2) = 1"
stranswers(3) = "Equality"
stranswers(4) = "n"
stranswers(5) = "will not"
stranswers(6) = "integers"
stranswers(7) = "multiply"
stranswers(8) = "sq1 not equal to sq2"
stranswers(9) = "logn"
stranswers(10) = "Capital"

End Sub

Private Sub Fill_in_the_Blanks_Paper4()
StrQuestions(1) = "The statement 'typedef int length' declares length as dash ."
StrQuestions(2) = "In the printf statement between the % and the conversion character there may be an h if the integer is to b printed as dash ."
StrQuestions(3) = "The type var-list is used to declare variables that refer to each dash  of a function in turn."
StrQuestions(4) = "An output stream dash  is associated with a program for error reporting."
StrQuestions(5) = "int dash (FILE *fp) returns non zero if error occurred on the stream."
StrQuestions(6) = "FILE is defibed witha dash  statement."
StrQuestions(7) = "The function dash (char *-s) excites the command contained in the characer string S."
StrQuestions(8) = "It is a bad error to dash  something not obtained by calling calloc or malloc."
StrQuestions(9) = "File descriptor is an dash ."
StrQuestions(10) = "The function lseek provide a way to read or write a file in dash  order."

stroptions(1) = "stderr"
stroptions(2) = "argument"
stroptions(3) = "arbitrary"
stroptions(4) = "free"
stroptions(5) = "ttpedef"
stroptions(6) = "integer"
stroptions(7) = "an integer type"
stroptions(8) = "system"
stroptions(9) = "ferror"
stroptions(10) = "short"

stranswers(1) = "An Integer Type"
stranswers(2) = "short"
stranswers(3) = "argument"
stranswers(4) = "stderr"
stranswers(5) = "ferror"
stranswers(6) = "typedef"
stranswers(7) = "system"
stranswers(8) = "free"
stranswers(9) = "integer"
stranswers(10) = "arbitrary"

End Sub

Private Sub Fill_in_the_Blanks_Paper5()
StrQuestions(1) = "Pointer arguments enable a function to access and dash  objects defined in the calling routine."
StrQuestions(2) = "double xx(void) meant that function xx has dash ."
StrQuestions(3) = "TO check a pointer variable (p) whether it currently refers to a valid dynamic variable,m we can use the following statement; if (dash )."
StrQuestions(4) = "If p is a pointer variable, statements" & Chr(13) & _
    "if(dash ) and if (p ! =NULL)" & Chr(13) & _
    "are equivalent."
StrQuestions(5) = "A polynomial like 3x5 - 2x3 + x2 can be represented using dash ."
StrQuestions(6) = " Deletion of node from a singly linked list , in generaly requires adjustment of pointer field of dash node(s)."
StrQuestions(7) = "Insertion of a node to a singly linked list, in general, requires adjustment of pointer field of dash  node(s)."
StrQuestions(8) = "Insertion of node at the beginning of a link list requires adjustment ofpointer field of dash node(s)."
StrQuestions(9) = "A contiguous list can be implemented using an dash ."
StrQuestions(10) = "We need to search for a correct place to insert an item in dash  sort."

stroptions(1) = "linked list"
stroptions(2) = "insertion"
stroptions(3) = "obe"
stroptions(4) = "head"
stroptions(5) = "two"
stroptions(6) = "change"
stroptions(7) = "p!=NULL"
stroptions(8) = "arrray"
stroptions(9) = "p"
stroptions(10) = "no argument"

stranswers(1) = "change"
stranswers(2) = "no argument"
stranswers(3) = "p!=NULL"
stranswers(4) = "p"
stranswers(5) = "linked list"
stranswers(6) = "obe"
stranswers(7) = "two"
stranswers(8) = "head"
stranswers(9) = "array"
stranswers(10) = "insertion"

End Sub

Private Sub Fill_in_the_Blanks_Paper6()
StrQuestions(1) = "Verifying that a list of n items is in the correct order (ascending or descending order) requires at least dash  comparison operations."
StrQuestions(2) = "The idea of dividing a problem into smaller and similar sub problems is called dash ."
StrQuestions(3) = "dash  is an example of sorting method which is designed sing divide and cinquer srategy."
StrQuestions(4) = "The function getchar() returns dash  when there is no ore input character."
StrQuestions(5) = "The data type double is actually double precision dash ."
StrQuestions(6) = "Operators *,/ and % have dash  precedence."
StrQuestions(7) = "Operators &,^,== have dash  precedence."
StrQuestions(8) = "x = f() + g();" & Chr(13) & _
                  "In this statement and g may be evaluated in dash  order."
StrQuestions(9) = "An dash  converts assembly language program to machine code."
StrQuestions(10) = "Braces { and } are used to group declarations and statements together into a dash  statement."

stroptions(1) = "any"
stroptions(2) = "compound"
stroptions(3) = "float"
stroptions(4) = "different"
stroptions(5) = "assembler"
stroptions(6) = "n-1"
stroptions(7) = "quick sort"
stroptions(8) = "same"
stroptions(9) = "EOF"
stroptions(10) = "divide and conquer"

stranswers(1) = "n-1"
stranswers(2) = "divide and conquer"
stranswers(3) = "quick sort"
stranswers(4) = "EOF"
stranswers(5) = "float"
stranswers(6) = "same"
stranswers(7) = "different"
stranswers(8) = "any"
stranswers(9) = "assembler"
stranswers(10) = "compound"

End Sub
Private Sub Fill_in_the_Blanks_Paper7()
StrQuestions(1) = "Any pointer can be cast to dash  without loss of information."
StrQuestions(2) = "The operator size of returns size of data type in dash ."
StrQuestions(3) = "for (i=o;i < N; i==)" & Chr(13) & _
    "{" & Chr(13) & _
    " " & Chr(13) & _
    "}" & Chr(13) & _
    "The above statement is typically used when N( the number of dash ) is known."
StrQuestions(4) = "switch (flag) {" & Chr(13) & _
            "}" & Chr(13) & _
    "In the above statement the variable flag should have dash  values."
StrQuestions(5) = "switch(choice = getchar())" & Chr(13) & _
    "{" & Chr(13) & _
        "case 'R': printf (" & Chr(4) & "red\n" & Chr(4) & ");" & Chr(13) & _
        "-------------;" & Chr(13) & _
    "default : printf(" & Chr(4) & "color\n" & Chr(4) & ");" & Chr(13) & _
    "}"
StrQuestions(6) = "Fill in the blanks for making the sentence correct :" & Chr(4) & _
    "for((count = \0dash  count 2 = N) dash  count <=(N-1)/2; count++)7. 7. struct{" & Chr(4) & _
        "unsigned a :1;" & Chr(13) & _
        "unsigned b: 3;" & Chr(13) & _
        "unsigned c: 2;" & Chr(13) & _
        "unsigned d: 1;" & Chr(13) & _
       "} v;" & Chr(13) & _
    "The above statement declares variable v as a structure with four dash. "
StrQuestions(7) = ""
StrQuestions(8) = "Given that a is an unsigned integer variable whose value is 0X 6db7, the expression." & Chr(13) & _
        "b = a << 6;" & Chr(13) & _
    "will assign b the value dash ."
StrQuestions(9) = "An array declared as int a[5][5][5];. If the address of the element a[1][3][2] is x, then the address of the element a[1][2][3] is dash  than x."
StrQuestions(10) = "The value returned by the following function, when it is called with an initial values 5 is dash ." & Chr(13) & _
    "int f(int i)" & Chr(13) & _
    "{" & Chr(13) & _
        "if(i<1) return 1 ;" & Chr(13) & _
        "else return(i+f(i-1));" & Chr(13) & _
    "}"

stroptions(1) = "bit-fields"
stroptions(2) = ", and ;"
stroptions(3) = "beak"
stroptions(4) = "less"
stroptions(5) = "iterations"
stroptions(6) = "16"
stroptions(7) = "integer"
stroptions(8) = "bytes"
stroptions(9) = "0 x 6dc0"
stroptions(10) = "void"

stranswers(1) = "void"
stranswers(2) = "bytes"
stranswers(3) = "iterations"
stranswers(4) = "integer"
stranswers(5) = "break"
stranswers(6) = ". And ;"
stranswers(7) = "bit fields"
stranswers(8) = "0 x 6dc0"
stranswers(9) = "less"
stranswers(10) = "16"

End Sub

Private Sub Fill_in_the_Blanks_Paper8()
StrQuestions(1) = "dash  refers to the passing of a variable's address to a function."
StrQuestions(2) = "dash  is the bitwise inclusive or operator."
StrQuestions(3) = "dash  data structure is used to implement recursion."
StrQuestions(4) = "To prevent the use of functions across different files, dash  storage class  is used."
StrQuestions(5) = "To share a variable across different files, dash  storage class is used."
StrQuestions(6) = "When the dash  statement is executed, the program skips the rest of the ststements in the loop and goes back to test the loop condition."
StrQuestions(7) = "The expression 11%3 evaluates to dash ."
StrQuestions(8) = "A function that does no return anything has return type dash ."
StrQuestions(9) = "A pointer to void can hold pointer to dash ."
StrQuestions(10) = "The contents of two pointers that point to adjacent variable of type float different by dash ."

stroptions(1) = "String"
stroptions(2) = "2"
stroptions(3) = "extern"
stroptions(4) = ""
stroptions(5) = "stack"
stroptions(6) = "void"
stroptions(7) = "I"
stroptions(8) = "static"
stroptions(9) = "call by reference"
stroptions(10) = "continue"

stranswers(1) = "call by reference"
stranswers(2) = "I"
stranswers(3) = "stack"
stranswers(4) = "static"
stranswers(5) = "extern"
stranswers(6) = "continue"
stranswers(7) = "2"
stranswers(8) = "void"
stranswers(9) = "string"
stranswers(10) = ""

End Sub

Private Sub Fill_in_the_Blanks_Paper9()
StrQuestions(1) = "The dash  operator is unary operator to find the value of a variable."
StrQuestions(2) = "? : is dash  operator."
StrQuestions(3) = "Dividing a number by 256 is equivalent to shifting the dash  of number to the right."
StrQuestions(4) = "The dash  operator always appears between name of structure variable and obe of its members."
StrQuestions(5) = "Every C program automatically opens at least dash  streams."
StrQuestions(6) = "In an C-expression with mixed data type containing variables of the float and double the result is on the type dash ."
StrQuestions(7) = "The unary operator ~ ca be used for performing dash  on a variable."
StrQuestions(8) = "The following C- expression" & Chr(13) & _
                  "1+2*3-5" & Chr(13) & _
                  "evaluates to the value dash ."
StrQuestions(9) = "Consider the following program" & Chr(13) & _
                  "#define Flag 01" & Chr(13) & _
                  "#define Flag 02" & Chr(13) & _
                  "main()" & Chr(13) & _
                  "{" & Chr(13) & _
                  "int status = 03;" & Chr(13) & _
                  "status &=~ (FLag1 | FLag2);" & Chr(13) & _
                  "}" & Chr(13) & _
                  "The value of the status at the end of the program is dash ."
StrQuestions(10) = "Each pointer is dash  to the type of variable to which it pointers."

stroptions(1) = "bit-wise"
stroptions(2) = "zero"
stroptions(3) = "double "
stroptions(4) = "2"
stroptions(5) = "obe's complement "
stroptions(6) = "bound"
stroptions(7) = ""
stroptions(8) = "&"
stroptions(9) = "8 bits"
stroptions(10) = "3"

stranswers(1) = "&"
stranswers(2) = "bit-wise"
stranswers(3) = "8 bits"
stranswers(4) = ""
stranswers(5) = "3"
stranswers(6) = "double"
stranswers(7) = "obe's complement"
stranswers(8) = "2"
stranswers(9) = "zero"
stranswers(10) = "bound"

End Sub

Private Sub Fill_in_the_Blanks_Paper10()
StrQuestions(1) = "Consider the following" & Chr(13) & _
                  "int X = 12" & Chr(13) & _
                  "X %=5;" & Chr(13) & _
                  "The value of X becomes dash ."
StrQuestions(2) = "The dash  program is used for resolving references to external variables."
StrQuestions(3) = "An array always contains data of dash  type."
StrQuestions(4) = "A pointer variable can be assigned dash  of ordinary variable."
StrQuestions(5) = "Two pointer variables can be compared provided both points to objects of the same dash ."
StrQuestions(6) = "The function dash  indicates an end of file condition for any steam orienteddata file."
StrQuestions(7) = " An individual union member can be accessed using the operators dash ."
StrQuestions(8) = "An identifier in C must begin with a dash  or dash ."
StrQuestions(9) = "In C while is dash  word."
StrQuestions(10) = "Quotation marks within the strings (in C) must be written as dash ."

stroptions(1) = "\"
stroptions(2) = ". and ->"
stroptions(3) = "linker"
stroptions(4) = "underscore"
stroptions(5) = "address"
stroptions(6) = "2"
stroptions(7) = "eof()"
stroptions(8) = "reserved"
stroptions(9) = "data type"
stroptions(10) = "same"

stranswers(1) = "2"
stranswers(2) = "linker"
stranswers(3) = "same"
stranswers(4) = "address"
stranswers(5) = "data type"
stranswers(6) = "feof()"
stranswers(7) = ". And ->"
stranswers(8) = "underscore"
stranswers(9) = "reserved"
stranswers(10) = "\"

End Sub


Private Sub Fill_in_the_Blanks_Paper11()
StrQuestions(1) = "#define A 66" & Chr(13) & _
                  "#define B 132" & Chr(13) & _
                  "char memo[A][B]" & Chr(13) & _
                  "with respect to the above statements memo is an array of character containingdash  elements."
StrQuestions(2) = "The following declaration statement int Z[12] = {0,0,8,9,0,6} will initialise value of Z[5] to dash ."
StrQuestions(3) = "dash  can be represented as one dimensional array of character type."
StrQuestions(4) = "Following statement ++customer.last payment.mont increments month field of a structure dash ."
StrQuestions(5) = "The function sprintf is meant for formatted write to a memory dash ."
StrQuestions(6) = "#if is a preprocessor directive for dash  compilation."
StrQuestions(7) = "The directive #elif is equiv. to dash  clause to the preprocessor."
StrQuestions(8) = "Storage for storing a structure type variable is allocated during dash  time."
StrQuestions(9) = "Data files are closed in a C program using dash  library function."
StrQuestions(10) = "dash  is used for multifunction previous definition of a macro or symbolic constant."


stroptions(1) = "else-if"
stroptions(2) = "fclose"
stroptions(3) = "c"
stroptions(4) = "#undef"
stroptions(5) = "Last Payment,Customer"
stroptions(6) = "Compile"
stroptions(7) = "6"
stroptions(8) = "buffer"
stroptions(9) = "66 x 32"
stroptions(10) = "Conditional"

stranswers(1) = "66 × 32"
stranswers(2) = "6"
stranswers(3) = "String"
stranswers(4) = "Last Payment, Customer"
stranswers(5) = "buffer"
stranswers(6) = "Conditional"
stranswers(7) = "else-if"
stranswers(8) = "Compile"
stranswers(9) = "fclose"
stranswers(10) = "#undef"

End Sub

Private Sub Matching_Columns_Paper1()
strcolumnA(1) = " A Single letter is a good choice for variable same under certain circumstances"
strcolumnA(2) = "Modular Design"
strcolumnA(3) = "Some Assertions are true at the beginning of each iteration of do while statement"
strcolumnA(4) = "Keep the programs as simple as possible"
strcolumnA(5) = "A single byte capable of holding one character of the local character set"
strcolumnA(6) = "Unsigned char"
strcolumnA(7) = "'\013'"
strcolumnA(8) = "\t represents a special character in C"
strcolumnA(9) = "Character with value 0"
strcolumnA(10) = "Value of YES is 1, is correct"

strcolumnB(1) = "enum boolean {NO, YES};"
strcolumnB(2) = "Can have values between 0 and 255"
strcolumnB(3) = "tab"
strcolumnB(4) = "char C;"
strcolumnB(5) = "Arbitrary byte sized bit pattern specified in octal digits"
strcolumnB(6) = "Loop invariant"
strcolumnB(7) = "'\0'"
strcolumnB(8) = "for(i = 0; i< 20;i ++)" & Chr(13) & _
                "{j++;" & Chr(13) & _
                "}"
strcolumnB(9) = "Each function should be performing only one task"
strcolumnB(10) = "Programs must be easy to debug."

strmulanswer(1) = "for(i = 0; i< 20;i ++)" & Chr(13) & _
                "{j++;" & Chr(13) & _
                "}"
strmulanswer(2) = "Each function should be performing only one task"
strmulanswer(3) = "Loop invariant"
strmulanswer(4) = "Programs must be easy to debug."
strmulanswer(5) = "char C;"
strmulanswer(6) = "Can have values between 0 and 255"
strmulanswer(7) = "Arbitrary byte sized bit pattern specified in octal digits"
strmulanswer(8) = "tab"
strmulanswer(9) = "'\0'"
strmulanswer(10) = "enum boolean {NO, YES};"

End Sub
Private Sub Matching_Columns_Paper2()
strcolumnA(1) = "int limit = MAXLIbe + 1;"
strcolumnA(2) = "const char mg[] =" & Chr(4) & warning & Chr(4) & ""
strcolumnA(3) = "int x,y,z;" & Chr(13) & _
                "x = 2;" & Chr(13) & _
                "y = 3;" & Chr(13) & _
                "z = x/y;"
strcolumnA(4) = "char A[30];" & Chr(13) & _
                "for( i = 0;(c = getchar())! =" & Chr(13) & _
                "EOF && i <29); ++i)" & Chr(13) & _
                "A[i] = c;" & Chr(13)
strcolumnA(5) = "int k;" & Chr(13) & _
                "char c;" & Chr(13) & _
                "c = k;" & Chr(13) & _
                "k = c;"
strcolumnA(6) = "d = c > ='0' && c< = '9'"
strcolumnA(7) = "int x = 3; n=4;" & Chr(13) & _
                "x = ++n;" & Chr(13) & _
                "x = x++;"
strcolumnA(8) = "while (0)" & Chr(13) & _
                "{" & Chr(13) & _
                "i = 2;" & Chr(13) & _
                "}"
strcolumnA(9) = "do " & Chr(13) & _
                "{" & Chr(13) & _
                "i = 2;" & Chr(13) & _
                "} while(0)"
strcolumnA(10) = "for (;;)" & Chr(13) & _
                 "i = 2;" & Chr(13) & _
                 "}"

strcolumnB(1) = "5 6"
strcolumnB(2) = "assignment operation takes place in an infinite loop"
strcolumnB(3) = "assignment operation takes place once only"
strcolumnB(4) = "your program should not remove any character from input stream which you cannot process"
strcolumnB(5) = "; value o 1 is assigned if the character is digit"
strcolumnB(6) = "assignment never takes place"
strcolumnB(7) = "interchanging this order of assignment will ensure no change in value of one of the variable"
strcolumnB(8) = "Division operation produces result after truncation in some cases."
strcolumnB(9) = "elements of the array can not be altered"
strcolumnB(10) = "predefined constants can be used for initialization of variables."

stranswers(1) = "predefined constants can be used for initialization of variables."
stranswers(2) = "elements of the array can not be altered"
stranswers(3) = "Division operation produces result after truncation in some cases."
stranswers(4) = "Your program should not remove any character from input stream which you cannot process"
stranswers(5) = "interchanging this order of assignment will ensure no change in value of one of the variable"
stranswers(6) = "; value o 1 is assigned if the character is digit"
stranswers(7) = "5 6"
stranswers(8) = "assignment never takes place"
stranswers(9) = "assignment operation takes place once only"
stranswers(10) = "assignment operation takes place in an infinite loop"


End Sub

Private Sub Matching_Columns_Paper3()
strcolumnA(1) = "if(n>0) {" & Chr(13) & _
                "if(d>e)" & Chr(13) & _
                "} z = a;" & Chr(13) & _
                "Else" & Chr(13) & _
                "z = b;"
strcolumnA(2) = "Proper indentation inessential for if else statements"
strcolumnA(3) = "Break statement"
strcolumnA(4) = "Exit statement"
strcolumnA(5) = "Put a break after the last case of a switch statement"
strcolumnA(6) = "goto statement"
strcolumnA(7) = "If you ere working in a library and were given a thousand index cards to put in an alphabetical order, then a good way would be to distribute them into piles according to the first letter and sort the individual piles similarly"
strcolumnA(8) = "In a particular sorting method, the list is first partitioned by an element called pivot such that elements greater than picot is on one side and those smaller than pivot is on the other side."
strcolumnA(9) = "An operator is C permits two different expressions to appear in situations where only one expression would ordinarily be used "
strcolumnA(10) = "func 1(int n)" & Chr(13) & _
                 "{ int i;" & Chr(13) & _
                 "long int prod = 1;" & Chr(13) & _
                 "if(n>1)" & Chr(13) & _
                 "for(i=2; i<=n; ++i)" & Chr(13) & _
                 "prod*= i;" & Chr(13) & _
                 "return(prod);" & Chr(13) & _
                 "}"

strcolumnB(1) = "Forces exit from a program"
strcolumnB(2) = "Divide and conquer"
strcolumnB(3) = "exit from some deeply bested structure (eg. two or more loops) at once"
strcolumnB(4) = "returned value does not match the function type"
strcolumnB(5) = "Indentations help in understanding if else statements"
strcolumnB(6) = "The comma operator(,)"
strcolumnB(7) = "quick sort"
strcolumnB(8) = "Forces immediate exit from any loop"
strcolumnB(9) = "unnecessary but good programming practice."
strcolumnB(10) = "Braces can alter the order of association between if and else statements."

stranswers(1) = "Braces can alter the order of association between if and else statements."
stranswers(2) = "Indentations help in understanding if else statements"
stranswers(3) = "Forces immediate exit from any loop"
stranswers(4) = "Forces exit from a program"
stranswers(5) = "unnecessary but good programming practice."
stranswers(6) = "exit from some deeply bested structure (eg. two or more loops) at once"
stranswers(7) = "Divide and conquer"
stranswers(8) = "quick sort"
stranswers(9) = "The comma operator(,)"
stranswers(10) = "returned value does not match the function type"

End Sub

Private Sub Matching_Columns_Paper4()
strcolumnA(1) = "Variables internal to a function; comes into existence when the function is called"
strcolumnA(2) = "Variables used by more than one function and are of permanent nature"
strcolumnA(3) = "Value of YES is 1, is incorrect"
strcolumnA(4) = "f(inta)"
strcolumnA(5) = "f(int *)"
strcolumnA(6) = "#include" & Chr(4) & "filename" & Chr(4) & ""
strcolumnA(7) = "#include <filename>"
strcolumnA(8) = "void swap(int x, int y)" & Chr(13) & _
                "{ int temp;" & Chr(13) & _
                "temp = x;" & Chr(13) & _
                "x = y;" & Chr(13) & _
                "y = temp" & Chr(13) & _
                "}"
strcolumnA(9) = "no space allocated for storage of character during compile time"
strcolumnA(10) = "#if SYSTEM == SYSV" & Chr(13) & _
                 "#define HDR " & Chr(4) & " sysv.h" & Chr(4)

strcolumnB(1) = "f(a[2])"
strcolumnB(2) = "yes = atoi(49)"
strcolumnB(3) = "conditional expressions must be constant integer expressions in this statement."
strcolumnB(4) = "Semantic error"
strcolumnB(5) = "file is searched in the directory where source is located"
strcolumnB(6) = "char char *s;"
strcolumnB(7) = "automatic variables"
strcolumnB(8) = "file is searched according to implementation"
strcolumnB(9) = "f(&a[2])"
strcolumnB(10) = "external variables"

stranswers(1) = "automatic variables"
stranswers(2) = "external variables"
stranswers(3) = "yes = atoi(49)"
stranswers(4) = "f(a[2])"
stranswers(5) = "f(&a[2])"
stranswers(6) = "file is searched in the directory where source is located"
stranswers(7) = "file is searched according to implementation"
stranswers(8) = "Semantic error"
stranswers(9) = "char char *s;"
stranswers(10) = "conditional expressions must be constant integer expressions in this statement."


End Sub
Private Sub Matching_Columns_Paper5()
strcolumnA(1) = "While (*p!='\0')" & Chr(13) & _
                "p++;"
strcolumnA(2) = "While (*p!=NULL)" & Chr(13) & _
                "p++;"
strcolumnA(3) = "int (*p)[10]"
strcolumnA(4) = "int *p[10]"
strcolumnA(5) = "int (*p)(void8,void*)"
strcolumnA(6) = "int *p(void*,void*)"
strcolumnA(7) = "int (*p[5])()"
strcolumnA(8) = "int (*(*p())"
strcolumnA(9) = "sizeof"
strcolumnA(10) = "#define"

strcolumnB(1) = "p is a pointer to a function that returns integer"
strcolumnB(2) = "compile time directive"
strcolumnB(3) = "p is function that returns pointer to integer"
strcolumnB(4) = "p is a pointer to a function that returns a pointer to integer"
strcolumnB(5) = "p is an array to pointer to functions which return integer"
strcolumnB(6) = "p is a pointer to a character string"
strcolumnB(7) = "p is a pointer to an array of integers"
strcolumnB(8) = "compile time operator"
strcolumnB(9) = "is pointer to an array of integer pointers"
strcolumnB(10) = "p is used in improper fashion"

stranswers(1) = "p is a pointer to a character string"
stranswers(2) = "p is used in improper fashion"
stranswers(3) = "p is a pointer to an array of integers"
stranswers(4) = "is pointer to an array of integer pointers"
stranswers(5) = "p is a pointer to a function that returns integer"
stranswers(6) = "p is function that returns pointer to integer"
stranswers(7) = "p is an array to pointer to functions which return integer"
stranswers(8) = "p is a pointer to a function that returns a pointer to integer"
stranswers(9) = "compile time operator"
stranswers(10) = "compile time directive"

End Sub
Private Sub Matching_Columns_Paper6()
strcolumnA(1) = "struct element * func (char *work,int n)"
strcolumnA(2) = "self -refrential structure ."
strcolumnA(3) = "A variable that can hold objects of different types"
strcolumnA(4) = "A function that returns objects of different types"
strcolumnA(5) = "You can define variables in C which occupies less space then character variables"
strcolumnA(6) = "getchar() has returned -1"
strcolumnA(7) = "returns un initialized storage dynamically ( in run time)"
strcolumnA(8) = "returns initialized storage in run time"
strcolumnA(9) = "isupper('a')"
strcolumnA(10) = "toupper('A')"

strcolumnB(1) = "calloc"
strcolumnB(2) = "A"
strcolumnB(3) = "bit fields"
strcolumnB(4) = "malloc"
strcolumnB(5) = "O"
strcolumnB(6) = "alternate form of declaration for better program readability"
strcolumnB(7) = "must be declared as union"
strcolumnB(8) = "end of file has come(most probably)"
strcolumnB(9) = "must be declared as void"
strcolumnB(10) = "useful for link list implementation"


stranswers(1) = "alternate form of declaration for better program readability"
stranswers(2) = "useful for link list implementation"
stranswers(3) = "must be declared as union"
stranswers(4) = "must be declared as void"
stranswers(5) = "bit fields"
stranswers(6) = "end of file has come(most probably)"
stranswers(7) = "malloc"
stranswers(8) = "calloc"
stranswers(9) = "O"
stranswers(10) = "A"


End Sub

Private Sub Matching_Columns_Paper7()
strcolumnA(1) = "lseek"
strcolumnA(2) = "fseek"
strcolumnA(3) = "r+"
strcolumnA(4) = "w+"
strcolumnA(5) = "int y[4][3] = {" & Chr(13) & _
                "1,3,5,2,4,6,3,5,7" & Chr(13) & _
                "};"
strcolumnA(6) = "int y [4][3]={" & Chr(13) & _
                "{1},{2},{3},{4}" & Chr(13) & _
                "}"
strcolumnA(7) = "In a sorting program, for the sake of generality, a programmer decided to use mecros for comparing keys."
strcolumnA(8) = "In a sorting program, for the sake of generality, a programmer for implementing the code for  comparing of keys;"
strcolumnA(9) = "int a[2]" & Chr(13) & _
                "float x,y;" & Chr(13) & _
                "char b[4]"
strcolumnA(10) = "In fopen " & Chr(4) & "w" & Chr(4) & " mode is used"

strcolumnB(1) = "Run time overhead is significant"
strcolumnB(2) = "y[3][0] is 4"
strcolumnB(3) = "y[3][1] is 0"
strcolumnB(4) = "Occupy 20 bytes of memory"
strcolumnB(5) = "valid mode for opening a file; permits read and write"
strcolumnB(6) = "To create text file for writing"
strcolumnB(7) = "valid mode for opening a file; permits read and write; previous content is discarded."
strcolumnB(8) = "random access in the file; file specified through file descriptor"
strcolumnB(9) = "run time overhead is minimal and hence a desirable strategy"
strcolumnB(10) = "random access in the file; file specified through file pointer."

stranswers(1) = "random access in the file; file specified through file pointer."
stranswers(2) = "random access in the file; file specified through file descriptor"
stranswers(3) = "y[3][0] is 4"
stranswers(4) = "valid mode for opening a file; permits read and write; previous content is discarded."
stranswers(5) = "y[3][1] is 0"
stranswers(6) = "y[3][0] is 4"
stranswers(7) = "Run time overhead is significant"
stranswers(8) = "run time overhead is minimal and hence a desirable strategy"
stranswers(9) = "Occupy 20 bytes of memory"
stranswers(10) = "To create text file for writing"


End Sub

Private Sub Matching_Columns_Paper8()
strcolumnA(1) = "getchar()"
strcolumnA(2) = "get(strl);"
strcolumnA(3) = "getc()"
strcolumnA(4) = "scanf(" & Chr(4) & "%s" & Chr(13) & ",strl);"
strcolumnA(5) = "In fopen " & Chr(4) & "a+" & Chr(4) & " mode is used"
strcolumnA(6) = "allows to define a global variable in one file and declare and use it in another file"
strcolumnA(7) = "Carriage return"
strcolumnA(8) = "x+=(2<<3)"
strcolumnA(9) = "local"
strcolumnA(10) = "x=(float)5/2"

strcolumnB(1) = "x=(float)(5/2)"
strcolumnB(2) = "\r"
strcolumnB(3) = "To Open text file for read/write"
strcolumnB(4) = "x = 5/(float) 2"
strcolumnB(5) = "Returns text character from input stream"
strcolumnB(6) = "x = x+2*2*2"
strcolumnB(7) = "if input string is " & Chr(4) & "Sardar Patel" & Chr(4) & " value stored in strl is " & Chr(4) & "Sardar Patel" & Chr(4)
strcolumnB(8) = "If input string is " & Chr(4) & "Sarder Patel" & Chr(4) & " value stored in strl is " & Chr(4) & "Sarder" & Chr(4)
strcolumnB(9) = "Returns text character from stdin"
strcolumnB(10) = "extern"

stranswers(1) = "Returns text character from stdin"
stranswers(2) = "if input string is " & Chr(4) & "Sardar Patel" & Chr(4) & " value stored in strl is " & Chr(4) & "Sardar Patel" & Chr(4)
stranswers(3) = "Returns text character from input stream"
stranswers(4) = "If input string is " & Chr(4) & "Sarder Patel" & Chr(4) & " value stored in strl is " & Chr(4) & "Sarder" & Chr(4)
stranswers(5) = "To Open text file for read/write"
stranswers(6) = "extern"
stranswers(7) = "\r"
stranswers(8) = "x = x+2*2*2"
stranswers(9) = "x=(float)(5/2)"
stranswers(10) = "x = 5/(float) 2"
End Sub

Private Sub Matching_Columns_Paper9()
strcolumnA(1) = "union"
strcolumnA(2) = "calloc()"
strcolumnA(3) = "malloc()"
strcolumnA(4) = "puts()"
strcolumnA(5) = "fputs()"
strcolumnA(6) = "#include"
strcolumnA(7) = "rvalue of data item remains unchanged during program execution"
strcolumnA(8) = "accomplishing indirection"
strcolumnA(9) = "a small buffer in memory that is capable of holding different types ofdata"
strcolumnA(10) = "address space of data item in memory"

strcolumnB(1) = "allocate and clear memory"
strcolumnB(2) = "union"
strcolumnB(3) = "preprocessor command."
strcolumnB(4) = "arrow operator"
strcolumnB(5) = "const type qualifier"
strcolumnB(6) = "lvalue"
strcolumnB(7) = "output character to screen"
strcolumnB(8) = "is a memory location that is used by several different variables, which may be of different type."
strcolumnB(9) = "allocates memory bit does not clear memory."
strcolumnB(10) = "does not append a new live character"

stranswers(1) = "is a memory location that is used by several different variables, which may be of different type."
stranswers(2) = "allocate and clear memory"
stranswers(3) = "allocates memory bit does not clear memory."
stranswers(4) = "output character to screen"
stranswers(5) = "does not append a new live character"
stranswers(6) = "preprocessor command."
stranswers(7) = "const type qualifier"
stranswers(8) = "arrow operator"
stranswers(9) = "union"
stranswers(10) = "lvalue"

End Sub
Private Sub Matching_Columns_Paper10()
strcolumnA(1) = "to convert the operand to the specified data type "
strcolumnA(2) = "visibility of data item "
strcolumnA(3) = "making function private to the file in which it is defined"
strcolumnA(4) = "allows to define a global variable in one file and declare and use it in another file"
strcolumnA(5) = "&"
strcolumnA(6) = "&="
strcolumnA(7) = "&&"
strcolumnA(8) = "automatic variable"
strcolumnA(9) = "Register variable"
strcolumnA(10) = "Static variable"

strcolumnB(1) = "increase in speed of execution"
strcolumnB(2) = "assignment operator"
strcolumnB(3) = "scope till the end of file"
strcolumnB(4) = "logical operator"
strcolumnB(5) = "static"
strcolumnB(6) = "cast operator"
strcolumnB(7) = "bitwise operator"
strcolumnB(8) = "scope till the end of the function in which declared"
strcolumnB(9) = "extern"
strcolumnB(10) = "scope"

stranswers(1) = "cast operator"
stranswers(2) = "scope"
stranswers(3) = "static"
stranswers(4) = "extern"
stranswers(5) = "bitwise operator"
stranswers(6) = "assignment operator"
stranswers(7) = "logical operator"
stranswers(8) = "scope till the end of the function in which declared"
stranswers(9) = "increase in speed of execution"
stranswers(10) = "scope till the end of file"

End Sub
Private Sub Matching_Columns_Paper11()
strcolumnA(1) = "Macro"
strcolumnA(2) = "Function"
strcolumnA(3) = "Switch statement"
strcolumnA(4) = "for(;;)"
strcolumnA(5) = "do-while statement"
strcolumnA(6) = "goto statement"
strcolumnA(7) = "Formal arguments of function"
strcolumnA(8) = "Actual argument of a function"
strcolumnA(9) = "call by reference"
strcolumnA(10) = "Call by value"

strcolumnB(1) = "function invocation"
strcolumnB(2) = "difficult to follow the flow of control in the program"
strcolumnB(3) = "unbending loop if no break statement inside the body"
strcolumnB(4) = "values of parameters in the calling program is affected"
strcolumnB(5) = "group of a few C statements with infrequent usage"
strcolumnB(6) = "Function definition"
strcolumnB(7) = "Frequently used group of some C statements."
strcolumnB(8) = "Multi level decision"
strcolumnB(9) = "Value of parameters in the calling program is unaffected"
strcolumnB(10) = "Body executed at least once"

stranswers(1) = "Frequently used group of some C statements."
stranswers(2) = "group of a few C statements with infrequent usage"
stranswers(3) = "Multi level decision"
stranswers(4) = "unbending loop if no break statement inside the body"
stranswers(5) = "Body executed at least once"
stranswers(6) = "difficult to follow the flow of control in the program"
stranswers(7) = "Value of parameters in the calling program is unaffected"
stranswers(8) = "function invocation"
stranswers(9) = "values of parameters in the calling program is affected"
stranswers(10) = "Body executed at least once"

End Sub

Private Sub TrueorFalse_Paper1()
StrQuestions(1) = "The printf () is a macro which takes in a variable number of arguments."
StrQuestions(2) = "A char type has an equivalent integer interpretation; so it is really a special kibe of short integer."
StrQuestions(3) = "12,245 is a legal string constant in 'C'."
StrQuestions(4) = "One cannot have octal or hexadecimal integer constants in 'C'."
StrQuestions(5) = "All functions in 'C' have to return an integer or a character value."
StrQuestions(6) = "When a variable is passed to function, by value , its value remains unchanged in the calling program."
StrQuestions(7) = "All macro substitutions in programmer one before compilation of the program."
StrQuestions(8) = "All files need not be opened with fopen(...) before it is used forreading or writing within he program."
StrQuestions(9) = "The functions getchar() and putchar() do the same thin so they are interchangeable in a program."
StrQuestions(10) = "An external variable is one which is essentially defined in a file other than the one in which it is used."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper2()
StrQuestions(1) = "Static variables are nothing but global constants. "
StrQuestions(2) = "It is possible to write a function which can return either an integer or a floating pointe value , but not a string."
StrQuestions(3) = "Since the library function sqrt()takes in a double precision variable as argrments , it is not possible to find the square root of an integer value."
StrQuestions(4) = "In program, the statement   #include 'filename' is replaced by the contents of the file 'filename', before compilation."
StrQuestions(5) = "The statement    #define Yes 1   would replace all the occurrences of YES without quotes in the program, by 1."
StrQuestions(6) = "The function sprintf() functions like fprintf(), but operates on a string, instead of afile."
StrQuestions(7) = "Typedef is a facility in C for specifying few data types."
StrQuestions(8) = "The statement    typedef int LENGTH  declares LENGTH as a variable of type int."
StrQuestions(9) = "A register declaration advises the compiler that the corresponding variable will be heavily in use."
StrQuestions(10) = " Recursive calls to a function are executed in a Last In First Out manner."

stranswers(1) = "False"
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

Private Sub TrueorFalse_Paper3()
StrQuestions(1) = "All automatic variables become static when a function is recursive."
StrQuestions(2) = "When a variable i is incriminated with the operation i++ in a statement, it mans that i has toe be incremented before the statement is executed "
StrQuestions(3) = "It is not possible to have bested if else statements in 'C'.)"
StrQuestions(4) = "If an integer variable is assigned a decimal value, it is not possible to print its hexadecimal equivalent without writing a program for the necessary conversion."
StrQuestions(5) = "If x is an integer variable and px is a pointer, the statement   px = *x assigns the address of x to px."
StrQuestions(6) = "Automatic arras cannot be initialized"
StrQuestions(7) = "Arrays cannot be returned by functions , however pointes to arrays can be returned."
StrQuestions(8) = "It is not possible to allocate memory dynamically in 'C'."
StrQuestions(9) = "The arguments to scanf() should be pointers to variables and not the variables names."
StrQuestions(10) = " Bitwise logical operators are applicable to float , double or integer variables only."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper4()
StrQuestions(1) = "The function getchar() is a library function which reads in a single character from standard input."
StrQuestions(2) = "The select and for loops test the termination condition at the bottom of the loop, while they do while tests the termination condition at the top."
StrQuestions(3) = "The break statement is used when it is required to exit from a loop other by the testing of termination condition."
StrQuestions(4) = "The continue statement  is used to skip some statements within a loop and start text iteration."
StrQuestions(5) = "The operator '==' is an assignment cum increment operator in 'C'."
StrQuestions(6) = "Members of enumerated data types are constants that are written as identifiers, but have signed integer values."
StrQuestions(7) = "Unsigned integers can have values ranging from a garbage minimum to a positive maximum."
StrQuestions(8) = "It is not possible to have formatted input \ output in 'C'."
StrQuestions(9) = "A linked list is a data structure which is created by dynamic allocation of memory."
StrQuestions(10) = " The maximum number of elements to be ever stored in a linked list has to be declared a priori like the dimension of an array."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper5()
StrQuestions(1) = "An array declared as A[100][100] can hold a maximum of 1000 elements."
StrQuestions(2) = "A structure is a composite data type which contains more than one mumber and all members are essentially of the same type."
StrQuestions(3) = "Malloc is a function which is used to allocate memory and returns pointers to variables."
StrQuestions(4) = "If A[10[] is a array,(A + ) points to he second element of the array."
StrQuestions(5) = "When a function accepts another functions name as argument, the formal argument declaration is 'datatype(*finction-name)();'."
StrQuestions(6) = "A bitwise exclusive -OR operation will return a 1, if any of the bits has a value 1 and the other has value 0 or 1."
StrQuestions(7) = "The decimal equivalent of 0x32 is 50."
StrQuestions(8) = "The ! operator is a ternary operator and has the highest precedence among all logical operators."
StrQuestions(9) = "The declaration  void unction - name()   indicates that function name returns nothing to the calling program."
StrQuestions(10) = " It's not possible to have switch statements bested within while or for loops."

stranswers(1) = "False"
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

Private Sub TrueorFalse_Paper6()
StrQuestions(1) = "int p(char*a[]) is a function that accepts an argument which is a pointerto a character array ans returns a n integer."
StrQuestions(2) = "int(*p[10] is a pointer to 10 element integer array."
StrQuestions(3) = "The expression(*t[2][3]+5) can be used for accessing as individual array element t[2][3][5]."
StrQuestions(4) = " int (*a)[100];  *a = (int *) malloc (10 * 5 * sizeof(int)       );  This program segment will assign to the variable , a pointer of the correct type.5. float *pv = NULL  is an example of correct initialization of pointer variable."
StrQuestions(5) = "struct{ int acct-no; char acct_type; float balance;  } customer, *pc;    pc = &customer  This program segment will initialize pc to null pointer"
StrQuestions(6) = "struct{ int red;     int blue;       int green;  }color, *cp;    cp = & color;   If these statements are there in a program then compiler would generate error for the statement (*cp).green = 2;    occurring in the same program."
StrQuestions(7) = "A structure can include obe or more pointers as members"
StrQuestions(8) = "struct( int red;     int blue;       int green;  {color, *cp = & color;  is a syntactically correct declaration."
StrQuestions(9) = " The -> operator ca be combined with the period operator to access a member of a structure that is itself a member of another structure."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "True"
End Sub

Private Sub TrueorFalse_Paper7()
StrQuestions(1) = "The operator -> is not left to right associative."
StrQuestions(2) = "We can pass only pointer to a completer structure and not a complete structure as function argument."
StrQuestions(3) = "struct pt{   int x;  int y;}  struct rect{   struct point R;         struct point L;     }; struct rect r, *rp = &r; Given the above program segment, r.R.x and (rp R).x are identical."
StrQuestions(4) = "struct{ int l;       char *st;   } *p;   Given the above declaration ++p -> len and (p++) -> len expressions increment the same variable."
StrQuestions(5) = "*p++ increments the content of the location pointed by p."
StrQuestions(6) = "(*p++ ->st) and (*p-> st++) are equivalent expressions."
StrQuestions(7) = "A structure can be passed to a function in C through Call by Value mechanism."
StrQuestions(8) = "It is illegal for a structure to contain itself as a member."
StrQuestions(9) = "A self referential structure includes a member which is a pointer to the structure type itself."
StrQuestions(10) = " Members that compose a union share the same storage area in computer's memory."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper8()
StrQuestions(1) = "Members that compose a union are of different data type and occupy different amount of space."
StrQuestions(2) = "A union cannot be a member of a union."
StrQuestions(3) = "fact = 1;    while(i<7   { i=i + 1;  fact = i * fact }   Invariant for the while loop is         'fact is the i! .AND. i n'"
StrQuestions(4) = "i = 1;   p = 1;  while(i<n)      { i = i + 1;          if a[i] < [p] then p = i;     }   Invariant for the while loop is  '(1 <= i<= N) .AND. (a[p] < a[1])'"
StrQuestions(5) = "When all elements to be sorted are equal, quick sort does a lot of unbecessary exchanges."
StrQuestions(6) = "Tower of Hanoi problem cannot be solved using only two pegs."
StrQuestions(7) = "dx = 2; x = 5; n = 100;  while(x<n)  { x = x + dx;   dx = abx(dx-6); printf(%dn,x);  }   The above program segment prints numbers which are not multiples of 2 and 3."
StrQuestions(8) = "Ones complement operator ~ is  a binary operator that causes the bits of its operands to be inverted."
StrQuestions(9) = " extern double val[];    is a valid declaration for external variable."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper9()
StrQuestions(1) = "Recursive functions provide saving in storage space."
StrQuestions(2) = "It is given that # define max(A,B) ((A)>(B) ? (A):(B);   On execution the statement  max(i++,j++)    will cause increment o larger value by 2."
StrQuestions(3) = "The program segment  int n;  scanf (%d,n);   will produce compiler error."
StrQuestions(4) = "The conversion specification %h is used to read the values in hexadecimal."
StrQuestions(5) = "Typedef declarations adds new name for some existing type."
StrQuestions(6) = "Vital revenue is valid variable name in C."
StrQuestions(7) = "During compilation , C program ignores leading zeros so that 123 is same as 0123."
StrQuestions(8) = "The size of the array defined as     char *a[25] will be 25 *sizeof(char)."
StrQuestions(9) = "While popping an element from the stack, the stack size will decease."
StrQuestions(10) = " The time taken to access the Nth element of a linked list will be more than the time taken to access the Nth element from a sequential list."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper10()
StrQuestions(1) = "To use binary searching, the table in which the search is carried out, beed not be in sorted order."
StrQuestions(2) = "In the case of call by reference the parameter will be always of pointer type."
StrQuestions(3) = "Wile inserting an element in the middle of a linked list, it is not becesary to shift the elements that occur after the inserted element."
StrQuestions(4) = "It is not possible to find the address of the first node given the address ofthe last node in a singly linked list."
StrQuestions(5) = "In C programming language, strings are represents using an array."
StrQuestions(6) = "Statements inside a do-while loop will be executed at least once."
StrQuestions(7) = "The && and 11 operators compares tow numeric values."
StrQuestions(8) = "Using union, a block of memory can be accessed in different ways."
StrQuestions(9) = "It is possible to transfer the control from a function to another function using goto statement."
StrQuestions(10) = " Elements of a structure have to be of the same type."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper11()
StrQuestions(1) = "A variable of type char can old the value 301"
StrQuestions(2) = "Quick sort will be more efficient, if the table to be sorted is initially almost sorted."
StrQuestions(3) = "Relational operators have higher precedence than arithmetic operators"
StrQuestions(4) = "You can assign one structure variable to another, provided they are ofthe same type."
StrQuestions(5) = "When a function returns a value, the entire function call can appear on theright side o the equality sign."
StrQuestions(6) = "The compiler will complain if you try to access array element 14 in the array A defined as intA[10]"
StrQuestions(7) = "Evaluation of the expression (exp1 && exp2) always causes evaluation of exp2."
StrQuestions(8) = "Consider the declarations    int x;  char ch= 'A';   After execution of x = ch - 65; x will have value zero."
StrQuestions(9) = "A register variable cannot be static."
StrQuestions(10) = " The maximum number of register class variables that can be declared in a function must be less that 0r equal to the number of machine registers that are available."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper12()
StrQuestions(1) = "The function main() cannot be recursive."
StrQuestions(2) = "The statement   for(;;);    will make the program non terminating."
StrQuestions(3) = "The external static variables cannot be accessed from outside the source file where it is declared."
StrQuestions(4) = "Every C program need not have main() function."
StrQuestions(5) = "'Unresolved symbol' error message from linker may indicate that the library specified is not correct."
StrQuestions(6) = "Constants do not have associated lvalue."
StrQuestions(7) = "An 'if-else' statement can be replaced by the operator?"
StrQuestions(8) = "Function definition and function declaration are one and the same thing."
StrQuestions(9) = "File descriptor is an integer value that serves as index number to describe the file."
StrQuestions(10) = " In an assignment statement the value on the left hand side of the equality sign is converted to the data type of the variable on the right hand side of the equality sign."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper13()
StrQuestions(1) = "The statement    x = x <>2;  is equivalent to the statement  x = x*4:"
StrQuestions(2) = "Consider the following C statement   while((c = getchar()) ! = EOF)  whee 'c' is of type int. There is an error in this statement."
StrQuestions(3) = "A variable declared a 'register variable' is registered with the value being initialized and it cannot be altered."
StrQuestions(4) = "Constant expression are evaluated at the run time and not at the compile time."
StrQuestions(5) = "In deciding on an algorithm for solving a problem it is important to consider onlu its time complexity."
StrQuestions(6) = "The function fscanf can take variable number of arguments."
StrQuestions(7) = "Macros and functions are different names for the same feature of C."
StrQuestions(8) = "#define can be used for defining symbolic constants."
StrQuestions(9) = "# if n defined is a valid preprocessor directive."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "True"
End Sub

Private Sub TrueorFalse_Paper14()
StrQuestions(1) = "The function fprintf is meant for unformatted write to data files."
StrQuestions(2) = "Recursive function for competition of the nth Fibonacci number is mere efficient than its alive formulation of the same function."
StrQuestions(3) = "Insertion sort cannot be implemented using arrays."
StrQuestions(4) = "Library function strcmp returns an integer value."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
End Sub

Private Sub multiplechoice_Paper1()
strmultiple(1) = "Arrays can be initialized provided they are"
strmuloption(1) = "automatic"
strmuloption(2) = "external"
strmuloption(3) = "static"
strmuloption(4) = "both b and c"

strmultiple(2) = "Arrays are passed as arguments to a function by"
strmuloption(5) = "value"
strmuloption(6) = "reference"
strmuloption(7) = "both a and b"
strmuloption(8) = "none of the above"

strmultiple(3) = "Which of the following statements is true after execution of the program   int a[10],i,*p; a[0] = 1;   a[1] = 2;   p=a;    (*p)++;"
strmuloption(9) = "a[0]"
strmuloption(10) = "a[1] = 3"
strmuloption(11) = "a[1] = 2 "
strmuloption(12) = "all"

strmultiple(4) = "It is necessary to declare the type of function in the calling program if"
strmuloption(13) = "the function returns an integer"
strmuloption(14) = "the function returns a non integer value"
strmuloption(15) = "the function is not defined in the same file"
strmuloption(16) = "none of the above"

strmultiple(5) = "The declaration void function 1(int) indicates the function1 is a function which"
strmuloption(17) = "has no arguments"
strmuloption(18) = "returns nothing"
strmuloption(19) = "both a and b"
strmuloption(20) = "none of the above"

strmultiple(6) = "Recursive functions are executed in a "
strmuloption(21) = "Last in first out order"
strmuloption(22) = "First in first out order"
strmuloption(23) = "parallel fashion"
strmuloption(24) = "anyone of the above"

strmultiple(7) = "When a function in recursively calls , all automatic variables"
strmuloption(25) = "are initialized during each execution of the function"
strmuloption(26) = "are retained from the last execution"
strmuloption(27) = "are maintained in a stack"
strmuloption(28) = "none of the above"

strmultiple(8) = "A static variable is one"
strmuloption(29) = "which cannot be initialized "
strmuloption(30) = "which is initialized once at the commencement of execution and cannot bechanged at run time."
strmuloption(31) = "retains its value throughout the life of the program."
strmuloption(32) = "which is same as an automatic variable nut is placed at the head of a program."

strmultiple(9) = "An external variable is one"
strmuloption(33) = "which is global accessible by all functions"
strmuloption(34) = "has declaration 'extern' associated with it when declared within a function"
strmuloption(35) = "will be initialized to 0 if not initialized"
strmuloption(36) = "all of the above"

strmultiple(10) = "Consider the following program fragment main(){int a,b,c;b=2;a=2*(b++);c=2*(++b);}which one of the given answers is correct?"
strmuloption(37) = "a = 4, c = 6"
strmuloption(38) = "a = 3, c = 8"
strmuloption(39) = "b = 3, c = 6"
strmuloption(40) = "a = 4, c = 8"

strmulanswer(1) = "both b and c"
strmulanswer(2) = "reference"
strmulanswer(3) = "a[1] = 3"
strmulanswer(4) = "the function returns a non integer value"
strmulanswer(5) = "returns nothing"
strmulanswer(6) = "Last in first out order"
strmulanswer(7) = "are initialized during each execution of the function"
strmulanswer(8) = "retains its value throughout the life of the program."
strmulanswer(9) = "all of the above"
strmulanswer(10) = "a = 4, c = 8"

End Sub

Private Sub multiplechoice_Paper2()
strmultiple(1) = "Given the following program fragment, which one of the alternatives is correct? main(){char status;int balance;balance = 1000;status = (balance >= 1000) ? 'C': 'O' ;}"
strmuloption(1) = "status = 'O'"
strmuloption(2) = "status = 'C'"
strmuloption(3) = "status = O;"
strmuloption(4) = "status = NIL"

strmultiple(2) = "What will be value of count after the following program is executed ? main(){int count, digit = 0;count = 1;while(digit <= 9) {printf(%d\n,++count);++digit;}"
strmuloption(5) = "10"
strmuloption(6) = "9"
strmuloption(7) = "12"
strmuloption(8) = "11"

strmultiple(3) = "What will be the value of sum after the following program is executed? main(){int sum, index,sum = 1;index = 9;do{  index = index - 1;  sum = 2 * sum;  }while(index>9);"
strmuloption(9) = "1"
strmuloption(10) = "2"
strmuloption(11) = "9"
strmuloption(12) = "0.5"

strmultiple(4) = "How many times will the print statement be executed? main(){int n;n = 10;while (n<10){ printf('hello'); --n;)}"
strmuloption(13) = "never"
strmuloption(14) = "once"
strmuloption(15) = "10"
strmuloption(16) = "9"

strmultiple(5) = "If the following loop is implemented  {int num;   num = 0;    do {--num;   printf('%d',num);   num++;  } while (num >=0)  }"
strmuloption(17) = "the loop will run infinitely many times"
strmuloption(18) = "the program will not enter the loop"
strmuloption(19) = "there will be a compilation error reported"
strmuloption(20) = "a run time error will be reported."

strmultiple(6) = "What is the final value of digit? main(){int digit;for (digit = 0; digit <= 9; ++digit)     printf('%d\n',digit);digit = 2 *digit;--digir;}"
strmuloption(21) = "19"
strmuloption(22) = "-1"
strmuloption(23) = "17"
strmuloption(24) = "16"

strmultiple(7) = "What is the final value of sum? main(){int sum = 1;for (;sum = 1;   printf('%d\n',++sum);}"
strmuloption(25) = "10"
strmuloption(26) = "9"
strmuloption(27) = "11"
strmuloption(28) = "none of the above"

strmultiple(8) = "What is the value of 'average ' after the following program is executed?main(){int sum, index; index = o;  sum = 0;for(;;) {   sum = sum+index;    ++index;    if(sum >= 100) break;   }   average = sum/index;    }"
strmuloption(29) = "91/14"
strmuloption(30) = "91/13"
strmuloption(31) = "105/15"
strmuloption(32) = "105/14"

strmultiple(9) = "What is the output of the following program? main(){int x,y,a;x = 2;y = 1;z = 1;if (x > y + z)      printf('Hello!'\n');else if (x < y + z) printf('Hi!\n');    else    printf('Hey!\n');}"
strmuloption(33) = "Hi!"
strmuloption(34) = "Hello!"
strmuloption(35) = "Hey!"
strmuloption(36) = "none "

strmultiple(10) = " What is the value of variable POLYGON? main(){int POLYGON,L,B;   L=B=2;POLYGON = (L ==B) ? 1:0;}"
strmuloption(37) = "0"
strmuloption(38) = "1"
strmuloption(39) = "2"
strmuloption(40) = "None of the above"


strmulanswer(1) = "status = 'C'"
strmulanswer(2) = "11"
strmulanswer(3) = "2"
strmulanswer(4) = "never"
strmulanswer(5) = "the loop will run infinitely many times"
strmulanswer(6) = "19"
strmulanswer(7) = "10"
strmulanswer(8) = "105/15"
strmulanswer(9) = "Hey!"
strmulanswer(10) = "1"

End Sub

Private Sub multiplechoice_Paper3()
strmultiple(1) = "#define  max(x,y) x = (x>y)?x:y is a macro definition. Which an find the maximum of two numbers x and y if"
strmuloption(1) = "x and y are both integers"
strmuloption(2) = "x and y are both declared as float"
strmuloption(3) = "x and y are both declared as double"
strmuloption(4) = "all the above are true"

strmultiple(2) = "Consider the macro definitions #define square(x) x*xand #define square(x) (x*x)when used in a program statement B = ++square(C)"
strmuloption(5) = "they will produce the same result"
strmuloption(6) = "they will not produce same results"
strmuloption(7) = "the second definition is not legal"
strmuloption(8) = "none of the above"

strmultiple(3) = "#define two(x) 2*x   #define double(x) x+x   main()   {    int num,sum,product;  num=1;  sum= --two(num); --sum; product= --double(num);    printf('%d%d\n',sum,product);   }   The output oftyhe above program is"
strmuloption(9) = "0 0"
strmuloption(10) = "0 1"
strmuloption(11) = "1 1"
strmuloption(12) = "1 0"

strmultiple(4) = "The following program fragment{int sum, index;index = 50;while(index >= 0) {sum = sum/index;-index;   }}"
strmuloption(13) = "will give a run time error"
strmuloption(14) = "will give a compilation error"
strmuloption(15) = "will give a linking error"
strmuloption(16) = "none of the above"

strmultiple(5) = "What is the output of the following program? main(){int B,X,Y,Z;X = 1;Y = 2;Z = 3;if (X > 1) | (y > 1) if (Z >1)   printf ('OK.\n');   else break;if ((X > 1) && (Z > 3))  printf('Bye\n');printf('Hello!');}"
strmuloption(17) = "OK. Bye"
strmuloption(18) = "Bye"
strmuloption(19) = "OK"
strmuloption(20) = "Hello!"

strmultiple(6) = "main(){float balance, loan;balance = 1000.0;loan= balance/10;if((balance > 500)(loan < 500))  printf('good account\n');   if(balance < 500)(loan < 500)   printf({'caution !\n');}What is the output of the above program?"
strmuloption(21) = "good account"
strmuloption(22) = "caution"
strmuloption(23) = "good account caution"
strmuloption(24) = "good account caution"

strmultiple(7) = "The following libes, if included in a program, will cause obe of the following errors. Indicate the correct obe.{double c; scanf('%c',c);}"
strmuloption(25) = "Runtime error"
strmuloption(26) = "Compilation error"
strmuloption(27) = "Typedef error"
strmuloption(28) = "No error"

strmultiple(8) = "A 'switch' statement is used to"
strmuloption(29) = "switch between functions in a program"
strmuloption(30) = "switch from one variable to another variable"
strmuloption(31) = "to choose from multiple possibilities which may arise due to different values of a single variable"
strmuloption(32) = "to use switching variables"

strmultiple(9) = "If the following variables are set to the values as shown below , then what is the value of the expression following it? answer = 2;marks = 10;!((answer < 5)(marks > 2))"
strmuloption(33) = "1"
strmuloption(34) = "0"
strmuloption(35) = "-1"
strmuloption(36) = "2"

strmultiple(10) = " What is the following program doing? main(){int digit = 0;doprintf('%d\n',digit++)while(digit , = 9);"
strmuloption(37) = "Adding 9 integers"
strmuloption(38) = "Adding integers from 1 to 9"
strmuloption(39) = "Displaying any 9 integers"
strmuloption(40) = "Displaying integers from 1 to 9"

strmulanswer(1) = "all the above are true"
strmulanswer(2) = "they will not produce same results"
strmulanswer(3) = "1 0"
strmulanswer(4) = "will give a run time error"
strmulanswer(5) = "OK"
strmulanswer(6) = "good account caution"
strmulanswer(7) = "Runtime error"
strmulanswer(8) = "to choose from multiple possibilities which may arise due to different values of a single variable"
strmulanswer(9) = "0"
strmulanswer(10) = "Displaying integers from 1 to 9"

End Sub

Private Sub multiplechoice_Paper4()
strmultiple(1) = "Consider the program fragments witch(choice){case 'R' : printf  (" & Chr(4) & "RED" & Chr(4) & ");case 'W' : printf (" & Chr(4) & "WHITE" & Chr(4) & ");case 'B' : printf(" & Chr(4) & "BLUE" & Chr(4) & ");default : printf (" & Chr(4) & "ERROR" & Chr(4) & ");break;}What should be the output if choice = 'R'?"
strmuloption(1) = "RED"
strmuloption(2) = "RED ERROR"
strmuloption(3) = "RED WHITE BLUE ERROR"
strmuloption(4) = "RED WHITE BLUE"

strmultiple(2) = "If c is a variable initialized to 1, how many times will the following loop be executed? while((c > 0) && (c < 60)) {   loop body   c++;}"
strmuloption(5) = "60"
strmuloption(6) = "59"
strmuloption(7) = "61"
strmuloption(8) = "1"

strmultiple(3) = "The statement #include <math.h> is written at the top of program to indicate"
strmuloption(9) = "the beginning of the program"
strmuloption(10) = "the program does heavy mathematical calculations"
strmuloption(11) = "that certain information about mathematical calculations functions are to be included at the beginning of the program "
strmuloption(12) = "none of the above"

strmultiple(4) = "If x and y are variables as declared below    double x = 0.005, y = -0.001;   What is the value of ceil(x+y), where ceil is a function to compute ceiling of a number?"
strmuloption(13) = "1"
strmuloption(14) = "0"
strmuloption(15) = "0.005"
strmuloption(16) = "0.5"

strmultiple(5) = "The for statement which can precede a loop to be executed 50 times or till a boolean variable " & Chr(4) & "found" & Chr(4) & " becomes false is given by"
strmuloption(17) = "for (i =0; i<= 50 found == false; i++)"
strmuloption(18) = "for (i =0; i< 50 found == true; i++)"
strmuloption(19) = "for (i =1; i<= 50 found == false; i++)"
strmuloption(20) = "none of the above"

strmultiple(6) = "The declarations typedef float height[100];height men, women;"
strmuloption(21) = "define men and women as 100 element floating point arrays."
strmuloption(22) = "define men and women as floating point variables"
strmuloption(23) = "define height, men and women as floating point variables "
strmuloption(24) = "are illegal"

strmultiple(7) = "Consider the following declarations typedef struct {   char name[20];  char middlename[5]; char surname [20];  } NAME  NAME class [20];"
strmuloption(25) = "class is an array of 20 characters only"
strmuloption(26) = "class is an array of 20 names where each name consists of a name, middlename and surname"
strmuloption(27) = "class is a new type"
strmuloption(28) = "none of the above"

strmultiple(8) = "What would be the values assigned to a,b,c if the statement scanf(" & Chr(4) & "%d %d%d" & Chr(4) & ", &a,&b,&c) is extended with input data item 123456?"
strmuloption(29) = "a = 12, b = 34, c = 56"
strmuloption(30) = "a = 1, b = 2, c = 3"
strmuloption(31) = "a = 123456 and nothing is assigned to b and c"
strmuloption(32) = "a and b are not assigned anything, c = 123456"

strmultiple(9) = "What would be the values assigned to a,b and c if the statement scanf (" & Chr(4) & "%3d,%3d,%3d" & Chr(4) & ",&a&b&c) is executed  with input data as 1234b5678b9( b denotes blank)?"
strmuloption(33) = "a = 123, b = 4, c = 567"
strmuloption(34) = "a = 123, b = 567, c = 9"
strmuloption(35) = "a = 123, b = 456, c = 789"
strmuloption(36) = "a = 1234, b = 5678, c = 9"

strmultiple(10) = " What would be the values of i,x and c if scanf(" & Chr(4) & "%3d,%5f,%c" & Chr(4) & ",&i&x&c) is executed with input data 10b256.875bT?"
strmuloption(37) = "i = 10, b = 56.875, C = T"
strmuloption(38) = "i = 100, b = 256.87, C = T"
strmuloption(39) = "i = 010, b = 256.87, c = '5'"
strmuloption(40) = "i = 10, b = 256.8, C = '7'"



strmulanswer(1) = "RED WHITE BLUE ERROR"
strmulanswer(2) = "59"
strmulanswer(3) = "that certain information about mathematical calculations functions are to be included at the beginning of the program "
strmulanswer(4) = "0"
strmulanswer(5) = "for (i =1; i<= 50 found == false; i++)"
strmulanswer(6) = "define men and women as 100 element floating point arrays."
strmulanswer(7) = "class is an array of 20 names where each name consists of a name, middlename and surname"
strmulanswer(8) = "a = 123456 and nothing is assigned to b and c"
strmulanswer(9) = "a = 123, b = 4, c = 567"
strmulanswer(10) = "i = 10, b = 256.8, C = '7'"
End Sub

Private Sub multiplechoice_Paper5()
strmultiple(1) = "What would be the assignments if char s[100]; int d; float f; fastemer b12345b5"
strmuloption(1) = "s = " & Chr(4) & "fasteber" & Chr(4) & ", d = 12345, f = 0.05"
strmuloption(2) = "s = " & Chr(4) & "fasteber" & Chr(4) & ", d = 123*45, f = 0.05"
strmuloption(3) = "s = " & Chr(4) & "fasteber" & Chr(4) & ",f=0.05"
strmuloption(4) = "s = " & Chr(4) & "fasteber" & Chr(4) & ", d = 0.05"

strmultiple(2) = "What would be the output of the following program fragment?{int i = 12345;float x = 145.678;printf(" & Chr(4) & "%3d,%5d,%8d" & Chr(4) & ",i,i,i);}"
strmuloption(5) = "123 123 123"
strmuloption(6) = "12345 12345 12345"
strmuloption(7) = "12345"
strmuloption(8) = "1 2 3"

strmultiple(3) = " The function fprintf is used in a program"
strmuloption(9) = "when too many printf calls have been already used in the program."
strmuloption(10) = "in place of printf, since printf uses more memory"
strmuloption(11) = "when the output is to be printed on to a file."
strmuloption(12) = "when the type of the variables to be printed are not known."

strmultiple(4) = "In the following statement    fprintf(fpt," & Chr(4) & "%n" & Chr(4) & ",i), the variable fpt is"
strmuloption(13) = "a character variable"
strmuloption(14) = "arbitrarily assigned a value"
strmuloption(15) = "a pointer to a file."
strmuloption(16) = "a special kind of variable called " & Chr(4) & "file" & Chr(4) & "."

strmultiple(5) = "Consider the following definitions struct circle{  float radius;   struct point centre;        } circle1;  struct point {      int x;      int y;          }   One can define a circle by the following assignment statement"
strmuloption(17) = "circle1.radius = 10; circle1.centre.x=0; center.y=0;"
strmuloption(18) = "circle1.radius = 10, circle1.x = circle1.y = 0;"
strmuloption(19) = "circle1.circle.radius = 10;   circle1.circle.centre.point.x =    circle1.circle.centre.point.y = 0;"
strmuloption(20) = "radius = 10; x = 0; y = 0;"

strmultiple(6) = "Consider the following declaration struct list {   int x;  struct list *bext;      }*head; The statement head.x = 100"
strmuloption(21) = "assigns 100 to one element of the structure list"
strmuloption(22) = "creates a node of type list and assigns a value to x"
strmuloption(23) = "creates a head of the list"
strmuloption(24) = "is an erroneous statement."

strmultiple(7) = "How many times will the following loop be executed?c1 = 'a';while(c1 >= 'a' && c1 <= 'z'){c1++;}"
strmuloption(25) = "25"
strmuloption(26) = "26"
strmuloption(27) = "0"
strmuloption(28) = "1"

strmultiple(8) = "What is the following function computing? text(n)int n;{int i;long int prod = 1;if (n > 1)for ( i = 2; i <= n;i ++)prod *= i;return(prod);}"
strmuloption(29) = "calculates the product of any n positive integers"
strmuloption(30) = "calculates the value of factorial n"
strmuloption(31) = "checks whether n is a non negative integer"
strmuloption(32) = "none of these"

strmultiple(9) = "What is the following function doing? test(a,b)int a,b;{int z = (abs(a) > = abs(b)) ? a:b;return(z);}"
strmuloption(33) = "finds the maximum of a and b;"
strmuloption(34) = "returns the minimum of a and b;"
strmuloption(35) = "returns the member whose absolute value is largest"
strmuloption(36) = "none of the above"

strmultiple(10) = " void summation(n)int n;{int i;sum = 0;i = 1;for(i = 1; i < = n; i++)sum +=i;}The above program returns"
strmuloption(37) = "the sum of n numbers"
strmuloption(38) = "the nth number"
strmuloption(39) = "sum of 1,2......,n"
strmuloption(40) = "none of the above"

strmulanswer(1) = "s = " & Chr(4) & "fasteber" & Chr(4) & ",f=0.05"
strmulanswer(2) = "12345 12345 12345"
strmulanswer(3) = "when the output is to be printed on to a file."
strmulanswer(4) = "a pointer to a file."
strmulanswer(5) = "circle1.radius = 10; circle1.centre.x=0; center.y=0;"
strmulanswer(6) = "is an erroneous statement."
strmulanswer(7) = "26"
strmulanswer(8) = "calculates the value of factorial n"
strmulanswer(9) = "returns the member whose absolute value is largest"
strmulanswer(10) = "none of the above"
End Sub

Private Sub multiplechoice_Paper6()
strmultiple(1) = "void Test(void){char c;if ((c=getchar())! = EOLN) Test();putchar(c);return;}The above function"
strmuloption(1) = "prints a line it is read from the standard input, character by character"
strmuloption(2) = "reads a line and prints it one character at a line     "
strmuloption(3) = "reads a line character by character"
strmuloption(4) = "reads a line character by character, reverses it and prints it."

strmultiple(2) = "What will be the output of the following program?{int i = 1234; j = 0177, k = 0xa08c;printf(" & Chr(4) & "%80x\n" & Chr(4) & ",i,j,k)}"
strmuloption(5) = "1234 1777 a08c"
strmuloption(6) = "1234 1abc a08c"
strmuloption(7) = "00001234 01aa ax"
strmuloption(8) = "1234 0777 ao8o"

strmultiple(3) = "What is the following function doing? main(){char libe[80];gets (libe);puts(libe);}"
strmuloption(9) = "prints horizontal straight lines on screen"
strmuloption(10) = "prints 80 vertical lines on screen"
strmuloption(11) = "reads in a line of 80 characters"
strmuloption(12) = "reads and prints lines composed of characters"

strmultiple(4) = "If i,j,k are integers, the scanf function to enter i,j,k such that i is decimal , j is octal and k is hexadecimal would be "
strmuloption(13) = "scanf (" & Chr(4) & "x %x %8x" & Chr(4) & ", i,j,k)"
strmuloption(14) = "scanf(" & Chr(4) & "%d %o %x" & Chr(4) & ",i,j,k);"
strmuloption(15) = "scanf(" & Chr(4) & "%d %o %x" & Chr(4) & ", &i,&j,&k)"
strmuloption(16) = "scanf(" & Chr(4) & "%d %8 %16d" & Chr(4) & ",&i,&j,&k)"

strmultiple(5) = "How many times will the following loop be executed if the input data item is 0.1234?while(c = getchar() ! = 0) {              }"
strmuloption(17) = "infinitely"
strmuloption(18) = "never"
strmuloption(19) = "once"
strmuloption(20) = "5 times"

strmultiple(6) = "The declaration " & Chr(4) & " unsigned u " & Chr(4) & " indicates"
strmuloption(21) = "u is an unsigned character"
strmuloption(22) = "u is an unsigned integer"
strmuloption(23) = "u is a character"
strmuloption(24) = "none of the above"

strmultiple(7) = "A declaration " & Chr(4) & "short int" & Chr(4) & " is used for variables"
strmuloption(25) = "which have a short duration in a program"
strmuloption(26) = "which have short names"
strmuloption(27) = "which may require less storage than normal integers"
strmuloption(28) = "all of the above"

strmultiple(8) = "In case of ordinary int variables"
strmuloption(29) = "the leftmost bit is reserved for sign"
strmuloption(30) = "the rightmost bit is reserved for sign"
strmuloption(31) = "no bit is reserved for sign"
strmuloption(32) = "none of the above"

strmultiple(9) = "What would be the value of c?{int c;float a,b;a = 245.05;b = 40.02;c = a + b;}"
strmuloption(33) = "285.07"
strmuloption(34) = "285.0"
strmuloption(35) = "2850"
strmuloption(36) = "285"

strmultiple(10) = " What would be the value of i and k?{int i,j,k;j = 5;i = 2*j/2;k = 2*(j/2);}"
strmuloption(37) = "i = 5, k = 5"
strmuloption(38) = "i = 4, k = 4"
strmuloption(39) = "i = 5, k = 4"
strmuloption(40) = "i = 4, k = 5"

strmulanswer(1) = "reads a line character by character, reverses it and prints it."
strmulanswer(2) = "1234 1777 a08c"
strmulanswer(3) = "reads and prints lines composed of characters"
strmulanswer(4) = "scanf(" & Chr(4) & "%d %o %x" & Chr(4) & ", &i,&j,&k)"
strmulanswer(5) = "infinitely"
strmulanswer(6) = "u is an unsigned integer"
strmulanswer(7) = "which may require less storage than normal integers"
strmulanswer(8) = "the leftmost bit is reserved for sign"
strmulanswer(9) = "285"
strmulanswer(10) = "i = 5, k = 4"

End Sub

Private Sub multiplechoice_Paper7()
strmultiple(1) = "If ASCII value of 'x' is 120. What is the value of i = ('x' - 'w')/3?"
strmuloption(1) = "1"
strmuloption(2) = "2"
strmuloption(3) = "0"
strmuloption(4) = "4"

strmultiple(2) = "How many times will the following loop be executed?{x = 5;if (x = 1){:x++;}"
strmuloption(5) = "never"
strmuloption(6) = "once"
strmuloption(7) = "5 times"
strmuloption(8) = "infinitely many times"

strmultiple(3) = "If an integer occupies 4 bytes and a character occupies 1 bytes of memory, each element of the following structure would occupy how many bytes? struct name{int age;char name[20];}"
strmuloption(9) = "5"
strmuloption(10) = "24"
strmuloption(11) = "21"
strmuloption(12) = "22"

strmultiple(4) = "A short integer occupies 2 bytes , an ordinary integer 4 bytes and a long integer occupies 8 bytes of memory. A structure is defined as struct TAB{short a;int b;long c;}TABLE [10];Then the total memory requirement for TABLE is"
strmuloption(13) = "14"
strmuloption(14) = "140"
strmuloption(15) = "40"
strmuloption(16) = "24"

strmultiple(5) = "The binary function sqrt operates on a double precision argument. If, i is an integer variable, which one of the following calls would correctly compute sqrt(i)?"
strmuloption(17) = "sqrt((double)i)"
strmuloption(18) = "(double) sqrt(i)"
strmuloption(19) = "(double)(sqrt(i))"
strmuloption(20) = "sqrt(i)"

strmultiple(6) = "Consider a part of a loop as shown  below: for(i = 0; i <=1000; i++){;:if(error<0.005) break;}The above loop will"
strmuloption(21) = "always run for 10000 times"
strmuloption(22) = "will never run for 10000 times"
strmuloption(23) = "may or may not run for 10000 times "
strmuloption(24) = "none of the above"

strmultiple(7) = "if (i>1) c = 2;else c = 3;switch(c) {case 2: printf(" & Chr(4) & "CAUTION" & Chr(4) & ");break;case 3: printf(" & Chr(4) & "GOOD BYE" & Chr(4) & ");break;default : printf(" & Chr(4) & "ERROR" & Chr(4) & ");}What is the output if i = 3 ?"
strmuloption(25) = "CAUTION"
strmuloption(26) = "GOOD BYE"
strmuloption(27) = "ERROR"
strmuloption(28) = "DEFAULT"

strmultiple(8) = "How many times will the following loop be executed = 500;while(x <= 500){x = x - 600;if(x<0) break;}"
strmuloption(29) = "0"
strmuloption(30) = "1"
strmuloption(31) = "500"
strmuloption(32) = "100"

strmultiple(9) = "Consider the following program fragment: for(cpunt=1, sum=0;count<=10;count++){scanf(" & Chr(4) & "%d" & Chr(4) & ",&x);if(c<0) continue;sum += x;}what would be the value of sum with the following input values?1,-1,2,-2,3,-3,4,-4,5,-5"
strmuloption(33) = "0"
strmuloption(34) = "30"
strmuloption(35) = "10"
strmuloption(36) = "15"

strmultiple(10) = " Consider the following program fragment sum = 0;do{scanf(" & Chr(4) & "%d" & Chr(4) & ",&x);if(x>0) printf(" & Chr(4) & "%d" & Chr(4) & ",x);elseif(x==0) break;sum += x;} while (sum < 10);what would bethe output if input is 5,2,0,3,0?"
strmuloption(37) = "5 2"
strmuloption(38) = "5 2 0"
strmuloption(39) = "5 2 0 3 "
strmuloption(40) = "5 2 3"

strmulanswer(1) = "0"
strmulanswer(2) = "infinitely many times"
strmulanswer(3) = "24"
strmulanswer(4) = "140"
strmulanswer(5) = "sqrt((double)i)"
strmulanswer(6) = "may or may not run for 10000 times "
strmulanswer(7) = "CAUTION"
strmulanswer(8) = "1"
strmulanswer(9) = "15"
strmulanswer(10) = "5 2"

End Sub

Private Sub multiplechoice_Paper8()
strmultiple(1) = "If nstart = 10, n = 0 and nstop = 100 what would be he output of the following program fragment? for(i = 0; n< nstop; i++){n = start + i*nstart;printf(" & Chr(4) & "%d" & Chr(4) & ",n);}"
strmuloption(1) = "10"
strmuloption(2) = "10 11 12 ... 100"
strmuloption(3) = "10 20 30 40 ... 110"
strmuloption(4) = "100"

strmultiple(2) = "What would be the output of the following program? sum = 0 ;for(i = 1;i <= 10; i++)if(i%2==0) sum += i;printf(" & Chr(4) & "%d" & Chr(4) & ",sum);"
strmuloption(5) = "45"
strmuloption(6) = "0"
strmuloption(7) = "40"
strmuloption(8) = "30"

strmultiple(3) = "If a =-11 and b = -3. What is the value of a%b?"
strmuloption(9) = "-3"
strmuloption(10) = "-2"
strmuloption(11) = "2"
strmuloption(12) = "3"

strmultiple(4) = "What would be the output of the following program? sum = 0;for(i= -100; i < 0; i++)sum +=abs(i);printf(" & Chr(4) & "%d" & Chr(4) & ",sum);"
strmuloption(13) = "100"
strmuloption(14) = "5050"
strmuloption(15) = "-5050"
strmuloption(16) = "-100"

strmultiple(5) = "What is the value of u1 and u2?int u1,u2;int v = 3;int *pv;u1 = 2*(v+5);pv = &v;u2 = 2*(*pv + 5);"
strmuloption(17) = "u1 = 16, u2 = 16"
strmuloption(18) = "u1 = 8, u2 = 16"
strmuloption(19) = "u1 = 16, u2 = 3"
strmuloption(20) = "u1 = 8,u2 = 3"

strmultiple(6) = "What would be the output of the following program? for(row = 1; row <= 100; row++){for(col = 0; col < 5; ++col)printf(" & Chr(4) & "%d" & Chr(4) & ", row);printf(" & Chr(4) & "\n" & Chr(4) & "); }"
strmuloption(21) = "1 2 3 4 5 }    1 2 3 4 5 } 100 times    1 2 3 4 5 }"
strmuloption(22) = "1 2 3 4 5 .... 100 "
strmuloption(23) = "1 2 3 4 5    6 7 8 9 10    .....   96 97 98 99 100"
strmuloption(24) = "1 6 ..... 96    2 7 ..... 97    3 8 ..... 98    4 9 ..... 99    5 10 ..... 100"

strmultiple(7) = "Consider their following program fragment  if (marks > 80) grade ='A';  else if((marks > 70) && (attendance > 75))     grade = 'B';     else if (attendance < 75) grade = 'R'      if ((marks < 60) && (attendance < 50))      grade = 'F' else grade = 'E';What is the value of grade if marks = 55 and attendance = 80?"
strmuloption(25) = "A"
strmuloption(26) = "B"
strmuloption(27) = "F"
strmuloption(28) = "E"

strmultiple(8) = "Suppose i,j,k are integer variables with values 1,2,3 respectively. What is the value of the expression:  !((j+k) > (i + 5)) ?    "
strmuloption(29) = "5"
strmuloption(30) = "0"
strmuloption(31) = "1"
strmuloption(32) = "6"

strmultiple(9) = "Suppose i = 7, f = 5.5 and c = 'W'. What is the value of the expression:  (c != 'p' i + f < = 10) ?"
strmuloption(33) = "1"
strmuloption(34) = "0"
strmuloption(35) = "-1"
strmuloption(36) = "not defined"

strmultiple(10) = " If ASCII code of '0' is 48. What is the value of:   to ascii('2')?"
strmuloption(37) = "48"
strmuloption(38) = "2"
strmuloption(39) = "0"
strmuloption(40) = "50"

strmulanswer(1) = "10 20 30 40 ... 110"
strmulanswer(2) = "30"
strmulanswer(3) = "-2"
strmulanswer(4) = "5050"
strmulanswer(5) = "u1 = 16, u2 = 16"
strmulanswer(6) = "1 2 3 4 5    6 7 8 9 10    .....   96 97 98 99 100"
strmulanswer(7) = "E"
strmulanswer(8) = "1"
strmulanswer(9) = "1"
strmulanswer(10) = "50"
End Sub

Private Sub multiplechoice_Paper9()
strmultiple(1) = "hat would be the output of the following program if input to he program is 'e'?main(){int lower,uppper;upper = getchar();upper = toupper(;ower);putchar(upper);}"
strmuloption(1) = "3"
strmuloption(2) = "e"
strmuloption(3) = "E"
strmuloption(4) = "nothing"

strmultiple(2) = "The function is lower(char) checks whether a character is in lower case or not. Therefore it should return"
strmuloption(5) = "0 or 1"
strmuloption(6) = "-1,0 or 1"
strmuloption(7) = "a character"
strmuloption(8) = "nothing"

strmultiple(3) = "If S is an array of 80 characters, then their value assigned to S through thestatement scanf(" & Chr(4) & "%s" & Chr(4) & ",S) with input 12345, would be"
strmuloption(9) = "" & Chr(4) & "12345" & Chr(4) & ""
strmuloption(10) = "nothing since 12345 is an integer"
strmuloption(11) = "S is an illegal name for a string"
strmuloption(12) = "%s cannot be used for reading in values of S."

strmultiple(4) = "The function fopen(" & Chr(4) & "filename" & Chr(4) & "," & Chr(4) & "r" & Chr(4) & ") returns"
strmuloption(13) = "nothing"
strmuloption(14) = "a value 0 or 1 depending on whether the file could be opened or not."
strmuloption(15) = "a pointer to FILE filename, if it exits"
strmuloption(16) = "a pointer to new file after creating it."

strmultiple(5) = "The expression ((fpt = fopen(" & Chr(4) & "samples" & Chr(4) & "," & Chr(4) & "w" & Chr(4) & ")) == NULL) would be true if"
strmuloption(17) = "the file sample does not exist while fopen is being executed"
strmuloption(18) = "the file " & Chr(4) & "sample" & Chr(4) & " could not be crated for writing "
strmuloption(19) = "fpt is not declared as a FILE pointer"
strmuloption(20) = "the file " & Chr(4) & "sample" & Chr(4) & " is read only."

strmultiple(6) = "If a is a unsigned integer variable whose value is hx6db7, what is the value of -a?"
strmuloption(21) = "h1hhh1"
strmuloption(22) = "hxhh1"
strmuloption(23) = "hx248"
strmuloption(24) = "hx9248"

strmultiple(7) = "If a = 0x6db7 and b = 0xa726, what is the value of a & b?"
strmuloption(25) = "0xca91"
strmuloption(26) = "0x6db7"
strmuloption(27) = "0xzb92"
strmuloption(28) = "0x2526"

strmultiple(8) = "IF a = 0x6db7 and b = 0xa726, what is the value of a^b?"
strmuloption(29) = "0x58d9"
strmuloption(30) = "0xca91"
strmuloption(31) = "0xceb7"
strmuloption(32) = "0x2526"

strmultiple(9) = "The expression a<< 6 shifts all bits of a six places to the left. If a 0x6db7, what is the value of a<< 6?"
strmuloption(33) = "0xa72b"
strmuloption(34) = "0xa2b"
strmuloption(35) = "0x6dc0"
strmuloption(36) = "0x1111"

strmultiple(10) = " The expression a >> 6 shifts all bits of a size places to the right. What is the value of a >> 6 if a = 0x6db7 ?"
strmuloption(37) = "0x1234"
strmuloption(38) = "0x0001"
strmuloption(39) = "0x0000"
strmuloption(40) = "0x1b6"


strmulanswer(1) = "E"
strmulanswer(2) = "0 or 1"
strmulanswer(3) = "" & Chr(4) & "12345" & Chr(4) & ""
strmulanswer(4) = "a pointer to FILE filename, if it exits"
strmulanswer(5) = "the file " & Chr(4) & "sample" & Chr(4) & " could not be crated for writing "
strmulanswer(6) = "hx9248"
strmulanswer(7) = "0x2526"
strmulanswer(8) = "0xca91"
strmulanswer(9) = "0x6dc0"
strmulanswer(10) = "0x1b6"
End Sub

Private Sub multiplechoice_Paper10()
strmultiple(1) = "The declaration   union id{   char color[12]; int size;} shirt, pant; denotes shirt and pant are variables of type id and"
strmuloption(1) = "each can have a value of color and size"
strmuloption(2) = "each can represent either a 12 character color or a integer size at a time"
strmuloption(3) = "shirt and pant are same as struct variables"
strmuloption(4) = "variable shirt and pant cannot be used simultaneously is a statement."

strmultiple(2) = "Identify the most appropriate sentence to describe union "
strmuloption(5) = "Unions are like structures"
strmuloption(6) = "Unions contain members of different data types which share the same storage area in memory."
strmuloption(7) = "Unions are less frequently used in programs"
strmuloption(8) = "Unions are used for set operations"

strmultiple(3) = "Consider the following declarations. union id{ char color; int size;    }struct {  char country;int date;union id; } flag; To assign a color to a flag, the correct statement would be "
strmuloption(9) = "flag.color = 'W';"
strmuloption(10) = "flag.id.color = 'W';"
strmuloption(11) = "flag.color = 'W';"
strmuloption(12) = "flage.id.color ='White';"

strmultiple(4) = "Consider the following declaration enum colors {black, blue, green};This represents"
strmuloption(13) = "black = 0, blue = 1. green = 2"
strmuloption(14) = "color[1] = 'black',color[2] = 'blue',color[3] = 'green'"
strmuloption(15) = "color = 'black' or color = 'blue' or 'color' = 'green'"
strmuloption(16) = "none of the above"

strmultiple(5) = "With the following declarations enum color{black = -1 blue, green}"
strmuloption(17) = "black = -1, blue = 2, green = 3 "
strmuloption(18) = "black = -1, blue = -2, green = -"
strmuloption(19) = "black = -1, blue = 0, green = 1"
strmuloption(20) = "this is an illegal declaration"

strmultiple(6) = "What is the output of the following program main(){int u = 1, v = 3;printf(" & Chr(4) & "%d%d" & Chr(4) & ",u,v);funct1(&u,&v);printf(" & Chr(4) & "%d%d\n" & Chr(4) & ",u,v);}void funct1(int *pu, int *pv)}*pu = 0;*pv = 0;return;}"
strmuloption(21) = "u = 1 v = 3 u = 1 v = 3"
strmuloption(22) = "u = 1 v = 3 u = 1 v = 1 "
strmuloption(23) = "u = 1 v = 3 u = 0 v = 0"
strmuloption(24) = "u = 3 v = 1 u = 3 v = 1"

strmultiple(7) = "What is the output of the following program ? main(){int i = 0;while(i < 5) {sum(i);i++;i++;}void sum(i)int i;{static k; printf(" & Chr(4) & "%d" & Chr(4) & ",k + 1);k++;}"
strmuloption(25) = "0 1 2 3 4"
strmuloption(26) = "1 2 3 4 5"
strmuloption(27) = "1 3 5 7 9"
strmuloption(28) = "0 2 4 6 8"

strmultiple(8) = "What are the outputs of the following program? main(){int u = 1;int v = 3;void funct1(int u,int v);void funct2(int *pu, int *pv); printf(" & Chr(4) & "u = %d" & Chr(4) & ", v = %d" & Chr(4) & ", u v);funct1(u,v);printf(" & Chr(4) & "u = %d" & Chr(4) & ", v = %d" & Chr(4) & ", u,v,);funct2((&u,&v);printf(" & Chr(4) & "u=%d,v=%d" & Chr(4) & ",u,v);}funct1(int u, int v){u = u +V;v = u - v;u = u - v;}funct2(int *pu, int *pv){*pu = *pu + pv;*pv = *pu - pv;*pu = *pu - pv;"
strmuloption(29) = "u = 1 v = 3 u = 1 v = 3 u = 1 v = 3"
strmuloption(30) = "u = 1 v = 3 u = 3 v = 1 u = 3 v = 1"
strmuloption(31) = "u = 1 v = 3 u = 3 v = 1 u = 1 v = 3"
strmuloption(32) = "u = 1 v = 3 u = 1 v = 3 u = 3 v = 3"

strmultiple(9) = " Consider the following fragment main(){void f1(int a);int i;for (i = 0; i < = 5; i++)f1(i);}void f1(int a){static int k;int prod = 1;while (k < 4) {prod *=a ; k++; }printf(" & Chr(4) & "%d" & Chr(4) & ", prod);}The output ofthe program will be "
strmuloption(33) = "0 1 2 3"
strmuloption(34) = "0 1"
strmuloption(35) = "0 1 2 3 4 5"
strmuloption(36) = "0 1 4 9 16 25"

strmulanswer(1) = "each can represent either a 12 character color or a integer size at a time"
strmulanswer(2) = "Unions contain members of different data types which share the same storage area in memory."
strmulanswer(3) = "flag.id.color = 'W';"
strmulanswer(4) = "black = 0, blue = 1. green = 2"
strmulanswer(5) = "black = -1, blue = 0, green = 1"
strmulanswer(6) = "u = 1 v = 3 u = 0 v = 0"
strmulanswer(7) = "0 2 4 6 8"
strmulanswer(8) = "u = 1 v = 3 u = 1 v = 3 u = 3 v = 3"
strmulanswer(9) = "0 1 2 3"
strmulanswer(10) = ""
End Sub

Private Sub multiplechoice_Paper11()
strmultiple(1) = "When a new element is inserted in the middle of a linked list, which of the following is true?"
strmuloption(1) = "only the elements that appear after the new element need to be moved"
strmuloption(2) = "only the elements that appear before the new element need to be moved."
strmuloption(3) = "the elements that appear before and after the new element need to moved."
strmuloption(4) = "none of the above"

strmultiple(2) = "The function sprintf() works like printf(), but operates on"
strmuloption(5) = "Data in a file;"
strmuloption(6) = "stderr"
strmuloption(7) = "stdin"
strmuloption(8) = "string"

strmultiple(3) = "If the space occupied by two strings s1 in 'C' are respectively m and n, the space occupied by the string obtained by concatenating s1 and s2 is always"
strmuloption(9) = "less than m + n"
strmuloption(10) = "equal to m + n"
strmuloption(11) = "greater than m + n"
strmuloption(12) = "none of the above"

strmultiple(4) = "Which of the following 'C' type is not a primitive data structure?"
strmuloption(13) = "int"
strmuloption(14) = "float"
strmuloption(15) = "char"
strmuloption(16) = "none of the above"

strmultiple(5) = "Suppose DATA array contains 1000000 elements. Using the binary search alogorithm, one requires only about n comparisons to find the location of an item in the DATA array, then n is"
strmuloption(17) = "10"
strmuloption(18) = "45"
strmuloption(19) = "20"
strmuloption(20) = "none of the above"

strmultiple(6) = "If the address of the 8th element in a linked list of integers is 1022, then the  address of 9th element is"
strmuloption(21) = "1024"
strmuloption(22) = "1026"
strmuloption(23) = "1023"
strmuloption(24) = "none of these"

strmultiple(7) = "What will be the value returned by the following function, when it is called with 11?recur(int num)(if ( ( num/2) !=0) return (recur(num/2)*10+num%2);else return 1;}"
strmuloption(25) = "Function does not return any value, because it goes into an infinite loop"
strmuloption(26) = "11"
strmuloption(27) = "1011"
strmuloption(28) = "none of these"

strmultiple(8) = "Consider the following set of elements float x,y;x = 7;y =10;x*=y*= y + 28.5;After the execution of the above set of statements , the value of x will be"
strmuloption(29) = "70"
strmuloption(30) = "2695"
strmuloption(31) = "1995"
strmuloption(32) = "none of these"

strmultiple(9) = "What does the following program statement do on executing the following statement under MS-DOS environment?   f1 = fopen(" & Chr(4) & "c:\bew\tools.dat" & Chr(4) & "," & Chr(4) & "wb" & Chr(4) & ")"
strmuloption(33) = "Opens file tools. data in the directory"
strmuloption(34) = "Return NULL, because the path name is improper."
strmuloption(35) = "Gives syntax error, because the mode of file opening is improper."
strmuloption(36) = "None of the above."

strmultiple(10) = " The output of the following program will be:#include <stdio.h>int x[] = (0,1,2,4};void s1(int *p1,int * p2){    int temp;   temp = *p1; *p1 = *p2;  *p2 = temp;}main(){ int i = 1;  x[i] = 2;   s1(&i,&x[i]);   printf(" & Chr(4) & "%d%d \n,i,x[i]);"
strmuloption(37) = "2 1"
strmuloption(38) = "1 2"
strmuloption(39) = "2 2"
strmuloption(40) = "none of these"

strmulanswer(1) = "none of the above"
strmulanswer(2) = "string"
strmulanswer(3) = "less than m + n"
strmulanswer(4) = "none of the above"
strmulanswer(5) = "20"
strmulanswer(6) = "none of these"
strmulanswer(7) = "1011"
strmulanswer(8) = "2695"
strmulanswer(9) = "Opens file tools. data in the directory"
strmulanswer(10) = "2 1"

End Sub

Private Sub multiplechoice_Paper12()
strmultiple(1) = "Consider  the following program:#include <stdio.h>main()int age;char sex;puts(" & Chr(4) & "What is your age?" & Chr(4) & ");scanf((" & Chr(4) & "%d" & Chr(4) & ",&age);puts(" & Chr(4) & "What is your sex?" & Chr(4) & ");sex=getchar();}What happens when the above program is compiled and executed?"
strmuloption(1) = "Accepts age and sex from standard input."
strmuloption(2) = "Accepts age bit it does not wait to input sex,"
strmuloption(3) = "Produces syntax error"
strmuloption(4) = "None of these"

strmultiple(2) = "Consider the following program:#include >stdio.h" & Chr(4) & "main(){int x = 105, y = 0, z=30;char k[4];strcpy(k," & Chr(4) & "FAIL" & Chr(4) & ");if(x/20==5&&(x*y>100||y+z*2<75))if(x<<2<=256)strcpy(k," & Chr(4) & "PASS" & Chr(4) & ");else strcpy(k," & Chr(4) & "pass" & Chr(4) & ");else strcpy(k," & Chr(4) & "fail" & Chr(4) & ");}The value of k on execution of the above program will be"
strmuloption(5) = "PASS"
strmuloption(6) = "FAIL"
strmuloption(7) = "pass"
strmuloption(8) = "None of the above"

strmultiple(3) = "Which of the following statements is correct regarding following C program?main(){    int i,j,*p,*s;  j=j=30; *(p=&i)*3;  p=(s=p,s=&j,s=&i);  *s+=*p+5;"
strmuloption(9) = "Produces syntax errors"
strmuloption(10) = "*s,*p an i are equal"
strmuloption(11) = "*s,*p and j are equal"
strmuloption(12) = "None of the above"

strmultiple(4) = "Which of the following constants does not produce over flow when assigned to an unsigned long integer variable? I. OXFFFFFFF    ii 42949672961  iii 07777777777"
strmuloption(13) = "all of them"
strmuloption(14) = "only i and iii"
strmuloption(15) = "only i and ii"
strmuloption(16) = "None of them"

strmultiple(5) = "Assume thati,j an k are integer variables and their alias are 8,5 and 0 respectively. What will be the values of variables i and k after executing the following expressions?k=(j>=5) ? (i<5) ? i-j-i; k-j: i;i+=(k)?(i)?(j):(i):(k);"
strmuloption(17) = "-3 and 3"
strmuloption(18) = "3 and -5"
strmuloption(19) = "3 and -3"
strmuloption(20) = "-5 and 3"

strmultiple(6) = "The library function exit() causes an exit from"
strmuloption(21) = "the loop in which it occurs"
strmuloption(22) = "the block in which it occurs"
strmuloption(23) = "the function in which it occurs"
strmuloption(24) = "none of these"

strmultiple(7) = "The getch() library function"
strmuloption(25) = "returns character when any key is pressed"
strmuloption(26) = "returns a character when enter is pressed"
strmuloption(27) = "returns and displays a character on the screen when any key is pressed"
strmuloption(28) = "returns none of the above"

strmultiple(8) = "A default argument has a value that"
strmuloption(29) = "may be supplied by the calling program or function"
strmuloption(30) = "related data items and variables"
strmuloption(31) = "integers with user defined names"
strmuloption(32) = "none of them"

strmultiple(9) = "A structure brings together a group of"
strmuloption(33) = "items of the same data type"
strmuloption(34) = "related data items and  variables"
strmuloption(35) = "integers with user defined names"
strmuloption(36) = "none of them"

strmultiple(10) = " The && and operators"
strmuloption(37) = "compare two numeric values"
strmuloption(38) = "combine two numeric values"
strmuloption(39) = "compare two Boolean values"
strmuloption(40) = "perform none of the above"

strmulanswer(1) = "Accepts age bit it does not wait ti input sex,"
strmulanswer(2) = "pass"
strmulanswer(3) = "Produces syntax errors"
strmulanswer(4) = "all of them"
strmulanswer(5) = "-5 and 3"
strmulanswer(6) = "none of these"
strmulanswer(7) = "returns a character when enter is pressed"
strmulanswer(8) = "may be supplied by the calling program or function"
strmulanswer(9) = "related data items and  variables"
strmulanswer(10) = "compare two Boolean values"

End Sub
Private Sub multiplechoice_Paper13()
strmultiple(1) = "The break statement causes an exit"
strmuloption(1) = "only the innermost loop"
strmuloption(2) = "only from the innermost switch"
strmuloption(3) = "from the innermost loop or switch"
strmuloption(4) = "none of the above"

strmultiple(2) = "In a for loop with a multi statement loop body, semicolon should appear following"
strmuloption(5) = " the for statement itself"
strmuloption(6) = "the crossing brace in the multiple statement loop body"
strmuloption(7) = "each statement within he loop body and the test expression"
strmuloption(8) = "none of the above"

strmultiple(3) = "When an array name is passed to a function , the function"
strmuloption(9) = "access exactly the same array with the same name  as the calling program"
strmuloption(10) = "accesses a copy of the array passed by the program"
strmuloption(11) = "refers to the array using the same name as that used by the calling program"
strmuloption(12) = "does none of the above"

strmultiple(4) = "The first element in a string is"
strmuloption(13) = "the name of the string"
strmuloption(14) = "the first character in the string"
strmuloption(15) = "the length of the string"
strmuloption(16) = "none of the above"

strmultiple(5) = "A pointer is"
strmuloption(17) = "address of a variable"
strmuloption(18) = "an indication of the variable to be accessed text."
strmuloption(19) = "a variable for storing address"
strmuloption(20) = "none of he above"

strmultiple(6) = "The output of main(){int i = 5, j = 5;while(j-->=3){static int i = 10; printf(" & Chr(4) & "%d," & Chr(4) & ",i++);printf(" & Chr(4) & "%d." & Chr(4) & ",i);}"
strmuloption(21) = "5, 6, 7, 8"
strmuloption(22) = "10, 11, 12, 1"
strmuloption(23) = "10, 10, 10, 1"
strmuloption(24) = "none of these"

strmultiple(7) = "The output of main(){intx,y;c = 0x0F;y = 0x0F;while (x = y^x) printf(" & Chr(4) & "d,%d," & Chr(4) & "x,y,y--);printf(" & Chr(4) & "%d." & Chr(4) & ",x);}"
strmuloption(25) = "16, 16, 15, 16"
strmuloption(26) = "0, 16, 15, 0"
strmuloption(27) = "0, 16, 15, 15"
strmuloption(28) = "0 e"

strmultiple(8) = "Static int *(*name[])()"
strmuloption(29) = "name is an array of pointers to function returning pointer to static integer"
strmuloption(30) = "name is a static array of pointers to function returning pointer to integer"
strmuloption(31) = "name is function returning pointer to array o pointers to static integer"
strmuloption(32) = "none of these"

strmultiple(9) = "The output of ain(){int j = 'b' ;switch (j) { case 'a' : putc ('a'); continue;case 'b': putc('b') ;continue; case 'c': putx('e');continue; default : putc('d') ;}}"
strmuloption(33) = "b"
strmuloption(34) = "b infinite times"
strmuloption(35) = "bed"
strmuloption(36) = "none of these"

strmultiple(10) = " What happens when the following C program is executed?#include <stdio.h>main(){FILE *fp1, *fp2; fp1 = fopen(" & Chr(4) & "obe" & Chr(4) & "," & Chr(4) & "w" & Chr(4) & ");fp2 = fopen(" & Chr(4) & "obe" & Chr(4) & "," & Chr(4) & "w" & Chr(4) & "); fputc('A',fp1); fputc('B',fp2);fclose(fp1);fclose(fp2);}"
strmuloption(37) = "results in syntax error - not possible to open same file"
strmuloption(38) = "run time error"
strmuloption(39) = "one will be created with character 'B'"
strmuloption(40) = "none of these"

strmulanswer(1) = "from the innermost loop or switch"
strmulanswer(2) = "each statement within he loop body and the test expression"
strmulanswer(3) = "access exactly the same array with the same name  as the calling program"
strmulanswer(4) = "the first character in the string"
strmulanswer(5) = "address of a variable"
strmulanswer(6) = "none of these"
strmulanswer(7) = "0 e"
strmulanswer(8) = "name is an array of pointers to function returning pointer to static integer"
strmulanswer(9) = "bed"
strmulanswer(10) = "one will be created with character 'B'"

End Sub

Private Sub multiplechoice_Paper14()
strmultiple(1) = "Output of main(){puts(concat)" & Chr(4) & "beWDELHI" & Chr(4) & "," & Chr(4) & "LUCKHNOW" & Chr(4) & ");}char *concat(char *s1,char *s2) {int l1 = strlen(s1),l2 = strlen(S2),i,j;char s[256]; strcpy(s,s1);for (i = l1,j =0; j< 12; j++,i++) s[i] = s2[j]; s[i] = '\0';return s;}"
strmuloption(1) = "beW DELHI LUCKHNOW"
strmuloption(2) = "beWDELHI LUCKHNOW"
strmuloption(3) = "beWDELHI"
strmuloption(4) = "none of these"

strmultiple(2) = "If an array is used as function argument, the array is passed "
strmuloption(5) = "by value"
strmuloption(6) = "by reference"
strmuloption(7) = "None of the above, as array cannot be used a function argument"
strmuloption(8) = "call by name"

strmultiple(3) = "Consider the following program, main(){int c = 50;for(;c;)c--;printf(" & Chr(4) & "%d\n" & Chr(4) & ",c);}The output of the program will be ,"
strmuloption(9) = "50"
strmuloption(10) = "-50"
strmuloption(11) = "0"
strmuloption(12) = "None of the above"

strmultiple(4) = "What is the output of the following program?# define row 3#define col 4int a[row][col] = {1,2,3,4,5,6,7,8,9,10,11,12);main(){int i,j,k = 99;for(i = 0; i < row; i ++)for ( j = 0; j < col; j++)if )a[i][j] < k)k = a[i][j];printf(" & Chr(4) & "%d\n" & Chr(4) & ",k);}"
strmuloption(13) = "99"
strmuloption(14) = "12"
strmuloption(15) = "1"
strmuloption(16) = "none of the above"

strmultiple(5) = "Consider the following program, main(){int x = 0, i,j;for (i = 0, j =10; i < 5, j > 0; i +=2, j--)++x;printf(" & Chr(4) & "%d" & Chr(4) & ",x);}The otput oftheabove program will be,"
strmuloption(17) = "10"
strmuloption(18) = "5"
strmuloption(19) = "0"
strmuloption(20) = "all of the above"

strmultiple(6) = "The output of the following program will be main(){int a, *ptr,b,c;a = 25;ptr = &a;b = a + 30;c = *ptr;printf(" & Chr(4) & "%d,%d,%d" & Chr(4) & ", a,b,c);}"
strmuloption(21) = "25,25,25"
strmuloption(22) = "25,55,25"
strmuloption(23) = "25,55,55"
strmuloption(24) = "none of the above"

strmultiple(7) = "Consider the following program fragment Main(){    int a,b,c;  b = 2;  a = 2 x (b++);  c = 2 x (++b);}The correct values of"
strmuloption(25) = "a = 4, c = 6"
strmuloption(26) = "a = 3, c = 8"
strmuloption(27) = "b = 3 , c = 6"
strmuloption(28) = "a = 4, c = 8"

strmultiple(8) = "The output of the following program will be#include <stdio.h>main(){   int a=10, j=3, i=2; test1(a);   a*=(i+j);   printf(" & Chr(4) & "%d" & Chr(4) & ",a);}test1(x)int x;{int k = 2;x* = k;return(x) ;}"
strmuloption(29) = "20"
strmuloption(30) = "10"
strmuloption(31) = "50"
strmuloption(32) = "none of the above"

strmultiple(9) = "int arr[] = {1,2,3,4};int count;incr() {return ++count;}main(){  arr[count++] = incr();  printf('arr[count = %d\n" & Chr(4) & ",arr[count]);}  The value printed by the above program is"
strmuloption(33) = "1"
strmuloption(34) = "2"
strmuloption(35) = "3"
strmuloption(36) = "none of the above"

strmultiple(10) = " #include <stdio.h>main(){   char ch ='A';   while(ch<='F') {    switch(ch){ case 'A' : case 'B': case 'C': case 'D: ch++; continue;acase 'E': case 'F': ch++;}putchar(ch);}}What will be the output of the above program?"
strmuloption(37) = "ABCDEF will be displayed"
strmuloption(38) = "FG will be displayed"
strmuloption(39) = "EFG will be displayed"
strmuloption(40) = "None of the above"

strmulanswer(1) = "beW DELHI LUCKHNOW"
strmulanswer(2) = "by reference"
strmulanswer(3) = "0"
strmulanswer(4) = "1"
strmulanswer(5) = "10"
strmulanswer(6) = "25,55,25"
strmulanswer(7) = "a = 4, c = 8"
strmulanswer(8) = "50"
strmulanswer(9) = "3"
strmulanswer(10) = "FG will be displayed"
End Sub

Private Sub multiplechoice_Paper15()
strmultiple(1) = "A data file must be closed using"
strmuloption(1) = "library function fprint()"
strmuloption(2) = "Exit function"
strmuloption(3) = "library function fclose()"
strmuloption(4) = "none of the above"

strmultiple(2) = "An existing data file (fold) will be opened for both reading and appending by the following four of the function call -"
strmuloption(5) = "fclose (" & Chr(4) & "fold" & Chr(4) & ",a+)"
strmuloption(6) = "fopen(" & Chr(4) & "fold" & Chr(4) & ",r+);"
strmuloption(7) = "fopen(" & Chr(4) & "fild" & Chr(4) & ",w+)"
strmuloption(8) = "fopen(" & Chr(4) & "fold" & Chr(4) & ",a+);"

strmultiple(3) = "The statement FILE *fpt;"
strmuloption(9) = "defines a pointer to pre defined structure type FILE"
strmuloption(10) = "defines a pointer to user defined structure type FILE."
strmuloption(11) = "defines a points to pre defined data type FILE DESCRIPTOR"
strmuloption(12) = "none of the above"

strmultiple(4) = "The functions fgets and fputs"
strmuloption(13) = "reads and write strings from or to data files"
strmuloption(14) = "reads and write strings from or to input / output stream"
strmuloption(15) = "reads and write records from or to data files."
strmuloption(16) = "none of the above"

strmultiple(5) = "When one dimensional character array of unspecified length is assigned an initial value,"
strmuloption(17) = "an arbitrary character is automatically added to th end of the string"
strmuloption(18) = "'10' is added to the end of the string"
strmuloption(19) = "length of the string is added to the end of the string"
strmuloption(20) = "any of the above"

strmultiple(6) = "When multidimensional arrays are assigned initial value"
strmuloption(21) = "Rightmost subscript increases most  rapidly"
strmuloption(22) = "Leftmost subscript increases most rapidly"
strmuloption(23) = "Rightmost subscript increases least"
strmuloption(24) = "any of the above"

strmultiple(7) = "What would be the value of r if initial value of i,k and j are 2,-3, -3 with reference to the following program segment. if{(i<0) || (j >=0)}{   if(j== k)   r = 'T';}else   if(i ==k)   r = 'F';"
strmuloption(25) = "T"
strmuloption(26) = "F"
strmuloption(27) = "-3"
strmuloption(28) = "2"

strmultiple(8) = "What is the output of the following program fragment?{ sum = 0;do {scanf(" & Chr(4) & "%d" & Chr(4) & ", & i);if(i <0) {i = -i;++flag;}sum +=i;}while(i != 0);}"
strmuloption(29) = "sum of flag number of input numbers"
strmuloption(30) = "sum of (flag - 1) number of input number"
strmuloption(31) = "sum of absolute values of (flag) number of input numbers."
strmuloption(32) = "none of the above"

strmultiple(9) = "In order o check whether an input character is a vowel of not, an efficient multiway decision making routine can be written using"
strmuloption(33) = "while loop"
strmuloption(34) = "switch"
strmuloption(35) = "if - then - else"
strmuloption(36) = "none of the above."

strmultiple(10) = " Consider enum colors {black,blue,iyan,green,red}Enumeration constants (eg. black,blue) represent -"
strmuloption(37) = "integer values such that   black 0   .   .   .   red 4"
strmuloption(38) = "integer values such that    black -1   .   .   .   red 3"
strmuloption(39) = "predefined floating point values"
strmuloption(40) = "none of the above"

strmulanswer(1) = "library function fclose()"
strmulanswer(2) = "fclose (" & Chr(4) & "fold" & Chr(4) & ",a+)"
strmulanswer(3) = "defibes a pointer to pre defined structure type FILE"
strmulanswer(4) = "reads and write strings from or to data files"
strmulanswer(5) = "'10' is added to the end of the string"
strmulanswer(6) = "Rightmost subscript increases most  rapidly"
strmulanswer(7) = "F"
strmulanswer(8) = "sum of absolute values of (flag) number of input numbers."
strmulanswer(9) = "switch"
strmulanswer(10) = "integer values such that   black 0   .   .   .   red 4"

End Sub

Private Sub questions_paper1()
StrQuestions(1) = "Write a program with bested if-else, that will print different messages at different times of day. The messages to be printed are" & Chr(13) & _
"12 midnight - 12 noon - good morning!" & Chr(13) & _
"12 noon - 6 pm - good evening!" & Chr(13) & _
"6 PM - 12 midnight - good night!" & Chr(13) & _
"else - 'Time invalid'" & Chr(13) & _
"Can you provide an implementation using switch statement? Compare both the programs."
StrQuestions(2) = "Write a C program that includes a recursive function to evaluate Legendry polynomials as" & Chr(13) & _
    "P0 = 1, P1 = x;" & Chr(13) & _
    "Pn = (2n - 1)/n * P n-1 - (n-1) / n * Pn-2" & Chr(13) & _
"Legendre polynomials are floating point quantities"
StrQuestions(3) = "Develop Functions in C for following tasks. " & Chr(13) & _
"a. deletion of the last element of a singly linked list" & Chr(13) & _
"b. insertion of an element at the beginning of singly linked list" & Chr(13) & _
"c. deletion of first entry of singly linked list."
StrQuestions(4) = "Develop a function in C that supplies a list into two other lists, so that the entries that were in odd numbered positions re now in one list(in the same relative order as before) and those from even numbered position are in the other new list."
StrQuestions(5) = "The following statement is designed to check the relative sizes of three integers which you may assume to be different -" & Chr(13) & _
"if (x > z) if (x < y) if (y < z) C = 1; else C = z; else" & Chr(13) & _
"if (y < z) C =3; else C = 4; else if ( x < y)" & Chr(13) & _
"if (s < z) C = 5 , else C = 6; else if(y < z) C = 7, else" & Chr(13) & _
"if ( z < x) if (z<y) C = 8; else c =9 ; else C = 10;" & Chr(13) & _
"a. Rewrite this statement in a form that is easier to read" & Chr(13) & _
"b. Identify an eliminate redundant checks" & Chr(13) & _
"c. Write a simpler, shorter statement that accomplishes the same result."
StrQuestions(6) = "A. The first few numbers of the Lucas sequence which is a variation of Fibonacci sequence are :" & Chr(13) & _
"1 3 4 7 11 18 29" & Chr(13) & _
    "a. Design a program in c to generate Lucas sequence" & Chr(13) & _
    "b. Given a =0; b=1 and c = 1 are the first three numbers of some sequence. All other number of this sequence are generated from the sum of their three most recent predecessors" & Chr(13) & _
    "B. Design a program in C, to generate this sequence."
StrQuestions(7) = "Design a program in C to implement insertion sort. The program should use binary search to locate the insertion position."
StrQuestions(8) = "Write a function in C, which converts upper case letters to lower case (without using library routine.)                 "
StrQuestions(9) = "Write a function invert(p,q,t) that returns p with the t bits that begin at position q inverted (I.e. 1 changed to 0 and vice versa), leaving others unchanged."
StrQuestions(10) = "Explain, with examples, following functionalities provided by the C preprocessor." & Chr(13) & _
    "a. file inclusion" & Chr(13) & _
    "b.macro substitution" & Chr(13) & _
    "c. conditional inclusion"

End Sub

Private Sub questions_paper2()
StrQuestions(1) = "a.  Write a 'C' program that will make p point to an integer variable to which q points. The item to which p was formerly pointing will be disposed of." & Chr(13) & _
"b. Write a C function that can compute sum of any geometric series."
StrQuestions(2) = "Write a program to find the particular value of x that causes the function" & Chr(13) & _
    "f(x) = x + cos(x) - 1 - sin(x), to be equal to 0."
StrQuestions(3) = "Suppose a given sum of money. A is placed into a savings bank account at the beginning of each year for n years. If the account earns interest at the rate of I percent annually, then the total amount of money accumulated after n years is given by" & Chr(13) & _
    "A{(1 + i/100) + (1 +1/100)2 + ---(1 + i/100)n }" & Chr(13) & _
    "Write a conversational - style C program to determine how mush money must be deposited at the beginning of each year so that at the rate of 5% interest per year, the money accumulated after 30 years is Rs. 100,000/-"
StrQuestions(4) = "Consider the following list of countries and their capitals" & Chr(13) & _
    "India new Delhi" & Chr(13) & _
    "Bangladesh Dacca" & Chr(13) & _
    "Sri Lanka   Colombo" & Chr(13) & _
    "England London" & Chr(13) & _
    "Germany Berlin" & Chr(13) & _
    "Write a program that will accept the name of a country as input ad output the corresponding capital. The program should run till the string 'END' is encountered & print 'DONT KNOW' for unknown countries."
StrQuestions(5) = "Write a C program that will encrypt a given text in the following manner. For each line of text" & Chr(13) & _
    "a. Convert each character to ASCII equivalents" & Chr(13) & _
    "b. Generate a positive random integer (using library function)" & Chr(13) & _
    "c. Add the integer to ASCII value of each character." & Chr(13) & _
    "d. Let A1 denote lowest permissible value in ASCII code & A2 denote highest permissible value in ASCII code. If the integer generated in step C, exceeds A2, then subtract A2 from the integer & add A1 to the remainder, ensuring that the number will remain between A1 and A2." & Chr(13) & _
    "e. Write out the encrypted character."
StrQuestions(6) = "Write a program that will accept real numbers as input, interactively and maintain the numbers in a linked list in a sorted manner."
StrQuestions(7) = "A graph of n vertices is represented by its adjacency matrix A." & Chr(13) & _
    "A = [aij] is defibed as" & Chr(13) & _
    "aij = 1 if there is an edge between vertex I and vertex j = 0 otherwise." & Chr(13) & _
    "Write a C program which will determine whether a path exists between any two vertices of a given graph."
StrQuestions(8) = "Write A C program that can maintain the name , roll number and marks of a class of students. The size of the class is variable. Include functions to compute the average marks of the class."
StrQuestions(9) = "Write a C program to create a personal file for the employees of an organization. It should contain name, employee code, birth date, educational qualification of the employees. Include a function to calculate the age of any employee on any day."
StrQuestions(10) = "Enhance the above C program by providing facilities to add more employee information, delete information about employees and modify information whenever necessary."
End Sub
Private Sub questions_paper3()
StrQuestions(1) = "Write a C program to maintain a customer account which can keep track of day to days transactions like withdrawal and deposit of money to customers accounts. The program should print the balance of a customer on request."
StrQuestions(2) = "Define a union of type 'ans' containing two members - an integer quantity and a floating quantity. Use this in a structure which is used to create a linked list of members. Compute the average and standard deviation of the numbers and store these in two nodes at the end of the list."
StrQuestions(3) = "Write a complete C program to enter and maintain a family tree. Start with yourself as the root and create up to 'n' levels, where n is to be entered as an input."
StrQuestions(4) = "Assume that a set of members is stored in file. Write a 'C' program to read in the data from this file and compute the mean, median & mode of these members."
StrQuestions(5) = "Write a C program which accepts read number and prints. The sum of digits (complete and fractional parts) of this number."
StrQuestions(6) = "Write a C function for finding the absolute value of the integer parameter passed to it (do not use any library function). Use a switch control structure. Can you enhance your program so that it can handle double precision quantities?"
StrQuestions(7) = "Write a C program that accepts command line arguments, and prints the arguments one by one  and also the number of arguments passed to it. Suggest an application where this can be used effectively."
StrQuestions(8) = "Write a C function to do string concatenation. The function should accept two character pointers corresponding to the first and second string and return a character pointer of the new string obtained by concatenating second to the first. Use dynamic memory allocation to create new string. Assume a library function int strlen(char *) is available to compute string length."
StrQuestions(9) = "Define an abstract data type for handling complex numbers. Use astructure called complex with two member variables of type float. Also write a C function to do complex number addition and multiplication. The function should accept two structure pointers and return back a structure pointer."
StrQuestions(10) = "Write a C program defining an Union which ca hold an 'integer' or 'float' string. Define a variable 'Union type' to keep track of the type of data stored in the union. Write a function to print the value stored in the union on passing a union pointer and the 'union type' variable to this function."
End Sub
Private Sub questions_paper4()
StrQuestions(1) = " What is a self referential structure? Define a structure to implement a node of adoubl linked list, in which, each node contains an alphabetical name and an array of 10 double precision values."
StrQuestions(2) = "Write a C program to sort a data file, whose file name is given as an argument in the command line. The sorted data is to be stored in another file, whose name is also passed as a command line argument."
StrQuestions(3) = "Give a generic method for converting a do- while loop as given below into a while loop." & Chr(13) & _
    "Do" & Chr(13) & _
        "Statements" & Chr(13) & _
    "While (expression);"
StrQuestions(4) = "Write a function which will accept an array of integers as an argument. It should find and return the smallest element in the array after sorting it. The calling program requires the sorted array. The size of the array can be defined to be a global constant."
StrQuestions(5) = "Write an interactive program to determine whether a given positive integer is" & Chr(13) & _
    "a. a prime number" & Chr(13) & _
    "b. a Fibonacci number" & Chr(13) & _
    "Validate the input in your program."
StrQuestions(6) = "a. Develop a function to calculate sum of first n odd integers." & Chr(13) & _
    "b. Develop a function to calculate sum of n even integers starting from a given even integer."
StrQuestions(7) = "Write  an interactive C program that will convert a positive integer quantity to a Roman numeral (i.e. 12 will converted to XII, 14 will be converted to XIV, so on)."
StrQuestions(8) = "Write recursive and iterative functions for computation of factorial of a number. Discuss relative merits and demerits of these implementations."
StrQuestions(9) = "a. What is the purpose of key word void in a function declaration? What is the function of void in function definition?" & Chr(13) & _
    "b. What is the advantage of using function prototypes in a program?" & Chr(13) & _
    "c. Explain meaning of following declarations:" & Chr(13) & _
        "i) int f(int a(;" & Chr(13) & _
        "ii) char f(void);" & Chr(13) & _
        "iii) unsigned f(unsigned a, unsigned b);"
StrQuestions(10) = "Write a C program that will accept two strings as input and compare them to find the length of the greatest common substring between the two."
End Sub
Private Sub questions_paper5()
StrQuestions(1) = "Write a program to find the length of the largest monotonically increasing subsequence in a sequence of real numbers."
StrQuestions(2) = "a.  What is the purpose of statements like" & Chr(13) & _
    "# if !defibed(HDR)" & Chr(13) & _
    "# defibe HDR" & Chr(13) & _
    "/*Contents of HDR*/" & Chr(13) & _
    "#End If" & Chr(13) & _
  "b. Write the code to implement the following." & Chr(13) & _
    "Test the name SYSTEM    to decide which header file is to be included into a file. The decision rules are" & Chr(13) & _
    "Case 1: System = sysv: Include sysv.h" & Chr(13) & _
    "Case 2: System = msdos: Include msdos.h" & Chr(13) & _
    "Case 3: System = bsd: Include bsd, h" & Chr(13) & _
    "Case 4 : Anything ele : Include default.h"
StrQuestions(3) = "Suggest appropriate declarations for each of the following situations involving pointers" & Chr(13) & _
    "a. Declare a function that accepts an argument which is a pointer to an integer quantity and returns a pointer to a six element character array." & Chr(13) & _
    "b. Declare a function that accepts an argument which is an integer array and returns a pointer to a character" & Chr(13) & _
    "c. Declare a 12-element array of pointers to functions. Each function will accept two double precision quantities as argument and will return a double precision quantity." & Chr(13) & _
    "d. Declare a pointe to a function that accepts an argument  which is an array of pointers to integer quantities and returns a pointer to a character."
StrQuestions(4) = "Write a poogram that will accept from the ser bunary, hexadecimal, octal or decimal integer and print it in some other form as requested by the user."
StrQuestions(5) = "Write a C program to convet infix expression to postfix form. Make becessary assumptions(Hint: consider operator precedencerules of C)."
StrQuestions(6) = "Write a C program to evaluate posfix form. Make becessary assumptions(Hint: consider operator precedence rules of C)"
StrQuestions(7) = "Develop a program in C for alphabetically sorting character strings. The program should have the facility for obtaining the strings in an interactive fashion. Collection of the character strings must be represented by an obe dimensional array of pointers."
StrQuestions(8) = "Write a program for analyzing a libe of text by examining each of the characters and determining into which of several different categories it falls, we want  to count  the numbers of vowels, constants , digits , white space characters and other characters."
StrQuestions(9) = "Write a function that will allow a floating point number to be raised to an integer power. In other words we wish to evaluate the expression" & Chr(13) & _
        "y = xn" & Chr(13) & _
    "where y and x are floating point variables and x is in integer variable."
StrQuestions(10) = "Write a loop that will geberate every third integer, beginning with I = 2 and continuing for all integers that are less then 100. Calculate sum of those integers that are evenly divisible by 5. Use two different methods to carry out the test" & Chr(13) & _
    "a. use conditional operator (?:)" & Chr(13) & _
    "b. use an if-else statement"
End Sub

Private Sub questions_paper6()
StrQuestions(1) = "a. Write an appropriate control structure that will examine value of a floating point variable called temp and print one of the following messages, depending on the value assigned to temp -" & Chr(13) & _
    "i) ICE, if the value of temp is less then 0" & Chr(13) & _
    "ii) WATER, if the value of temp lies between 0 and 100." & Chr(13) & _
    "iii) STEAM, if the value of temp exceeds 100." & Chr(13) & _
    "Can a switch statement be used in this problem?" & Chr(13) & _
    "b. Write a loop that will calculate the sum of every third integer, beginning with 2 or all values less then 100. Use do - while statement."
StrQuestions(2) = "a. Suppose pointer variable points to a structure that contains an array as a member. How can an element of the embedded array be accessed?" & Chr(13) & _
    "b. How is a member of a union variable assigned an initial value? In what ways do initializations of union variable differ from that of structure variable?" & Chr(13) & _
    "c. What are the ways by which a complete structure can be returned by a function?"
StrQuestions(3) = "Consider a simple customer billing system. In this system each customer record contains customer's name, street address, city and state, account status (current, overdue or delinquent), previous balance, current payment, new balance and payment date. Customer records need to be maintained and updated as soon as some transactions are made. Status of the account will be determined in the following fashion:" & Chr(13) & _
    "A) if the current payment is greater than zero but less then 10 percent of the previous outstanding balance, the account will be overdue." & Chr(13) & _
    "B) if there is an outstanding balance and the current payment is zero, account will be delinquent" & Chr(13) & _
    "C) Otherwise account is current" & Chr(13) & _
    "Design and develop a C program which will interactively read I customer records, update the customer record for a transaction and then display the updated record."
StrQuestions(4) = "Develop an interactive, file oriented C program that maintains   list of names, address and telephone numbers in alphabetical order. The program should have provision for" & Chr(13) & _
    "a. insertion and deletion of records from the file" & Chr(13) & _
    "b. searching for a particular record in the file."
StrQuestions(5) = "a. Suppose that a is an unsigned integer variable whose value is 0X6aC7. Now extract  the leftmost 6 bits of this value and assign them to the unsigned integer variable. Assign 0 to remaining bits of b. What is the value of b? Write a program for the above task." & Chr(13) & _
    "b. Suppose a is an unsigned integer variable whose value id 0x6db7; shift it to the right by six places and assign the resulting bit pattern to the variable b, What is the value of b? Write a program for the above task."
StrQuestions(6) = "Write a C program that will illustrate equivalence between " & Chr(13) & _
    "a. Shifting a binary number to the left by n bits and multiplying the binary number by 2n" & Chr(13) & _
    "b. shifting a binary number to the right by n bits and dividing the binary number by 2n "
StrQuestions(7) = "Write a C program that will encode and decode contents of a text file by replacing each character with its one's complement."
StrQuestions(8) = "Write  preprocessor directives for each of the following situations;" & Chr(13) & _
    "a. if the symbolic constant BOOLEAN has been defined, define the symbolic constants TRUE and FALSE  so that their values are 1 and 0, respectively and negate the definitions of the symbolic constants YES and NO." & Chr(13) & _
    "b. if flag has value 0, define the symbolic constant COLOR to have value of 1. Otherwise, if the value of flag is less then 3, define COLOR to have value of 2; and if the value of flag equals or exceed 3, define COLOR to have a value of 3."
StrQuestions(9) = "a. What is the top down design methodology? What are its advantages?" & Chr(13) & _
    "b. What is the difference between compilation and execution of a program?" & Chr(13) & _
    "c. What is a debugger?" & Chr(13) & _
    "d. How does the operation of interpreter differ from compiler?" & Chr(13) & _
    "e. What is the function o a linker?"
StrQuestions(10) = "a. Write a C program for calculating volume and surface area of a sphere given diameter of the sphere. Your program must be as general as possible. Design a strategy for testing your program. Provide justifications for your approach." & Chr(13) & _
    "b. Write a for loop that will read characters in a character type array called text and write the characters backwards into another character array called  back text. Assume that text contains 80 characters. Use comma operator within for loop."
End Sub

Private Sub questions_paper7()
StrQuestions(1) = "Write a function to delete the first nod of a singly linked list. Each node in the list stores an alphanumeric string."
StrQuestions(2) = "Explain the Quick sort method. How efficient is Quick sort? Write a C function to sort all the elements in an array X between the position lb and ub(lb is the lower bound and ub is upper bound)."
StrQuestions(3) = "Each node of the linked list stores the employee number and name of the employee in an organization." & Chr(13) & _
    "Write a C program for the following operations" & Chr(13) & _
    "a. Create am alphabetical list of employees" & Chr(13) & _
    "B.Print the Contents of the created Fields"
StrQuestions(4) = "Explain pointers. What are the advantages and disadvantages of linked representation of data over sequential representation?"
StrQuestions(5) = "Write a C program to find the frequency of words from a given text file. The list of words and their corresponding frequency should be in the alphabetical order of words."
StrQuestions(6) = "Write a function mult(n,m,err) that uses the shift operators >> and << to compute and return the value of (n*2m) . Assure that both n and m are unsigned positive integers and are 32 bits in size. The 'err' parameter is set by the function to zero if calculation is successful and 1 if an overflow occurs. An overflow occurs when the resultant value doe not fit in the 32 bit unsigned integer." & Chr(13) & _
    "Do not use add, subtract, multiply or divide to compute the result."
StrQuestions(7) = "Write a program to read a text file WRONG.DAT and copy to another file RIGHT.DAT after making corrections in the text as follows:" & Chr(13) & _
    "a. Read a word from the file WRONG.DAT. A word is a sequence of characters terminated by a space, tab, comma, full stop, semicolon or the newline character." & Chr(13) & _
    "b. Replace the first character of the word to uppercase, if the previous word was terminated by a full stop, else by the lower case alphabet." & Chr(13) & _
    "c. Replace the word 'a' with the word 'an' if the text word starts with a vowel. (Alphabets a,e,i,o,u, or A,E,I,O,U known as vowels)"
StrQuestions(8) = "A data file consists of information about 40 students such as student name(Max 30 characters), student age(integer) and student height(integer). Write a C program to create another file with the records stored in the order of age and height i.e. use age as primary key and height as the secondary key."
StrQuestions(9) = "Write C program to remove all comment lines from a syntax fee source program."
StrQuestions(10) = "Write a C program to implement a CHANGE editor command that is used in the form CHANGE (string1, string2,string3) to replace the first occurrence(if there is one) of string2 in string 1 by string 3."
End Sub

Private Sub questions_paper8()
StrQuestions(1) = "Write a program for accomplishing the following:" & Chr(13) & _
    "Read a six digit positive integer. If he number is even, the add up its digits else multiply the individual digits and print the result."
StrQuestions(2) = "Write a program to read in a string and output the frequency of each character, in that string."
StrQuestions(3) = "a. Write a statement that moves the current position 13 bytes backward." & Chr(13) & _
    "b. Write a decelerator for main() that will enable command line arguments." & Chr(13) & _
    "c. Write a statement that displays the address of the variable testvar." & Chr(13) & _
    "d. What is the difference between intarr[3] and *(intarr+3)?" & Chr(13) & _
    "e. Write a statement that defines a obe dimensional array called double Array of type of type double that holds 100 elements." & Chr(13) & _
    "f. Write a statement that defines an array coins of type int and initializes it to the values of the penny, nickel, dime, quarter, half dollar and dollar." & Chr(13) & _
    "g. What does this statement define?" & Chr(13) & _
        "Employee empist[1000]" & Chr(13) & _
"h. Write a statement that defines that defines an array called many birds that holds 50 objects of type bird."
StrQuestions(4) = "a. Explain the following giving examples." & Chr(13) & _
    "i) do-while ii) switch  iii) for" & Chr(13) & _
     "b. Explain the following" & Chr(13) & _
    "i) fwrite() with examples" & Chr(13) & _
    "ii) single and double dimension array to represent string type of data." & Chr(13) & _
    "iii) Recursion with an example" & Chr(13) & _
    "iv) Pointers to function"
StrQuestions(5) = "a. Write expressions for the following:" & Chr(13) & _
    "i) the address of var" & Chr(13) & _
    "ii) the variable var used as a reference argument" & Chr(13) & _
    "iii) the data type pointer to char" & Chr(13) & _
    "b. Write a definition for a variable of type pointer to float"
StrQuestions(6) = "Write a program to input two strings consisting of maximum 80 characters. Examine both the strings and remove all the common character s from both these strings. Display the resultant strings."
StrQuestions(7) = "Write a function to send all -ve elements o an array to the end without altering the original sequence of i.e. if array contains" & Chr(13) & _
    "5,-4,3,-2,6,-11,12,-8,9" & Chr(13) & _
"then the return array will be" & Chr(13) & _
    "5,3,6,12,9,-4,-2,-11,-8" & Chr(13) & _
"write a C program to test the above function"
StrQuestions(8) = "An employee file consists of information about employees such as employee babe, employee designation, and total salary. Write a C program to create another file with the records stored in the order of designation and total salary i.e. use designation as primary key and total salary as secondary key."
StrQuestions(9) = "A bank maintains their customer database with account number as the key for corresponding records. The signature of each employee is scanned and kept as bitmaps in different binary files having account number as the file names. These bitmaps are in rectangular form of varying sizes. The first two bytes of the bitmaps stores the length and the text two bytes stores the breadth of the image in terms of pixels and each of the following bytes correspond to eight pixels on the bitmap. A function putBitmap(int x,int y, char *p) is available to display such a bit map which is pointed by p on the screen with top left corner of the rectangular bit map at the address x,y where x is the row number , y is the column number." & Chr(13) & _
    "Write a C program to combine all the bitmap files to produce a single file. This file should also contain a table of pointers to the different bitmaps that are available in the file. Provision should be made to add bitmaps corresponding to new customers. Also write a function to display the signature of a customer on the screen, by extracting the corresponding bitmap from this combined file."
StrQuestions(10) = "The student register of a class contains the roll number, name and sex of all the students. The name is sorted as surname first followed by first name and then the second name (like GANDHI SUBHASH MANUBHAI). There can be more that one blank character between surname, first name and second name. The maximum length of the name of a student is 30 characters and it contains only English alphabets and blank spaces. Write a program to print the register in alphabetical order of the student name. In this register the name of each female student must appear in capital letters. For a male student name should appear with first letter as capital and the remaining as small letters (for surname, first name as well as second name). Eg. Gandhi Subhash Manubhai."
End Sub

Private Sub questions_paper9()
StrQuestions(1) = "Write a C program to search for a particular string in a given text file . The program should take the file name, the string to be searched and search option as command line arguments. The program should display the lines that are containing given string along with the line numbers, if the  search option is -a. If the search option is -b, the program should display those lines with line numbers, which are not containing the given string. If the option is -r, the program should ask the user to input the new string that is to be replaced with; and interactively replace its occurrences. (i.e. program should ask the user whether to replace each time the string is found. If user types 'Y' in its response, the string is to be changed; else the string should be kept as such.)"
StrQuestions(2) = "Write a C program to" & Chr(13) & _
    "a. read an input source file and find all those lines which contain an extern string as a first word. Separate all the words contained in these lines to form a list of extern identifiers." & Chr(13) & _
    "b. Search the rest of the file for all those lines which contain any of the extern variables from the list prepared in (a)." & Chr(13) & _
    "c. Extract the C function from the given C source code, for which body of the C function contains any of the extern variables found in the list prepared in (a). Create the output file with these extracted C functions only."
StrQuestions(3) = "a. Write a C program which accepts a real number and prints the sum of digits(complete and fraction parts) o this number." & Chr(13) & _
    "b. Write a C program to read a number having more that 5 digits and print whether the given number is divisible by 11 or not by using the algorithm which states that a number is divisible by 11 if and only if the different of the sums of digits at odd positions and even positions is either zero or divisible by 11."
StrQuestions(4) = "Write a C program to print a file on the standard output. The filename is passed as command line argument. You may use the following standard C library functions." & Chr(13) & _
    "a.fopen b.printf" & Chr(13) & _
    "c. getc and     d. putc"
StrQuestions(5) = "Write a macro to find the gcd(greatest common divisor) of two integers. How does a macro differ from a function? Illustrate the same or the gcd macro."
StrQuestions(6) = "Write a C program to define a structure with a 'int' an 'float' members. Also shoe how you can dynamically allocate memory or this structure in your program."
StrQuestions(7) = "Write a C function called isdigit(int) which should return a non zero if C is a digit and 0 if not"
StrQuestions(8) = "Write the function pow(X,Y) which is a part of standard library. X and Y are of the type double. The function should return the value XY."
StrQuestions(9) = "Write a C program to implement the standard C library function defined by" & Chr(13) & _
        "char *strstr(cs,st)" & Chr(13) & _
    "where 'cs' and 'st' are character pointer. The function should return a pointer to the first occurrence of the string 'st' in 'cs' and NULL if not present."
StrQuestions(10) = "Write a program which reads from the standard input and writes into the standard output after doing the following primitive typesetting" & Chr(13) & _
    "a. Replacing one or more blank between words by a single blank." & Chr(13) & _
    "b. Every sentence should start with an upper case character."

End Sub

Private Sub questions_paper10()
StrQuestions(1) = "Write the C program " & Chr(13) & _
    "int getlibe(char s[], int lim)" & Chr(13) & _
    "Which should read obe full libe from the standard input and store it in the array s of size 'lim'? For lines longer than 'lim' characters it should store the first 'lim' characters. The function return the actual number of the character in the line read."
StrQuestions(2) = "Write a C function to find the determinant value of a matrix." & Chr(13) & _
    "int det(int*,int)" & Chr(13) & _
    "The two dimensional array of matrix element is passed by reference to the function along with the size of the square matrix. The function should return the determinant value."
StrQuestions(3) = "Write the C function to implement the standard library function" & Chr(13) & _
    "int strcmp(str1,str2)" & Chr(13) & _
    "Where str1 and str2 are character pointers. The function should return -1 or 0 or +1 depending on whether str1 is lexicographically 'less than' or 'equal to' or 'greater than' str2."
StrQuestions(4) = "Write a program to search for a given word in a dictionary. The dictionary is a lexicographically sorted array of character strings. Use binary search method. You can make use of the standard C library function to compare two strings."
StrQuestions(5) = "Implement a stack using a one dimensional array. Write procedures to 'push' and 'pop' items into and from the stack. Take care of stack overflow and underflow."
StrQuestions(6) = "Write a program that recognizes whether a telephone number is valid or not by checking for the following criterion. A telephone number is a 10 digit string whose first digit is a '0' followed by 2 digits lying within 1-4. The remaining positions may be occupied by any digit from 0-9, other than the 4th position which can be occupied by any digit from 1-9."
StrQuestions(7) = "Write a program that will read in a text from a file, replace all occurrences of the string ' And ' in the ext by ' Or ' and write the modified text onto another file. (The substrings won't be modified)."
StrQuestions(8) = "Write a program for a grocery store which should accept a unique code for number for an item and the quantity o that item sold to a customer. The unit price of that item is then to be retrieve from a file of prices and then cost is computed. The bill should contain the code for all items bought, their price & the total cost."
StrQuestions(9) = "Write a program that will accept a set of names separated by new line and check whether they are written properly. A name should begin with an upper case alphabet, following which each string in a name should begin with an upper case letter. The only non alphabetic character allowed in a name are '.' And '-' ) period and hyphen)."
StrQuestions(10) = "a. Write a C expression to determine the number of bytes required to store a variable of following type" & Chr(13) & _
    "typedef struct{" & Chr(13) & _
    "char flag;" & Chr(13) & _
    "int  i;" & Chr(13) & _
    "double x;" & Chr(13) & _
    "} mystructType;" & Chr(13) & _
    "b. What is the difference between call by value and call by reference?"

End Sub


