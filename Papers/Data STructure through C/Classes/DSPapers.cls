VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSPapers"
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
 
Public Function DS(PaperType As String, DSPaper As Paper, mCat As Category, Number As Integer)
'ptype is the paper type
'DSpaper is the number of paper as paper1,paper2 ...
'mcat is the category
'number is the number of questions, or options or answers to be taken

 StrPaperType = PaperType
' StrPaper = Paper
If PaperType = "Fill in the Blanks" Then
   If DSPaper = Paper1 Then Fill_in_the_Blanks_Paper1
   If DSPaper = Paper2 Then Fill_in_the_Blanks_Paper2
   If DSPaper = Paper3 Then Fill_in_the_Blanks_Paper3
   If DSPaper = Paper4 Then Fill_in_the_Blanks_Paper4
   If DSPaper = Paper5 Then Fill_in_the_Blanks_Paper5
   If DSPaper = Paper6 Then Fill_in_the_Blanks_Paper6
   If DSPaper = Paper7 Then Fill_in_the_Blanks_Paper7
   If DSPaper = Paper8 Then Fill_in_the_Blanks_Paper8
   If DSPaper = Paper9 Then Fill_in_the_Blanks_Paper9
   If DSPaper = Paper10 Then Fill_in_the_Blanks_Paper10
   If DSPaper = Paper11 Then Fill_in_the_Blanks_Paper11
   If DSPaper = Paper12 Then Fill_in_the_Blanks_Paper12
   If DSPaper = Paper13 Then Fill_in_the_Blanks_Paper13
   If DSPaper = Paper14 Then Fill_in_the_Blanks_Paper14
   If DSPaper = Paper15 Then Fill_in_the_Blanks_Paper15
   If DSPaper = Paper16 Then Fill_in_the_Blanks_Paper16
   If DSPaper = Paper17 Then Fill_in_the_Blanks_Paper17
   If DSPaper = Paper18 Then Fill_in_the_Blanks_Paper18
   If DSPaper = Paper19 Then Fill_in_the_Blanks_Paper19
   If DSPaper = Paper20 Then Fill_in_the_Blanks_Paper20
   
ElseIf PaperType = "Matching Columns" Then
       If DSPaper = Paper1 Then Matching_Columns_Paper1
       If DSPaper = Paper2 Then Matching_Columns_Paper2
       If DSPaper = Paper3 Then Matching_Columns_Paper3
       If DSPaper = Paper4 Then Matching_Columns_Paper4
       If DSPaper = Paper5 Then Matching_Columns_Paper5
       If DSPaper = Paper6 Then Matching_Columns_Paper6
       If DSPaper = Paper7 Then Matching_Columns_Paper7
       If DSPaper = Paper8 Then Matching_Columns_Paper8
       If DSPaper = Paper9 Then Matching_Columns_Paper9
       If DSPaper = Paper10 Then Matching_Columns_Paper10
       If DSPaper = Paper11 Then Matching_Columns_Paper11
       If DSPaper = Paper12 Then Matching_Columns_Paper12
       If DSPaper = Paper13 Then Matching_Columns_Paper13
       If DSPaper = Paper14 Then Matching_Columns_Paper14
       If DSPaper = Paper15 Then Matching_Columns_Paper15
       If DSPaper = Paper16 Then Matching_Columns_Paper16
       If DSPaper = Paper17 Then Matching_Columns_Paper17
       If DSPaper = Paper18 Then Matching_Columns_Paper18
       If DSPaper = Paper19 Then Matching_Columns_Paper19
       
ElseIf PaperType = "True Or False" Then
       If DSPaper = Paper1 Then TrueorFalse_Paper1
       If DSPaper = Paper2 Then TrueorFalse_Paper2
       If DSPaper = Paper3 Then TrueorFalse_Paper3
       If DSPaper = Paper4 Then TrueorFalse_Paper4
       If DSPaper = Paper5 Then TrueorFalse_Paper5
       If DSPaper = Paper6 Then TrueorFalse_Paper6
       If DSPaper = Paper7 Then TrueorFalse_Paper7
       If DSPaper = Paper8 Then TrueorFalse_Paper8
       If DSPaper = Paper9 Then TrueorFalse_Paper9
       If DSPaper = Paper10 Then TrueorFalse_Paper10
       If DSPaper = Paper11 Then TrueorFalse_Paper11
       If DSPaper = Paper12 Then TrueorFalse_Paper12
       If DSPaper = Paper13 Then TrueorFalse_Paper13
       If DSPaper = Paper14 Then TrueorFalse_Paper14
       If DSPaper = Paper15 Then TrueorFalse_Paper15
       If DSPaper = Paper16 Then TrueorFalse_Paper16
       If DSPaper = Paper17 Then TrueorFalse_Paper17
       If DSPaper = Paper18 Then TrueorFalse_Paper18
       If DSPaper = Paper19 Then TrueorFalse_Paper19
       If DSPaper = Paper20 Then TrueorFalse_Paper20
       
ElseIf PaperType = "Descriptive Questions" Then
       If DSPaper = Paper1 Then questions_paper1
       If DSPaper = Paper2 Then questions_paper2
       If DSPaper = Paper3 Then questions_paper3
       If DSPaper = Paper4 Then questions_paper4
       If DSPaper = Paper5 Then questions_paper5
       If DSPaper = Paper6 Then questions_paper6
       If DSPaper = Paper7 Then questions_paper7
       If DSPaper = Paper8 Then questions_paper8
       If DSPaper = Paper9 Then questions_paper9
       
ElseIf PaperType = "Multiple Choice" Then
       If DSPaper = Paper1 Then multiplechoice_Paper1
       If DSPaper = Paper2 Then multiplechoice_Paper2
       If DSPaper = Paper3 Then multiplechoice_Paper3
       If DSPaper = Paper4 Then multiplechoice_Paper4
       If DSPaper = Paper5 Then multiplechoice_Paper5
       If DSPaper = Paper6 Then multiplechoice_Paper6
       If DSPaper = Paper7 Then multiplechoice_Paper7
       If DSPaper = Paper8 Then multiplechoice_Paper8
       If DSPaper = Paper9 Then multiplechoice_Paper9
       If DSPaper = Paper10 Then multiplechoice_Paper10
       If DSPaper = Paper11 Then multiplechoice_Paper11
       If DSPaper = Paper12 Then multiplechoice_Paper12
       If DSPaper = Paper13 Then multiplechoice_Paper13
       If DSPaper = Paper14 Then multiplechoice_Paper14
       If DSPaper = Paper15 Then multiplechoice_Paper15
       If DSPaper = Paper15 Then multiplechoice_Paper16
       If DSPaper = Paper15 Then multiplechoice_Paper17
       If DSPaper = Paper15 Then multiplechoice_Paper18
       If DSPaper = Paper15 Then multiplechoice_Paper19
       If DSPaper = Paper15 Then multiplechoice_Paper20
              
End If

If mCat = Fill_Questions Then
   DS = StrQuestions(Number)
ElseIf mCat = Fill_Options Then
       DS = stroptions(Number)
ElseIf mCat = Fill_Answers Then
       DS = stranswers(Number)
       
ElseIf mCat = Matching_ColumnA Then
       DS = strcolumnA(Number)
ElseIf mCat = Matching_ColumnB Then
       DS = strcolumnB(Number)
ElseIf mCat = Matching_Answers Then
       DS = stranswers(Number)
       
ElseIf mCat = Questions Then
       DS = StrQuestions(Number)
       
ElseIf mCat = TF_Questions Then
       DS = StrQuestions(Number)
ElseIf mCat = TF_Answers Then
       DS = stranswers(Number)
       
ElseIf mCat = MC_Questions Then
       DS = strmuloption(Number)
       
ElseIf mCat = MC_Options_1 Then
       DS = strmuloption(Number)
       
ElseIf mCat = MC_Options_2 Then
       DS = strmuloption(Number)
       
ElseIf mCat = MC_Options_3 Then
       DS = strmuloption(Number)
       
ElseIf mCat = MC_Options_4 Then
       DS = strmuloption(Number)
       
ElseIf mCat = MC_Answer Then
       DS = strmulanswer(Number)
       
End If

End Function

Private Sub Fill_in_the_Blanks_Paper1()
StrQuestions(1) = "dash isthe process ofwriting an algorithm in the correct syntax of a computer language."
StrQuestions(2) = "dash is the processs ofrunning the program on sample data chosen to find errors."
StrQuestions(3) = "An object in terms of a simple case of itself, is called a dash function."
StrQuestions(4) = "In recursivve calls there must be a dash of the sequence."
StrQuestions(5) = "A line of carswaiting at a toll booth isan example of dash concept."
StrQuestions(6) = "The number ofprimitive operations that can be appplied on a queue is dash."
StrQuestions(7) = "In dash the intrinsic ordering of the elements determine the results of its basic operation"
StrQuestions(8) = "An dash is a collection of items into which items can be inserted arbitrarily and from which only the smallest item can be removed."
StrQuestions(9) = "A dash allows deletion of only the largest item."
StrQuestions(10) = "The elements ofa priority queue can be dash that are ordered on one or several fields."

stroptions(1) = "Way out"
stroptions(2) = "Recursive"
stroptions(3) = "Descending priority queue"
stroptions(4) = "cc"
stroptions(5) = "Complex structure"
stroptions(6) = "Coding"
stroptions(7) = "Ascending priority queue"
stroptions(8) = "Testing"
stroptions(9) = "FIFO"
stroptions(10) = "3(three)"

stranswers(1) = "Coding"
stranswers(2) = "Testing"
stranswers(3) = "Recursive"
stranswers(4) = "Way out"
stranswers(5) = "FIFO"
stranswers(6) = "3(three)"
stranswers(7) = "Priority queue"
stranswers(8) = "Ascending priority queue"
stranswers(9) = "Descending priority queue"
stranswers(10) = "omplex structure"

End Sub


Private Sub Fill_in_the_Blanks_Paper2()
StrQuestions(1) = "In top down approach the main program should not become dash entity whose task can not be dexcried."
StrQuestions(2) = "In top down approach, each subprogram should do dash task/s butdo it well."
StrQuestions(3) = "Data structure, defined at logical level iscalled dash data type."
StrQuestions(4) = "Top down and bottom up are two methods used for dash programming."
StrQuestions(5) = "The order of magnitude of function is the same as order of the term in functin that increase dash relative to tN."
StrQuestions(6) = "Array has elements of dash data type."
StrQuestions(7) = "Array may be accessed by dash."
StrQuestions(8) = "Array is dash data structure."
StrQuestions(9) = "If elements of a row are stored next to one another, the aray is said to be stored in dash order."
StrQuestions(10) = "Stack is dash data structure."

stroptions(1) = "Dynamic"
stroptions(2) = "Modular"
stroptions(3) = "Same"
stroptions(4) = "Vague"
stroptions(5) = "Abstract"
stroptions(6) = "Only one"
stroptions(7) = "Fastest"
stroptions(8) = "Index"
stroptions(9) = "Linear"
stroptions(10) = "Row major"

stranswers(1) = "Vague"
stranswers(2) = "Only one"
stranswers(3) = "Abstract"
stranswers(4) = "Modular"
stranswers(5) = "Fastest"
stranswers(6) = "Same"
stranswers(7) = "Index"
stranswers(8) = "Linear"
stranswers(9) = "Row major"
stranswers(10) = "Dynamic"

End Sub

Private Sub Fill_in_the_Blanks_Paper3()
StrQuestions(1) = "Stack has dash structure."
StrQuestions(2) = "dash is used to indicate the end of the list."
StrQuestions(3) = "A list is a dash data structure."
StrQuestions(4) = "The next address field is known as dash."
StrQuestions(5) = "The method of interpreting a bit pattern is called a dash."
StrQuestions(6) = "A queue can be defined as a dash."
StrQuestions(7) = "The problem of Tower of Hanoi is very difficult to solve using dash process."
StrQuestions(8) = "The term headof the queue is same as the term dash ofthe queue."
StrQuestions(9) = "In top down approach, dash consideration is postponed."
StrQuestions(10) = "A stack may ne represented by a dash linked list."

stroptions(1) = "LIFO"
stroptions(2) = "Dynamic"
stroptions(3) = "Pointer"
stroptions(4) = "Null pointer"
stroptions(5) = "Detailed"
stroptions(6) = "Front"
stroptions(7) = "Waiting line"
stroptions(8) = "Linear"
stroptions(9) = "Data type"
stroptions(10) = "Non recursive"

stranswers(1) = "LIFO"
stranswers(2) = "Null pointer"
stranswers(3) = "Dynamic"
stranswers(4) = "Pointer"
stranswers(5) = "Data type"
stranswers(6) = "Waiting line"
stranswers(7) = "Non recursive"
stranswers(8) = "Front"
stranswers(9) = "Detailed"
stranswers(10) = "Linear"

End Sub

Private Sub Fill_in_the_Blanks_Paper4()
StrQuestions(1) = "In linked list, the identitt of next element is dash defined."
StrQuestions(2) = "dash access of element in linked list is not possible."
StrQuestions(3) = "Besides data field, each node of linked list contains at least dash more field."
StrQuestions(4) = "In linked list, successive elements dash occupy adjacent space in memory."
StrQuestions(5) = "dash is used to provide linear order in linked list."
StrQuestions(6) = "End of linked list is marked by putting dash in the next address field ofthe last node."
StrQuestions(7) = "List pointer variable contains the address of dash node."
StrQuestions(8) = "Attempting to delete a node in dash linked list results in underflow."
StrQuestions(9) = "Attempting to create a new node in empty dash in linked listm results in overflow."
StrQuestions(10) = "Traverses, in linked list, is dash to taverses in linear array."

stroptions(1) = "Empty"
stroptions(2) = "Explicitly"
stroptions(3) = "One"
stroptions(4) = "Pointer"
stroptions(5) = "Direct"
stroptions(6) = "Null pointer"
stroptions(7) = "First"
stroptions(8) = "Free space"
stroptions(9) = "Need not"
stroptions(10) = "Array"

stranswers(1) = "Explicitly"
stranswers(2) = "Direct"
stranswers(3) = "One"
stranswers(4) = "Need not"
stranswers(5) = "Pointer"
stranswers(6) = "Null pointer"
stranswers(7) = "First"
stranswers(8) = "Empty"
stranswers(9) = "Free space"
stranswers(10) = "Array"

End Sub
Private Sub Fill_in_the_Blanks_Paper5()
StrQuestions(1) = " Searching is dash  in sorted as well as unsorted linked list."
StrQuestions(2) = "Inserting the node at the dash of the list, in unsorted linked list, is easier."
StrQuestions(3) = "Copying all or part of linked list isdone by taking first null list and inserting appropriate elements dash."
StrQuestions(4) = "Deletion of a node, in linked list, involves keeping track of address of node which immediately dash the node that is to be deleted."
StrQuestions(5) = "Header node in header linked list is always at the dash of the list."
StrQuestions(6) = "In grounded header linked list the last node contains dash."
StrQuestions(7) = "In circular header linked list the last points to dash node."
StrQuestions(8) = "Polynomials in memory can be rpresented by dash list."
StrQuestions(9) = "Given the location ofth node to be deleted, in doubly linked list, it is possible o delete the node dash traversing th part of the list."
StrQuestions(10) = "For representing polynomial in memory, using linked list, each node myst have dash fields."

stroptions(1) = "Header"
stroptions(2) = "Beginning"
stroptions(3) = "Header linked"
stroptions(4) = "Proceeds"
stroptions(5) = "Without"
stroptions(6) = "Null pointer"
stroptions(7) = "3"
stroptions(8) = "Possible"
stroptions(9) = "One by one"
stroptions(10) = "Start(beginning)"

stranswers(1) = "Possible"
stranswers(2) = "Beginning"
stranswers(3) = "One by one"
stranswers(4) = "Proceeds"
stranswers(5) = "Start(beginning)"
stranswers(6) = "Null pointer"
stranswers(7) = "Header"
stranswers(8) = "Header linked"
stranswers(9) = "Null pointer"
stranswers(10) = "3"

End Sub

Private Sub Fill_in_the_Blanks_Paper6()
StrQuestions(1) = "Stack data structure can also be represented through _________  list."
StrQuestions(2) = "Informatioin of most recently inserted nod in _________ can be had by POP operation."
StrQuestions(3) = "FRONT = REAR pointer refers to _________ queue."
StrQuestions(4) = "Inserting a node after a given node, in double linked list, requires _________ pointer changes."
StrQuestions(5) = "Technique for collecting all deleted space in free storage list is known as _________ collection."
StrQuestions(6) = "_________memory allocation is known as dynamic meory allocation."
StrQuestions(7) = "_________ memory allocation is known as static memoty allocation."
StrQuestions(8) = "In _________ method of memoty allocatioin, the first entry, which has free block eaual to or more thn required one, is taken."
StrQuestions(9) = "In _________method of memoty allocation, the entry which is smallest, amongst all the entries which are equal orr bigger than the rqquired one, is chosen. "
StrQuestions(10) = "In best fit method _________ table is used."

stroptions(1) = "Compile time"
stroptions(2) = "Four"
stroptions(3) = "First fit"
stroptions(4) = "Stack"
stroptions(5) = "Best fit"
stroptions(6) = "Run time"
stroptions(7) = "Sorted"
stroptions(8) = "Linked"
stroptions(9) = "Empty"
stroptions(10) = "Garbage"

stranswers(1) = "Linked"
stranswers(2) = "Stack"
stranswers(3) = "Empty"
stranswers(4) = "Four"
stranswers(5) = "Garbage"
stranswers(6) = "Run time"
stranswers(7) = "Compile time"
stranswers(8) = "First fit"
stranswers(9) = "Best fit"
stranswers(10) = "Sorted"

End Sub

Private Sub Fill_in_the_Blanks_Paper7()
StrQuestions(1) = "_________ is most appropriate data structure to represent linked list."
StrQuestions(2) = "_________ to struct is used to represent link of linked list in C."
StrQuestions(3) = "_________ order can be had while building the singly linked list."
StrQuestions(4) = "Initialised _________ pointer variable can be substituted by initialized global variable."
StrQuestions(5) = "_________ linked list cn bescanned in either of direction."
StrQuestions(6) = "Push operation in stack may result in _________."
StrQuestions(7) = "Pop operation in stack may result in _________."
StrQuestions(8) = "_________ functin can be used to allocate memory dynamically."
StrQuestions(9) = "_________ is used to return memory to the system."
StrQuestions(10) = "calloc() and malloc() are very _________ functins, used for allocating memory space."

stroptions(1) = "Overflow "
stroptions(2) = "Underflow"
stroptions(3) = "Doubly"
stroptions(4) = "Malloc()"
stroptions(5) = "Pointer"
stroptions(6) = "Free()"
stroptions(7) = "Static"
stroptions(8) = "Similar"
stroptions(9) = "Struct"
stroptions(10) = "Sorted"

stranswers(1) = "Struct"
stranswers(2) = "Pointer"
stranswers(3) = "Sorted"
stranswers(4) = "Static"
stranswers(5) = "Doubly"
stranswers(6) = "Overflow"
stranswers(7) = "Underflow"
stranswers(8) = "Malloc()"
stranswers(9) = "Free()"
stranswers(10) = "Similar"

End Sub

Private Sub Fill_in_the_Blanks_Paper8()
StrQuestions(1) = "In a tree, a node that has no children is called a _________ node."
StrQuestions(2) = "The maximum level of any leaf in the tree isalso known as the _________ of the tree."
StrQuestions(3) = "Two nodes are _________ if they are left and right children of the same node."
StrQuestions(4) = "A B-tree, of order 11, contains at least _________ keys in each nn root node."
StrQuestions(5) = "In the n-(n-1) tree each node has a maximum of _________ keys and n sons."
StrQuestions(6) = "Applications of a hash function, to a record key, results in a value known as a _________."
StrQuestions(7) = "A B-tree of order n is also called _________ tree."
StrQuestions(8) = "A balanced order-n multiway search tree, in which each nn root node contains at least n/2 keys is called a _________."
StrQuestions(9) = "_________ is a part of compiler hat keeps record of names and thrie attributes."
StrQuestions(10) = "In a _________ each node pointes to two subtrees, one containing keys less than the key n the given node and the other containing keys greater than or equal to the key, in the given node."

stroptions(1) = "Siblings"
stroptions(2) = "N-1"
stroptions(3) = "N-(n-1)"
stroptions(4) = "Depth"
stroptions(5) = "Binary search tree"
stroptions(6) = "Symbol table"
stroptions(7) = "Leaf"
stroptions(8) = "5"
stroptions(9) = "B-tree or order n"
stroptions(10) = "Hash key"

stranswers(1) = "Leaf"
stranswers(2) = "Depth"
stranswers(3) = "Siblings"
stranswers(4) = "5"
stranswers(5) = "N-1"
stranswers(6) = "Hash key"
stranswers(7) = "N-(n-1)"
stranswers(8) = "B-tree or order n"
stranswers(9) = "Symbol table"
stranswers(10) = "Binary search tree"

End Sub

Private Sub Fill_in_the_Blanks_Paper9()
StrQuestions(1) = "The _________  of a node, in a binary tree, is defined as the heightt of its left subtree minus the height of its right subtree."
StrQuestions(2) = "In order transversal of a _________ is performed first for the inorder transversal ofa binary tree."
StrQuestions(3) = "The _________ of a binary tree isthe maximum levvel of its leaves."
StrQuestions(4) = "In an AVL tree the htights oftwh two subtrees ofevery bode never differ by more than _________."
StrQuestions(5) = "A NULL left link, of any node, is replaced with an address of inorder _________ while performing inorder threadingoff a given binary tree."
StrQuestions(6) = "The preorder transversal of a binary tree begins with processing of the _________ "
StrQuestions(7) = "_________ refers to theprocess of deriving storage address from a record key."
StrQuestions(8) = "If a binary tree is threaed for a postorrder transversal, a NULL right link of any node isreplaced by the address of its _________."
StrQuestions(9) = "In _________ the balance ofany node must be 1, -1, or 0."
StrQuestions(10) = "A B-tree of order 24 contains at least _________ keys in each nn root node."

stroptions(1) = "1"
stroptions(2) = "Height"
stroptions(3) = "Root node"
stroptions(4) = "Predecessor"
stroptions(5) = "Left subtree"
stroptions(6) = "12"
stroptions(7) = "Successor"
stroptions(8) = "Balance"
stroptions(9) = "An AVL tree"
stroptions(10) = "Hashing"

stranswers(1) = "Balance"
stranswers(2) = "Left subtree"
stranswers(3) = "Height"
stranswers(4) = "1"
stranswers(5) = "Predecessor"
stranswers(6) = "Root node"
stranswers(7) = "Hashing"
stranswers(8) = "Successor"
stranswers(9) = "An AVL tree"
stranswers(10) = "12"

End Sub

Private Sub Fill_in_the_Blanks_Paper10()
StrQuestions(1) = "A dircted tree, in which outdegree ofenery node isless than or equal to m, is known as _________."
StrQuestions(2) = "In a _________ outdegree of ever node is either 2 or 0."
StrQuestions(3) = "A node containing 3 keys in a multiway search tre must ave _________ subtrees."
StrQuestions(4) = "The maximum difference between the heights of two subtrees on ant node of a balanced binary tree is _________."
StrQuestions(5) = "_________ refers to the operation in which each node of the tree isprocessed exactly once."
StrQuestions(6) = "The level of a root node off any tree is _________."
StrQuestions(7) = "In a _________ the left subtree contains nodes whose keys are less than the key associated with the root node."
StrQuestions(8) = "A_________ tree has no NULL links."
StrQuestions(9) = "In a directed tree, the _________ node has an indegree 0."
StrQuestions(10) = "The minimum number off keys contained in each non root node of a B-tree of orer 15 is _________."

stroptions(1) = "0"
stroptions(2) = "Threaded"
stroptions(3) = "Complete binary tree"
stroptions(4) = "m-ary tree"
stroptions(5) = "1"
stroptions(6) = "7"
stroptions(7) = "Root"
stroptions(8) = "Tree transversal"
stroptions(9) = "Binary search tree"
stroptions(10) = "4"

stranswers(1) = "m-ary tree"
stranswers(2) = "Complete binary tree"
stranswers(3) = "4"
stranswers(4) = "1"
stranswers(5) = "Tree transversal"
stranswers(6) = "0"
stranswers(7) = "Binary search tree"
stranswers(8) = "Threaded"
stranswers(9) = "Root"
stranswers(10) = "7"

End Sub

Private Sub Fill_in_the_Blanks_Paper11()
StrQuestions(1) = "A balanced binary tree is also known as _________."
StrQuestions(2) = "In a _________ tree, NULL linksare replaced itht thrad links."
StrQuestions(3) = "In a B-tree oforder n, each non root node contains at least _________ keys."
StrQuestions(4) = "Total number of cycles in a tree must be _________."
StrQuestions(5) = "In a _________ tree oforder n, each node has lessthan or eaual  to n subtrees and 1 fewer key than it has subtrees."
StrQuestions(6) = "In a _________ tree, the total number of nodes at a level i is always 2 raised to (i-1)."
StrQuestions(7) = "Insertions in an existing lexical ordred Binary tree are performed at a _________ levvel."
StrQuestions(8) = "In a _________ traversal of a tree, the root node is processed last."
StrQuestions(9) = "In a _________ tree, the outdegree of any node mus be either 0 or m."
StrQuestions(10) = "In a _________ tree, each node cn have at most 2 offsprings."

stroptions(1) = "Mutiway search"
stroptions(2) = "Threaded"
stroptions(3) = "Complete m-ary"
stroptions(4) = "Binary"
stroptions(5) = "An AVL tree"
stroptions(6) = "Leaf"
stroptions(7) = "Postorder"
stroptions(8) = "n/2"
stroptions(9) = "Complete binary"
stroptions(10) = "0"

stranswers(1) = "An AVL tree"
stranswers(2) = "Threaded"
stranswers(3) = "n/2"
stranswers(4) = "0"
stranswers(5) = "Mutiway search"
stranswers(6) = "Complete binary"
stranswers(7) = "Leaf"
stranswers(8) = "Postorder"
stranswers(9) = "Complete m-ary"
stranswers(10) = "Binary"

End Sub

Private Sub Fill_in_the_Blanks_Paper12()
StrQuestions(1) = "If graph _________ then diagonal of adjacency matrix has all zeros."
StrQuestions(2) = "_________ is equal to number of 1/s in row or column of adjacency matrix."
StrQuestions(3) = "If x is adjacency matrix of a graph, then i, jth entry of xk, gives _________ between ith and jth vertex."
StrQuestions(4) = "If graph has n vertices and e edges then adjacency matrix ofthe graph has n rows and _________ columns."
StrQuestions(5) = "All edges are scanned before moving to other vertex, in _________ search."
StrQuestions(6) = "In _________ search, move to adjacent univisitedvertex is made as soon as possible."
StrQuestions(7) = "Depth first search is also called _________."
StrQuestions(8) = "Fixed length method of string representatin results in _________."
StrQuestions(9) = "Insertion and deletiion are easily handled in _________ representation"
StrQuestions(10) = "Linked list and _________ method of string representaion allows dynamic allocation of storage."

stroptions(1) = "Index table"
stroptions(2) = "N"
stroptions(3) = "No self loops"
stroptions(4) = "Breadth first"
stroptions(5) = "No. Of edge sequences of length k"
stroptions(6) = "Depth first"
stroptions(7) = "Degree of a vertex of a graph"
stroptions(8) = "Back tracking"
stroptions(9) = "Linked list"
stroptions(10) = "Wastage of space"

stranswers(1) = "No self loops"
stranswers(2) = "Degree of a vertex of a graph"
stranswers(3) = "No. Of edge sequences of length k"
stranswers(4) = "N"
stranswers(5) = "Breadth first"
stranswers(6) = "Depth first"
stranswers(7) = "Back tracking"
stranswers(8) = "Wastage of space"
stranswers(9) = "Linked list"
stranswers(10) = "Index table"

End Sub

Private Sub Fill_in_the_Blanks_Paper13()
StrQuestions(1) = "Possible entries in adjancenct matrix are 1 and dash"
StrQuestions(2) = "If insert operation is used in linked list representatin then inserted string is dash."
StrQuestions(3) = "In internal sorting methods all data resides in dash"
StrQuestions(4) = "Other name for insertin sort is dash."
StrQuestions(5) = "Tie for bubble sort/selectin sort/insertion sort is proprotional to dash where n is no. Of elements to be sorted."
StrQuestions(6) = "In merge sort, file is divided into subfiles which are independently sorted and then dash."
StrQuestions(7) = "In dash the array is partitined in sch a wa that all elements less than or equal to t, appear before t and those appearing after t, are greater than t, where t is partitioning elements."
StrQuestions(8) = "A heap is a tree structure where largest/smallest elemens is available at the dash."
StrQuestions(9) = "Worst case search efficienct of binary search is dash."
StrQuestions(10) = "In binary search keys muxt be dash"

stroptions(1) = "Ordered"
stroptions(2) = "Root"
stroptions(3) = "Quick sort"
stroptions(4) = "n2 "
stroptions(5) = "Shuttle sort"
stroptions(6) = "Merged"
stroptions(7) = "Internal memory"
stroptions(8) = "Not accessible"
stroptions(9) = "Log2 n+1"
stroptions(10) = "0"

stranswers(1) = "0"
stranswers(2) = "Not accessible"
stranswers(3) = "Internal memory"
stranswers(4) = "Shuttle sort"
stranswers(5) = "n2"
stranswers(6) = "Merged"
stranswers(7) = "Quick sort"
stranswers(8) = "Root"
stranswers(9) = "Log2 n+1"
stranswers(10) = "Ordered"

End Sub

Private Sub Fill_in_the_Blanks_Paper14()
StrQuestions(1) = "B-tree indexing uses a dash for keeping index."
StrQuestions(2) = "dash refers to finding the position where a string pattern P first appears in a given string Q."
StrQuestions(3) = "dash was developed to solve problems involving string manipulation."
StrQuestions(4) = "Breadth firstsearc huses dash to hold nodes that are waiting to be processed."
StrQuestions(5) = "Depth first search uses dash to hold nodes that are waiting to be processed."
StrQuestions(6) = "When the number of vvertices of graph changes due to insertion and deletion of vertices, then the sitable representation is dash."
StrQuestions(7) = "Adjacency list is a list formed for each node, connectin it to dash nodes."
StrQuestions(8) = "Adjacency list may ne stored as dash or linked list."
StrQuestions(9) = "index table method of string represenattion creates the problem of dash."
StrQuestions(10) = "A dash is acomplete binary tree with value at each node is at least as much as the valued at children nodes(if they exist)."

stroptions(1) = "Index Sequential Method"
stroptions(2) = "Same address"
stroptions(3) = "Prime"
stroptions(4) = "Collision"
stroptions(5) = "Highly volatile file"
stroptions(6) = "Quadratic"
stroptions(7) = "Main memory"
stroptions(8) = "Linear"
stroptions(9) = "Hashing"
stroptions(10) = "Address"

stranswers(1) = "Main memory"
stranswers(2) = "Address"
stranswers(3) = "Hashing"
stranswers(4) = "Collision"
stranswers(5) = "Linear"
stranswers(6) = "Quadratic"
stranswers(7) = "Prime"
stranswers(8) = "Same address"
stranswers(9) = "Index Sequential Method"
stranswers(10) = "Highly volatile file"

End Sub

Private Sub Fill_in_the_Blanks_Paper15()
StrQuestions(1) = "B-tree indexing uses a dash for keeping index."
StrQuestions(2) = "dash refers to finding the position where a string pattern P first appears in a given string Q."
StrQuestions(3) = "dash was developed to solve problems involving string manipulation."
StrQuestions(4) = "Breadth firstsearc huses dash to hold nodes that are waiting to be processed."
StrQuestions(5) = "Depth first search uses dash to hold nodes that are waiting to be processed."
StrQuestions(6) = "When the number of vvertices of graph changes due to insertion and deletion of vertices, then the sitable representation is dash."
StrQuestions(7) = "Adjacency list is a list formed for each node, connectin it to dash nodes."
StrQuestions(8) = "Adjacency list may ne stored as dash or linked list."
StrQuestions(9) = "index table method of string represenattion creates the problem of dash."
StrQuestions(10) = "A dash is acomplete binary tree with value at each node is at least as much as the valued at children nodes(if they exist)."

stroptions(1) = "Adjacent"
stroptions(2) = "Garbage collection"
stroptions(3) = "Adjacent list"
stroptions(4) = "Queue"
stroptions(5) = "Stack"
stroptions(6) = "Heap"
stroptions(7) = "Array"
stroptions(8) = "Pattern matching"
stroptions(9) = "SNOBOL"
stroptions(10) = "Binary tree"

stranswers(1) = "Binary tree"
stranswers(2) = "Pattern matching"
stranswers(3) = "SNOBOL"
stranswers(4) = "Queue"
stranswers(5) = "Stack"
stranswers(6) = "Adjacent list"
stranswers(7) = "Adjacent"
stranswers(8) = "Array"
stranswers(9) = "Garbage collection"
stranswers(10) = "Heap"

End Sub

Private Sub Fill_in_the_Blanks_Paper16()
StrQuestions(1) = "In strictly traingular matrix the diagonal elements are dash."
StrQuestions(2) = "Heap sort is dash than bubble sort."
StrQuestions(3) = "For a given binary tree, we can have dash threaded binary tree."
StrQuestions(4) = "A tree can be considered as a/an dash threaded binary tree."
StrQuestions(5) = "A priority queue does not follow dash rule strictly."
StrQuestions(6) = "Recursion uses dash as an internal data structure "
StrQuestions(7) = "dash oes not require extra storage."
StrQuestions(8) = "In B-tree of order m, no node has more than dash subtrees."
StrQuestions(9) = "Indexed dash access method is implemented using more than one level of indices."
StrQuestions(10) = "Matrices, in which non zero entries tend tocluster around the diagonal for each ros, are called dash matrix."

stroptions(1) = "Quick sort"
stroptions(2) = "FIFO"
stroptions(3) = "m"
stroptions(4) = "Band"
stroptions(5) = "More than one"
stroptions(6) = "Sequential"
stroptions(7) = "Stack"
stroptions(8) = "Faster"
stroptions(9) = "Acyclic"
stroptions(10) = "Zero"

stranswers(1) = "Zero"
stranswers(2) = "Faster"
stranswers(3) = "More than one"
stranswers(4) = "Acyclic"
stranswers(5) = "FIFO"
stranswers(6) = "Stack"
stranswers(7) = "Quick sort"
stranswers(8) = "m"
stranswers(9) = "Sequential"
stranswers(10) = "Band"

End Sub

Private Sub Fill_in_the_Blanks_Paper17()
StrQuestions(1) = "dash is aimed at enduring efficiency in searching for rcords in a tree."
StrQuestions(2) = "In a tree, any node with out degree zero, can be considered as a/an dash node."
StrQuestions(3) = "dash occurs when a hash functin is biased towards placement of kets in a given region withing the storage space."
StrQuestions(4) = "Time required to access each element of a two dimensional array is dash."
StrQuestions(5) = "To present a complete binary tree of level 3, an array of at least dash elements is necessary."
StrQuestions(6) = "In an array, elements A[i] and A[i+1] are stored in dash  locatins in memory."
StrQuestions(7) = "If we store elements in an array, program of a sorted array containing 10 elements at least dash elements have to be moved."
StrQuestions(8) = "To insert an element at 7th position ofa sorted array containing 10 elements at least dash elements have to be moved."
StrQuestions(9) = "If two polynomials of order 10 nad 5 are stored in arrays A[10,2] and B[5,2] , the array storing their product should be of dimension dash."
StrQuestions(10) = "A mxn matrix can be stored in a/an dash of same size."

stroptions(1) = "Same"
stroptions(2) = "15,2"
stroptions(3) = "Random"
stroptions(4) = "Consecutive"
stroptions(5) = "Array"
stroptions(6) = "3"
stroptions(7) = "Leaf"
stroptions(8) = "15"
stroptions(9) = "Height balancing"
stroptions(10) = "Collision"

stranswers(1) = "Height balancing"
stranswers(2) = "Leaf"
stranswers(3) = "Collision"
stranswers(4) = "Same"
stranswers(5) = "15"
stranswers(6) = "Consecutive"
stranswers(7) = "Random"
stranswers(8) = "3"
stranswers(9) = "15,2"
stranswers(10) = "Array"
End Sub

Private Sub Fill_in_the_Blanks_Paper18()
StrQuestions(1) = "If A is a mxn matrix with k non zero elements where o<k<<mxn , then A can bestored using minimum space in dash arrray."
StrQuestions(2) = "Row major ordering of a/an dash array means rows are stored in asequence, starting from the first element."
StrQuestions(3) = "If an A[3][2], is represented in row mahor order, the element A[1,0] will be located in the dash position from beginning ofthe array."
StrQuestions(4) = "Asequential search for an element in a sorted array would take dash time."
StrQuestions(5) = "Binary search for an element in a sorted array would take dash time."
StrQuestions(6) = "A stack is a ordered list in which all insertions and deletions are at dash end."
StrQuestions(7) = "If the elements A,B,C,D.E are added to a stack in that order the first element to be deleted will be dash."
StrQuestions(8) = "Ifth elements A,B,C,D,E are added to a stack in that order, the first element to be deleted will be dash."
StrQuestions(9) = "If TOP points to the top of stack, the n TOP is dash by 1 when an element is popped from the stack."
StrQuestions(10) = "If A[n] denotes a stack where bottom of the stack is denoted by A[1], then stack overflow condition occurs when an element is dash with TOP equal to n."

stroptions(1) = "Sorted"
stroptions(2) = "Pushed"
stroptions(3) = "A"
stroptions(4) = "Decreased"
stroptions(5) = "4th"
stroptions(6) = "E"
stroptions(7) = "One"
stroptions(8) = "[k+1,3]"
stroptions(9) = "O(n)"
stroptions(10) = "Rectangular"

stranswers(1) = "[k+1,3]"
stranswers(2) = "Rectangular"
stranswers(3) = "4th"
stranswers(4) = "O(n)"
stranswers(5) = "Sorted"
stranswers(6) = "One"
stranswers(7) = "E"
stranswers(8) = "A"
stranswers(9) = "Decreased"
stranswers(10) = "Pushed"

End Sub

Private Sub Fill_in_the_Blanks_Paper19()
StrQuestions(1) = "When processes are scheduledfor execution in a computer they are stored in a dash."
StrQuestions(2) = "A programming language that supposts nested procedure call is implemented using dash."
StrQuestions(3) = "When functions are recursively called during execution, they are executed in dash fashion."
StrQuestions(4) = "Circular queues are implemented usng arays because of dash utilisation of memory space."
StrQuestions(5) = "Postfix expresions are evaluated in a dash scan(or pass) ofthe expression."
StrQuestions(6) = "A dash representation is preferred to an array, when on wants to frequently insert an element in the middle of a sequence."
StrQuestions(7) = "The advantage of linked list representation is that nodes can be created dash."
StrQuestions(8) = "If a set of integers are stored in a linked list as shown below" & Chr(13) & _
                  "[5] dash>[6] dash>[8] dash|" & Chr(13) & _
                  "then, to insert a node containing 7 at its correct place we have to do dash pointer adjustments."
StrQuestions(9) = "The depth of a complete binary tree with n nodes is dash."
StrQuestions(10) = "A tree isa connected graph with no dash."

stroptions(1) = "Linked list"
stroptions(2) = "Floor(logn) + 1"
stroptions(3) = "Dynamically"
stroptions(4) = "Queue"
stroptions(5) = "Two"
stroptions(6) = "Optimum"
stroptions(7) = "Cycles"
stroptions(8) = "LIFO"
stroptions(9) = "Single"
stroptions(10) = "Stacks"

stranswers(1) = "Queue"
stranswers(2) = "Stacks"
stranswers(3) = "LIFO"
stranswers(4) = "Optimum"
stranswers(5) = "Single"
stranswers(6) = "Linked list"
stranswers(7) = "Dynamically"
stranswers(8) = "Two"
stranswers(9) = "Floor(logn) + 1"
stranswers(10) = "Cycles"

End Sub

Private Sub Fill_in_the_Blanks_Paper20()
StrQuestions(1) = "A minimal spanning tree ofa graph with n nodes will have dash nodes."
StrQuestions(2) = "Graphs can be represented by dash lists."
StrQuestions(3) = "When a tree issearched in a breadth first manner, all children of a node v are visited dash siblings of v."
StrQuestions(4) = "Depth first traversal of graph provides a dash tree ofthe graph."
StrQuestions(5) = "Kruskal’s algorithm, for building a minimum cost spanning tree, considers edges for inclusion into thetree in a dash order."
StrQuestions(6) = "In Kruskal’s algrithm, for construction of a minimal cost spanig tree for a graph, the selected edges always form a forest, while in Prim’s algorithm for MST generation, the selected edges always form a dash."
StrQuestions(7) = "It is desirable to keep an extra node at the front of a list which does not represent an item in the list is called dash."
StrQuestions(8) = "A graph can depict dash relationship between nodes."

stroptions(1) = "Spanning tree"
stroptions(2) = "Tree"
stroptions(3) = "One to many"
stroptions(4) = "Adjacency"
stroptions(5) = "After"
stroptions(6) = "Header node"
stroptions(7) = "N-1"
stroptions(8) = "Increasing"

stranswers(1) = "N-1"
stranswers(2) = "Adjacency"
stranswers(3) = "After"
stranswers(4) = "Spanning tree"
stranswers(5) = "Increasing"
stranswers(6) = "Tree"
stranswers(7) = "Header node"
stranswers(8) = "One to many"

End Sub

Private Sub Matching_Columns_Paper1()
strcolumnA(1) = "To identify and develop useful mathematical entities and operations"
strcolumnA(2) = "Stack"
strcolumnA(3) = "Queue"
strcolumnA(4) = "Call itself until conditions are satisfied"
strcolumnA(5) = "Brute force search division of array"
strcolumnA(6) = "Division of array"
strcolumnA(7) = "Order of elements based on the sdquence in which they are inserted"
strcolumnA(8) = "In between natural languge and programme language"
strcolumnA(9) = "Set of test data"
strcolumnA(10) = "A logical module for a specific problem relative to a particular data structure"

strcolumnB(1) = "FIFO list"
strcolumnB(2) = "High level data type"
strcolumnB(3) = "LIFO list"
strcolumnB(4) = "Binary search"
strcolumnB(5) = "Recursion"
strcolumnB(6) = "Liner search"
strcolumnB(7) = "Stack/queue"
strcolumnB(8) = "Verification"
strcolumnB(9) = "Top down approach"
strcolumnB(10) = "Algorithm"

stranswers(1) = "High level data type"
stranswers(2) = "LIFO list"
stranswers(3) = "FIFO list"
stranswers(4) = "Recursion"
stranswers(5) = "Liner search"
stranswers(6) = "Binary search"
stranswers(7) = "Stack/queue"
stranswers(8) = "Algorithm"
stranswers(9) = "Verification"
stranswers(10) = "Top down approach"

End Sub

Private Sub Matching_Columns_Paper2()
strcolumnA(1) = "Postponement ofdetailed consideration"
strcolumnA(2) = "Data structure as defined at logical level"
strcolumnA(3) = "Top Down approach"
strcolumnA(4) = "Efficiency of algorithm"
strcolumnA(5) = "Structured data type having elements of same type"
strcolumnA(6) = "Accessing elements of array"
strcolumnA(7) = "Array"
strcolumnA(8) = "Row major order/column major order"
strcolumnA(9) = "Stack"
strcolumnA(10) = "Dynamic data structure"

strcolumnB(1) = "LIFO"
strcolumnB(2) = "Linked data structure"
strcolumnB(3) = "Algorithm"
strcolumnB(4) = "Index"
strcolumnB(5) = "Array"
strcolumnB(6) = "Stack"
strcolumnB(7) = "Order of magnitude"
strcolumnB(8) = "Abstract data type"
strcolumnB(9) = "Step wise refinement"
strcolumnB(10) = "Linked data structure"

stranswers(1) = "Algorithm"
stranswers(2) = "Abstract data type"
stranswers(3) = "Step wise refinement"
stranswers(4) = "Order of magnitude"
stranswers(5) = "Array"
stranswers(6) = "Index"
stranswers(7) = "Linked data structure"
stranswers(8) = "Linked data structure"
stranswers(9) = "LIFO"
stranswers(10) = "Stack"

End Sub

Private Sub Matching_Columns_Paper3()
strcolumnA(1) = "Array of pointers"
strcolumnA(2) = "Minimize the data movement"
strcolumnA(3) = "Data items in a linked list"
strcolumnA(4) = "Data in the logical order"
strcolumnA(5) = "Data in the physical order"
strcolumnA(6) = "Don’t lose sight of the forest for its trees."
strcolumnA(7) = "List with no codes"
strcolumnA(8) = "List"
strcolumnA(9) = "Array"
strcolumnA(10) = "Method of interpreting a bit"

strcolumnB(1) = "Linked list"
strcolumnB(2) = "Linked list"
strcolumnB(3) = "Pointer sort"
strcolumnB(4) = "Array"
strcolumnB(5) = "Null list"
strcolumnB(6) = "Algorithm"
strcolumnB(7) = "Static data structure"
strcolumnB(8) = "Dynamic data structure"
strcolumnB(9) = "Logical order"
strcolumnB(10) = "Data type"

stranswers(1) = "Pointer sort"
stranswers(2) = "Linked list"
stranswers(3) = "Logical order"
stranswers(4) = "Linked list"
stranswers(5) = "Array"
stranswers(6) = "Algorithm"
stranswers(7) = "Null list"
stranswers(8) = "Dynamic data structure"
stranswers(9) = "Static data structure"
stranswers(10) = "Data type"

End Sub

Private Sub Matching_Columns_Paper4()
strcolumnA(1) = "Identification of next element in linked list"
strcolumnA(2) = "Access of elements in linked list"
strcolumnA(3) = "To determine if list is empty"
strcolumnA(4) = "Memory space allocation for successive elements in linked list"
strcolumnA(5) = "Used to provide linear order in linked list"
strcolumnA(6) = "End of linked list"
strcolumnA(7) = "Address of the first node in linked list"
strcolumnA(8) = "Attempting to delete node in empty lnked list"
strcolumnA(9) = "Attempting to create a node hwne free space is not available"
strcolumnA(10) = "Similarity in traversal of linear array and linked list"

strcolumnB(1) = "List pointer variable"
strcolumnB(2) = "Explicit"
strcolumnB(3) = "Underflow"
strcolumnB(4) = "Operation on list"
strcolumnB(5) = "Overlfow"
strcolumnB(6) = "Pointer"
strcolumnB(7) = "Linear structure"
strcolumnB(8) = "Sequential"
strcolumnB(9) = "Need not be contigous"
strcolumnB(10) = "Null pointer"

stranswers(1) = "Explicit"
stranswers(2) = "Sequential"
stranswers(3) = "Operation on list"
stranswers(4) = "Need not be contigous"
stranswers(5) = "Pointer"
stranswers(6) = "Null pointer"
stranswers(7) = "List pointer variable"
stranswers(8) = "Underflow"
stranswers(9) = "Overlfow"
stranswers(10) = "Linear structure"

End Sub

Private Sub Matching_Columns_Paper5()
strcolumnA(1) = "Easier to insert a node in usorted linked list at"
strcolumnA(2) = "Inserting appropriate elements one by one in null list"
strcolumnA(3) = "In linked list, keeping track of address of the node which immediately proceeds the node, marked for"
strcolumnA(4) = "Special node in jeader linked list"
strcolumnA(5) = "Header linked list with last node having null pointer"
strcolumnA(6) = "Header linked list having last pointing back to header node"
strcolumnA(7) = "Representation of polynomial in memory"
strcolumnA(8) = "Deleting the node whose address is known without traversing the list"
strcolumnA(9) = "Number of fields in a node of linked list representing polynomial"
strcolumnA(10) = "Get a node store new element and insert the node at top"

strcolumnB(1) = "Header linked list"
strcolumnB(2) = "Doubly lnked list"
strcolumnB(3) = "Beginning"
strcolumnB(4) = "3"
strcolumnB(5) = "Deletion"
strcolumnB(6) = "PUSH in stack"
strcolumnB(7) = "Grounded header list"
strcolumnB(8) = "Copy"
strcolumnB(9) = "Header node"
strcolumnB(10) = "Circular header list"

stranswers(1) = "Beginning"
stranswers(2) = "Copy"
stranswers(3) = "Deletion"
stranswers(4) = "Header node"
stranswers(5) = "Grounded header list"
stranswers(6) = "Circular header list"
stranswers(7) = "Header linked list"
stranswers(8) = "Doubly lnked list"
stranswers(9) = "3"
stranswers(10) = "PUSH in stack"

End Sub

Private Sub Matching_Columns_Paper6()
strcolumnA(1) = "Get the alue of most recently inserted node and delete"
strcolumnA(2) = "FRONT = REAR pointer"
strcolumnA(3) = "Four pointer changes in of a node after given node"
strcolumnA(4) = "Technique to collect deleted space in freestorage list"
strcolumnA(5) = "Run time memory allocation"
strcolumnA(6) = "Compile time memory allocation"
strcolumnA(7) = "Smalles amongst all the entries which are qual to or bigger than the required one for memoty allocation"
strcolumnA(8) = "The first entry which is equal to or more than the required one for memory allocation"
strcolumnA(9) = "Memory allocation"
strcolumnA(10) = "    C-data structure to represent linked list"

strcolumnB(1) = "Best fit method"
strcolumnB(2) = "First fit method"
strcolumnB(3) = "POP in stack"
strcolumnB(4) = "Fragmentation"
strcolumnB(5) = "Doubly linked list"
strcolumnB(6) = "Struct"
strcolumnB(7) = "Dynamic memory allocation"
strcolumnB(8) = "Empty queue"
strcolumnB(9) = "Garbage collection"
strcolumnB(10) = "Static memoty allocation  "

stranswers(1) = "POP in stack"
stranswers(2) = "Empty queue"
stranswers(3) = "Doubly linked list"
stranswers(4) = "Garbage collection"
stranswers(5) = "Dynamic memory allocation"
stranswers(6) = "Static memoty allocation "
stranswers(7) = "Best fit method"
stranswers(8) = "First fit method"
stranswers(9) = "Fragmentation"
stranswers(10) = "Struct"

End Sub

Private Sub Matching_Columns_Paper7()
strcolumnA(1) = "For link of a list"
strcolumnA(2) = "Building the list by inserting each new item in the proper sequence of the character"
strcolumnA(3) = "Can’t be traversed in reversed order"
strcolumnA(4) = "Substitute for initialized static pointer variable"
strcolumnA(5) = "Can be scanned in either direction"
strcolumnA(6) = "Underflow in stack"
strcolumnA(7) = "Overflow in stack"
strcolumnA(8) = "For dynamic memoty allocation"
strcolumnA(9) = "To return memoty to the system"
strcolumnA(10) = "    malloc() returning ponter to "

strcolumnB(1) = "PUSH operation"
strcolumnB(2) = "Doubly linked list"
strcolumnB(3) = "Malloc()"
strcolumnB(4) = "Pointer to struct"
strcolumnB(5) = "Free()"
strcolumnB(6) = "Singly linked list"
strcolumnB(7) = "Void"
strcolumnB(8) = "Sorted linked list"
strcolumnB(9) = "Initializd global variable"
strcolumnB(10) = "POP operation"

stranswers(1) = "Pointer to struct"
stranswers(2) = "Sorted linked list"
stranswers(3) = "Singly linked list"
stranswers(4) = "Initializd global variable"
stranswers(5) = "Doubly linked list"
stranswers(6) = "POP operation"
stranswers(7) = "PUSH operation"
stranswers(8) = "Malloc()"
stranswers(9) = "Free()"
stranswers(10) = "Void"

End Sub
Private Sub Matching_Columns_Paper8()
strcolumnA(1) = "A node with no subtrees"
strcolumnA(2) = "A dictionary containing names and their associated values"
strcolumnA(3) = "A node with indegree 0 in a directed tree"
strcolumnA(4) = "A tree in which outdegree of each node is less than or equal to 2."
strcolumnA(5) = "Outdegree of each node is either 2 or 0"
strcolumnA(6) = "Root node isprocessed at last."
strcolumnA(7) = "Threaded binary trees."
strcolumnA(8) = "Root node is processed first."
strcolumnA(9) = "Inorder traversal of left subtree is performed first"
strcolumnA(10) = "A balanced binary tree."

strcolumnB(1) = "Postorder traversal"
strcolumnB(2) = "Binary tree"
strcolumnB(3) = "Root node"
strcolumnB(4) = "Inorder traversal"
strcolumnB(5) = "Preorder traversal"
strcolumnB(6) = "Leaf node"
strcolumnB(7) = "Symbol table"
strcolumnB(8) = "AVL tree"
strcolumnB(9) = "No NULL links"
strcolumnB(10) = "Complete binary tree"

stranswers(1) = "Leaf node"
stranswers(2) = "Symbol table"
stranswers(3) = "Root node"
stranswers(4) = "Binary tree"
stranswers(5) = "Complete binary tree"
stranswers(6) = "Postorder traversal"
stranswers(7) = "No NULL links"
stranswers(8) = "Preorder traversal"
stranswers(9) = "Inorder traversal"
stranswers(10) = "AVL tree"

End Sub
Private Sub Matching_Columns_Paper9()
strcolumnA(1) = "Maximum level of leaves of a tree"
strcolumnA(2) = "Height of left subtree of a node minum height of its right subtree."
strcolumnA(3) = "Binary search tree."
strcolumnA(4) = "Each non root node contains at least n/2 keys."
strcolumnA(5) = "Tree that contains at least 5 keys in each non root node"
strcolumnA(6) = "B-tree of order n."
strcolumnA(7) = "A function that transforms a key into a table index, at which record is placed"
strcolumnA(8) = "Occurs for records having the same hash key"
strcolumnA(9) = "NULL links are replaced with thread links."
strcolumnA(10) = "Each node has less than or equal to n subtrees, and 1 fewer key than it has subtrees."

strcolumnB(1) = "B-tree of order 11"
strcolumnB(2) = "Hash function"
strcolumnB(3) = "Balance of a tree"
strcolumnB(4) = "Multiway search tree of order n"
strcolumnB(5) = "Hash collision"
strcolumnB(6) = "B-tree of order n"
strcolumnB(7) = "An n-(n-1) tree"
strcolumnB(8) = "Height of a tree"
strcolumnB(9) = "Threaded tree"
strcolumnB(10) = "Lexical ordered binary tree"

stranswers(1) = "Height of a tree"
stranswers(2) = "Balance of a tree"
stranswers(3) = "Lexical ordered binary tree"
stranswers(4) = "B-tree of order n"
stranswers(5) = "B-tree of order 11"
stranswers(6) = "An n-(n-1) tree"
stranswers(7) = "Hash function"
stranswers(8) = "Hash collision"
stranswers(9) = "Threaded tree"
stranswers(10) = "Multiway search tree of order n"

End Sub
Private Sub Matching_Columns_Paper10()
strcolumnA(1) = "Tree consisting of nodes which can have at most 2 offsprings."
strcolumnA(2) = "Tree in which heights of the two subtrees of any node neer differ by more than 1."
strcolumnA(3) = "Level at which insertion in an existing lexically ordered binary tree is performed"
strcolumnA(4) = "A tree that has structureal links and thread links both. "
strcolumnA(5) = "A value resulting from application of a hash function to a record key."
strcolumnA(6) = "Process of deriving a storage address from a record key"
strcolumnA(7) = "Left & right children of the same node."
strcolumnA(8) = "Each node of this tree has a maximum of n-1 keys and n sons"
strcolumnA(9) = "Minimum number of keys contained in each non root node of a B-tree of order 15."
strcolumnA(10) = "Total number of cycles in a tree"

strcolumnB(1) = "threaded tree"
strcolumnB(2) = "Hashing"
strcolumnB(3) = "Siblings"
strcolumnB(4) = "Binary tree"
strcolumnB(5) = "AVL tree"
strcolumnB(6) = "N-(n-1) tree"
strcolumnB(7) = "Hash key "
strcolumnB(8) = "Leaf level"
strcolumnB(9) = "0"
strcolumnB(10) = "7"

stranswers(1) = "Binary tree"
stranswers(2) = "AVL tree"
stranswers(3) = "Leaf level"
stranswers(4) = "threaded tree"
stranswers(5) = "Hash key "
stranswers(6) = "Hashing"
stranswers(7) = "Siblings"
stranswers(8) = "N-(n-1) tree"
stranswers(9) = "7"
stranswers(10) = "0"

End Sub
Private Sub Matching_Columns_Paper11()
strcolumnA(1) = "Level of a root node of a tree"
strcolumnA(2) = "A directed tree in which outdegree of every node is less than or equal to m."
strcolumnA(3) = "Part of a compiler that keeps track ofnames and their attributes"
strcolumnA(4) = "Outdegree of every node 2 or 0 in this three"
strcolumnA(5) = "A NULL left link of a node is replaced by the address of its inorder predecessor."
strcolumnA(6) = "A tree in which balance ofeach node must be either 1, -1 or 0."
strcolumnA(7) = "Its left subtree contains nodes whose keys are lexically less than the key associated with root node."
strcolumnA(8) = "Number of subtrees for a node containing 3 keys in a multiway search tree. "
strcolumnA(9) = "Maximum difference beween the heights of two subtrees of any node of a balanced binary tree."
strcolumnA(10) = "Each node of the tree is processed exactly once."

strcolumnB(1) = "AVL tree"
strcolumnB(2) = "M-ary tree"
strcolumnB(3) = "Symbol table"
strcolumnB(4) = "Inorder threading"
strcolumnB(5) = "Lexical ordered binary tree"
strcolumnB(6) = "4"
strcolumnB(7) = "0"
strcolumnB(8) = "Complete binary tree"
strcolumnB(9) = "Tree traversal"
strcolumnB(10) = "1"

stranswers(1) = "0"
stranswers(2) = "M-ary tree"
stranswers(3) = "Symbol table"
stranswers(4) = "Complete binary tree"
stranswers(5) = "Inorder threading"
stranswers(6) = "AVL tree"
stranswers(7) = "Lexical ordered binary tree"
stranswers(8) = "4"
stranswers(9) = "1"
stranswers(10) = "Tree traversal"

End Sub
Private Sub Matching_Columns_Paper12()
strcolumnA(1) = "A vertex with degree 1"
strcolumnA(2) = "Used to store indices"
strcolumnA(3) = "In this sorting rechnique the ata set is spslit in two parts & each part is individually sorted"
strcolumnA(4) = "Used in linked collision processing to store synonyms"
strcolumnA(5) = "this sort technique reorders he data set elements in such a way that all elements less than partitioning elelents are in one part and greater of equal elements are in other part."
strcolumnA(6) = "In this digraph indegree ofa vertex is equal to outdegree of that vertex"
strcolumnA(7) = "This sort finds the location LOC of smallest element  in A(K)…A(N) and then interchange A(LOC) with A(K), 2<K<n"
strcolumnA(8) = "Search method with worst case efficience n log n+1"
strcolumnA(9) = "This sort puts each elements A(K) into its proper position in the previoysly sorted array A(1),…a(K-1), 2<K<n."
strcolumnA(10) = "    In this method of collision processing, all synonyms are chained together"

strcolumnB(1) = "Balanced digraph"
strcolumnB(2) = "B+ -Tree"
strcolumnB(3) = "Binary search"
strcolumnB(4) = "Pendant vertex"
strcolumnB(5) = "Merge sort"
strcolumnB(6) = "Selection sort"
strcolumnB(7) = "Linked collision processing"
strcolumnB(8) = "Quick sort"
strcolumnB(9) = "Overflow area"
strcolumnB(10) = "Insertion sort"

stranswers(1) = "Pendant vertex"
stranswers(2) = "B+ -Tree"
stranswers(3) = "Merge sort"
stranswers(4) = "Overflow area"
stranswers(5) = "Quick sort"
stranswers(6) = "Balanced digraph"
stranswers(7) = "Selection sort"
stranswers(8) = "Binary search"""
stranswers(9) = "Insertion sort"
stranswers(10) = "Linked collision processing"

End Sub
Private Sub Matching_Columns_Paper13()
strcolumnA(1) = "Also known as sort within sort"
strcolumnA(2) = "Also known as back tracking"
strcolumnA(3) = "A symmetric bit matrix"
strcolumnA(4) = "Complete binary tree with value at aprent node is at least as much as the values at children nodes (if they exist)"
strcolumnA(5) = "Number of edges incident on the vertex"
strcolumnA(6) = "ertex with no incident edges."
strcolumnA(7) = "In this traversal all incident edges on the vertex are scanned before moving to other vertex."
strcolumnA(8) = "This method of graph representation is useful if total number of edges is small"
strcolumnA(9) = "String representation which results in wastage of memory."
strcolumnA(10) = "    Key to address transformation"

strcolumnB(1) = "Hashing"
strcolumnB(2) = "Breadth first"
strcolumnB(3) = "Isolated vertex"
strcolumnB(4) = "Adjacency list"
strcolumnB(5) = "Depth first search"
strcolumnB(6) = "Max Heap"
strcolumnB(7) = "Degree of a vertex"
strcolumnB(8) = "Adjacency matrix"
strcolumnB(9) = "Multiple sorting"
strcolumnB(10) = "Fixed length representation for strings"

stranswers(1) = "Multiple sorting"
stranswers(2) = "Depth first search"
stranswers(3) = "Adjacency matrix"
stranswers(4) = "Max Heap"
stranswers(5) = "Fixed length representation for strings"
stranswers(6) = "Isolated vertex"
stranswers(7) = "Breadth first"
stranswers(8) = "Adjacency list"
stranswers(9) = "Fixed length representation for strings"
stranswers(10) = "Hashing"

End Sub
Private Sub Matching_Columns_Paper14()
strcolumnA(1) = "Zero matrix"
strcolumnA(2) = "In this collision processing some addresses may remain unchecked for acancy."
strcolumnA(3) = "Keys generating same address in hashing"
strcolumnA(4) = "This method puts the key causing collision at first vacant position"
strcolumnA(5) = "A way to represent graphs in wich the start and end vertices of all edges are to be given."
strcolumnA(6) = "This occurs when two or more keys generated same address."
strcolumnA(7) = "Way to represent strings useful when lots of insertions & deletions take place"
strcolumnA(8) = "Sorting techniques in which the complete data must be in memory."
strcolumnA(9) = "    String operation which puts two strings together"
strcolumnA(10) = "    A Hashing technique which makes use of prime number"

strcolumnB(1) = "Synonym"
strcolumnB(2) = "Edge listing"
strcolumnB(3) = "Linear collision processing"
strcolumnB(4) = "Collision"
strcolumnB(5) = "Adjacency matrix for a graph without edges"
strcolumnB(6) = "Quadrativ collision processing"
strcolumnB(7) = "Concatenation"
strcolumnB(8) = "Division remainder method"
strcolumnB(9) = "Internal sorting techniques"
strcolumnB(10) = "Linked list representation"

stranswers(1) = "Adjacency matrix for a graph without edges"
stranswers(2) = "Quadrativ collision processing"
stranswers(3) = "Synonym"
stranswers(4) = "Linear collision processing"
stranswers(5) = "Edge listing"
stranswers(6) = "Collision"
stranswers(7) = "Linked list representation"
stranswers(8) = "Internal sorting techniques"
stranswers(9) = "Concatenation"
stranswers(10) = "Division remainder method"

End Sub
Private Sub Matching_Columns_Paper15()
strcolumnA(1) = "Used in depth first search"
strcolumnA(2) = "Searching technique using levels of indexes."
strcolumnA(3) = "A string"
strcolumnA(4) = "Used in breadth first search"
strcolumnA(5) = "This search technique has, on an average, n/2 comparisons."
strcolumnA(6) = "A representation of graph in the form of n* e matrix where, n is number of vertices and e is number of edges."
strcolumnA(7) = "Finding position where a string pattern P first appears in given string Q"
strcolumnA(8) = "SNOBOL"
strcolumnA(9) = "This problem arises in index table method of string representation"
strcolumnA(10) = "    Arranging data according to some predetermined order."

strcolumnB(1) = "Garbage collection"
strcolumnB(2) = "String manipulation language"
strcolumnB(3) = "Array of characters"
strcolumnB(4) = "Sorting"
strcolumnB(5) = "Stack"
strcolumnB(6) = "Sequential search"
strcolumnB(7) = "ISAM"
strcolumnB(8) = "Queue"
strcolumnB(9) = "Adjoint matrix"
strcolumnB(10) = "Pattern matching"

stranswers(1) = "Sequential search"
stranswers(2) = "ISAM"
stranswers(3) = "Array of characters"
stranswers(4) = "Queue"
stranswers(5) = "Sequential search"
stranswers(6) = "Adjoint matrix"
stranswers(7) = "Pattern matching"
stranswers(8) = "String manipulation language"
stranswers(9) = "Garbage collection"""
stranswers(10) = "Sorting"

End Sub
Private Sub Matching_Columns_Paper16()
strcolumnA(1) = "Priority queue"
strcolumnA(2) = "AVL tree"
strcolumnA(3) = "Adjacency list"
strcolumnA(4) = "Nested procedure"
strcolumnA(5) = "Garbage collection"
strcolumnA(6) = "Postfix"
strcolumnA(7) = "Sequential memoty"
strcolumnA(8) = "    Hashing"
strcolumnA(9) = "B-tee"
strcolumnA(10) = "Circular queue"

strcolumnB(1) = "Height balanced"
strcolumnB(2) = "Graph"
strcolumnB(3) = "Process scheduling"
strcolumnB(4) = "Reverse polish"
strcolumnB(5) = "Stack"
strcolumnB(6) = "Dynamic memory management"
strcolumnB(7) = "Indexing technique"
strcolumnB(8) = "Disk"
strcolumnB(9) = "Most efficient queue using arrays"
strcolumnB(10) = "Tree indexing"

stranswers(1) = "Process scheduling"
stranswers(2) = "Height balanced"
stranswers(3) = "Graph"
stranswers(4) = "Stack"
stranswers(5) = "Dynamic memory management"
stranswers(6) = "Reverse polish"
stranswers(7) = "Disk"
stranswers(8) = "Indexing technique"
stranswers(9) = "Tree indexing"
stranswers(10) = "Most efficient queue using arrays"

End Sub
Private Sub Matching_Columns_Paper17()
strcolumnA(1) = "String"
strcolumnA(2) = "Btree"
strcolumnA(3) = "Symbol table"
strcolumnA(4) = "FIFO"
strcolumnA(5) = "Sapanning tree"
strcolumnA(6) = "File structure"
strcolumnA(7) = "Structure"
strcolumnA(8) = "Boundary tag"
strcolumnA(9) = "Dequee"
strcolumnA(10) = "Rank"

strcolumnB(1) = "Hashing"
strcolumnB(2) = "Record"
strcolumnB(3) = "Parsing"
strcolumnB(4) = "Secondary storage"
strcolumnB(5) = "Garbage collection"
strcolumnB(6) = "Input restricted"
strcolumnB(7) = "Matrix"
strcolumnB(8) = "Scheduling"
strcolumnB(9) = "Height balanced"
strcolumnB(10) = "Breadth first search"

stranswers(1) = "Hashing"
stranswers(2) = "Height balanced"
stranswers(3) = "Parsing"
stranswers(4) = "Scheduling"
stranswers(5) = "Breadth first search"
stranswers(6) = "Secondary storage"
stranswers(7) = "Record"
stranswers(8) = "Garbage collection"
stranswers(9) = "Input restricted"
stranswers(10) = "Matrix"

End Sub
Private Sub Matching_Columns_Paper18()
strcolumnA(1) = "Leaf"
strcolumnA(2) = "Linear probing"
strcolumnA(3) = "Non linear data structure"
strcolumnA(4) = "LIFO"
strcolumnA(5) = "Negative numbers"
strcolumnA(6) = "Folding"
strcolumnA(7) = "Di-graph"
strcolumnA(8) = "Threaded"
strcolumnA(9) = "Game tree"
strcolumnA(10) = "External sort"

strcolumnB(1) = "Tree"
strcolumnB(2) = "Stack"
strcolumnB(3) = "Hashing"
strcolumnB(4) = "Binary tree"
strcolumnB(5) = "Directed edge"
strcolumnB(6) = "Zero sum"
strcolumnB(7) = "Key generation"
strcolumnB(8) = "Collision management"
strcolumnB(9) = "Merging"
strcolumnB(10) = "2’s complement"

stranswers(1) = "Linear probing"
stranswers(2) = "Collision management"
stranswers(3) = "Tree"
stranswers(4) = "Stack"
stranswers(5) = "2’s complement"
stranswers(6) = "Key generation"
stranswers(7) = "Directed edge"
stranswers(8) = "Binary tree"
stranswers(9) = "Zero sum"
stranswers(10) = "Merging"

End Sub
Private Sub Matching_Columns_Paper19()
strcolumnA(1) = "Quick sort"
strcolumnA(2) = "ISAM"
strcolumnA(3) = "Dope ector"
strcolumnA(4) = "Random access"
strcolumnA(5) = "Topological ordering"
strcolumnA(6) = "Best fit"
strcolumnA(7) = "Forest"
strcolumnA(8) = "n-queen"
strcolumnA(9) = "List in a list"
strcolumnA(10) = "Dynamic allocation"

strcolumnB(1) = "Indexing"
strcolumnB(2) = "Partitioning and exchanging"
strcolumnB(3) = "Array"
strcolumnB(4) = "Memory allocation"
strcolumnB(5) = "Sparse matrix"
strcolumnB(6) = "Graph"
strcolumnB(7) = "Pointers in C"
strcolumnB(8) = "Binary tree based representation"
strcolumnB(9) = "Backtracking"
strcolumnB(10) = "Generalized list"

stranswers(1) = "Partitioning and exchanging"
stranswers(2) = "Indexing"
stranswers(3) = "Sparse matrix"
stranswers(4) = "Array"
stranswers(5) = "Graph"
stranswers(6) = "Memory allocation"
stranswers(7) = "Binary tree based representation"
stranswers(8) = "Backtracking"
stranswers(9) = "Generalized list"
stranswers(10) = "Pointers in C"

End Sub

Private Sub questions_paper1()
StrQuestions(1) = "Wtie PUSH and POP functions ins C language simulating push and pop operations of stack which are implemented by array of integers."
StrQuestions(2) = "Write INSERt and DELETE functions in C language simulating insertion and deetion in queue which is implemented by array of characters."
StrQuestions(3) = "Write INSERT and DELETE functions in C language simulating inserting and deletion in circular queue which is implemented by array if characters."
StrQuestions(4) = "Write a C program to sort the names stored in a 2 dimensional array."
StrQuestions(5) = "Write a C program to create a singly linked list of records sorted in ascending order. Key record is of type integer."
StrQuestions(6) = "Write a C program to delete a node containing a gien information in singly linked list. Make necessary assumptions."
StrQuestions(7) = "Write a C program to print the information from each node insingly linked list. Make necessaty assumptions."
StrQuestions(8) = "Write a C program to create a doubly linked list in asceding sorted order of the information. Information firld is of type integer."
StrQuestions(9) = "Write a C program to delete a node containing a gien information in doubly linked list. Make necessary assumptions."
StrQuestions(10) = "Write a C program to print the information, from each node in reverse order in doubly linked list, in which pointer to last node is TAIL."
End Sub

Private Sub questions_paper2()
StrQuestions(1) = "Write PUSH and POP function, in C language simulating push and pop operations of a stack which is implemented by singly linked list., Make necessary assumptions."
StrQuestions(2) = "Write INSERt and DELETE functions in C language simulating insertion and deletion operations of a queue which implemented bu singly linked list. Make necessaty assumptions."
StrQuestions(3) = "    Write INSERT and DELETE functions, in C language, simulating nsertion and deletion operations in circular queue implemented by singly linked list. Make necessaty assumptions."
StrQuestions(4) = "a.  Write a C program to delete a node, containing a gien information from queue, which is implemented by singly linked list." & Chr(13) & _
                  "b.  Write a C program to delete a node, containing a gien information from circular queue which is implemented by singly linked list."
StrQuestions(5) = "Explain the process of obtaining a binary tree which represents a given ordered tree. Give one example."
StrQuestions(6) = "Write a general non recursie C program, for postorder traversal of a binary tree. Explain it giving one example."
StrQuestions(7) = "Write and explain a general non recursive C program for a preorder traversal of a binary tree. Explain it giving one example."
StrQuestions(8) = "Explain the method of representing any ordered forest by a unique binary trees."
StrQuestions(9) = "Explain the process of threading of a binary tree for inorder traversal with at least two examples."
StrQuestions(10) = "Discuss the advantages and disadvantages of a threaded stoage representation for binary trees."
End Sub

Private Sub questions_paper3()
StrQuestions(1) = "Explain how a lexical ordered binary tree can be created. Take one insertion sequence of names and expain the procedure of cration of a lexical ordered binary tree through repeated insertions."
StrQuestions(2) = "Describe different methods of implementing a symbol table. Copare the efficiencies of the insertion and searching mechanisms offered by these mechanisms."
StrQuestions(3) = "Describe traversal of a threaded binary tree from its unthreaded counterpart giving any one example."
StrQuestions(4) = "What is hashing? List important features of a good hashing algorithm. Explain how it is useful to implement symbol tables."
StrQuestions(5) = "Write a C program to insert string INSTR to MASTERSTR at position POS. Use data structure of your choice, suitable to the problem."
StrQuestions(6) = "Write a C program for concatenating two strings, STR1 and STR2 to get new string STR3. Strings are to be stored using fixed length metod. What are the drawbacks.?"
StrQuestions(7) = "Write a C program for garbage collection, i.e. to reclaim usale workspace. This proglem arises due to index table meth of string handling."
StrQuestions(8) = "Write a recursive C progam for quick sort. Trace your algorithm for sample data. Analyze your algorithm."
StrQuestions(9) = "Write a recursive C program for merge sort. Trace your program for samle data. Analyze your algorighm."
StrQuestions(10) = "Write a C program for insertion/.selection/bubble sort and analyze them."
End Sub
Private Sub questions_paper4()
StrQuestions(1) = "Write a C program for division remainder techniques of hashing. Also handle colliwion by using linked colliwion processing."
StrQuestions(2) = "Define a B-tree. Taking an example to show how insertion is handled here."
StrQuestions(3) = "Write a C program for depth first search/Breadth first search of a graph."
StrQuestions(4) = "Write in short on applications of graphs in computer science."
StrQuestions(5) = "Write a simple pattern matching algorithm, which compares a given pattern P with each of the substring of T, moving from left to right, until you get a match."
StrQuestions(6) = "Explain binary search method. Write C function for recursive and non recursive implementation of binary search. Compare these imolementations in terms of time and memory requirements."
StrQuestions(7) = "Explain the quicksort method. Discusss how quick sort behaves when the array to be sorted contains the same key."
StrQuestions(8) = "Write a c function to ealate the postfix expression P using stack P: 5 6 2 + * 12 4 / -" & Chr(13) & _
                  "Explain how priority of the operators are taken care of futing the process of converting an infix expression to postfix form."
StrQuestions(9) = "A lending library keeps a list of issued books. Each item in the issue list has the following information:" & Chr(13) & _
                    "a.  Accession number of the book, which is an integer;" & Chr(13) & _
                    "b.  Issue code, which is I, B or L, indicating issued to a member, sent for binding or issued to another library respectively; and" & Chr(13) & _
                    "c.  Identity of member to whom issued, which is an alpha numeric quantity" & Chr(13) & _
                    "The services provided b yteh library are:" & Chr(13) & _
                    "i.  Provides status of a book on request from a member" & Chr(13) & _
                    "ii. When a book is issued, it is added to the listof issued books;" & Chr(13) & _
                    "iii.    When a book is returned it is deleted from the list of issued books." & Chr(13) & _
                    "Design efficient data structures for storing the above information." & Chr(13) & _
                    "Design C procedures for implementation of the services."
StrQuestions(10) = "What is B-tree? Discuss an algorithm for deletion of a node from B-tree."
End Sub
Private Sub questions_paper5()
StrQuestions(1) = "In n-dimensional space, v1 and v2 are the end points of a line. 80% of the components of v1 and v2 are zero. Write a C function for computing length of the line v12 using appropriate data structure representation for the sparse vectors v1 and v2. Justify your choice of represenatation."
StrQuestions(2) = "Write a C program to deermine if an input string is of the from X x Y, where X is a string of arbitrary length composed only of 1 and 2. Y is the reerse stirng of X and x is any alphabet. Eg. 122b221 and 12121a12121 are valid strings."
StrQuestions(3) = "Write a C program to read ina parenthesized infix expression and " & Chr(13) & _
                  "a.  Chech whether partentheses match" & Chr(13) & _
                  "b.  Remove all extra parentheses keeping the value of the expression intact."
StrQuestions(4) = "A stack contains a set of 10 names. It is required to assign the names, from the stack, to variables X,Y,Z I in the following order" & Chr(13) & _
                    "X = 3rd name in the stack from bottom" & Chr(13) & _
                    "Y = 3rd name in the list from top" & Chr(13) & _
                    "Z = 5th name in the stack" & Chr(13) & _
                    "Write a C program to perform these operations without changing content or configuration of the stack. You cannot use random access for accessing th4e elements in the stack."
StrQuestions(5) = "Write a C program to implement a CHANGE editor command that isised in the form CHANGE(string1,string2,string3) to replace the fitst occurrence of string 2(if any) in string 1 by string 3."
StrQuestions(6) = "What is garbage collection? Explain the boundary tag method."
StrQuestions(7) = "Give the heap that results when starting with empty heap nodes are inserted successively for the keys" & Chr(13) & _
                    "EASY QUESTION" & Chr(13) & _
                    "Assume aplhanemrric ordering of the keys"
StrQuestions(8) = "Which positions could be occupied by the 3rd largest key in a heap of size 32? Which positions could not be occupied by the 3rd smallest key in a heap of size 32?"
StrQuestions(9) = "Write a C program for deletion of a record from the position d in a heap."
StrQuestions(10) = "Write a sorting program which implements a recursive merge sort with a cutoff to insertion sort for subsets with less than M elements. Discuss about possible values of M which would make your program most efficient."
End Sub
Private Sub questions_paper6()
StrQuestions(1) = "Write a recursive program to compute height of a binary tree: longest distance from the root to a leaf node. Prove that your algorithm is correct."
StrQuestions(2) = "Draw the binary search tree that results from inserting into an initially empty tree records with the keys" & Chr(13) & _
                    "EASY QUESTION" & Chr(13) & _
                    "and then deleting the Q."""
StrQuestions(3) = "a.  Suppose that the number of items to be put into hash table is known in advance. Under what conditions will chaining be preferred to double hashing?" & Chr(13) & _
                  "b.  Which hashing method would you use for an application in which many equal keys are likely to be present?"
StrQuestions(4) = "a.  Gie the ‘next’ table for Knuth-Morris-Pratt algorithm for the pattern ABRADACABRA." & Chr(13) & _
                  "b.  Suggest an algorithm for searching a text file for a string of 50 consecutive blanks."
StrQuestions(5) = "a.  Suppose that it is known that the running time of one algorithm is always about NlogN and that the running time ofanother algorithm is always about N(3). What can you say about relative performance of the algorithms?" & Chr(13) & _
                  " b.  Explain the difference beween O(1) and O(2)."
StrQuestions(6) = "a.  Give an example of a tree for which the stack in a preorder traversal uses more space than the queue in a breadth first traversal of a tree." & Chr(13) & _
                  "b.  Draw the parse tree for the expression (A + B) * C + (D+E)"
StrQuestions(7) = "Write a C program for level order traversal (breadth first) of a forest represented as a binary tree."
StrQuestions(8) = "Write an recursie “divide and conquer” C rogram to draw an approximation to the line segment connecting two points (x1,y1) and (x2,y2) by drawing points using only integer coordinates. You should start from a point close to the middle."
StrQuestions(9) = "Consider the following problem: we imagine N people have decided to commit suicide by arranging themselves in a circle and killng the Mth person around the circle, closing ranks as each person drops out of the circle. Problem is to find out which person is the last to die. Write a C program to simulate execution sequence to solve the problem."
StrQuestions(10) = "What is acircular queue? Write a C program to insert an item in the circular queue. Write another C function for printing elements of the queue in reverse order."
End Sub
Private Sub questions_paper7()
StrQuestions(1) = "Write a C function for inserting an item into a hsh table with open addressing and linear probing."
StrQuestions(2) = "A tri diagonal matrix is a square matrix in which all entiries are 0 except possibly thos on the main diagonal and on the diagonals immediately above and below it. That is, T is a tri diagonal matrix means that T[i][j]  = 0 unless | i-j| is less than or equal to 1." & Chr(13) & _
                  "Deice a space efficient storage scheme forsuch matrices. And give the corresponding idex function." & Chr(13) & _
                  "Write a C function for obtaining transpose of the matrix using the storage scheme devised by you."
StrQuestions(3) = "Draw three leels, showing the firsttw omoves of the game tree for the game of tic tac toe. Discuss how symmetry can be usedfor reducing the number of nodes at ech level of the game tree."
StrQuestions(4) = "Write a recursive C program for solving the eight queen’s problem." & Chr(13) & _
                  "What is the maximum depth of recursion in the eight queen’s problem."
StrQuestions(5) = "Write C functions required for deletion of a node from an AVL tree."
StrQuestions(6) = "Consider an array based implementation of a heap. Design a C function that will delete an item with index p from a heap and restore heap properties of the resulting list. Estimate time and space requirements of your function."
StrQuestions(7) = "Insert the keys below into an initially empty B-tree of order 3" & Chr(13) & _
                  "a g f b k d h m j e s I r x l n t u p" & Chr(13) & _
                  "Show B-tree at each step."
StrQuestions(8) = "If a key in a B – tree is not in a leaf, prove that its immediate predecessor and immediate successor( under the natural order) are in leaves."
StrQuestions(9) = "What is the smallest number of keys that ,when inserted in an appropriate order, will force a B- tree of order 5 to have height 2 ?"
StrQuestions(10) = "A dequeue is a list in which additions and deletions can be made at either the head or tail. With dequeue stored as circularly linked list, provide algorithms to add a node  to either end of a dequee and to delete a node from either end."
End Sub
Private Sub questions_paper8()
StrQuestions(1) = "Write a function that will concatenate two linked lists. Function should have two paralmeters – pointers to the beginning of the lists and the function should link second list at the end of the first list." & Chr(13) & _
                  "Write a similar function that splits a list in two. Function should have as its parameters pointer to the list and pointer to the node at which split takes place"
StrQuestions(2) = "A soubly lnked list can be made circular by adjusting appropriate pointers. Suggest the pointer adjustment." & Chr(13) & _
                  "Discuss the advantages and disadvantages of a circularly doubly linked list in doing varioys list operations."
StrQuestions(3) = "    Write a C function to delete a node p from a singly linked list using a single pointer variable."
StrQuestions(4) = "A binary tree is a Fibonacci tree if it is balanced and has the minimum number of nodes among all balanced binary trees with its depth." & Chr(13) & _
                  "Develop an algorithm to generate a Fibonacci tree of depth d."
StrQuestions(5) = "Can you find a way to use the left and right pointer fields of terminal and internal nodes that contain null pointers in a binary tree to create a threaded binary trees so that they can be traversed without using stack. Provide a C routine for traversal using your scheme."
StrQuestions(6) = "Suppose that you are given a maze with the structure of a binary tree. The maze has a strating point and an exit. A node associated with the starting point, with the exit and with eery point in the maze where one of two paths must be taken. Modify preorder traversal scheme to design an algorithm for printing out a path from starting node to the exit point."
StrQuestions(7) = "Design a procedure BUFFER in c that has the following three parameters – pointer to the beginning of queue Q, E of type element of the queue and OP, a character. If OP is R, then the procedure returns after removing top of the queue and copying it  E. If OP is n, then BUFFER returns after inserting E at the end of queue."
StrQuestions(8) = "Discuss the advantages and disadvantages of designing a general sorting function as a hybrid betwwn quicksort and shell sort. What criteria qould you use to switch from one to another?"
StrQuestions(9) = "Explain the selection sort method." & Chr(13) & _
                  "What is the advantage of selection sort over all other methods of sorting?"
StrQuestions(10) = "Is it true that if we represent polynomials as circularly linkes lists then we can erase a polynomial more easily? Discuss." & Chr(13) & _
                   "Provide C functions for addition and subtraction function for these polynomials."
End Sub
Private Sub questions_paper9()
StrQuestions(1) = "What are sparse matrices?" & Chr(13) & _
                  "Develop data structures and appropriate procedures to adapt the generalized dope vector method to a sparse 3 dimensional array."
StrQuestions(2) = "Write a C program to store the roll number and names of the students in a binary search tree" & Chr(13) & _
                  "Write a C function which searches for the roll number of a student in the binary tree."
StrQuestions(3) = "Proide a C function for implementation of multiplication of two long integers. Suggest appropriate data structure for storage of the two integers."
StrQuestions(4) = "Write a C function that sorts a file by first applying the radix sort to the most significant r digits. If then uses insertion sort to sort the entire field."
StrQuestions(5) = "Provide a C program for Kruskal’s minimum spanning tree algorithm. Discuss time and space complexity of the algorithm."
End Sub

Private Sub TrueorFalse_Paper1()
StrQuestions(1) = "In 1’s complement method the positive ‘0’ and negative ‘0’ is separately represented."
StrQuestions(2) = "The binary number system is the only means by which bits can be used to represent integers."
StrQuestions(3) = "The limiting factor on the precision of numbers that can be represented on a particular machine, is the number of significant digits in the mantissa."
StrQuestions(4) = "The method of interpreting a bit pattern is called a data type."
StrQuestions(5) = "A data type is a collection of values and a set of operations on values."
StrQuestions(6) = "Patching an old program is easier than starting afresh."
StrQuestions(7) = "A queue can be defined as a waiting line."
StrQuestions(8) = "Recursive function must be clled directly."
StrQuestions(9) = "In many cases we observe that, a recursie solution isn them ost natural and logical way of solving problem."
StrQuestions(10) = "The problem of Tower of Hanoi is very difficult to solve using non recousive process."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper2()
StrQuestions(1) = "The number of total elements in a queue is fixed."
StrQuestions(2) = "The number of elements is the queue, at any time, is  -(q.near – q.front)."
StrQuestions(3) = "The operation of removing and element, from a queue logically, involves manipulation of only one element."
StrQuestions(4) = "The extra node kept in front of a list is known as list header."
StrQuestions(5) = "Dynamic implementation of list may be time consuming to call upon the system to allocate and free storage."
StrQuestions(6) = "In dynamic implementation of list, the address computation is complex."
StrQuestions(7) = "The term head of the queue is same as the term front of the queue."
StrQuestions(8) = "Queue can be created by setting up an ordinary contiguous array to hold the items."
StrQuestions(9) = "The concept ofcucular array is very difficult to implement."
StrQuestions(10) = "Data structures are programming language dependent."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper3()
StrQuestions(1) = "The major activities on ordered list are accesing, inserting and deleting."
StrQuestions(2) = "In top down approach detailed consideration is postponed."
StrQuestions(3) = "Data structure as defined at logical level, is called abstract data type."
StrQuestions(4) = "Top down approach uses divide and conquer approach."
StrQuestions(5) = "O(N) (linear time) is better than O(1) (constant time)"
StrQuestions(6) = "Array is linear data structure."
StrQuestions(7) = "Address ofan element is invariant whether two dimensional array is implemented by using row-major order or column major order."
StrQuestions(8) = "List null can be used to initialize list ass empty list."
StrQuestions(9) = "A stack may be representated by a linear linked list."
StrQuestions(10) = "For stack or queue, arbitrary amount of memory can be allocated."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper4()
StrQuestions(1) = "In linked list the logical order of elements, is same as their physical arrangement."
StrQuestions(2) = "Random access of the elements in linked list is possible."
StrQuestions(3) = "In linked list each node must have at least two fields."
StrQuestions(4) = "In linked list , successive elements need not occupy adjacent space in memory."
StrQuestions(5) = "Pointer is used to provide the lnear order in linked list."
StrQuestions(6) = "Null pointer variable contains the address of last node, in linked list."
StrQuestions(7) = "List pointer variable coantains the address of last node, in linked list."
StrQuestions(8) = "Deletion operation, in non empty linked list, may result in overflow condition."
StrQuestions(9) = "Attempting to create a new node when free space pool has no space, results in overflow condition."
StrQuestions(10) = "Traversal, in linear array and in linked list, are similar."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper5()
StrQuestions(1) = "Searching the linked list requires the list must be in sorted order."
StrQuestions(2) = "Inserting a node in unsorted linked list in the middle of the list, is always easier."
StrQuestions(3) = "In copy all or part of the linked list, each element is inserted in null list one by one."
StrQuestions(4) = "Deletion of node in linked list involves only keeping track of each of the addresses of he node that is to be deleted."
StrQuestions(5) = "Header linked list always contains a special node, called header node , at the beginning of the list."
StrQuestions(6) = "In grounded header list last node points back to the header node."
StrQuestions(7) = "In circular header list, last node points back to the header node."
StrQuestions(8) = "Polynomials in memory can be represented by header linked list."
StrQuestions(9) = "Deleting the node whose address is known, in double linked list, requires at least some part of the list must be traversed."
StrQuestions(10) = "Two field node of linked list is sufficient to represent polynomial in memory."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper6()
StrQuestions(1) = "Push operation in stack is performed at he rear end."
StrQuestions(2) = "Only the most recently inserted element can be accessed through POP operation in stack."
StrQuestions(3) = "FRONT = REAR pointer means queue is empty."
StrQuestions(4) = "Deletion, in double linked, requires more pointer changes than in singly linked list."
StrQuestions(5) = "Garbage collection is a technique to collect deleted space into free storage list."
StrQuestions(6) = "Dynamic memory allocation takes place at run time."
StrQuestions(7) = "Static memory allocation takes place at run time."
StrQuestions(8) = "Free table of memory is sorted on size, in first fit method."
StrQuestions(9) = "Free table of memory is sorted on size of space, in best fit method."
StrQuestions(10) = "The first fit method claims reductin in wastage of space due to fragmentation."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper7()
StrQuestions(1) = "Union is most suitable data structure to represent in a linked list."
StrQuestions(2) = "Pointer is used to represent link of linked list."
StrQuestions(3) = "A singly linked list can be constructed in sorted order."
StrQuestions(4) = "Initialized static pointer ariable can be substituted bu initialized global variable."
StrQuestions(5) = "Doubly linked list can be scanned in either direction."
StrQuestions(6) = "POP operation in stack may result in overflow."
StrQuestions(7) = "PUSH operation in stack may result in underflow."
StrQuestions(8) = "malloc() allocates memory statically."
StrQuestions(9) = "It is possible to return previoysly allocated memory to system."
StrQuestions(10) = "calloc() function and malloc() return pointer to void."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub
Private Sub TrueorFalse_Paper8()
StrQuestions(1) = "Every tree can be uniquely represented by a binary tree."
StrQuestions(2) = "Tree traversal is a procedure by which each node in the tree is processed exactly twice, in a systematic manner."
StrQuestions(3) = "In preorder traversal of a binary tree, the left subtree is traversed after processing the root node."
StrQuestions(4) = "In postorder traversal, of a binary tree, the root node is processed first."
StrQuestions(5) = "For inorder traversal, of a binary tree, the inorder traversal of its left subtree should be performed first."
StrQuestions(6) = "A binary tree is threaded according to a particular traversal order."
StrQuestions(7) = "For inorder threading of a binary tree, if the left link of some node is NULL, it is replaced by the address of its successor."
StrQuestions(8) = "While threading of a binary tree, a NULL right link is replaced by the address of the successor of the node in question."
StrQuestions(9) = "Each node in a balanced binaty tree has a balance of 1,-1 or 0."
StrQuestions(10) = "A balanced binaty tree( AVL tree) is a binary tree in which the heights of the two subtrees ,of every node, are always equal."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper9()
StrQuestions(1) = "In a threaded binary tree, the left link of a node always denots a structural link."
StrQuestions(2) = "The balance of a node in a bonary tree, is defined as the hgiht of its left subtree, minus the height of its right subtree."
StrQuestions(3) = "In a threaded binary tree the left orr right link of a node is always a thread link."
StrQuestions(4) = "The height of a binary tree is the maximum level of its leaves."
StrQuestions(5) = "The process of determining predecessor and successor nodes, for any given node, is easy for unthreaded tree compared with its threaded counterpart."
StrQuestions(6) = "The inorder traversal, of a threaded tree, is faster than that of its unthreaded version."
StrQuestions(7) = "Insertions into and deletions from, a threaded tree are faster than that of its unthreaded version."
StrQuestions(8) = "Insertions, in an existing lexical ordered binary tree, can only be performed at a leaf level."
StrQuestions(9) = "Searching of a node in a lexical ordered binary tree, is slow because the number of comparisons required greatly increases."
StrQuestions(10) = "Symbol table is an important part of operating system, keeping record of names and their associated values."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "Ture"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper10()
StrQuestions(1) = "If a binary search method is used for accessing a symbol table, the ration of insertion time to access time, for the table entries, will be quite high."
StrQuestions(2) = "A B-tree, of order 11, contains at least 4 keys in each non root node."
StrQuestions(3) = "Each non root node of a B-tree of order n contains at least n/2 keys."
StrQuestions(4) = "The difference eween heights of the left and right subtrees of each node, in a balanced binary tree is 1, -1 or 0."
StrQuestions(5) = "For any node, of an AVL tree, the height of its left subtree must be less than the height of its right subtree must be less than the height of its right subtree."
StrQuestions(6) = " In a binary search tree, each node contains a single key and points to three subtrees."
StrQuestions(7) = "If a node ofmultiway search tree contains 3 keys, it should have 4 subtrees."
StrQuestions(8) = "A 4-5 tree is a B-tree of order 5, as is a 5-4 tree."
StrQuestions(9) = "In a tree, a node with no subtrees is called a leaf node."
StrQuestions(10) = "A binary tree consists of nodes which have exactly 2 offsprings."

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
StrQuestions(1) = "The leel of any node of a tree is the length of its path from the root."
StrQuestions(2) = "A B tree of order n is also called an (n-1) – n tree."
StrQuestions(3) = "Hashing refers to the process of deriving astorage address from a record key."
StrQuestions(4) = "A threaded binary tree canhave NULL links but no structural links."
StrQuestions(5) = "The total number of cucles contained in any tree is 0."
StrQuestions(6) = "The total number of cycles contained in any tree is 0."
StrQuestions(7) = "If , in given directed tree, the othdegree of every node is less than of equal to 2, it is called a complete binary tree."
StrQuestions(8) = "Deletion of a node from threaded tree is time consuming but insertion ito it is not."
StrQuestions(9) = "In an m-ary tree the outdegree of all nodes is less than or equal to m."
StrQuestions(10) = "A B-tree of order 13, contains at least 12 keys in each non root node."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper12()
StrQuestions(1) = "Adjacency matrix, of a graph, is a n * e matrix, where n is no. of vertices and e is no. of edges."
StrQuestions(2) = "Adjacency matrix of a graph is a symmetric matrix."
StrQuestions(3) = "There are only two 1’s in a row or column of an adjacency matrix of a graph."
StrQuestions(4) = "If x is adjacency matrix of a graph, then I,jth entry of xk, gives no. of paths of length k between ith & jth vertex."
StrQuestions(5) = "    In depth first search all edges incident on a vertex are scanned before moing to other vertex."
StrQuestions(6) = "In depth first search move to adjacent unvisited vertex is made as soon as possible."
StrQuestions(7) = "Other name for breadth first search is back tracking."
StrQuestions(8) = "Fixed length method of string representation results in wastage of space."
StrQuestions(9) = "It is not advisable to use linked list representation of strings if there are lots of insertion and deletion."
StrQuestions(10) = "Adjacency matrix is a bit matrix."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper13()
StrQuestions(1) = "Linked list and index table method of string representation."
StrQuestions(2) = "After insert operation in linked list representation of strings, inserted string may be accessed."
StrQuestions(3) = "In internal sorting methods it is necessaty for all data to reside in internal memory."
StrQuestions(4) = "Insertion sort is also called bubble sort."
StrQuestions(5) = "While sorting on multiple keys the first key is least significant."
StrQuestions(6) = "In merge sort files to be merged must be sorted."
StrQuestions(7) = "Quick sort and merge sort are example of divide and conquer algorithms."
StrQuestions(8) = "Insertion sort is o(n)."
StrQuestions(9) = "Worst case efficiency of binary search is n log n+1"
StrQuestions(10) = "In binary search keys must be ordered."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper14()
StrQuestions(1) = "In binary search all keys may not reside in main memory."
StrQuestions(2) = "Hashing is also known as KAT."
StrQuestions(3) = "Division remainder method of hashing uses a prime number to arrive at address."
StrQuestions(4) = "Folding method of Hashing, rules out possibility of collision."
StrQuestions(5) = "There is no hashing method without the possibility of collision."
StrQuestions(6) = "Multiple hashing uses more than one hashing function to resolve collision."
StrQuestions(7) = "In linked collision processing, initial record is in prime area and all other records, which produce collision are put in overflow area."
StrQuestions(8) = "ISAM stands for index storage access method."
StrQuestions(9) = "Use of ISAM is advisable for files with lots of insertions and deletions."
StrQuestions(10) = "B+ tree is a binary tree."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper15()
StrQuestions(1) = "Hashing algorithm cannot be used with non numerical keys."
StrQuestions(2) = "Pattern matching refers to finding whether two strings are identical."
StrQuestions(3) = "Adjacency list representation is useful when number of edges is small."
StrQuestions(4) = "Breadth first search uses stack, to hold nods that are waiting to be processed."
StrQuestions(5) = "Depth first search uses queue, to hold nodes that are waiting to be processed."
StrQuestions(6) = "SNOBOL was developed to solve problems involving string manipulation."
StrQuestions(7) = "Adjacency list representation is suitable when, number of vertices of graph changes due to insertion and deletion of vertices."
StrQuestions(8) = "Adjacency list representation is suitable for breadth first search but not for depth first search."
StrQuestions(9) = "Average case time for Quick sort is O(logn)."
StrQuestions(10) = "In linear collision processing, the key causing collision is kept at first vacant position."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper16()
StrQuestions(1) = "Stack is an example of linear list."
StrQuestions(2) = "Without pointe data type it is not possibleto implement hierarchical data structures."
StrQuestions(3) = "A queue can be implemented using acircular array with front and rear indices and one position left vacant."
StrQuestions(4) = "It is easy to nsert an element in alinear list implemented using an array."
StrQuestions(5) = "Major drawback of linked llists is that they are not suitable for random access."
StrQuestions(6) = "A doubly linked list can be made circular buy appropriately setting the values only in the first node of the list."
StrQuestions(7) = "A polynomial in two variables can be represented as a linked list of structures."
StrQuestions(8) = "In order toperform arithmetic with large integers(integers with values which exceed the limit in a computer) we need to use linear lists."
StrQuestions(9) = "Bubble sort algorithm requires O(N) comparisons for sorting a sorted list."
StrQuestions(10) = "Quick sort algorithm requires O(NlogN) cmarisons for sorting a sorted list."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper17()
StrQuestions(1) = "Sequential search for small n (number of elements in a list) can be faster than binary searh."
StrQuestions(2) = "If a<b, then a(n) is O(b(n)), but b(n) is not O(a(n))."
StrQuestions(3) = "If f(n) is O(g)(n))  and h(n) is any non ero function then the function f(n)h(n) is not O(g(n)h(n))."
StrQuestions(4) = "If we do not know that a given list is ordered, then we need to adopt sequential search strategy."
StrQuestions(5) = "Recursive functions and iterative functions, using stacks, can accomplish exactly the same tasks."
StrQuestions(6) = "Tail recursion can be replaced by efficiently constructed loop."
StrQuestions(7) = "A binary tree always has a root."
StrQuestions(8) = "A heap need not be a binary tree."
StrQuestions(9) = "If the nodes in a complete binary tree are labeled in inorder sequence, then each node is exactly as many levels above the leaves as the highest power of wo that diides this label."
StrQuestions(10) = "In an AVL tree left and right subtree will always hae the same height."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper18()
StrQuestions(1) = "Merge sort is a stable sorting alogirithm."
StrQuestions(2) = "Linear probing is a collision resolution policy for hashing."
StrQuestions(3) = "For sorting elements ofan array where each element is a large record, it is always desirable to use an index array for reducing overhead of comparison."
StrQuestions(4) = "Radix search algorithms work with bits or components of the keys instead of the complete keys."
StrQuestions(5) = "An articulation point in a connected graph is a vertx hat, if deleted, would break the graph nto two or more components."
StrQuestions(6) = "Breadth first graph search alogorithm uses a stack to keep track of nodes visited."
StrQuestions(7) = "Graphs can be used forprocessig sets i.e. collection of objects."
StrQuestions(8) = "We cn find a spanning tree of a graph using depth first search algorithm."
StrQuestions(9) = "A bit array in C can be used forstoring adjacency matrix of an undirected grph but not a directed graph."
StrQuestions(10) = "Topolgical sorting can be used for schedultng tasks."

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

Private Sub TrueorFalse_Paper19()
StrQuestions(1) = "Queue is an useful dat structure for any simulatin application."
StrQuestions(2) = "A priority queue is implemented using an array of stacks."
StrQuestions(3) = "A pre order and post order traversal sequence uniquely defines a tree."
StrQuestions(4) = "An n-ary tree can be represented using binary tree."
StrQuestions(5) = "First fit technique is used in memory management."
StrQuestions(6) = "Adjacency matrix of a graph of n nodes with each node only connected to itself is an identity matrix."
StrQuestions(7) = "In C strings are represented using ‘\s’ terminated arrays."
StrQuestions(8) = "Representation of a priority queue as a heap is advantageous for small n(number of elements)."
StrQuestions(9) = "Binary search trees can be used for implementation of symbol tables because it requires only O(N) comparisons for each access."
StrQuestions(10) = "A B-tree is an m-way balanced tree."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"

End Sub

Private Sub TrueorFalse_Paper20()
StrQuestions(1) = "Threaded binary trees are useful for implementation of non recursive tre traversal alogirithms."
StrQuestions(2) = "Selection sort inimizes data movement and hence useful for array based implementations involving large structures."
StrQuestions(3) = "A field of a structre in C cannot be of union type."
StrQuestions(4) = "Pointers in C provides for address arithmetic independent of the underlying hardware."
StrQuestions(5) = "An array based implementation of binary tree endures optima usage of space."
StrQuestions(6) = "An array based implementation of complete binary tree ensures optimal usage of space."
StrQuestions(7) = "If we want to delete leaves of a binary tree we should use inorder traversal algorithm."
StrQuestions(8) = "A complete graph is one in which every node is connected to every other node."
StrQuestions(9) = "Binary trees ae usedfor evaluation of arithmetic expressions."
StrQuestions(10) = "Preorder traversal of a binary search tree produces a sorted sequence of records."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub

Private Sub multiplechoice_Paper1()
strmultiple(1) = "The most widely used method for interpreting bit setting as non negative integer is –"
strmuloption(1) = "Binary number system"
strmuloption(2) = "BCD system"
strmuloption(3) = "ASCII system"
strmuloption(4) = "None of the above"

strmultiple(2) = "Any string of bits of length n represents a unique non negative integer between –"
strmuloption(5) = "0 and 2n-1 – 1"
strmuloption(6) = "0 and 2n-1"
strmuloption(7) = "1 and 2n-1"
strmuloption(8) = "none of the above"

strmultiple(3) = "With 1’s complement method, the range of numbers that cn be represented using n bits is –"
strmuloption(9) = "A 1 followed by using n-1 zeros to a 0 followed by n-1 ones"
strmuloption(10) = "A 0 followed by n -1 ones to a 1 followed by n – 1 zeros "
strmuloption(11) = "-2n-1 to 2n-1"
strmuloption(12) = " none of the above"

strmultiple(4) = "Using 2’s complement method, the range of numbers that can be represented using n bits is –"
strmuloption(13) = "A 1 followed by n ones to a 1 followed by n zeros"
strmuloption(14) = "A 1 followed by n – 1 zeros to a 0 followed by n-1 1’s"
strmuloption(15) = "-2n-1 to 2n-1"
strmuloption(16) = "none of the above"

strmultiple(5) = "Key concept for representing a real number is am mantissa times –"
strmuloption(17) = "10 raised to an exponent"
strmuloption(18) = "2 raised to an exponent"
strmuloption(19) = "base raised to an exponent"
strmuloption(20) = "none of the above"

strmultiple(6) = "The set of native data types that a particular computer can support is determined by –"
strmuloption(21) = "Type of hardware company"
strmuloption(22) = "What functions have been wired into hardware"
strmuloption(23) = "What software support is required"
strmuloption(24) = "None of these"

strmultiple(7) = "While considering data structure implementation, the factor/s under consideration is/are –"
strmuloption(25) = "Time"
strmuloption(26) = "Time and space"
strmuloption(27) = "Time, space and processor"
strmuloption(28) = "None of the above"

strmultiple(8) = "Pick out invalid statement from following: Queues can be used for –"
strmuloption(29) = "The line printer"
strmuloption(30) = "Access to disk storage"
strmuloption(31) = "Function call"
strmuloption(32) = "None of the above"

strmultiple(9) = "p = 1for(x = n ;x>0;x = x-1)p = x * preturn(p)The above set of steps is an example of :"
strmuloption(33) = "iterative code"
strmuloption(34) = "recursive code"
strmuloption(35) = "implied function"
strmuloption(36) = "none of these"

strmultiple(10) = "Pick out invalid statement from following: Algorithm must be –"
strmuloption(37) = "Efficient"
strmuloption(38) = "Concise and compact"
strmuloption(39) = "Free of ambiguity"
strmuloption(40) = "None of these"

strmulanswer(1) = "Binary number system"
strmulanswer(2) = "0 and 2n-1"
strmulanswer(3) = "A 1 followed by using n-1 zeros to a 0 followed by n-1 ones"
strmulanswer(4) = "A 1 followed by n – 1 zeros to a 0 followed by n-1 1’s"
strmulanswer(5) = "base raised to an exponent"
strmulanswer(6) = "What functions have been wired into hardware"
strmulanswer(7) = "Time and space"
strmulanswer(8) = "Function call"
strmulanswer(9) = "iterative code"
strmulanswer(10) = "None of these"
End Sub

Private Sub multiplechoice_Paper2()
strmultiple(1) = "In top down approach –"
strmuloption(1) = "A problem is subdivided into sub problems; each one is attacked without worrying about other."
strmuloption(2) = "A problem is tackled from beginning to end in one go."
strmuloption(3) = "Subproblems are solved first; then all solutions of subproblems are put together to solve the main problem"
strmuloption(4) = "None of these"

strmultiple(2) = "  In bottom up approach –"
strmuloption(5) = "A problem is subdivided into supproblems ; each ne is attacked without worrying about others."
strmuloption(6) = "A problem is tackled from beginning to end in ne go."
strmuloption(7) = "Subproblems are solved first; then all solution of subproblems are put together to solve the main problem."
strmuloption(8) = "None of these"

strmultiple(3) = "  Modular programming uses-"
strmuloption(9) = "Only top down method"
strmuloption(10) = "Only bottom up method"
strmuloption(11) = "Both a & b"
strmuloption(12) = "None of these"

strmultiple(4) = "  In analysis of algorithm, approximate relationship between the size of the job and the amount of work required to do it is expressed by using –"
strmuloption(13) = "Order of magnitude or Bid-O"
strmuloption(14) = "Central tendency"
strmuloption(15) = "Differential equation."
strmuloption(16) = "None of these"

strmultiple(5) = "  Which is better computing time(in analysis of algorithm)?"
strmuloption(17) = "O(N)"
strmuloption(18) = "O(2n)"
strmuloption(19) = "O(log2N)"
strmuloption(20) = "None of these"

strmultiple(6) = "Structured data type made up of finite collection of ordered elements, all of which are of the same data type is –"
strmuloption(21) = "Record"
strmuloption(22) = "Array"
strmuloption(23) = "File"
strmuloption(24) = "None of these"

strmultiple(7) = "Elements of array are accessed by –"
strmuloption(25) = "Accessing function in built in data structure."
strmuloption(26) = "Mathematical function"
strmuloption(27) = "Index"
strmuloption(28) = "None of these"

strmultiple(8) = "Array is –"
strmuloption(29) = "Linear data structure"
strmuloption(30) = "Non linear data structure"
strmuloption(31) = "Complex data structure"
strmuloption(32) = "None of these"

strmultiple(9) = "  Row major order in two dimensional array refers to –"
strmuloption(33) = "All elements of a row are stored n memory in sequence followed by next row in sequence and so on."
strmuloption(34) = "All elements of a row are stored in memory in sequence followed by next column in sequence and so on."
strmuloption(35) = "All elements of a column are stored in memory in sequeice followed by next column in sequence and so on."
strmuloption(36) = "None of these"

strmultiple(10) = " A data structure, in which a element is added and remoe only from only one end, is known as –"
strmuloption(37) = "Queue"
strmuloption(38) = "Stack"
strmuloption(39) = "Array"
strmuloption(40) = "None of the above"

strmulanswer(1) = "A problem is subdivided into sub problems; each one is attacked without worrying about other."
strmulanswer(2) = "Subproblems are solved first; then all solution of subproblems are put together to solve the main problem."
strmulanswer(3) = "Both a & b"
strmulanswer(4) = "Order of magnitude or Bid-O"
strmulanswer(5) = "O(log2N)"
strmulanswer(6) = "Array"
strmulanswer(7) = "Index"
strmulanswer(8) = "Linear data structure"
strmulanswer(9) = "All elements of a row are stored n memory in sequence followed by next row in sequence and so on."
strmulanswer(10) = "Stack"
End Sub

Private Sub multiplechoice_Paper3()
strmultiple(1) = "Stack is –"
strmuloption(1) = "Static data structure"
strmuloption(2) = "Dynamic data structure"
strmuloption(3) = "In built data structure."
strmuloption(4) = "None of these"

strmultiple(2) = "In linked list, a node contains at least –"
strmuloption(5) = "Node address field, data field"
strmuloption(6) = "Node number , data field"
strmuloption(7) = "Next address field, information field"
strmuloption(8) = "None of these"

strmultiple(3) = "Which is not true ?"
strmuloption(9) = "Abstract data type is the useful tool for specifying the logical properties of data type"
strmuloption(10) = "While defining an abstract data type as a mathematical concept, the space andefficiency is not of major concern."
strmuloption(11) = "Every abstract data type can be implemented using any programming language."
strmuloption(12) = "None of these"

strmultiple(4) = "Which is not true ?"
strmuloption(13) = "Recursion involves entering and exiting a block"
strmuloption(14) = "Recursion may maintain non useful local variables and temporary variables through a stack"
strmuloption(15) = "In a recursive process stacking and unstacking local variables is done by compiler to ensure that no problem arises"
strmuloption(16) = "None of the above"

strmultiple(5) = "  In a queue ( where q.rear and q.front are pointer to the two ends of a queue) :"
strmuloption(17) = "The number of total elements is fixed"
strmuloption(18) = "If q.rear > q.front, it is empty"
strmuloption(19) = "The number of elements at any time is (q.near – a.front - 1)"
strmuloption(20) = "None of the above"

strmultiple(6) = "  Implementation of list in a dynamic fashion is –"
strmuloption(21) = "To call upon the system to allocate and free storage may not be time consuming"
strmuloption(22) = "A set of nodes is not reserved in advance for use"
strmuloption(23) = "The address computation is complex"
strmuloption(24) = "None of the above"

strmultiple(7) = "  Queue-"
strmuloption(25) = "Can be created by setting up an ordinary contiguous array to hold the items"
strmuloption(26) = "Can take care of delete operation automatically"
strmuloption(27) = "Need a one pointer to handle addition and deletion of an item"
strmuloption(28) = "None of the above"

strmultiple(8) = "  Data structure –"
strmuloption(29) = "May be helpful to develop efficient algorithms in different phases of data processing"
strmuloption(30) = "Need not give relationship between data items"
strmuloption(31) = "Is programming language dependent"
strmuloption(32) = "None of the above"

strmultiple(9) = "  In linked list, the logical order ofelements-"
strmuloption(33) = "Is same as their physical arrangement"
strmuloption(34) = "Is not necessarily equivalent to their physical arrangement"
strmuloption(35) = "Is determined by their physical arrangement"
strmuloption(36) = "None of the above"

strmultiple(10) = " Direct or random access of element is not possible in –"
strmuloption(37) = "Linked list"
strmuloption(38) = "Array"
strmuloption(39) = "String"
strmuloption(40) = "None of the above"

strmulanswer(1) = "Dynamic data structure"
strmulanswer(2) = "Next address field, information field"
strmulanswer(3) = "Every abstract data type can be implemented using any programming language."
strmulanswer(4) = "In a recursive process stacking and unstacking local variables is done by compiler to ensure that no problem arises"
strmulanswer(5) = "None of the above"
strmulanswer(6) = "A set of nodes is not reserved in advance for use"
strmulanswer(7) = "Can be created by setting up an ordinary contiguous array to hold the items"
strmulanswer(8) = "May be helpful to develop efficient algorithms in different phases of data processing"
strmulanswer(9) = "Is not necessarily equivalent to their physical arrangement"
strmulanswer(10) = "Linked list"
End Sub

Private Sub multiplechoice_Paper4()
strmultiple(1) = "  Odd man out (in linked list)-"
strmuloption(1) = "Modify"
strmuloption(2) = "Replace"
strmuloption(3) = "To determine(if it is empty)"
strmuloption(4) = "None of the above"

strmultiple(2) = "  The nth node in singly linked list, is accessed via(where n > 1)-"
strmuloption(5) = "The head node"
strmuloption(6) = "The tail node"
strmuloption(7) = "(n-1) nodes."
strmuloption(8) = "None of the above"

strmultiple(3) = "  In linked list, the successive elements –"
strmuloption(9) = "Must occupy contiguous space in memory"
strmuloption(10) = "Need not occupy contiguous space in memory"
strmuloption(11) = "Must not occupy contiguous space in memory."
strmuloption(12) = "None of the above"

strmultiple(4) = "  Linear order in linked list is provided through –"
strmuloption(13) = "Index number"
strmuloption(14) = "The implied position of the node"
strmuloption(15) = "Pointer"
strmuloption(16) = "None of the above"

strmultiple(5) = "  Null pointer is used to tell –"
strmuloption(17) = "End of linked list"
strmuloption(18) = "Empty pointer field of a structure"
strmuloption(19) = "The linked list is empty"
strmuloption(20) = "All of the above"

strmultiple(6) = "  List pointer variable in linked list contains address of the –"
strmuloption(21) = "Fillowing node in the list"
strmuloption(22) = "Current node in the list"
strmuloption(23) = "First node in the list"
strmuloption(24) = "None of the above"

strmultiple(7) = "  Underflow condition in linked list may occur when attempting to-"
strmuloption(25) = "Insert a new node when there is no free space for it"
strmuloption(26) = "Delete a non existent node in the list"
strmuloption(27) = "Delete a node in empty list"
strmuloption(28) = "None of the above"

strmultiple(8) = "  Overflow condition in linked list may occur when attempting to –"
strmuloption(29) = "Create a node when free space pool is empty traverse the nodes when free space pool is empty"
strmuloption(30) = "Traverse the nodes when free space pool is empty"
strmuloption(31) = "Create a node when linked list is empty"
strmuloption(32) = "None of the above"

strmultiple(9) = "  Because of linear structure of linked list having natural linear ordering there is similarity between linked list and array in-"
strmuloption(33) = "Insertion of a node"
strmuloption(34) = "Deletion of a node"
strmuloption(35) = "Traversal of elements of list"
strmuloption(36) = "None of the above"

strmultiple(10) = " Searching a linked list requires liked list be created –"
strmuloption(37) = "In sorted order only"
strmuloption(38) = "In any order"
strmuloption(39) = "Without underflow condition"
strmuloption(40) = "None of the above"


strmulanswer(1) = "None of the above"
strmulanswer(2) = "(n-1) nodes."
strmulanswer(3) = "Need not occupy contiguous space in memory"
strmulanswer(4) = "Pointer"
strmulanswer(5) = "All of the above"
strmulanswer(6) = "First node in the list"
strmulanswer(7) = "Delete a node in empty list"
strmulanswer(8) = "Create a node when free space pool is empty traverse the nodes when free space pool is empty"
strmulanswer(9) = "Traversal of elements of list"
strmulanswer(10) = "In any order"
End Sub

Private Sub multiplechoice_Paper5()
strmultiple(1) = "  Copying all or part of linked list consists of –"
strmuloption(1) = "Taking a null string and adding appropriate elements to the list one by one"
strmuloption(2) = "Taking a null string and putting the address of the beginning of the list"
strmuloption(3) = "Changing only address of the address of beginning of the list"
strmuloption(4) = "None of the above"

strmultiple(2) = "  Deletion of a node in linked list involves keeping track of the address of the node –"
strmuloption(5) = "Which immediately follows the node that is to be deleted"
strmuloption(6) = "Which immediately prededs the node that is to be deleted"
strmuloption(7) = "That is to be deleted"
strmuloption(8) = "None of the above"

strmultiple(3) = "  Header of a linked list is a special node at the –"
strmuloption(9) = "End of the linked list"
strmuloption(10) = "At middle of the list"
strmuloption(11) = "Beginning of the list"
strmuloption(12) = "None of the above"

strmultiple(4) = "  Header linked list in which last node contains the null pointer is called –"
strmuloption(13) = "Grounded header list"
strmuloption(14) = "Circular header list"
strmuloption(15) = "General header list"
strmuloption(16) = "None of the above"

strmultiple(5) = "  Header linked list in which last node points the header node is called –"
strmuloption(17) = "Grounded header list"
strmuloption(18) = "Circular header list"
strmuloption(19) = "General header list"
strmuloption(20) = "None of the above"

strmultiple(6) = "  Polynomials in memory may be maintained through –"
strmuloption(21) = "Linked list with header node"
strmuloption(22) = "One dimensional array"
strmuloption(23) = "Stack"
strmuloption(24) = "None of the above"

strmultiple(7) = "  Having address of the node to be deleted in double linked list, the node can be deleted –"
strmuloption(25) = "With out traversing the list"
strmuloption(26) = "Only after traversing the list form the head"
strmuloption(27) = "Only traversing the list from the  tail"
strmuloption(28) = "Noen of the above"

strmultiple(8) = "  Representing the polynomial in memory using linked list requires each node having –"
strmuloption(29) = "Two fields"
strmuloption(30) = "Three fields"
strmuloption(31) = "More than three fields"
strmuloption(32) = "None of the above"

strmultiple(9) = "  “Get a node, store new element and insert the new node at the top” refers to insert operation in non empty –"
strmuloption(33) = "stack"
strmuloption(34) = "queue"
strmuloption(35) = "array"
strmuloption(36) = "none of the above"

strmultiple(10) = " String concatenation means –"
strmuloption(37) = "Combining two strings"
strmuloption(38) = "Extracting a substring out of a string"
strmuloption(39) = "Partitioning a string into two strings"
strmuloption(40) = "None of these"

strmulanswer(1) = "Taking a null string and adding appropriate elements to the list one by one"
strmulanswer(2) = "Which immediately follows the node that is to be deleted"
strmulanswer(3) = "Beginning of the list"
strmulanswer(4) = "General header list"
strmulanswer(5) = "Grounded header list"
strmulanswer(6) = "One dimensional array"
strmulanswer(7) = "With out traversing the list"
strmulanswer(8) = "Two fields"
strmulanswer(9) = "stack"
strmulanswer(10) = "Extracting a substring out of a string"
End Sub

Private Sub multiplechoice_Paper6()
strmultiple(1) = "  “Get the alue of most recently inserted node and delete the node” refers to operation –"
strmuloption(1) = "POP"
strmuloption(2) = "PUSH"
strmuloption(3) = "EMPTY"
strmuloption(4) = "None of the above"

strmultiple(2) = "  “FRONT = REAR” pointer refers to empty –"
strmuloption(5) = "stack"
strmuloption(6) = "queue"
strmuloption(7) = "array"
strmuloption(8) = "none of the above"

strmultiple(3) = "  Inserting a node after a given node in doubly linked list requires –"
strmuloption(9) = "One pointer changes"
strmuloption(10) = "Two pointer changes"
strmuloption(11) = "Four pointer changes"
strmuloption(12) = "None of the above"

strmultiple(4) = "  A technique which collects all deletes space onto free storage list is called –"
strmuloption(13) = "Static memory allocation"
strmuloption(14) = "Garbage collection"
strmuloption(15) = "Dynamic memory allocation"
strmuloption(16) = "None of the above"

strmultiple(5) = "  Memory allocation at the run time is known as –"
strmuloption(17) = "Static memory allocation"
strmuloption(18) = "Dynamic memory allocation"
strmuloption(19) = "Paging"
strmuloption(20) = "None of the above"

strmultiple(6) = "  Memory allocation at the compile time is known as –"
strmuloption(21) = "Static memory allocation"
strmuloption(22) = "Dynamic memory allocation"
strmuloption(23) = "Paging"
strmuloption(24) = "None of the above"

strmultiple(7) = "  In memory allocation, the first entry which has free block equal to or more than required one is selected –"
strmuloption(25) = "First fit method"
strmuloption(26) = "Best fit method"
strmuloption(27) = "Worst fit method"
strmuloption(28) = "None of the above"

strmultiple(8) = "  In memory allocation an entry which is smallest amongst all entries which equals to or is bigger than the required one iseselected in –"
strmuloption(29) = "First fit method"
strmuloption(30) = "Best fit method"
strmuloption(31) = "Worst fit method"
strmuloption(32) = "None of the above"

strmultiple(9) = "  Reduction in wastage due to fragmentation is minimum in –"
strmuloption(33) = "First fit method"
strmuloption(34) = "Best fit method"
strmuloption(35) = "Worst fit method"
strmuloption(36) = "None of the above"

strmultiple(10) = " Most appropriate data structure in C to represent linked list is –"
strmuloption(37) = "Array"
strmuloption(38) = "Struct"
strmuloption(39) = "Union"
strmuloption(40) = "None of the above"


strmulanswer(1) = "POP"
strmulanswer(2) = "queue"
strmulanswer(3) = "Four pointer changes"
strmulanswer(4) = "Garbage collection"
strmulanswer(5) = "Dynamic memory allocation"
strmulanswer(6) = "Static memory allocation"
strmulanswer(7) = "First fit method"
strmulanswer(8) = "Best fit method"
strmulanswer(9) = "Best fit method"
strmulanswer(10) = "Struct"
End Sub

Private Sub multiplechoice_Paper7()
strmultiple(1) = "  Link of linked list in C is of type –"
strmuloption(1) = "Unsigned integer"
strmuloption(2) = "Pointer to integer"
strmuloption(3) = "Pointer to struct"
strmuloption(4) = "None of the above"

strmultiple(2) = "  Sorted list can be had while –"
strmuloption(5) = "Creating linked list"
strmuloption(6) = "Rearranging the already created list"
strmuloption(7) = "In no way"
strmuloption(8) = "None of the above"

strmultiple(3) = "  struct address{char name[40]; struct address *next.}info;func(struct address *j;{static struct address *I =0;if(i==0) i=j;else i-> next =j;j -> next =0;I = j;}builds the linked list by placing each element"
strmuloption(9) = "in the beginning"
strmuloption(10) = "on the end"
strmuloption(11) = "at the appropriate place in the list"
strmuloption(12) = "none of the above"

strmultiple(4) = "  struct  address {char name[40]; struct address *next;} info;func(struct address *head);{ while (head) {printf(head -> nema);head= head -> next;}}prints the –"
strmuloption(13) = "name in element of linked list pointed by non null head"
strmuloption(14) = "name in all elements of linked list starting from the element"
strmuloption(15) = "none of the above"
strmuloption(16) = "all of the above"

strmultiple(5) = "  A list which can be scanned in either direction is –"
strmuloption(17) = "Singly linked list"
strmuloption(18) = "Doubly linked list"
strmuloption(19) = "Sorted linked list"
strmuloption(20) = "None of the above"

strmultiple(6) = "  datatype *s;datatype *p;datatype *I; p = malloc(100); t = p; s = p + (100/sizeof (datatype)) – sizeof(datatpye);func(datatype j);{if (p > s) {printf(“Overflow\n”); return}*p = j; p++}is referring to –"
strmuloption(21) = "push operation of stack"
strmuloption(22) = "pop operation of stack"
strmuloption(23) = "push operation of queue"
strmuloption(24) = "none of the above"

strmultiple(7) = "  malloc function returns zero value when –"
strmuloption(25) = "memory is successfully allocated"
strmuloption(26) = "allocation error occurs"
strmuloption(27) = "all memory is cleared"
strmuloption(28) = "none of the above"

strmultiple(8) = "  Previously allocates memory can be returned to the system by using –"
strmuloption(29) = "Malloc()"
strmuloption(30) = "Free()"
strmuloption(31) = "Calloc()"
strmuloption(32) = "None of thea bove"

strmultiple(9) = "  malloc() returns pointer to –"
strmuloption(33) = "integer"
strmuloption(34) = "structure"
strmuloption(35) = "character"
strmuloption(36) = "none of the above"

strmultiple(10) = " Function /s associates with dynamic allocation of memory –"
strmuloption(37) = "Calloc"
strmuloption(38) = "Malloc"
strmuloption(39) = "Free"
strmuloption(40) = "All/none of the above"


strmulanswer(1) = "Pointer to struct"""
strmulanswer(2) = "Creating linked list"
strmulanswer(3) = "none of the above"
strmulanswer(4) = "name in all elements of linked list starting from the element"
strmulanswer(5) = "Doubly linked list"
strmulanswer(6) = "push operation of stack"
strmulanswer(7) = "memory is successfully allocated"
strmulanswer(8) = "Free()"
strmulanswer(9) = "character"
strmulanswer(10) = "All/none of the above"
End Sub

Private Sub multiplechoice_Paper8()
strmultiple(1) = "  Which of the following process isfaster for threaded trees compared with their unthreaded counterparts?"
strmuloption(1) = "Insertion"
strmuloption(2) = "Deletion"
strmuloption(3) = "Traversal"
strmuloption(4) = "None of these"

strmultiple(2) = "  Which of the following is NOT a required feature of a good hashing algorithm? It should –"
strmuloption(5) = "Be repeatable"
strmuloption(6) = "Allow even distribution of records throughout the allocates space"
strmuloption(7) = "Minimize synonyms"
strmuloption(8) = "None of the above"

strmultiple(3) = "  Which of the following statements is TRUE in view of a threaded binary tree? It can have –"
strmuloption(9) = "NULL links but no structural links"
strmuloption(10) = "Only structural links but no NULL links"
strmuloption(11) = "Structural links and NULL links"
strmuloption(12) = "None of the above"

strmultiple(4) = "  Which of the following steps is performed first for inorder traversal of a binary tree?"
strmuloption(13) = "Traversal of the left subtree in postorder"
strmuloption(14) = "Processing of the root node"
strmuloption(15) = "Traversal of the left subtree in inorder"
strmuloption(16) = "None of the above"

strmultiple(5) = "  Part of a compiler that keeps record of names of variables and their associated attributes/ values is known as a –"
strmuloption(17) = "Parser"
strmuloption(18) = "Symbol table"
strmuloption(19) = "Lexical analyzer"
strmuloption(20) = "None of the these"

strmultiple(6) = "  A B-tree of order n is also called –"
strmuloption(21) = "(n-n) -1 tree"
strmuloption(22) = "n-(n-2) tree"
strmuloption(23) = "(n-1) –n tree"
strmuloption(24) = "none of these"

strmultiple(7) = "  The preorder traversal of a binary tree begins with –"
strmuloption(25) = "Processing of the root node"
strmuloption(26) = "Traversal of the right subtree in preorder"
strmuloption(27) = "Traversal of the left subtree in preorder"
strmuloption(28) = "None of the above"

strmultiple(8) = "  Hashing refers to the process of deriving –"
strmuloption(29) = "A record key from storage address"
strmuloption(30) = "Storage address from a record key"
strmuloption(31) = "A floating point code from a record key"
strmuloption(32) = "None of the above"

strmultiple(9) = "  If a binary search technique is used for accessing and implementing a symbol table, which of the following statements is TRUE? The ratio of –"
strmuloption(33) = "Insertion time to access time for its entries is quite high"
strmuloption(34) = "Access time to insertion time is quite high"
strmuloption(35) = "Search time to insertion time is very high"
strmuloption(36) = "None of the above"

strmultiple(10) = " If a binary tree is threaded for an inorder traversal order, a NULL right link of any node is replaced by the address of its –"
strmuloption(37) = "Successor "
strmuloption(38) = "Predecessor"
strmuloption(39) = "Root"
strmuloption(40) = "None of these"

strmulanswer(1) = "Traversal"
strmulanswer(2) = "None of the above"
strmulanswer(3) = "None of the above"
strmulanswer(4) = "Traversal of the left subtree in inorder"
strmulanswer(5) = "Symbol table"
strmulanswer(6) = "(n-1) –n tree"
strmulanswer(7) = "Processing of the root node"
strmulanswer(8) = "Storage address from a record key"
strmulanswer(9) = "Insertion time to access time for its entries is quite high"
strmulanswer(10) = "Successor "
End Sub

Private Sub multiplechoice_Paper9()
strmultiple(1) = "  The postorder traversal of a binary tree begins with –"
strmuloption(1) = "The postorder traversal of the left subtree."
strmuloption(2) = "Processing of the root node"
strmuloption(3) = "The postorder traversal of the right subtree"
strmuloption(4) = "None of the above"

strmultiple(2) = "  If a binary tree is threaded for an inorder traversal order NULL left link of any node is replaced by the address of its –"
strmuloption(5) = "Successor"
strmuloption(6) = "Predecessor"
strmuloption(7) = "Root"
strmuloption(8) = "None of these"

strmultiple(3) = "  Which of the following figures is not possible as a balance of any code of an AVL tree?"
strmuloption(9) = "1"
strmuloption(10) = "-1"
strmuloption(11) = "0"
strmuloption(12) = "none of these"

strmultiple(4) = "  The inorder traversal of some binary tree produced he sequence DBEAFC, and the postorder traversal of the same tree produced the sequence DEBFCA. Which of the following is a correct preorder traversal sequence?"
strmuloption(13) = "DBAECF"
strmuloption(14) = "ABEDFC"
strmuloption(15) = "ABDECF"
strmuloption(16) = "None of these"

strmultiple(5) = "  The inorder traversal of some binary tree produced the sequence DBEAFC, and thr postorder traversal of the same tree produced the sequence DEBFCA. What will be the totalnymber of nodes in the left subtree of the given tree?"
strmuloption(17) = "1"
strmuloption(18) = "4"
strmuloption(19) = "5"
strmuloption(20) = "none of these"

strmultiple(6) = "  A balanced binary tree is a binary tree in which the heights of the two subtrees of every node never differ by more than –"
strmuloption(21) = "2"
strmuloption(22) = "1"
strmuloption(23) = "3"
strmuloption(24) = "none of these"

strmultiple(7) = "  What is the minimum number of keys contained in each non root node of a B-tree of order 11?"
strmuloption(25) = "4"
strmuloption(26) = "5"
strmuloption(27) = "3"
strmuloption(28) = "none of these"

strmultiple(8) = "  The postorder traversal of some binary tree produced th sequence CDBFEA, and the inorder traversal of the same tree produced the sequence CBDAFE. What will be the total number of nodes in its right subtree?"
strmuloption(29) = "2"
strmuloption(30) = "3"
strmuloption(31) = "4"
strmuloption(32) = "none of these"

strmultiple(9) = "  The postorder traversal of some binary tree produced the sequence CDBFEA, and the inorder traversal of the same tree produced the sequence CBDAFE. What will be the total number of nodes in its left subtree?"
strmuloption(33) = "2"
strmuloption(34) = "5"
strmuloption(35) = "3"
strmuloption(36) = "none of these"

strmultiple(10) = " Which of the following sttements is TRUE? A B-tree of order –"
strmuloption(37) = "24 contains at least 12 keys in each non root node"
strmuloption(38) = "25 contains at least 12 keys in each node"
strmuloption(39) = "24 contains at least 23 keys in each non root node"
strmuloption(40) = "none of the above"

strmulanswer(1) = "The postorder traversal of the left subtree."
strmulanswer(2) = "Predecessor"
strmulanswer(3) = "none of these"
strmulanswer(4) = "ABDECF"
strmulanswer(5) = "none of these"
strmulanswer(6) = "1"
strmulanswer(7) = "5"
strmulanswer(8) = "2"
strmulanswer(9) = "3"
strmulanswer(10) = "24 contains at least 12 keys in each non root node"
End Sub

Private Sub multiplechoice_Paper10()
strmultiple(1) = "  the inorder traversal of some binary tree produced the sequence CBDA, and the postorder traversal of the same tree produced the sequence CDBA. Which of the following statements is TRUE for the given tree? Its –"
strmuloption(1) = "left subtree is empty, and the total number of nodes in its right subtree is 3."
strmuloption(2) = "Right subtree is empty, and the left subtree contains 3 nodes."
strmuloption(3) = "Both left subtree and right subtree must be non empty."
strmuloption(4) = "None of the above"

strmultiple(2) = "  Which of the following statements is TRUE in view of an threaded binary tree? It can have –"
strmuloption(5) = "Structural links but no thread link"
strmuloption(6) = "Thread links but no structural links"
strmuloption(7) = "Thread links and can also have structural links"
strmuloption(8) = "None of the above"

strmultiple(3) = "  Which of the following terms refers to the balance of a node? Height of its-"
strmuloption(9) = "Left subtree minus height of its right subtree"
strmuloption(10) = "Right subtree minus height of its left subtree"
strmuloption(11) = "Left subtree plus height of tits right subtree"
strmuloption(12) = "None of these"

strmultiple(4) = "  To maintain order, insertions in a lexical ordered binary tree can only be performed at –"
strmuloption(13) = "A leaf level"
strmuloption(14) = "An intermediate level"
strmuloption(15) = "A root node"
strmuloption(16) = "None of these"

strmultiple(5) = "  Which of the following statements is TRUE in view of a complete binary tree?"
strmuloption(17) = "The number of nodes at each level is 1 less than some power of 2"
strmuloption(18) = "The outdegree of every node is exactly equal to 2 or 0"
strmuloption(19) = "The total number of nodes in the tree isalqays some pwer of 2"
strmuloption(20) = "None of the above"

strmultiple(6) = "  Consider the following statements for an AL tree.i. Heights of the two subtrees of every node never differ by more than 1.ii.   Each node of the tree has a balance of 1, -1, or 0.Which of the following statements is TRUE?"
strmuloption(21) = "(i) is FALSe and (ii) is FALSE"
strmuloption(22) = "(i) is FALSE and (ii) TRUE"
strmuloption(23) = "(i) is TRUE and (ii) is FALSE"
strmuloption(24) = "(i) is TRUE and (ii) is TRUE"

strmultiple(7) = "  Which of the following statements is TRUE in view of threaded bnary treed?"
strmuloption(25) = "Insertion into a threaded tree is time consuming but deletion from it is not."
strmuloption(26) = "Both insertions into and deletions from a threaded tree are more time consuming"
strmuloption(27) = "Deletion from a threaded tree is time consuming but insertin into it is not"
strmuloption(28) = "None of the above"

strmultiple(8) = "  Searching of a node in a lexical ordered binary tree is fast, because the number of comparisons needed to determine if an item exists in the tree will be –"
strmuloption(29) = "1 always"
strmuloption(30) = "2 always"
strmuloption(31) = "0 always"
strmuloption(32) = "none of these"

strmultiple(9) = "  Which of the  following statements is TRUe in iew of a multiway search tree? If a node has –"
strmuloption(33) = "4 subtrees, it contains 3 keys"
strmuloption(34) = "5 keys, it has 7 subtrees"
strmuloption(35) = "6 subtrees , it contains 6 keys"
strmuloption(36) = "none of the above"

strmultiple(10) = ".    Which of the following statements is TRUE in view of a lexical ordered binary tree?"
strmuloption(37) = "Its left subtree contains nodes whose keys are lexically greater than the key associated with the root node"
strmuloption(38) = "Its left subtree contains nodes whose keys are lexically less than the key associated with the root node"
strmuloption(39) = "Its root node contains a key which is lexically greater than most of the keys associated with the nodes of the right subtree"
strmuloption(40) = "None of the above"

strmulanswer(1) = "Right subtree is empty, and the left subtree contains 3 nodes."
strmulanswer(2) = "Thread links and can also have structural links"
strmulanswer(3) = "Left subtree minus height of its right subtree"
strmulanswer(4) = "A leaf level"
strmulanswer(5) = "The outdegree of every node is exactly equal to 2 or 0"
strmulanswer(6) = "(i) is TRUE and (ii) is TRUE"
strmulanswer(7) = "Both insertions into and deletions from a threaded tree are more time consuming"
strmulanswer(8) = "none of these"
strmulanswer(9) = "4 subtrees, it contains 3 keys"
strmulanswer(10) = "Its left subtree contains nodes whose keys are lexically less than the key associated with the root node"
End Sub

Private Sub multiplechoice_Paper11()
strmultiple(1) = ". If in a given directed tree the outdegree of eery node is less than or equal to m, then it is called a –"
strmuloption(1) = "Threaded binary tree"
strmuloption(2) = "Complete m-ary tree"
strmuloption(3) = "M-ary tree"
strmuloption(4) = "None of these"

strmultiple(2) = ". If for the gien directed tree outdegree of every node is exactly equal to 2 or 0, and the no. of nodes at level I is 2 raised to the power of i-1, the tree is called –"
strmuloption(5) = "Binary tree"
strmuloption(6) = "Complete binary tree"
strmuloption(7) = "M-ary tree"
strmuloption(8) = "None of these"

strmultiple(3) = ". While performing inorder threading of a binary tree, if the left link represents a structural link, then it is –"
strmuloption(9) = "Replaced with a NULL link"
strmuloption(10) = "Left unchanged"
strmuloption(11) = "Replaced with a new thread link"
strmuloption(12) = "None of the above"

strmultiple(4) = ". While performing inorder threading of a binary tree, if the right link represents a NULL link, then it is-"
strmuloption(13) = "Replaced with a structural link"
strmuloption(14) = "Left unchanged"
strmuloption(15) = "Replaced with a new thread link"
strmuloption(16) = "None of the above"

strmultiple(5) = ". For which of the following trees , is the process of traversal is faster than others ?"
strmuloption(17) = "Threaded  binary tree"
strmuloption(18) = "Lexical ordered binary tree"
strmuloption(19) = "M-ary tree"
strmuloption(20) = "None of these"

strmultiple(6) = ". How many cycles should be contained in a tree ?"
strmuloption(21) = "0"
strmuloption(22) = "at least 1"
strmuloption(23) = "any number"
strmuloption(24) = "none of these"

strmultiple(7) = ". What is the level of a root node of a tree?"
strmuloption(25) = "-1"
strmuloption(26) = "0"
strmuloption(27) = "1"
strmuloption(28) = "none of these"

strmultiple(8) = ". In which of the following trees must balance of each node be either 1, -1, or 0?"
strmuloption(29) = "Threaded  tree"
strmuloption(30) = "Lexical ordered binary tree"
strmuloption(31) = "AVL tree"
strmuloption(32) = "none of these"

strmultiple(9) = ". Level of any node of a tree is –"
strmuloption(33) = "Its distance from the root"
strmuloption(34) = "Height of its left subtree minus height of its right subtree"
strmuloption(35) = "Height of its right subtree minus height of its left subtree"
strmuloption(36) = "None of the above"

strmultiple(10) = "Which of the following sttements is TRUe in iew of a threaded storage representation of a binary tree?"
strmuloption(37) = "Its inorder threading is different from its postorder threading"
strmuloption(38) = "The number of NULL links in a threaded tree is very small, but it is non zero."
strmuloption(39) = "The number of thread links is reduced to zero."
strmuloption(40) = "None of the above"

strmulanswer(1) = "M-ary tree"
strmulanswer(2) = "Complete binary tree"
strmulanswer(3) = "Left unchanged"
strmulanswer(4) = "Replaced with a new thread link"
strmulanswer(5) = "Threaded  binary tree"
strmulanswer(6) = "0"
strmulanswer(7) = "0"
strmulanswer(8) = "AVL tree"
strmulanswer(9) = "Its distance from the root"
strmulanswer(10) = "Its inorder threading is different from its postorder threading"
End Sub

Private Sub multiplechoice_Paper12()
strmultiple(1) = "Which string operation is not available in ‘C’?"
strmuloption(1) = "Concatenation"
strmuloption(2) = "Pattern matching"
strmuloption(3) = "Reversing the string"
strmuloption(4) = "All of these are available"

strmultiple(2) = "If graph G has no edges then corresponding adjacency matrix is –"
strmuloption(5) = "Unit matrix"
strmuloption(6) = "Zero matrix"
strmuloption(7) = "Matrix with all 1’s"
strmuloption(8) = "None of these"

strmultiple(3) = "Language developed for string manipulation –"
strmuloption(9) = "SNOBOL"
strmuloption(10) = "LISP"
strmuloption(11) = "Ada"
strmuloption(12) = "None of these"

strmultiple(4) = "Worst case efficiency of binary search is –"
strmuloption(13) = "Log2 n+1"
strmuloption(14) = "N"
strmuloption(15) = "N(2)"
strmuloption(16) = "None of these"

strmultiple(5) = "    What property is not expected from good hashing technique should –"
strmuloption(17) = "Produce keys uniformly distributed oer the range"
strmuloption(18) = "Easy to program"
strmuloption(19) = "Produce no collisions"
strmuloption(20) = "All properties are expected"

strmultiple(6) = "    What is not true for linear collision processing?"
strmuloption(21) = "It is easier to program"
strmuloption(22) = "It may induce more collision"
strmuloption(23) = "It requires Space for links"
strmuloption(24) = "All are true"

strmultiple(7) = "Graph structure is available in –"
strmuloption(25) = "C"
strmuloption(26) = "C++"
strmuloption(27) = "Pascal"
strmuloption(28) = "None of these"

strmultiple(8) = "A digraph in which outdegree is same as indegree is called –"
strmuloption(29) = "Balanced"
strmuloption(30) = "Symmetric"
strmuloption(31) = "Regular"
strmuloption(32) = "None of these"

strmultiple(9) = "A vertex with degree one in a graph is called –"
strmuloption(33) = "Leaf"
strmuloption(34) = "Pendant vertex"
strmuloption(35) = "End vertex"
strmuloption(36) = "None of these"

strmultiple(10) = "Which one is not a representation of a graph?"
strmuloption(37) = "Adjacency matrix"
strmuloption(38) = "Edge listing"
strmuloption(39) = "Adjacency list"
strmuloption(40) = "All represent graphs"

strmulanswer(1) = "Concatenation"
strmulanswer(2) = "Zero matrix"
strmulanswer(3) = "SNOBOL"
strmulanswer(4) = "Log2 n+1"
strmulanswer(5) = "Produce no collisions"
strmulanswer(6) = "It requires Space for links"
strmulanswer(7) = "None of these"
strmulanswer(8) = "Balanced"
strmulanswer(9) = "Pendant vertex"
strmulanswer(10) = "All represent graphs"
End Sub

Private Sub multiplechoice_Paper13()
strmultiple(1) = ". In an adjacency matrix parallel edges are given by –"
strmuloption(1) = "Similar columns"
strmuloption(2) = "Similar rows"
strmuloption(3) = "Not representable"
strmuloption(4) = "None of these"

strmultiple(2) = ". If x is the adjacency matrix of a graph then (I,j)th entry of x represents –"
strmuloption(5) = "No. of edge sequences of length k between ith & jth vertex"
strmuloption(6) = "No. of paths of length k between ith & jth vertex."
strmuloption(7) = "No. of walk of length k between ith & jth vertex"
strmuloption(8) = "None of these"

strmultiple(3) = ". Breadth first search –"
strmuloption(9) = "Scans all incident edges before moving to other vertes"
strmuloption(10) = "Scans adjacent unisited vertex as soon as possible"
strmuloption(11) = "Is same as backtracking"
strmuloption(12) = "None of these"

strmultiple(4) = ". What is not true about linked list representation for strings ?"
strmuloption(13) = "Insertion and deletion are easier to handle"
strmuloption(14) = "It allows dynamic allocation of memory"
strmuloption(15) = "Working with substrings is easy"
strmuloption(16) = "All are true"

strmultiple(5) = ". Name the sort for which time is not proportional to n2"
strmuloption(17) = "Selection sort"
strmuloption(18) = "Bubble sort"
strmuloption(19) = "Quick sort"
strmuloption(20) = "None of these"

strmultiple(6) = ". The element at the rootof heap is –"
strmuloption(21) = "Largest"
strmuloption(22) = "Smallest"
strmuloption(23) = "Depending on type of heap it may be smallest or largest"
strmuloption(24) = "None of these"

strmultiple(7) = ". this rort does not use diide and conquer methodology-"
strmuloption(25) = "merge sort"
strmuloption(26) = "quick sort"
strmuloption(27) = "bubble sort"
strmuloption(28) = "all use diide and conquer methodology"

strmultiple(8) = ". Name the sort in which array to be sorted is partitioned again and again in such a way that all elements less than or equal to partition element appear before it and thos which are greater appear after it."
strmuloption(29) = "Merge sort"
strmuloption(30) = "Quick sort"
strmuloption(31) = "Selection sort"
strmuloption(32) = "None of these"

strmultiple(9) = ". Worst case efficienct of this search is n."
strmuloption(33) = "Sequential search"
strmuloption(34) = "Binary search"
strmuloption(35) = "Indexed search"
strmuloption(36) = "None of these"

strmultiple(10) = "Which is not a hashing technique?"
strmuloption(37) = "Diision remainder method"
strmuloption(38) = "Folding"
strmuloption(39) = "Mid square"
strmuloption(40) = "All are hashing technique"

strmulanswer(1) = "Not representable"
strmulanswer(2) = "No. of edge sequences of length k between ith & jth vertex"
strmulanswer(3) = "Scans all incident edges before moving to other vertes"
strmulanswer(4) = "Working with substrings is easy"
strmulanswer(5) = "Quick sort"
strmulanswer(6) = "Depending on type of heap it may be smallest or largest"
strmulanswer(7) = "bubble sort"
strmulanswer(8) = "Quick sort"
strmulanswer(9) = "Sequential search"
strmulanswer(10) = "All are hashing technique"
End Sub

Private Sub multiplechoice_Paper14()
strmultiple(1) = ". Synonym means-"
strmuloption(1) = "Two keys generating same address"
strmuloption(2) = "Two keys aing same values"
strmuloption(3) = "Two keys generating addresses on same sector of track"
strmuloption(4) = "None of these"

strmultiple(2) = ". Extra space in each record is kept for this collision processing method –"
strmuloption(5) = "Linear collision processing"
strmuloption(6) = "Linked collision processing"
strmuloption(7) = "Quadratic collision processing"
strmuloption(8) = "None of these"

strmultiple(3) = ". this method of collision processing requires prime area and overflow area."
strmuloption(9) = "Linked collision processing"
strmuloption(10) = "Quadratic collision processing"
strmuloption(11) = "Linear collision processing"
strmuloption(12) = "None of these"

strmultiple(4) = ". In this method of collision processing some of the addresses may remain unchecked"
strmuloption(13) = "Linked collision processing"
strmuloption(14) = "Linear collision processing"
strmuloption(15) = "Quadratic collision processing"
strmuloption(16) = "None of these"

strmultiple(5) = ". In which case adjacency list representation of graph is not useful?"
strmuloption(17) = "When no. of edges is small"
strmuloption(18) = "When no. of vertices are changing due to insertion and deletion"
strmuloption(19) = "In breadth first traversal"
strmuloption(20) = "It is useful in all above cases"

strmultiple(6) = ". Which method of traversal uses queue to hold nodes that are waiting to be processed?"
strmuloption(21) = "Breadth first"
strmuloption(22) = "Depth first"
strmuloption(23) = "D-search"
strmuloption(24) = "None of these"

strmultiple(7) = ". Which method of traversal does not use stack to hold nodes that are waiting to be processed?"
strmuloption(25) = "Breadth first"
strmuloption(26) = "Depth first"
strmuloption(27) = "D-search"
strmuloption(28) = "None of these"

strmultiple(8) = ". Pattern matching refers to –"
strmuloption(29) = "Finding position where a string pattern P first appears in given string Q"
strmuloption(30) = "Checking whether two strings are identical"
strmuloption(31) = "Checking whether a pattern P occurs in given string Q "
strmuloption(32) = "None of these"

strmultiple(9) = ". This sort inserts each elements A(k) into its proper position in the previoysly sorted subarray A(1), …, A(k-1)."
strmuloption(33) = "Insertion sort"
strmuloption(34) = "Radix sort"
strmuloption(35) = "Merge sort"
strmuloption(36) = "Bubble sort"

strmultiple(10) = ".    This sort finds location LOC of smallest elements in A(K),…, A(N) nd then interchange A(LOC) whti A(K) for k =1,…N-1."
strmuloption(37) = "Shuffle sort"
strmuloption(38) = "Quich sort"
strmuloption(39) = "Heap sort"
strmuloption(40) = "None of these"

strmulanswer(1) = "Two keys generating same address"
strmulanswer(2) = "Linked collision processing"
strmulanswer(3) = "Linked collision processing"
strmulanswer(4) = "Quadratic collision processing"
strmulanswer(5) = "It is useful in all above cases"
strmulanswer(6) = "Breadth first"
strmulanswer(7) = "Breadth first"
strmulanswer(8) = "Finding position where a string pattern P first appears in given string Q"
strmulanswer(9) = "Insertion sort"
strmulanswer(10) = "Shuffle sort"
End Sub

Private Sub multiplechoice_Paper15()
strmultiple(1) = ". A suitable structure for breadth first and depth first traversal of graphs."
strmuloption(1) = "Adjacency matrix"
strmuloption(2) = "Edge listing"
strmuloption(3) = "Adjacency list"
strmuloption(4) = "None of these"

strmultiple(2) = ". In this search keys must be ordered"
strmuloption(5) = "Sequential search"
strmuloption(6) = "Hashing"
strmuloption(7) = "Binary search"
strmuloption(8) = "None of these"

strmultiple(3) = ". This searching method requires that all keys must reside in internal memory"
strmuloption(9) = "Binary search"
strmuloption(10) = "Sequential search"
strmuloption(11) = "Hashing "
strmuloption(12) = "None of these"

strmultiple(4) = ". In this sort, file is diided into subfules which are to be independently sorted and the nmerged."
strmuloption(13) = "Quick sort"
strmuloption(14) = "Heap sort"
strmuloption(15) = "Bubble sort"
strmuloption(16) = "None of these"

strmultiple(5) = ". With this method of string representation problem of garbage collection arises."
strmuloption(17) = "Fixed length method"
strmuloption(18) = "Index table method"
strmuloption(19) = "Linked list method"
strmuloption(20) = "None of these"

strmultiple(6) = ". Back tracking is annothe name for this method oftraersal"
strmuloption(21) = "Depth first"
strmuloption(22) = "Breadth first"
strmuloption(23) = "D-search"
strmuloption(24) = "None of these"

strmultiple(7) = ". what is not true for adjacency matrix of a graph?"
strmuloption(25) = "It is symmetric matrix"
strmuloption(26) = "It is bit matrix"
strmuloption(27) = "Diagonal has all zeros"
strmuloption(28) = "All are true statements"

strmultiple(8) = ". What is not true about hashing."
strmuloption(29) = "Also called KAT"
strmuloption(30) = "It is an algorithm producing address given a key."
strmuloption(31) = "ISAM is a technique for hashing"
strmuloption(32) = "All are true."

strmultiple(9) = ". What is not true about collision?"
strmuloption(33) = "It occurs when two or more key produce same address"
strmuloption(34) = "Every hashing algorithm may produce it"
strmuloption(35) = "Another hashing function may be used to resolve collision"
strmuloption(36) = "All are true"

strmultiple(10) = ".    A two person perfect information game can be represented by a –"
strmuloption(37) = "Binary tree"
strmuloption(38) = "N-ary tree"
strmuloption(39) = "Graph"
strmuloption(40) = "AVL tree"

strmulanswer(1) = "Adjacency list"
strmulanswer(2) = "Binary search"
strmulanswer(3) = "Binary search"
strmulanswer(4) = "None of these"
strmulanswer(5) = "Index table method"
strmulanswer(6) = "Depth first"
strmulanswer(7) = "It is symmetric matrix"
strmulanswer(8) = "ISAM is a technique for hashing"
strmulanswer(9) = "Another hashing function may be used to resolve collision"
strmulanswer(10) = "N-ary tree"
End Sub

Private Sub multiplechoice_Paper16()
strmultiple(1) = ". Which of the following is an example of linear data structure?"
strmuloption(1) = "Linked list"
strmuloption(2) = "binary tree"
strmuloption(3) = "graph"
strmuloption(4) = "tries"

strmultiple(2) = ". Dynamic storage allocation cannot be used in C without-"
strmuloption(5) = "Enumerated data types"
strmuloption(6) = "Uuions"
strmuloption(7) = "Pointer"
strmuloption(8) = "Files"

strmultiple(3) = ". Average case time complexity of the quicksort algorithm is more than –"
strmuloption(9) = "O(Nlog(2) N)"
strmuloption(10) = "O(Nlnn)"
strmuloption(11) = "O(N(2))"
strmuloption(12) = "O(N(3))"

strmultiple(4) = "Worst case eme complexity of the heap sort algorithm is–"
strmuloption(13) = "O(Nlog(2)N)"
strmuloption(14) = "O(NlnN)"
strmuloption(15) = "O(N(2))"
strmuloption(16) = "O(N(3))"

strmultiple(5) = ". A sparse matrix is better represented using a/an –"
strmuloption(17) = "Array"
strmuloption(18) = "binary tree"
strmuloption(19) = "multi linked list"
strmuloption(20) = "stack"

strmultiple(6) = ". atoi() is a C function which converts character strings to –"
strmuloption(21) = "equivalent floating point value"
strmuloption(22) = "equivalent integer value"
strmuloption(23) = "equivalent octal value"
strmuloption(24) = "equivalent hex value"

strmultiple(7) = ". A binary tree with the peoperty that the key stored at each node is greater than those stored in tis children, is a –"
strmuloption(25) = "AVL tree"
strmuloption(26) = "Heap"
strmuloption(27) = "Binary search tree"
strmuloption(28) = "None of the above"

strmultiple(8) = ". Row major ordering of a rectangular array means –"
strmuloption(29) = "Columns are stored before rows"
strmuloption(30) = "Rows are stored in a sequence starting from the first element"
strmuloption(31) = "Greatest element of each row is stored first in memory."
strmuloption(32) = "None of the above"

strmultiple(9) = ". Folding is a method of generating –"
strmuloption(33) = "Index function for a triangular matrix"
strmuloption(34) = "Header node of a circularly linked list"
strmuloption(35) = "A hash function"
strmuloption(36) = "None of the above"

strmultiple(10) = ".    The technique of linear probing for collision resolution can lead to-"
strmuloption(37) = "Clustering"
strmuloption(38) = "Radix sort"
strmuloption(39) = "Efficient storage utilization"
strmuloption(40) = "Overflow"


strmulanswer(1) = "Linked list"
strmulanswer(2) = "Pointer"
strmulanswer(3) = "O(Nlog(2) N)"
strmulanswer(4) = "O(Nlog(2)N)"
strmulanswer(5) = "multi linked list"
strmulanswer(6) = "equivalent integer value"
strmulanswer(7) = "Heap"
strmulanswer(8) = "Rows are stored in a sequence starting from the first element"
strmulanswer(9) = "A hash function"
strmulanswer(10) = "Clustering"
End Sub

Private Sub multiplechoice_Paper17()
strmultiple(1) = ". In C, it is best to implement chained hash table as an array of –"
strmuloption(1) = "Strings"
strmuloption(2) = "Structures"
strmuloption(3) = "Pointers"
strmuloption(4) = "None of the above"

strmultiple(2) = ". For sorting contiguous list ofrecords quicksort may be preferred oer mergesort because –"
strmuloption(5) = "It requires less time always"
strmuloption(6) = "It does not require extra space for an auxilliary storage"
strmuloption(7) = "It requires more more programming effort"
strmuloption(8) = "Some programming languages does not support recursion"

strmultiple(3) = ". In Quicksort, a desirable choice for the partitioning element will be–  "
strmuloption(9) = "First element of the list"
strmuloption(10) = "Last element of the list"
strmuloption(11) = "Median of the list"
strmuloption(12) = "A randomly chosen element of the list"

strmultiple(4) = ". Recursion proides an attractive mechanism for designing efficient program for –"
strmuloption(13) = "Generating the nth Fibonacci number"
strmuloption(14) = "Generating n!"
strmuloption(15) = "Searching with backtrack"
strmuloption(16) = "Computing square root of a number"

strmultiple(5) = ". For implementation of recursie and re-enntrant programs we need to use computer memory as a/an –"
strmuloption(17) = "Queue "
strmuloption(18) = "Heap"
strmuloption(19) = "Array"
strmuloption(20) = "Stack"

strmultiple(6) = ". A dynamic data structure where we cn search for desired records in O(log n) time is "
strmuloption(21) = "Heap"
strmuloption(22) = "Binary search tree"
strmuloption(23) = "Circularly liked list"
strmuloption(24) = "Array"

strmultiple(7) = ". Inorder traversal of the binary search tree implies isiting records in –"
strmuloption(25) = "The order of increasing magnitude of their key"
strmuloption(26) = "The order of decreasing magnitude of their key"
strmuloption(27) = "Arbitraryt order"
strmuloption(28) = "None of the above"

strmultiple(8) = ". Knuth-Morris –Pratt string matchng algorithm, for finding an instance of a string of length p, in another string of length  m runs, in an amount of time at worst proportional to - "
strmuloption(29) = "M"
strmuloption(30) = "M + p"
strmuloption(31) = "P"
strmuloption(32) = "M*p"

strmultiple(9) = ". In C strings are stored in –"
strmuloption(33) = "Linked list of characters"
strmuloption(34) = "Linked list of unsigned characters"
strmuloption(35) = "Array of characters"
strmuloption(36) = "Array of integers"

strmultiple(10) = ".    Macro expansion in C is an example of –"
strmuloption(37) = "String matching and manipulation operation"
strmuloption(38) = "Application of hashing for a practial problem"
strmuloption(39) = "String operation which does not require string matching."
strmuloption(40) = "None of the above"


strmulanswer(1) = "Pointers"
strmulanswer(2) = "It does not require extra space for an auxilliary storage"
strmulanswer(3) = "Median of the list"
strmulanswer(4) = "Searching with backtrack"
strmulanswer(5) = "Stack"
strmulanswer(6) = "Binary search tree"
strmulanswer(7) = "The order of increasing magnitude of their key"
strmulanswer(8) = "M"
strmulanswer(9) = "Array of characters"
strmulanswer(10) = "String matching and manipulation operation"
End Sub

Private Sub multiplechoice_Paper18()
strmultiple(1) = ". A threaded binary tree has following problem –"
strmuloption(1) = "More time consuming tree traversal"
strmuloption(2) = "Non sequential memory allocation"
strmuloption(3) = "Additional storage requirement"
strmuloption(4) = "None of the above"

strmultiple(2) = ". Representation of list structures with items that are either sublists or individual items, by a singly linked list is –"
strmuloption(5) = "Possible with an additional tag field (storing information or address of the sublist in the information field of the node)"
strmuloption(6) = "Not possible"
strmuloption(7) = "Possible without additional tag field"
strmuloption(8) = "Not possible without additional link fields"

strmultiple(3) = ". Compaction is –"
strmuloption(9) = "Compression of information fields of a list"
strmuloption(10) = "Moving all emory locations currently in use into a single contiguous region of memory"
strmuloption(11) = "Conversion of an link list to an array"
strmuloption(12) = "None of the above"

strmultiple(4) = ". Reference count may be maintained for memory locations used in a linked list for the purpose of –"
strmuloption(13) = "Copying"
strmuloption(14) = "Compaction"
strmuloption(15) = "Reclamation"
strmuloption(16) = "Traversal"

strmultiple(5) = ". Stability of a sorting algorithm is important for –"
strmuloption(17) = "Sorting records on the basis of multiple keys"
strmuloption(18) = "Worst case performance of the sorting algorithm"
strmuloption(19) = "Sorting alphanumeric keys because they are likely to be same"
strmuloption(20) = "None of the above"

strmultiple(6) = ". Sparse matrices are preresented using –"
strmuloption(21) = "Singly linked list where each node contains non zero elements of the matrix"
strmuloption(22) = "Circularly linked list for each row and column with each node corresponding to non zero element"
strmuloption(23) = "A binary tree"
strmuloption(24) = "Doubly linked list"

strmultiple(7) = ". In which collision processing method is it not required to detect a given list position, if it is occupied or not ?"
strmuloption(25) = "Quadratic"
strmuloption(26) = "Linked"
strmuloption(27) = "Rehashing"
strmuloption(28) = "None of the above"

strmultiple(8) = ". When a new element is inserted in the middle of a linked list –"
strmuloption(29) = "Only the elements that appear after the new element need to be moved"
strmuloption(30) = "Only the elements that appear before the new element need to be moved"
strmuloption(31) = "The elements before and after the new element need to be moved"
strmuloption(32) = "None of the above"

strmultiple(9) = ". When we concatenate two strings s1 and s2 o size m and n, size of the resultant string is always –"
strmuloption(33) = "Of size m +n"
strmuloption(34) = "Of size less than m + n"
strmuloption(35) = "Of size m* n"
strmuloption(36) = "Of size max(m,n)"

strmultiple(10) = ".    Following data structure can suffer from overflow problem –"
strmuloption(37) = "Stack implemented using array"
strmuloption(38) = "Linked list"
strmuloption(39) = "Circularly linked list"
strmuloption(40) = "None of the above"

strmulanswer(1) = "Additional storage requirement"
strmulanswer(2) = "Possible with an additional tag field (storing information or address of the sublist in the information field of the node)"
strmulanswer(3) = "Moving all emory locations currently in use into a single contiguous region of memory"
strmulanswer(4) = "Reclamation"
strmulanswer(5) = "Sorting records on the basis of multiple keys"
strmulanswer(6) = "Circularly linked list for each row and column with each node corresponding to non zero element"
strmulanswer(7) = "None of the above"
strmulanswer(8) = "None of the above"
strmulanswer(9) = "Of size m +n"
strmulanswer(10) = "Stack implemented using array"
End Sub

Private Sub multiplechoice_Paper19()
strmultiple(1) = ". Consider the following definition of a structure struct node{ int data;struct node *lptr, *rptr;}This is suitable for an element of "
strmuloption(1) = "Singly linked list"
strmuloption(2) = "Circularly linked list"
strmuloption(3) = "Hash table"
strmuloption(4) = "Queue"

strmultiple(2) = ". The tree traversal technique in which the root is traversed before its children is known as –"
strmuloption(5) = "Post order traversal"
strmuloption(6) = "Pre-order traversal"
strmuloption(7) = "Inorder traversal"
strmuloption(8) = "None of the above"

strmultiple(3) = ". A data structure in which insertion and deletion can take place at both the end is called –"
strmuloption(9) = "Dequeue"
strmuloption(10) = "Stack"
strmuloption(11) = "Circular queue"
strmuloption(12) = "None of the above"

strmultiple(4) = ". The maximum number of nodes in a binary tree, of depth 5 is –"
strmuloption(13) = "31"
strmuloption(14) = "16"
strmuloption(15) = "32"
strmuloption(16) = "15"

strmultiple(5) = ". In two’s complement representation of an integer, if all the bits are 1, then the number is –"
strmuloption(17) = "1"
strmuloption(18) = "-1"
strmuloption(19) = "-32768"
strmuloption(20) = "none of the above"

strmultiple(6) = ". A is an array of size mxn, stored in the row major order. If the address of the first element in the array is M, the address of the element A(I,j) (A(0,0)) is the first elemtnt of the array and each element occupies one location in memory ) is "
strmuloption(21) = "M + (i-1) * m + j – 1"
strmuloption(22) = "M + i*m+j"
strmuloption(23) = "M + (j-1)*m + i-1"
strmuloption(24) = "M + (i-1) *n + j-1"

strmultiple(7) = ". A binary tree of depth d is almost complete binary tree, if –"
strmuloption(25) = "Each leaf in the tree is either at leel d or at level d-1"
strmuloption(26) = "For any node n in the tree with a right descendant at level d, all the left descendants of n that are leaes are also at level d."
strmuloption(27) = "Both (a) and (b) are true. "
strmuloption(28) = "None of these"

strmultiple(8) = ". Let P be the queue of integers defined as follows –#define MAXQ 500struct queue {int items[MAXQ];int front,rear;}q;To insert an elements in the queue we can use –"
strmuloption(29) = "++q.items[q.rear] = x"
strmuloption(30) = "q.items[++q.rear] = x"
strmuloption(31) = "q.items[q.rear] ++ = x"
strmuloption(32) = "none of the above"

strmultiple(9) = ". Adjacency matrix of a digraph is –"
strmuloption(33) = "Identity matrix"
strmuloption(34) = "Symmetric matrix"
strmuloption(35) = "Asymmetric matrix"
strmuloption(36) = "None of the above"

strmultiple(10) = ".    Let p be a pointer to a structure. A member mem of that structure is referenced by –"
strmuloption(37) = "*p.mem"
strmuloption(38) = "*(p.m)"
strmuloption(39) = "(*p).mem"
strmuloption(40) = "none of the above"

strmulanswer(1) = "Circularly linked list"
strmulanswer(2) = "Pre-order traversal"
strmulanswer(3) = "Dequeue"
strmulanswer(4) = "31"
strmulanswer(5) = "-1"
strmulanswer(6) = "M + (i-1) *n + j-1"
strmulanswer(7) = "For any node n in the tree with a right descendant at level d, all the left descendants of n that are leaes are also at level d."
strmulanswer(8) = "q.items[++q.rear] = x"
strmulanswer(9) = "Asymmetric matrix"
strmulanswer(10) = "(*p).mem"
End Sub

Private Sub multiplechoice_Paper20()
strmultiple(1) = ". Using arrays most efficient implementation of queue is as –"
strmuloption(1) = "Linear queue"
strmuloption(2) = "Priority queue"
strmuloption(3) = "Circular queue"
strmuloption(4) = "None of the above"

strmultiple(2) = ". The best fit policy for selection of a memory block for satisfying a memory request involves –"
strmuloption(5) = "Finding the first block that meets the requirement"
strmuloption(6) = "Finding the smallest block that meets the requirement"
strmuloption(7) = "Finding the largest block that meets the requirement"
strmuloption(8) = "None of the above"

strmultiple(3) = ". A n-ary tree can be represented using –"
strmuloption(9) = "Priority queue"
strmuloption(10) = "Input restricted dequeue"
strmuloption(11) = "Binary tree"
strmuloption(12) = "None of the above"

strmultiple(4) = ". A complete k-ary tree can be efficiently represented using –"
strmuloption(13) = "Binary tree"
strmuloption(14) = "An array"
strmuloption(15) = "A graph"
strmuloption(16) = "None of the above"

strmultiple(5) = ". We can efficiently reverse a string using a -   "
strmuloption(17) = "Linear queue"
strmuloption(18) = "Circular queue"
strmuloption(19) = "Stack"
strmuloption(20) = "Doubly linked list"

strmultiple(6) = ". We can convert an infix expression to a postfix expression using a –"
strmuloption(21) = "Stack"
strmuloption(22) = "Queue"
strmuloption(23) = "Dequeue"
strmuloption(24) = "None of the above"

strmultiple(7) = ". Simplest way to organize the information about available memory blocks in a system is –"
strmuloption(25) = "AVL tree"
strmuloption(26) = "Doubly linked circular list"
strmuloption(27) = "Stack"
strmuloption(28) = "File"

strmulanswer(1) = "Circular queue"
strmulanswer(2) = "Finding the smallest block that meets the requirement"
strmulanswer(3) = "Binary tree"
strmulanswer(4) = "An array"
strmulanswer(5) = "Stack"
strmulanswer(6) = "Stack"
strmulanswer(7) = "Doubly linked circular list"

End Sub


