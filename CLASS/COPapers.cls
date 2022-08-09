VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COPapers"
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
 
Public Function CO(PaperType As String, CoPaper As Paper, mCat As Category, Number As Integer)
'ptype is the paper type
'itpaper is the number of paper as paper1,paper2 ...
'mcat is the category
'number is the number of questions, or options or answers to be taken

 StrPaperType = PaperType
' StrPaper = Paper
If PaperType = "Fill in the Blanks" Then
   If CoPaper = Paper1 Then Fill_in_the_Blanks_Paper1
   If CoPaper = Paper2 Then Fill_in_the_Blanks_Paper2
   If CoPaper = Paper3 Then Fill_in_the_Blanks_Paper3
   If CoPaper = Paper4 Then Fill_in_the_Blanks_Paper4
   If CoPaper = Paper5 Then Fill_in_the_Blanks_Paper5
   If CoPaper = Paper6 Then Fill_in_the_Blanks_Paper6
   If CoPaper = Paper7 Then Fill_in_the_Blanks_Paper7
   If CoPaper = Paper8 Then Fill_in_the_Blanks_Paper8
   If CoPaper = Paper9 Then Fill_in_the_Blanks_Paper9
   If CoPaper = Paper10 Then Fill_in_the_Blanks_Paper10
   
ElseIf PaperType = "Matching Columns" Then
       If CoPaper = Paper1 Then Matching_Columns_Paper1
       If CoPaper = Paper2 Then Matching_Columns_Paper2
       If CoPaper = Paper3 Then Matching_Columns_Paper3
       If CoPaper = Paper4 Then Matching_Columns_Paper4
       If CoPaper = Paper5 Then Matching_Columns_Paper5
       If CoPaper = Paper6 Then Matching_Columns_Paper6
       If CoPaper = Paper7 Then Matching_Columns_Paper7
       If CoPaper = Paper8 Then Matching_Columns_Paper8
       If CoPaper = Paper9 Then Matching_Columns_Paper9
       If CoPaper = Paper10 Then Matching_Columns_Paper10
       
ElseIf PaperType = "True Or False" Then
       If CoPaper = Paper1 Then TrueorFalse_Paper1
       If CoPaper = Paper2 Then TrueorFalse_Paper2
       If CoPaper = Paper3 Then TrueorFalse_Paper3
       If CoPaper = Paper4 Then TrueorFalse_Paper4
       If CoPaper = Paper5 Then TrueorFalse_Paper5
       If CoPaper = Paper6 Then TrueorFalse_Paper6
       If CoPaper = Paper7 Then TrueorFalse_Paper7
       If CoPaper = Paper8 Then TrueorFalse_Paper8
       If CoPaper = Paper9 Then TrueorFalse_Paper9
       
ElseIf PaperType = "Descriptive Questions" Then
       If CoPaper = Paper1 Then questions_paper1
       If CoPaper = Paper2 Then questions_paper2
       If CoPaper = Paper3 Then questions_paper3
       If CoPaper = Paper4 Then questions_paper4
       
ElseIf PaperType = "Multiple Choice" Then
       If CoPaper = Paper1 Then multiplechoice_Paper1
       If CoPaper = Paper2 Then multiplechoice_Paper2
       If CoPaper = Paper3 Then multiplechoice_Paper3
       If CoPaper = Paper4 Then multiplechoice_Paper4
       If CoPaper = Paper5 Then multiplechoice_Paper5
       If CoPaper = Paper6 Then multiplechoice_Paper6
       If CoPaper = Paper7 Then multiplechoice_Paper7
       If CoPaper = Paper8 Then multiplechoice_Paper8
       If CoPaper = Paper9 Then multiplechoice_Paper9
       If CoPaper = Paper10 Then multiplechoice_Paper10
       If CoPaper = Paper11 Then multiplechoice_Paper11
       If CoPaper = Paper12 Then multiplechoice_Paper12
End If

If mCat = Fill_Questions Then
   CO = StrQuestions(Number)
ElseIf mCat = Fill_Options Then
       CO = stroptions(Number)
ElseIf mCat = Fill_Answers Then
       CO = stranswers(Number)
       
ElseIf mCat = Matching_ColumnA Then
       CO = strcolumnA(Number)
ElseIf mCat = Matching_ColumnB Then
       CO = strcolumnB(Number)
ElseIf mCat = Matching_Answers Then
       CO = stranswers(Number)
       
ElseIf mCat = Questions Then
       CO = StrQuestions(Number)
       
ElseIf mCat = TF_Questions Then
       CO = StrQuestions(Number)
ElseIf mCat = TF_Answers Then
       CO = stranswers(Number)
       
ElseIf mCat = MC_Questions Then
       CO = strmultiple(Number)
       
ElseIf mCat = MC_Options_1 Then
       CO = strmuloption(Number)
       
ElseIf mCat = MC_Options_2 Then
       CO = strmuloption(Number)
       
ElseIf mCat = MC_Options_3 Then
       CO = strmuloption(Number)
       
ElseIf mCat = MC_Options_4 Then
       CO = strmuloption(Number)
       
ElseIf mCat = MC_Answer Then
       CO = strmulanswer(Number)
       
End If


End Function

Private Sub Fill_in_the_Blanks_Paper1()
StrQuestions(1) = "To set the least significant bit, of an 8-bit register A to 0, we must carry out a logical dash  operation of A with the mask dash Hex."
StrQuestions(2) = "A dash structure is useful in the evaluation off postfix(reverse Polish) arithmetic expressions."
StrQuestions(3) = "Content addressable memory is useful in implementing the dash in an associative cache."
StrQuestions(4) = "In a 4-bit reflected grey code sequence, the code that immediately follows 0101 is dash."
StrQuestions(5) = "Synchronous circuits are characterized by the use of dash."
StrQuestions(6) = "In the 8-bit one’s complement representation, dash Hex  and dash Hex are both representations for zero."
StrQuestions(7) = "When Booth’s multiplication algorithm is used for computing 45 × 36 (both decimal numbers) with 45 as the multiplicand ,the number of additions required, is dash and the number of subtractions is dash."
StrQuestions(8) = "If the CPU must not honor any other interrupts while executing an interrupt service subroutine, then a suitable software solution, to the problem is to dash."
StrQuestions(9) = "A handshake based protocol for data transfer is an example of dash type of data transfer."
StrQuestions(10) = "A memory that loses its contents upon power down is known as dash memory."

stroptions(1) = "Clock signal"
stroptions(2) = "00,FF"
stroptions(3) = "Asynchronous"
stroptions(4) = "0111"
stroptions(5) = "Volatile"
stroptions(6) = "Tag directory"
stroptions(7) = "Stack"
stroptions(8) = "AND,01H"
stroptions(9) = "2,2"
stroptions(10) = "Disable Interrupts"

stranswers(1) = "AND,01H"
stranswers(2) = "Stack"
stranswers(3) = "Tag directory"
stranswers(4) = "0111"
stranswers(5) = "Clock signal"
stranswers(6) = "00,FF"
stranswers(7) = "2,2"
stranswers(8) = "Disable Interrupts"
stranswers(9) = "Asynchronous"
stranswers(10) = "Volatile"

End Sub
Private Sub Fill_in_the_Blanks_Paper2()
StrQuestions(1) = "The performance of an instruction pipeline is adversely affected by dash  instructions."
StrQuestions(2) = "In an 8-bit multiplication using Booth’s algorithm, dash number of additions/subtractions are required when the multiplier is 85 decimal."
StrQuestions(3) = "An 8-bit register contains the value -35 decimal in the signed-magnitude format. The contents of the register in Hexadecimal are dash."
StrQuestions(4) = "A combinational circuit takes a 32-bit number as input and gives as output the number of zero bits in the input number. The number of output bits in the circuit is dash."
StrQuestions(5) = "The smallest number of JK flip-flop required, to implement a ring counter with 16 outputs , is dash."
StrQuestions(6) = "An EPROM, with 16 bit address bus and 4-bit data bus, contains dash bits."
StrQuestions(7) = "A CPU requests data from memory by sending a 1 on the MEMREQ signal and a 1 on the READ signal. When the data read operation is completed, the memory raises a signal called MFC(Memory Function Complete). This ode of data transfer is dash."
StrQuestions(8) = "Data in registers R1 and R2 must be added and the sum should be stored in register R3. The instruction set of the CPU contains two address instructions ADD and MOV. The number of instructions, required to complete the operation, is dash."
StrQuestions(9) = "Daisy chaining is a method to implement dash."
StrQuestions(10) = "An interrupt from a device is ignored by the CPU, if the corresponding dash bit, is set to 0."

stroptions(1) = "256 K"
stroptions(2) = "5"
stroptions(3) = "Asynchronous"
stroptions(4) = "A3"
stroptions(5) = "Branch"
stroptions(6) = "Interrupt Priority"
stroptions(7) = "4"
stroptions(8) = "Mask"
stroptions(9) = "8"
stroptions(10) = "2"

stranswers(1) = "Branch"
stranswers(2) = "8"
stranswers(3) = "A3"
stranswers(4) = "5"
stranswers(5) = "4"
stranswers(6) = "256 K"
stranswers(7) = "Asynchronous"
stranswers(8) = "2"
stranswers(9) = "Interrupt Priority"
stranswers(10) = "Mask"
End Sub

Private Sub Fill_in_the_Blanks_Paper3()
StrQuestions(1) = "The concept of a stored program computer refers to the fact that dash."
StrQuestions(2) = "A self recursive subroutine is one which dash."
StrQuestions(3) = "A computer, with a 32 bit address bus, can be connected to dash MB of memory."
StrQuestions(4) = "A CPU, whose program counter has 24 bits can address at least dash words of storage."
StrQuestions(5) = "A floating point representation is using an 8 bit, Excess 128 exponent. When the value 2.5 is represented in normalize format, along with a hidden bit, the value of the exponent in binary is dash."
StrQuestions(6) = "In a compiler, syntax errors are checked by dash."
StrQuestions(7) = "In a compiler syntax errors are checked by dash."
StrQuestions(8) = "In a program variable X is defined as an integer and variable Y is defined as an array of integers. If the programmer uses the statement “X = Y”, then there has been a dash error."
StrQuestions(9) = "A compiler cannot detect dash type of errors made by the programmer."
StrQuestions(10) = "When an interrupt occurs, the CPU saves the program counter on stack before branching to the dash."

stroptions(1) = "Parser"
stroptions(2) = "Semantic"
stroptions(3) = "Logical"
stroptions(4) = "Both data and memory are stored in the same memory"
stroptions(5) = "Lexical analyzer"
stroptions(6) = "81H"
stroptions(7) = "Interrupt service subroutine"
stroptions(8) = "4096"
stroptions(9) = "16M"
stroptions(10) = "Calls itself"

stranswers(1) = "Both data and memory are stored in the same memory"
stranswers(2) = "Calls itself"
stranswers(3) = "4096"
stranswers(4) = "16M"
stranswers(5) = "81H"
stranswers(6) = "Parser"
stranswers(7) = "Lexical analyzer"
stranswers(8) = "Semantic"
stranswers(9) = "Logical"
stranswers(10) = "Interrupt service subroutine"
End Sub

Private Sub Fill_in_the_Blanks_Paper4()
StrQuestions(1) = "Two columns in a Microprogram can be merged into a single column if the combination dash does not appear in the same row of the two columns."
StrQuestions(2) = "In a daisy chain implementation of interrupt priority, the interrupt requests of all the devices are dash, before connecting to the interrupt signal of the CPU."
StrQuestions(3) = "In a two level memory hierarchy, the cache hit ratio must be dash so that the average access time is 10 ns. It is known that cache memory access time is 8 ns and main memory access time is 40 ns."
StrQuestions(4) = "Every interrupt service subroutine must end with a dash instruction."
StrQuestions(5) = "In a computer program, it is necessary to use a constant, to store the Avogadro number. The constant must be declared of type dash."
StrQuestions(6) = "A PU, with a memory address register of 28 bits, can address up to dash MB of memory."
StrQuestions(7) = "A digital computer, in which data memory is separate from instruction memory, is said to have dash architecture."
StrQuestions(8) = "The range of integers, that can be represented using 12 bit two’s complement representation, is to dash."
StrQuestions(9) = "In a computer, where I/O operations can also be performed using LOAD and STORE instructions, uses the dash scheme to connect I/O devices to the CPU."
StrQuestions(10) = "If it is desired that, no other interrupts be honored when an interrupt X is being serviced, then the CPU can execute a dash instruction at the beginning of the Interrupt Service Subroutine."

stroptions(1) = "Floating point"
stroptions(2) = "Return"
stroptions(3) = "Harvard"
stroptions(4) = "-2048 to +2047"
stroptions(5) = "ORed"
stroptions(6) = "11"
stroptions(7) = "Disable interrupts, Interrupt service subroutine"
stroptions(8) = "256"
stroptions(9) = "0.9"
stroptions(10) = "Memory Mapped I/O"

stranswers(1) = "11"
stranswers(2) = "ORed"
stranswers(3) = "0.9"
stranswers(4) = "Return"
stranswers(5) = "Floating point"
stranswers(6) = "256"
stranswers(7) = "Harvard"
stranswers(8) = "-2048 to +2047"
stranswers(9) = "Memory Mapped I/O"
stranswers(10) = "Disable interrupts, Interrupt service subroutine"

End Sub

Private Sub Fill_in_the_Blanks_Paper5()
StrQuestions(1) = "When input/output is performed, using instructions such as IN and OUT which are meant for I/O devices, the mapping scheme used is called dash."
StrQuestions(2) = "The number of 456 × 4 RAM chips required to construct a 2 KB cache is dash."
StrQuestions(3) = "Address lines and data lines are sometimes multiplexed in order to restrict the dash."
StrQuestions(4) = "If the register, which has been named in an instruction, contains the address of the actual at, we say the addressing mode is dash."
StrQuestions(5) = "When the addition of two positive numbers results in a negative value, the dash flag will be set."
StrQuestions(6) = "In a grey code sequence , the dash between any two successive numbers is 1."
StrQuestions(7) = "The kernel of an operating system in likely to be written in dash language of the underlying processor in order to improve the performance."
StrQuestions(8) = "A queue is a first in dash out structure."
StrQuestions(9) = "The dash addressing mode is ideally suited for addressing array elements."
StrQuestions(10) = "When an 8-bit two’s complement number is to be right shifted, we must set bit (MSB) to dash."

stroptions(1) = "Hamming Distance"
stroptions(2) = "Assembly"
stroptions(3) = "Previous value of Bit 7"
stroptions(4) = "Overflow"
stroptions(5) = "I/O Mapped I/O"
stroptions(6) = "I/O pins"
stroptions(7) = "16"
stroptions(8) = "First"
stroptions(9) = "Indexed"
stroptions(10) = "Indirect"

stranswers(1) = "I/O Mapped I/O"
stranswers(2) = "16"
stranswers(3) = "I/O pins"
stranswers(4) = "Indirect"
stranswers(5) = "Overflow"
stranswers(6) = "Hamming Distance"
stranswers(7) = "Assembly"
stranswers(8) = "First"
stranswers(9) = "Indexed"
stranswers(10) = "Previous value of Bit 7"
End Sub

Private Sub Fill_in_the_Blanks_Paper6()
StrQuestions(1) = "The dash bit is used in a write back cache to avoid unnecessary updates of main memory pages."
StrQuestions(2) = "The minimum number of bits, required to represent toe number -53 using 2’2 complement notation is dash."
StrQuestions(3) = "The sum output in a full adder, with inputs A,B, and C, is given by the Boolean expression dash."
StrQuestions(4) = "The smallest negative integer that can be  represented in 8-bit two’s complement notation is dash."
StrQuestions(5) = "The amount  of storage required, in a horizontally organized microstore, is dash than the amount of storage required in a vertically organized microstore."
StrQuestions(6) = "A hardwired control unit is dash , when compared to a microprogrammed control unit."
StrQuestions(7) = "In early microprocessors, designers used the multiplexing of address and data lines, in order to save the number of dash."
StrQuestions(8) = "First in first out is  a dash  policy in a fully associative cache."
StrQuestions(9) = "Taking the dash of a number and adding 1, will give us the two’s complement of the original number."
StrQuestions(10) = "Y is the two’s complement of a number X. The two’s complement of Y will give us dash."

stroptions(1) = "Faster"
stroptions(2) = "I/O pins"
stroptions(3) = "X"
stroptions(4) = "Replacement"
stroptions(5) = "More"
stroptions(6) = "Dirty"
stroptions(7) = "6"
stroptions(8) = "ABC"
stroptions(9) = "One’s complement"
stroptions(10) = "-128"

stranswers(1) = "Dirty"
stranswers(2) = "6"
stranswers(3) = "A?B?C"
stranswers(4) = "-128"
stranswers(5) = "More"
stranswers(6) = "Faster"
stranswers(7) = "I/O pins"
stranswers(8) = "Replacement"
stranswers(9) = "One’s complement"
stranswers(10) = "X"
End Sub

Private Sub Fill_in_the_Blanks_Paper7()
StrQuestions(1) = "An instruction that fetches the top of the stack is dash."
StrQuestions(2) = "In a half adder circuit SUM is realized using dash operation."
StrQuestions(3) = "There is no need of separate I/O instructions in case of dash."
StrQuestions(4) = "A CALL instruction is dash addressing mode."
StrQuestions(5) = "If the most significant bit is changed when two numbers are added in signed magnitude method, this is called dash."
StrQuestions(6) = "If Accumulator contains 10100011.execution of arithmetic shift instruction once on this data, results in dash."
StrQuestions(7) = "The maximum numbers of microoperations possible, when the microsinstruction format contains 3 fields, each of 4 bits in length, are dash."
StrQuestions(8) = "The numbers of address lines used, to address I/O in case of 8086 microprocessor, are dash."
StrQuestions(9) = "The process of accessing information on CD-ROM is dash."
StrQuestions(10) = "The time required to position R/W head, on appropriate cylinder, is called dash."

stroptions(1) = "Semi random"
stroptions(2) = "EX-OR"
stroptions(3) = "11010001"
stroptions(4) = "16"
stroptions(5) = "3"
stroptions(6) = "Overflow"
stroptions(7) = "I/O mapped I/O"
stroptions(8) = "PUSH"
stroptions(9) = "Implied stack"
stroptions(10) = "Seek Time"

stranswers(1) = "PUSH"
stranswers(2) = "EX-OR"
stranswers(3) = "I/O mapped I/O"
stranswers(4) = "Implied stack"
stranswers(5) = "Overflow"
stranswers(6) = "11010001"
stranswers(7) = "3"
stranswers(8) = "16"
stranswers(9) = "Semi random"
stranswers(10) = "Seek time"
End Sub

Private Sub Fill_in_the_Blanks_Paper8()
StrQuestions(1) = "A register that keeps track of the next instruction to be executed is dash."
StrQuestions(2) = "The cache memory concept is based on the principle of dash."
StrQuestions(3) = "Booth algorithm gives a procedure for multiplying two binary integers in  dash representation."
StrQuestions(4) = "In a six segment pipe lined processor, number of clocks required to process 120 tasks are dash."
StrQuestions(5) = "In an assembly language program, ORG is a dash type of instruction."
StrQuestions(6) = "An associative memory is also called dash addressable memory."
StrQuestions(7) = "A dash is called universal combinational logic circuit."
StrQuestions(8) = "Excess -3 code is a dash code."
StrQuestions(9) = "dash architecture has separate program and data memories."
StrQuestions(10) = "The hamming distance between 101101 and 110010 is dash."

stroptions(1) = "5"
stroptions(2) = "Multiplexer"
stroptions(3) = "Content"
stroptions(4) = "Self complimentary"
stroptions(5) = "Harvard"
stroptions(6) = "Locality of reference"
stroptions(7) = "Signed 2’s complement"
stroptions(8) = "Program counter"
stroptions(9) = "Pseudo"
stroptions(10) = "125"

stranswers(1) = "Program counter"
stranswers(2) = "Locality of reference"
stranswers(3) = "Signed 2’s complement"
stranswers(4) = "125"
stranswers(5) = "Pseudo"
stranswers(6) = "Content"
stranswers(7) = "Multiplexer"
stranswers(8) = "Self complimentary"
stranswers(9) = "Harvard"
stranswers(10) = "5"
End Sub

Private Sub Fill_in_the_Blanks_Paper9()
StrQuestions(1) = "dash construction can be used to clear a word in accumulator, when there is no CLEAR instruction."
StrQuestions(2) = "In dash addressing mode, the operand itself is contained in an instruction dash"
StrQuestions(3) = "A Boolean expression ABC + ABCD + AB + AC/B can be simplified as dash."
StrQuestions(4) = "Shifting data in  a register , in the left direction by one bit, is equivalent to dash by two."
StrQuestions(5) = "The memory that appears between CPU and main memory is dash."
StrQuestions(6) = "A decimal number 13 is represented in BBCD as dash."
StrQuestions(7) = "An instruction cycle consists of dash and execute cycles."
StrQuestions(8) = "The address selection part, of a microprogrammed part, is called dash."
StrQuestions(9) = "8086 microprocessor can prefetch bytes of program code from main memory."
StrQuestions(10) = "The type of RAM used CACHE is dash."

stroptions(1) = "Opcode fetch"
stroptions(2) = "AB"
stroptions(3) = "00010011"
stroptions(4) = "Multiply"
stroptions(5) = "Cache"
stroptions(6) = "XOR"
stroptions(7) = "Immediate"
stroptions(8) = "Microprogram sequence"
stroptions(9) = "6"
stroptions(10) = "SRAM"

stranswers(1) = "XOR"
stranswers(2) = "Immediate"
stranswers(3) = "AB"
stranswers(4) = "Multiply"
stranswers(5) = "Cache"
stranswers(6) = "00010011"
stranswers(7) = "Opcode fetch"
stranswers(8) = "Microprogram sequence"
stranswers(9) = "6"
stranswers(10) = "SRAM"
End Sub

Private Sub Fill_in_the_Blanks_Paper10()
StrQuestions(1) = "The output of a 3-state buffer is in dash state, when its enable input is not selected."
StrQuestions(2) = "In general, stack is a dash list."
StrQuestions(3) = "Stored program computer is also called dash."
StrQuestions(4) = "25.125(base 10) is represented in the binary form as dash."
StrQuestions(5) = "RISC architectures can achieve execution each segment in dash clock cycle."
StrQuestions(6) = "An instruction that increments program counter and does nothing is dash."
StrQuestions(7) = "The number of address lines in 8086 microprocessor are dash."
StrQuestions(8) = "A system software that converts high level language into object code is a dash."
StrQuestions(9) = "In a 1’s complement subtraction method, a carry resulted in most significant bit column, is dash."
StrQuestions(10) = "The DMA transfer technique, where transfer of one data word at a time, is called dash."

stroptions(1) = "Single"
stroptions(2) = "Cycle stealing"
stroptions(3) = "LIFO"
stroptions(4) = "NOP"
stroptions(5) = "11001.001"
stroptions(6) = "High impedance"
stroptions(7) = "Von-Neumann"
stroptions(8) = "20"
stroptions(9) = "Compiler"
stroptions(10) = "Added to least significant column"

stranswers(1) = "High impedance"
stranswers(2) = "LIFO"
stranswers(3) = "Von-Neumann"
stranswers(4) = "11001.001"
stranswers(5) = "Single"
stranswers(6) = "NOP"
stranswers(7) = "20"
stranswers(8) = "Compiler"
stranswers(9) = "Added to least significant column"
stranswers(10) = "Cycle stealing"
End Sub

Private Sub Matching_Columns_Paper1()
strcolumnA(1) = "PUSH"
strcolumnA(2) = "Microprogram"
strcolumnA(3) = "Macro"
strcolumnA(4) = "Two’s complement"
strcolumnA(5) = "One’s complement"
strcolumnA(6) = "NOP"
strcolumnA(7) = "BRANCH"
strcolumnA(8) = "LRU"
strcolumnA(9) = "QUEUE"
strcolumnA(10) = "STACK"

strcolumnB(1) = "Zero address instruction"
strcolumnB(2) = "One address instruction"
strcolumnB(3) = "One address instruction"
strcolumnB(4) = "Control sequence"
strcolumnB(5) = "Single representation for zero"
strcolumnB(6) = "Assembly time expansion"
strcolumnB(7) = "Bit inversion"
strcolumnB(8) = "Page replacement"
strcolumnB(9) = "LIFO"
strcolumnB(10) = "FIFO"

stranswers(1) = "One address instruction"
stranswers(2) = "Control sequence"
stranswers(3) = "Assembly time expansion"
stranswers(4) = "Single representation for zero"
stranswers(5) = "Bit inversion"
stranswers(6) = "Zero address instruction"
stranswers(7) = "One address instruction"
stranswers(8) = "Page replacement"
stranswers(9) = "FIFO"
stranswers(10) = "LIFO"
End Sub


Private Sub Matching_Columns_Paper2()
strcolumnA(1) = "Set the MSB of 8-bit register to 1."
strcolumnA(2) = "Set the LSB of 8-bit resister to 0."
strcolumnA(3) = "Test if 8-bit number is negative."
strcolumnA(4) = "Test if 8-bit number is odd."
strcolumnA(5) = "Does not require parentheses."
strcolumnA(6) = "Requires periodic refreshing"
strcolumnA(7) = "Associative search"
strcolumnA(8) = "Array access"
strcolumnA(9) = "Pointer access"
strcolumnA(10) = "Loading constants"

strcolumnB(1) = "Dynamic RAM"
strcolumnB(2) = "OR with mask 80H"
strcolumnB(3) = "AND with mask 80H"
strcolumnB(4) = "AND with mask 01H"
strcolumnB(5) = "AND with mask FEH"
strcolumnB(6) = "Infix"
strcolumnB(7) = "Content addressable RAM"
strcolumnB(8) = "Index addressing"
strcolumnB(9) = "Indirect addressing Immediate addressing"
strcolumnB(10) = "Immediate Addressing"

stranswers(1) = "OR with mask 80H"
stranswers(2) = "AND with mask FEH"
stranswers(3) = "AND with mask 80H"
stranswers(4) = "AND with mask 01H"
stranswers(5) = "Infix"
stranswers(6) = "Dynamic RAM"
stranswers(7) = "Content addressable RAM"
stranswers(8) = "Index addressing"
stranswers(9) = "Indirect addressing Immediate addressing"
stranswers(10) = "Immediate Addressing"
End Sub

Private Sub Matching_Columns_Paper3()
strcolumnA(1) = "Hierarchical memory"
strcolumnA(2) = "Dynamic memory"
strcolumnA(3) = "Cache memory"
strcolumnA(4) = "Magnetic disk memory"
strcolumnA(5) = "Associative memory"
strcolumnA(6) = "Ferrite core memory"
strcolumnA(7) = "4 MB"
strcolumnA(8) = "½ kb  × 8 EPROM"
strcolumnA(9) = "Tag directory"
strcolumnA(10) = "0000 – 03FF"

strcolumnB(1) = "Secondary storage"
strcolumnB(2) = "Content addressable storage"
strcolumnB(3) = "Primary storage"
strcolumnB(4) = "Memory refresh"
strcolumnB(5) = "Locality of reference"
strcolumnB(6) = "Static RAM"
strcolumnB(7) = "Associative cache"
strcolumnB(8) = "22 bits"
strcolumnB(9) = "1024 bits"
strcolumnB(10) = "4096 bits"

stranswers(1) = "Locality of reference"
stranswers(2) = "Memory refresh"
stranswers(3) = "Static RAM"
stranswers(4) = "Secondary storage"
stranswers(5) = "Content addressable storage"
stranswers(6) = "Primary storage"
stranswers(7) = "22 bits"
stranswers(8) = "4096 bits"
stranswers(9) = "Associative cache"
stranswers(10) = "1024 bits"

End Sub

Private Sub Matching_Columns_Paper4()
strcolumnA(1) = "Carry generate & propagate signals"
strcolumnA(2) = "Cascade of full adders"
strcolumnA(3) = "-128 to +127"
strcolumnA(4) = "-512 to + 511"
strcolumnA(5) = "-127 to +127"
strcolumnA(6) = "-256 to +256"
strcolumnA(7) = "Carry output of a half adder"
strcolumnA(8) = "Sum output of a full adder"
strcolumnA(9) = "Propagate signal of a carry lookahead adder"
strcolumnA(10) = "Both 000 octal and 777 octal are representations for 0"

strcolumnB(1) = "Logical AND"
strcolumnB(2) = "Logical OR"
strcolumnB(3) = "Carry lookahead adder"
strcolumnB(4) = "Ripple carry adder"
strcolumnB(5) = "Logical XOR"
strcolumnB(6) = "8-bit signed magnitude"
strcolumnB(7) = "8-bit two’s complement"
strcolumnB(8) = "9-bit two’s complement"
strcolumnB(9) = "10-bit two’s complement"
strcolumnB(10) = "10-bit two’s complement"

stranswers(1) = "Carry lookahead adder"
stranswers(2) = "Ripple carry adder"
stranswers(3) = "8-bit two’s complement"
stranswers(4) = "10-bit two’s complement"
stranswers(5) = "8-bit signed magnitude"
stranswers(6) = "10-bit two’s complement"
stranswers(7) = "Logical AND"
stranswers(8) = "Logical XOR"
stranswers(9) = "Logical OR"
stranswers(10) = "9-bit two’s complement"
End Sub

Private Sub Matching_Columns_Paper5()
strcolumnA(1) = "-23 decimal"
strcolumnA(2) = "32 decimal"
strcolumnA(3) = "23 decimal"
strcolumnA(4) = "32 hexadecimal"
strcolumnA(5) = "45 decimal"
strcolumnA(6) = "Excess 128 format"
strcolumnA(7) = "80 Hexadecimal"
strcolumnA(8) = "FF Hexadecimal"
strcolumnA(9) = "Number of possible functions of 2 inputs"
strcolumnA(10) = "Two's complement of -43 decimal"

strcolumnB(1) = "00100011 BCD"
strcolumnB(2) = "16 decimal"
strcolumnB(3) = "50 decimal"
strcolumnB(4) = "E9 Hexadecimal"
strcolumnB(5) = "231 base 4"
strcolumnB(6) = "40 octal"
strcolumnB(7) = "8-bit exponent"
strcolumnB(8) = "Zero"
strcolumnB(9) = "-1"
strcolumnB(10) = "00110101 binary"

stranswers(1) = "E9 Hexadecimal"
stranswers(2) = "40 octal"
stranswers(3) = "00100011 BCD"
stranswers(4) = "50 decimal"
stranswers(5) = "231 base 4"
stranswers(6) = "8-bit exponent"
stranswers(7) = "Zero"
stranswers(8) = "-1"
stranswers(9) = "16 decimal"
stranswers(10) = "00110101 binary"

End Sub

Private Sub Matching_Columns_Paper6()
strcolumnA(1) = "Asynchronous communication"
strcolumnA(2) = "Flip Flop"
strcolumnA(3) = "Scratch pad memory"
strcolumnA(4) = "RISC"
strcolumnA(5) = "IO mapped IO"
strcolumnA(6) = "Harvard architecture"
strcolumnA(7) = "Stack addressing"
strcolumnA(8) = "Karnaugh map"
strcolumnA(9) = "Nano memory"
strcolumnA(10) = "Hard disk drive"

strcolumnB(1) = "Second level of control memory"
strcolumnB(2) = "Separate input and output instructions"
strcolumnB(3) = "PUSH and POP"
strcolumnB(4) = "Simplify Boolean function"
strcolumnB(5) = "UART"
strcolumnB(6) = "One instruction per clock cycle"
strcolumnB(7) = "Digital signal processor TMS320C30"
strcolumnB(8) = "Bistable multivibrator"
strcolumnB(9) = "Secondary memory"
strcolumnB(10) = "Set of registers"

stranswers(1) = "UART"
stranswers(2) = "Bistable multivibrator"
stranswers(3) = "Set of registers"
stranswers(4) = "One instruction per clock cycle"
stranswers(5) = "Separate input and output instructions"
stranswers(6) = "Digital signal processor TMS320C30"
stranswers(7) = "PUSH and POP"
stranswers(8) = "Simplify Boolean function"
stranswers(9) = "Second level of control memory"
stranswers(10) = "Secondary memory"

End Sub

Private Sub Matching_Columns_Paper7()
strcolumnA(1) = "Third generation computers"
strcolumnA(2) = "Universal logic gates"
strcolumnA(3) = "Racing problems do not exist"
strcolumnA(4) = "Divided by 10"
strcolumnA(5) = "Synchronous communication"
strcolumnA(6) = "LRU"
strcolumnA(7) = "LRU"
strcolumnA(8) = "Multiplication"
strcolumnA(9) = "Division"
strcolumnA(10) = "Associative memory"

strcolumnB(1) = "NAND and NOR"
strcolumnB(2) = "Decade counter"
strcolumnB(3) = "SSI/MSI ICs"
strcolumnB(4) = "Master/Slave JK Flip-flop"
strcolumnB(5) = "Block mode transmission"
strcolumnB(6) = "CAM"
strcolumnB(7) = "ORG"
strcolumnB(8) = "Booth algorithm"
strcolumnB(9) = "Restoration algorithm"
strcolumnB(10) = "Page replacement policy"

stranswers(1) = "SSI/MSI ICs"
stranswers(2) = "NAND and NOR"
stranswers(3) = "Master/Slave JK Flip-flop"
stranswers(4) = "Decade counter"
stranswers(5) = "Block mode transmission"
stranswers(6) = "Page replacement policy"
stranswers(7) = "ORG"
stranswers(8) = "Booth algorithm"
stranswers(9) = "Restoration algorithm"
stranswers(10) = "CAM"

End Sub

Private Sub Matching_Columns_Paper8()
strcolumnA(1) = "R1 <- R2 + R3"
strcolumnA(2) = "An instruction to provide small delays"
strcolumnA(3) = "Operating system"
strcolumnA(4) = "BCD code of decimal number 15"
strcolumnA(5) = "Storage element consisting of two cross coupled NOR gates"
strcolumnA(6) = "High speed memory between CPU and main memory"
strcolumnA(7) = "Nibble"
strcolumnA(8) = "Hamming code"
strcolumnA(9) = "A circuit with one input and many outputs"
strcolumnA(10) = "A single chip computer"

strcolumnB(1) = "00010101"
strcolumnB(2) = "Flip flop"
strcolumnB(3) = "Error detecting and correcting code"
strcolumnB(4) = "Demultiplexer"
strcolumnB(5) = "UNIX"
strcolumnB(6) = "NOP"
strcolumnB(7) = "Register transfer"
strcolumnB(8) = "Micro controller"
strcolumnB(9) = "Microprocessor"
strcolumnB(10) = "Cache memory"

stranswers(1) = "Register transfer"
stranswers(2) = "NOP"
stranswers(3) = "UNIX"
stranswers(4) = "00010101"
stranswers(5) = "Flip flop"
stranswers(6) = "Cache Memory"
stranswers(7) = "Micro controller"
stranswers(8) = "Error detecting and correcting code"
stranswers(9) = "Demultiplexer"
stranswers(10) = "Microprocessor"
End Sub

Private Sub Matching_Columns_Paper9()
strcolumnA(1) = "Instruction register"
strcolumnA(2) = "Correction of 6 "
strcolumnA(3) = "( a b )' = a' + b'"
strcolumnA(4) = "FIFO"
strcolumnA(5) = "DOS hidden file"
strcolumnA(6) = "Primary memory"
strcolumnA(7) = "Capability to process more than one program concurrently"
strcolumnA(8) = "Non volatile memory"
strcolumnA(9) = "Output from printer"
strcolumnA(10) = "Relative addressing"

strcolumnB(1) = "Hard copy"
strcolumnB(2) = "RAM"
strcolumnB(3) = "Uses base register"
strcolumnB(4) = "Multi programming"
strcolumnB(5) = "De-Morgan's theorem"
strcolumnB(6) = "IO.SYS"
strcolumnB(7) = "ROM"
strcolumnB(8) = "BCD addition"
strcolumnB(9) = "Page replacement policy"
strcolumnB(10) = "Operation code"

stranswers(1) = "Operation code"
stranswers(2) = "BCD addition"
stranswers(3) = "De-Morgan's theorem"
stranswers(4) = "Page replacement policy"
stranswers(5) = "IO.SYS"
stranswers(6) = "Multi programming"
stranswers(7) = "RAM"
stranswers(8) = "ROM"
stranswers(9) = "Hard copy"
stranswers(10) = "Uses base register"

End Sub

Private Sub Matching_Columns_Paper10()
strcolumnA(1) = "Address is an operand"
strcolumnA(2) = "Memory requires refresh"
strcolumnA(3) = "Bidirectional read/write control"
strcolumnA(4) = "Program control transfer"
strcolumnA(5) = "Horizontal format"
strcolumnA(6) = "Segment registers"
strcolumnA(7) = "Clock generator"
strcolumnA(8) = "Seek time"
strcolumnA(9) = "State table"
strcolumnA(10) = "Cache memory"

strcolumnB(1) = "A stable multivibrator"
strcolumnB(2) = "8086 CPU"
strcolumnB(3) = "Sequential basic circuits"
strcolumnB(4) = "Hard disk drive"
strcolumnB(5) = "Control memory"
strcolumnB(6) = "DMA controller"
strcolumnB(7) = "DRAM"
strcolumnB(8) = "Locality of reference"
strcolumnB(9) = "Sub routing"
strcolumnB(10) = "Immediate addressing"

stranswers(1) = "Immediate addressing"
stranswers(2) = "DRAM"
stranswers(3) = "DMA controller"
stranswers(4) = "Sub routing"
stranswers(5) = "Control memory"
stranswers(6) = "8086 CPU"
stranswers(7) = "A stable multivibrator"
stranswers(8) = "Hard disk drive"
stranswers(9) = "Sequential basic circuits"
stranswers(10) = "Locality of reference"

End Sub

Private Sub TrueorFalse_Paper1()
StrQuestions(1) = "A dirty bit is associated with each block of a write back cache."
StrQuestions(2) = "The construction of a divider based on the non restoring division algorithm requires a subtractor and requires with shift right capability."
StrQuestions(3) = "In a Karnaugh map of four variables A,B,C and D, the term AB will cover a strip of 2 squares."
StrQuestions(4) = "The word firmware is used to describe a hardwired control unit."
StrQuestions(5) = "A linker generates he executable version of the program by binding seeral source programs."
StrQuestions(6) = "Macros are expanded during the time of source assembly."
StrQuestions(7) = "A finite state machine, with 255 states, can be implemented using 13 flip-flops."
StrQuestions(8) = "A 4-stage pipelined floating point adder can give a speedup greater than 4, if we endure that all stages require the same processing time."
StrQuestions(9) = "Wait cycles are unnecessary while executing al instruction in which only regiser addressing is employed."
StrQuestions(10) = "There are 4 Boolean functions of two Boolean variables."

stranswers(1) = "True"
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

Private Sub TrueorFalse_Paper2()
StrQuestions(1) = "In a fully associative cache memory, the entire cache memory is implemented as associated memory."
StrQuestions(2) = "The bit ratio of a 128 bit KB cache will be better than the hit ratio of a 64 KB cache."
StrQuestions(3) = "A static RAM does not require a refresh input."
StrQuestions(4) = "Vertically organized microcode requires lese amount of storage when compared to horizontally organized microcode."
StrQuestions(5) = "Vertically organized microcode is faster when compared to horizontally organized microcode."
StrQuestions(6) = "Hardwired controllers are faster when compared to microcontrollers."
StrQuestions(7) = "The number of additions and subtractions, required by Booth's multiplication algorithm, is less than or equal to  the number of additions and subtractions required in the conventional shift and add multiplication algorithm."
StrQuestions(8) = "In a carry look ahead adder, which adds two numbers A and B, the 'propagate' signal, at stage I, can be generated by either ORing or EXORing the inputs Ai and Bi."
StrQuestions(9) = "The width of the program counter of a CPU, which can address 200 MB of main memory, is at least 32 bits."
StrQuestions(10) = "The carry output off a full adder is set to 1, if at least two of the three inputs to the full adder are 1."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub
Private Sub TrueorFalse_Paper3()
StrQuestions(1) = "A 'stored program ' computer is one in which the instructions are fethched bth e CPU from a memory."
StrQuestions(2) = "A tag directory is unnecessary in direct mapped caches."
StrQuestions(3) = "A computer with a 32-bit address bus can be connected to 2048 MB of memory."
StrQuestions(4) = "A CPU whose program counter has 24 bits can address at least 16 Mwords of storage."
StrQuestions(5) = "A floating point representation is using an 8-bit , Excess-128 exponent. When the value 2.5 is represented in normalized format, with one hidden bit, the value of the exponent in 81 Hex."
StrQuestions(6) = "When the multiplier is 01(decimal), Booth's multiplication algorithm requires one addition and one subtraction."
StrQuestions(7) = "A queue is a first in first out structure."
StrQuestions(8) = "When a right shift operation is performed on two's complement numbers, zeros must be shifted from the left."
StrQuestions(9) = "The source program is separated into tokens by a parser."
StrQuestions(10) = "A programming error,  such as forgetting to declare a label in the program, cannot be detected by a single pass assembler."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub
Private Sub TrueorFalse_Paper4()
StrQuestions(1) = "An assembly language program, which uses macros, will be faster than a program which uses subroutines."
StrQuestions(2) = "An assembly language program, which uses macros, will consume less space compared to a program which uses subroutines."
StrQuestions(3) = "Unlike subroutines we cannot pass parameters to macros."
StrQuestions(4) = "Macros are expanded during program execution."
StrQuestions(5) = "Dynamic memories are good candidates for building the cache memory of a computer."
StrQuestions(6) = "When the multiplier is 127(decimal), Booth’s multiplication algorithm requires one addition operation."
StrQuestions(7) = "Stack is a first in first out structure."
StrQuestions(8) = "There are two representations for zero in 1’s complement number system."
StrQuestions(9) = "PUSH is a zero address instruction."
StrQuestions(10) = "In a memory mapped I/O address of a memory location is 40H and on I/O device also has the address of 40H."

stranswers(1) = "True"
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
Private Sub TrueorFalse_Paper5()
StrQuestions(1) = "Exclusive OR operation on the contents of a register with itself sets the carry flag of the status register."
StrQuestions(2) = "Exclusive OR operation on the contents of a register with itself sets the carry flag of the status register."
StrQuestions(3) = "A software interrupt is an instruction executed by the CPU."
StrQuestions(4) = "An instruction cycle always consists of one memory read operation."
StrQuestions(5) = "A horizontal format of the control word in a microprogrammmed control unit is faster."
StrQuestions(6) = "Magnetic drums provide faster access than disks because there is no rotational latency."
StrQuestions(7) = "Machine code is generated for instructions each time a subroutine is called in an assembly program."
StrQuestions(8) = "A 4-bit binary adder can be realized using three full adders and one half adder."
StrQuestions(9) = "When a DMA request is received from an I/O device. CPU completes the present machine cycle."
StrQuestions(10) = "When two eight bit numbers x and y represented in signed magnitude representation x = 0101 1000B and y = 01000101B are added, it results in and OVERFLOW."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub
Private Sub TrueorFalse_Paper6()
StrQuestions(1) = "In Boolean algebra, ABCD D is zero."
StrQuestions(2) = "he FANOUT of standard TTL gate is 10."
StrQuestions(3) = "A sixteen bit microprocessor means it has 16 address lines."
StrQuestions(4) = "A binary sequence 10100100B has an even parity."
StrQuestions(5) = "The zero status flag can be set to 1 by loading an accumulator register with zero."
StrQuestions(6) = "If a processor contains a prefetch buffer for instructions, Instruction register can get the next instruction to be executed, from the prefecth bugger always."
StrQuestions(7) = "In an asynchronous communication interface, hardware transmit and receive clocks can be different."
StrQuestions(8) = "In a vectored interrupt, the branch address is assigned to a fixed memory location."
StrQuestions(9) = "In a digital computer, ROM portion of the main memory is needed for storing an initial program called Bootstrap program."
StrQuestions(10) = "A page refers to the organization of address space, while a block refers to the organization of memory space."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub
Private Sub TrueorFalse_Paper7()
StrQuestions(1) = "In Von Neumann's computer, there is no distinction between data and instruction."
StrQuestions(2) = "Binary(12)10 and BCD(12)10 have same representation."
StrQuestions(3) = "In Boolean algebra w x + y x + w y can be reduced to w x + y x."
StrQuestions(4) = "Decoding of instruction takes place in several steps of the execution phase."
StrQuestions(5) = "Program counter counts the number of instructions executed at a particular instant of time."
StrQuestions(6) = "More MOS gates can be packed than Bipolar gates in the same space."
StrQuestions(7) = "Instruction Register and program counter are part of EU in an 8086 CPU."
StrQuestions(8) = "In an indexed addressing mode, index register contains the base address of a date array and the instruction contains the offset address."
StrQuestions(9) = "The use of second level of control memory called nanocontrol memory can reduce the total control memory size."
StrQuestions(10) = "When main memory is shared by multiple processors, the data inconsistency between main memory and cache memory can be eliminated by adapting to write through policy."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub
Private Sub TrueorFalse_Paper8()
StrQuestions(1) = "A location counter is used to place he machine code generated by the assembler in the memory."
StrQuestions(2) = "An stable multivibrator can be used as a flip flop."
StrQuestions(3) = "Decimal equivalent of the octal number 125.1 is 85.125"
StrQuestions(4) = "MAR keeps track of the next instruction to be executed."
StrQuestions(5) = "Single pass assembler is quite larger than two pass assembler."
StrQuestions(6) = "The set of microinstructions are stored in RAM."
StrQuestions(7) = "All vectored interrupts are hardware interrupts."
StrQuestions(8) = "In a computer with cache memory, the write back policy is superior to write through policy, with regard to the speed."
StrQuestions(9) = "When an interrupt occurs, the CPU completes he presently executing instruction and then transfers the program control ISR."
StrQuestions(10) = "A vertical format of the micro program control word is faster."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "True"
stranswers(10) = "False"
End Sub
Private Sub TrueorFalse_Paper9()
StrQuestions(1) = "Static RAM requires every bit to be refreshed every few milliseconds."
StrQuestions(2) = "Instruction register is not accessible to the assembly programs of almost all computers."
StrQuestions(3) = "Register reference instructions can be  executed faster than that of memory referenced instructions."
StrQuestions(4) = "A half duplex transmission system is one that is capable of transmitting in one direction only."
StrQuestions(5) = "With I/O mapped I/O all instructions that refer to memory are also available for I/O."
StrQuestions(6) = "Booth algorithm gives a procedure for multiplying binary integers in signed 2's complement representation."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "True"
End Sub

Private Sub questions_paper1()
StrQuestions(1) = "a.  Show the steps involved in applying the Booth's multiplication algorithm to the numbers 3 and 5. Assume 5 is the multiplier. Assume 4-bit representation." & Chr(13) & _
                  "b.  Design a combinational circuit which receives 4 inputs W,X,Y and Z & generates two outputs P and Q. P is set to 1 if at least two of the inputs re 1. Q is set to 0 if at least two of the inputs are 0. Give the sum of Products expressions for P and Q." & Chr(13) & _
                  "c.  A 16-bit computer is to be interfaced to 2 k words of EPROM memory and 30 k words of RAM. Given 1K X 8 EPROM chips and 2K X 8 RAM chips. Explain how you will design the memory subsystem. Show the address decoding and indicate the address range for the EPROMs and RAMs."
StrQuestions(2) = "a.  Distinguish between random access storage and sequential access storage devices. Give an example of each storage type. What are the typical applications of each storage type?" & Chr(13) & _
                  "b.  How many address lines are present in a 256 K X 8 RAM? How many such RAMs will be required to construct a 512K X 32 memory bank?" & Chr(13) & _
                  "c.  What is the motivation for hierarchical memory organization in modern computers? Explain why memory hierarch works successfully in von Neumann computers."
StrQuestions(3) = "a.  Explain the differences between an assembler and a compiler." & Chr(13) & _
                  "b.  What are the advantages and disadvantages of developing a soft ware package using" & Chr(13) & _
                  "i)  a high level language and" & Chr(13) & _
                  "ii) an assembly language" & Chr(13) & _
                  "c.  What is meant by the kernel of an operating system?" & Chr(13) & _
                  "d.  Explain how real time operating systems are different"
StrQuestions(4) = "a.  A floating point number system uses 16 bits for representing a number. The most significant bit is the sign bit. The least significant 9 bits represent the mantissa. The remaining 6 bits represent the exponent. Assume that numbers are stored in the normalized format with one hidden bit. Give the representation of -1.6 × 10(3 ) in this number system." & Chr(13) & _
                  "b.  In the above floating point number representation, what is the value representated by 0000100110000000 ?" & Chr(13) & _
                  "c.  What are the advantages of two's complement number system over one's complement number system?" & Chr(13) & _
                  "d.  Show an assembly language program, using the 8085 or the 8086 instruction set, to place the value -25 in a location X. Assume that two's complement representation is used."
StrQuestions(5) = "a.  Explain the meaning of the sentence : 8086 is a 16 bit microprocessor." & Chr(13) & _
                  "b.  What is the purpose of the stack pointer register in the 8086 CPU?" & Chr(13) & _
                  "c.  Explain the role of stacks in programming." & Chr(13) & _
                  "d.  In an 8-bit microcomputer(e.g. 8085), three memory locations, A,B , and C, store 8-bit values. Let A0,B0 and C0 denote the least significant bits of A,B, and C. Show an assembly language program to compute the value (A0.B0 + B0.C0 + C0.A0). Here, the '.' Operator stands for logical AND and the " + " operator stands for logical OR. The result of the computation should be stored in D0, where D0 is the least significant bit of location D."
StrQuestions(6) = "a.  Distinguish between the terms 'software interrupt' and 'non maskable interrupt'." & Chr(13) & _
                  "b.  What is meant by 'disabling interrupts' ? When are interrupts disabled?" & Chr(13) & _
                  "c.  What is meant by 'interrupt priority'?" & Chr(13) & _
                  "d.  What is meant by 'device polling'?" & Chr(13) & _
                  "e.  What is the role of an interrupt controller in a computer?"
StrQuestions(7) = "a.  The function f(w,x,y,z) on four Boolean variables w,x,y,z is defined as follows: If at least two of the inputs are zero, then f(w,x,y,z) is equal to 1; otherwise the function evaluates to zero. Show the Karnaugh map for the function and derive a minimal sum of products implementation of the function, using only two input NAND gates." & Chr(13) & _
                  "b.  Prove that any digital circuit  can be constructed using only 2-input XOR and 2-input AND gates."
StrQuestions(8) = "a.  Consider a full subtractor circuit which subtracts input Y and the borrow B(in) from input X; derive minimum SOP expressions for the outputs D(difference) and B(out) (output borrow) of the subtractor." & Chr(13) & _
                  "b.  Show how a 4-bit subtraction circuit can be constructed, using the circuit designed in part (a) above." & Chr(13) & _
                  "c.  Show how the D output of part (a) can be realized using a 4-to-1 multiplexer."
StrQuestions(9) = "Show the design of a 4-bit synchronous shift register which has the following capabilities : shift right, shift left, parallel load, and reset to 0. Assume the appropriate control inputs. Use D flip flops."
StrQuestions(10) = "a.  Using logic gates design a circuit which receives and 8- bit number X = X7,X6,X5,X4,X3,X2,X0 and generates the 9-bit number Y = X +1" & Chr(13) & _
                   "b. Design a combinational shifter which receives and 8- bit number X and a 3-bit  input Y; the output Z of the shifter is X shifter right by as many bits as the number Y."
End Sub

Private Sub questions_paper2()
StrQuestions(1) = "a.  Show a reflected 4-bit grey code sequence. Explain how the 5-bit grey code sequence can be constructed starting from the 4-bit grey code sequence" & Chr(13) & _
                  "b.  Show a binary to grey code conversion circuit assuming that the input is a 5-bit number" & Chr(13) & _
                  "c.  Given an 8-bit number X, show a circuit using logic gates to compute the two's complement of X."
StrQuestions(2) = "a.  Design a synchronous counter which counts the sequence 5,4,3,1,8,5,4,3,1,8,… Use JK flip flops and logic gates in you design." & Chr(13) & _
                  "b.  Design a circuit that coverts a 4-bit binary number into an equivalent BCD number. Assume that the input is represented in unsigned binary format." & Chr(13) & _
                  "c.  What is meant by a critical race? How is the problem of races circumvented in digital systems?"
StrQuestions(3) = "a.  Distinguish between on Neumann machines, Harvard architecture, and Turing machines." & Chr(13) & _
                  "b.  Distinguish between macros and subroutines" & Chr(13) & _
                  "c.  Distinguish between software interrupts and vector interrupts."
StrQuestions(4) = "The two level memory hierarchy of a computer consists of cache memory and main memory. The hit ratio of the hierarchical memory has been measured as 0.95. The CPU of the computer operates at 120 MHz." & Chr(13) & _
                  "a.  Derive an expression for the average access time of the hierarchical memory." & Chr(13) & _
                  "b.  What is the access time of the cache memory, assuming that the average access time is not more than the cycle time of the CPU?" & Chr(13) & _
                  "c.  Define the term 'Memory Bandwidth'. What is the memory bandwidth of the hierarchical memory in this example?" & Chr(13) & _
                  "d.  Assume that the cache is 128 KB in size and the main memory is 4 MB in size. The block size is 1 KB and direct mapping policy is used in the cache. The CPU is a 32 bit microprocessor. The CPU generates a write operation at memory location 00000100H. Assuming write back policy, explain how the cache controller will handle the write request. Similarly explain how a read request to location 00000100H will be handled."
StrQuestions(5) = "a.  Write a subroutine called FUNCX in some assembly language (e.g. 8085 or 8086). The subroutine expects two 8-bit numbers X and Y on the top of the stack when the sub routing is called. The unction of the subroutine is to compute(X+Y-1) and store it back on the stack." & Chr(13) & _
                  "b.  A program computes the value (a+b+1), where a is the value stored in the A register and b is the value stored in the B register. At the end of the computation, the resulting value is stored in C register. Write this program using the FUNCX subrouting."
StrQuestions(6) = "a.  The charge on an electron is 1.6 × 10(-19) coulombs. Represent this number in the 32 bit floating point representation. Assume that the most significant bit is the sign bit S. the least significant 23 bits are for the mantissa M. The remaining 8 bits are for the exponent E in Excess - 128 format. The mantissa is to stored in the normalized format, where the 1 to the immediate left of the binary point in the mantissa is not stored." & Chr(13) & _
                  "Give your answer in hexadecimal" & Chr(13) & _
                  "b.  A 32 bit register contains he value DACB0000(hexadecimal). If this value is to be interpreted as a floating point number in the above representation scheme, what is the number being represented?" & Chr(13) & _
                  "c.  Give the merits and demerits of the floating point and fixed point representations for storing real numbers."
StrQuestions(7) = "a.  Explain the difference between a macro and a subroutine. Give the relative merits of macros and subroutines in assembly language programming." & Chr(13) & _
                  "b.  Explain the difference between polling and interrupt drive I/O." & Chr(13) & _
                  "c.  Show the logic circuit for generating the look ahead carry signals in a 4-bit carry look ahead adder." & Chr(13) & _
                  "d.  What is meant by Booth's recoding of a multiplier? Recode the multiplier 198 (decimal) using Booth's recoding scheme. How many additions and subtractions will be required, when this recoding is used?"
StrQuestions(8) = "a.  Show the organization of a microprogrammed control unit and explain its operation." & Chr(13) & _
                  "b.  What is the difference between horizontal microcode organization and vertical microcode organization?" & Chr(13) & _
                  "c.  A microprogrammed controller is to be designed for heating 29 control signals. What is the size of the microstore? The size of the micro program counter is give to be 8 bits. It is known that there are 3 bits reserved in the microinstruction register for condition selection."
StrQuestions(9) = "a.  Explain the term locality of reference. What is its significance? Distinguish between temporal and spatial locality." & Chr(13) & _
                  "b.  Using a block diagram, explain how a virtual address is translated into a physical address in a computer which uses virtual memory." & Chr(13) & _
                  "c.  How do the following influence the performance of a virtual memory system ?" & Chr(13) & _
                  "i)  size of a page" & Chr(13) & _
                  "ii) Replacement policy" & Chr(13) & _
                       Chr(13) & _
                  "d.  A computer has a direct mapped cache with 16 one word blocks. The cache is initially empty. What is the observed hit ratio when the CPU generates the following word address sequence:" & Chr(13) & _
                  "1,4,8,5,20,17,19,56,9,11,4,43,5,6,9,17 ?"
StrQuestions(10) = "a.  A hard disk has a rotating speed of 4500 resolutions per minute. The seek time is 20 ms, and the data transfer rate is 2 MB per second. The disk is organized into 512 bytes sectors. What is the average time to read a sector?" & Chr(13) & _
                   "b.  Describe a typical asynchronous handshaking protocol used for data transfer between the CPU and an I/O device." & Chr(13) & _
                   "c.  What is the role of a bus arbiter in a computer system? Describe two bus arbitration strategies commonly used in computers?"
End Sub

Private Sub questions_paper3()
StrQuestions(1) = "a.  The Fibonacci number sequence is 0,1,1,2,3,4,8,13,… The sequence is defined by the recurrence Fib(N) = Fib (N-1) + Fib(N-2) and the initial conditions Fib(0) = 0 and Fib(1) = 1." & Chr(13) & _
                  "Show an assembly language program to compute Fib(N), given N." & Chr(13) & _
                  "Assume that the value of N is present in a memory location N. The final answer must be placed in a memory location FIBN. None that Fib(N) can be stored as an 8-bit number only up to some value of N. If the number specified in location N is larger than this value, you must write a 1 in the location ERROR and halr. Otherwise, compute Fib(N) store it in FIBN, and write a 0 in location ?ERROR and halt." & Chr(13) & _
                  "b.  What will be the bit pattern stored in FIBN and ERROR when you run your program for the following values of N?" & Chr(13) & _
                  "N = 09H, N = 0BH, N = 0FH and N = 10H" & Chr(13) & _
                  "c.  Explain how you can handle large values of N by suitably modifying your program."
StrQuestions(2) = "a.  What is the difference between application software and system software?" & Chr(13) & _
                  "b.  What is the name of a program which the computer executes as soon as it is powered on? Explain the functions of such a program." & Chr(13) & _
                  "c.  Classify the following programs as application software or system software:" & Chr(13) & _
                  "i)  Text editor" & Chr(13) & _
                  "ii) Sort/merge utility" & Chr(13) & _
                  "iii)    spell checker" & Chr(13) & _
                  "iv) Debugger" & Chr(13) & _
                  "v)  Compiler for Fortran" & Chr(13) & _
                  "vi) Desk Top publishing software" & Chr(13) & _
                  "vii)    A database management system" & Chr(13) & _
                  "viii)   Cross assembler" & Chr(13) & _
                  "ix) BASIC interpreter"
StrQuestions(3) = "a.  When we write an assembly language program, what are the four essential components that we include in every line? Explain the significance of each of these." & Chr(13) & _
                  "b.  What is meant by a macro? Compare the relative merits and demerits of using macros instead of subroutines in a program." & Chr(13) & _
                  "c.  What is meant by a symbol table? Explain its use."
StrQuestions(4) = "a.  Design a finite state machine which generates the sequence 001, 011, 101, 110, 001, 011, 101, 110… Use minimum hardware. Use JK flip flops as memory elements." & Chr(13) & _
                  "b.  Two 4-digit BCD numbers are stored in 16 bit registers A and B." & Chr(13) & _
                  "The sum of these BCD numbers must be generated and stored in a 16 bit register C. You have a 4-bit parallel adder available. Other auxiliary components, such as logic gates and registers are also available. Show a circuit that can accomplish the specified task. Explain its operation."
StrQuestions(5) = "a.  The Berger code of a number is the number of bits which are set to 1. for instance, the Berger  code of 0110 is 2, the Berger code of 1101 is 3, and so on. Design a circuit which computes the Berger code of an 8-bit number. You may use both combinational and sequential building blocks. Explain the operation of the circuit." & Chr(13) & _
                  "b.  An 8-bit number is available in register A of the CPU. Write a subrouting in an assembly language to computer the Berger code of the number and store it in register B. Provide adequate comments."
StrQuestions(6) = "a.  Explain the difference between combinational and sequential logic" & Chr(13) & _
                  "b.  Show the realization of D-FF from a J k flip flop and explain" & Chr(13) & _
                  "c.  Design a digital circuit that performs the four logic operations of exclusive-OR, exclusive-NOR, NOR and NAND. Use tow selection variables. Show he logic diagram of one typical stage."
StrQuestions(7) = "a.  What is the advantage of using a CPU register for temporary data storage over using a memory location?" & Chr(13) & _
                  "b.  The memory unit of a computer has 256 K words of 32 bits each. The computer has an instruction format with four fields: an operation code field; a mode field to specify one of seven addressing modes; a register address field to specify one of 60 processor registers; and a memory address. Specify the instruction format and the number of bits in each field if the instruction length is equal to one memory word."
StrQuestions(8) = "a.  What is the difference between zero address , one address and two address instructions? Illustrate with the help of examples." & Chr(13) & _
                  "b.  Write an assembly language segment which finds the frequency of digits of a message consisting of 35 numeric digits of MSG. the frequency through 9 are to be given at DIG0,DIG1, DIG@, DIG#, _____, DIG( respectively."
StrQuestions(9) = "a.  Give the schematic diagram of a microprogram sequence, explaining how it implements conditional branch and other functions required to be performed by the control unit of a digital computer." & Chr(13) & _
                  "b.  Describe horizontal and vertical micro code formats used in control unit design."
StrQuestions(10) = "a.  What is associative memory? Explain how it is used in address mapping in a cache memory system." & Chr(13) & _
                   "b.  Explain the macro assembler."
End Sub

Private Sub questions_paper4()
StrQuestions(1) = "a.  An n-bit digital multiplexer is a universal logic circuit. Explain." & Chr(13) & _
    "b.  Design a 4 bit shift register with the following capabilities using D-flip flops" & Chr(13) & _
    "i)  Serial in - Parallel out" & Chr(13) & _
    "ii) Parallel inn - parallel out" & Chr(13) & _
    "d.  Develop the state diagram of a 2-bit binary counter."
StrQuestions(2) = "a.  Explain the indexed addressing and Based addressing modes with examples." & Chr(13) & _
                  "b.  Assume the stack operations consume one time unit, register to register operations, one time unit; register/ stack to / from memory operations, four time units; and memory to memory operations, seen time units. Calculate the time units needed to execute the following high level language statement for zero address, one address, two memory address and three memory address, and multiple register computers. State any other assumptions made. A = (B +C)/(C-B) * D"
StrQuestions(3) = "Given 1K × 8 bits of RAM and ROM chips, show the circuit schematic to construct 16k RAM and 8k ROM to interface to a microprocessor having sixteen address lines. Assume suitable address spaces for ROM and RAM"
StrQuestions(4) = "a.  Distinguish between a control memory and main memory used in a digital computer." & Chr(13) & _
                  "b.  Describe the sequence of micro operations of an interrupt return cycle. Consider the interrupt to be a vectored interrupt."
StrQuestions(5) = "Write an assembly language program that defines a list of 10 readings starting at location MY DATA in the data segment. All these readings are byte wide. It is required to sum up all the non zero readings from this list and find out their average value. Store the results in locations NONZSUM and AG respectively, defined in the data segment"
StrQuestions(6) = "a.   Compare programmed I/O with DMA based I/O with respect to program design complexity, I/O band width and hardware requirements." & Chr(13) & _
                  "b.  Explain the basic principle of Booth's algorithm for multiplication of signed 2's complement numbers. Give the algorithm in flops chart."
StrQuestions(7) = "a.  Differentiate between cache memory and virtual memory." & Chr(13) & _
                  "b.  A digital computer has a memory unit of 64k × 16 and a cache memory of 512 words. The cache memory uses direct memory with a block size of 8 words." & Chr(13) & _
                  "i)  How many bits are there in the age, index, block and word fields of the address format." & Chr(13) & _
                  "ii) How may bits are there in each word of the cache and how are they grouped?" & Chr(13) & _
                  "iii)    How many blocks can the cache accommodate?"
StrQuestions(8) = "a.  Explain how a one pass assembler can account for forward references." & Chr(13) & _
                  "b.  Explain passing techniques employed in a compiler."
StrQuestions(9) = "a.  Design a 4-bit BCD adder to add two 4-bit numbers. Make sure that the result of addition is also a valid BCD number." & Chr(13) & _
                  "b.  Explain the fixed point binary number division using 'Restoring' method with the help of a flow chart"
StrQuestions(10) = "a.  Explain the difference between synchronous I/O and asynchronous I/O" & Chr(13) & _
                   "b.  How many characters per second, can be transmitted over a 2400- baud line, in each of the following modes (Assume a character code of 8-bits)" & Chr(13) & _
                   "i)  Synchronous serial transmission" & Chr(13) & _
                   "ii) Asynchronous serial transmission with one start and one stop bit" & Chr(13) & _
                   "iii)Asynchronous serial transmission with one start bit and two stop bits." & Chr(13) & _
                        Chr(13) & _
                   "d.  In most computers an interrupt is recognized only after the execution of the instruction. Consider the possibility of acknowledging the interrupt at any time during the interrupt at any time during the execution of the instruction. Discuss the difficulty that may arise."
End Sub

Private Sub multiplechoice_Paper1()
strmultiple(1) = "When an interrupt occurs, the processor completes the current ___________ before jumping to the interrupt service subroutine."
strmuloption(1) = "Microinstruction it is executing"
strmuloption(2) = "Instruction it is executing"
strmuloption(3) = "Current macro it is executing"
strmuloption(4) = "Procedure call it is executing"

strmultiple(2) = "We wish to test if the number present in the 8-bit register A is an odd number. For this purpose, we must -"
strmuloption(5) = "AND the a register with mask 01H and test if Z flag is set to 0."
strmuloption(6) = "OR the A register with mask 01H and test if Z flag is set to 0."
strmuloption(7) = "AND the A register with mask 0H and test if Z flag is set to ."
strmuloption(8) = "R the A register with mask 01H and test if Z flag is set to ."

strmultiple(3) = "Parenthesis are required to indicate the order of evaluation of arithmetic operations in -"
strmuloption(9) = "Prefix expressions"
strmuloption(10) = "Infix expressions"
strmuloption(11) = "Postfix expressions."
strmuloption(12) = "Both a and c"

strmultiple(4) = "An 8-bit carry look ahead adder that adds two 8-bit numbers X and Y and a carry input C in can be used to subtract an 8-bit number B from an 8-bit number A by setting -"
strmuloption(13) = "X = A, Y = B, Cin = 1"
strmuloption(14) = "X = A , Y = B', Cin = 1"
strmuloption(15) = "X = B , Y = A, Cin = 1"
strmuloption(16) = "X = B, Y = A', Cin = 1"

strmultiple(5) = "In an 16 bit two's complement representation, the value - 3 is represented by "
strmuloption(17) = "8003 Hex"
strmuloption(18) = "FFFA Hex"
strmuloption(19) = "8FF3 Hex"
strmuloption(20) = "FFFD Hex"

strmultiple(6) = "If even parity mechanism is being used for data transfer, which of the following data bytes has been received incorrectly?"
strmuloption(21) = "00 H"
strmuloption(22) = "03 H"
strmuloption(23) = "D0 H"
strmuloption(24) = "FF H"

strmultiple(7) = "Which of the following will be able to verify if a given string is a valid parenthesized arithmetic expression?"
strmuloption(25) = "A lexical analyzer"
strmuloption(26) = "A parser"
strmuloption(27) = "A postfix expression evaluator"
strmuloption(28) = "None of the above"

strmultiple(8) = "The minimum number of flip flops required to implement a 6-bit ring couner is -"
strmuloption(29) = "4"
strmuloption(30) = "8"
strmuloption(31) = "16"
strmuloption(32) = "65536"

strmultiple(9) = "Assuming that there are no branch instructions or data hazards, the number of cycles required to process 100 instructions through a 3-stage instruction pipeline is -"
strmuloption(33) = "300"
strmuloption(34) = "103"
strmuloption(35) = "102"
strmuloption(36) = "97"

strmultiple(10) = "The number of select input lines in an 8 to 1 multiplexer is -"
strmuloption(37) = "1"
strmuloption(38) = "8"
strmuloption(39) = "256"
strmuloption(40) = "None of the above"

strmulanswer(1) = "Instruction it is executing"
strmulanswer(2) = "AND the A register with mask 0H and test if Z flag is set to ."
strmulanswer(3) = "Infix expressions"
strmulanswer(4) = "X = A , Y = B', Cin = 1"
strmulanswer(5) = "FFFD Hex"
strmulanswer(6) = "D0 H"
strmulanswer(7) = "A parser"
strmulanswer(8) = "4"
strmulanswer(9) = "102"
strmulanswer(10) = "None of the above"

End Sub

Private Sub multiplechoice_Paper2()
strmultiple(1) = "The register A of an 8-bit computer contains the value - 25 decimal. The computer executes the INR A instruction which increments the contents of A by 1. The contents of A after the instruction execution are, in hexadecimal -"
strmuloption(1) = "F8"
strmuloption(2) = "FB"
strmuloption(3) = "E8"
strmuloption(4) = "EB"

strmultiple(2) = "The advantage of horizontally organized microcode over vertically organized microcode is -"
strmuloption(5) = "Smaller number of storage locations are required"
strmuloption(6) = "Smaller number of control steps are required"
strmuloption(7) = "Generation of control signals is faster"
strmuloption(8) = "None of the above"

strmultiple(3) = "The least negative value that the product of two 8-bit two's complement numbers can take is -"
strmuloption(9) = "-2(14)"
strmuloption(10) = "-2(15)"
strmuloption(11) = "-2(16)"
strmuloption(12) = "None of the above"

strmultiple(4) = "The value 38 Hex is stored in register B of an 8-bit computer. The computer executes an instruction to complement all the bits of register B. Assuming two's complement representation. He decimal value in register B after the execution of the instruction is "
strmuloption(13) = "57"
strmuloption(14) = "-57"
strmuloption(15) = "-56"
strmuloption(16) = "none of the above"

strmultiple(5) = "A Toggle flip flop can be constructed using a JK flip flop by connecting the -"
strmuloption(17) = "Toggle input to J and the inserted form of toggle input to K"
strmuloption(18) = "The toggle input to J"
strmuloption(19) = "Inverted form of toggle input to K"
strmuloption(20) = "None of the above"

strmultiple(6) = "Two 4-bit carry look ahead adders A1 and A2 are cascaded in order to form an 8-bit adder. The carry output from adder A1 is given as carry input to adder A2. When wiring the circuit, the carry output from A1 is accidentally connected to the ground signal. Which of the following situations reveals this fault?"
strmuloption(21) = "Addition of 44H and 22H"
strmuloption(22) = "Addition of F0H and FFH"
strmuloption(23) = "Addition of 0FH and 01H"
strmuloption(24) = "Addition of F4H and 22H"

strmultiple(7) = "The fully associative cache memory of a computer has 4 blocks of 256 bytes each. The cache is initially empty. The CPU requests for blocks 0,2,3,1,4,0,2,3,1. Assuming the LRU replacement policy, the number of misses in this access sequence is -"
strmuloption(25) = "4"
strmuloption(26) = "7"
strmuloption(27) = "8"
strmuloption(28) = "9"

strmultiple(8) = "Which of the following  is not an advantage of a subrouting?"
strmuloption(29) = "Faster execution of programs"
strmuloption(30) = "Easy readability of programs"
strmuloption(31) = "Saving in memory storage"
strmuloption(32) = "Better documentation of programs"

strmultiple(9) = "A page replacement policy is not necessary in"
strmuloption(33) = "Fully associative caches"
strmuloption(34) = "Set associative caches"
strmuloption(35) = "Directly mapped caches"
strmuloption(36) = "Write through caches"

strmultiple(10) = "Which of the following representations requires he least number of bits to store the number +255?"
strmuloption(37) = "Binary coded decimal"
strmuloption(38) = "Two's complement"
strmuloption(39) = "1's complement"
strmuloption(40) = "unsigned binary"

strmulanswer(1) = "E8"
strmulanswer(2) = "Generation of control signals is faster"
strmulanswer(3) = "None of the above"
strmulanswer(4) = "-57"
strmulanswer(5) = "None of the above"
strmulanswer(6) = "Addition of 0FH and 01H"
strmulanswer(7) = "7"
strmulanswer(8) = "Faster execution of programs"
strmulanswer(9) = "Directly mapped caches"
strmulanswer(10) = "unsigned binary"
End Sub

Private Sub multiplechoice_Paper3()
strmultiple(1) = "In a stack computer, there is support for -"
strmuloption(1) = "PUSH and POP instructions only"
strmuloption(2) = "Zero address instructions only"
strmuloption(3) = "Zero address instructions, PUSH, and POP"
strmuloption(4) = "None of the above"

strmultiple(2) = "In the instruction LOAD A, #32, the first operand is the destination operand and is in the register addressing mode. The second operand is the source operand and is in the immediate addressing mode. This means that the -"
strmuloption(5) = "Contents of location 32 will be moved to register A."
strmuloption(6) = "Contents of location 32 will be moved to the address stored in register A."
strmuloption(7) = "Value 32 will be moved into register A."
strmuloption(8) = "Value in register A will be moved into the location 32."

strmultiple(3) = "A printer device is connected to a computer. The status of the printer is indicated by bit 3 of the 8-bit status register. (Assume LSB position is bit 0.) If bit 3 is 1, then the printer is ready, otherwise the printer is not ready. To test if the printer is ready -"
strmuloption(9) = "AND the status register with 08H and test for Z flag = 0."
strmuloption(10) = "AND the status register with 04H and test for Z flag = 1"
strmuloption(11) = "OR the status register with 08H and test for Z flag = 0"
strmuloption(12) = "None of the above"

strmultiple(4) = "The reverse polish expression 'eab+cab+*/+' stands for -"
strmuloption(13) = "(e+(a+d)/(c*(a+b)))"
strmuloption(14) = "(e+a*(d+c)/(a+b)))"
strmuloption(15) = "(e+(a+d)+c*(a/b))"
strmuloption(16) = "none of the above"

strmultiple(5) = "Floating point division involves-"
strmuloption(17) = "Fixed point division of mantissas and fixed point addition of exponents"
strmuloption(18) = "Floating point division of mantissas and fixed point addition of exponents"
strmuloption(19) = "Fixed point division of mantissas and fixed point subtraction of exponents"
strmuloption(20) = "Floating point division of mantissas and fixed point subtraction of exponents"

strmultiple(6) = "In order to obtain the best possible speedup from a k-stage pipeline -"
strmuloption(21) = "K must be equal to log2n  where n is the number of inputs handled by the pipeline."
strmuloption(22) = "The clock signal to the individual stages must be skewed by an amount equal to the latch delay."
strmuloption(23) = "The delays of individual stages must be made equal"
strmuloption(24) = "All of the above"

strmultiple(7) = "The minimum hardware required to construct a 3-to-8 decoder is using -"
strmuloption(25) = "Two 2 to 4 decoders"
strmuloption(26) = "Two 2 to 4 decoders and a 1 to 2 decoder"
strmuloption(27) = "Three 2 to 4 decoders"
strmuloption(28) = "Depends on the technology (TTL, CMOS, etc.)"

strmultiple(8) = "Two instructions I1 and I2 are of the form D1 = f(S1) and D2 = g(S2) reapectively. Here S1 is the set of source operands for I1 and S2 is the set of source operands for I2. D1 and D2 are the destination operands for I1 and I2 respectively. F and g are functions (such as ADD, SUB etc.). The instructions I1 and I2 can be executed concurrently if -"
strmuloption(29) = "D1 does not belong to S2 and D2 does not belong to S1."
strmuloption(30) = "D1 is not the same as D2."
strmuloption(31) = "Both a and b above"
strmuloption(32) = "None of the above"

strmultiple(9) = "The value 35 (decimal) is loaded in the 8-bit register A of a microprocessor. The processor executes the 'rotate register A right' instruction 4 times. The value in the register will now be -"
strmuloption(33) = "53(decimal)"
strmuloption(34) = "83(decimal)"
strmuloption(35) = "82 (decimal)"
strmuloption(36) = "none of the above"

strmultiple(10) = "A microprocessor is a processor with a reduced -"
strmuloption(37) = "Instruction set"
strmuloption(38) = "Power requirement"
strmuloption(39) = "MIPS performance"
strmuloption(40) = "None of the above"

strmulanswer(1) = "Zero address instructions, PUSH, and POP"
strmulanswer(2) = "Value 32 will be moved into register A."
strmulanswer(3) = "AND the status register with 04H and test for Z flag = 1"
strmulanswer(4) = "(e+(a+d)/(c*(a+b)))"
strmulanswer(5) = "Fixed point division of mantissas and fixed point subtraction of exponents"
strmulanswer(6) = "The delays of individual stages must be made equal"
strmulanswer(7) = "Two 2 to 4 decoders and a 1 to 2 decoder"
strmulanswer(8) = "Both a and b above"
strmulanswer(9) = "none of the above"
strmulanswer(10) = "None of the above"

End Sub

Private Sub multiplechoice_Paper4()
strmultiple(1) = "In a CPU, the current value of program counter is saved in a special register called periods context register before the control transfers to the subroutine. Such a CPU cannot support -"
strmuloption(1) = "Recursive subroutines"
strmuloption(2) = "Branch instructions"
strmuloption(3) = "Interrupts"
strmuloption(4) = "None of the above"

strmultiple(2) = "In an 8-bit Booth's multiplication algorithm, the largest number of additions that will be ever required is -"
strmuloption(5) = "8"
strmuloption(6) = "4"
strmuloption(7) = "3"
strmuloption(8) = "2"

strmultiple(3) = "In a microprogram with 10 columns and 10 rows, it is observed that all the columns are compatible to one another i.e. no two control signals are activated in the same clock cycle. We can reduce the size of the microprogram to -"
strmuloption(9) = "1 column and 1 row"
strmuloption(10) = "4 columns and 1 row"
strmuloption(11) = "4  columns and 4 rows"
strmuloption(12) = "4 columns and 10 rows"

strmultiple(4) = "A compute readable image has 256 × 256  picture elements. Each picture element has a great level. There are at most 256 grey levels. To store one image, we require -"
strmuloption(13) = "128 KB"
strmuloption(14) = "64 KB"
strmuloption(15) = "8 KB"
strmuloption(16) = "none of the above"

strmultiple(5) = "The Hamming distance between two n bit binary numbers is defined as the number of bit positions they differ. If we have a function called onescount(x) which returns the number of 1's in the number x, then a correct way to compute the Hamming distance between  two numbers x and y is -"
strmuloption(17) = "Onescount(x) + onescount(y) - n"
strmuloption(18) = "Onescount(x&y) , where & is the bitwise AND operation"
strmuloption(19) = "Onescount (x|y), where | is the bitwise OR operation"
strmuloption(20) = "Onescount(x@y) , where @ is the bitwise XOR operation"

strmultiple(6) = "In a computer fractions are represented using a 16 bit fixed point notation and the binary point is assumed to be at the left of the most significant bit. The representation for 0.625 in this notation is -"
strmuloption(21) = "C0 in Hex"
strmuloption(22) = "AA in Hex"
strmuloption(23) = "D0 in Hex"
strmuloption(24) = "None of the above"

strmultiple(7) = "Two 4-bit two's complement numbers are added using a ripple carry adder. The range of the sum output is -"
strmuloption(25) = "-128 to +127"
strmuloption(26) = "-256 to +256"
strmuloption(27) = "-512 to +511"
strmuloption(28) = "-256 to + 256"

strmultiple(8) = "Cycle stealing refers to -"
strmuloption(29) = "Reduction in the number of clock cycles for memory access in a cache memory"
strmuloption(30) = "Overlapping of the memory refresh operation and memory read operations in dynamic memories"
strmuloption(31) = "Reduction in instruction cycle time through instruction pipelining."
strmuloption(32) = "None of the above"

strmultiple(9) = "Two negative numbers are added and the result is observed to be positive. The most likely reason for this phenomenon is -"
strmuloption(33) = "A dynamic fault in the adder circuit"
strmuloption(34) = "An underflow due to limited precision of the adder"
strmuloption(35) = "Incorrect representation of the input numbers"
strmuloption(36) = "A wiring fault in the adder circuit"

strmultiple(10) = "A computer program includes a subroutine which is declared as an 'external' subroutine. This means that the subroutine"
strmuloption(37) = "Is denied in the kernel of the operating system."
strmuloption(38) = "Uses external devices"
strmuloption(39) = "Is defined elsewhere"
strmuloption(40) = "Resides on a different compute and will be downloaded dynamically at run time."

strmulanswer(1) = "Recursive subroutines"
strmulanswer(2) = "4"
strmulanswer(3) = "4 columns and 10 rows"
strmulanswer(4) = "64 KB"
strmulanswer(5) = "Onescount(x@y) , where @ is the bitwise XOR operation"
strmulanswer(6) = "None of the above"
strmulanswer(7) = "-256 to +256"
strmulanswer(8) = "None of the above"
strmulanswer(9) = "An underflow due to limited precision of the adder"
strmulanswer(10) = "Is defined elsewhere"
End Sub

Private Sub multiplechoice_Paper5()
strmultiple(1) = "In an n- bit Booth's multiplication algorithm, the largest number of operations required are -"
strmuloption(1) = "N additions and n subtractions"
strmuloption(2) = "Log2(n) additions and log2(n) subtractions"
strmuloption(3) = "N-1 additions and n-1 subtractions"
strmuloption(4) = "None of the above"

strmultiple(2) = "In a microprogram with 10 rows and 10 columns, it is observed that only two columns c1 and c3 are compatible i.e. the signals c1 and c3 are not activated in the same clock cycle. In the vertical format, this microprogram will contain:"
strmuloption(5) = "9 rows and 10 columns"
strmuloption(6) = "10 rows and 9 columns"
strmuloption(7) = "9 rows and 9 columns"
strmuloption(8) = "none of the above"

strmultiple(3) = "The difference between a COMPARE instruction and a SUBTRACT instruction is -"
strmuloption(9) = "COMPARE instruction does not set any flags"
strmuloption(10) = "SUBTRACT instruction does not set any flags"
strmuloption(11) = "COMPARE instruction leaves the registers unchanged"
strmuloption(12) = "COMPARE instruction does not require the ALU."

strmultiple(4) = "Which of the following  subroutines in unlikely to be part of the kernel of an operating system?"
strmuloption(13) = "Process scheduling algorithm"
strmuloption(14) = "Sorting algorithm"
strmuloption(15) = "Page swapping algorithm"
strmuloption(16) = "Maintenance of the real time clock"

strmultiple(5) = "In the IEEE 754 floating point representation standard, the base is -"
strmuloption(17) = "23"
strmuloption(18) = "127"
strmuloption(19) = "16"
strmuloption(20) = "none of the above"

strmultiple(6) = "The sum of two hexadecimal numbers 23D and 9AA gives the hexadecimal number -"
strmuloption(21) = "AF7"
strmuloption(22) = "BF6"
strmuloption(23) = "BE7"
strmuloption(24) = "BE5"

strmultiple(7) = "The simplified form of the expression AB + ABC' is -"
strmuloption(25) = "AB"
strmuloption(26) = "A(B+C)"
strmuloption(27) = "A(B+C')"
strmuloption(28) = "None of the above"

strmultiple(8) = "In an 8-bit microcomputer, the locations 100 Hex and 101 Hex contain alues F3 Hex and FE Hex respectively. A program fetches these two values, compares them on relative magnitude, stores the large of the n location 102 Hex. Assuming two's complement representation, the value stored in location 102 is -"
strmuloption(29) = "F3 Hex"
strmuloption(30) = "FE Hex"
strmuloption(31) = "102 Hex"
strmuloption(32) = "none of the above"

strmultiple(9) = "The memory map of a microcomputer is as follows: the first 8 KB of the meory cosists of EPROM memory and the remaining 56 KB of memory consists of RAM memory. The valid RAM locations are :"
strmuloption(33) = "0000 - 1 FFF"
strmuloption(34) = "1 FFF - FFFF"
strmuloption(35) = "we cannot determine the memory map of the RAM without additional information"
strmuloption(36) = "none of the above"

strmultiple(10) = "the basic reason why hierarchical memory organization in computers can offer very small average access time is -"
strmuloption(37) = "the advances in semiconductor memory technology have led to he development of very high speed memories"
strmuloption(38) = "the computer programs display a locality of reference"
strmuloption(39) = "due to advances in software technology, computer programs have become small enough o fit into cache memory"
strmuloption(40) = "operating systems provide advanced memory management functions which reduce the average access time considerably."

strmulanswer(1) = "None of the above"
strmulanswer(2) = "none of the above"
strmulanswer(3) = "COMPARE instruction leaves the registers unchanged"
strmulanswer(4) = "Sorting algorithm"
strmulanswer(5) = "none of the above"
strmulanswer(6) = "BE7"
strmulanswer(7) = "None of the above"
strmulanswer(8) = "FE Hex"
strmulanswer(9) = "none of the above"
strmulanswer(10) = "the computer programs display a locality of reference"
End Sub

Private Sub multiplechoice_Paper6()
strmultiple(1) = "Which of the following is not a universal building block?"
strmuloption(1) = "Three input NAND gate"
strmuloption(2) = "Two input NOR gate"
strmuloption(3) = "Two input multiplexer"
strmuloption(4) = "Two input EXOR gate"

strmultiple(2) = "A four bit register contains a number X. After two rotate right operations and one rotate left operation on the register, the value in the register is 6(decimal). The number X is equal to -"
strmuloption(5) = "14 in octal"
strmuloption(6) = "12 in octal"
strmuloption(7) = "10 in octal"
strmuloption(8) = "03 in octal"

strmultiple(3) = "X and Y are two digit BCD numbers. It is known that X + Y is equal to 82(BCD) and X-Y = 04(BCD). Calculate the value of X and express it as an unsigned binary number. The value of X is equal to -"
strmuloption(9) = "01000011"
strmuloption(10) = "0000101"
strmuloption(11) = "00101011"
strmuloption(12) = "00100111"

strmultiple(4) = "The minimum number of JK flip flops required to construct a sequential circuit which has 23 states is -"
strmuloption(13) = "23"
strmuloption(14) = "6"
strmuloption(15) = "There is insufficient information"
strmuloption(16) = "none of the above"

strmultiple(5) = "Which one of the following strings is not described by the regular expression? (0+1) 1 * (10)*"
strmuloption(17) = "The empty string"
strmuloption(18) = "0"
strmuloption(19) = "10"
strmuloption(20) = "110"

strmultiple(6) = "Which of the following sets is not regular ?"
strmuloption(21) = "The set of all binary strings which begin with a 0 and end with a 1."
strmuloption(22) = "The set off all binary strings which have at least one occurrence of the pattern 101 in them."
strmuloption(23) = "The set of all binary strings which represent multiples of two."
strmuloption(24) = "None of the above"

strmultiple(7) = "In order to correct a single error, the minimum number of check bits that must be added to 4 data bits is -"
strmuloption(25) = "2"
strmuloption(26) = "3"
strmuloption(27) = "4"
strmuloption(28) = "there is insufficient data to answer the question"

strmultiple(8) = "The fundamental mode of operation of asynchronous sequential circuits implies that -"
strmuloption(29) = "We cannot feed pulses as input to the circuit"
strmuloption(30) = "The flow table of the circuit cannot be reduced"
strmuloption(31) = "There exists a homing sequence for each of the states in the flow table"
strmuloption(32) = "None of the above"

strmultiple(9) = "Which of the following is not an advantage of asynchronous circuits?"
strmuloption(33) = "Higher speed"
strmuloption(34) = "Lesser power consumption"
strmuloption(35) = "Smaller design effort"
strmuloption(36) = "No need to provide clock generation circuitry"

strmultiple(10) = "The minimum number of bits required to represent numbers in the range -23 to +31 is -"
strmuloption(37) = "8"
strmuloption(38) = "7"
strmuloption(39) = "6"
strmuloption(40) = "5"

strmulanswer(1) = "Two input EXOR gate"
strmulanswer(2) = "12 in octal"
strmulanswer(3) = "01000011"
strmulanswer(4) = "none of the above"
strmulanswer(5) = "The empty string"
strmulanswer(6) = "None of the above"
strmulanswer(7) = "3"
strmulanswer(8) = "None of the above"
strmulanswer(9) = "Smaller design effort"
strmulanswer(10) = "6"

End Sub

Private Sub multiplechoice_Paper7()
strmultiple(1) = "If even parity is used in transmitting 8-bit data, which of the following bytes has been received with a transmission error?"
strmuloption(1) = "10000001"
strmuloption(2) = "11001100"
strmuloption(3) = "00000000"
strmuloption(4) = "none of the above"

strmultiple(2) = "A CPU has a 16 bit program counter. This means that the CPU can address -"
strmuloption(5) = "16K memory locations"
strmuloption(6) = "32 K memory locations"
strmuloption(7) = "64K memory locations"
strmuloption(8) = "256K memory locations"

strmultiple(3) = "How many memory locations re required to store the instruction LX1 H, 0800H in an 8085 assembly language program?"
strmuloption(9) = "1"
strmuloption(10) = "2"
strmuloption(11) = "3"
strmuloption(12) = "4"

strmultiple(4) = "How many memory fetches (including instruction fetch) are required to execute the instruction LX1 H,0800H?"
strmuloption(13) = "2"
strmuloption(14) = "3"
strmuloption(15) = "4"
strmuloption(16) = "5"

strmultiple(5) = "If interrupts arrive on the three lines INTR, RST 6.5, and RST 7.5, which of them will the 8085 CPU acknowledge?"
strmuloption(17) = "INTR"
strmuloption(18) = "RST 6.5"
strmuloption(19) = "RST 7.5"
strmuloption(20) = "The 8085 will poll the interrupt statuses and acknowledge the interrupt which was acknowledged least recently."

strmultiple(6) = "The smallest negative number that can be stored using 1 byte of storage and the two's complement number system is -"
strmuloption(21) = "-127"
strmuloption(22) = "-128"
strmuloption(23) = "-255"
strmuloption(24) = "-256"

strmultiple(7) = "An IBM compatible PC386 SX system uses -"
strmuloption(25) = "A 16 bit CPU"
strmuloption(26) = "A 32 bit CPU"
strmuloption(27) = "A 64 bit CPU"
strmuloption(28) = "A 386 bit CPU"

strmultiple(8) = "A 3.5 double density floppy disk which does not have any surface defects can be formatted to store a maximum of -"
strmuloption(29) = "1.2 MB of data"
strmuloption(30) = "1.22 MB of data"
strmuloption(31) = "1.4 MB of data"
strmuloption(32) = "1.44 MB of data"

strmultiple(9) = "There are ____________ possible functions of 3 variables."
strmuloption(33) = "3"
strmuloption(34) = "8"
strmuloption(35) = "64"
strmuloption(36) = "256"

strmultiple(10) = "Which of the following is not a universal building block?"
strmuloption(37) = "2-input NAND gate"
strmuloption(38) = "3-input NAND gate"
strmuloption(39) = "2 to 1 multiplexer"
strmuloption(40) = "2-input XOR gate"

strmulanswer(1) = "none of the above"
strmulanswer(2) = "64K memory locations"
strmulanswer(3) = "3"
strmulanswer(4) = "3"
strmulanswer(5) = "RST 7.5"
strmulanswer(6) = "-127"
strmulanswer(7) = "A 32 bit CPU"
strmulanswer(8) = "1.44 MB of data"
strmulanswer(9) = "256"
strmulanswer(10) = "2-input XOR gate"

End Sub

Private Sub multiplechoice_Paper8()
strmultiple(1) = "The language BASIC is normally used along with -"
strmuloption(1) = "An assembler"
strmuloption(2) = "A compiler"
strmuloption(3) = "An interpreter"
strmuloption(4) = "A computer aided software engineering(CASE) tool"

strmultiple(2) = "A diamond shaped box in a flow chart refers to -"
strmuloption(5) = "An assignment statement"
strmuloption(6) = "Printing of a file"
strmuloption(7) = "An if then else decision"
strmuloption(8) = "Data input from a deice"

strmultiple(3) = "A Spooler allows printing -"
strmuloption(9) = "The contents of the screen using print screen option"
strmuloption(10) = "Of graphic files on a non graphics printer"
strmuloption(11) = "A file in background while we do other tasks on the computer"
strmuloption(12) = "Of different fonts on a laser jet printer"

strmultiple(4) = "A linker is a program which generates -"
strmuloption(13) = "The executable program from object programs"
strmuloption(14) = "An object program from an assembly language program"
strmuloption(15) = "An assembly language program from a high level language program"
strmuloption(16) = "The executable program from high level language program"

strmultiple(5) = "Pseudo code refers to a program which -"
strmuloption(17) = "Can be executed only sing virtual memory"
strmuloption(18) = "Can be executed without compilation and linking"
strmuloption(19) = "Is written in a high level language without following any syntax"
strmuloption(20) = "Is written in an assembly language for a non existent architecture"

strmultiple(6) = "Which of the following violates the principles of structured programming?"
strmuloption(21) = "The use of global variables"
strmuloption(22) = "The use of 'go to' statements in a program"
strmuloption(23) = "The use of 'do while' statements"
strmuloption(24) = "The use of recursive procedures"

strmultiple(7) = "A debugging tool is a program which-"
strmuloption(25) = "Removes bugs from a user program"
strmuloption(26) = "Removes viruses from the computer"
strmuloption(27) = "Helps the user find bugs in his program"
strmuloption(28) = "Displays the errors in a user program"

strmultiple(8) = "Which of the following is not an output of an assembler?"
strmuloption(29) = "Executable program"
strmuloption(30) = "Source listing with line numbers and errors"
strmuloption(31) = "A symbol table"
strmuloption(32) = "Object program containing machine opcodes"

strmultiple(9) = "Which of the following is not an advantage of Dynamic RAMs?"
strmuloption(33) = "High density"
strmuloption(34) = "Low cost"
strmuloption(35) = "High speed"
strmuloption(36) = "No need for memory refresh"

strmultiple(10) = "When the value 37H is divided by 17H, the remainder is"
strmuloption(37) = "00H"
strmuloption(38) = "03H"
strmuloption(39) = "07h"
strmuloption(40) = "09H"

strmulanswer(1) = "An interpreter"
strmulanswer(2) = "An if then else decision"
strmulanswer(3) = "A file in background while we do other tasks on the computer"
strmulanswer(4) = "The executable program from object programs"
strmulanswer(5) = "Is written in a high level language without following any syntax"
strmulanswer(6) = "The use of 'go to' statements in a program"
strmulanswer(7) = "Helps the user find bugs in his program"
strmulanswer(8) = "Executable program"
strmulanswer(9) = "High speed"
strmulanswer(10) = "09H"

End Sub

Private Sub multiplechoice_Paper9()
strmultiple(1) = "The unit distance code(gray code) of a given binary number (1001)n is -"
strmuloption(1) = "(1110)B"
strmuloption(2) = "(0110)B"
strmuloption(3) = "(1101)B"
strmuloption(4) = "(1111)B"

strmultiple(2) = "The computer architecture having stored program concept is -"
strmuloption(5) = "Hazard"
strmuloption(6) = "Von-Neumann"
strmuloption(7) = "Pascal"
strmuloption(8) = "None of the above"

strmultiple(3) = "A modulo-20 counter can be designed using-"
strmuloption(9) = "20-flipflops"
strmuloption(10) = "5-flipflops"
strmuloption(11) = "4-flipflops"
strmuloption(12) = "none of the above"

strmultiple(4) = "The memory that requires refreshing of data is -"
strmuloption(13) = "SRAM"
strmuloption(14) = "EAROM"
strmuloption(15) = "DRAM"
strmuloption(16) = "Bubble memory"

strmultiple(5) = "A scheme, in which the address specifies which memory word contains the address of the operand, is called -"
strmuloption(17) = "Immediate addressing"
strmuloption(18) = "Indirect addressing"
strmuloption(19) = "Direct addressing"
strmuloption(20) = "Based addressing"

strmultiple(6) = "A bootstrap is -"
strmuloption(21) = "A memory device"
strmuloption(22) = "A program to start up a computer"
strmuloption(23) = "Hardware of computer used to check memory"
strmuloption(24) = "An assembler"

strmultiple(7) = "The process of accessing information on a CD-ROM is -"
strmuloption(25) = "Random"
strmuloption(26) = "Random"
strmuloption(27) = ""
strmuloption(28) = "None of the above"

strmultiple(8) = "A program, embedded in semi conductor chips during their manufacture is called -"
strmuloption(29) = "Humanware"
strmuloption(30) = "Firmware"
strmuloption(31) = "Hardware"
strmuloption(32) = "Software"

strmultiple(9) = "An array processor consists of -"
strmuloption(33) = "Multiple control units and multiple ALU units"
strmuloption(34) = "Single control unit and multiple ALU units"
strmuloption(35) = "Multiple control units and single ALU unit"
strmuloption(36) = "Single control unit and single ALU unit"

strmultiple(10) = "Virtual memory system allows the employment of -"
strmuloption(37) = "More than address space"
strmuloption(38) = "The full address space"
strmuloption(39) = "More than the hard disk space"
strmuloption(40) = "None of the above"

strmulanswer(1) = "(1111)B"
strmulanswer(2) = "Pascal"
strmulanswer(3) = "5-flipflops"
strmulanswer(4) = "EAROM"
strmulanswer(5) = "Direct addressing"
strmulanswer(6) = "A program to start up a computer"
strmulanswer(7) = "Semi-random"
strmulanswer(8) = "Firmware"
strmulanswer(9) = "Multiple control units and single ALU unit"
strmulanswer(10) = "More than address space"

End Sub

Private Sub multiplechoice_Paper10()
strmultiple(1) = "In a J-K flip flop, the function K=J is used to realize –"
strmuloption(1) = "T-flip flop"
strmuloption(2) = "S-R flip flop"
strmuloption(3) = "D- flip flop"
strmuloption(4) = "M/S JK flip flop"

strmultiple(2) = "Serial to parallel data conversion is done, using –"
strmuloption(5) = "Counter "
strmuloption(6) = "Shift register"
strmuloption(7) = "Ring counter"
strmuloption(8) = "ADC"

strmultiple(3) = "CACHE memory is implemented, using"
strmuloption(9) = "Dynamic RAM"
strmuloption(10) = "Static RAM"
strmuloption(11) = "EAROM"
strmuloption(12) = "EPROM"

strmultiple(4) = "A 4-bit ring counter is initially loaded with –"
strmuloption(13) = "0000"
strmuloption(14) = "0110"
strmuloption(15) = "1001"
strmuloption(16) = "1000"

strmultiple(5) = "A nano control memory is implemented to –"
strmuloption(17) = "Improve the speed of execution"
strmuloption(18) = "Reduce the overall control memory size"
strmuloption(19) = "Reduce the complexity of hardware"
strmuloption(20) = "None of the above"

strmultiple(6) = "The major objective, in choosing page replacement policy, is to –"
strmuloption(21) = "Minimize the hit ratio"
strmuloption(22) = "Maximize the hit ratio"
strmuloption(23) = "Reduce the size of the page"
strmuloption(24) = "None of the above"

strmultiple(7) = "The basic circuit of ECL supports the –"
strmuloption(25) = "AND logic"
strmuloption(26) = "EX-OR logic"
strmuloption(27) = "NOT logic"
strmuloption(28) = "OR-NOR logic"

strmultiple(8) = "Ready to run machine language program is produced by –"
strmuloption(29) = "Compiler"
strmuloption(30) = "Loader"
strmuloption(31) = "Assembler"
strmuloption(32) = "All of the above"

strmultiple(9) = "In a 5 stage pipe lined processor, each stage is assumed to take 1 unit time. If there are 100 tasks to be executed, then it requires –"
strmuloption(33) = "500 time units"
strmuloption(34) = "105 time units"
strmuloption(35) = "104 time units"
strmuloption(36) = "none of the above"

strmultiple(10) = "Booth algorithm gives procedure for multiplying binary integers in –"
strmuloption(37) = "Signed magnitude representation"
strmuloption(38) = "Unsigned representation"
strmuloption(39) = "2’s complement representation"
strmuloption(40) = "none of the above"

strmulanswer(1) = "D- flip flop"
strmulanswer(2) = "Shift register"
strmulanswer(3) = "Dynamic RAM"
strmulanswer(4) = "1000"
strmulanswer(5) = "Reduce the overall control memory size"
strmulanswer(6) = "Maximize the hit ratio"
strmulanswer(7) = "OR-NOR logic"
strmulanswer(8) = "Loader"
strmulanswer(9) = "104 time units"
strmulanswer(10) = "2’s complement representation"

End Sub

Private Sub multiplechoice_Paper11()
strmultiple(1) = "If K = J, in a J-K flip flop this results in-"
strmuloption(1) = "T- flip flop"
strmuloption(2) = "D flip flop"
strmuloption(3) = "S-R flip flop"
strmuloption(4) = "None defined"

strmultiple(2) = "Racing condition in J-K flip flop is eliminated by -"
strmuloption(5) = "Cascading a D-flip flop to the Q output of JK flip flop"
strmuloption(6) = "Using master slave JK flipflop"
strmuloption(7) = "Connecting J and K together"
strmuloption(8) = "None of the above"

strmultiple(3) = "CPU gets the address of next instruction to be processed from -"
strmuloption(9) = "Instruction register"
strmuloption(10) = "Memory address register"
strmuloption(11) = "Index register"
strmuloption(12) = "Program counter"

strmultiple(4) = "Which of the following algorithms is not related to removal, in demand paged memory management?"
strmuloption(13) = "FIFO"
strmuloption(14) = "LRU"
strmuloption(15) = "Round robin"
strmuloption(16) = "All of the above"

strmultiple(5) = "Floating point multiplication requires -"
strmuloption(17) = "Fixed point multiplication of the mantissa and fixed point addition of exponents"
strmuloption(18) = "Fixed point addition of mantissa and fixed point multiplication of exponents"
strmuloption(19) = "Fixed point addition of mantissa and fixed point addition of exponents"
strmuloption(20) = "None of these"

strmultiple(6) = "In a compiler, the recognition of basic elements and creation of uniform symbols is carried out during -"
strmuloption(21) = "Lexical analysis"
strmuloption(22) = "Syntax analysis"
strmuloption(23) = "Interpretation"
strmuloption(24) = "Code generation"

strmultiple(7) = "In a floppy disk, data are stored on -"
strmuloption(25) = "A single spiral track like a phonograph record"
strmuloption(26) = "A single circular track containing multiple sectors"
strmuloption(27) = "Multiple spiral tracks"
strmuloption(28) = "None of the above"

strmultiple(8) = "Fetch operations are not required in -"
strmuloption(29) = "Direct addressing"
strmuloption(30) = "Immediate addressing"
strmuloption(31) = "Indirect addressing"
strmuloption(32) = "Register addressing"

strmultiple(9) = "In an 8088 microprocessor -"
strmuloption(33) = "Internal and external data bus sizes are equal"
strmuloption(34) = "Internal data bus is 8-bit wide and external data bus is 16 bit wide"
strmuloption(35) = "Internal data bus is 16 bit wide and external data bus is 8-bit wide"
strmuloption(36) = "None of these"

strmultiple(10) = "In case of 68000 Assembly language program, the order of operands is -"
strmuloption(37) = "Source operand is followed by destination operand"
strmuloption(38) = "Only source operand present and destination operand is implied"
strmuloption(39) = "Destination operand is followed by source operand"
strmuloption(40) = "In any order"

strmulanswer(1) = "T- flip flop"
strmulanswer(2) = "Using master slave JK flipflop"
strmulanswer(3) = "Program counter"
strmulanswer(4) = "Round robin"
strmulanswer(5) = "Fixed point multiplication of the mantissa and fixed point addition of exponents"
strmulanswer(6) = "Lexical analysis"
strmulanswer(7) = "None of the above"
strmulanswer(8) = "Register addressing"
strmulanswer(9) = "Internal data bus is 16 bit wide and external data bus is 8-bit wide"
strmulanswer(10) = "Source operand is followed by destination operand"
End Sub

Private Sub multiplechoice_Paper12()
strmultiple(1) = "CARRY, in a half adder, can be obtained using -"
strmuloption(1) = "EX-OR gate"
strmuloption(2) = "And gate"
strmuloption(3) = "Or gate"
strmuloption(4) = "EX-NOR gate"

strmultiple(2) = "A memory device, in which a bit is stored as a charge across the stray capacitance -"
strmuloption(5) = "SRAM"
strmuloption(6) = "EPROM"
strmuloption(7) = "DRAM"
strmuloption(8) = "Bubble memory"

strmultiple(3) = "Assembler directives -"
strmuloption(9) = "Allocate storage for constants and program variables"
strmuloption(10) = "Are converted into machine instructions to be included into object code."
strmuloption(11) = "Are instructions to assembler to tell how assembly of source program is to be performed"
strmuloption(12) = "None of these"

strmultiple(4) = "To transfer the contents of register R1 to register R2 enable -"
strmuloption(13) = "Input of register R1 and input of register R2."
strmuloption(14) = "Input of register R1 and output of register R2"
strmuloption(15) = "Output of register R1 and input of register R2"
strmuloption(16) = "None of these"

strmultiple(5) = "The key technology used, in fourth generation computers, is -"
strmuloption(17) = "MSI"
strmuloption(18) = "SSI"
strmuloption(19) = "LSI & VLSI"
strmuloption(20) = "Transistor"

strmultiple(6) = "The  following logic building block can be used to implement any combinational logic circuit -"
strmuloption(21) = "Decoder"
strmuloption(22) = "Multiplexer"
strmuloption(23) = "EX-OR gate"
strmuloption(24) = "Encoder"

strmultiple(7) = "In based addressing mode, instruction contains -"
strmuloption(25) = "Base address"
strmuloption(26) = "Displacement"
strmuloption(27) = "Absolute address"
strmuloption(28) = "None of the above"

strmultiple(8) = "Stack based architecture supports -"
strmuloption(29) = "Two address instructions"
strmuloption(30) = "Three address instructions"
strmuloption(31) = "Zero address instructions"
strmuloption(32) = "None of the above"

strmultiple(9) = "Read/write head of hard disk drive -"
strmuloption(33) = "Does not make physical contact with magnetic media while reading/writing"
strmuloption(34) = "Makes physical contact with magnetic media while reading/ writing"
strmuloption(35) = "Makes physical contact while writing only"
strmuloption(36) = "None of these"

strmulanswer(1) = "And gate"
strmulanswer(2) = "DRAM"
strmulanswer(3) = "Allocate storage for constants and program variables"
strmulanswer(4) = "Output of register R1 and input of register R2"
strmulanswer(5) = "LSI & VLSI"
strmulanswer(6) = "Multiplexer"
strmulanswer(7) = "Displacement"
strmulanswer(8) = "Zero address instructions"
strmulanswer(9) = "Does not make physical contact with magnetic media while reading/writing"
End Sub
