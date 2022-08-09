VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCNPapers"
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
 
Public Function DCN(PaperType As String, DCNPaper As Paper, mCat As Category, Number As Integer)
'ptype is the paper type
'DCNpaper is the number of paper as paper1,paper2 ...
'mcat is the category
'number is the number of questions, or options or answers to be taken

StrPaperType = PaperType
' StrPaper = Paper
If PaperType = "Fill in the Blanks" Then
   If DCNPaper = Paper1 Then Fill_in_the_Blanks_Paper1
   If DCNPaper = Paper2 Then Fill_in_the_Blanks_Paper2
   If DCNPaper = Paper3 Then Fill_in_the_Blanks_Paper3
   If DCNPaper = Paper4 Then Fill_in_the_Blanks_Paper4
   If DCNPaper = Paper5 Then Fill_in_the_Blanks_Paper5
   If DCNPaper = Paper6 Then Fill_in_the_Blanks_Paper6
   If DCNPaper = Paper7 Then Fill_in_the_Blanks_Paper7
   If DCNPaper = Paper8 Then Fill_in_the_Blanks_Paper8
   If DCNPaper = Paper9 Then Fill_in_the_Blanks_Paper9
   If DCNPaper = Paper10 Then Fill_in_the_Blanks_Paper10
   
ElseIf PaperType = "Matching Columns" Then
       If DCNPaper = Paper1 Then Matching_Columns_Paper1
       If DCNPaper = Paper2 Then Matching_Columns_Paper2
       If DCNPaper = Paper3 Then Matching_Columns_Paper3
       If DCNPaper = Paper4 Then Matching_Columns_Paper4
       If DCNPaper = Paper5 Then Matching_Columns_Paper5
       If DCNPaper = Paper6 Then Matching_Columns_Paper6
       If DCNPaper = Paper7 Then Matching_Columns_Paper7
       If DCNPaper = Paper8 Then Matching_Columns_Paper8
       If DCNPaper = Paper9 Then Matching_Columns_Paper9
       If DCNPaper = Paper10 Then Matching_Columns_Paper10
       
ElseIf PaperType = "True Or False" Then
       If DCNPaper = Paper1 Then TrueorFalse_Paper1
       If DCNPaper = Paper2 Then TrueorFalse_Paper2
       If DCNPaper = Paper3 Then TrueorFalse_Paper3
       If DCNPaper = Paper4 Then TrueorFalse_Paper4
       If DCNPaper = Paper5 Then TrueorFalse_Paper5
       If DCNPaper = Paper6 Then TrueorFalse_Paper6
       If DCNPaper = Paper7 Then TrueorFalse_Paper7
       If DCNPaper = Paper8 Then TrueorFalse_Paper8
       If DCNPaper = Paper9 Then TrueorFalse_Paper9
       If DCNPaper = Paper10 Then TrueorFalse_Paper10
         
ElseIf PaperType = "Descriptive Questions" Then
       If DCNPaper = Paper1 Then questions_paper1
       If DCNPaper = Paper2 Then questions_paper2
       If DCNPaper = Paper3 Then questions_paper3
       If DCNPaper = Paper4 Then questions_paper4
       If DCNPaper = Paper5 Then questions_paper5
       If DCNPaper = Paper6 Then questions_paper6
       If DCNPaper = Paper7 Then questions_paper7
       
ElseIf PaperType = "Multiple Choice" Then
       If DCNPaper = Paper1 Then multiplechoice_Paper1
       If DCNPaper = Paper2 Then multiplechoice_Paper2
       If DCNPaper = Paper3 Then multiplechoice_Paper3
       If DCNPaper = Paper4 Then multiplechoice_Paper4
       If DCNPaper = Paper5 Then multiplechoice_Paper5
       If DCNPaper = Paper6 Then multiplechoice_Paper6
       If DCNPaper = Paper7 Then multiplechoice_Paper7
       If DCNPaper = Paper8 Then multiplechoice_Paper8
       If DCNPaper = Paper9 Then multiplechoice_Paper9
       If DCNPaper = Paper10 Then multiplechoice_Paper10
End If

If mCat = Fill_Questions Then
   DCN = StrQuestions(Number)
ElseIf mCat = Fill_Options Then
       DCN = stroptions(Number)
ElseIf mCat = Fill_Answers Then
       DCN = stranswers(Number)
       
ElseIf mCat = Matching_ColumnA Then
       DCN = strcolumnA(Number)
ElseIf mCat = Matching_ColumnB Then
       DCN = strcolumnB(Number)
ElseIf mCat = Matching_Answers Then
       DCN = stranswers(Number)
       
ElseIf mCat = Questions Then
       DCN = StrQuestions(Number)
       
ElseIf mCat = TF_Questions Then
       DCN = StrQuestions(Number)
ElseIf mCat = TF_Answers Then
       DCN = stranswers(Number)
       
ElseIf mCat = MC_Questions Then
       DCN = strmultiple(Number)
       
ElseIf mCat = MC_Options_1 Then
       DCN = strmuloption(Number)
       
ElseIf mCat = MC_Options_2 Then
       DCN = strmuloption(Number)
       
ElseIf mCat = MC_Options_3 Then
       DCN = strmuloption(Number)
       
ElseIf mCat = MC_Options_4 Then
       DCN = strmuloption(Number)
       
ElseIf mCat = MC_Answer Then
       DCN = strmulanswer(Number)
       
End If

End Function

Private Sub Fill_in_the_Blanks_Paper1()
StrQuestions(1) = "A seven-bit code can generate dash possible characters."
StrQuestions(2) = "Synchronous transmission is almost dash % efficient than asynchronous transmission."
StrQuestions(3) = "Each asynchronous character is preceded by a dash bit which has the opposite polarity to the idle condition."
StrQuestions(4) = "Digital circuits are dash susceptible to distortion and interference."
StrQuestions(5) = "Statistical multiplexers may also be classed as dash ."
StrQuestions(6) = "The upper limit on the rate at which information can be transmitted by a system of fixed BW and SNR is called the dash."
StrQuestions(7) = "Due to less dash effect, coaxial cable has a lower degree of attenuation at high frequencies than twisted pair cable."
StrQuestions(8) = "dash fibers contains many hundreds of modes."
StrQuestions(9) = "All communication satellites are dash."
StrQuestions(10) = "The approximate roundtrip delay in satellite communication is dash ? sec."

stroptions(1) = "Start"
stroptions(2) = "Concentrators"
stroptions(3) = "Geostationary"
stroptions(4) = "Information capacity"
stroptions(5) = "20%"
stroptions(6) = "Multimode"
stroptions(7) = "256"
stroptions(8) = "Skin"
stroptions(9) = "Less"
stroptions(10) = "128"

stranswers(1) = "128"
stranswers(2) = "20%"
stranswers(3) = "Start"
stranswers(4) = "Less"
stranswers(5) = "Concentrators"
stranswers(6) = "Information capacity"
stranswers(7) = "Skin"
stranswers(8) = "Multimode"
stranswers(9) = "Geostationary"
stranswers(10) = "256"

End Sub

Private Sub Fill_in_the_Blanks_Paper2()
StrQuestions(1) = "dash modulation technique is used in high speed modems."
StrQuestions(2) = "DM(Delta Modulation) is a dash coding technique."
StrQuestions(3) = "With 8 bit code per sample, the bit rate becomes dash kbps for voice signals."
StrQuestions(4) = "Modems, operating at data rates up to 1.8 kbps, generally employ dash modulation technique."
StrQuestions(5) = "The difference between the decoded signal and the original signal is referred to as the dash."
StrQuestions(6) = "At the physical level of IEEE 802.3 standard, dash code is used."
StrQuestions(7) = "Hub polling is usually employed on dash lines which have large line turnaround delay."
StrQuestions(8) = "Packet switching greatly reduces the store and forward dash at each IMP."
StrQuestions(9) = "Two processes at the same layer in different networks are called dash process."
StrQuestions(10) = "Usually the dash layer manages the routing of packet in a network."

stroptions(1) = "Quantizing error or Quantizing noise"
stroptions(2) = "Delay"
stroptions(3) = "Manchester II (biphase level)"
stroptions(4) = "Source"
stroptions(5) = "Network"
stroptions(6) = "FSK"
stroptions(7) = "64"
stroptions(8) = "Peer"
stroptions(9) = "Half duplex"
stroptions(10) = "QAM (Quadrature Amplitude Modulation)"

stranswers(1) = "QAM (Quadrature Amplitude Modulation)"
stranswers(2) = "Source"
stranswers(3) = "64"
stranswers(4) = "FSK"
stranswers(5) = "Quantizing error or Quantizing noise"
stranswers(6) = "Manchester II (biphase level)"
stranswers(7) = "Half duplex"
stranswers(8) = "Delay"
stranswers(9) = "Peer"
stranswers(10) = "Network"

End Sub

Private Sub Fill_in_the_Blanks_Paper3()
StrQuestions(1) = "The layer ‘N’ that provides service is called dash."
StrQuestions(2) = "Token bus is an example of dash layer protocol."
StrQuestions(3) = "A logic high in RS-232C is a voltage between -3V and dash under load."
StrQuestions(4) = "Error detection and correction are the functions of dash layer."
StrQuestions(5) = "Effectiveness of a code for error detection is usually measured in terms of dash distance."
StrQuestions(6) = "Sliding window protocol is a dash control protocol."
StrQuestions(7) = "The receiver window size is normally kept bigger than the maximum sequence number in case of dash protocol."
StrQuestions(8) = "LAPB (Link Access Protocol) is a dash standard."
StrQuestions(9) = "IEEE 802.2 is a dash protocol."
StrQuestions(10) = "IEEE dash represents token ring protocol."

stroptions(1) = "Sliding window with selective repeat"
stroptions(2) = "LLC( Logical Link Control)"
stroptions(3) = "Data-link"
stroptions(4) = "Flow"
stroptions(5) = "-15V"
stroptions(6) = "CCITT"
stroptions(7) = "Hamming"
stroptions(8) = "Service provider"
stroptions(9) = "MAC(Medium Access Control) / Data Link Layer"
stroptions(10) = "802.8"

stranswers(1) = "Service provider"
stranswers(2) = "MAC(Medium Access Control) / Data Link Layer"
stranswers(3) = "-15V"
stranswers(4) = "Data-link"
stranswers(5) = "Hamming"
stranswers(6) = "Flow"
stranswers(7) = "Sliding window with selective repeat"
stranswers(8) = "CCITT"
stranswers(9) = "LLC( Logical Link Control)"
stranswers(10) = "802.8"

End Sub

Private Sub Fill_in_the_Blanks_Paper4()
StrQuestions(1) = "LLC protocol data unit is commonly known as dash."
StrQuestions(2) = "In pure ALOHA the frames are of dash length."
StrQuestions(3) = "Slotted ALOHA achieves double efficient by reducing the packet dash probability."
StrQuestions(4) = "Ethernet has a dash topology."
StrQuestions(5) = "The maximum number of repeaters that can be connected for the largest configuration of Ethernet is dash."
StrQuestions(6) = "FDDI is a high performance token dash LAN runs at 100 MPS."
StrQuestions(7) = "dash Fibers are used in FDDI."
StrQuestions(8) = "Using repeaters, the maximum cable length in Ethernet networks could be dash."
StrQuestions(9) = "Token bus employs dash encoding."
StrQuestions(10) = "Network addresses are called dash."

stroptions(1) = "Multimode"
stroptions(2) = "2.5 km"
stroptions(3) = "Analog modulation"
stroptions(4) = "Ring"
stroptions(5) = "Collision"
stroptions(6) = "Four"
stroptions(7) = "Fixed"
stroptions(8) = "LAN frame"
stroptions(9) = "Bus"
stroptions(10) = "NSAPS (Network Service Access Points)"

stranswers(1) = "LAN frame"
stranswers(2) = "Fixed"
stranswers(3) = "Collision"
stranswers(4) = "Bus"
stranswers(5) = "Four"
stranswers(6) = "Ring"
stranswers(7) = "Multimode"
stranswers(8) = "2.5 km"
stranswers(9) = "Analog modulation"
stranswers(10) = "NSAPS (Network Service Access Points)"
End Sub


Private Sub Fill_in_the_Blanks_Paper5()
StrQuestions(1) = "One of the major functions of network layer is to dash packets form source node to the destination node."
StrQuestions(2) = "Congestion in a network can be controlled by introducing dash control."
StrQuestions(3) = "X.25 normally assumes the use of dash as the physical level standard."
StrQuestions(4) = "In X.25 packet transferred across the DTE/DCE interface begins with a dash header."
StrQuestions(5) = "X.25 is a connection dash protocol."
StrQuestions(6) = "ISDN is designed to provide circuit switched and packet switched connections at dash kbps."
StrQuestions(7) = "The capacity of ISDN D channel is dash kbps."
StrQuestions(8) = "Basic access rate in ISDN channel is dash kbps"
StrQuestions(9) = "The ISDN physical layer frame format is dash bits."
StrQuestions(10) = "For the 2.048 mbps data rate, the ISDN channel structure will be as follows: dash numbers of B channels plus one 64 kbps D channel."

stroptions(1) = "64"
stroptions(2) = "144 or 192"
stroptions(3) = "3 Bytes"
stroptions(4) = "30"
stroptions(5) = "X.21"
stroptions(6) = "48"
stroptions(7) = "Flow"
stroptions(8) = "16 or 64"
stroptions(9) = "Route"
stroptions(10) = "Oriented"

stranswers(1) = "Route"
stranswers(2) = "Flow"
stranswers(3) = "X.21"
stranswers(4) = "3 Bytes"
stranswers(5) = "Oriented"
stranswers(6) = "64"
stranswers(7) = "16 or 64"
stranswers(8) = "144 or 192"
stranswers(9) = "48"
stranswers(10) = "30"

End Sub
Private Sub Fill_in_the_Blanks_Paper6()
StrQuestions(1) = "The bandwidth of an audio signal is dash."
StrQuestions(2) = "The bandwidth of a video signal is dash."
StrQuestions(3) = "Electromagnetic waves having frequencies between dash are termed as microwaves."
StrQuestions(4) = "The VHF band of radio frequencies is dash "
StrQuestions(5) = "The light accepting capability of an optical fiber is measured by its dash"
StrQuestions(6) = "Analog to digital conversion introduces dash noise."
StrQuestions(7) = "For enabling perfect reconstruction of an analog signal from its samples the sampling frequency should be greater than the dash frequency."
StrQuestions(8) = "Intermodal dispersion is absent in dash fibers."
StrQuestions(9) = "dash fiber suffers from high signal dispersion."
StrQuestions(10) = "dash reduces the probability of error due to transmission"

stroptions(1) = "Quantization"
stroptions(2) = "30 - 3000 MHz"
stroptions(3) = "Coding"
stroptions(4) = "Single-mode"
stroptions(5) = "6 MHz"
stroptions(6) = "Numerical aperture"
stroptions(7) = "3 KHz"
stroptions(8) = "Multi mode"
stroptions(9) = "Nyquist"
stroptions(10) = "30-300 MHz."

stranswers(1) = "3 KHz"
stranswers(2) = "6 MHz"
stranswers(3) = "30 - 3000 MHz"
stranswers(4) = "30-300 MHz."
stranswers(5) = "Numerical aperture"
stranswers(6) = "Quantization"
stranswers(7) = "Nyquist"
stranswers(8) = "Single-mode"
stranswers(9) = "Multi mode"
stranswers(10) = "Coding"

End Sub

Private Sub Fill_in_the_Blanks_Paper7()
StrQuestions(1) = "Highest  data rates can be supported by dash."
StrQuestions(2) = "The process by which a signal is made compatible with a transmission channel is called dash."
StrQuestions(3) = "dash allows data to be transmitted using a smaller bandwidth."
StrQuestions(4) = "A graded index fiber is a dash fiber."
StrQuestions(5) = "A modem has a modulator and a dash."
StrQuestions(6) = "Hamming codes are used for error dash."
StrQuestions(7) = "Parity checking is used for error dash."
StrQuestions(8) = "Many modern modems for data communication use digital dash modulation."
StrQuestions(9) = "In FDM consecutive channels are separated by dash."
StrQuestions(10) = "Digital signals can also use dash in addition o analog multiplexing schemes."

stroptions(1) = "compression"
stroptions(2) = "Phase"
stroptions(3) = "Modulation"
stroptions(4) = "Guard bands"
stroptions(5) = "Correction"
stroptions(6) = "Multi mode"
stroptions(7) = "Single mode"
stroptions(8) = "TDM"
stroptions(9) = "Demodulator"
stroptions(10) = "Detection"

stranswers(1) = "Single mode"
stranswers(2) = "Modulation"
stranswers(3) = "compression"
stranswers(4) = "Multi mode"
stranswers(5) = "Demodulator"
stranswers(6) = "Correction"
stranswers(7) = "Detection"
stranswers(8) = "Phase"
stranswers(9) = "Guard bands"
stranswers(10) = "TDM"

End Sub

Private Sub Fill_in_the_Blanks_Paper8()
StrQuestions(1) = "The total data rate of  a time division multiplexed signal exceeds the sum of he data rates of the individual channels because of the inclusion of dash bits."
StrQuestions(2) = "The sequence of operations in dash are sampling quantization and encoding."
StrQuestions(3) = "The dash process in a PCM modulator introduces an error in the original signal and the final demodulated signal."
StrQuestions(4) = "PCM stands for dash code modulation."
StrQuestions(5) = "dash shift keying can be obtained by multiplying a carrier wave by a polar NRZ(+1,-1) digital signal."
StrQuestions(6) = "dash shift keying can be obtained by applying a digital signal in NRZ format to a VCO."
StrQuestions(7) = "A differential encoder followed by a BPSK modulator generates dash modulation."
StrQuestions(8) = "dash shift  keying modulation is very susceptible to noise."
StrQuestions(9) = "A different carrier frequency for each channel is used in dash."
StrQuestions(10) = "The carrier recovery circuit is not required for a dash encoded PSK."

stroptions(1) = "Frequency "
stroptions(2) = "FDM"
stroptions(3) = "Pulse"
stroptions(4) = "Amplitude"
stroptions(5) = "PCM"
stroptions(6) = "Phase"
stroptions(7) = "Quantization"
stroptions(8) = "Synchronization"
stroptions(9) = "Differentially"
stroptions(10) = "DPSK"

stranswers(1) = "Synchronization"
stranswers(2) = "PCM"
stranswers(3) = "Quantization"
stranswers(4) = "Pulse"
stranswers(5) = "Phase"
stranswers(6) = "Frequency "
stranswers(7) = "DPSK"
stranswers(8) = "FDM"
stranswers(9) = "Differentially"
stranswers(10) = "Amplitude"

End Sub

Private Sub Fill_in_the_Blanks_Paper9()
StrQuestions(1) = "For the case of a unipolar NRZ binary signal, dash is known as on -off keying."
StrQuestions(2) = "Most of the existing fiber optic communication systems use dash keying."
StrQuestions(3) = "An dash source is most frequently used in multimode fiber because of its large core diameter."
StrQuestions(4) = "High data rate monomode fiber optic systems largely depend upon dash sources."
StrQuestions(5) = "An optical signal travels through the fiber waveguide by total internal reflection at the core to dash interface."
StrQuestions(6) = "An dash photodiode has a higher sensitivity than PIN photodiode."
StrQuestions(7) = "Media access control in FDDI is based on dash passing."
StrQuestions(8) = "FDDI uses token passing on dash topology."
StrQuestions(9) = "The local telephone exchanges are usually networked by using a higher level switching known as the dash."
StrQuestions(10) = "A dash can handle local calls within a company or factory without intervening with the PSTN."

stroptions(1) = "Trunk autamatic exchange(TAX)"
stroptions(2) = "Aalanche"
stroptions(3) = "Amplitude"
stroptions(4) = "Cladding"
stroptions(5) = "Ring"
stroptions(6) = "LED"
stroptions(7) = "PABX"
stroptions(8) = "ASK"
stroptions(9) = "Token"
stroptions(10) = "Laser diode"

stranswers(1) = "ASK"
stranswers(2) = "Amplitude"
stranswers(3) = "LED"
stranswers(4) = "Laser diode"
stranswers(5) = "Cladding"
stranswers(6) = "Aalanche"
stranswers(7) = "Token"
stranswers(8) = "Ring"
stranswers(9) = "Trunk autamatic exchange(TAX)"
stranswers(10) = "PABX"

End Sub

Private Sub Fill_in_the_Blanks_Paper10()
StrQuestions(1) = "Currently high bandwidth digital telephony is supported within local areas supported by a dash."
StrQuestions(2) = "A large number of telephone channels are dash over a long distance trunk."
StrQuestions(3) = "The transmission circuits that interconnect two different switching centers in the PSTN are known as dash circuits."
StrQuestions(4) = "A computer sends digital data over the telephone network using a dash."
StrQuestions(5) = "Run length coding for compression is used in digital dash."
StrQuestions(6) = "A FAX machine includes dash , coding data modem, printer and telephone dialing facilities."
StrQuestions(7) = "Satellite transponders use dash amplifiers."
StrQuestions(8) = "A dash satellite needs no tracking by earth stations."
StrQuestions(9) = "A satellite transponder must be equipped with a dash receiver."
StrQuestions(10) = "Future satellite communications may use dash multiple access schemes."

stroptions(1) = "Scanning "
stroptions(2) = "Low noise"
stroptions(3) = "FAX"
stroptions(4) = "Trunk"
stroptions(5) = "Wide band"
stroptions(6) = "PABX"
stroptions(7) = "Modem"
stroptions(8) = "Code division"
stroptions(9) = "Geostationary"
stroptions(10) = "Multiplexed"

stranswers(1) = "PABX"
stranswers(2) = "Multiplexed"
stranswers(3) = "Trunk"
stranswers(4) = "Modem"
stranswers(5) = "FAX"
stranswers(6) = "Scanning "
stranswers(7) = "Low noise"
stranswers(8) = "Geostationary"
stranswers(9) = "Wide band"
stranswers(10) = "Code division"

End Sub

Private Sub Matching_Columns_Paper1()
strcolumnA(1) = "Multiplexing and data concentration techniques arethe means ofsharing"
strcolumnA(2) = "Frequency hopping and spread spectrum  techniquesare involved in"
strcolumnA(3) = "If a channel cannot transmit below f1Hz and above f2Hz it is called"
strcolumnA(4) = "Maximum data rate gien by Nyauist theorem is not practically  achievable mainly due to"
strcolumnA(5) = "Bounded media for physical channels"
strcolumnA(6) = "LOS (line of sight) transmission"
strcolumnA(7) = "Satellite communication"
strcolumnA(8) = "QPSK"
strcolumnA(9) = "INSAT"
strcolumnA(10) = "Transmitter andreceier"

strcolumnB(1) = "C band ( 4 – 6 GHz)"
strcolumnB(2) = "Geostationary satellite"
strcolumnB(3) = "Channel noise"
strcolumnB(4) = "SHF band (3 to 30 GHz)"
strcolumnB(5) = "CDM (Code Diision Multiplexing)"
strcolumnB(6) = "Transponder"
strcolumnB(7) = "Digital modulation"
strcolumnB(8) = "Twisted pair wire"
strcolumnB(9) = "Pasband"
strcolumnB(10) = "Communication channel"

stranswers(1) = "Communication channel"
stranswers(2) = "CDM (Code Diision Multiplexing)"
stranswers(3) = "Pasband"
stranswers(4) = "Channel noise"
stranswers(5) = "Twisted pair wire"
stranswers(6) = "SHF band (3 to 30 GHz)"
stranswers(7) = "C band ( 4 – 6 GHz)"
stranswers(8) = "Digital modulation"
stranswers(9) = "Geostationary satellite"
stranswers(10) = "Communication channel"
End Sub

Private Sub Matching_Columns_Paper2()
strcolumnA(1) = "Coherent detection"
strcolumnA(2) = "DPCM"
strcolumnA(3) = "Nyquist rate"
strcolumnA(4) = "Full duplex operation is possible over single transmission channel"
strcolumnA(5) = "Manchester code"
strcolumnA(6) = "IMP (Interface message processor)"
strcolumnA(7) = "Star topology"
strcolumnA(8) = "High utilization of outgoing communication"
strcolumnA(9) = "Message transport  across  thesubnet"
strcolumnA(10) = "Headers and interfaces"

strcolumnB(1) = "Concentrator"
strcolumnB(2) = "Network components"
strcolumnB(3) = "Point to point  subnet"
strcolumnB(4) = "ARPANET"
strcolumnB(5) = "Minimum sampling rate"
strcolumnB(6) = "Demodulation"
strcolumnB(7) = "Store and forward  switching"
strcolumnB(8) = "FDM"
strcolumnB(9) = "Source coding"
strcolumnB(10) = "Baseband encoding"

stranswers(1) = "Demodulation"
stranswers(2) = "Source coding"
stranswers(3) = "Minimum sampling rate"
stranswers(4) = "FDM"
stranswers(5) = "Baseband encoding"
stranswers(6) = "ARPANET"
stranswers(7) = "Point to point  subnet"
stranswers(8) = "Concentrator"
stranswers(9) = "Store and forward  switching"
stranswers(10) = "Baseband encoding"

End Sub

Private Sub Matching_Columns_Paper3()
strcolumnA(1) = "7-layer architecture"
strcolumnA(2) = "TCP(Transmission control protocol)"
strcolumnA(3) = "FDDI( Fiber Distributed Data Interface)"
strcolumnA(4) = "PDU and SDU"
strcolumnA(5) = "X.21"
strcolumnA(6) = "DQDB"
strcolumnA(7) = "RC,LRC, CRC"
strcolumnA(8) = "Piggybacking"
strcolumnA(9) = "Upper bound on the number of frames"
strcolumnA(10) = "Link layer protocol"

strcolumnB(1) = "Fiber Optic Network"
strcolumnB(2) = "Bit oriented protocoo (BOPs)"
strcolumnB(3) = "MAN protocol"
strcolumnB(4) = "Connection oriented protocol"
strcolumnB(5) = "Stop and wait ARQ"
strcolumnB(6) = "Physical layer protocol"
strcolumnB(7) = "ISO-OSI"
strcolumnB(8) = "Window size"
strcolumnB(9) = "Error detecting codes"
strcolumnB(10) = "Peer communication"

stranswers(1) = "ISO-OSI"
stranswers(2) = "Connection oriented protocol"
stranswers(3) = "Fiber Optic Network"
stranswers(4) = "Peer communication"
stranswers(5) = "Physical layer protocol"
stranswers(6) = "MAN protocol"
stranswers(7) = "Error detecting codes"
stranswers(8) = "Stop and wait ARQ"
stranswers(9) = "Window size"
stranswers(10) = "Bit oriented protocoo (BOPs)"

End Sub

Private Sub Matching_Columns_Paper4()
strcolumnA(1) = "BISYNC"
strcolumnA(2) = "TDM, FDM"
strcolumnA(3) = "SPADE"
strcolumnA(4) = "Reservatioon ALOHA"
strcolumnA(5) = "Ethernet"
strcolumnA(6) = "CSMA/CD"
strcolumnA(7) = "RIU (Ring Interface Unit)"
strcolumnA(8) = "LED as source "
strcolumnA(9) = "N-connect"
strcolumnA(10) = "Multipath routing"

strcolumnB(1) = "FDDI"
strcolumnB(2) = "Collision resolution protocol"
strcolumnB(3) = "Satellite networks"
strcolumnB(4) = "Network layer primitive"
strcolumnB(5) = "Intelsat satellite"
strcolumnB(6) = "Character oriented protocol"
strcolumnB(7) = "Network layer routing"
strcolumnB(8) = "Token  ring"
strcolumnB(9) = "Local Area networks"
strcolumnB(10) = "Channel allocation  scheme"

stranswers(1) = "Character oriented protocol"
stranswers(2) = "Channel allocation  scheme"
stranswers(3) = "Intelsat satellite"
stranswers(4) = "Satellite networks"
stranswers(5) = "Local Area networks"
stranswers(6) = "Collision resolution protocol"
stranswers(7) = "Token  ring"
stranswers(8) = "FDDI"
stranswers(9) = "Network layer primitive"
stranswers(10) = "Network layer routing"

End Sub

Private Sub Matching_Columns_Paper5()
strcolumnA(1) = "Needle packet"
strcolumnA(2) = "Choke packets"
strcolumnA(3) = "LAPB(Link Access Procedure Balanced)"
strcolumnA(4) = "irtual Call"
strcolumnA(5) = "The rate of transmission of symbols"
strcolumnA(6) = "Videotex, Teletex"
strcolumnA(7) = "Digital bit pipe"
strcolumnA(8) = "H11,H12"
strcolumnA(9) = "SS#5, SS#6"
strcolumnA(10) = "TSAP, TCEP"

strcolumnB(1) = "Virtual circuit service of X.25"
strcolumnB(2) = "ISDN channel"
strcolumnB(3) = "ISDN"
strcolumnB(4) = "Transport service primitives"
strcolumnB(5) = "ISDN services"
strcolumnB(6) = "Data Link control protocol"
strcolumnB(7) = "Centralized routing"
strcolumnB(8) = "ISDN signaling schemes"
strcolumnB(9) = "Baud"
strcolumnB(10) = "Congestion control mechanism"

stranswers(1) = "Centralized routing"
stranswers(2) = "Congestion control mechanism"
stranswers(3) = "Data Link control protocol"
stranswers(4) = "Virtual circuit service of X.25"
stranswers(5) = "Baud"
stranswers(6) = "ISDN services"
stranswers(7) = "ISDN"
stranswers(8) = "ISDN channel"
stranswers(9) = "ISDN signaling schemes"
stranswers(10) = "Transport service primitives"

End Sub

Private Sub Matching_Columns_Paper6()
strcolumnA(1) = "A signal taking discrete alues at discrete interals oftime"
strcolumnA(2) = "Leased and switchedservices"
strcolumnA(3) = "A tansmission channel immune to electromagnetic interference"
strcolumnA(4) = "A metallic transmission media that can be used at a frequency of 10 MHz."
strcolumnA(5) = "A metallic transmission media that can be used for frequencies not exceeding 10 raise to power 5 Hz"
strcolumnA(6) = "A wired circuit connecting a telephone subscriber with the local exchange"
strcolumnA(7) = "A metallic  transmission media which is shielded from electrical noise"
strcolumnA(8) = "Microwave frequencies"
strcolumnA(9) = "Frequencies than can be transmitted over acoazial cable"
strcolumnA(10) = "Frequencies that can be transmitted over a twisted wire pair"

strcolumnB(1) = "Optical fiber"
strcolumnB(2) = "30 – 300 MHz"
strcolumnB(3) = "Local loop"
strcolumnB(4) = "0 – 100 KHz"
strcolumnB(5) = "Coaxial cable "
strcolumnB(6) = "Digital signal"
strcolumnB(7) = "60 KHz – 100 MHz"
strcolumnB(8) = "Coaxial cable"
strcolumnB(9) = "Twisted wire pair"
strcolumnB(10) = "ISDN"

stranswers(1) = "Digital signal"
stranswers(2) = "ISDN"
stranswers(3) = "Optical fiber"
stranswers(4) = "Coaxial cable "
stranswers(5) = "Twisted wire pair"
stranswers(6) = "Local loop"
stranswers(7) = "Coaxial cable"
stranswers(8) = "Coaxial cable"
stranswers(9) = "60 KHz – 100 MHz"
stranswers(10) = "0 – 100 KHz"

End Sub


Private Sub Matching_Columns_Paper7()
strcolumnA(1) = "Voice band freuqncies"
strcolumnA(2) = "A signal that consists of distinct tones at distincttime intervals"
strcolumnA(3) = "A modulation scheme with constant amplitude"
strcolumnA(4) = "A modulation scheme with constant carrier frequency"
strcolumnA(5) = "Conversion of binary ASK to FSK"
strcolumnA(6) = "Tracking the carrier frequency"
strcolumnA(7) = "Transmission of data qithout carrier modulation"
strcolumnA(8) = "error correcting code"
strcolumnA(9) = "Digital phase modulation scheme that can be demodulated non coherently"
strcolumnA(10) = "Synchronization word"

strcolumnB(1) = "TDM"
strcolumnB(2) = "PLL"
strcolumnB(3) = "DPSK"
strcolumnB(4) = "FSK"
strcolumnB(5) = "Convolutional code"
strcolumnB(6) = "FSK signal"
strcolumnB(7) = "PSK"
strcolumnB(8) = "VCO"
strcolumnB(9) = "Baseband communication"
strcolumnB(10) = "300 Hz – 3400 Hz"

stranswers(1) = "300 Hz – 3400 Hz"
stranswers(2) = "FSK signal"
stranswers(3) = "FSK"
stranswers(4) = "PSK"
stranswers(5) = "VCO"
stranswers(6) = "PLL"
stranswers(7) = "Baseband communication"
stranswers(8) = "Convolutional code"
stranswers(9) = "DPSK"
stranswers(10) = "TDM"

End Sub
Private Sub Matching_Columns_Paper8()
strcolumnA(1) = "Sharing of transmission media"
strcolumnA(2) = "Errr detection scheme"
strcolumnA(3) = "Combination of ASK and PSK"
strcolumnA(4) = "Single error correcting block code"
strcolumnA(5) = "Adding redundancy to tachle channel errors "
strcolumnA(6) = "Viterbi algorithm"
strcolumnA(7) = "Two wire to four wire conversion"
strcolumnA(8) = "DTMF"
strcolumnA(9) = "Store and forward system"
strcolumnA(10) = "Intermodal dispersion"

strcolumnB(1) = "Tone dialing"
strcolumnB(2) = "Convolutional code"
strcolumnB(3) = "Channel coding"
strcolumnB(4) = "CRC"
strcolumnB(5) = "Multimode fiber"
strcolumnB(6) = "Hybrid"
strcolumnB(7) = "Multiplexing"
strcolumnB(8) = "Packet switched network"
strcolumnB(9) = "QAM"
strcolumnB(10) = "Hamming code"

stranswers(1) = "Multiplexing"
stranswers(2) = "CRC"
stranswers(3) = "QAM"
stranswers(4) = "Hamming code"
stranswers(5) = "Channel coding"
stranswers(6) = "Convolutional code"
stranswers(7) = "Hybrid"
stranswers(8) = "Tone dialing"
stranswers(9) = "Packet switched network"
stranswers(10) = "Multimode fiber"

End Sub

Private Sub Matching_Columns_Paper9()
strcolumnA(1) = "Reduced intermodal dispersion in optical fiber"
strcolumnA(2) = "Spontaneous  light emission"
strcolumnA(3) = "A high monochromatic and directional optical source"
strcolumnA(4) = "Guard band"
strcolumnA(5) = "Digitization of speech"
strcolumnA(6) = "Conversion of an analog signal to discrete levels"
strcolumnA(7) = "Two level quantization of image"
strcolumnA(8) = "Digital FAX compression technique"
strcolumnA(9) = "Equipment channel that connects the receie antenna with the transmit antenna of a satellite"
strcolumnA(10) = "A low low permanent connection between two optical  fiber"

strcolumnB(1) = "FAX"
strcolumnB(2) = "Transponder"
strcolumnB(3) = "Run length coding"
strcolumnB(4) = "Quantization"
strcolumnB(5) = "Laser"
strcolumnB(6) = "Splice"
strcolumnB(7) = "FCM"
strcolumnB(8) = "PCM"
strcolumnB(9) = "Graded index"
strcolumnB(10) = "LED"

stranswers(1) = "Graded index"
stranswers(2) = "LED"
stranswers(3) = "Laser"
stranswers(4) = "FCM"
stranswers(5) = "PCM"
stranswers(6) = "Quantization"
stranswers(7) = "FAX"
stranswers(8) = "Run length coding"
stranswers(9) = "Transponder"
stranswers(10) = "Splice"

End Sub

Private Sub Matching_Columns_Paper10()
strcolumnA(1) = "Phase reversal of a carrier according to a binary signal"
strcolumnA(2) = "Channel coding"
strcolumnA(3) = "A waveguide for optical signals"
strcolumnA(4) = "Undersampling"
strcolumnA(5) = "Anti aliasing filters"
strcolumnA(6) = "SDM, FDM"
strcolumnA(7) = "RS- 232"
strcolumnA(8) = "RS- 232"
strcolumnA(9) = "conversion of optical signal to electrical signal"
strcolumnA(10) = "Sever bandwidth limitation"

strcolumnB(1) = "Satellite communication"
strcolumnB(2) = "Aliasing"
strcolumnB(3) = "Serial data transfer"
strcolumnB(4) = "Bandlimiting"
strcolumnB(5) = "Error correction"
strcolumnB(6) = "Photodiode"
strcolumnB(7) = "Analog multiplexing"
strcolumnB(8) = "BPSK"
strcolumnB(9) = "Optical fiber"
strcolumnB(10) = "PABX"

stranswers(1) = "BPSK"
stranswers(2) = "Error correction"
stranswers(3) = "Optical fiber"
stranswers(4) = "Aliasing"
stranswers(5) = "Bandlimiting"
stranswers(6) = "Analog multiplexing"
stranswers(7) = "Serial data transfer"
stranswers(8) = "PABX"
stranswers(9) = "Photodiode"
stranswers(10) = "Intersymbol interference "

End Sub

Private Sub TrueorFalse_Paper1()
StrQuestions(1) = "In signalling bit rate and baud rate are always equal."
StrQuestions(2) = "Asyncfhronous transmission is less efficient than synchrounous transmission."
StrQuestions(3) = "For base band transmission a modem is required"
StrQuestions(4) = "Digital signals can also be transmitted over analos channels"
StrQuestions(5) = "Digital multiplexing is simpler than analod multiplixing"
StrQuestions(6) = "Bandwidth limitation caues signal shape distortion"
StrQuestions(7) = "White noise is a natural noise"
StrQuestions(8) = "Attenuation in channel changes the shape of the signal"
StrQuestions(9) = "Twisted wire pair is susceptible to a great deal of electrical interference"
StrQuestions(10) = "Coaxial cable can provide data rates over 10 Mbps and frequencies up to 400 MHz.        "

stranswers(1) = "False"
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

Private Sub TrueorFalse_Paper2()
StrQuestions(1) = "Fiber optic cables have low MTBF(mean Time Between Failures)"
StrQuestions(2) = "Intermodulation intererrence is present in FDMA"
StrQuestions(3) = "PSK has a practical limit at 4.8 kbps."
StrQuestions(4) = "RZ and NRZ are two basic classes of encoding scheme in baseband transmission"
StrQuestions(5) = "NRZ encoding has inherent error correcting capabilities."
StrQuestions(6) = "Polling and concentration are either statistical TDM or statistical multiplexing"
StrQuestions(7) = "Message switching provides end to end transpost of messages"
StrQuestions(8) = "Network layer in OSI architecturer provides a true end to end connection"
StrQuestions(9) = "Data compression is performed usuall in the presentation"
StrQuestions(10) = "FTAM(File Transfer Access and Management) is a presentation layer protocol"

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"

End Sub
Private Sub TrueorFalse_Paper3()
StrQuestions(1) = "FRP, SMTP, and TELNET are standard protocols for the application layer"
StrQuestions(2) = "The maximum permitted data rate in case of 232-C is 20 kbps and maximum cable length is 15 meter."
StrQuestions(3) = "'Framming' refers to the process of encapsulation and decapsulation"
StrQuestions(4) = "Stop and wait protocol is generally used to detectframes with error(s) at the receiving DLC module"
StrQuestions(5) = "HDLC(Hig level data link control) protocol ) uses selective repeat strategies"
StrQuestions(6) = "For stop and wait protocol the window size is 1"
StrQuestions(7) = "Character oriented protocols are best suited for synchronous transmission"
StrQuestions(8) = "HDLC(High level data link control) protocol is a bit oriented protocol"
StrQuestions(9) = "IEEE 802.3 describes CSMA/CD protocol"
StrQuestions(10) = "IEEE 802.4 describes token bus protocol"

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
Private Sub TrueorFalse_Paper4()
StrQuestions(1) = "The channel utilization of the pure ALOHA is 37% (approx)."
StrQuestions(2) = "In Ethernet, the permitted maximum cable length in a single segment is 500 meters."
StrQuestions(3) = "Ethernet protocol is a n acknowledgement type of protocol."
StrQuestions(4) = "FDDI can spread over distance of up to 2000 kms"
StrQuestions(5) = "FDDI cbaling consists of wo fiber rings"
StrQuestions(6) = "Manchester encoding is used in FDDI."
StrQuestions(7) = "Ethernet employs manchester encoding"
StrQuestions(8) = "Network throughput can be increased by reducing meand packet delay"
StrQuestions(9) = "Adaptive algorithms do not change their routing decisions"
StrQuestions(10) = "Static multipath routing increases the reliability of the subnet"

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "true"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"

End Sub
Private Sub TrueorFalse_Paper5()
StrQuestions(1) = "The lockup phenomenon in congestion is known as deadlocks"
StrQuestions(2) = "Modulo 16 is the basic frame format in X.25 link level protocol"
StrQuestions(3) = "A virtual call is a dynamically stablished virtual circuit using a call set up and call clearing procedure"
StrQuestions(4) = "Flos control and error control in X.25 are implemented using sequence numbers"
StrQuestions(5) = "The network layer provides only connectionless services to the trahsport layer"
StrQuestions(6) = "ISDN supports only switched applications"
StrQuestions(7) = "ISDN B channel has a capacity of 64 kbops"
StrQuestions(8) = "SS#7 signaling has three layers"
StrQuestions(9) = "X.400 is a mai lsystem protocol"
StrQuestions(10) = "The structure of the directory services in X.500 is similar to that of X.400 MHS"

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"

End Sub
Private Sub TrueorFalse_Paper6()
StrQuestions(1) = "For satellite communication phase modulation schemas are better than amplitude modulation"
StrQuestions(2) = "Coherent demodulation schemes usually have simpler implementations than non coherent demodulation schemes"
StrQuestions(3) = "An FSK receiver must icorporate carrier recovery"
StrQuestions(4) = "A PSK reveiver must incorporate carrier recovery"
StrQuestions(5) = "For the same modulation scheme , if a non coherent detector exists, it sill prodice lower errors than a coherent detector."
StrQuestions(6) = "The refractive index at the core of an optical fiber is always higher than the cladding"
StrQuestions(7) = "The refractive index difference for a single mode fiber is smaller than that of a multimode fiber"
StrQuestions(8) = "A laser diode is perfectly monochromatic"
StrQuestions(9) = "The maximum length of a multimode optical fiber link is usuall limited buy the loss in the fiber"
StrQuestions(10) = "Pulse broadening in an optical fiber due to dispersion depends on the quality of the optical source"

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "True"

End Sub
Private Sub TrueorFalse_Paper7()
StrQuestions(1) = "Multimode fibres have a lower numerical aperture than single mode fibers"
StrQuestions(2) = "Pure silica does not cause any attenuation of optical signals"
StrQuestions(3) = "An APD receiver always produces better SNR than a PIN receiver"
StrQuestions(4) = "Shot noice is caused during the presence of an optical signal only"
StrQuestions(5) = "Thermal oise is independent of the optical signal"
StrQuestions(6) = "Single mode fibers do not have any dispersion"
StrQuestions(7) = "At the zero dispersio wavelength it is possible to rnsmit dat at infinitely high data rates"
StrQuestions(8) = "A surface emitting ype of LED has a more diretional radiation pattern than an edge limiting type  of LED"
StrQuestions(9) = "FDDI can support both synchronous and asynchrounous services"
StrQuestions(10) = "FDDI is not suited for real time applications"

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "Treu"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "False"

End Sub
Private Sub TrueorFalse_Paper8()
StrQuestions(1) = "An FSK signal is less susceptible to noise than ASK"
StrQuestions(2) = "A DPSK signal can be demodulated without recovering the carrier"
StrQuestions(3) = "FSK requires more bandwidth than PSK"
StrQuestions(4) = "In digital FAX transmitter the scanned image is digitaized using PCM"
StrQuestions(5) = "FDM does not require any bandwidth overhead whereas TDM does."
StrQuestions(6) = "Telephone subscribers are conneted to the local exchange in a bus topology."
StrQuestions(7) = "In between two trunk automatic exchanges , there are a many wire cables as the number of channels between them"
StrQuestions(8) = "PCM used in digital telephoney uses a no uniform quantizer"
StrQuestions(9) = "In order to provide error correcting capability some redundancy myst be introduced in the transmitted dat"
StrQuestions(10) = "Manchester coding is a powerful error correcting  scheme"

stranswers(1) = "True"
stranswers(2) = "True"
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
StrQuestions(1) = "Manchester coding makes clock recovery easier"
StrQuestions(2) = "The bandwidth of a Manchester coded signal is twice to that of a NRZ coded signal"
StrQuestions(3) = "Hamming codes can be used for detecting errors only, but not for error correction"
StrQuestions(4) = "CRC(Cyclic redundancy check) can detect as well as correct errors."
StrQuestions(5) = "DPSK modulation uses the information in two consecutive bits"
StrQuestions(6) = "A geostationary satellite must have zero inclination"
StrQuestions(7) = "There is only one geostationary orbit"
StrQuestions(8) = "Satellite transponders use different polarizations in adjacent transponder channels"
StrQuestions(9) = "Non uniform quantization in PCM increases the signal to noise ration for voice transmission"
StrQuestions(10) = "Quantization noise in a PCM system increase down the number of levels of quantization is increased"

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
Private Sub TrueorFalse_Paper10()
StrQuestions(1) = "Given an idela lowpass chanel of bandwidth B, data transmission at a rate R >- 2B band will always cause intersymbol interference"
StrQuestions(2) = "PCM is not very efficient for high SNR requirements"
StrQuestions(3) = "A digital communication system is always more nois than an analog communication system because of the quantization process thoug hi t is more reliable"
StrQuestions(4) = "FDM is more appropriate than TDM in multiplexing digital signals having variable data rates"
StrQuestions(5) = "When the same number of channels are to be multiplexed fortransmission, FDM always requires a greater bandwidth than TDM."
StrQuestions(6) = "Carrier modulation is used for frequency division multiplexing."
StrQuestions(7) = "Carrier modulation does not change the bandwidth of the signal"
StrQuestions(8) = "PCM allows transmission of an analog signal to be transmitted over a digital channel"
StrQuestions(9) = "Interference can becompletely removed by good filtering."
StrQuestions(10) = "In tone dialing telephones, 8 different tones are used to endcode all thedialed digits."

stranswers(1) = "True"
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

Private Sub questions_paper1()
StrQuestions(1) = "a. A system is operating a t the Nyquist rate of 2 bits per Hz of BW. Dtermine the theoretical average signal to noise power ratio. Express the same in dB." & Chr(13) & _
                  "b. Calculate the channel capacity of a channel which has a BW of 10 MHz and a dB value of 30."
StrQuestions(2) = "a. Consider a PCM system in which 30 channels are to be time division multiplexed. The BW of each channel is 3 KHz. The sampling rate is 33.33 % higher than the theoretical minimum, and 8 bits are used for each sample. Determine the required bit rate and find the minimum required transmission band width." & Chr(13) & _
                  "b.  Define simplex, half duplex and full duplex start bit, 2 stop bits, 1 parity and 8 data bits. Calculate used data rate."
StrQuestions(3) = "A telephone line has a pass band from 300 Hz to 3.3 KHz. What data rate is permitted if a 30 dB SNR is desired."
StrQuestions(4) = "a. For designing a backbone computer network which one of the following will you choose as the topology of your network and explain why?" & Chr(13) & _
                     "- Star" & Chr(13) & _
                     "- Tree" & Chr(13) & _
                     "- Partial interconnection," & Chr(13) & _
                     "- Complete interconnection." & Chr(13) & _
                   "b. What are the disadvantages of circuit switching? How can one get rid of those problems? Suggest alternative switching techniques and compare with circuit switching."
StrQuestions(5) = "a. Describe the OSI network  architecture proposed by the ISO clearly indicating the function performed  by each layer." & Chr(13) & _
                  "b. Explain the following terms with respect to a computer network:" & Chr(13) & _
                  "- Protocol" & Chr(13) & _
                  "- Peer process" & Chr(13) & _
                  "- Entity" & Chr(13) & _
                  "- Interface" & Chr(13) & _
                  "- SDU and PDU"
StrQuestions(6) = "Consider a subnet having three IMPS A, B and C interconnected using three 20 kbps unidirectional links AB, BC and CA. Assume that packet arrival rate follows Poisson distribution with mean 1000 bits. Mean traffic from A to B, A to C , B to C, B to A., C to A and C to B are 6,5,9,3,5, and 7 packets/sec respectively." & Chr(13) & _
                  "a)  What is the utilization of link CA?" & Chr(13) & _
                  "b)  What is the mean number of packets/ sec that wait at B to be transmitted to C?" & Chr(13) & _
                  "c)  What is the mean time a packet arriving at A waits before its transmission to B begins?" & Chr(13) & _
                  "d)  What is the mean packet delay for each of the six different types of traffic (eg. Traffic from B to A)?"
StrQuestions(7) = "A stop and wait protocol uses a 100 kbps satellite link which employs a round trip propagation delay of 250 ms approximately. Find out the percentage of time the sender is blocked to wait for acknowledgement, if the frame size is -" & Chr(13) & _
                   "a)  1000 bits" & Chr(13) & _
                   "b)  2000 bits, and" & Chr(13) & _
                   "c)  10000 bits"""
StrQuestions(8) = "a)  Compare bit oriented and character oriented data link protocols in terms of advantages and disadvantages." & Chr(13) & _
                  "b)  What is the purpose of bit stuffing? What is the corresponding function in a character oriented protocol?"
StrQuestions(9) = "a)  What is meant by a pure ALOHA and a slotted ALOHA channel?" & Chr(13) & _
                  "b)  Why is the channel throughput double in slotted ALOHA compared to pure ALOHA?"
StrQuestions(10) = "Consider a slotted ALOHA system with a mean offered load of G pkts/pkt. Time" & Chr(13) & _
                   "a)  What is the probability that a transmission requires exactly in L attempts?" & Chr(13) & _
                   "b)  What is the expected number of transmissions per packet?"
End Sub

Private Sub questions_paper2()
StrQuestions(1) = "a)  Describe the features that distinguish a LAN from a WAN." & Chr(13) & _
                  "b)  Why is the efficiency off LAN protocols relatively higher than WAN protocols?" & Chr(13) & _
                  "c)  What modifications have been made in the OSI seven layered architecture to accommodate LANS?"""
StrQuestions(2) = "a)  Describe the protocol used in the Ethernet." & Chr(13) & _
                  "b)  Explain how collision occurs in the Ethernet." & Chr(13) & _
                  "c)  How is the Ethernet protocol different from the ALOHA protocol?"
StrQuestions(3) = "a)  Show the waveform of the electrical signal in the Ethernet as well as in the token ring corresponding to the bit pattern 0110010011" & Chr(13) & _
                  "b)  What are the functions of the AM in the token ring LAN?"
StrQuestions(4) = "a)  Explain the priority management scheme in a token ring LAN." & Chr(13) & _
                  "b)  Explain how a token ring LAN can still function even if a station or a link fails."
StrQuestions(5) = "In a token ring LAN, there are 256 stations and the distance between two neighboring stations is 10 meters. The data rate is 1 Mbps and the velocity of signal propagation is 2.5 × 10 (raise to power 5) Km/sec. At a certain time, 116 stations are powered off 50 stations are active and each active station is holding the token for 10 ms. Determine the current scan time and channel efficient."
StrQuestions(6) = "a)  Ethernet does not have any priority scheme. Mention two application areas where this can create problem." & Chr(13) & _
                  "b)  In a 2 Km long 10 M bps Ethernet, the frame length is 256 bits. Assuming a signal propagation velocity of 2.5 × 10 (raise to power 5) km/sec, find out the efficient of the channel."
StrQuestions(7) = "a)  Discuss connection less and connection oriented services offered by the network layer." & Chr(13) & _
                  "b)  In X.25 why is the virtual circuit number used buy a sending station different from the virtual circuit number used by a receiving station?"
StrQuestions(8) = "a)  Explain the flow control mechanism in X.25 protocol." & Chr(13) & _
                  "b)  In X.25 why is the virtual circuit number used by a sending station different from the virtual circuit number used by a receiving station?"
StrQuestions(9) = "    How long does it take transmit a 10 inch by 12 inch image by facsimile (Fax) over an ISDN B channel? The facsimile digitizes the image into 200 pixels per inch and assigns time of facsimile over a 64 kbps ISDN line."
StrQuestions(10) = "Consider a document with 300 × 400 mm of text with 0.25 char /mm and 0.12 lines of text per mm. Assume digitization of 400 × 400 pixels per mm, find out the transmission time of facsimile over a 64 kbps ISDN line."
End Sub

Private Sub questions_paper3()
StrQuestions(1) = "Considering  RAM access time is 100 msec, determine the number of lines a PBX can handle using time division switching?"
StrQuestions(2) = "a)  Draw and explain the FDDI class A and Class B node connections." & Chr(13) & _
                  "b)  Discuss the FSSI token and frame formats."
StrQuestions(3) = "a)  Describe the difference between synchronous and asynchronous transmission." & Chr(13) & _
                  "b)  Compare the following transmission media with examples:" & Chr(13) & _
                  "i)  Twisted pair and coaxial cable" & Chr(13) & _
                  "ii) Coaxial cable and optical fibre" & Chr(13) & _
                  "iii)    Terrestrial link and satellite"
StrQuestions(4) = "A fibre optic system requires 2 µ W of received power. The cable is 8 km long and exhibits 305 dB/km loss. There is a 1.25 dB loss at both the source and the receiver. Calculate the required level of optical power at the optical source."
StrQuestions(5) = "Answer the following in the context off a token ring LAN" & Chr(13) & _
                  "a.  What s latency buffer and why is it 30 bits long?" & Chr(13) & _
                  "b.  Why should each station know the address of its upstream neighbor. How does it know?" & Chr(13) & _
                  "c.  What happens if a token is lost or a duplicate token appears on the ring?" & Chr(13) & _
                  "d.  What is beaconing?"
StrQuestions(6) = "A certain token ring LAN has 20 stations each located at a distance of 10 meters from its left neighbors. Assume a 4 Mbps data rate, signal propagation velocity of 2.5 × 10 (raise to power 5) km/s and 20 % ring utilization, find -" & Chr(13) & _
                  "A.The wake Time" & Chr(13) & _
                  "b.  The scan time"
StrQuestions(7) = "A fiber optic token ring used as a MAN is 300 km long and runs at 200 Mbps. After sending a frame, as tation breaks the ring and drains the frame from it before regenerating the token. The signal propagation speed within the fibre is 3 × 10(raise to power 5) km/sec and the maximum frame size is 1024 bytes. Ignoring all kinds of overheads, calculates the maximum efficiency of the ring."
StrQuestions(8) = "Assume that a signal changes once every millisecond , find the baud rate and the bit rate if the system transmits" & Chr(13) & _
                  "a.Two" & Chr(13) & _
                  "b.  Four and" & Chr(13) & _
                  "c.Seven Level" & Chr(13) & _
                  "d.  Signals."
StrQuestions(9) = "Find the number of the following devices that could be accommodated by a T-1 type TDM line if 1% of the lne capacity isreserved for synch ronizatin purposes." & Chr(13) & _
                    "a.  1200 - bps computer terminal" & Chr(13) & _
                    "b.  9600 - bps computer ports" & Chr(13) & _
                    "c.  64 k bps PCM voice frequency lines"
StrQuestions(10) = "Consider a simple telephone network consisting of two end offices and one toll center with a 1- MHz full duplex trunk between each end office and the roll center. The average telephone is used to make four calls per 8 hour work day, with a mean call duration of size minutes. 10% of the calls are long distance. What is the maximum number of telephones and end office can support."
End Sub

Private Sub questions_paper4()
StrQuestions(1) = "a.  Compare the addressing schemes in HDLC, LLC and LAP - D." & Chr(13) & _
                  "b.  Is an IDN necessary for ISDN? Explain."
StrQuestions(2) = "a.  List the major disadvantages with the layered approach to protocols." & Chr(13) & _
                 "b.  Discuss the need or lack of need for a network layer(OSI layer 3) in a broadcast network."
StrQuestions(3) = "Assume a local network is to attached to a long haul X.25 network via a single station which appears to the X.25 network as a single DTE. What gateway logic in the station would allow other local network stations to have access to the X.25 network?"
StrQuestions(4) = " a) What is the advantage in connecting the outer conductor of a coaxial cable to ground?" & Chr(13) & _
                  "b)  Explain clearly the step index and graded index fibers. Also sketch the light path n each case." & Chr(13) & _
                  "c)  Illustrate the operation of a three stage switch in digital telephony."
StrQuestions(5) = "a)  What is channel capacity? How is it related to bandwidth?" & Chr(13) & _
                    "b)  In a certain communication channel, the signal power is 100 W and the noise power is 10W. In order to send information at the rate of 10 kb/s, what is the required bandwidth?" & Chr(13) & _
                    "c)  What    is channel? Explain physically and as a concept."
StrQuestions(6) = "a)  What is multiplexing? What are the different types of multiplexing possible for analog signals? Explains." & Chr(13) & _
                  "b)    A band of frequencies from 100 to 160 KHz is being allocated for channels. Each channel is 5 KHz wide, with a 1- KHz guard band. Sketch the channel assignments from 100 to 120 KHz."
StrQuestions(7) = "a.  Explain why a digital communication scheme can be made more resistant to channel noise than  an analog scheme." & Chr(13) & _
                  "b.  How can an analog signal be converted to a digital signal? Give a block diagram of the process and indicate the block(s) where error is introduced due to digitization?"
StrQuestions(8) = "a)  What is Nyquist frequency? What is the effect when an analog signal is sampled at less than the Nyquist frequency?" & Chr(13) & _
                  "b)  A system is designed to sample analog signals, concert them to digital form with a 4-bit converter, and transmit them. What bit rate is required if the analog signal consists of frequencies between 400 Hz to 3400 Hz?"
StrQuestions(9) = "a.  What is a Manchester code? What are the advantages and disadvantages of Manchester coding over unipolar NRZ coding?" & Chr(13) & _
                  "b.  Draw a circuit that will convert an NRZ coded signal to a Manchester coded signal."
StrQuestions(10) = ""
End Sub

Private Sub questions_paper5()
StrQuestions(1) = "a.  What are the various methods by which a common channel may be shared by several digital sources. Describe any one of them." & Chr(13) & _
                  "b.  A system can support a data rate of 100 kbits/. How many users can it multiplex if each user is a 3 KHZ bandwidth signal, sampled at the Nyquist rate, and 7-bit digitization?"
StrQuestions(2) = "a.  In time division multiplexing, describe the concept of framing and its importance." & Chr(13) & _
    "b.  Why are framing bits kept at a minimum? What is the risk of doing that?" & Chr(13) & _
    "c.   TDM system is designed to have framing rake up no more than 5 percent of the frame time. The system supports 10 users, each of which has been digitized to 12 bits. How many frame bits can be used?"
StrQuestions(3) = "In a PCM system consisting a PCM modulator, a digital transmission channel, and a PCM demodulator, the reconstructed analog signal at the output of the demodulator ocontains some noise, een in the absence of transmission errors." & Chr(13) & _
                  "a.  What is the noise called? Explain the reason for its presence." & Chr(13) & _
                  "b.  How can this noise be minimized? At what cost?" & Chr(13) & _
                  "c.  A 2 - bit PCM modulator is used with a 0-1  signal. What is the binary digital value that will occur for the following inputs: 0.4 V, 0.78 V."
StrQuestions(4) = "a.  Draw a block schematic of a digital communication systm and briefly describe the function if each block." & Chr(13) & _
                  "b.  For a binary pattern of 1011 0100 sketch the NRZ, RZ, and Manchester coded patterns."
StrQuestions(5) = "In the following base band communication system the encoder converts the input digital data(binary) to unipolar NRZ format. The channel is modeled as a first order lopass filter with a bandwidth B which is greater than the data rate R of the digital signal. The receiver uses a threshold detector with an appropriate timing recovery circuit." & Chr(13) & _
                    "a.  For an input data stream of 0110111010, sketch the signals at points A,B and C. Assume that the channel is noiseless." & Chr(13) & _
                    "b.  What consideration should be taken by the timing recovery circuit at the receiver for the case where channel noise is present? For the noisy channel care, what is the optimum threshold for the threshold detector circuit? Explain your answer."
StrQuestions(6) = "a.  What is meant by an eye-diagram? What information can be derived from an eye- diagram?" & Chr(13) & _
                  "b.  Draw a block schematic of a base band communication system and estimate the minimum required bandwidth of the channel for a transmitted data rate of R k bits/ s in RZ format."
StrQuestions(7) = "a.  What is binary phase shift keying(BPSJK)? Draw a block diagram of a BPSK modulator for a binary input available in bipolar NRZ format. Sketch the modulated output for  an input bit stream 1011, assuming a carrier frequency of 4 KHz and a data rate of 1k bits sec." & Chr(13) & _
                  "b.  How is BPSK demodulated? Explain the concept of synchronization error in a BPSK receiver and its effect on the recovered signal."
StrQuestions(8) = "a)  What is differential phase shift keying? How is it achieved? Take a suitable stream of binary data of your choice and indicate the phases of a binary DPSK waveform, clearly illustrating the DPSK modulation scheme." & Chr(13) & _
                  "b)  Explain the advantage of a differential phase shift keying modulation over ordinary PSK. What is the price paid for achieving this advantage?"
StrQuestions(9) = "a)  A binary ASK modulator applies on off keying of a carrier of frequency f0. The spectrum fo the modulated signal has a bandwidth equal to 2R. The same binary signal is applied to an FSK modulator with mark and space frequencies f1 and f0 respectively, f1 > f0. Estimate the bandwidth of the FSK signal by sketching the spectrum of the ASK and FSK signals. What should be the minimum separation between f1 and f0?" & Chr(13) & _
                  "b)  Draw the block schematics of a coherent and a non coherent receiver for binary FSK. Which one gives a better performance?"
StrQuestions(10) = ""
End Sub

Private Sub questions_paper6()
StrQuestions(1) = "a)  What is meant by QPSK? Draw the block diagram of QPSK modulate and derive the trigonometric formulae representing the modulated signal corresponding to different input bits, clearly showing the frequency and phase terms." & Chr(13) & _
                  "b)  Explain what is meant by carrier recovery. Give a block diagram of any carrier recovery scheme known to you. Explain the schema"
StrQuestions(2) = "a)  Explain the mechanism of error detection by sending a parity check bit for every transmitted word of length n. How can the error detection capability of his scheme be improved? AT what cost?" & Chr(13) & _
                  "b)  A forward error correcting scheme can correct up to two errors, the codeword length being 23 bits. Calculate probability of word error, given that the probability of bit error in transmission is 1 in 100."
StrQuestions(3) = "a)  What is an error detection scheme? Give examples." & Chr(13) & _
                    "b)  What is a forward error correcting scheme? Give an example of a forward error correcting scheme and explain how error correction is achieved?" & Chr(13) & _
                    "c)  What is meant by coding gain? What is its implication in terms of the data rate and transmission bandwidth?"
StrQuestions(4) = "Describe the concept of repetition encoding for error detection and error correction. Using an example of triple repetition encoding in a channel having a probability of bit error = 0.01." & Chr(13) & _
                    "a)  Find the output bit error rate at the decoder when the code is used for error detection only." & Chr(13) & _
                    "b)  Find the output bit error rate when the code is used for error correction with majority logic decoding."
StrQuestions(5) = "a)  Explain the concept of block forward error correcting codes. In your answer bring out the definition of code rate and error correcting ability." & Chr(13) & _
                  "b)  A digital signal utilizes 8 - bit codeword. Calculate the probability of areceive codeword containing at least 2 errors when the bit error probability in transmission is 0.01."
StrQuestions(6) = "a)  With the help of neat ray diagrams, explain the light ran propagation mechanisms in single mode, step index multimode, and graded index fibers. What are the approximate dimensions of core and cladding in each case?" & Chr(13) & _
                  "b)  Describe the nature of attenuation of an optical signal through a glass fiber and how it varies with the wavelength of the optical signal. State the principle causes of attenuation in a fiber."
StrQuestions(7) = "a)  Describe the structure of a n optical fiber and explain the mechanism of light propagation along the fiber." & Chr(13) & _
                  "b)  What are the relative merits and demerits of a single mode fiber in comparison to a multi mode fiber. Describe the structure and composition differences between the two."
StrQuestions(8) = "a)  What limits the maximum transmission rate of optical pulses along an optical fiber? Describe some of the methods used in manufacturing the fiber to overcome these limitations." & Chr(13) & _
                  "b)  What is meant by inter modal dispersion? What are the parameters which determine inter modal dispersion?"
StrQuestions(9) = "a)  Illustrate with the help of a block schematic the different components ofa typical long distance fiber optic link. Mention the various components of signal loss in the link between the optical source and the photo detector," & Chr(13) & _
                  "b)  What advantages does a semiconductor laser diode(LD) have over light emitting diodes(LED) for fiber transmission? When is an LED preferred over an LD in fiber optic communication?"
StrQuestions(10) = ""

End Sub

Private Sub questions_paper7()
StrQuestions(1) = "a)  Describe the various functional blocks of digital FAX machine?" & Chr(13) & _
                  "b)  An electronic scanner in a FAX transmitter produces 1.11 million pixels per document page. Each pixel is quantized to two levels and the digital image is compressed at an average rate of 35 to 1 using run length coding and Hoffman code. If he maximum transmission bandwidth is 1 KHz, find the document transmission time."
StrQuestions(2) = "a)  Explain the term full duplex as applied to telephony." & Chr(13) & _
                  "b)  What is the type of multiplexing  used in telephone trunk circuits to transmits a large number of voice channels by the same transmitter? How is it achieved?"
StrQuestions(3) = "a)  What does one need in order to transmit computer data over a telephone line? What are its functional blocks?" & Chr(13) & _
                  "b)  What is the role of a PABX? What can a PABX do for the phone user that the regular phone company exchange may not be able to do?"
StrQuestions(4) = "a)  Describe the nature of signals produced on the subscribers' loop by a pulse dialer." & Chr(13) & _
                    "b)  Describe the nature of the signals produced by a Tone dialer." & Chr(13) & _
                    "c)  A DS1 digital signal contains 24 eight bit PCM channels per frame time division multiplexed on a cable. One synchronizing bit per frame is supplied, and the sampling rate is 8 KHZ. Determine the bit rate on the cable."
StrQuestions(5) = "a)  What is the function of a modem? Why is it necessary?" & Chr(13) & _
                  "b)  A modem uses 1000 Kz for 0, and 2000 Hz for 1. Sketch the waveform seen on the signal line for the following pattern of data bits: 110010, assuming that the data is being sent at 1000 baud. Show the timing of a singles bit." & Chr(13) & _
                  "c)  What is the goal of ISDN?"
StrQuestions(6) = "a)  Compare FEC to ARZ. What are the advantages and drawbacks of each?" & Chr(13) & _
                  "b)  What is aliasing ? What is done to limit aliasing?"
StrQuestions(7) = "a)  Give examples of transmission media in which security is poor, fair, or good." & Chr(13) & _
                "b)  What is Manchester coding? Why is it used? Compare the band width it uses to that of NRA." & Chr(13) & _
                "c)  A communications system uses an 8-bit  converter at exactly the Nyquist rate for signals with 30 KHz bandwidth. Find the bit rate at the output."
StrQuestions(8) = "a)  What is meant by simplex, half duplex, and full duplex communications systems. Give representative examples of each." & Chr(13) & _
                    "b)  What is bandwidth compression? Taking the example of the data transmitted by a fax machine, explain how bandwidth compression may be done." & Chr(13) & _
                    "c)  Is there any component in a data transmission system that increases the message bandwidth? If yes, state the function of the component."
StrQuestions(9) = ""
StrQuestions(10) = ""

End Sub

Private Sub multiplechoice_Paper1()
strmultiple(1) = "The first successful packet broadcast network is -"
strmuloption(1) = "TYMNET"
strmuloption(2) = "ALOHA"
strmuloption(3) = "ARPANET"
strmuloption(4) = "BITNET"

strmultiple(2) = "RABMN network belongs to -"
strmuloption(5) = "    American Defense"
strmuloption(6) = "European bank"
strmuloption(7) = "DOE, India"
strmuloption(8) = "DOE, India"

strmultiple(3) = "The medium size network is termed as -"
strmuloption(9) = "WAN"
strmuloption(10) = "LAN"
strmuloption(11) = "MAN"
strmuloption(12) = "VAN"

strmultiple(4) = "Which of the following standards is followed in India?"
strmuloption(13) = "IFIP"
strmuloption(14) = "CCITT"
strmuloption(15) = "ISO"
strmuloption(16) = "ANSI"

strmultiple(5) = "Which of the following is not a goal of a computer network?"
strmuloption(17) = "Resource sharing "
strmuloption(18) = "Low reliability"
strmuloption(19) = "Load sharing"
strmuloption(20) = "To avoid physical movement"

strmultiple(6) = "Computers in a computer network are connected via -"
strmuloption(21) = "Telephone line only"
strmuloption(22) = "Satellite channel only"
strmuloption(23) = "Either telephone line or satellite channel or both"
strmuloption(24) = "Coaxial cable only"

strmultiple(7) = "E-mail stands for -"
strmuloption(25) = "Electronic mail"
strmuloption(26) = "Excess mail"
strmuloption(27) = "Educational mail"
strmuloption(28) = "Express mail"

strmultiple(8) = "Modem refers to -"
strmuloption(29) = "modulation"
strmuloption(30) = "modulator"
strmuloption(31) = "modulator and demodulator"
strmuloption(32) = "modern modes of communication"

strmultiple(9) = "The latest modulation technique used by data modems is -"
strmuloption(33) = "QPSK"
strmuloption(34) = "FSK"
strmuloption(35) = "DPSK"
strmuloption(36) = "ASK"

strmultiple(10) = "DTE is popularly known as -"
strmuloption(37) = "Data Terminal Equipment"
strmuloption(38) = "Digital Type Equipment"
strmuloption(39) = "Discrete Time Equipment"
strmuloption(40) = "Dual Time Equipment"

strmulanswer(1) = "ALOHA"
strmulanswer(2) = "DOE, India"
strmulanswer(3) = "MAN"
strmulanswer(4) = "CCITT"
strmulanswer(5) = "Low reliability"
strmulanswer(6) = "Either telephone line or satellite channel or both"
strmulanswer(7) = "Electronic mail"
strmulanswer(8) = "modulator and demodulator"
strmulanswer(9) = "DPSK"
strmulanswer(10) = "Data Terminal Equipment"
End Sub

Private Sub multiplechoice_Paper2()
strmultiple(1) = "Signals through a compute network are -"
strmuloption(1) = "Continuous"
strmuloption(2) = "Discrete"
strmuloption(3) = "Either a or b"
strmuloption(4) = "Neither a nor b"

strmultiple(2) = "Messages are transferred in a computer network through -"
strmuloption(5) = "Circuit switching "
strmuloption(6) = "Message switching"
strmuloption(7) = "Packet switching"
strmuloption(8) = "Multiplexing"

strmultiple(3) = "Which one of the following is not a guided transmission line?"
strmuloption(9) = "Pair of wires"
strmuloption(10) = "Waveguides"
strmuloption(11) = "Laser beam"
strmuloption(12) = "Fiber optic channel"

strmultiple(4) = "Signals that involve human communication are generally –"
strmuloption(13) = "Digital"
strmuloption(14) = "Analog  "
strmuloption(15) = "Either digital or analog"
strmuloption(16) = "None of these"

strmultiple(5) = "A digital channel implies that the channel –"
strmuloption(17) = "is digitized"
strmuloption(18) = "is carrying digital data"
strmuloption(19) = "adopts digital modulation techniques"
strmuloption(20) = "is without carrier"

strmultiple(6) = "Frequency range of the human voice is usually –"
strmuloption(21) = "100 – 4000 Hz"
strmuloption(22) = "200 – 600 Hz"
strmuloption(23) = "199 – 200 Hz"
strmuloption(24) = "400 – 800 Hz"

strmultiple(7) = "In twisted pair cables the mode of communication is –"
strmuloption(25) = "FCM"
strmuloption(26) = "CDMA    "
strmuloption(27) = "TDM"
strmuloption(28) = "TDMA"

strmultiple(8) = "DCE is popularly known as –"
strmuloption(29) = "Data circuit terminating equipment"
strmuloption(30) = "Discrete circuit equipment"
strmuloption(31) = "Dual circuit equipment"
strmuloption(32) = "Discrete component equipment"

strmultiple(9) = "Which one of the following is a balanced interface?"
strmuloption(33) = "RS-232C"
strmuloption(34) = "RS-423"
strmuloption(35) = "RS-449"
strmuloption(36) = "RS-422"

strmultiple(10) = "Digitization refers to –"
strmuloption(37) = "Sampling"
strmuloption(38) = "Quantization"
strmuloption(39) = "Sampling and quantization"
strmuloption(40) = "Either sampling or quantization"

strmulanswer(1) = "Discrete"
strmulanswer(2) = "Packet switching"
strmulanswer(3) = "Laser beam"
strmulanswer(4) = "Analog  "
strmulanswer(5) = "is carrying digital data"
strmulanswer(6) = "100 – 4000 Hz"
strmulanswer(7) = "TDM"
strmulanswer(8) = "Data circuit terminating equipment"
strmulanswer(9) = "RS-422"
strmulanswer(10) = "Sampling and quantization"
End Sub

Private Sub multiplechoice_Paper3()
strmultiple(1) = "If R is the data rate in a binary wave form then the bit period Y is equal to –"
strmuloption(1) = "1/R"
strmuloption(2) = "R"
strmuloption(3) = "1/R2"
strmuloption(4) = "Log2 R"

strmultiple(2) = "If there are M possible signal levels, each signal element is coded as K = [log2 M] binary digits then the bit rate are –"
strmuloption(5) = "M times the baud rate"
strmuloption(6) = "K times the baud rate"
strmuloption(7) = "Equal to the baud rate"
strmuloption(8) = "log2M times the baud rate"

strmultiple(3) = "Telephone networks usually use –"
strmuloption(9) = "Simplex transmission"
strmuloption(10) = "Half-duple transmission"
strmuloption(11) = "Full duplex transmission"
strmuloption(12) = "None of these"

strmultiple(4) = "Data networks usually use –"
strmuloption(13) = "Simplex transmission"
strmuloption(14) = "Half-duple transmission"
strmuloption(15) = "Full duplex transmission"
strmuloption(16) = "    None of these"

strmultiple(5) = "For long distance data transmission the preferable mode of communication is –"
strmuloption(17) = "Serial transmission"
strmuloption(18) = "Parallel transmission"
strmuloption(19) = "Either a or b"
strmuloption(20) = "Both a and b"

strmultiple(6) = "The transmission of a digital signal at its original frequency without modulation is called –"
strmuloption(21) = "Broad band signaling"
strmuloption(22) = "Base band signaling"
strmuloption(23) = "Digital signaling"
strmuloption(24) = "Pass band signaling"

strmultiple(7) = "During transmission the distortion of the signal depends on –"
strmuloption(25) = "The length of transmission"
strmuloption(26) = "The frequency of the signal"
strmuloption(27) = "Both of these"
strmuloption(28) = "Both of these"

strmultiple(8) = "Signal rate of a communication channel depends on –"
strmuloption(29) = "Frequency and time responses of the channel"
strmuloption(30) = "Frequency responses of the channel only"
strmuloption(31) = "Time of responses of the channel only"
strmuloption(32) = "None of these"

strmultiple(9) = "The maximum data rate of a channel having a BW = BHz and number of discrete signal values = V, is given by –"
strmuloption(33) = "[B log2 V] bps"
strmuloption(34) = "[log2 V] bps"
strmuloption(35) = "[B/2 log2 V] bps"
strmuloption(36) = "[2B log2 VV] bps"

strmultiple(10) = "For a channel of BW = BHz and signal to noise ratio = S/N dB the channel capacity is limited by –"
strmuloption(37) = "[B log2 (1+ S/N)] bps"
strmuloption(38) = "[2B log2 (1+S/N)] bps"
strmuloption(39) = "[B/2 log2 (1 + S/N)] bps"
strmuloption(40) = "[Log2 (1+S/N]bps"


strmulanswer(1) = "1/R"
strmulanswer(2) = "K times the baud rate"
strmulanswer(3) = "Half-duple transmission"
strmulanswer(4) = "Full duplex transmission"
strmulanswer(5) = "Serial transmission"
strmulanswer(6) = "Base band signaling"
strmulanswer(7) = "Both of these"
strmulanswer(8) = "Frequency and time responses of the channel"
strmulanswer(9) = "[2B log2 VV] bps"
strmulanswer(10) = "[B log2 (1+ S/N)] bps"
End Sub

Private Sub multiplechoice_Paper4()
strmultiple(1) = "In a coaxial cable the signal power decreases –"
strmuloption(1) = "lineally with the propagation distance"
strmuloption(2) = "inversely with the propagation distance"
strmuloption(3) = "exponentially with the propagation distance"
strmuloption(4) = "does not depend on the propagation distance"

strmultiple(2) = "Twisted wire pair can accommodate a maximum data rate of –"
strmuloption(5) = "2400 bps"
strmuloption(6) = "4800 bps"
strmuloption(7) = "9600 bps"
strmuloption(8) = "19200 bps"

strmultiple(3) = "Coaxial cables are widely used in –"
strmuloption(9) = "Telephone networks"
strmuloption(10) = "cable TV networks"
strmuloption(11) = "Both in telephone and cable TV networks"
strmuloption(12) = "Computer networks"

strmultiple(4) = "A very popular base band LAN Ethernet is essentially a –"
strmuloption(13) = "Coaxial cable"
strmuloption(14) = "Twisted wire pair"
strmuloption(15) = "Optical fiber"
strmuloption(16) = "None of these"

strmultiple(5) = "In satellite communication up link frequency differs from the down link frequency in order to –"
strmuloption(17) = "avoid interference"
strmuloption(18) = "permit full duplex operation"
strmuloption(19) = "follow the convention"
strmuloption(20) = "none of these"

strmultiple(6) = "Angle modulation is referred as –"
strmuloption(21) = "phase modulation only"
strmuloption(22) = "frequency modulation only"
strmuloption(23) = "Amplitude and frequency modulation"
strmuloption(24) = "Frequency and phase modulation"

strmultiple(7) = "ASK is rarely used in modems because, it –"
strmuloption(25) = "shifts only between ‘on’ and ‘off’ state"
strmuloption(26) = "takes care of the amplitude only"
strmuloption(27) = "is highly susceptible to noise"
strmuloption(28) = "none of these"

strmultiple(8) = "QAM (Quadrature Amplitude Modulation) is –"
strmuloption(29) = "PSK combined with AM"
strmuloption(30) = "FSK combined with AM"
strmuloption(31) = "ASK combined with AM"
strmuloption(32) = "QPSK combined with AM"

strmultiple(9) = "Channel coding is used to –"
strmuloption(33) = "secure the channel"
strmuloption(34) = "minimize the interference in the channel"
strmuloption(35) = "protect the information against channel noise"
strmuloption(36) = "protect against unnecessary tapping of signal from the channel"

strmultiple(10) = "A band limited signal can be completely characterized by sampling its amplitude the rate (R) of which , is –"
strmuloption(37) = "R > f max"
strmuloption(38) = "R >= f max"
strmuloption(39) = "R > ½ f max"
strmuloption(40) = "R > 2 f max"

strmulanswer(1) = "exponentially with the propagation distance"
strmulanswer(2) = "9600 bps"
strmulanswer(3) = "Both in telephone and cable TV networks"
strmulanswer(4) = "Coaxial cable"
strmulanswer(5) = "permit full duplex operation"
strmulanswer(6) = "Frequency and phase modulation"
strmulanswer(7) = "is highly susceptible to noise"
strmulanswer(8) = "PSK combined with AM"
strmulanswer(9) = "protect the information against channel noise"
strmulanswer(10) = "R > 2 f max"
End Sub

Private Sub multiplechoice_Paper5()
strmultiple(1) = "The sampling interval of telephone voice signal is –"
strmuloption(1) = "125 ?sec "
strmuloption(2) = "25 ?sec"
strmuloption(3) = "50 ?sec "
strmuloption(4) = "100 ?sec"

strmultiple(2) = "A transmitted bit, in base band transmission, can be a/an –"
strmuloption(5) = "current pulse"
strmuloption(6) = "voltage pulse"
strmuloption(7) = "optical power pulse"
strmuloption(8) = "All of these    "

strmultiple(3) = "IEEE 802.5 employs –"
strmuloption(9) = "Manchester coding"
strmuloption(10) = "Manchester – II coding"
strmuloption(11) = "50% differential Manchester coding"
strmuloption(12) = "Differential Manchester coding"

strmultiple(4) = "Radio broadcasting is a common example of –"
strmuloption(13) = "FDM"
strmuloption(14) = "TDM"
strmuloption(15) = "Both of these"
strmuloption(16) = "None of these"

strmultiple(5) = "In network terminology, ‘SAP’ refers to –"
strmuloption(17) = "secured application"
strmuloption(18) = "service access points"
strmuloption(19) = "simple acce3ss protocol"
strmuloption(20) = "service access protocol"

strmultiple(6) = "Which one of the following is the correct structure for the ARPANET?"
strmuloption(21) = "7-layer structure"
strmuloption(22) = "6-layer structure"
strmuloption(23) = "4-layer structure"
strmuloption(24) = "3-layer structure"

strmultiple(7) = "Double bit errors can be detected through –"
strmuloption(25) = "VRC"
strmuloption(26) = "CRC"
strmuloption(27) = "LRC"
strmuloption(28) = "None of these"

strmultiple(8) = "Most commonly used protocol in DLC(Data Link Control) procedures is –"
strmuloption(29) = "sliding window protocol(in general)"
strmuloption(30) = "stop and wait sliding window protocol"
strmuloption(31) = "sliding window protocol with go-back-N"
strmuloption(32) = "sliding window with selective repeat"

strmultiple(9) = "Which one of the following is a not a Bit oriented protocol?"
strmuloption(33) = "BISYNC"
strmuloption(34) = "DDCMP"
strmuloption(35) = "HDLC"
strmuloption(36) = "SDLC"

strmultiple(10) = "‘Bit Surfing’ is a common technique available in –"
strmuloption(37) = "Character oriented protocol"
strmuloption(38) = "Bit oriented protocol"
strmuloption(39) = "Sliding window protocol with selective repeat"
strmuloption(40) = "Sliding window with go-back-N"

strmulanswer(1) = "125 ?sec "
strmulanswer(2) = "All of these    "
strmulanswer(3) = "50% differential Manchester coding"
strmulanswer(4) = "FDM"
strmulanswer(5) = "service access points"
strmulanswer(6) = "4-layer structure"
strmulanswer(7) = "CRC"
strmulanswer(8) = "sliding window protocol with go-back-N"
strmulanswer(9) = "BISYNC"
strmulanswer(10) = "Bit oriented protocol"
End Sub

Private Sub multiplechoice_Paper6()
strmultiple(1) = "Packet generation in ALOHA networks follows –"
strmuloption(1) = "Poisson distribution"
strmuloption(2) = "Gaussian distribution"
strmuloption(3) = "Normal distribution"
strmuloption(4) = "None of the above"

strmultiple(2) = "Token ring uses"
strmuloption(5) = "Manchester encoding"
strmuloption(6) = "Analog modulation encoding"
strmuloption(7) = "Differential Manchester encoding"
strmuloption(8) = "None of the above"

strmultiple(3) = "Which of the following types of signal requires the highest bandwidth for transmission?"
strmuloption(9) = "Speech"
strmuloption(10) = "Video"
strmuloption(11) = "music"
strmuloption(12) = "facsimile"

strmultiple(4) = "An example of base band data communication may be found in –"
strmuloption(13) = "Broadcast radio"
strmuloption(14) = "TV"
strmuloption(15) = "Telephone"
strmuloption(16) = "Satellite links"

strmultiple(5) = "The advantage of the optical fiber is –"
strmuloption(17) = "Low attenuation "
strmuloption(18) = "High bandwidth"
strmuloption(19) = "Low interference"
strmuloption(20) = "All of the above"

strmultiple(6) = "In a telephone channel, voice signals are band limited to –"
strmuloption(21) = "1 KHz   "
strmuloption(22) = "3.4 KHz"
strmuloption(23) = "8 KHz"
strmuloption(24) = "20 KHz"

strmultiple(7) = "Attenuation in an optical fiber can be as small as –"
strmuloption(25) = "20 dB/km"
strmuloption(26) = "2 dB/km"
strmuloption(27) = "0.2 dB/km"
strmuloption(28) = "0.02 dB/km"

strmultiple(8) = "A twisted wire cable cannot be used for the transmission of –"
strmuloption(29) = "analog voice"
strmuloption(30) = "digitized voice"
strmuloption(31) = "analog video"
strmuloption(32) = "None of the above"

strmultiple(9) = "The local loop of a telephone channel consists of the following transmission medium –"
strmuloption(33) = "twin-wire"
strmuloption(34) = "microwave"
strmuloption(35) = "satellite"
strmuloption(36) = "coaxial cable"

strmultiple(10) = "The following transmission link supports the lowest transmission bandwidth –"
strmuloption(37) = "wireless links"
strmuloption(38) = "coaxial cable"
strmuloption(39) = "twisted wire pair"
strmuloption(40) = "optical fibers"

strmulanswer(1) = "Poisson distribution"
strmulanswer(2) = "Differential Manchester encoding"
strmulanswer(3) = "Video"
strmulanswer(4) = "Telephone"
strmulanswer(5) = "All of the above"
strmulanswer(6) = "3.4 KHz"
strmulanswer(7) = "0.2 dB/km"
strmulanswer(8) = "analog video"
strmulanswer(9) = "twin-wire"
strmulanswer(10) = "twisted wire pair"
End Sub

Private Sub multiplechoice_Paper7()
strmultiple(1) = "The following modulation requires the lowest bandwidth –"
strmuloption(1) = "ASK"
strmuloption(2) = "PSK"
strmuloption(3) = "FSK"
strmuloption(4) = "QPSK"

strmultiple(2) = "An FSK signal can be demodulated using a/an –"
strmuloption(5) = "envelope detector"
strmuloption(6) = "low pass filter"
strmuloption(7) = "frequency discriminator"
strmuloption(8) = "all of the above"

strmultiple(3) = "An FSK signal can be detected by a –"
strmuloption(9) = "FM discriminator"
strmuloption(10) = "differential detector"
strmuloption(11) = "matched filler"
strmuloption(12) = "all of the above"

strmultiple(4) = "The following modulation can be detected by a non coherent receiver –"
strmuloption(13) = "PSK"
strmuloption(14) = "QPSK"
strmuloption(15) = "FSK"
strmuloption(16) = "All of the above"

strmultiple(5) = "A base band PAM signal applied to a voltage controlled oscillator produces he following modulation- "
strmuloption(17) = "PSK"
strmuloption(18) = "FSK"
strmuloption(19) = "DPSK"
strmuloption(20) = "ASK"

strmultiple(6) = "Intersymbol interference is a result of –"
strmuloption(21) = "noise is the channel    "
strmuloption(22) = "bandwidth limitation of the channel"
strmuloption(23) = "demodulation defects"
strmuloption(24) = "carrier recovery error"

strmultiple(7) = "A PSK signal can be demodulated using a –"
strmuloption(25) = "frequency discriminator"
strmuloption(26) = "balanced modulator"
strmuloption(27) = "voltage controlled oscillator"
strmuloption(28) = "None of the above"

strmultiple(8) = "The following modulation produces a pass band signal –"
strmuloption(29) = "PSK"
strmuloption(30) = "DPSK"
strmuloption(31) = "FSK"
strmuloption(32) = "All of these"

strmultiple(9) = "When an analog signal which is band limited to 5 KHz s to be transmitted using an 8-bit PCM, the minimum data rate required is –"
strmuloption(33) = "8 kbits/sec"
strmuloption(34) = "16 kbits/sec"
strmuloption(35) = "32 kbits/sec"
strmuloption(36) = "64 kbits/sec"

strmultiple(10) = "The number of quantization levels in an 8-bit PCM encoder are –"
strmuloption(37) = "8"
strmuloption(38) = "16"
strmuloption(39) = "128"
strmuloption(40) = "256"


strmulanswer(1) = "PSK"
strmulanswer(2) = "frequency discriminator"
strmulanswer(3) = "all of the above"
strmulanswer(4) = "FSK"
strmulanswer(5) = "FSK"
strmulanswer(6) = "bandwidth limitation of the channel"
strmulanswer(7) = "balanced modulator"
strmulanswer(8) = "All of these"
strmulanswer(9) = "64 kbits/sec"
strmulanswer(10) = "256"
End Sub

Private Sub multiplechoice_Paper8()
strmultiple(1) = "An example of full duplex transmission is –"
strmuloption(1) = "telephone channel"
strmuloption(2) = "satellite channel"
strmuloption(3) = "broadcast radio"
strmuloption(4) = "None of these"

strmultiple(2) = "An example of half duplex transmission is –"
strmuloption(5) = "telephone channel"
strmuloption(6) = "satellite channel"
strmuloption(7) = "broadcast radio"
strmuloption(8) = "None of these"

strmultiple(3) = "A phase locked loop may e applied for demodulating –"
strmuloption(9) = "PSK"
strmuloption(10) = "FSK"
strmuloption(11) = "QPSK"
strmuloption(12) = "None of the above"

strmultiple(4) = "Differential encoding of PSK is used to avoid –"
strmuloption(13) = "carrier recovery in the receiver"
strmuloption(14) = "timing recovery in the receiver"
strmuloption(15) = "improve noise immunity"
strmuloption(16) = "All of the above"

strmultiple(5) = "A conventional PABX uses –"
strmuloption(17) = "circuit switching"
strmuloption(18) = "packet switching"
strmuloption(19) = "both circuit and packet switching"
strmuloption(20) = "none of these"

strmultiple(6) = "The lowest optical signal attenuation, in a fiber made of para silica, is obtained at –"
strmuloption(21) = "800 nm"
strmuloption(22) = "1100 nm"
strmuloption(23) = "1300 nm"
strmuloption(24) = "1500 nm"

strmultiple(7) = "Single mode fibers have the following advantage over multimode fibers –"
strmuloption(25) = "lower attenuation"
strmuloption(26) = "lower dispersion"
strmuloption(27) = "lower cost"
strmuloption(28) = "easier to manufacture"

strmultiple(8) = "A graded index fiber has the following advantage over a step index fiber –"
strmuloption(29) = "lower attenuation"
strmuloption(30) = "lower dispersion"
strmuloption(31) = "lower cost"
strmuloption(32) = "easier to manufacture"

strmultiple(9) = "A laser diode has the following advantage over the LED as a fiber optic source –"
strmuloption(33) = "high power"
strmuloption(34) = "coherent optical output"
strmuloption(35) = "high directivity"
strmuloption(36) = "All of these"

strmultiple(10) = "The intrinsic later in a PIN photodiode is used for –"
strmuloption(37) = "increasing bandwidth of the receiver"
strmuloption(38) = "increasing high absorption"
strmuloption(39) = "reducing noise at the receiver"
strmuloption(40) = "None of these"

strmulanswer(1) = "telephone channel"
strmulanswer(2) = "satellite channel"
strmulanswer(3) = "PSK"
strmulanswer(4) = "carrier recovery in the receiver"
strmulanswer(5) = "circuit switching"
strmulanswer(6) = "1500 nm"
strmulanswer(7) = "lower dispersion"
strmulanswer(8) = "lower dispersion"
strmulanswer(9) = "All of these"
strmulanswer(10) = "increasing high absorption"
End Sub

Private Sub multiplechoice_Paper9()
strmultiple(1) = "An avalanche photodiode produces gain by the effect of –"
strmuloption(1) = "optical amplification"
strmuloption(2) = "electrical amplification"
strmuloption(3) = "avalanche breakdown"
strmuloption(4) = "All of these"

strmultiple(2) = "The following can support highest rates of data transmission-"
strmuloption(5) = "step index multi mode fiber"
strmuloption(6) = "graded index multi mode fiber"
strmuloption(7) = "coaxial cable"
strmuloption(8) = "single mode fiber"

strmultiple(3) = "Shot noise is generated at the –"
strmuloption(9) = "photo detector"
strmuloption(10) = "pre-amplifier"
strmuloption(11) = "optical fiber"
strmuloption(12) = "None of these"

strmultiple(4) = "A double heterojunction in an LED source provides –"
strmuloption(13) = "carrier confinement"
strmuloption(14) = "optical confinement"
strmuloption(15) = "both a and b"
strmuloption(16) = "None of these"

strmultiple(5) = "The lowest intramodal dispersion in silica fibers may b found near-"
strmuloption(17) = "800 nm"
strmuloption(18) = "1300 nm"
strmuloption(19) = "1500 nm"
strmuloption(20) = "1650 nm"

strmultiple(6) = "Dispersion shifted fibers are made by the use of specific reference index profiles to change the following properties of a fiber –"
strmuloption(21) = "material dispersion"
strmuloption(22) = "waveguide dispersion"
strmuloption(23) = "intermodal dispersion"
strmuloption(24) = "attenuation"

strmultiple(7) = "The cladding layer in an optical fiber is used for –"
strmuloption(25) = "providing protection to the fiber"
strmuloption(26) = "providing better dispersion properties"
strmuloption(27) = "confining the optical signal within the core"
strmuloption(28) = "providing better attenuation properties"

strmultiple(8) = "The numerical aperture of an optical fiber gives a measure of its –"
strmuloption(29) = "thickness"
strmuloption(30) = "core refractive index"
strmuloption(31) = "cladding refractive index"
strmuloption(32) = "light gathering capability of the fiber"

strmultiple(9) = "A laser diode source is used with a non zero extinction ratio for –"
strmuloption(33) = "higher lifetime of the laser"
strmuloption(34) = "simpler driving circuit"
strmuloption(35) = "lower noise figure at the receiver"
strmuloption(36) = "higher speed of operation"

strmultiple(10) = "Carrier recovery is usually difficult in a satellite communication system because of –"
strmuloption(37) = "large distance"
strmuloption(38) = "signal fading"
strmuloption(39) = "high frequency of transmission"
strmuloption(40) = "lower power in the satellite"

strmulanswer(1) = "avalanche breakdown"
strmulanswer(2) = "single mode fiber"
strmulanswer(3) = "photo detector"
strmulanswer(4) = "both a and b"
strmulanswer(5) = "1300 nm"
strmulanswer(6) = "waveguide dispersion"
strmulanswer(7) = "confining the optical signal within the core"
strmulanswer(8) = "light gathering capability of the fiber"
strmulanswer(9) = "higher speed of operation"
strmulanswer(10) = "signal fading"
End Sub

Private Sub multiplechoice_Paper10()
strmultiple(1) = "In a telephone network an echo may be introduced at the -"
strmuloption(1) = "switch"
strmuloption(2) = "multiplexer"
strmuloption(3) = "transmission cable"
strmuloption(4) = "hybrid"

strmultiple(2) = "In digital telephone transmission the digitization of speech signals is doneby -"
strmuloption(5) = "4 bit PCM"
strmuloption(6) = "6 bit PCM"
strmuloption(7) = "8 bit PCM"
strmuloption(8) = "16 bit PCM"

strmultiple(3) = "A 30 channel PCM signal in digital telephony will have a data rate of -"
strmuloption(9) = "64 kb/sec"
strmuloption(10) = "240 kb/sec"
strmuloption(11) = "1 mb/sec"
strmuloption(12) = "2.048 mb/sec"

strmultiple(4) = "The following is not a part of a FAX transmitter -"
strmuloption(13) = "PCM encoder"
strmuloption(14) = "scanning"
strmuloption(15) = "data compaction"
strmuloption(16) = "digital modulation"

strmultiple(5) = "In time division multiplexing of n channels, each having a data rate of R kb/sec, the multiplexed signal will have a net bit rate of -"
strmuloption(17) = "nR kb/sec"
strmuloption(18) = "slightly more htna nR kb/sec"
strmuloption(19) = "2nR kb/sec"
strmuloption(20) = "slightly more than 2nR kb/sec"

strmultiple(6) = "The overhead in a time division multiplexed signal is required for -"
strmuloption(21) = "synchronization"
strmuloption(22) = "clock recovery"
strmuloption(23) = "error control"
strmuloption(24) = "none of these"

strmultiple(7) = "If insufficient channel separation is used in frequency division multiplexing systems, the effect is -"
strmuloption(25) = "intersymbol interference"
strmuloption(26) = "intermodulation"
strmuloption(27) = "echo"
strmuloption(28) = "cross talk"

strmultiple(8) = "A satellite must be equipped with a -"
strmuloption(29) = "narrowband receiver"
strmuloption(30) = "baseband receiver"
strmuloption(31) = "a tunable receiver"
strmuloption(32) = "a wideband receiver"

strmultiple(9) = "The downlink and uplink channels of a satellite are separated -"
strmuloption(33) = "in time"
strmuloption(34) = "in frequency"
strmuloption(35) = "in space"
strmuloption(36) = "not separated"

strmultiple(10) = "Satellite to ground communication takes place through -"
strmuloption(37) = "medium wave"
strmuloption(38) = "short wave"
strmuloption(39) = "microwave"
strmuloption(40) = "optical signals"

strmulanswer(1) = "hybrid"
strmulanswer(2) = "8 bit PCM"
strmulanswer(3) = "2.048 mb/sec"
strmulanswer(4) = "PCM encoder"
strmulanswer(5) = "slightly more htna nR kb/sec"
strmulanswer(6) = "synchronization"
strmulanswer(7) = "cross talk"
strmulanswer(8) = "a wideband receiver"
strmulanswer(9) = "in frequency"

strmulanswer(10) = "microwave"
End Sub




