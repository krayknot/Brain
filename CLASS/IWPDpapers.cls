VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IWPDpapers"
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
 
 
Public Function IWPD(PaperType As String, IWPDPaper As Paper, mCat As Category, Number As Integer)
'ptype is the paper type
'IWPDpaper is the number of paper as paper1,paper2 ...
'mcat is the category
'number is the number of questions, or options or answers to be taken

  StrPaperType = PaperType
' StrPaper = Paper
If PaperType = "Fill in the Blanks" Then
   If IWPDPaper = Paper1 Then Fill_in_the_Blanks_Paper1
   If IWPDPaper = Paper2 Then Fill_in_the_Blanks_Paper2
   If IWPDPaper = Paper3 Then Fill_in_the_Blanks_Paper3
   If IWPDPaper = Paper4 Then Fill_in_the_Blanks_Paper4
   If IWPDPaper = Paper5 Then Fill_in_the_Blanks_Paper5
   If IWPDPaper = Paper6 Then Fill_in_the_Blanks_Paper6
   If IWPDPaper = Paper7 Then Fill_in_the_Blanks_Paper7
   If IWPDPaper = Paper8 Then Fill_in_the_Blanks_Paper8
   If IWPDPaper = Paper9 Then Fill_in_the_Blanks_Paper9
   If IWPDPaper = Paper10 Then Fill_in_the_Blanks_Paper10
   If IWPDPaper = Paper11 Then Fill_in_the_Blanks_Paper11
   If IWPDPaper = Paper12 Then Fill_in_the_Blanks_Paper12
      
ElseIf PaperType = "Matching Columns" Then
       If IWPDPaper = Paper1 Then Matching_Columns_Paper1
       If IWPDPaper = Paper2 Then Matching_Columns_Paper2
       If IWPDPaper = Paper3 Then Matching_Columns_Paper3
       If IWPDPaper = Paper4 Then Matching_Columns_Paper4
       If IWPDPaper = Paper5 Then Matching_Columns_Paper5
       If IWPDPaper = Paper6 Then Matching_Columns_Paper6
       If IWPDPaper = Paper7 Then Matching_Columns_Paper7
       If IWPDPaper = Paper8 Then Matching_Columns_Paper8
       If IWPDPaper = Paper9 Then Matching_Columns_Paper9
       If IWPDPaper = Paper10 Then Matching_Columns_Paper10
       If IWPDPaper = Paper11 Then Matching_Columns_Paper11
       If IWPDPaper = Paper12 Then Matching_Columns_Paper12
       If IWPDPaper = Paper13 Then Matching_Columns_Paper13
       If IWPDPaper = Paper14 Then Matching_Columns_Paper14
              
ElseIf PaperType = "True Or False" Then
       If IWPDPaper = Paper1 Then TrueorFalse_Paper1
       If IWPDPaper = Paper2 Then TrueorFalse_Paper2
       If IWPDPaper = Paper3 Then TrueorFalse_Paper3
       If IWPDPaper = Paper4 Then TrueorFalse_Paper4
       If IWPDPaper = Paper5 Then TrueorFalse_Paper5
       If IWPDPaper = Paper6 Then TrueorFalse_Paper6
       If IWPDPaper = Paper7 Then TrueorFalse_Paper7
       If IWPDPaper = Paper8 Then TrueorFalse_Paper8
       If IWPDPaper = Paper9 Then TrueorFalse_Paper9
       If IWPDPaper = Paper10 Then TrueorFalse_Paper10
       If IWPDPaper = Paper11 Then TrueorFalse_Paper11
       If IWPDPaper = Paper12 Then TrueorFalse_Paper12
       If IWPDPaper = Paper13 Then TrueorFalse_Paper13
       If IWPDPaper = Paper14 Then TrueorFalse_Paper14
              
ElseIf PaperType = "Descriptive Questions" Then
       If IWPDPaper = Paper1 Then questions_paper1
       If IWPDPaper = Paper2 Then questions_paper2
       If IWPDPaper = Paper3 Then questions_paper3
       If IWPDPaper = Paper4 Then questions_paper4
       If IWPDPaper = Paper5 Then questions_paper5
       If IWPDPaper = Paper6 Then questions_paper6
       
ElseIf PaperType = "Multiple Choice" Then
       If IWPDPaper = Paper1 Then multiplechoice_Paper1
       If IWPDPaper = Paper2 Then multiplechoice_Paper2
       If IWPDPaper = Paper3 Then multiplechoice_Paper3
       If IWPDPaper = Paper4 Then multiplechoice_Paper4
       If IWPDPaper = Paper5 Then multiplechoice_Paper5
       If IWPDPaper = Paper6 Then multiplechoice_Paper6
       If IWPDPaper = Paper7 Then multiplechoice_Paper7
       If IWPDPaper = Paper8 Then multiplechoice_Paper8
       If IWPDPaper = Paper9 Then multiplechoice_Paper9
       If IWPDPaper = Paper10 Then multiplechoice_Paper10
       If IWPDPaper = Paper11 Then multiplechoice_Paper11
End If

If mCat = Fill_Questions Then
   IWPD = StrQuestions(Number)
ElseIf mCat = Fill_Options Then
       IWPD = stroptions(Number)
ElseIf mCat = Fill_Answers Then
       IWPD = stranswers(Number)
       
ElseIf mCat = Matching_ColumnA Then
       IWPD = strcolumnA(Number)
ElseIf mCat = Matching_ColumnB Then
       IWPD = strcolumnB(Number)
ElseIf mCat = Matching_Answers Then
       IWPD = stranswers(Number)
       
ElseIf mCat = Questions Then
       IWPD = StrQuestions(Number)
       
ElseIf mCat = TF_Questions Then
       IWPD = StrQuestions(Number)
ElseIf mCat = TF_Answers Then
       IWPD = stranswers(Number)
       
ElseIf mCat = MC_Questions Then
       IWPD = strmultiple(Number)
       
ElseIf mCat = MC_Options_1 Then
       IWPD = strmuloption(Number)
       
ElseIf mCat = MC_Options_2 Then
       IWPD = strmuloption(Number)
       
ElseIf mCat = MC_Options_3 Then
       IWPD = strmuloption(Number)
       
ElseIf mCat = MC_Options_4 Then
       IWPD = strmuloption(Number)
       
ElseIf mCat = MC_Answer Then
       IWPD = strmulanswer(Number)
              
End If

End Function

Private Sub Fill_in_the_Blanks_Paper1()
StrQuestions(1) = "The lock rate at which the pulses are transmitted by a serial device is called dash."
StrQuestions(2) = "The process of directing a message to the recipient's host is called dash."
StrQuestions(3) = "The programming language to create web pages is dash."
StrQuestions(4) = "Unique address to identify a page is called dash."
StrQuestions(5) = "Areas on an image map that viewers can click to open different URLs are called dash."
StrQuestions(6) = "Network is defined as a collection of dash that are able to communicate with each other."
StrQuestions(7) = "Small agglomeratioins of hosts are called dash."
StrQuestions(8) = "In IP, the scheme of directing data to a remote host is dash and packets are referred as dash."
StrQuestions(9) = "Mapping of hostnames onto IP-addresses is called dash."
StrQuestions(10) = "Mapping of IP addresses to hardware addresses is called dash."

stroptions(1) = "Baud rate"
stroptions(2) = "Routing"
stroptions(3) = "HTML"
stroptions(4) = "URL"
stroptions(5) = "hotspots"
stroptions(6) = "hosts"
stroptions(7) = "sites"
stroptions(8) = "routing,datagram"
stroptions(9) = "hostname resolution"
stroptions(10) = "address resolution"

stranswers(1) = "hostname resolution"
stranswers(2) = "routing,datagram"
stranswers(3) = "hotspots"
stranswers(4) = "hosts"
stranswers(5) = "URL"
stranswers(6) = "sites"
stranswers(7) = "Routing"
stranswers(8) = "HTML"
stranswers(9) = "Baud rate"
stranswers(10) = "address resolution"

End Sub

Private Sub Fill_in_the_Blanks_Paper2()
StrQuestions(1) = "The maximum size of datagram that can be processed by that particular piece of hardware is called dash."
StrQuestions(2) = "The mechanism used to map IP addresses onto Ethernet addresses is dash."
StrQuestions(3) = "A dash is a host that is connected to two or more physical networks simultaneously and is configured to switch packets between them."
StrQuestions(4) = "The current state of a existing serial connection between two points, i.e. the number of bits transferred per second is measured by dash."
StrQuestions(5) = "PHP is a scripting language to create dash."
StrQuestions(6) = "A dash is an intelligent connecting wire between two network cards."
StrQuestions(7) = "A dash is an intelligent insulator."
StrQuestions(8) = "Text containing connections withing it to other documents is called dash."
StrQuestions(9) = "Bridges operate at the dash layer of the OSI model."
StrQuestions(10) = "In hypertext terminology nodes without further links are called dash."

stroptions(1) = "Leaf nodes"
stroptions(2) = "Maximum Transfer unit or MTU"
stroptions(3) = "Gateway"
stroptions(4) = "Data-link"
stroptions(5) = "ARP"
stroptions(6) = "Dynaimc HTMl pages"
stroptions(7) = "Hyperlinks"
stroptions(8) = "Bridge"
stroptions(9) = "Firewall"
stroptions(10) = "Bit rate"

stranswers(1) = "Bridge"
stranswers(2) = "Dynaimc HTMl pages"
stranswers(3) = "Hyperlinks"
stranswers(4) = "Maximum Transfer unit or MTU"
stranswers(5) = "Gateway"
stranswers(6) = "Leaf nodes"
stranswers(7) = "Data-link"
stranswers(8) = "ARP"
stranswers(9) = "Bit rate"
stranswers(10) = "Firewall"

End Sub

Private Sub Fill_in_the_Blanks_Paper3()
StrQuestions(1) = "A hypertext link connects two nodes and is normally directed in the sense that it points from one node called dash node to another called dash node."
StrQuestions(2) = "The term dash is used when referring to systems having a variety of node types like text, graphics, sound etc."
StrQuestions(3) = "Hypertext is a network of dash."
StrQuestions(4) = "A dash is a collection of data organized around a specific topic and linked to another body of information"
StrQuestions(5) = "A dash is information embedded in a node that connects it in some way to another node."
StrQuestions(6) = "A dash is a visual cue in a node that alerts a user that a link exists."
StrQuestions(7) = "A dash is a visual representation of link in a node."
StrQuestions(8) = "A dash is a record of the nodes that a user has accessed in viewing a hypertext network."
StrQuestions(9) = "dash are devices which convert digital signals into analog for transmission over telephone lines."
StrQuestions(10) = "dash performs the opposite function of modem"

stroptions(1) = "button/hotspot"
stroptions(2) = "codec"
stroptions(3) = "trail"
stroptions(4) = "modems"
stroptions(5) = "anchor,destination"
stroptions(6) = "node"
stroptions(7) = "Nodes"
stroptions(8) = "Link"
stroptions(9) = "Button/hotspot"
stroptions(10) = "Hypermedia"

stranswers(1) = "Nodes"
stranswers(2) = "button/hotspot"
stranswers(3) = "codec"
stranswers(4) = "trail"
stranswers(5) = "modems"
stranswers(6) = "anchor,destination"
stranswers(7) = "Button/hotspot"
stranswers(8) = "node"
stranswers(9) = "Hypermedia"
stranswers(10) = "Link"
End Sub

Private Sub Fill_in_the_Blanks_Paper4()
StrQuestions(1) = "Frequencies below dash MHz are radio frequencies and higher are microwave frequencies."
StrQuestions(2) = "dash is a satellite based nationwide network connecting all the districts of India."
StrQuestions(3) = "A dash is a public discussion area where people can post messages"
StrQuestions(4) = "Programs called dash run automated lists."
StrQuestions(5) = "Hypertext documents on the Web are commonly referred to as dash."
StrQuestions(6) = "dash is called the biggest BBS."
StrQuestions(7) = "In Anonymous FTP you log in as dash and your password is dash."
StrQuestions(8) = "Unsolicited commercial e-mail is called dash"
StrQuestions(9) = "A dash is a grid of rectangles, each of which can hold text and images."
StrQuestions(10) = "Each rectangle in a tabl that can hold information is called a dash."

stroptions(1) = "Pages"
stroptions(2) = "Usenet"
stroptions(3) = "Bulletin board"
stroptions(4) = "Spam"
stroptions(5) = "List servers"
stroptions(6) = "Anonymous , your e-mail address"
stroptions(7) = "Table"
stroptions(8) = "NICNET"
stroptions(9) = "1000"
stroptions(10) = "cell"

stranswers(1) = "Usenet"
stranswers(2) = "Pages"
stranswers(3) = "Anonymous , your e-mail address"
stranswers(4) = "Table"
stranswers(5) = "Bulletin board"
stranswers(6) = "Spam"
stranswers(7) = "NICNET"
stranswers(8) = "1000"
stranswers(9) = "cell"
stranswers(10) = "List servers"

End Sub

Private Sub Fill_in_the_Blanks_Paper5()
StrQuestions(1) = "The cells of a table are arranged vertically into dash and horizontally into dash"
StrQuestions(2) = "The width of a border between any two table ells is called dash."
StrQuestions(3) = "Empty space along the inner edge of a cell is called dash."
StrQuestions(4) = "HTML commands surrounded by < and > are called dash."
StrQuestions(5) = "A link is associated with something on a web page by using the dash tag."
StrQuestions(6) = "An dash is away to make different parts of a graphic link to different web pages."
StrQuestions(7) = "The part of a web site where guests can make public comments about the site is called dash."
StrQuestions(8) = "Address where all host part its are 0 points to dash."
StrQuestions(9) = "Address where all host part nits are 1 is the dash."
StrQuestions(10) = "Address 0.0.0.0 is called the dash"

stroptions(1) = "Image map"
stroptions(2) = "<A>"
stroptions(3) = "default route"
stroptions(4) = "Padding"
stroptions(5) = "Guestbook"
stroptions(6) = "broadcast address"
stroptions(7) = "network"
stroptions(8) = "columns, rows"
stroptions(9) = "tgs"
stroptions(10) = "spacing"

stranswers(1) = "Guestbook"
stranswers(2) = "broadcast address"
stranswers(3) = "network"
stranswers(4) = "Image map"
stranswers(5) = "<A>"
stranswers(6) = "Padding"
stranswers(7) = "columns, rows"
stranswers(8) = "tgs"
stranswers(9) = "spacing"
stranswers(10) = "default route"

End Sub
Private Sub Fill_in_the_Blanks_Paper6()
StrQuestions(1) = "Address 127.0.0.0 is called the dash."
StrQuestions(2) = "Two major components of HTML are dash and dash."
StrQuestions(3) = "IP addresses are dash bit numbers."
StrQuestions(4) = "Language used for creating Web pages is called dash."
StrQuestions(5) = "dash in computers is a device that protects a private network from public part(the Internet)"
StrQuestions(6) = "An analog device that sends digital signals over ordinary telephone lines is dash."
StrQuestions(7) = "The process by which a host with multiple network connections decides where to deliver IP datagram is called dash."
StrQuestions(8) = "dash is seen by many as a replacement for POP mail."
StrQuestions(9) = "The maximum size of datagram that can be processed by a piece of hardware is called dash."
StrQuestions(10) = "dash interface allows you to develop and test networking software without using a real network."

stroptions(1) = "32"
stroptions(2) = "loopback address"
stroptions(3) = "loopback"
stroptions(4) = "MTU"
stroptions(5) = "tags, attributes"
stroptions(6) = "IMAP"
stroptions(7) = "Modem"
stroptions(8) = "Firewall"
stroptions(9) = "routing"
stroptions(10) = "html"

stranswers(1) = "MTU"
stranswers(2) = "tags, attributes"
stranswers(3) = "32"
stranswers(4) = "loopback address"
stranswers(5) = "loopback"
stranswers(6) = "Firewall"
stranswers(7) = "routing"
stranswers(8) = "html"
stranswers(9) = "IMAP"
stranswers(10) = "Modem"

End Sub

Private Sub Fill_in_the_Blanks_Paper7()
StrQuestions(1) = "dash supply information on a given domain or set of domains."
StrQuestions(2) = "dash is IP protocol for parallel line interface."
StrQuestions(3) = "dash is an application for resolving hostnames to IP addresses."
StrQuestions(4) = "Mail header is offset from mail body by a dash."
StrQuestions(5) = "dash adds dynamic content to Web pages."
StrQuestions(6) = "dash is used for remote login."
StrQuestions(7) = "dash displays Internet resources in a hierarchical menu format."
StrQuestions(8) = "dash allows one communication channel to send many transmissions at the same time."
StrQuestions(9) = "WWW documents and files are called dash."
StrQuestions(10) = "dash is a software used to view Web pages."

stroptions(1) = "gopher"
stroptions(2) = "pages"
stroptions(3) = "name servers"
stroptions(4) = "PLIP"
stroptions(5) = "browser"
stroptions(6) = "Telnet"
stroptions(7) = "multiplexer"
stroptions(8) = "Blank line"
stroptions(9) = "BIND"
stroptions(10) = "SSI"

stranswers(1) = "PLIP"
stranswers(2) = "gopher"
stranswers(3) = "pages"
stranswers(4) = "multiplexer"
stranswers(5) = "name servers"
stranswers(6) = "Blank line"
stranswers(7) = "BIND"
stranswers(8) = "SSI"
stranswers(9) = "browser"
stranswers(10) = "Telnet"

End Sub

Private Sub Fill_in_the_Blanks_Paper8()
StrQuestions(1) = "Using browser to travel the web is called dash."
StrQuestions(2) = "Internet Explorer and dash are the two most popular Web browsers."
StrQuestions(3) = "dash is a connection from on Web page to another."
StrQuestions(4) = "One special kind of link that makes it easier for people visiting the Web site tosend you e-mail is called a dash link."
StrQuestions(5) = "Alternative to mailto like is to use a powerful but complex program called dash gateway."
StrQuestions(6) = "A text description of the image for the benefit of non graphical users is called dash label."
StrQuestions(7) = "dash is the text that is slightly raised above a line."
StrQuestions(8) = "A monospace font similar to one a typewriter can produce is available as an otion for Web page formatting as dash."
StrQuestions(9) = "If you want to make sure that text on a page will appear the way you typed it, you can designate the text as dash."
StrQuestions(10) = "The element of a Web page, which is a good way to break up a heavy text document, is dash."

stroptions(1) = "surfing"
stroptions(2) = "headings"
stroptions(3) = "teletype"
stroptions(4) = "mailto"
stroptions(5) = "CGI"
stroptions(6) = "preformatted"
stroptions(7) = "superscript"
stroptions(8) = "Alt"
stroptions(9) = "link"
stroptions(10) = "Netscape Navigator"

stranswers(1) = "Surfing"
stranswers(2) = "Netscape Navigator"
stranswers(3) = "link"
stranswers(4) = "mailto"
stranswers(5) = "mailto"
stranswers(6) = "Alt"
stranswers(7) = "superscript"
stranswers(8) = "teletype"
stranswers(9) = "CGI"
stranswers(10) = "preformatted"

End Sub

Private Sub Fill_in_the_Blanks_Paper9()
StrQuestions(1) = "With most of the Web browsers, Strong formatting is equivalent to dash type text formatting."
StrQuestions(2) = "Width of the border between any two table cells is called dash."
StrQuestions(3) = "The empty space along the inner edge of cell is called dash."
StrQuestions(4) = "Putting one set of HTML tags inside another is called dash."
StrQuestions(5) = "dash program allows remote login into another computer."
StrQuestions(6) = "dash program determines if a node is responding to TCP/IP communication."
StrQuestions(7) = "dash connects a network to another of the same ring through a point to point link."
StrQuestions(8) = "IP fragments the data into smaller units if it is greater than dash in size."
StrQuestions(9) = "IP protocol is an implementation of the OSI dash layer."
StrQuestions(10) = "IP routes data frames called Internet dash around the Internet."

stroptions(1) = "spacing"
stroptions(2) = "bold"
stroptions(3) = "datagram"
stroptions(4) = "network"
stroptions(5) = "telnet"
stroptions(6) = "64 KB"
stroptions(7) = "Ping"
stroptions(8) = "nesting"
stroptions(9) = "router"
stroptions(10) = "padding"

stranswers(1) = "64 KB"
stranswers(2) = "Ping"
stranswers(3) = "spacing"
stranswers(4) = "bold"
stranswers(5) = "datagram"
stranswers(6) = "telnet"
stranswers(7) = "network"
stranswers(8) = "router"
stranswers(9) = "padding"
stranswers(10) = "nesting"

End Sub

Private Sub Fill_in_the_Blanks_Paper10()
StrQuestions(1) = "Nodes on a TCP/IP network are either hosts or dash."
StrQuestions(2) = "any nodes on a TCP/IP network that run application software or re terminals are dash."
StrQuestions(3) = "Any node that routes TCP/IP packets between networks is called a TCP/IP dash node."
StrQuestions(4) = "dash program traces the route of ah IP packet through the Internet."
StrQuestions(5) = "A security gateway that protects a network from intrusion from outside sources is called dash."
StrQuestions(6) = "dash allow two or more companies to share a parts of their intranets related to ring projects."
StrQuestions(7) = "A dash connects to a number of clients; it acts on behalf of other clients and sends requests from the clients to the server."
StrQuestions(8) = "HTML dash contain special formatting commands and are contained withing a less than (<) and a greater than (>) symbol."
StrQuestions(9) = "To enter a new line in HTML text we use dash tag."
StrQuestions(10) = "Each part of a list in HTML is defined with the dash tag."

stroptions(1) = "Proxy"
stroptions(2) = "gateways"
stroptions(3) = "<BR>"
stroptions(4) = "<LI>"
stroptions(5) = "Firewall"
stroptions(7) = "Extranets"
stroptions(6) = "tags"
stroptions(7) = "Extranets"
stroptions(8) = "gateway"
stroptions(9) = "hosts"
stroptions(10) = "traceroute"

stranswers(1) = "Firewall"
stranswers(2) = "Proxy"
stranswers(3) = "gateways"
stranswers(4) = "<LI>"
stranswers(5) = "gateway"
stranswers(6) = "<BR>"
stranswers(7) = "hosts"
stranswers(8) = "traceroute"
stranswers(9) = "tags"
stranswers(10) = "Extranets"

End Sub


Private Sub Fill_in_the_Blanks_Paper11()
StrQuestions(1) = "Some browsers allow the type of numbered list to be denied with the dash tag."
StrQuestions(2) = "Type of a bullet point in an unordered list is defined with dash tag allowed by some browsers."
StrQuestions(3) = "Default color of a HTML link is dash."
StrQuestions(4) = "HTML uses 24 bits to denote colors, giving a maximum of dash colors."
StrQuestions(5) = "GIF images can contain dash colors."
StrQuestions(6) = "A horizontal line an be added in a Web page with dash tag."
StrQuestions(7) = "dash allow an HTML document to analyze, parse and store information received from a form."
StrQuestions(8) = "dash is the language of choice for writing CGI scripts as it has the advantage of running on any computer not being a pre compiled language."
StrQuestions(9) = "dash tag is used to define the start of a list of menu options."
StrQuestions(10) = "A Java program is called from an HTM: script with the dash tag."

stroptions(1) = "<OL TYPE =x>"
stroptions(2) = "Blue"
stroptions(3) = "APPLEt"
stroptions(4) = "<SELECT>"
stroptions(5) = "PERL"
stroptions(6) = "<UL TYPE = shape>"
stroptions(7) = "16.77 million"
stroptions(8) = "<HR>"
stroptions(9) = "CGI scripts"
stroptions(10) = "256 "

stranswers(1) = "CGI scripts"
stranswers(2) = "<OL TYPE =x>"
stranswers(3) = "<HR>"
stranswers(4) = "<SELECT>"
stranswers(5) = "APPLEt"
stranswers(6) = "PERL"
stranswers(7) = "<UL TYPE = shape>"
stranswers(8) = "<SELECT>"
stranswers(9) = "16.77 million"
stranswers(10) = "256 "

End Sub

Private Sub Fill_in_the_Blanks_Paper12()
StrQuestions(1) = "dash examines the network address field and define the best route for the mail acket."
StrQuestions(2) = "Java applets have file names with dash extension."
StrQuestions(3) = "An dash is a way to make parts of a graphic like to different web pages."
StrQuestions(4) = "dash is a Microsoft technology that is similar to Java  Beans in function and allows  interactive programs to be offered on Web pages."
StrQuestions(5) = "A program written in Java that is presented on a Web page and runs automatically when a page loads is called an dash."
StrQuestions(6) = "dash is a temporary stofage area for pages that a Web browser has rcently viewed."
StrQuestions(7) = "An Internet addressing system that associates a domain name with numeric address is called dash."
StrQuestions(8) = "Receiving a file form another site and sotring it on your system is called dash."
StrQuestions(9) = ""
StrQuestions(10) = ""

stroptions(1) = "DNS"
stroptions(2) = "Routers"
stroptions(3) = "Image map"
stroptions(4) = "ActiveX"
stroptions(5) = "downloading"
stroptions(6) = ".class"
stroptions(7) = "Cache"
stroptions(8) = "Applet"
stroptions(9) = ""
stroptions(10) = ""

stranswers(1) = "Applet"
stranswers(2) = "Cache"
stranswers(3) = "Routers"
stranswers(4) = "Image map"
stranswers(5) = "downloading"
stranswers(6) = "DNS"
stranswers(7) = "ActiveX"
stranswers(8) = ".class"
stranswers(9) = ""
stranswers(10) = ""

End Sub

Private Sub questions_paper1()
StrQuestions(1) = "Discuss briefly as to what is the internet and how it worked. Is the Internet the same as the World Wide Web?"
StrQuestions(2) = "How are the Web and Internet are interrelated?"
StrQuestions(3) = "What applications are needed to use the Internet?"
StrQuestions(4) = "How is the Internet changing to meet information demands?"
StrQuestions(5) = "Describe the Web as a huge global market. What makes the prospect so bright?"
StrQuestions(6) = "HTML is more of a script than a language. Do you agree? Give reasons."
StrQuestions(7) = "Name some of the search tools(engines)"
StrQuestions(8) = "Discuss the HTML features not supported by all browsers."
StrQuestions(9) = "What is a n interlaced image?"
StrQuestions(10) = "Differentiate between POP and IMAP."

End Sub

Private Sub questions_paper2()
StrQuestions(1) = " Write a short note about Usenet."
StrQuestions(2) = " What is IP multicasting?"
StrQuestions(3) = " What IP Masquerading?"
StrQuestions(4) = " What is a Proxy server?"
StrQuestions(5) = " What is firewall? What are its uses?"
StrQuestions(6) = " Describe an Intranet? What hardware and software will be needed to establish one?"
StrQuestions(7) = " Discuss IMAP as replacement for POP."
StrQuestions(8) = " Discuss the communication architecture with reference to OSI reference model."
StrQuestions(9) = " Discuss various applications of hypertext mechanism."
StrQuestions(10) = "Write a short note on ISDN."
End Sub
Private Sub questions_paper3()
StrQuestions(1) = " What is a search engine? Discuss with few examples."
StrQuestions(2) = " What is netiquette? Give some examples."
StrQuestions(3) = " What is flaming? How is it different from spamming?"
StrQuestions(4) = " What is a mail header? What does it contain?"
StrQuestions(5) = " What is a MIME attachment? How does it enhance the e mail facility?"
StrQuestions(6) = " What are the uses of mail filters?"
StrQuestions(7) = " What are mailing lists? How are they different from newsgroups?"
StrQuestions(8) = " What is push technology? Give examples."
StrQuestions(9) = " What is a web server? How does it function?"
StrQuestions(10) = "    What is web casting? Discuss its uses?"
End Sub
Private Sub questions_paper4()
StrQuestions(1) = " What is a shareware? How is it different from freeware? Discuss with examples."
StrQuestions(2) = " What is streaming media? How is it implemented in browsers?"
StrQuestions(3) = " What is VRML? How does it differ from HTML?"
StrQuestions(4) = " What are plug-ins? How are they useful?"
StrQuestions(5) = " What is hypermedia? How is it different from hypertext?"
StrQuestions(6) = " Consider the following HTTP request:    GET /art/gallery.dll/animals?   animal = dog&color = black HTTP/1.0 Connection: Keep-Alive   User-Agent: Mozilla/3.0b4Gold (WinNT: 1)    Host: WWW.Tsite.com:1024    Accept: image/gif, image/x-xbitmap,image/jpeg,image/pjpeg, *.*  Describe the various headers in the above request."
StrQuestions(7) = " What are the advantages of using e -m ail? What are the disadvantages ?"
StrQuestions(8) = " Give an overview of SMTP. Discuss its drawbacks. Discuss theenhancements available in MIME to overcome these drawbacks."
StrQuestions(9) = " Discuss the advantages and disadvantages of the World Wide Web.  "
StrQuestions(10) = "What is a URL?  What does it contain? Describe its inplementation in FTP, e=mail, HTTP and telnet programs."

End Sub
Private Sub questions_paper5()
StrQuestions(1) = " Discuss  advantages of JAVA over conventional software languages."
StrQuestions(2) = " What is a scripting language? How is it different from the other high-level computer languages?"
StrQuestions(3) = " What are the differences between JAA and JavaScript? What advantage JavaScript has over HTML?"
StrQuestions(4) = " Discuss various network topologies, along with their advantages and disadvantages."
StrQuestions(5) = " discuss the broad features of ActiveX technology and compare it with java."
StrQuestions(6) = " Discuss the major implementation of mail servers and mail clients in different OS platforms like Unix, DOS, Windows 9x  and Mac OS."
StrQuestions(7) = " What is IRC? Discuss how it is implemented in various environments."
StrQuestions(8) = " discuss the features of any two of the following email programs:    (a) Microsoft Outlook   (b) Pegasus Mail    (c) Pine    (d) Eudora Mail (e) Netscape Messanger"
StrQuestions(9) = " Compare the advantages and disadvantages of using a text based Web browser like Lynx with a graphic Web browser like Internet Explorer?"
StrQuestions(10) = "Discuss the features of any two of the following browsers:   (a) Netscape Navigator (b) Internet explorer   (c) NCSA Mosaic     (d) Lynx"

End Sub
Private Sub questions_paper6()
StrQuestions(1) = "Discuss the features of any Web authoring and editing tool."
StrQuestions(2) = "What is a Web server application? How is it different from a Web Browser?"
StrQuestions(3) = "Discuss how secure connections are supported through SSL. How is the technique helpful in e-commerce?"
StrQuestions(4) = "What is  off line browsing? Discuss its usefulness?"
StrQuestions(5) = "What is a mail server program? Discuss its uses."
StrQuestions(6) = "HTML forms are a great alternative to CGI based forms since they are very easy to implement  and maintain and require no complicates setup on the web site. - Discuss."
StrQuestions(7) = "Write short  notes on any two of the following  (a) Freeware   (b) Shareware   (c) Middleware  (d) Author ware  (e) Vaporware"
StrQuestions(8) = "Describe the procedure of building a web site."
StrQuestions(9) = "What are the languages/methods available for web database programming. Discuss them briefly."
StrQuestions(10) = "What are the software utilities you would like to have if you were a progessional web designer and programmer? Why?"

End Sub

Private Sub TrueorFalse_Paper1()
StrQuestions(1) = "Host names are case sensitive."
StrQuestions(2) = "IP addresses are 32-bit numbers."
StrQuestions(3) = "The scheme of directing data to a remote host is called routing."
StrQuestions(4) = "Internet Engineering Task Force(IETF) assigns port numbers to well known services."
StrQuestions(5) = "NIC is the central authority for assigning numbers to Internet sites."
StrQuestions(6) = "Every machine must be assigned a number unique to the networking envirinment."
StrQuestions(7) = "IP always checks the datagram for errors before transmitting it."
StrQuestions(8) = "It is the responsibility of the communicating hosts to check integrity and completeness of data."
StrQuestions(9) = "A leased line is not connected to the telephone exchange and hence does not proide dial tone or ring signal."
StrQuestions(10) = "Internet hosts are often computers, but need not be."

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""
End Sub

Private Sub TrueorFalse_Paper2()
StrQuestions(1) = " Small agglomerations of hosts are called sites."
StrQuestions(2) = " PPP is a network protocol to send network traffic along parallel lines."
StrQuestions(3) = " Routing involves error checking as well as cost optimization."
StrQuestions(4) = " VRML is a programming language that is replacing HTML for creating web pages."
StrQuestions(5) = " Individual files that make up a web site are called pages."
StrQuestions(6) = " UUCP allows to forward hobs and files through several hosts."
StrQuestions(7) = " UUCP means a low cost dial up network link."
StrQuestions(8) = " ARPANET was precursor to the Internet."
StrQuestions(9) = " To add a host to thinnest installation, you have to disrupt network service."
StrQuestions(10) = "Adding a host to thick Ethernet does bring down the network."

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""
End Sub

Private Sub TrueorFalse_Paper3()
StrQuestions(1) = "In Ethernet network system, if data collision occurs, the send is aborted."
StrQuestions(2) = "The main benefit of IP is that it turns physically dissimilar networks into one apparently homogenous network."
StrQuestions(3) = "IP is hardware dependent."
StrQuestions(4) = "mapping of hostnames onto IP addresses is called address resolution."
StrQuestions(5) = "PPP is a serial protocol."
StrQuestions(6) = "PPP has more features than SLIP."
StrQuestions(7) = "SLIP has advantage over PPP in that it is not limited to transporting IP-datagrams, but can allow transmission of any type of datagram."
StrQuestions(8) = "IP, by intent is not reliable."
StrQuestions(9) = "If IP is unable to forward a data , it discards it."
StrQuestions(10) = "TCP is responsible for reliability of communication."

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""
End Sub

Private Sub TrueorFalse_Paper4()
StrQuestions(1) = "TCP allows several concurrent logins to the same host."
StrQuestions(2) = "TCP/IP defines an abstract interface through which hardware is addressed to hide the diversity of equipment that may be used in a networking environment."
StrQuestions(3) = "The maximum size of datagram that can be processed by that particular piece of hardware is called Maximum Transfer Unit(MTU)."
StrQuestions(4) = "If you are running a local network that does not have TCP/IP traffic with other networked, you may assign the host address according to you personal preference."
StrQuestions(5) = "The IP address format is often referred to as dotted quad notation."
StrQuestions(6) = "Host numbers with octets all 0 or all 255 are reserved for special purposes."
StrQuestions(7) = "Network 127.0.0.0 is reserved for IP traffic local to your host."
StrQuestions(8) = "Loop back interface (Network 127.0.0.1) allows you to develop and test networking software without ever using a 'real' network."
StrQuestions(9) = "Hosts with identical IP-network numbers should be found n the same network."
StrQuestions(10) = "IP allows you to subdivide an IP-network into several subnets."

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""
End Sub

Private Sub TrueorFalse_Paper5()
StrQuestions(1) = "The network administrator generates subnets."
StrQuestions(2) = "Sub netting is completely invisible to the outside world."
StrQuestions(3) = "A host can have more that one IP-address."
StrQuestions(4) = "A Gateway is assigned one IP-address irrespective of the number of networks it is connected to."
StrQuestions(5) = "DNS organizes host names in a hierarchy of networks."
StrQuestions(6) = "DNS is a giant distributed of name servers."
StrQuestions(7) = "Name servers supply information on a given domain or set of domains."
StrQuestions(8) = "A host may have more than one name, but one of these names must be identified as the official one."
StrQuestions(9) = "Ping is the networking equivalent of a zonal device and is used toverigy that a given address is reachable."
StrQuestions(10) = "UUCP executes program on a remote host."

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""
End Sub

Private Sub TrueorFalse_Paper6()
StrQuestions(1) = "The values denoted by bit rate and baud ate are usually same."
StrQuestions(2) = "Hardware handshake is optional for serial communication."
StrQuestions(3) = "Domain name is generally used to set the host's NIS domain, whichmay be entirely different from the DNS domain your host belongs to."
StrQuestions(4) = "One need not configure the loop back interface, unless doing some network testing or development work as applications do not use it during normal operation."
StrQuestions(5) = "Routers connect networks to one another."
StrQuestions(6) = "When you see high packet loss rates when pinging a host behind one or more gateways, this may hint at a very congested network and not at technical deficiencies."
StrQuestions(7) = "PLIP are point to point links."
StrQuestions(8) = "SLIP, PPP, and PLIP are all point to point links."
StrQuestions(9) = "SLIP and PPP are different from PLIP as they dial up modem for connection."
StrQuestions(10) = "BIND is an application for resolving hostnames to IP addresses."

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""
End Sub

Private Sub TrueorFalse_Paper7()
StrQuestions(1) = "PAP and CHAP are both client authentication protocols in PPP."
StrQuestions(2) = "With PPP, each system may not enquire its peer to authenticate itself using authentication protocol."
StrQuestions(3) = "PAP is comparatively more secure than CHAP."
StrQuestions(4) = "UUCP is never used over a dial up network."
StrQuestions(5) = "UUCP does not call the remote system immediately to execute a hob, rather stores the job description temporarily."
StrQuestions(6) = "UUCP associates a grade with each job it transmits."
StrQuestions(7) = "Unlike TCP-IP networking a hostname in UUCP networking does not have to meet any standards."
StrQuestions(8) = "With UUCP you cannot execute file transfers with systems which are not directly connected."
StrQuestions(9) = "Mail header is offset from mail body by a line of asterisks."
StrQuestions(10) = "Recipient's log-in name must form a pat of his email address."

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""
End Sub

Private Sub TrueorFalse_Paper8()
StrQuestions(1) = "On UUCP networks, he transport software performs the mail routing."
StrQuestions(2) = "The advantage of bank path notation in e-mail is hat you do not have to remember the network topology."
StrQuestions(3) = "Usenet maps can be used to easily generate full paths from your host to any destination site."
StrQuestions(4) = "You cannot query a name server over UUCP."
StrQuestions(5) = "UUCP transport programs cannot deal with fully qualified domain names."
StrQuestions(6) = "It is possible to force mail into a misconfigured remote system and to get a reply."
StrQuestions(7) = "Push technology places a heavy load on the client system."
StrQuestions(8) = "The problem with pulling news is that it needs tight control by theserver."
StrQuestions(9) = "NNTP allows the knowledgeable to insert articles into news stream with false sender."
StrQuestions(10) = "The routing table allows multiple routes per destination."

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""
End Sub

Private Sub TrueorFalse_Paper9()
StrQuestions(1) = "Spamming and flaming are very much the same."
StrQuestions(2) = "IMAP is more powerful than POP3."
StrQuestions(3) = "POP is preferable to IMAP in situations like remote or roaming access, because snooping of passwords is a concern."
StrQuestions(4) = "No matter what machine you go to, if you open a telnet connection to 127.0.0.1 you will reach the local host."
StrQuestions(5) = "A fixed IP address cannot be allotted if host is mobile."
StrQuestions(6) = "IP multicast allows an arbitrary number of IP hosts on disparate IP networks to have IP datagram simultaneously routed to them."
StrQuestions(7) = "In Token ring protocol collisions are avoided by allowing only one staion on the LAN the right to transmit at a time."
StrQuestions(8) = "In asynchronous transfer mode, the network packages data into standard size blocks, which can be efficiently conveyed from point to point."
StrQuestions(9) = "POP is preferable  to IMAP as it allows manipulation of messages in the server, remote creation and deletion of folders in the server, concurrent access to shared mail folders etc."
StrQuestions(10) = "IP is a connectionless protocol."

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""
End Sub

Private Sub TrueorFalse_Paper10()
StrQuestions(1) = "IP is cheap, fast but unreliable."
StrQuestions(2) = "SSI adds dynamic content to web pages."
StrQuestions(3) = "Java runtime interpreter is not required to run Java programs in Mac and UNIX environments."
StrQuestions(4) = "TCP/IP is platform independent."
StrQuestions(5) = "You cannot telnet/FTP from the bridge."
StrQuestions(6) = "Bridge is a part of the path."
StrQuestions(7) = "Every port on a bridge is assigned the same physical Ethernet address."
StrQuestions(8) = "A bridge does not deal with packets the way a hub does, and hence contributes to timing problem on a network."
StrQuestions(9) = "IMAP allows users to access their e-mail from different computers without having to download it."
StrQuestions(10) = "IP masquerade allows a set of machines to invisibly access the Internet hidden behind a gateway system."

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""
End Sub

Private Sub TrueorFalse_Paper11()
StrQuestions(1) = "IP numbers are assigned to interfaces and not hosts."
StrQuestions(2) = "PLIP interface is normally faster than SLIP interface."
StrQuestions(3) = "The role of TCP is to divide a message into packets and specify  the destination address."
StrQuestions(4) = "Some sites allow you to use FTP even if you do not have an account or password."
StrQuestions(5) = "You cannot look at the contents of a file while you are connected via FTP."
StrQuestions(6) = "You can participate in multimedia conferences via Internet through dial up modems."
StrQuestions(7) = "You do not have to pay for reading magazine through the net."
StrQuestions(8) = "Newsreader is a hardware implementation of the Usenet."
StrQuestions(9) = "E-mail can be used to send movies."
StrQuestions(10) = "Usenet and Mailing List facilities are the same."

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""
End Sub

Private Sub TrueorFalse_Paper12()
StrQuestions(1) = "Anonymous FTP allows viewing and retrieving a file from the archive of a host without having an account on that machine."
StrQuestions(2) = "Telnet is used for remote login."
StrQuestions(3) = "Gopher displays Internet resources in the hierarchical menu format."
StrQuestions(4) = "WWW is a large scale implementation of client server architecture."
StrQuestions(5) = "Gopher links a list of menus and hyperlinks to other menus.  "
StrQuestions(6) = "HTTP is used for creating home pages for the World Wide Web."
StrQuestions(7) = "Asynchronous communication is slower than synchronous communication."
StrQuestions(8) = "In full duplex mode of communication the channel is bi directional, whereas in half duplex mode it is unidirectional."
StrQuestions(9) = "Failure of a station in star network is easy to detect and the station can be removed without disrupting the network."
StrQuestions(10) = "Tree topology is an extension of bus topology."

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""
End Sub

Private Sub TrueorFalse_Paper13()
StrQuestions(1) = "Multiplexer allows one communication channel to send many transmissions at the same time."
StrQuestions(2) = "Failure of a stating in ring topology disrupts the network."
StrQuestions(3) = "CSMA/CD access method is used in star networks."
StrQuestions(4) = "Interfaces are rules of communication between similar modules, whereas protocols are rules of communication between dissimilar modules."
StrQuestions(5) = "Radio waves have very high data transfer rate and can be made secure through encryption."
StrQuestions(6) = "Hypertext is non linear text."
StrQuestions(7) = "The windows help system is an example of hypertext."
StrQuestions(8) = "A computer is not necessary to connect to the Net."
StrQuestions(9) = "HTTP is a non persistent, stateless protocol."
StrQuestions(10) = "The header of a document is removed before it is displayed in the browser."


stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""
End Sub

Private Sub TrueorFalse_Paper14()
StrQuestions(1) = "An Internet server may not be a part of the WWW."
StrQuestions(2) = "The Web is as old as the Net."
StrQuestions(3) = "The FTP sites are technically not part of the Web."
StrQuestions(4) = "E-mail is a synchronous form of communication."
StrQuestions(5) = "Text looks the same in every  browser, but graphics may look different."
StrQuestions(6) = "Tables are the easiest things to create with HTML programming."
StrQuestions(7) = "To make lines between cells on a web page disappear, use spacing value of 0."
StrQuestions(8) = "An image in JPEG format will take less time to load than in GIF format."

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
End Sub

Private Sub Matching_Columns_Paper1()
strcolumnA(1) = "Opera"
strcolumnA(2) = "CU-SeeMee"
strcolumnA(3) = "Java 2"
strcolumnA(4) = "Sambar"
strcolumnA(5) = "PkZip"
strcolumnA(6) = "Eudora"
strcolumnA(7) = "ICQ 99"
strcolumnA(8) = "Pegasus"
strcolumnA(9) = "Claris"
strcolumnA(10) = "Adobe SiteMill"

strcolumnB(1) = "Web designer"
strcolumnB(2) = "Online chat"
strcolumnB(3) = "Web server"
strcolumnB(4) = "E-mail"
strcolumnB(5) = "Video conferencing software"
strcolumnB(6) = "Web Browser"
strcolumnB(7) = "E-mail"
strcolumnB(8) = "Homepage builder"
strcolumnB(9) = "Development environment"
strcolumnB(10) = "Compression and decompression"

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""

End Sub

Private Sub Matching_Columns_Paper2()
strcolumnA(1) = "Adobe"
strcolumnA(2) = "Microsoft"
strcolumnA(3) = "Corel"
strcolumnA(4) = "McAfee"
strcolumnA(5) = "Netscape"
strcolumnA(6) = "Novvell"
strcolumnA(7) = "IBM"
strcolumnA(8) = "Symantec"
strcolumnA(9) = "Borland(Inprise)"
strcolumnA(10) = "Apple"

strcolumnB(1) = "QuickTime"
strcolumnB(2) = "Lotus Notes"
strcolumnB(3) = "System Utilities"
strcolumnB(4) = "Acrobat"
strcolumnB(5) = "LAN software"
strcolumnB(6) = "ActivveX"
strcolumnB(7) = "Webserver software"
strcolumnB(8) = "Office suite"
strcolumnB(9) = "virus scanner"
strcolumnB(10) = "Java tools"

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""

End Sub

Private Sub Matching_Columns_Paper3()
strcolumnA(1) = "UUCP"
strcolumnA(2) = "IP"
strcolumnA(3) = "CSLIP"
strcolumnA(4) = "PPP"
strcolumnA(5) = "TCP"
strcolumnA(6) = "TCP/IP"
strcolumnA(7) = "UDP"
strcolumnA(8) = "ARP"
strcolumnA(9) = "RARP"
strcolumnA(10) = "ICMP"

strcolumnB(1) = "Abstract hardwre interface"
strcolumnB(2) = "More features than SLIP"
strcolumnB(3) = "Provides reliability to IP"
strcolumnB(4) = "Internetworking"
strcolumnB(5) = "Compression of IP headers"
strcolumnB(6) = "Communicate error messages"
strcolumnB(7) = "Unix network"
strcolumnB(8) = "Send single packets to destination"
strcolumnB(9) = "Map IP addresses onto Ethernet addresses"
strcolumnB(10) = "Procedure for bootstrapping diskless clients"

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""

End Sub

Private Sub Matching_Columns_Paper4()
strcolumnA(1) = "Packet switched network"
strcolumnA(2) = "Network Information System"
strcolumnA(3) = "DNS"
strcolumnA(4) = "Bounced Mail"
strcolumnA(5) = "X.400"
strcolumnA(6) = "X.500"
strcolumnA(7) = "Bangpath Notation"
strcolumnA(8) = "Hybrid addresses"
strcolumnA(9) = "Elm"
strcolumnA(10) = "Address resolution"

strcolumnB(1) = "Broadcast datagram"
strcolumnB(2) = "Directory server"
strcolumnB(3) = "Returned to sender"
strcolumnB(4) = "Immediate packet forwarding"
strcolumnB(5) = "Giant distributed database"
strcolumnB(6) = "Mail addressing system"
strcolumnB(7) = "Mix different types of address operations"
strcolumnB(8) = "Electronic mail"
strcolumnB(9) = "Yellow pages"
strcolumnB(10) = "Uses exclamation mark"

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""

End Sub

Private Sub Matching_Columns_Paper5()
strcolumnA(1) = "Browse"
strcolumnA(2) = "Anchor"
strcolumnA(3) = "AOL"
strcolumnA(4) = "Cache"
strcolumnA(5) = "Encryption"
strcolumnA(6) = "Geocities"
strcolumnA(7) = "Java"
strcolumnA(8) = "Bookmarks"
strcolumnA(9) = "ActiveX"
strcolumnA(10) = "Guestbook"

strcolumnB(1) = "Favorites "
strcolumnB(2) = "Applet"
strcolumnB(3) = "Memory storage"
strcolumnB(4) = "Free internet sites"
strcolumnB(5) = "Interactive programming"
strcolumnB(6) = "Surfing"
strcolumnB(7) = "Hospitals"
strcolumnB(8) = "visitor comments"
strcolumnB(9) = "Security"
strcolumnB(10) = "ISP"

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""

End Sub

Private Sub Matching_Columns_Paper6()
strcolumnA(1) = "Virtual host"
strcolumnA(2) = "Ipvv6"
strcolumnA(3) = "DHCP"
strcolumnA(4) = "Gateway"
strcolumnA(5) = "Bridge"
strcolumnA(6) = "Hotspots"
strcolumnA(7) = "IP Routing"
strcolumnA(8) = "URL"
strcolumnA(9) = "Lynx"
strcolumnA(10) = "Datagram"

strcolumnB(1) = "Multiple network connections"
strcolumnB(2) = "Text mode browsers"
strcolumnB(3) = "New Internet protocol"
strcolumnB(4) = "More than one server on one machine"
strcolumnB(5) = "Mail packets"
strcolumnB(6) = "Intelligent network connector"
strcolumnB(7) = "Hyperlinds"
strcolumnB(8) = "Dynamically assignIP addresses"
strcolumnB(9) = "Network traffic"
strcolumnB(10) = "Unique page address"

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""

End Sub

Private Sub Matching_Columns_Paper7()
strcolumnA(1) = "FTP "
strcolumnA(2) = "VRML"
strcolumnA(3) = "Firewall"
strcolumnA(4) = "IP of Filtering Firewall"
strcolumnA(5) = "Proxy Server"
strcolumnA(6) = "Secure Socket Layer"
strcolumnA(7) = "Domain Name Server"
strcolumnA(8) = "Leased Line"
strcolumnA(9) = "MIME"
strcolumnA(10) = "Web Server"

strcolumnB(1) = "Net wide database"
strcolumnB(2) = "RSA data security"
strcolumnB(3) = "Make network connections for you"
strcolumnB(4) = "Multimedia mail"
strcolumnB(5) = "Intelligent Insulator"
strcolumnB(6) = "3D designs"
strcolumnB(7) = "File download"
strcolumnB(8) = "Store web pages"
strcolumnB(9) = "Block all but select network traffic"
strcolumnB(10) = "Fixed point to point data comunication"

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""

End Sub

Private Sub Matching_Columns_Paper8()
strcolumnA(1) = "Address with all host bits 0"
strcolumnA(2) = "Address with all host bits 1"
strcolumnA(3) = "Address 0.0.0.0"
strcolumnA(4) = "Address 127.0.0.0"
strcolumnA(5) = "Address 127.0.0.1"
strcolumnA(6) = ""
strcolumnA(7) = ""
strcolumnA(8) = ""
strcolumnA(9) = ""
strcolumnA(10) = ""

strcolumnB(1) = "Default Route"
strcolumnB(2) = "Loopback address"
strcolumnB(3) = "Loopback interface"
strcolumnB(4) = "Network"
strcolumnB(5) = "Broadcast address"
strcolumnB(6) = ""
strcolumnB(7) = ""
strcolumnB(8) = ""
strcolumnB(9) = ""
strcolumnB(10) = ""

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""

End Sub

Private Sub Matching_Columns_Paper9()
strcolumnA(1) = "RS-232"
strcolumnA(2) = "Thick Ehternet"
strcolumnA(3) = "Thinnet"
strcolumnA(4) = "Twisted pair cable"
strcolumnA(5) = "Push technology"
strcolumnA(6) = "Pull tchnology"
strcolumnA(7) = "Star network"
strcolumnA(8) = "Token Ring network"
strcolumnA(9) = "Server side includes"
strcolumnA(10) = "Ipv4"

strcolumnB(1) = "Maximum 232 Web addresses"
strcolumnB(2) = "10 base – T"
strcolumnB(3) = "10 base – 5"
strcolumnB(4) = "10 base – 2"
strcolumnB(5) = "Tight server control"
strcolumnB(6) = "Dynamic Web pages"
strcolumnB(7) = "Load on server"
strcolumnB(8) = "CSMA/CD"
strcolumnB(9) = "PC serial communication standard"
strcolumnB(10) = "No data collision"

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""

End Sub

Private Sub Matching_Columns_Paper10()
strcolumnA(1) = "IMAP"
strcolumnA(2) = "IP Masquerade"
strcolumnA(3) = "Anonymous FTP"
strcolumnA(4) = "Gopher"
strcolumnA(5) = "Bulletin Board"
strcolumnA(6) = "Pine"
strcolumnA(7) = "PageMill"
strcolumnA(8) = "Usenet"
strcolumnA(9) = "Network News Transfer Protocol"
strcolumnA(10) = "INN"

strcolumnB(1) = "E- mail program"
strcolumnB(2) = "Public domain area"
strcolumnB(3) = "Stream oriented connection"
strcolumnB(4) = "Internet News"
strcolumnB(5) = "Remote login"
strcolumnB(6) = "Newsgroups"
strcolumnB(7) = "Invisibly access Internet"
strcolumnB(8) = "Access e-mail without downloading "
strcolumnB(9) = "Menu driven"
strcolumnB(10) = "HTML designer"

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""

End Sub

Private Sub Matching_Columns_Paper11()
strcolumnA(1) = "Archie"
strcolumnA(2) = "Articles"
strcolumnA(3) = "Bandwidth"
strcolumnA(4) = "Bitnet"
strcolumnA(5) = "Bulletin Board Service"
strcolumnA(6) = "Emoticons"
strcolumnA(7) = "Encryption"
strcolumnA(8) = "Listserv"
strcolumnA(9) = "Mail Bombing"
strcolumnA(10) = "Request for comments"

strcolumnB(1) = "mailing Lists"
strcolumnB(2) = "Online documents that have to do with technical standards for the Internet."
strcolumnB(3) = "An autonomous network of research sitesData capacity of a channel"
strcolumnB(4) = "Net – based service to locate files for downloading"
strcolumnB(5) = "Combination of keyboard characters to express feelings in e – mail communication"
strcolumnB(6) = "Flooding an e-mail account/server"
strcolumnB(7) = "A forum for users to exchange information "
strcolumnB(8) = "Scrambling of a message for security"
strcolumnB(9) = "Messages posted to newsgroups"
strcolumnB(10) = ""

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""

End Sub

Private Sub Matching_Columns_Paper12()
strcolumnA(1) = "Acrobat"
strcolumnA(2) = "Adobe"
strcolumnA(3) = "Java"
strcolumnA(4) = "ASP"
strcolumnA(5) = "POTS"
strcolumnA(6) = ""
strcolumnA(7) = ""
strcolumnA(8) = ""
strcolumnA(9) = ""
strcolumnA(10) = ""

strcolumnB(1) = "Telephone service"
strcolumnB(2) = "Portable Document Format"
strcolumnB(3) = "Microsoft"
strcolumnB(4) = "Sun"
strcolumnB(5) = "Postscript"
strcolumnB(6) = ""
strcolumnB(7) = ""
strcolumnB(8) = ""
strcolumnB(9) = ""
strcolumnB(10) = ""

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""

End Sub

Private Sub Matching_Columns_Paper13()
strcolumnA(1) = "Java"
strcolumnA(2) = "MPEG 4"
strcolumnA(3) = "Direct X"
strcolumnA(4) = "MIT"
strcolumnA(5) = "NCP"
strcolumnA(6) = "MP-3"
strcolumnA(7) = "Ethernet"
strcolumnA(8) = "GUI"
strcolumnA(9) = "Portal"
strcolumnA(10) = "Thin client"

strcolumnB(1) = "Network Computer"
strcolumnB(2) = "Macintosh systems"
strcolumnB(3) = "CD quality music"
strcolumnB(4) = "One stop Web site"
strcolumnB(5) = "Palo Alto Research centre"
strcolumnB(6) = "Realistic game play"
strcolumnB(7) = "First protocol for the Internet"
strcolumnB(8) = "Interactive video"
strcolumnB(9) = "First WAN"
strcolumnB(10) = "Hardware independence"

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = ""
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = ""
stranswers(9) = ""
stranswers(10) = ""

End Sub

Private Sub Matching_Columns_Paper14()
strcolumnA(1) = "Ken Thompson & Dennis Ritchie"
strcolumnA(2) = "Vint Cerf & Bob kahn"
strcolumnA(3) = "Paul Zimmerman"
strcolumnA(4) = "Mark Andreesen"
strcolumnA(5) = "Sun Microsystems"
strcolumnA(6) = "Java"
strcolumnA(7) = "IBM"
strcolumnA(8) = "MPEG encoding"
strcolumnA(9) = "4GL"
strcolumnA(10) = "E-commerce"

strcolumnB(1) = "IBM"
strcolumnB(2) = "TCP"
strcolumnB(3) = "VGA standard"
strcolumnB(4) = "Business on the Web "
strcolumnB(5) = "Mosaic"
strcolumnB(6) = "SQL Queries"
strcolumnB(7) = "Pretty Good Privacy"
strcolumnB(8) = "Software to create animatins for Net viewing"
strcolumnB(9) = "James Gosling"
strcolumnB(10) = "Java"

stranswers(1) = ""
stranswers(2) = ""
stranswers(3) = ""
stranswers(4) = ""
stranswers(5) = "Java"
stranswers(6) = ""
stranswers(7) = ""
stranswers(8) = "Audio video compression"
stranswers(9) = ""
stranswers(10) = "Business on the Web"

End Sub

Private Sub multiplechoice_Paper1()
strmultiple(1) = "The following is not a part of header field in a mail message:"
strmuloption(1) = "Subject   "
strmuloption(2) = "From:"
strmuloption(3) = "Message-id"
strmuloption(4) = "None of above"

strmultiple(2) = "E-mail address consists of:"
strmuloption(5) = "Name of the machine handling the mail"
strmuloption(6) = "User identification"
strmuloption(7) = "Both"
strmuloption(8) = "None of the above"

strmultiple(3) = "What is a spider?"
strmuloption(9) = "a new internet game"
strmuloption(10) = "a new game for Mac enthusiasts"
strmuloption(11) = "an Internet search engine"
strmuloption(12) = "a spy program"

strmultiple(4) = "What do you need to put your web pages on the WWW?"
strmuloption(13) = "a connection to the Internet"
strmuloption(14) = "a Web browser"
strmuloption(15) = "a Web Server"
strmuloption(16) = "All of the above"

strmultiple(5) = "What are the new HTML features not supported by all browsers?"
strmuloption(17) = "tables"
strmuloption(18) = "forms"
strmuloption(19) = "frames"
strmuloption(20) = "all of the above"

strmultiple(6) = "The following is not a paragraph style in HTML."
strmuloption(21) = "Normal"
strmuloption(22) = "Preformatted"
strmuloption(23) = "Address"
strmuloption(24) = "All of the above"

strmultiple(7) = "The following is not a character style in HTML:"
strmuloption(25) = "emphasis"
strmuloption(26) = " variable"
strmuloption(27) = "code"
strmuloption(28) = "bold"

strmultiple(8) = "The following is not a type of list in HTML:"
strmuloption(29) = "Bulleted"
strmuloption(30) = "Glossary"
strmuloption(31) = "Menu"
strmuloption(32) = "None of the above"

strmultiple(9) = "Protocols in network mean"
strmuloption(33) = "Languages for communication between computers"
strmuloption(34) = "Rules for exchange of messages between two or more hosts"
strmuloption(35) = "Both "
strmuloption(36) = "None of the above"

strmultiple(10) = "UUCP stands for"
strmuloption(37) = "Unix to Unix Copy"
strmuloption(38) = "User to User communication Protocol"
strmuloption(39) = "Anonymous FTP"
strmuloption(40) = "Unknown User Control Protocol"


strmulanswer(1) = ""
strmulanswer(2) = ""
strmulanswer(3) = ""
strmulanswer(4) = ""
strmulanswer(5) = ""
strmulanswer(6) = ""
strmulanswer(7) = ""
strmulanswer(8) = ""
strmulanswer(9) = ""
strmulanswer(10) = ""
End Sub

Private Sub multiplechoice_Paper2()
strmultiple(1) = "Which of the following is true about UUCP?"
strmuloption(1) = "its main advantage is low band width"
strmuloption(2) = "it makes use of telephone equipment which place a tight limit on maximum transfer rate"
strmuloption(3) = "UUCP links are rarely permanent connections"
strmuloption(4) = "it is cheap but slow"

strmultiple(2) = "To make your computer a UUCP, node, you need "
strmuloption(5) = "a modem"
strmuloption(6) = "a working UUCP implementation"
strmuloption(7) = "another UUCP node that is willing to feed mail and news"
strmuloption(8) = "all of the above"

strmultiple(3) = "Usenet is mainly a provider of"
strmuloption(9) = "news "
strmuloption(10) = "mail"
strmuloption(11) = "utilities"
strmuloption(12) = "help services"

strmultiple(4) = "The protocol used by packet radio or ham radio is called"
strmuloption(13) = "X.25"
strmuloption(14) = "AX.25"
strmuloption(15) = "X.400"
strmuloption(16) = "X.500"

strmultiple(5) = "For the sites on the Internet, the IP addresses are assigned by:"
strmuloption(17) = "NIC"
strmuloption(18) = "IETF"
strmuloption(19) = "IEEE"
strmuloption(20) = "You are free to choose any 332- bit number"

strmultiple(6) = "UUCP does not need to know the following:"
strmuloption(21) = "the serial device the modem is attached to"
strmuloption(22) = "modem/communication speed"
strmuloption(23) = "a valid log-in ID"
strmuloption(24) = "chat script"

strmultiple(7) = "SMTP stands for"
strmuloption(25) = "Server Mail Transfer Protocol"
strmuloption(26) = "Simple Mail Transfer Protocol"
strmuloption(27) = "Small Mail Transmission Protocol"
strmuloption(28) = "Short Mail Transmission Protocol"

strmultiple(8) = "Mail routing involves.    Finding a path.   Error checking.    Speed optimizationd.    Cost optimization"
strmuloption(29) = "all of the above"
strmuloption(30) = " only 1 & 2"
strmuloption(31) = ".    only 1,2, & 3   "
strmuloption(32) = " only 1 & 3"

strmultiple(9) = "On the Internet the mailing route is"
strmuloption(33) = "specified by the user"
strmuloption(34) = "done by IP"
strmuloption(35) = "done by TCP"
strmuloption(36) = "all of the above"

strmultiple(10) = "Smart host routing is a technique used in"
strmuloption(37) = "UUCP"
strmuloption(38) = "WWW"
strmuloption(39) = "Dial up PPP"
strmuloption(40) = "Internet"


strmulanswer(1) = ""
strmulanswer(2) = ""
strmulanswer(3) = ""
strmulanswer(4) = ""
strmulanswer(5) = ""
strmulanswer(6) = ""
strmulanswer(7) = ""
strmulanswer(8) = ""
strmulanswer(9) = ""
strmulanswer(10) = ""
End Sub

Private Sub multiplechoice_Paper3()
strmultiple(1) = "The routing scheme inside the domain is"
strmuloption(1) = "completely visible to the outside worked"
strmuloption(2) = "regulated by IETF"
strmuloption(3) = "subject to Govt. media policies"
strmuloption(4) = "none of the above"

strmultiple(2) = "The following is true about mail aliases.  They provide shorthand for mail addresses.   They invoke a program with a mail message as the input. They send mail to the filed.    All systems require aliases for postmaster to be RFC compliant"
strmuloption(5) = "1 only"
strmuloption(6) = "1 & 4"
strmuloption(7) = "all of the above"
strmuloption(8) = "1,2 and 4"

strmultiple(3) = "The following is not true about NNTPs.  it stands for Network News Transfer Protocol.   it is a stream oriented client server connection.   it is a server side software packaged.  News transport is lower than UUCP networks"
strmuloption(9) = "1 & 2"
strmuloption(10) = "none"
strmuloption(11) = "3 & 4"
strmuloption(12) = "1,2 & 3"

strmultiple(4) = "The arrangement and meaning of Internet mail headers are defined by an Internet standard called:"
strmuloption(13) = "RFC 822"
strmuloption(14) = "RFC 1597"
strmuloption(15) = "RFC 920"
strmuloption(16) = "RFC 782"

strmultiple(5) = "Which of the following is not an example of common top level  domain:"
strmuloption(17) = "com"
strmuloption(18) = "org"
strmuloption(19) = "govt"
strmuloption(20) = "mil"

strmultiple(6) = "Which of the following statements is false? a. ISDN is a series of standards that specify a general purpose switched data network b..    An ISDN call creates a synchronous point to point data service to the destination c. ISDN is generally delivered in a high speed link that is broken down into a number of discrete channels i.    All true"
strmuloption(21) = "All true"
strmuloption(22) = "Only 3"
strmuloption(23) = "2 & 3"
strmuloption(24) = "1 & 3"

strmultiple(7) = "How long is IP address in IPv4?"
strmuloption(25) = "4 bytes"
strmuloption(26) = "8 bytes"
strmuloption(27) = "16 bytes"
strmuloption(28) = "6 bytes"

strmultiple(8) = "You may have a bridge if you have several(three or more) computers: a.   to save on price of a hub when you have an extra Ethernet card b.    to save on bother of learning how to do IP forwarding and other tricks c.    To avoid maintenance work in future when things change around"
strmuloption(29) = "all true"
strmuloption(30) = "1 & 2 true"
strmuloption(31) = "1 & 3 are true"
strmuloption(32) = "2 & 3 are true"

strmultiple(9) = "You want  a firewall if:  a.  You are trying to protect your network from external accessb.   You are tying to deny access to the world outside from your network c.   To protect your network from viruses"
strmuloption(33) = "1 & 2"
strmuloption(34) = "all"
strmuloption(35) = "3 only"
strmuloption(36) = "1 & 3"

strmultiple(10) = "The following is not a layer in the OSI reference model:"
strmuloption(37) = "Session layer"
strmuloption(38) = "Physical layer"
strmuloption(39) = "    Protocol layer"
strmuloption(40) = "Data link layer"

strmulanswer(1) = ""
strmulanswer(2) = ""
strmulanswer(3) = ""
strmulanswer(4) = ""
strmulanswer(5) = ""
strmulanswer(6) = ""
strmulanswer(7) = ""
strmulanswer(8) = ""
strmulanswer(9) = ""
strmulanswer(10) = ""
End Sub

Private Sub multiplechoice_Paper4()
strmultiple(1) = "Which of the following networking solution is suitable for networking in a building?"
strmuloption(1) = "LAN"
strmuloption(2) = "WAN"
strmuloption(3) = "MAN"
strmuloption(4) = "All of the above"

strmultiple(2) = "Data transfer rate for LAN is normally:"
strmuloption(5) = "1-10 mbps"
strmuloption(6) = "10-20 mbps"
strmuloption(7) = "1-2 Mbps"
strmuloption(8) = "1-100 Mbps"

strmultiple(3) = "Which of the following topology share a single channel on which all station can receive and transmit?"
strmuloption(9) = "bus"
strmuloption(10) = "tree"
strmuloption(11) = "star"
strmuloption(12) = "none of the above"

strmultiple(4) = "Which LAN access method used the right to transmission by a special bit pattern?"
strmuloption(13) = "CSMA/CD"
strmuloption(14) = "Ring topology"
strmuloption(15) = "Token passing"
strmuloption(16) = "None of the above"

strmultiple(5) = "Which of the following topology is least affected by addition/removal of a node?"
strmuloption(17) = "RING"
strmuloption(18) = "STAR"
strmuloption(19) = "BUS"
strmuloption(20) = "None of the above"

strmultiple(6) = "Protocols are used to communicate between"
strmuloption(21) = "Similar modules"
strmuloption(22) = "Dissimilar modules"
strmuloption(23) = "both"
strmuloption(24) = "none of the above"

strmultiple(7) = "Which of the following is used to connect two LANs using a same LAN protocols over a wide area?"
strmuloption(25) = "Router"
strmuloption(26) = "Bridges"
strmuloption(27) = "Gateways"
strmuloption(28) = "None of the above"

strmultiple(8) = "Which of the following services are not offered by CompuServe?"
strmuloption(29) = "Electronic mail box"
strmuloption(30) = "BBS"
strmuloption(31) = "News"
strmuloption(32) = "Information on computer hardware and software"

strmultiple(9) = "The following is not a domain name"
strmuloption(33) = "com  "
strmuloption(34) = "dev"
strmuloption(35) = "gov"
strmuloption(36) = "org"

strmultiple(10) = "The following is not a Windows e-mail program"
strmuloption(37) = "Outlook"
strmuloption(38) = "Pegasus"
strmuloption(39) = "Eudora"
strmuloption(40) = "Pine"

strmulanswer(1) = ""
strmulanswer(2) = ""
strmulanswer(3) = ""
strmulanswer(4) = ""
strmulanswer(5) = ""
strmulanswer(6) = ""
strmulanswer(7) = ""
strmulanswer(8) = ""
strmulanswer(9) = ""
strmulanswer(10) = ""
End Sub

Private Sub multiplechoice_Paper5()
strmultiple(1) = "Which of the following protocols is used by Internet mail?"
strmuloption(1) = "HTTP"
strmuloption(2) = "TCP/IP"
strmuloption(3) = "FTP"
strmuloption(4) = "PPP"

strmultiple(2) = "One of the following web sites does not contain a search engine:"
strmuloption(5) = "www.lycos.com"
strmuloption(6) = "www.yahoo.com"
strmuloption(7) = "www.excie.com"
strmuloption(8) = "none of the above"

strmultiple(3) = "The following is a synonym for the Internet"
strmuloption(9) = "cyberspace"
strmuloption(10) = "World Wide Web"
strmuloption(11) = "Gopher Space"
strmuloption(12) = "Intranet"

strmultiple(4) = "Where does the title of a web page appear when it is browsed?"
strmuloption(13) = "in search engines"
strmuloption(14) = "at top of page"
strmuloption(15) = "at top of browser window"
strmuloption(16) = "none of the above"

strmultiple(5) = "What is the most commonly used file name for the main page of a WWW site?"
strmuloption(17) = "default.html"
strmuloption(18) = "index.html"
strmuloption(19) = "start.html"
strmuloption(20) = "home.html"

strmultiple(6) = "What is another term for a bulleted list?"
strmuloption(21) = "Unnumbered list"
strmuloption(22) = "Numbered list"
strmuloption(23) = "Hit list"
strmuloption(24) = "Sequenced list"

strmultiple(7) = "What text is the placeholder for a bulleted list?"
strmuloption(25) = "No.  "
strmuloption(26) = "#"
strmuloption(27) = ">"
strmuloption(28) = "!"

strmultiple(8) = "Which of the following is not a feature of tables on web pages?"
strmuloption(29) = "every cell in a row has the same size"
strmuloption(30) = "rows are only a little bigger than the widest item in any of their cells"
strmuloption(31) = "every row has the same size"
strmuloption(32) = "none of the above"

strmultiple(9) = "What is the full form of ISDN?"
strmuloption(33) = "Innovations Subscribers Don't need"
strmuloption(34) = "I Still Don't know"
strmuloption(35) = "Integrated Subscriber Developer Network"
strmuloption(36) = "Integrated Services Digital Network"

strmultiple(10) = "Which of the following is true?"
strmuloption(37) = "Web server applications extend he functionality and capability of existing Web servers."
strmuloption(38) = "Web server applications receive HTTP request messages for the Web server and perform actions requested in those messages."
strmuloption(39) = "They formulate response and pass it back to the server."
strmuloption(40) = "All of the above"


strmulanswer(1) = ""
strmulanswer(2) = ""
strmulanswer(3) = ""
strmulanswer(4) = ""
strmulanswer(5) = ""
strmulanswer(6) = ""
strmulanswer(7) = ""
strmulanswer(8) = ""
strmulanswer(9) = ""
strmulanswer(10) = ""
End Sub

Private Sub multiplechoice_Paper6()
strmultiple(1) = "What do you call the documents found on the WWW?"
strmuloption(1) = "A link"
strmuloption(2) = "A page"
strmuloption(3) = "A node"
strmuloption(4) = "A spider"

strmultiple(2) = "The following is not a reason for using a text based browser such as Lynx:"
strmuloption(5) = "Pages load more quickly"
strmuloption(6) = "There is no possibility that objectionable images can be viewed."
strmuloption(7) = "Pages look better formatted with Lynx."
strmuloption(8) = "It is less taxing on the hardware resources."

strmultiple(3) = "Who created Mosaic, the first graphical web browser?"
strmuloption(9) = "Bill Gates"
strmuloption(10) = "Mark Anderson"
strmuloption(11) = " Steve Jobs."
strmuloption(12) = "Steve Ballmer"

strmultiple(4) = "If you don't put any links on your page, you have"
strmuloption(13) = "forced most visitors to use their browser's 'Back' button to leave the page."
strmuloption(14) = "created a dead end."
strmuloption(15) = "missed the opportunity to lead visitors to other parts of you Web site."
strmuloption(16) = "All of the above."

strmultiple(5) = "Which browser does not support the use of different fonts on page?"
strmuloption(17) = "Lynx"
strmuloption(18) = "Internet Explorer 3.0"
strmuloption(19) = "Netscape Navigator 4.0"
strmuloption(20) = "AOL"

strmultiple(6) = "Which of the link colors you are going to see the least often?"
strmuloption(21) = "Visited links."
strmuloption(22) = "Active links"
strmuloption(23) = "Normal links"
strmuloption(24) = "None of the above"

strmultiple(7) = "What is the term used to describe how background images can displayed repeatedly on a Web page?"
strmuloption(25) = "Cascading"
strmuloption(26) = "Tiling"
strmuloption(27) = "Centre"
strmuloption(28) = "None of the above"

strmultiple(8) = "The 'M' in HTML stands for"
strmuloption(29) = "Machine"
strmuloption(30) = "Modeling"
strmuloption(31) = "Markup"
strmuloption(32) = "Margin"

strmultiple(9) = "What part of a Web page should the <TITLE> tags be used in?"
strmuloption(33) = "The heading"
strmuloption(34) = "The body"
strmuloption(35) = "The footnote"
strmuloption(36) = "Anywhere"

strmultiple(10) = "The following is not a field in IP datagram:"
strmuloption(37) = "Version"
strmuloption(38) = "Header length"
strmuloption(39) = "Header checksum"
strmuloption(40) = "Source and destination IP address."


strmulanswer(1) = ""
strmulanswer(2) = ""
strmulanswer(3) = ""
strmulanswer(4) = ""
strmulanswer(5) = ""
strmulanswer(6) = ""
strmulanswer(7) = ""
strmulanswer(8) = ""
strmulanswer(9) = ""
strmulanswer(10) = ""
End Sub

Private Sub multiplechoice_Paper7()
strmultiple(1) = "Messages are carried between Internet hosts using ICMP, with a standard IP header. The following is not a typical such message:"
strmuloption(1) = "Destination unreachable"
strmuloption(2) = "Redirection"
strmuloption(3) = "Echo request/echo reply"
strmuloption(4) = "Source quench"

strmultiple(2) = "The following is not a User authentication method:"
strmuloption(5) = "Secure ID"
strmuloption(6) = "CHAP"
strmuloption(7) = "PAP"
strmuloption(8) = "None of the above"

strmultiple(3) = "Which of the following is not one of the used of repeaters?"
strmuloption(9) = "Boost signal power"
strmuloption(10) = "Clean signal pulses"
strmuloption(11) = "Increase maximum interconnection length"
strmuloption(12) = "None of the above"

strmultiple(4) = "Which of the following is not a way to refer to a Java program?"
strmuloption(13) = "An application"
strmuloption(14) = "A class"
strmuloption(15) = "An applet"
strmuloption(16) = "A script"

strmultiple(5) = "Which of the following statements is false?"
strmuloption(17) = "JPEG files cannot be animated like GIF files."
strmuloption(18) = "GIF files are commonly used for animation."
strmuloption(19) = "JPG files are easy for animation when they contain less than 256 colors."
strmuloption(20) = "None of the above"

strmultiple(6) = "What is the most common way animations are used on the Web?"
strmuloption(21) = "Menu choices"
strmuloption(22) = "Advertising banners"
strmuloption(23) = "Games"
strmuloption(24) = "All of the above"

strmultiple(7) = "Which company developed QuickTime movie format?"
strmuloption(25) = "Microsoft"
strmuloption(26) = "Apple"
strmuloption(27) = "Sun  "
strmuloption(28) = "Oracle"

strmultiple(8) = "Which of the following is not a good choice for a Web page many of commands?"
strmuloption(29) = "Java applet"
strmuloption(30) = "QuickTime"
strmuloption(31) = "Animated GIF file"
strmuloption(32) = "None of the above"

strmultiple(9) = "Which of the following is never used as meaning of PNG acronym?"
strmuloption(33) = "PNG's Not GIF"
strmuloption(34) = "Proctor 'N' Gamble"
strmuloption(35) = "Portable Network Graphics"
strmuloption(36) = "None of the above"

strmultiple(10) = "Which browser first popularized the use of plug-in-software?"
strmuloption(37) = "Internet Explorer"
strmuloption(38) = "Netscape Navigator"
strmuloption(39) = "NCSA Mosaic"
strmuloption(40) = "Chimera"

strmulanswer(1) = ""
strmulanswer(2) = ""
strmulanswer(3) = ""
strmulanswer(4) = ""
strmulanswer(5) = ""
strmulanswer(6) = ""
strmulanswer(7) = ""
strmulanswer(8) = ""
strmulanswer(9) = ""
strmulanswer(10) = ""
End Sub

Private Sub multiplechoice_Paper8()
strmultiple(1) = "What is another term for Web publishing?"
strmuloption(1) = "Downloading"
strmuloption(2) = "Uploading"
strmuloption(3) = "Freeloading"
strmuloption(4) = "Unloading"

strmultiple(2) = "What does FTP stand for?"
strmuloption(5) = "Faster Than Post"
strmuloption(6) = "File Transfer Protocol"
strmuloption(7) = "Files to Publish"
strmuloption(8) = "Format to Program"

strmultiple(3) = "Which of the following is not an e-mail client?"
strmuloption(9) = "MS Outlook"
strmuloption(10) = "Eudora Pro"
strmuloption(11) = "Pegasus"
strmuloption(12) = "Post office"

strmultiple(4) = "Which of the following is not a search tool for the Net "
strmuloption(13) = "Veronica"
strmuloption(14) = "Archie"
strmuloption(15) = "Excite"
strmuloption(16) = "Find fast"

strmultiple(5) = "Which of the following is not a part of TCP/IP suite of protocols?"
strmuloption(17) = "TCP"
strmuloption(18) = "UUCP"
strmuloption(19) = "UDP"
strmuloption(20) = "ICMP"

strmultiple(6) = "Which of the following technologies doesn't help in creating dynamic web pages?"
strmuloption(21) = "ASP"
strmuloption(22) = "PHP"
strmuloption(23) = "Perl"
strmuloption(24) = "Cold fusion"

strmultiple(7) = "Which of the following is not a client end scripting language?"
strmuloption(25) = "HTML"
strmuloption(26) = "XML"
strmuloption(27) = "XHTML"
strmuloption(28) = "None of the above"

strmultiple(8) = "What is MP-3?"
strmuloption(29) = "A sound player on Windows 9x platform"
strmuloption(30) = "New multimedia content on the web"
strmuloption(31) = "Format for compressed sound files"
strmuloption(32) = "Graphics presentation software"

strmultiple(9) = "Which of the following utilities provides encryption facility?"
strmuloption(33) = "Pretty Good Privacy"
strmuloption(34) = "Win-secure-It"
strmuloption(35) = "Both"
strmuloption(36) = "None"

strmultiple(10) = "Which of the following is a special protocol for e-mail?"
strmuloption(37) = "SMTP"
strmuloption(38) = "POP"
strmuloption(39) = "IMAP"
strmuloption(40) = "MIME"


strmulanswer(1) = ""
strmulanswer(2) = ""
strmulanswer(3) = ""
strmulanswer(4) = ""
strmulanswer(5) = ""
strmulanswer(6) = ""
strmulanswer(7) = ""
strmulanswer(8) = ""
strmulanswer(9) = ""
strmulanswer(10) = ""
End Sub

Private Sub multiplechoice_Paper9()
strmultiple(1) = "Which of the following browsers is different from others in most respects?"
strmuloption(1) = "Netscape Communicator"
strmuloption(2) = "Internet Explorer"
strmuloption(3) = "Opera"
strmuloption(4) = "Lynx"

strmultiple(2) = "Which of the following is file/data compression technique?"
strmuloption(5) = "Huffman Compression"
strmuloption(6) = "Shannon Fano coding"
strmuloption(7) = "Lempel Ziv-Welch"
strmuloption(8) = "All of the above"

strmultiple(3) = "Which of the following browsers in not free?"
strmuloption(9) = "Netscape navigator"
strmuloption(10) = "Internet explorer"
strmuloption(11) = "Opera"
strmuloption(12) = "Lynx"

strmultiple(4) = "Which of the following web sites provides free e-mail service?"
strmuloption(13) = "www.eudoramail.com"
strmuloption(14) = "www.hotmail.com"
strmuloption(15) = "www.netaddress.com"
strmuloption(16) = "None of the above"

strmultiple(5) = "Which among the following is a Web server application?"
strmuloption(17) = "Opera"
strmuloption(18) = "Claris"
strmuloption(19) = "Pegasus"
strmuloption(20) = "Sambler"

strmultiple(6) = "What is a shareware?"
strmuloption(21) = "Software available for use and distribution free of charge"
strmuloption(22) = "Software available for trial period after which it must be registered for use"
strmuloption(23) = "Software which is used in a multi-user/network system"
strmuloption(24) = "Software which is shared across a network"

strmultiple(7) = "What is MPEG related to?"
strmuloption(25) = "Method of audio-video compression"
strmuloption(26) = "Multimedia presentations"
strmuloption(27) = "Realistic game play"
strmuloption(28) = "E-commerce"

strmultiple(8) = "Which of the following is related to Internet security?"
strmuloption(29) = "CERT"
strmuloption(30) = "MIT"
strmuloption(31) = "DSL"
strmuloption(32) = "W3c"

strmultiple(9) = "What are emoticons?"
strmuloption(33) = "Combination of key board characters to express feelings in e-mail"
strmuloption(34) = "Emotional mail"
strmuloption(35) = "A multimedia mail attachment with pictures depicting extreme emotions"
strmuloption(36) = "None of the above"

strmultiple(10) = "What is encryption?"
strmuloption(37) = "Scrambling of message for security"
strmuloption(38) = "Invisibly access Internet"
strmuloption(39) = "Access e-mail without downloading"
strmuloption(40) = "Messages posted to newsgroups"

strmulanswer(1) = ""
strmulanswer(2) = ""
strmulanswer(3) = ""
strmulanswer(4) = ""
strmulanswer(5) = ""
strmulanswer(6) = ""
strmulanswer(7) = ""
strmulanswer(8) = ""
strmulanswer(9) = ""
strmulanswer(10) = ""
End Sub

Private Sub multiplechoice_Paper10()
strmultiple(1) = "What are messages posted to news groups called?"
strmuloption(1) = "Archie"
strmuloption(2) = "Articles"
strmuloption(3) = "Listserv"
strmuloption(4) = "Bitnet"

strmultiple(2) = "What is X.400 about?"
strmuloption(5) = "A directory server"
strmuloption(6) = "Procedure for bootstrapping diskless clients"
strmuloption(7) = "Mail addressing system"
strmuloption(8) = "Mapping of IP address"

strmultiple(3) = "What does Internet Control Message Protocol(ICMP) do?"
strmuloption(9) = "Map IP addresses onto Ethernet Addresses"
strmuloption(10) = "Communicate error messages"
strmuloption(11) = "Send single packets to destination"
strmuloption(12) = "compress IP headers"

strmultiple(4) = "What is bandwidth?"
strmuloption(13) = "The speed of data flow through modem"
strmuloption(14) = "the capacity of an ISDN connection"
strmuloption(15) = "Data capacity of a channel"
strmuloption(16) = "None of the above"

strmultiple(5) = "Which of the following is true about TCP?"
strmuloption(17) = "It is responsible for reliability of communication"
strmuloption(18) = "It allows several concurrent log-ins"
strmuloption(19) = "It helps in defining an abstract interface through which hardware is addressed"
strmuloption(20) = "All of the above"

strmultiple(6) = "Which of the following is not true about IP?"
strmuloption(21) = "It turns physically dissimilar networks into one apparently hmogenous one"
strmuloption(22) = "It is hardware independent"
strmuloption(23) = "It ,by intent, not reliable"
strmuloption(24) = "It is responsible for reliability of communications"

strmultiple(7) = "What is a Web site?"
strmuloption(25) = "A computer that stores web pages"
strmuloption(26) = "A program that allows surfers to view web pages"
strmuloption(27) = "A collection of web pages that contain related information"
strmuloption(28) = "A program that runs that runs on a web server to process forms"

strmultiple(8) = "What are CGI scripts?"
strmuloption(29) = "Browser add-in programs"
strmuloption(30) = "Programs that extend the capabilities of a server"
strmuloption(31) = "Both of the above"
strmuloption(32) = "None of the above"

strmultiple(9) = "What is SSL?"
strmuloption(33) = "The IP protocol to dynamically assign IP addresses"
strmuloption(34) = "It allows secure connection"
strmuloption(35) = "It is a derivative of the scripting language, PHP "
strmuloption(36) = "It is a web authoring tool"

strmultiple(10) = "What is a domain?"
strmuloption(37) = "It is a host that is connected to two networks"
strmuloption(38) = "It is a collection of sites that are related in some sense"
strmuloption(39) = "It is precursor to the WWW"
strmuloption(40) = "None of the above"


strmulanswer(1) = ""
strmulanswer(2) = ""
strmulanswer(3) = ""
strmulanswer(4) = ""
strmulanswer(5) = ""
strmulanswer(6) = ""
strmulanswer(7) = ""
strmulanswer(8) = ""
strmulanswer(9) = ""
strmulanswer(10) = ""
End Sub

Private Sub multiplechoice_Paper11()
strmultiple(1) = "What is Usenet?"
strmuloption(1) = "It is used to forward jobs and files through several hosts"
strmuloption(2) = "It is a low cost dialup network"
strmuloption(3) = "It is a provider of news"
strmuloption(4) = "None of the above"

strmultiple(2) = "Which of the following was a precursor to the Internet?"
strmuloption(5) = "Gopher"
strmuloption(6) = "ARPANET"
strmuloption(7) = "World Wide Web"
strmuloption(8) = "IETF"

strmultiple(3) = "What is a hot spot?"
strmuloption(9) = "A visual representation of a link in a node"
strmuloption(10) = "A network of nodes"
strmuloption(11) = "A hyperlink that can be multi directional"
strmuloption(12) = "A record of nodes that a user has accessed"

strmultiple(4) = "What is a leaf node?"
strmuloption(13) = "No further links"
strmuloption(14) = "A node with multi directional links"
strmuloption(15) = "A node full of hot spots"
strmuloption(16) = "None of the above"

strmulanswer(1) = ""
strmulanswer(2) = ""
strmulanswer(3) = ""
strmulanswer(4) = ""
strmulanswer(5) = ""
strmulanswer(6) = ""
strmulanswer(7) = ""
strmulanswer(8) = ""
strmulanswer(9) = ""
strmulanswer(10) = ""
End Sub




