VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGPapers"
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
 
Public Function CG(PaperType As String, CGPaper As Paper, mCat As Category, Number As Integer)
'ptype is the paper type
'CGpaper is the number of paper as paper1,paper2 ...
'mcat is the category
'number is the number of questions, or options or answers to be taken

 StrPaperType = PaperType
' StrPaper = Paper
If PaperType = "Fill in the Blanks" Then
   If CGPaper = Paper1 Then Fill_in_the_Blanks_Paper1
   If CGPaper = Paper2 Then Fill_in_the_Blanks_Paper2
   If CGPaper = Paper3 Then Fill_in_the_Blanks_Paper3
   If CGPaper = Paper4 Then Fill_in_the_Blanks_Paper4
   If CGPaper = Paper5 Then Fill_in_the_Blanks_Paper5
   If CGPaper = Paper6 Then Fill_in_the_Blanks_Paper6
   If CGPaper = Paper7 Then Fill_in_the_Blanks_Paper7
   If CGPaper = Paper8 Then Fill_in_the_Blanks_Paper8
   If CGPaper = Paper9 Then Fill_in_the_Blanks_Paper9
   If CGPaper = Paper10 Then Fill_in_the_Blanks_Paper10
   
ElseIf PaperType = "Matching Columns" Then
       If CGPaper = Paper1 Then Matching_Columns_Paper1
       If CGPaper = Paper2 Then Matching_Columns_Paper2
       If CGPaper = Paper3 Then Matching_Columns_Paper3
       If CGPaper = Paper4 Then Matching_Columns_Paper4
       If CGPaper = Paper5 Then Matching_Columns_Paper5
       If CGPaper = Paper6 Then Matching_Columns_Paper6
       If CGPaper = Paper7 Then Matching_Columns_Paper7
       If CGPaper = Paper8 Then Matching_Columns_Paper8
       If CGPaper = Paper9 Then Matching_Columns_Paper9
       
ElseIf PaperType = "True Or False" Then
       If CGPaper = Paper1 Then TrueorFalse_Paper1
       If CGPaper = Paper2 Then TrueorFalse_Paper2
       If CGPaper = Paper3 Then TrueorFalse_Paper3
       If CGPaper = Paper4 Then TrueorFalse_Paper4
       If CGPaper = Paper5 Then TrueorFalse_Paper5
       If CGPaper = Paper6 Then TrueorFalse_Paper6
       If CGPaper = Paper7 Then TrueorFalse_Paper7
       If CGPaper = Paper8 Then TrueorFalse_Paper8
       If CGPaper = Paper9 Then TrueorFalse_Paper9
       If CGPaper = Paper10 Then TrueorFalse_Paper10
       If CGPaper = Paper11 Then TrueorFalse_Paper11
       If CGPaper = Paper12 Then TrueorFalse_Paper12
       If CGPaper = Paper13 Then TrueorFalse_Paper13
       If CGPaper = Paper14 Then TrueorFalse_Paper14
       If CGPaper = Paper15 Then TrueorFalse_Paper15
       
ElseIf PaperType = "Descriptive Questions" Then
       If CGPaper = Paper1 Then questions_paper1
       If CGPaper = Paper2 Then questions_paper2
       If CGPaper = Paper3 Then questions_paper3
       If CGPaper = Paper4 Then questions_paper4
       If CGPaper = Paper5 Then questions_paper5
       If CGPaper = Paper6 Then questions_paper6
       If CGPaper = Paper7 Then questions_paper7
       
ElseIf PaperType = "Multiple Choice" Then
       If CGPaper = Paper1 Then multiplechoice_Paper1
       If CGPaper = Paper2 Then multiplechoice_Paper2
       If CGPaper = Paper3 Then multiplechoice_Paper3
       If CGPaper = Paper4 Then multiplechoice_Paper4
       If CGPaper = Paper5 Then multiplechoice_Paper5
       If CGPaper = Paper6 Then multiplechoice_Paper6
       If CGPaper = Paper7 Then multiplechoice_Paper7
       If CGPaper = Paper8 Then multiplechoice_Paper8
       If CGPaper = Paper9 Then multiplechoice_Paper9
       If CGPaper = Paper10 Then multiplechoice_Paper10
              
End If

If mCat = Fill_Questions Then
   CG = StrQuestions(Number)
ElseIf mCat = Fill_Options Then
       CG = stroptions(Number)
ElseIf mCat = Fill_Answers Then
       CG = stranswers(Number)
       
ElseIf mCat = Matching_ColumnA Then
       CG = strcolumnA(Number)
ElseIf mCat = Matching_ColumnB Then
       CG = strcolumnB(Number)
ElseIf mCat = Matching_Answers Then
       CG = stranswers(Number)
       
ElseIf mCat = Questions Then
       CG = StrQuestions(Number)
       
ElseIf mCat = TF_Questions Then
       CG = StrQuestions(Number)
ElseIf mCat = TF_Answers Then
       CG = stranswers(Number)
       
ElseIf mCat = MC_Questions Then
       CG = strmultiple(Number)
       
ElseIf mCat = MC_Options_1 Then
       CG = strmuloption(Number)
       
ElseIf mCat = MC_Options_2 Then
       CG = strmuloption(Number)
       
ElseIf mCat = MC_Options_3 Then
       CG = strmuloption(Number)
       
ElseIf mCat = MC_Options_4 Then
       CG = strmuloption(Number)
       
ElseIf mCat = MC_Answer Then
       CG = strmulanswer(Number)
End If


End Function

Private Sub Fill_in_the_Blanks_Paper1()
StrQuestions(1) = "Most general purpose graphics systems now include a keyboard and a dash as the major input devices."
StrQuestions(2) = "An RGB color system with dash bits of storage per pixel is generally referred to as a full color system."
StrQuestions(3) = "Jaggies or stair casing are an instance of a phenomenon known as dash. dash line drawing algorithm is very popular because it is accurate and it uses only incremental, integer calculations."
StrQuestions(4) = "dash line drawing algorithm is very popular because it is accurate and it uses only incremental, integer calculations."
StrQuestions(5) = "In dash pattern, characters are represented using rectangular grid patterns."
StrQuestions(6) = "Distance objects appear dash than objects closer to the viewing position under perspective projection."
StrQuestions(7) = "The number of principal vanishing points in a projection is determined by the umber of principal axes interesting the dash ."
StrQuestions(8) = "Cubic Bezier curves are generated with dash control points and the same number of blending functions."
StrQuestions(9) = "When two cubic Bezier curves, determined by control points {A,B,C,D} and {P,Q,R,S} are jointed with first order continuity, if we choose D = P and by positioning the second control point Q of the second curve. The position of this curve is at dash."
StrQuestions(10) = "When an object is being viewed on the X = Z plane from the point (0,0,100) it will give rise to dash principal vanishing points."

stroptions(1) = "Bresenham"
stroptions(2) = "Four"
stroptions(3) = "View plane"
stroptions(4) = "Smaller"
stroptions(5) = "Aliasing"
stroptions(6) = "Two"
stroptions(7) = "Mouse"
stroptions(8) = "24"
stroptions(9) = "Bitmap"
stroptions(10) = "D + (D-C)"

stranswers(1) = "Mouse"
stranswers(2) = "24"
stranswers(3) = "Aliasing"
stranswers(4) = "Bresenham"
stranswers(5) = "Bitmap"
stranswers(6) = "Smaller"
stranswers(7) = "View plane"
stranswers(8) = "Four"
stranswers(9) = "D + (D-C)"
stranswers(10) = "Two"

End Sub

Private Sub Fill_in_the_Blanks_Paper2()
StrQuestions(1) = "The term dash applies only to 1-bit per pixel bi-level systems."
StrQuestions(2) = "A display which stores the image as plotters and storage tubes do but allows selective erasing is the dash ."
StrQuestions(3) = "A dash  system CRT has three phosphor color dots at each pixel position."
StrQuestions(4) = "A dash  system produces jagged lines because lines are plotted as discrete point sets."
StrQuestions(5) = "Far away objects look smaller in size under dash projection."
StrQuestions(6) = "All the basic transformations can be expressed as matrices if we treat points in dash  coordinates."
StrQuestions(7) = "The parametric equation of a line passing through the points P and Q is dash "
StrQuestions(8) = "In GKS, the term dash  is used to identify various combinations of graphics hardware and software."
StrQuestions(9) = "dash are the properties of the output primitives."
StrQuestions(10) = "In the Cohen Sutherland line clipping algorithm if the result of the logical and operation of the region codes of the end points of line segment is not 0000 then the segment is dash"

stroptions(1) = "Attributes"
stroptions(2) = "Plasma panel"
stroptions(3) = "Shadow mask"
stroptions(4) = "Raster"
stroptions(5) = "Bitmap"
stroptions(6) = "tP + (1 – t) Q, t is a parameter"
stroptions(7) = "Homogenous"
stroptions(8) = "Perspective"
stroptions(9) = "Rejected"
stroptions(10) = "Workstation"

stranswers(1) = "Bitmap"
stranswers(2) = "Plasma panel"
stranswers(3) = "Shadow mask"
stranswers(4) = "Raster"
stranswers(5) = "Perspective"
stranswers(6) = "Homogenous"
stranswers(7) = "tP + (1 – t) Q, t is a parameter"
stranswers(8) = "Workstation"
stranswers(9) = "Attributes"
stranswers(10) = "Rejected"

End Sub

Private Sub Fill_in_the_Blanks_Paper3()
StrQuestions(1) = "Points in the screen are represented in dash coordinates."
StrQuestions(2) = "The basic building blocks for pictures are referred to as dash."
StrQuestions(3) = "One type of digitizer is the dash which is used to input 2D coordinates by activating a stylus at selected positions on a flat surface."
StrQuestions(4) = "An area on a display device to which a window is mapped is called a dash."
StrQuestions(5) = "The process of cutting off the lines which are outside the window so that only the lines within the window are displayed is known as dash."
StrQuestions(6) = "In a perspective projection, if the centre of projection is placed at infinity then it will reduce to dash projection."
StrQuestions(7) = "A parallel projection is called a/an dash projection of the direction of projection is perpendicular to the plane of projection."
StrQuestions(8) = "If the dash control point, of one Bezier curve is equal to the first control point of another Bezier curve, then the two Bezier curves are joined without any break."
StrQuestions(9) = "If we choose the knot vector to be {0,1,2,3,4,5,6} then the resulting B-spline blending functions will become dash."
StrQuestions(10) = "Generally a graphics system first converts world coordinate positions to dash device coordinates in the range from 0 to 1."

stroptions(1) = "Clipping"
stroptions(2) = "Output primitives"
stroptions(3) = "Viewport"
stroptions(4) = "Integer"
stroptions(5) = "Periodic"
stroptions(6) = "Orthographic"
stroptions(7) = "Parallel"
stroptions(8) = "Last"
stroptions(9) = "Normalized"
stroptions(10) = "Graphics Tablet"

stranswers(1) = "Integer"
stranswers(2) = "Output primitives"
stranswers(3) = "Graphics Tablet"
stranswers(4) = "Viewport"
stranswers(5) = "Clipping"
stranswers(6) = "Parallel"
stranswers(7) = "Orthographic"
stranswers(8) = "Last"
stranswers(9) = "Periodic"
stranswers(10) = "Normalized"

End Sub

Private Sub Fill_in_the_Blanks_Paper4()
StrQuestions(1) = "To have a steady picture on a refreshing display device, it has to be refreshed at least dash times per second."
StrQuestions(2) = "The portion of the memory used to hold the pixels is called the dash."
StrQuestions(3) = "Joystick is a dash."
StrQuestions(4) = "A tablet is composed of a flat surface and a dash."
StrQuestions(5) = "The display coordinate systems are referred to as  dash coordinates in the case of a video monitor."
StrQuestions(6) = "A dash consists of a small, vertical lever mounted on a base that is used to steer the screen cursor around."
StrQuestions(7) = "A dash keyboard is a small hand held box used to position the screen cursor."
StrQuestions(8) = "The process of converting the rasterized picture from the memory to the rigid display pattern of video is called dash conversion."
StrQuestions(9) = "Documents are scanned at dash dpi."
StrQuestions(10) = "dash is constructed by filling the region between two glass plates with a mixture of gases."

stroptions(1) = "Locator device"
stroptions(2) = "Scan"
stroptions(3) = "Screen"
stroptions(4) = "Joystick"
stroptions(5) = "Frame buffer"
stroptions(6) = "Stylus"
stroptions(7) = "Plasma panel"
stroptions(8) = "Keyboard"
stroptions(9) = "300"
stroptions(10) = "30"

stranswers(1) = "30"
stranswers(2) = "Frame buffer"
stranswers(3) = "Locator device"
stranswers(4) = "Stylus"
stranswers(5) = "Screen"
stranswers(6) = "Joystick"
stranswers(7) = "Keyboard"
stranswers(8) = "Scan"
stranswers(9) = "300"
stranswers(10) = "Plasma panel"

End Sub

Private Sub Fill_in_the_Blanks_Paper5()
StrQuestions(1) = "dash facility helps one to move back and forth in a document."
StrQuestions(2) = "A vanishing point is generated with a dash view."
StrQuestions(3) = "The transformation involved in dash is simply translation."
StrQuestions(4) = "To obtain the dash view of an object, it has to be rotated about the  X axis by 90 degrees."
StrQuestions(5) = "dash displays are used in calculators and laptop computers."
StrQuestions(6) = "A device used for drawing or painting on the screen is dash."
StrQuestions(7) = "A mouse provides dash changes in the cursor movement."
StrQuestions(8) = "dash methods stretch out a line from a starting position as the cursor is moved."
StrQuestions(9) = "In computer animation after the key frames are selected, the intermediate frames are called dash."
StrQuestions(10) = "Transformation of object shapes from one form to another is called dash."

stroptions(1) = "LCD"
stroptions(2) = "Top"
stroptions(3) = "Perspective"
stroptions(4) = "Graphics tablet"
stroptions(5) = "Rubber-band"
stroptions(6) = "Panning"
stroptions(7) = "Hypertext"
stroptions(8) = "In-betweens"
stroptions(9) = "Morphing"
stroptions(10) = "Relative"

stranswers(1) = "Hypertext"
stranswers(2) = "Perspective"
stranswers(3) = "Panning"
stranswers(4) = "Top"
stranswers(5) = "LCD"
stranswers(6) = "Graphics tablet"
stranswers(7) = "Relative"
stranswers(8) = "Rubber-band"
stranswers(9) = "In-betweens"
stranswers(10) = "Morphing"

End Sub

Private Sub Fill_in_the_Blanks_Paper6()
StrQuestions(1) = "In dash curve drawing technique control points may be placed randomly on the screen for the curve segments to join smoothly."
StrQuestions(2) = "Playing instrumental music in the background of text document on display is very easy with dash files."
StrQuestions(3) = "A dash is a picture stored as a set of pixels that correspond to the grid of dots on a computer screen."
StrQuestions(4) = "A dash refers to a text document that has been linked."
StrQuestions(5) = "The graphical representation of a surface defined by a function of the form F(x,y,z) = 0 is made possible through dash algorithm."
StrQuestions(6) = "dash curves start from a control point and end at a control point."
StrQuestions(7) = "By making use of dash in display monitors, it is possible to provide variety of color palettes."
StrQuestions(8) = "The dash holds the picture information which is displayed at regular intervals on the screen."
StrQuestions(9) = "A dash can not be used to input the location of point on the screen."
StrQuestions(10) = "The dash clipping algorithm works only for rectangular windows."

stroptions(1) = "Link text"
stroptions(2) = "MIDI"
stroptions(3) = "B-spline"
stroptions(4) = "Floating Horizon"
stroptions(5) = "Mouse"
stroptions(6) = "Bezier"
stroptions(7) = "Sutherland-Cohen"
stroptions(8) = "Bitmap"
stroptions(9) = "Lookup tables"
stroptions(10) = "AVI"

stranswers(1) = "B-spline"
stranswers(2) = "MIDI"
stranswers(3) = "Bitmap"
stranswers(4) = "Link text"
stranswers(5) = "Floating Horizon"
stranswers(6) = "Bezier"
stranswers(7) = "Lookup tables"
stranswers(8) = "AVI"
stranswers(9) = "Mouse"
stranswers(10) = "Sutherland-Cohen"

End Sub

Private Sub Fill_in_the_Blanks_Paper7()
StrQuestions(1) = "dash is a point plotting device."
StrQuestions(2) = "The convex polygon boundary that encloses a setoff control points is called the dash."
StrQuestions(3) = "dash is an operation by which a cursor can be moved to any position while the line/curve is drawn."
StrQuestions(4) = "The dash for character generation lends itself to changes of scale."
StrQuestions(5) = "dash operations are performed on an object to move the existing vertices, edges and faces."
StrQuestions(6) = "The dash parallel rays coming from infinity intersect the view plane with some angle."
StrQuestions(7) = "The process of determining the portion of a primitive lying within a region is called dash."
StrQuestions(8) = "The number of times image is redrawn on CRT is known as dash."
StrQuestions(9) = "A pixel’s value is used not to control the beam directly, but rather as an index into dash."
StrQuestions(10) = "One of the effects of aliasing is reflected in dash."

stroptions(1) = "Jagged edges"
stroptions(2) = "Clipping"
stroptions(3) = "Look up table"
stroptions(4) = "Oblique projection"
stroptions(5) = "Refresh rate"
stroptions(6) = "Tweaking"
stroptions(7) = "Stroke method"
stroptions(8) = "Raster CRT graphics display"
stroptions(9) = "Rubber banding"
stroptions(10) = "Convex hull"

stranswers(1) = "Raster CRT graphics display"
stranswers(2) = "Convex hull"
stranswers(3) = "Rubber banding"
stranswers(4) = "Stroke method"
stranswers(5) = "Tweaking"
stranswers(6) = "Oblique projection"
stranswers(7) = "Clipping"
stranswers(8) = "Refresh rate"
stranswers(9) = "Look up table"
stranswers(10) = "Jagged edges"
End Sub

Private Sub Fill_in_the_Blanks_Paper8()
StrQuestions(1) = "dash 44 kHz would represent high quality sound."
StrQuestions(2) = "The origin of the co ordinate system in a graphics screen is located at dash of the screen."
StrQuestions(3) = "The no. of pixels put 'ON' for a line moving from 2,8 to 8,12 would be dash."
StrQuestions(4) = "A cube is rotating on the screen; at ay point in time, the maximum number of visible faces would be dash."
StrQuestions(5) = "dash can be carried out with translation transformation."
StrQuestions(6) = "Length of the vector I+J is dash that of vector I-J."
StrQuestions(7) = "The most important parameter of scanner is its dash"
StrQuestions(8) = "dash means to look at any portion of the screen more closely."
StrQuestions(9) = "The voltage required to drive the picture tube of graphics monitor is dash the supply voltage."
StrQuestions(10) = "Mid point subdivision clipping algorithm requires dash calculations."

stroptions(1) = "16-bit"
stroptions(2) = "6"
stroptions(3) = "Panning"
stroptions(4) = "DPI"
stroptions(5) = "Greater Than"
stroptions(6) = "Top Left"
stroptions(7) = "3"
stroptions(8) = "Same as"
stroptions(9) = "Zooming"
stroptions(10) = "Integer"

stranswers(1) = "16-bit"
stranswers(2) = "Top Left"
stranswers(3) = "6"
stranswers(4) = "3"
stranswers(5) = "Panning"
stranswers(6) = "Same as"
stranswers(7) = "DPI"
stranswers(8) = "Zooming"
stranswers(9) = "Greater Than"
stranswers(10) = "Integer"

End Sub

Private Sub Fill_in_the_Blanks_Paper9()
StrQuestions(1) = "The dash interpolate the curve between the sampled points."
StrQuestions(2) = "The phenomenon of dash is associated with parallel projections only."
StrQuestions(3) = "Corresponding to world coordinate window, we have a rectangular region in screen coordinates known as dash."
StrQuestions(4) = "The Cohen – Sutherland line clipping is a dash approach."
StrQuestions(5) = "While joining any two curve segments, dash conditions have to be met."
StrQuestions(6) = "Both flood fill and boundary fill algorithms are sometimes called dash algorithms."
StrQuestions(7) = "The basic step in floating point horizon approach for hidden surface removal is dash."
StrQuestions(8) = "In raster scan display we have a requirement of memory called dash."
StrQuestions(9) = "The dash display is made up of 6 layers of crystalline material."
StrQuestions(10) = "The dash algorithm is used for generating line segments."

stroptions(1) = "Blending function"
stroptions(2) = "Foreshortening"
stroptions(3) = "View port"
stroptions(4) = "Non parametric"
stroptions(5) = "Continuity "
stroptions(6) = "Speed fill"
stroptions(7) = "Sorting"
stroptions(8) = "Frame buffer"
stroptions(9) = "LCD"
stroptions(10) = "DDA"

stranswers(1) = "Blending function"
stranswers(2) = "Foreshortening"
stranswers(3) = "View port"
stranswers(4) = "Non parametric"
stranswers(5) = "Continuity"
stranswers(6) = "Speed fill"
stranswers(7) = "Sorting"
stranswers(8) = "Frame buffer"
stranswers(9) = "LCD"
stranswers(10) = "DDA"

End Sub

Private Sub Fill_in_the_Blanks_Paper10()
StrQuestions(1) = "dash is a locator device."
StrQuestions(2) = "Equation of a line passing through origin and (2,1) is dash as that of a line passing through(4,2) and (8.4)."
StrQuestions(3) = "Screen coordinates must be in dash form."
StrQuestions(4) = "Vanishing points are associated with dash projection."
StrQuestions(5) = "The cross product of two vectors dash in the plane containing the vectors."
StrQuestions(6) = "Panning involves on dash transformations."
StrQuestions(7) = "Intersection of a line with a clipping window can result in dash value."
StrQuestions(8) = "Rotation about X axis followed b rotation about Y axis is dash as rotation about Y axis followed by rotation about X axis."
StrQuestions(9) = "The ratio of height to width of a world object remains same in dash projection."
StrQuestions(10) = "The dash line clipping algorithm can be easily implement in hardware."

stroptions(2) = "Isometric"
stroptions(3) = "Not same"
stroptions(6) = "Joystick"
stroptions(7) = "Midpoint subdivision"
stroptions(8) = "Does not lie"
stroptions(10) = "Translation"
stroptions(10) = "Integer"
stroptions(10) = "Real"
stroptions(10) = "Same"
stroptions(10) = "Perspective"
 
stranswers(1) = "Joystick"
stranswers(2) = "Same"
stranswers(3) = "Integer"
stranswers(4) = "Perspective"
stranswers(5) = "Does not lie"
stranswers(6) = "Translation"
stranswers(7) = "Real"
stranswers(8) = "Not same"
stranswers(9) = "Isometric"
stranswers(10) = "Midpoint subdivision"
End Sub

Private Sub Matching_Columns_Paper1()
strcolumnA(1) = "The conversion of continuous analog signal to discrete form is called"
strcolumnA(2) = "The number of samples per second is"
strcolumnA(3) = "The reciprocal of sampling rate or the no. of seconds per sample is"
strcolumnA(4) = "Analog input devices convert the user movement changes in either voltage or the amperage by varying the"
strcolumnA(5) = "Digital input devices that are actually analog deices that collect input information in"
strcolumnA(6) = "Strip microphones are used in this input deice"
strcolumnA(7) = "Most of the touch sensitive screens are designed using"
strcolumnA(8) = "Most common input device is"
strcolumnA(9) = "The difference between conditioned signal and the quantized signal is called"
strcolumnA(10) = "    The characters on the keyboard are encoded using"

strcolumnB(1) = "Aperture time"
strcolumnB(2) = "Resistance of the circuit"
strcolumnB(3) = "Analog to Digital conversion"
strcolumnB(4) = "Discrete form"
strcolumnB(5) = "Sampling rate"
strcolumnB(6) = "Light Emitting Diode (LED’s)"
strcolumnB(7) = "Amplitude uncertainty"
strcolumnB(8) = "Acoustic tablet"
strcolumnB(9) = "American Standard code for information interchange(ASCII)"
strcolumnB(10) = "Keyboard"

stranswers(1) = "Analog to Digital conversion"
stranswers(2) = "Sampling rate"
stranswers(3) = "Aperture time"
stranswers(4) = "Resistance of the circuit"
stranswers(5) = "Discrete form"
stranswers(6) = "Acoustic tablet"
stranswers(7) = "Light Emitting Diode (LED’s)"
stranswers(8) = "Keyboard"
stranswers(9) = "Amplitude uncertainty"
stranswers(10) = "American Standard code for information interchange(ASCII)"

End Sub
Private Sub Matching_Columns_Paper2()
strcolumnA(1) = "In a CRT the unit used to turn the electron beam On and OFF is"
strcolumnA(2) = "The duration of phosphorescence exhibited by a phosphor is called"
strcolumnA(3) = "The technique of splitting a raster scan pattern into two separate patterns is"
strcolumnA(4) = "The coordinate system used in computers to locate points is"
strcolumnA(5) = "The coordinate system we normally use to locate points"
strcolumnA(6) = "The path the electron when returning to the left side of the CRT is"
strcolumnA(7) = "The aspect ratio of 12 in X 16 in display"
strcolumnA(8) = "The coordinates after fitting(900,1000) to a 512 × 384 memory map would be"
strcolumnA(9) = "Plotter is similar to the display device"
strcolumnA(10) = "This display device can be viewed only from a limited angle"

strcolumnB(1) = "Left handed Cartesian coordinate system"
strcolumnB(2) = "225,250"
strcolumnB(3) = "Horizontal retrace"
strcolumnB(4) = "3:4"
strcolumnB(5) = "Interlacing"
strcolumnB(6) = "Memory tube"
strcolumnB(7) = "Control electrode"
strcolumnB(8) = "Persistence"
strcolumnB(9) = "Right handed Cartesian coordinate system"
strcolumnB(10) = "Liquid crystal display"

stranswers(1) = "Control electrode"
stranswers(2) = "Persistence"
stranswers(3) = "Interlacing"
stranswers(4) = "Left handed Cartesian coordinate system"
stranswers(5) = "Right handed Cartesian coordinate system"
stranswers(6) = "Horizontal retrace"
stranswers(7) = "3:4"
stranswers(8) = "225,250"
stranswers(9) = "Memory tube"
stranswers(10) = "Liquid crystal display"

End Sub
Private Sub Matching_Columns_Paper3()
strcolumnA(1) = "Computer graphics systems often require more than on processor because"
strcolumnA(2) = "A special area of memory dedicated to graphics only is called"
strcolumnA(3) = "The page of memory a processor can reach with direct addressing is called"
strcolumnA(4) = "In a good graphics, programming these operations are normally avoided"
strcolumnA(5) = "The technique of turning the electron beam off while retracing is called"
strcolumnA(6) = "A flicker free graphics device is"
strcolumnA(7) = "The process of representing continuous graphics objects as collection of discrete pixels is called"
strcolumnA(8) = "The binary values stored in the display memory"
strcolumnA(9) = "The phenomenon of emitting the light by a phosphor after it has need exposed to a high energy mea is called"
strcolumnA(10) = "The phenomenon of emitting the light by a phosphor while it is being exposed to a high energy mean is called"

strcolumnB(1) = "Storage tube display"
strcolumnB(2) = "Retrace blanking"
strcolumnB(3) = "Floating point calculation"
strcolumnB(4) = "Pixel on and pixel-of information"
strcolumnB(5) = "The calculation can overload the main processor"
strcolumnB(6) = "Fluorescence"
strcolumnB(7) = "The Zero Page"
strcolumnB(8) = "The frame buffer"
strcolumnB(9) = "Phosphorescence"
strcolumnB(10) = "Scan Conversion"

stranswers(1) = "The calculation can overload the main processor"
stranswers(2) = "The frame buffer"
stranswers(3) = "The Zero Page"
stranswers(4) = "Floating point calculation"
stranswers(5) = "Retrace blanking"
stranswers(6) = "Storage tube display"
stranswers(7) = "Scan Conversion"
stranswers(8) = "Pixel on and pixel-of information"
stranswers(9) = "Phosphorescence"
stranswers(10) = "Fluorescence"

End Sub
Private Sub Matching_Columns_Paper4()
strcolumnA(1) = "A graphics system in which the user dynamically controls the presentation of graphics models on a compute display is called"
strcolumnA(2) = "The scan conversion algorithm which uses integer addition, subtraction and multiplication by 2 is"
strcolumnA(3) = "Unequal intensity of diagonal lines is cone of the side effects of "
strcolumnA(4) = "The designer can compensate by"
strcolumnA(5) = "The initial pixel for starting the filling algorithm/.process is called"
strcolumnA(6) = "Dense matrices of dots in graphics is called"
strcolumnA(7) = "Data compression for run-length encoded pictures is about"
strcolumnA(8) = "The disadvantages with run length encoded pictures is"
strcolumnA(9) = "Active edge lists are used in"
strcolumnA(10) = "The alpha numeric CRT terminal uses this technique for real time operations of real time scam conversions"

strcolumnB(1) = "It is difficult to add or delete lines or text from the picture"
strcolumnB(2) = "Seed point"
strcolumnB(3) = "Cell encoding"
strcolumnB(4) = "Interactive graphics"
strcolumnB(5) = "10:1"
strcolumnB(6) = "Remaining consistent"
strcolumnB(7) = "Bresenham’s line algorithm"
strcolumnB(8) = "Pointillism"
strcolumnB(9) = "Scan conversion"
strcolumnB(10) = "Real time scan conversion"

stranswers(1) = "Interactive graphics"
stranswers(2) = "Bresenham’s line algorithm"
stranswers(3) = "Scan conversion"
stranswers(4) = "Remaining consistent"
stranswers(5) = "Seed point"
stranswers(6) = "Pointillism"
stranswers(7) = "10:1"
stranswers(8) = "It is difficult to add or delete lines or text from the picture"
stranswers(9) = "Real time scan conversion"
stranswers(10) = "Cell encoding"

End Sub
Private Sub Matching_Columns_Paper5()
strcolumnA(1) = "Over striking can be eliminated by"
strcolumnA(2) = "The geometric transformation in which the object is displaced a given distance and direction from its original position"
strcolumnA(3) = "The geometric transformation, in which the dimensions of an object are expanded or compressed"
strcolumnA(4) = "The mathematical statement of the view point in which the object itself is moved"
strcolumnA(5) = "The mathematical statement of the view point in which the object is held stationary and the background is moved"
strcolumnA(6) = "The  mathematical statement of he view point in which the object is held stationery and the background is moved "
strcolumnA(7) = "The transformation in which the transformation matrix is 1, a     b,a"
strcolumnA(8) = "he reflection about the line y = x can be attained by"
strcolumnA(9) = "The property that the adjacent pixels except at boundary edges, are likely to have the same characteristics"
strcolumnA(10) = "In a raster scan graphics device adjacent pixels on a scan line are likely to have the same characteristics. This property is"

strcolumnB(1) = "Rotation"
strcolumnB(2) = "Reversing the coordinates i.e., (x’,y’) =(y,x)"
strcolumnB(3) = "Spatial coherence"
strcolumnB(4) = "Scan line coherence"
strcolumnB(5) = "Checking each pixel before writing and avoid writing to pixel which was already written "
strcolumnB(6) = "Shearing"
strcolumnB(7) = "Translation"
strcolumnB(8) = "Geometric transformation"
strcolumnB(9) = "Scaling"
strcolumnB(10) = "Coordinate transformation"

stranswers(1) = "Checking each pixel before writing and avoid writing to pixel which was already written "
stranswers(2) = "Translation"
stranswers(3) = "Rotation"
stranswers(4) = "Scaling"
stranswers(5) = "Geometric transformation"
stranswers(6) = "Coordinate transformation"
stranswers(7) = "Shearing"
stranswers(8) = "Reversing the coordinates i.e., (x’,y’) =(y,x)"
stranswers(9) = "Spatial coherence"
stranswers(10) = "Scan line coherence"

End Sub
Private Sub Matching_Columns_Paper6()
strcolumnA(1) = "The clipping algorithm which uses the normal vector for determining whether a point on a line is inside or outside the window"
strcolumnA(2) = "the clipping algorithm which is more suitable for hardware implementation"
strcolumnA(3) = "In a perspective projection the eye of the artist is placed in"
strcolumnA(4) = "The points formed by apparent intersection  of lines parallel to the three principal axes X, Y, Z"
strcolumnA(5) = "The property that the farther an object is form the centre of projection the smaller it appears"
strcolumnA(6) = "Projections of lines that are not parallel to the view plane appear to meet at some point on the view plane. Such points are called"
strcolumnA(7) = "The property that objects behind the centre of projection are projected upside  down and backward onto the view plane is called"
strcolumnA(8) = "If the projection vector has the direction of the view plane normal then the projection is called"
strcolumnA(9) = "If the projection vector has the direction of the view plane not normal (arbitrary) then the projection is called"
strcolumnA(10) = "Orthographic projections in which the direction of projection is not parallel to any of the principal axes are called"

strcolumnB(1) = "Vanishing points"
strcolumnB(2) = "Principal vanishing points"
strcolumnB(3) = "Perspective foreshortening"
strcolumnB(4) = "View confusion"
strcolumnB(5) = "The centre of projection of the view plane"
strcolumnB(6) = "Oblique"
strcolumnB(7) = "Axonometric projections"
strcolumnB(8) = "Cyrus beck algorithm"
strcolumnB(9) = "Orthographic"
strcolumnB(10) = "Sutherland and Hodgeman"

stranswers(1) = "Cyrus beck algorithm"
stranswers(2) = "Sutherland and Hodgeman"
stranswers(3) = "The centre of projection of the view plane"
stranswers(4) = "Principal vanishing points"
stranswers(5) = "Perspective foreshortening"
stranswers(6) = "Vanishing points"
stranswers(7) = "View confusion"
stranswers(8) = "Orthographic"
stranswers(9) = "Oblique"
stranswers(10) = "Axonometric projections"

End Sub
Private Sub Matching_Columns_Paper7()
strcolumnA(1) = "The projection in which the direction of projection makes equal angles with all the three principal axes"
strcolumnA(2) = "The projection in which the direction of projection makes equal angles with exactly two of the three principal axes"
strcolumnA(3) = "The projection in which the direction of projection makes unequal angles with all the three principal axes"
strcolumnA(4) = "The projection in which the direction of projection is chosen so that there is no foreshortening of lines perpendicular to XY-plane"
strcolumnA(5) = "The projection in which the direction of projection is chosen so that the lines perpendicular to XY planes are foreshortened by half their lengths"
strcolumnA(6) = "The transformations applied to the object(data representing the object in world coordinates) prior o viewing"
strcolumnA(7) = "The projection plane used in three dimensional viewing of an object or picture"
strcolumnA(8) = "The 3D region I would coordinate space that will be clipped and projected onto the view plane"
strcolumnA(9) = "The transformations applied to the desired object to produce a two dimensional view of the object, described in normal device coordinates"
strcolumnA(10) = "Two dimensional transformations applied to the two dimensional image in normalized  device space to produce the final view"

strcolumnB(1) = "Cabinet"
strcolumnB(2) = "Image transformation"
strcolumnB(3) = "Diametric projection"
strcolumnB(4) = "Trimetric projections"
strcolumnB(5) = "Isometric projections"
strcolumnB(6) = "Cavalier projections"
strcolumnB(7) = "Viewing volume"
strcolumnB(8) = "Viewing transformations"
strcolumnB(9) = "Object transformations"
strcolumnB(10) = "View plane"

stranswers(1) = "Isometric projections"
stranswers(2) = "Diametric projection"
stranswers(3) = "Trimetric projections"
stranswers(4) = "Cavalier projections"
stranswers(5) = "Cabinet"
stranswers(6) = "Object transformations"
stranswers(7) = "View plane"
stranswers(8) = "Viewing volume"
stranswers(9) = "Viewing transformations"
stranswers(10) = "Image transformation"

End Sub
Private Sub Matching_Columns_Paper8()
strcolumnA(1) = "The method in which the clipping is done directly  against the view volume"
strcolumnA(2) = "Depth sort or priority algorithm is also called"
strcolumnA(3) = "One of the hidden surface algorithm to deal with non ;polygonal, no planar surface patches"
strcolumnA(4) = "The algorithms used to determine which objects and surfaces will obscure those objects and surfaces that are in back of them and rendering a more realistic image"
strcolumnA(5) = "It is easier to locate hidden surfaces when this projection is used."
strcolumnA(6) = "The method in which the clipping is done against the canonical view volume"
strcolumnA(7) = "The major difficulty in the implementation of Painter’s algorithm"
strcolumnA(8) = "The underlying concept of Painter’s algorithm"
strcolumnA(9) = "The property of raster displays makes the technique of ray tracing possible"
strcolumnA(10) = "this major advantage of ray tracing"

strcolumnB(1) = "Canonical clipping"
strcolumnB(2) = "It is independent of projection technique"
strcolumnB(3) = "Direct clipping"
strcolumnB(4) = "Sorting the polygons by depth and then paints each polygon onto the screen starting with the most distant polygon."
strcolumnB(5) = "Painter’s algorithm"
strcolumnB(6) = "The color of the pixel is set to the color of the object if there is an intersection of the projected ray  and object otherwise its color is set to the background."
strcolumnB(7) = "Hidden surface algorithm "
strcolumnB(8) = "Parallel projection"
strcolumnB(9) = "Ray tracing algorithm"
strcolumnB(10) = "The concept of depth of a polygon"

stranswers(1) = "Direct clipping"
stranswers(2) = "Painter’s algorithm"
stranswers(3) = "Ray tracing algorithm"
stranswers(4) = "Hidden surface algorithm "
stranswers(5) = "Parallel projection"
stranswers(6) = "Canonical clipping"
stranswers(7) = "The concept of depth of a polygon"
stranswers(8) = "Sorting the polygons by depth and then paints each polygon onto the screen starting with the most distant polygon."
stranswers(9) = "he color of the pixel is set to the color of the object if there is an intersection of the projected ray  and object otherwise its color is set to the background."
stranswers(10) = "It is independent of projection technique"

End Sub
Private Sub Matching_Columns_Paper9()
strcolumnA(1) = "The graphics standard which represents a series of consistently defined graphics procedures and a syntax for passing information to these procedures"
strcolumnA(2) = "The parameters required by POLYLINE"
strcolumnA(3) = "In GKS the number of coordinates required to draw a square"
strcolumnA(4) = "In GKS new primitives are built by"
strcolumnA(5) = "The principle used in Seed fill algorithms using stack"
strcolumnA(6) = "In back face removal any point(x1,y1,z1) specified in the right handed coordinate system is on the “inside” of the surface if the plane satisfies the inequality"
strcolumnA(7) = "The aliasing problem arises with"
strcolumnA(8) = "The algorithm used for generating line segments"
strcolumnA(9) = "The purpose of display processor"
strcolumnA(10) = "The main step in floating point horizon approach for hidden surface removal"

strcolumnB(1) = "Sorting"
strcolumnB(2) = "Five"
strcolumnB(3) = "Graphics Kernel System(GKS)"
strcolumnB(4) = "(Ax1+By1+Cz1+d)"
strcolumnB(5) = "Raster display"
strcolumnB(6) = "Digital differential analyzer (DDA)"
strcolumnB(7) = "A command data list comprising the basic primitives of GKS"
strcolumnB(8) = "To free CPU from the graphic chores"
strcolumnB(9) = "n,X,Y where X & Y are arrays of x and y coordinates and n is length of arrays"
strcolumnB(10) = "FILO"

stranswers(1) = "Graphics Kernel System(GKS)"
stranswers(2) = "n,X,Y where X & Y are arrays of x and y coordinates and n is length of arrays"
stranswers(3) = "Five"
stranswers(4) = "A command data list comprising the basic primitives of GKS"
stranswers(5) = "FILO"
stranswers(6) = "(Ax1+By1+Cz1+d)"
stranswers(7) = "Raster display"
stranswers(8) = "Digital differential analyze (DDA)"
stranswers(9) = "Digital differential analyzer (DDA)"
stranswers(10) = "Sorting"
End Sub

Private Sub Matching_Columns_Paper10()
strcolumnA(1) = "The convolution operation in initializing is also called as"
strcolumnA(2) = "The technique for using minimum number of intensity levels(generally black and white) to obtain increased visualization"
strcolumnA(3) = "The process of extracting a position of data base/a picture"
strcolumnA(4) = "The right hand coordinate system in which the display area of the virtual display device corresponds to the unit square whose lower left hand corner is at the origin "
strcolumnA(5) = "The coordinate system that corresponds to the device or workstation where the mage or model is to be displayed"
strcolumnA(6) = "A finite region to perform a viewing transformation"
strcolumnA(7) = "The mapping of window onto the display area of the device or onto a sub regin of the device"
strcolumnA(8) = "The aspect ratio of a rectangular window or view port is "
strcolumnA(9) = "The operation to display only a portion of a picture through a selected window"
strcolumnA(10) = "The reverse of clipping is -"

strcolumnB(1) = "Clipping"
strcolumnB(2) = "Normalized device coordinate system"
strcolumnB(3) = "window"
strcolumnB(4) = "Shielding"
strcolumnB(5) = "Half toning"
strcolumnB(6) = "Physical device coordinate system"
strcolumnB(7) = "Clipping        "
strcolumnB(8) = "Filtering"
strcolumnB(9) = "(s[max] - x[min]) / (y[max] - y [min])"
strcolumnB(10) = "view port"

stranswers(1) = "Filtering"
stranswers(2) = "Half toning"
stranswers(3) = "Clipping"
stranswers(4) = "Shielding"
stranswers(5) = "Half toning"
stranswers(6) = "window"
stranswers(7) = "view port"
stranswers(8) = "(s[max] - x[min]) / (y[max] - y [min])"
stranswers(9) = "Clipping"
stranswers(10) = "Shielding"

End Sub

Private Sub TrueorFalse_Paper1()
StrQuestions(1) = "The physical size of the pixel remains same on all the monitors"
StrQuestions(2) = "MIDI files occupy less space than WAV files."
StrQuestions(3) = "The use of graphics in video games is going down every day as faster processors are coming in the market."
StrQuestions(4) = "The resolution of a mouse must be compatible with that of the monitor."
StrQuestions(5) = "Any color or shade can be displayed o a VGA monitor."
StrQuestions(6) = "A laptop needs les electric power because it uses special graphics processor."
StrQuestions(7) = "A 300 dpi image scanner is good enough to input text material in the computer."
StrQuestions(8) = "A mirror reflection of a figure can be obtained by a combination of translation and scaling transformations."
StrQuestions(9) = "Straight lines sometimes appear broken (jagged) on the monitor as the screen is not really flat."
StrQuestions(10) = "The viewport must always be larger than the window."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper2()
StrQuestions(1) = "A cubic Bezier curve can be drawn with 4 control points C1,C2,C3,C4 only if C2 and C3 which have distance (separate) coordinate values."
StrQuestions(2) = "    If the control points of a cubic Bezier curve lie on a circle, then the curve would also be a circular arc."
StrQuestions(3) = "While generating the projection of 3D objects on the screen, if the distance of the viewer from the object is very large then the parallel and the perspective views may be more or less similar."
StrQuestions(4) = "The vanishing point in perspective projection appears on the screen only when the viewer is at a great distance from the object."
StrQuestions(5) = "The resolution of a screen dictates the number of colors that may be displayed."
StrQuestions(6) = "A cubic Bezier cure cannot be drawn if the control points are located at the vertices of a rectangle."
StrQuestions(7) = "A calculator or a laptop computer uses a very thin electronic gun for display."
StrQuestions(8) = "A shadow mask CRT has three phosphor color dots at each pixel position."
StrQuestions(9) = "A cubic Bezier curve may or may not lie within the polygon formed by the four control points."
StrQuestions(10) = "The Depth Buffer method is not applicable to arbitrarily shaped cured objects."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper3()
StrQuestions(1) = "The color variations, in a shadow mask CRT, is obtained by varying the intensity levels of the three electron beams."
StrQuestions(2) = "A passive matrix LCD display is better than an active matrix LCD display."
StrQuestions(3) = "Morphing is applicable only to static objects, and not for moving objects."
StrQuestions(4) = "To draw the mirror reflection of a Bezier curve segment, it is sufficient ot just take the reflection of the 4 control points instead of all the points on the curve."
StrQuestions(5) = "A color monitor using 24 bits per pixel in the frame buffer would produce a better quality picture than one using 8 bits per pixel."
StrQuestions(6) = "To store sound in a multimedia system, a mini tape recorder must be attached to the computer."
StrQuestions(7) = "The smoothness of the cubic Bezier curve decreases if the control points are placed very far from each other."
StrQuestions(8) = "A mirror reflection of an object about an axis parallel to X or Y axis can be obtained by using translation operation."
StrQuestions(9) = "The line drawing algorithm makes use of the line equation y = mmx+b. The algorithm is not useful in practical situations for lines which are nearly vertical as the value of m, the slope, could be around 50º or more."
StrQuestions(10) = "The Sutherland Cohen algorithm fails to clip a line if it is vertical and partly lying within the window."

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

Private Sub TrueorFalse_Paper4()
StrQuestions(1) = "Three control points are sufficient to draw a cubic curve."
StrQuestions(2) = "A clipping algorithm can be made more efficient if totally visible and totally invisible lines could be identified and separated from the remaining  lines."
StrQuestions(3) = "To obtain the from , top and side views of an object, it is very important to follow a particular sequence."
StrQuestions(4) = "The second and third control point of a cubic Bezier curve decides the starting slope of the curve."
StrQuestions(5) = "The depth Buffer method to detect visible surfaces needs a knowledge of the plane equation for each surface under consideration."
StrQuestions(6) = "In computer animation each frame in the sequence is called key frame."
StrQuestions(7) = "The purpose of the display processor is to free the CPU from the border of the graphics work."
StrQuestions(8) = "If the control points for Bezier curve are not placed properly the resulting curve may not be smooth in appearance."
StrQuestions(9) = "A TIFF picture file created on an IBM compatible PC can be downloaded on a Macintosh and viewed without any changes."
StrQuestions(10) = "A parametric cubic curve is always confined to the polygon form ed by the 4 control points used for drawing the curve."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"

End Sub

Private Sub TrueorFalse_Paper5()
StrQuestions(1) = "Light pens may give false readings due to background lighting in a room."
StrQuestions(2) = "Characters defined as bitmaps can be arbitrarily scaled and rotated"
StrQuestions(3) = "One can not produce a complex graphics image on a point plotting refreshing display device."
StrQuestions(4) = "A class of video devices that have reduced volume, weight and power requirements compared to a CRT is called flat panel display."
StrQuestions(5) = "The reflection along the line y = x is equivalent to the reflection along the X-axis followed by counter clockwise rotation of 90 degrees."
StrQuestions(6) = "The shape of a picture can be changed b applying translation transformation."
StrQuestions(7) = "(2,-3,6) is a homogenous coordinate representation of the point(1/3,-1/2,)."
StrQuestions(8) = "The orthographic projection of parallel lines converges to a vanishing point."
StrQuestions(9) = "Using back face removal method we can identify all the hidden surfaces of a single convex polyhedron."
StrQuestions(10) = "When all the control points of a Bezier cubic curve lie in a straight line then the curve also becomes a straight line."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "False"
stranswers(4) = "True"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper6()
StrQuestions(1) = "Dragging and rubber banding are techniques used n interactive graphics."
StrQuestions(2) = "Light pens are not popular because they are very expensive."
StrQuestions(3) = "Bresenham circle generation algorithm involves only integer additions, subtractions and multiplications."
StrQuestions(4) = "Applying a 2D rotation followed by scaling transformation is the same as applying first the scaling transformation and then the rotation."
StrQuestions(5) = "Clipping algorithm for non rectangular windows does not exit."
StrQuestions(6) = "Projections are classified into two categories: Orthographic and oblique projections."
StrQuestions(7) = "Cubic curves are the lowest degree polynomial curves that are non planar in 3D."
StrQuestions(8) = "Closed Bezier cubic curves are obtained by taking the first and the last control points to be the same."
StrQuestions(9) = "Back  face removal method is applied for curved surfaces."
StrQuestions(10) = "Floating horizon algorithm uses two one dimensional arrays: one to hold the minimum and the other the maximum projected y-values, for a finite set of projected X-values."

stranswers(1) = "True"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "True"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper7()
StrQuestions(1) = "Storing intensity information of each screen point in a raster system is the main reason for not being able to produce realistic display of scenes containing shading and color patterns."
StrQuestions(2) = "Joystick and light pen are both locator devices."
StrQuestions(3) = "The mouse reports relative movement while the tablet reports absolute position."
StrQuestions(4) = "The transformation, reflection about the origin, as the matrix representation: -1  0   0   -1"
StrQuestions(5) = "The reflection along the line y = -x is equivalent to reflection along the Y-axis followed by a clockwise rotation of 90 degrees."
StrQuestions(6) = "The Cohen Sutherland line clipping algorithm can be extended to non rectangular regions of clipping."
StrQuestions(7) = "In a perspective projection, points behind the eye are known as vanishing points."
StrQuestions(8) = "The polygonal face of a convex polyhedron having outward normal(1,1,1) will be self hidden if we view the face from the point(2,1,0)."
StrQuestions(9) = "The tangent to a Bezier curve at an end point is along the line joining that end point to the adjacent control point."
StrQuestions(10) = "The general form of the GKS function for specifying a sequence of n connected 2D straight line segments is: polyline (n,x,y)."

stranswers(1) = "False"
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

Private Sub TrueorFalse_Paper8()
StrQuestions(1) = "Liquid crystal displays are commonly used in small systems such as portable, laptop computers."
StrQuestions(2) = "Scan converting a straight line segment means locating the pixel positions closes to the line path and storing the intensity for each position in the frame buffer."
StrQuestions(3) = "Expressing positions ins homogenous coordinates allow us to represent the no. of seconds all geometric transformation as matrix multiplications."
StrQuestions(4) = "The Cohen Sutherland line clipping algorithm can not be extendable to 3D case."
StrQuestions(5) = "The overall design style for a setoff characters is called a typeface."
StrQuestions(6) = "When an object is being viewed on the x= 5 plane from the point(500,0,0) it will give rise to two principal; vanishing points."
StrQuestions(7) = "In the definition of a cubic Bezier curve the control points should be taken as distinct."
StrQuestions(8) = "In contrast to storage tube display a calligraphic refresh CRT displays uses a very high persistence phosphor."
StrQuestions(9) = "The valuator function provides multiple values."
StrQuestions(10) = "The light pen is the only true pick device."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper9()
StrQuestions(1) = "In back face surface removal, any point(x’,y’,z’) specified in the right handed coordinate system is on the “inside” of the surface if the plane satisfies the inequality Ax’,By’,Cz’,D<0"
StrQuestions(2) = "The knot vector is specified only for Bezier curves."
StrQuestions(3) = "GKS is a standardized method of developing graphics programs."
StrQuestions(4) = "Cohen Sutherland algorithm uses the end point codes of lines for clipping of lines in a window."
StrQuestions(5) = "Reflection about the line y = x is equivalent to a reflection relative to the Y axis followed by a counter clock wise rotation of 45 degrees."
StrQuestions(6) = "The scan line seed fill algorithm is not applicable to boundary defined regions."
StrQuestions(7) = "The aliasing problem arises only with raster display"
StrQuestions(8) = "For x=t × t, y = t , the transformation" & Chr(13) & _
                     "0  -2  2" & Chr(13) & _
                     "[x y 1]       -2   2   -2" & Chr(13) & _
                     "1  0   1" & Chr(13) & _
                     "yields points that like on a unit circle."
StrQuestions(9) = "Parallel projectors from infinity intersect the plane of projection at an oblique angle yielding oblique projection."
StrQuestions(10) = "Orthographic projection invariable involves vanishing points."

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

Private Sub TrueorFalse_Paper10()
StrQuestions(1) = "Parametric representation of circle of radius r is represented as x = r cos ? and y = r sin ? where 0 <= ? <= 2 ? (T)"
StrQuestions(2) = "B-spline curves and surfaces are the mostly used class of approximating splines."
StrQuestions(3) = "The 3D rotation about Z axis is same as the 2D rotation in homogeneous coordinates."
StrQuestions(4) = "The highest resolutions of image scanners that use laser light sources have resolutions greater than 2000 dpi."
StrQuestions(5) = "An uninterrupted span in scan line feed is a group of contiguous pixels on a single scan."
StrQuestions(6) = "Morphing is a graphics technique used to produce a transformation of one person or object into another,"
StrQuestions(7) = "Shadow mask methods are not commonly used in the raster scan method."
StrQuestions(8) = "The purpose of display processor is to free the CPU from the graphics chores"
StrQuestions(9) = "Data gloves can be used to grasp “real” objects."
StrQuestions(10) = "Touch panels allow displayed objects or screen positions to be selected with a touch of finger."

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

Private Sub TrueorFalse_Paper11()
StrQuestions(1) = "For a cube located at origin, it is possible to located a viewing point from which four sides are visible."
StrQuestions(2) = "Coordinates of a point ly9ng at infinity cannot be used with perspective matrix."
StrQuestions(3) = "In 3D graphics, in homogenous coordinate system, represents a point lying on X axis at infinity."
StrQuestions(4) = "Rotation of 30 degrees about Y axis is same as rotation of 30 degrees about Z axis and followed by clockwise rotation of 90 degrees about X axis."
StrQuestions(5) = "Uniform scaling in X and Y cannot distort a figure."
StrQuestions(6) = "Rotation about Y-axis will transfer a vector on Y axis to YZ plane"
StrQuestions(7) = "For a Bezier curve, as parameter t is changed in equal steps, the corresponding x coordinate also changes by equal amounts."
StrQuestions(8) = "Distinct control points are needed to draw two smoothly Bezier curves."
StrQuestions(9) = "The back face removal algorithm would fail to declare hidden surfaces when the viewing point is shifted to back side of an object."
StrQuestions(10) = "Floating Horizon algorithm can only handle increasing y values but not decreasing values."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper12()
StrQuestions(1) = "The ending slope of a Bezier curve can be determined without actually drawing the curve."
StrQuestions(2) = "The Bezier curve drawn with control point sequence T1, T2 T3 T4 would have an altogether different shape if the sequence is T1 T2 T3 T4 chosen."
StrQuestions(3) = "A video sequence is an arbitrary collection of image frames."
StrQuestions(4) = "In a multimedia computer usually audio and video information are stored in separate files."
StrQuestions(5) = "Audio captured in 133 MHz multimedia PC will sound the same in a 200 MHz multimedia PC."
StrQuestions(6) = "In a multimedia PC, while the text is stored in digital form, the pictures are stored in analog form."
StrQuestions(7) = "Sutherland Cohen line clipping algorithms can handle lines with negative slopes."
StrQuestions(8) = "In perspective projection all parallel lines of an object must appear to mee at a vanishing point."
StrQuestions(9) = "Given control points T1,T2,T3,T4 a Bezier curve can be drawn even if T1 T2 and T4 are collinear."
StrQuestions(10) = "For a B-spline curve, the order in which control points are elected, determines the shape of the curve."

stranswers(1) = "True"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "True"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper13()
StrQuestions(1) = "Bresenham’s circle drawing algorithm is implemented through showing equation of the circle."
StrQuestions(2) = "The dot product of two vectors 3I and 3J is zero."
StrQuestions(3) = "The most important parameter of a screen is its rate of scanning."
StrQuestions(4) = "A line moving from (5,5) to(10,8) would involve putting “ON” 3 pixels on the screen."
StrQuestions(5) = "The last column of any transformation matrix, used in computer graphics, can consist of all zero elements."
StrQuestions(6) = "Panning achieved through transformation matrix, used in computer graphics, can consist of all zero elements."
StrQuestions(7) = "The co ordinate of a point(100, -30) as reflected about the y-axis would be (-100,-30)."
StrQuestions(8) = "A Bezier curve is drawn with control points T1,T2,T3,T4 now T2 is moved slightly up on the line joining T1 and T2. This could not cause any change in shape of the curve."
StrQuestions(9) = "A graphics command has been expected to draw a circle on the screen. The process does not do anything till the next command is given by the user."
StrQuestions(10) = "The voltage required to drive the picture tube of a graphics motion is much greater than that required for an ordering motion."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "False"
stranswers(4) = "False"
stranswers(5) = "False"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub TrueorFalse_Paper14()
StrQuestions(1) = "An ink jet printer is used once or twice in six months. It can be used very effectively in this manner for a number of years."
StrQuestions(2) = "The operation of a multimedia PC can be said to be identical to that of a VCR as both deal with sound and picture and any program can be viewed  on demand."
StrQuestions(3) = "The cross product of two vector 2I and 2J would lie along the Z-axis."
StrQuestions(4) = "It is possible that for a Bezier curve the polynomials for x(t) and y(t) are identical."
StrQuestions(5) = "Each term of a vector 2I + 6J + 3K is multiplied by 3. The length of the vector would be 3 times the original length."
StrQuestions(6) = "The length of a vector I + J + K would be grater than that of vector I + J + K."
StrQuestions(7) = "If end point codes of lines are 0000 and 0000 for he Sutherland Cohen algorithm, then no intersection need to be carried out as the line would be completely visible."
StrQuestions(8) = "By varying t in equal interval along a Bezier curve the generated points would be at equal interval from each other "
StrQuestions(9) = "For a Bezier cure t = 0.5 yields a point which is equidistant from the end points."
StrQuestions(10) = "A 600 dpi scanner would pick up 600 pixels per inch from a document."

stranswers(1) = "False"
stranswers(2) = "False"
stranswers(3) = "True"
stranswers(4) = "False"
stranswers(5) = "True"
stranswers(6) = "False"
stranswers(7) = "True"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "True"
End Sub

Private Sub TrueorFalse_Paper15()
StrQuestions(1) = "A CD-ROM contains video games. When the game is played on a PC, the score of the player would be stored o the CD ROM itself."
StrQuestions(2) = "The file extension for Windows Bitmap is .BMP."
StrQuestions(3) = "MIDI requires less processing speed that digital audio files in multimedia application."
StrQuestions(4) = "It is possible to store a hand written document in a computer."
StrQuestions(5) = "Text is a fundamental building block of multimedia application."
StrQuestions(6) = "Higher the image resolution, the more disc space it requires the more time it takes to the displayed, and the longer it takes to print."
StrQuestions(7) = "The resolution of a captured image can be changed any number  of times without affecting the surfaces of original image."
StrQuestions(8) = "8MB RAM is most comfortable for develop most multimedia applications using standard packages."
StrQuestions(9) = "MIDI system record actual sound in digital file."
StrQuestions(10) = "One should liberally use high quality sounds in the digital file."

stranswers(1) = "False"
stranswers(2) = "True"
stranswers(3) = "True"
stranswers(4) = "True"
stranswers(5) = "False"
stranswers(6) = "True"
stranswers(7) = "False"
stranswers(8) = "False"
stranswers(9) = "False"
stranswers(10) = "False"
End Sub

Private Sub questions_paper1()
StrQuestions(1) = "a)  Give the transformation for rotating a 2D figure about an arbitrary point by 60 degrees clockwise." & Chr(13) & _
                  "b)  What are the uses of isometric projections?" & Chr(13) & _
                  "c)  How many control points are needed to generate a bicubic Bezier surface patch? Explain why?" & Chr(13) & _
                  "d)  Describe salient features of GKS."
StrQuestions(2) = "a)  A B-Spline curve is generated with the points A,B,C,D,E,F,GH. How would the shape of the curve change if the points are specified as A,B,C,D,E,F,G? Why?" & Chr(13) & _
                  "b)  briefly explain working of a tablet." & Chr(13) & _
                  "c)  How are top and side views generated on the screen?"
StrQuestions(3) = "a.  Given control points(10,100), (50,100),(70,120) and(100,150) calculate coordinates of any 4 points lying on the corresponding Bezier curve." & Chr(13) & _
                  "b.  Which clipping algorithm is best suited for hardware implementation? Give how this algorithm works?"
StrQuestions(4) = "a)  Consider a square with diagonal vertices(10,10) and (90,90). Draw a rough sketch of the cubic B-spline curve generated assuming that all the vertices of the square and the mid points of the sides form a set of consecutive control points." & Chr(13) & _
                  "b)  How are periodic B-Spline curves different from non-periodic B-spline curves?"
StrQuestions(5) = "a)  Indicate the transformation matrices top, side and front views of an object on the XY Plane (Y-vertical, Z toward the viewer)" & Chr(13) & _
                  "b)  Derive the transformation matrix for rotating an object about a line passing through a point (a,b,c) and the origin."
StrQuestions(6) = "a)  show that for a Beize curve of degree n, the ith derivative at the starting point depends on the first i+1 points where I <n." & Chr(13) & _
                  "b)  Discuss the continuity properties of 3D Bezier curve and surface patches. What are the restrictions on placement of control points for smooth variations?"
StrQuestions(7) = "What are the various input devices used for Graphics and explain briefly how they operate."
StrQuestions(8) = "a)  Distinguish between CGA, Ga and SVGA resolutions." & Chr(13) & _
                  "b)  What are plasma panels and how do they work?" & Chr(13) & _
                  "c)  Describe briefly the calligraphic refresh display systems." & Chr(13) & _
                  "d)  How can the light pen differentiate between two points on "
StrQuestions(9) = "Develop a transformation matrix to obtain isometric projections of pyramid of height 10 units and having 6 units size of square base. Make necessary assumptions if required."
StrQuestions(10) = ""

End Sub

Private Sub questions_paper2()
StrQuestions(1) = "a.  Explain the Sutherland Cohen line clipping algorithm. Is this applicable to any type of window? Justify your answer." & Chr(13) & _
                  "b.  Define vanishing points. Is the location of vanishing point directly related to the viewing point? Explain how."
StrQuestions(2) = "a.  Describe the Z-buffer algorithm for hidden surface removal." & Chr(13) & _
                  "b.  A mirror is vertically placed such that it passes through (20,0) and (0,20) find the reflected view of a triangle with vertices (30,40), (50,)) and (40,70) in this mirror."
StrQuestions(3) = "a.  Distinguish between windowport and view port." & Chr(13) & _
                  "b.  In 2D clipping how are lines grouped into visible, invisible and partially visible categories?" & Chr(13) & _
                  "c.  What is an oblique view of an object?"
StrQuestions(4) = "a.  Explain Scan line Seed Fill Algorithm." & Chr(13) & _
                  "b.  Describe briefly Bresenham's circle drawing algorithm."
StrQuestions(5) = "a.  describe the floating horizon method of removing hidden lines while displaying the surfaces." & Chr(13) & _
                  "b.  Work out the transformation matrix for generating a two vanishing point perspective projection for a cuboids. Show how the location of vanishing points on the screen can be computed."
StrQuestions(6) = "A Cube is placed at the origin of 3D system such that all its vertices have positive coordinate values and sides are parallel to the three principal axes. Indicate a convenient position of a viewer at which hw can see a 2-point perspective projection. Verify that such a view is generated."
StrQuestions(7) = "a.  Discuss the merits and demerits of various clipping algorithms." & Chr(13) & _
                  "b.  Explain briefly Explicit line clipping algorithm."
StrQuestions(8) = "a.  Determine the transformation matrix for carrying out rotation about an arbitrary axis." & Chr(13) & _
                  "b.  Develop a transformation matrix to obtain isometric view of cuboid of size 100 × 70 × 50 lying at the origin."
StrQuestions(9) = "a.  Explain how a CD-ROM . Drive reads contents from a CD." & Chr(13) & _
                  "b.  Enumerate various applications of Multimedia in social life."
StrQuestions(10) = "Write short notes on the following:" & Chr(13) & _
                   "a.  Hidden Surface removal" & Chr(13) & _
                   "b.Hypertext" & Chr(13) & _
                   "c.  Creation of MIDI files and generation of sounds." & Chr(13) & _
                   "D.Tweaking "

End Sub

Private Sub questions_paper3()
StrQuestions(1) = "a.  Discuss merits and demerits of various techniques for hidden surface removal." & Chr(13) & _
                   "b.  Describe briefly Phong Shading"
StrQuestions(2) = "Write short notes on the following" & Chr(13) & _
                  "a.  Applications of Animation" & Chr(13) & _
                  "b.  Use of Homogenous coordinate system" & Chr(13) & _
                  "c.  Dragging: Rubber band methods"
StrQuestions(3) = "Write short notes on the following" & Chr(13) & _
                  "a.Character generation" & Chr(13) & _
                  "b.  Floating horizon technique of Hidden surface removal." & Chr(13) & _
                  "c.  Describe what is Morphing in Animation"
StrQuestions(4) = "a.  Explain briefly about computing the Vanishing point." & Chr(13) & _
                  "b.  What is video conferencing and what are the issues involved in setting up such a facility?"
StrQuestions(5) = "Design an algorithm for a light pen tracking a cross"
StrQuestions(6) = "    A polygon is to be drawn for the following vertices (5,9),(25,7),(25,3),(75,3),(75,7). Give a window and viewport specification which would allow the use ofthese coordinates in creating the polygon."
StrQuestions(7) = "    What is the maximum number additional vertices which can be created when -" & Chr(13) & _
                  "a.  An n sided convex polygon is clipped by a line" & Chr(13) & _
                  "b.  An n sided convex polygon is clipped by a square" & Chr(13) & _
                  "Justify your answer"
StrQuestions(8) = "Write a 2 × 2 transformation matrix for each of the following:" & Chr(13) & _
                  "a.  Scale the entire picture four times as large" & Chr(13) & _
                  "b.  Scale the entire picture one third as large"
StrQuestions(9) = "Give a 3 × 3 homogeneous coordinate transform matrix to -" & Chr(13) & _
                   "a.  rotate the image clockwise by 90 degrees" & Chr(13) & _
                   "b.  Shift the image to the right by 10 units" & Chr(13) & _
                   "c.  Scale the image by twice as large"
StrQuestions(10) = ""

End Sub

Private Sub questions_paper4()
StrQuestions(1) = "a.  Cubic polynomial curves are more preferred in Graphics rather than any other polynomial curve. Comment and justify your answer" & Chr(13) & _
                  "b.  Outline the algorithmic steps for extending cohen Sutherland line clipping algorithm to three dimensional case."
StrQuestions(2) = "Write short notes on the following" & Chr(13) & _
                  "a.  Rubber band technique" & Chr(13) & _
                  "b.  Liquid Crystal Display" & Chr(13) & _
                  "c.  Homogenous coordinate system"
StrQuestions(3) = "a.  Outline the Bresenham circle generation algorithm" & Chr(13) & _
                  "b.  Explain the back face removal method"
StrQuestions(4) = "a.  Explain what is shadow mask phenomenon" & Chr(13) & _
                  "b.  What are orthographic projections? When do we need them?"
StrQuestions(5) = "a.  If we interchange the sequence of three geometric transformations, Translation, Rotation and Scaling. What is the effect on the resultant image? Justify your answer" & Chr(13) & _
                  "b.  Explain the Floating horizon algorithm."
StrQuestions(6) = "Draw two concentric circles and show how o fill in the area between the two circles."
StrQuestions(7) = "Draw ellipses on the screen such that all the major axes pass through origin and angle between any two major axes is 45 degrees."
StrQuestions(8) = "What are the common formats for storing audio and video in a multimedia systems? What do you understand by the term 'Hypertext'?"
StrQuestions(9) = "In 2D graphics obtain the 3 × 3 transformation matrix for translating ap oit by -1,2. calculate the inverse of this matrix and show that the result is a matrix which translates a point y by 1, -2?"
StrQuestions(10) = ""

End Sub

Private Sub questions_paper5()
StrQuestions(1) = "In 2D graphics the following transformations matrix would reflect a point about the diagonal line passing through the origin a and (10,10)." & Chr(13) & _
                  "[x y 1] = [ x y 1]          0   1   0" & Chr(13) & _
                  "1   0   0" & Chr(13) & _
                  "0   0   1" & Chr(13) & _
                  "Show that this is same as the coordination of matrix for 45 degrees clockwiserotation, followed by reflection about the x-zxis and finally by rotation of 45 degrees anticlockwise rotation (about the origin)"
StrQuestions(2) = "A unit cube is located at origin such that 3 of its edges are parallel to x,y and z axes. The vertices on from face have the coordinates." & Chr(13) & _
                  "The cube is being viewed from (0,0,5). Obtain the perspective projection points for the front and back faces of he cube. Can you indicate what happens to edges parallel to X axis, Y- axis and Z-Axis. Do they all meet at some vanishing point?"
StrQuestions(3) = "For the previous problem on cube located at the origin, imagine a person located at (10,6,8) looking at the centre of the cube. Find out which faces of thecube would be is able to him.(Hint: use Backface Removal Algorithm)."
StrQuestions(4) = "For the previous problem on cube located at the origin, find the oblique viewed(parallel projection) on XY plane, with light rays falling along the line(1,0,0) - (9,0,6)."
StrQuestions(5) = "A triangle PQR has its vertices located at P(80,50), Q(60,10), R(100,10). It is desired to obtain its reflection about an axis parallel to Y-axis and passing through the point A(30,10). Work out the necessary transformation matrix and also the coordinates of the vertices of the reflected triangle."
StrQuestions(6) = "Write a platform segment or give the algorithm for drawing a pie chart on the screen."
StrQuestions(7) = "A box is to be shown on the screen with 3 faces being visible. Assume a suitable position for the viewer for this purpose and indicate a method to find out which of 3 faces (out 0f 6) would be visible to the viewer."
StrQuestions(8) = "List the various options available in PAINTBRUSH under WINDWS."
StrQuestions(9) = "A clipping window ABCD is located as follows:" & Chr(13) & _
                  "A(100,10)       B(160,10)" & Chr(13) & _
                  "C(160,40)       D(100,40)" & Chr(13) & _
                  "Using Sutherland Cohen clipping algorithm find the ias able portion of the lie segments EF and GH." & Chr(13) & _
                  "E(50,0)     F(70,80)" & Chr(13) & _
                  "G(120,20)       H(140,80)"
StrQuestions(10) = "Discuss the role of tablet and mouse as graphical input devices."
End Sub

Private Sub questions_paper6()
StrQuestions(1) = "A cubic Bezier curve segment is described by the control points P1(20,20) , P2(40,80), P3(80,80), P4(90,50). Another curve segment is described by Q1(a,b), Q2(c,20), Q3(150,20),Q4(180,20). Determine the values a,b, and c so that the two cure segments join smoothly. Also draw a rough sketch of the two curves."
StrQuestions(2) = "What is the mechanism of showing text material on a graphics screen" & Chr(13) & _
                  "Name the various text attributes usually made available to a graphics programmer."
StrQuestions(3) = "Describe how a 3D object is presented on the screen using perspective proection. Take a simple object for illustration."
StrQuestions(4) = "What is the role of shadow masks used in graphics monitors? What do you understand by VGa and SVGA monitors?"
StrQuestions(5) = "What are the common formats for storing audio, video and still pictures in a multimedia system? What do you understand by the term 'hypertext'? Work out the transformation needed to carry out window to view port mapping."
StrQuestions(6) = "    A line is to be drawn on the screen starting from (8,4) to (2,0). Using a line drawing algorithm find out which pixels would be needed to be shown on the screen."
StrQuestions(7) = "Describe an algorithm to fill the interior of an arbitrary shaped object on the screen."
StrQuestions(8) = "    Describe how computer graphics can be effectively used for teaching & training ."
StrQuestions(9) = "Assume that we are observing the unit cube laced at origin from (20,40,20). Determine which faces are visible and hidden."
StrQuestions(10) = "You are given a line in 3D going from P1(-5,4,2) to P2 (5,-6,18). Project the one to the X = 0 plane (YZ plane ) when the viewer is placed at (10,0,10)."
End Sub

Private Sub questions_paper7()
StrQuestions(1) = "A unit cube is placed at the origin with its edges parallel to the X,Y and Z axes and 3 off the edges touching the axis. Find the  unitcube's orthographic proection on the XY plane after it has been rotted by 450 about the Y-axis."
StrQuestions(2) = "In a perspective transformation -" & Chr(13) & _
                  "a.  How will the image be changed if the objects and the projection lane are fixed and the centre of projection is moved?" & Chr(13) & _
                  "b.  How will the image be changed if the objects and center of projection are fixed and the projection plane is moved?"
StrQuestions(3) = "Assign a 4 bit code tot the end points of a line such that the bit will be 1 when the point is on the screen and 0 when it is off the screen. State the criteria fortrivvially acceptin and rejecting a line segment."
StrQuestions(4) = "Why do we study cubic polynomial cures rather that any other polynomial cure in computer graphics?"
StrQuestions(5) = "Obtain the matrix transformation for an object protected on the XY plane when viewed from (0,0,5). Find the vanishing points of this transformation?"
StrQuestions(6) = "Write a program in C(or an algorithm) which would draw brick tiling on the complete video screen.(Assume any resolution monitor for the monitor)"
StrQuestions(7) = "Give an algorithm for filling a rectangular figure"
StrQuestions(8) = "Give the explicit form of the 3 × 3 matrix "
StrQuestions(9) = ""
StrQuestions(10) = ""
End Sub

Private Sub multiplechoice_Paper1()
strmultiple(1) = "Video conferencing -"
strmuloption(1) = "permits two or more persons to communicate with each other while they can see each other on the computer terminal."
strmuloption(2) = "Refers to videos used in conferences for presenting ideas"
strmuloption(3) = "Is possible only when one uses video-disks"
strmuloption(4) = "Involves sending pictures through high speed networks and audio over telephone lines."

strmultiple(2) = "In 2D graphics reflection of a point about the diagonal line y = x is given by [x' y' 1] = { x y 1} T." & Chr(13) & _
                 "The matrix T is -"
strmuloption(5) = " -1  0  0" & Chr(13) & _
                 "0 1 0" & Chr(13) & _
                 "0 0 1"
strmuloption(6) = "0 1 0 " & Chr(13) & _
                  "0 1 0" & Chr(13) & _
                  "0 0 1"
strmuloption(7) = "0 1 1" & Chr(13) & _
                  "1 0 0" & Chr(13) & _
                  "0 0 1"
strmuloption(8) = "1 1 0" & Chr(13) & _
                  "1 1 0" & Chr(13) & _
                  "0 0 1"

strmultiple(3) = "Interlacing -"
strmuloption(9) = "refers to mixing various shades on the graphics screen."
strmuloption(10) = "Refers to displaying alternate columns on the screen."
strmuloption(11) = "Refers to displaying alternate rows on the screen"
strmuloption(12) = "Is another term for refreshing the screen"

strmultiple(4) = "Choose the correct one, of the following statements related to Bezier curves -"
strmuloption(13) = "A closed Bezier curve can be drawn by taking the first and last control points to be the same."
strmuloption(14) = "If first 3 control points are in a straight line, no curve can be drawn."
strmuloption(15) = "Given a sequence of 4 control points, the Bezier curve can take two different shapes."
strmuloption(16) = "To draw a curve on the screen, all the control points must be located within the boundaries of the screen."

strmultiple(5) = "Raster scan display means that the screen is scanned -"
strmuloption(17) = "top to bottom and right to left"
strmuloption(18) = "left to right and top to bottom"
strmuloption(19) = "bottom to top and left to right"
strmuloption(20) = "bottom to top and right to left"

strmultiple(6) = "To provide comfort to the compute user, the graphics screen must be refreshed at the rate of -"
strmuloption(21) = "5 frames per second"
strmuloption(22) = "20 fames per second"
strmuloption(23) = "50 frames per second"
strmuloption(24) = "200 frames per second"

strmultiple(7) = "The following transformation in 2D graphics" & Chr(13) & _
                 "[x' y' 1 ] = [x y 1]     0 - 1 0" & Chr(13) & _
                 "-1   0 0" & Chr(13) & _
                 "0  0  1" & Chr(13) & _
                 "will result in -"
strmuloption(25) = "rotation about y axis"
strmuloption(26) = "reflection about diagonal with negation slope."
strmuloption(27) = "Invalid operation"
strmuloption(28) = "Reduction in size"

strmultiple(8) = "An image scanner an be used for storing -"
strmuloption(29) = "text material"
strmuloption(30) = "engineering drawings"
strmuloption(31) = "pictures"
strmuloption(32) = "all of the above"

strmultiple(9) = "Sometimes lines appear broken or show staircase effect on the screen as -"
strmuloption(33) = "the screen is not really flat"
strmuloption(34) = "the line drawing algorithm allows only specific pixels to by put ON"
strmuloption(35) = "the algorithm is not efficient for slanting lines."
strmuloption(36) = "There are numerical errors in calculation."

strmultiple(10) = "    A circle drawn on the screen appears to be elliptical as -"
strmuloption(37) = "the screen has a rectangular shape"
strmuloption(38) = "the CRT is not completely spherical"
strmuloption(39) = "it is due to aspect ratio of the monitor"
strmuloption(40) = "out eyes are not at the same level as the screen."

strmulanswer(1) = "permits two or more persons to communicate with each other while they can see each other on the computer terminal."
strmulanswer(2) = "0 1 1" & Chr(13) & _
                  "1 0 0" & Chr(13) & _
                  "0 0 1"
strmulanswer(3) = "Refers to displaying alternate rows on the screen"
strmulanswer(4) = "A closed Bezier curve can be drawn by taking the first and last control points to be the same."
strmulanswer(5) = "left to right and top to bottom"
strmulanswer(6) = "50 frames per second"
strmulanswer(7) = "reflection about diagonal with negation slope."
strmulanswer(8) = "all of the above"
strmulanswer(9) = "the line drawing algorithm allows only specific pixels to by put ON"
strmulanswer(10) = "it is due to aspect ratio of the monitor"

End Sub

Private Sub multiplechoice_Paper2()
strmultiple(1) = "Tick the transformation which will result in invalid operation in 2D graphics -"
strmuloption(1) = "1   -1   0" & Chr(13) & _
                  "-1   -1   0" & Chr(13) & _
                  "0    -1  1  " & Chr(13) & _
strmuloption(2) = "0  0   0  " & Chr(13) & _
                  "-1  0  0" & Chr(13) & _
                  "0 -1 1"
strmuloption(3) = "0    0   0  " & Chr(13) & _
                  "0  -1   0" & Chr(13) & _
                  "-1  1   1"
strmuloption(4) = "1  0  0  " & Chr(13) & _
                  "0 -1  0"

strmultiple(2) = "One of the following applications is not directly linked with multimedia -"
strmuloption(5) = "advertising"
strmuloption(6) = "games development"
strmuloption(7) = "CBT(computer based teaching)"
strmuloption(8) = "Accounting"

strmultiple(3) = "Dragging in computer graphics can be achieved through following transformations -"
strmuloption(9) = "translation"
strmuloption(10) = "rotation"
strmuloption(11) = "scaling"
strmuloption(12) = "mirror reflection"

strmultiple(4) = "    One of the following terms is not associated with the field of computer graphics -"
strmuloption(13) = "buttons"
strmuloption(14) = "deadlock"
strmuloption(15) = "icon"
strmuloption(16) = "fonts"

strmultiple(5) = "A cube is placed at the origin with its edges touching the X, Y and Z axes. The vanishing point observed by viewer on Z axis would change if the object undergoes a small amount of -"
strmuloption(17) = "translation along X direction"
strmuloption(18) = "translation along Y direction"
strmuloption(19) = "translation along Z direction"
strmuloption(20) = "vanishing point would not change"

strmultiple(6) = "In 2D graphics[4 2 0] represents  a point -"
strmuloption(21) = "lying at infinity"
strmuloption(22) = "(4,2)"
strmuloption(23) = "(4,2) and also (2,1)"
strmuloption(24) = "Invalid representation"

strmultiple(7) = "The floating Horizon technique is used in computer graphics for -"
strmuloption(25) = "area filling"
strmuloption(26) = "hidden surface removal"
strmuloption(27) = "polygon clipping"
strmuloption(28) = "zooming"

strmultiple(8) = "For 2 Bezier curves to join smoothly -"
strmuloption(29) = "all control points of both curves must overlap."
strmuloption(30) = "At least one control point should be common."
strmuloption(31) = "Last two points of first curve must be same as first, two points of second curve."
strmuloption(32) = "Last two points of first curve must be collinear with first two points of second curve."

strmultiple(9) = "5 minutes of stereo music are stored in MIDI form as well as a WAV file. The digital audio is 16 bit at 44.1 KHz. Pick the correct statement regarding storage requirement of files."
strmuloption(33) = "30 KB for MIDI and 30 KB for WAV"
strmuloption(34) = "30 KB for MIDI and 50 KB for WAV"
strmuloption(35) = "50 MB for MIDI and 30 KB for WAV"
strmuloption(36) = "50 MB for MIDI and 50 MB for WAV"

strmultiple(10) = "A rectangle has been drawn on the screen. It is desired to carry out a zoom-in process to double the size of the rectangle. This process would involve."
strmuloption(37) = "only scaling"
strmuloption(38) = "scaling and translation"
strmuloption(39) = "scaling and rotation"
strmuloption(40) = "translation, scaling and translation back "

strmulanswer(1) = "1  0  0  " & Chr(13) & _
                  "0 -1  0"
strmulanswer(2) = "Accounting"
strmulanswer(3) = "translation"
strmulanswer(4) = "deadlock"
strmulanswer(5) = "vanishing point would not change"
strmulanswer(6) = "lying at infinity"
strmulanswer(7) = "hidden surface removal"
strmulanswer(8) = "Last two points of first curve must be collinear with first two points of second curve."
strmulanswer(9) = "30 KB for MIDI and 50 KB for WAV"
strmulanswer(10) = "translation, scaling and translation back "

End Sub

Private Sub multiplechoice_Paper3()
strmultiple(1) = "If A = -2     1    0        b = 10  6  1" & Chr(13) & _
                 "1     2    0" & Chr(13) & _
                 "0     0    1"
strmuloption(1) = "b.A can ot be determined as dimensions do not match"
strmuloption(2) = "A.b is same as b.A"
strmuloption(3) = "b.A = [-14  2   1]"
strmuloption(4) = "b.A = [22   -14     1]"

strmultiple(2) = "One of the following pairs has similar functions. Tick the correct pair -"
strmuloption(5) = "track ball and light pen"
strmuloption(6) = "digitizing tablet and the mouse"
strmuloption(7) = "mouse and light pen"
strmuloption(8) = "light pen and digitizing tablet"

strmultiple(3) = "The cross product of two vectors (I - K) and (2I + J) is -"
strmuloption(9) = "2I"
strmuloption(10) = "I - 3J - K"
strmuloption(11) = "3I + J - K"
strmuloption(12) = "3I - J + K"

strmultiple(4) = "The maximum number of control points for drawing B-spline curve is -"
strmuloption(13) = "4"
strmuloption(14) = "    8"
strmuloption(15) = "10"
strmuloption(16) = "could be anything"

strmultiple(5) = "Tick the correct statement "
strmuloption(17) = "When the mouse is moved the cursor moves on the screen. The same effect can be observed when the mouse is picked up and placed in a new location."
strmuloption(18) = "Usually no nits are allocated to Blue color in RGB monitors as it contributes very little to overall colors on the screen."
strmuloption(19) = "The refresh buffer or the display buffer or the display buffer holds the picture information which is displayed on the screen at regular intervals."
strmuloption(20) = "Each of the color monitor screens are painted with 3 layers of red, green and blue phosphorescence material"

strmultiple(6) = "To determine the visibility of a line against a rectangular window using endpoint codes, it is found that the codes for the line are 0101 and 0001. The line would be -"
strmuloption(21) = "totally visible"
strmuloption(22) = "totally invisible"
strmuloption(23) = "partially visible"
strmuloption(24) = "lying along one of the edges of the window"

strmultiple(7) = "In computer graphics, a cubic polynomial is chosen to draw a Bezier curve. Pick out the WRONG one out of the reasons given below -"
strmuloption(25) = "not possible to draw various shapes with lower order polynomial"
strmuloption(26) = "the order of polynomial must be odd."
strmuloption(27) = "very high order results in more computations."
strmuloption(28) = "It guarantees control on end points and the sloped at the end points so that segments may joint smoothly."

strmultiple(8) = "A 14 inches monitor implies -"
strmuloption(29) = "the screen is 14 inches wide"
strmuloption(30) = "the screen is 14 inches high"
strmuloption(31) = "a circle of 14 inches diameter is available for display"
strmuloption(32) = "the diagonal across the screen is 14 inches"

strmultiple(9) = "The Dot product of two vectors(2I - 3J + K) and (I  + J - K)"
strmuloption(33) = "-2"
strmuloption(34) = "1"
strmuloption(35) = "3I - 2J"
strmuloption(36) = "2I-3J-K"

strmultiple(10) = "One of the following sampling rate for storing digital audio is not preferred for multimedia applications."
strmuloption(37) = "44.1 KHz , mono"
strmuloption(38) = "22.05 KHz, stereo"
strmuloption(39) = "11 KHz,stereo"
strmuloption(40) = "5.5 KHz mono"


strmulanswer(1) = "b.A = [-14  2   1]"
strmulanswer(2) = "digitizing tablet and the mouse"
strmulanswer(3) = "3I - J + K"
strmulanswer(4) = "could be anything"
strmulanswer(5) = "The refresh buffer or the display buffer or the display buffer holds the picture information which is displayed on the screen at regular intervals."
strmulanswer(6) = "totally invisible"
strmulanswer(7) = "the order of polynomial must be odd."
strmulanswer(8) = "the diagonal across the screen is 14 inches"
strmulanswer(9) = "-2"
strmulanswer(10) = "5.5 KHz mono"
End Sub

Private Sub multiplechoice_Paper4()
strmultiple(1) = "The scanline seed Fill algorithm -"
strmuloption(1) = "can work for a figure with holes"
strmuloption(2) = "cannot work for a figure with holes"
strmuloption(3) = "can work only or fiture with straight edges"
strmuloption(4) = "can not work for a concave figure"

strmultiple(2) = "In 3D graphics, the transformation" & Chr(13) & _
                 "1   0   0   0" & Chr(13) & _
                 "0   1   0   0.5" & Chr(13) & _
                 "0   0   0   0" & Chr(13) & _
                 "2   0.5 0   1" & Chr(13) & _
                 "Will result in –"
strmuloption(5) = "parallel projection"
strmuloption(6) = "Isometric projection"
strmuloption(7) = "perspective projection"
strmuloption(8) = "generation of top view"

strmultiple(3) = "A perspective view is being generated by a viewer located at distance b from the origin on Z-zxis. If this distance is doubled, the distance of vanishing point from the origin, would -"
strmuloption(9) = "be at twice the distance"
strmuloption(10) = "remain unchanged"
strmuloption(11) = "be half the original"
strmuloption(12) = "be infinity distance"

strmultiple(4) = "The dot product of two vectors would -"
strmuloption(13) = "increase as the angle between them increases"
strmuloption(14) = "decrease as the angle between them increases"
strmuloption(15) = "remain constant irrespective of angle between the vectors"
strmuloption(16) = "decrease as the angle between them decreases."

strmultiple(5) = "45 degrees Rotation about y axis will move a vector on X-axis to -"
strmuloption(17) = "XZ plane"
strmuloption(18) = "XY plane"
strmuloption(19) = "YZ plane"
strmuloption(20) = "Z axis"

strmultiple(6) = "To draw two smoothly joining cubic Bezier curves segments one needs -"
strmuloption(21) = "4 control points"
strmuloption(22) = "7 control points"
strmuloption(23) = "8 control points"
strmuloption(24) = "9 control points"

strmultiple(7) = "In 2D graphics, the transformation -" & Chr(13) & _
                 "1   0   0" & Chr(13) & _
                 "0   1   0" & Chr(13) & _
                 "-1  2   1" & Chr(13) & _
                 "will result in -    "
strmuloption(25) = "scaling"
strmuloption(26) = "translation"
strmuloption(27) = "rotation"
strmuloption(28) = "mirror reflection"

strmultiple(8) = "The following choices represent points in homogenous coordinates system. Indicte the point lying at infinity:"
strmuloption(29) = "[5  3   1   0]"
strmuloption(30) = "[5  3   1   1]"
strmuloption(31) = "[50 30  10  10]"
strmuloption(32) = "[50 30  10  10000]"

strmultiple(9) = "In homogenous coordinate system the 3D point (6,12,0) can be represented in a number of ways. Pick out the INCORRECT choice :"
strmuloption(33) = "[6  12  0   1]"
strmuloption(34) = "[12 24      0   2]"
strmuloption(35) = "[3  6   0   ½]"
strmuloption(36) = "[3  6   0   0]"

strmultiple(10) = "One of the following DOES Not represent point lying on the y axis at infinity."
strmuloption(37) = "[0  1   0   0]"
strmuloption(38) = "[1  0   0   0]"
strmuloption(39) = "[0  5   0   0]"
strmuloption(40) = "0   10  0   0]"


strmulanswer(1) = "can work for a figure with holes"
strmulanswer(2) = "perspective projection"
strmulanswer(3) = "be at twice the distance"
strmulanswer(4) = "decrease as the angle between them increases"
strmulanswer(5) = "XZ plane"
strmulanswer(6) = "7 control points"
strmulanswer(7) = "translation"
strmulanswer(8) = "[5  3   1   0]"
strmulanswer(9) = "[3  6   0   0]"
strmulanswer(10) = "[1  0   0   0]"
End Sub

Private Sub multiplechoice_Paper5()
strmultiple(1) = "A line is parallel to one of the edges of a window and located far from it. The algorithm that would fail to indicate this is -"
strmuloption(1) = "Explicit line clipping algorithm"
strmuloption(2) = "Midpoint subdivision algorithm"
strmuloption(3) = "Sutherland Cohen algorithm"
strmuloption(4) = "There is no such algorithm"

strmultiple(2) = "A point(x,y,z) is transformed to (x’ y’ z’) using one of the following equations:" & Chr(13) & _
                 "[x’ y’  z’  1] = [x     y   z   1]      1 0 0 0" & Chr(13) & _
                 "0 0 1 0" & Chr(13) & _
                 "0 -1 0 0" & Chr(13) & _
                 "0 0 0 1" & Chr(13) & _
                 "OR" & Chr(13) & _
                 "X’ = 1 0 0 0    x" & Chr(13) & _
                 "y’ =  0 0 -1 0  y" & Chr(13) & _
                 "z’ =  0 1  0  0 z" & Chr(13) & _
                 "1  =  0 0  0  1 1" & Chr(13) & _
                 "What does this transform correspond to?"
strmuloption(5) = "translation by unit distance"
strmuloption(6) = "rotation by 90 degrees about X – axis."
strmuloption(7) = "rotation by 90 degrees about Z-axis"
strmuloption(8) = "rotation by 90 degrees about Z-axis"

strmultiple(3) = "To store good quality sound the audio signal in a multimedia PC is sampled at the rate of -"
strmuloption(9) = "44.1 Hz "
strmuloption(10) = "4.41 KHz"
strmuloption(11) = "44. KHz"
strmuloption(12) = "4.41 MHz"

strmultiple(4) = "Pick out thecorrect statement for Bezier curves."
strmuloption(13) = "Best curves are drawn whn starting and end slope of the curves are same."
strmuloption(14) = "A curve is proper only if the cubic polynomials for x and y are identical."
strmuloption(15) = "The control points may lie anywhere on the screen."
strmuloption(16) = "The control points should be so placed that vvalues of x coordinate are in the increasing order."

strmultiple(5) = "A cube of side 5 is placed at the origin such that its edges are parallel to the X,Y and Z axes. A viewer standing on the Z-axis is looking at the cube. The perspective view so generated would have a vanishing point. This would be the point where all edges – "
strmuloption(17) = "of the cube parallel to X-axis appear to meet"
strmuloption(18) = "of the cube parallel to Y-axis appear to meet"
strmuloption(19) = "of the cube parallel to Z-axis appear to meet"
strmuloption(20) = "on top of the cube appear to meet"

strmultiple(6) = "The use of lookup tables in display monitor is to -"
strmuloption(21) = "increase the resolution of the picture on the screen"
strmuloption(22) = "increase the range of coloir shades"
strmuloption(23) = "provide a variety ofcolour palettes but number of available colours remain same"
strmuloption(24) = "store the colour information of the picture"

strmultiple(7) = "The back face remoal algorithm is being applied to an object, one of whose sides is a quadrilateral. The vvertices of theside, read clockwise are A,B,C,D,a. Pick out thecorrect statement from the choices given below:"
strmuloption(25) = "The method is not applicable because theside has to beractangular or square."
strmuloption(26) = "Vectors AB and CD can be used to determine the inward normal"
strmuloption(27) = "Vectors BC and AB can be used todetermne the inward normal"
strmuloption(28) = "Vectors DA and BC can be used to determine the inward normal"

strmultiple(8) = "The cross product of two vectors -"
strmuloption(29) = "lies in the same plane as the vectors"
strmuloption(30) = "has a smaller magnitude than either of the vectors"
strmuloption(31) = "has a large magnitude than either of the vectors."
strmuloption(32) = "is perpendicular to plane containing the vectors    "

strmultiple(9) = "The format for storing digital audio in multimedia applications is-"
strmuloption(33) = "JPEG"
strmuloption(34) = "TIFF"
strmuloption(35) = "WAV"
strmuloption(36) = "BMP"

strmultiple(10) = "Pick out the INCORRECT statement related to authoring tools for multimedia -"
strmuloption(37) = "Authorware issimple and icon driven"
strmuloption(38) = "Animation in Authorware are better than those in Director"
strmuloption(39) = "Compare to Director, Authorware lacks capacity to synchronize sound effectively "
strmuloption(40) = "Director has no built in capabilities for electronic spreadsheet and charting   "

strmulanswer(1) = "There is no such algorithm"
strmulanswer(2) = "rotation by 90 degrees about X – axis."
strmulanswer(3) = "44. KHz"
strmulanswer(4) = "The control points may lie anywhere on the screen."
strmulanswer(5) = "of the cube parallel to Z-axis appear to meet"
strmulanswer(6) = "provide a variety ofcolour palettes but number of available colours remain same"
strmulanswer(7) = "Vectors BC and AB can be used todetermne the inward normal"
strmulanswer(8) = "is perpendicular to plane containing the vectors    "
strmulanswer(9) = "WAV"
strmulanswer(10) = "Animation in Authorware are better than those in Director"
End Sub

Private Sub multiplechoice_Paper6()
strmultiple(1) = "A line is parallel to one of the edges of a window and located far from it. The algorithm that would fail to indicate this is -"
strmuloption(1) = "Explicit line clipping algorithm"
strmuloption(2) = "Midpoint subdivision algorithm"
strmuloption(3) = "Sutherland Cohen algorithm"
strmuloption(4) = "There is no such algorithm"

strmultiple(2) = "A point(x,y,z) is transformed to (x' y' z') using one of the following equations:" & Chr(13) & _
                 "[x' y'  z'  1] = [x     y   z   1]      1 0 0 0" & Chr(13) & _
                 "0 0 1 0" & Chr(13) & _
                 "0 -1 0 0" & Chr(13) & _
                 "0 0 0 1" & Chr(13) & _
                 "OR" & Chr(13) & _
                 "X' = 1 0 0 0    x" & Chr(13) & _
                 "y' =  0 0 -1 0  y" & Chr(13) & _
                 "z' =  0 1  0  0 z" & Chr(13) & _
                 "1  =  0 0  0  1 1" & Chr(13) & _
                 "What does this transform correspond to?"
strmuloption(5) = "translation by unit distance"
strmuloption(6) = "rotation by 90 degrees about X - axis."
strmuloption(7) = "rotation by 90 degrees about Z-axis"
strmuloption(8) = "rotation by 90 degrees about Z-axis"

strmultiple(3) = "To store good quality sound the audio signal in a multimedia PC is sampled at the rate of -"
strmuloption(9) = "44.1 Hz "
strmuloption(10) = "4.41 KHz"
strmuloption(11) = "44. KHz"
strmuloption(12) = "4.41 MHz"

strmultiple(4) = "Pick out the correct statement for Bezier curves."
strmuloption(13) = "Best curves are drawn when starting and end slope of the curves are same."
strmuloption(14) = "A curve is proper only if the cubic polynomials for x and y are identical."
strmuloption(15) = "The control points may lie anywhere on the screen."
strmuloption(16) = "The control points should be so placed that values of x coordinate are in the increasing order."

strmultiple(5) = "A cube of side 5 is placed at the origin such that its edges are parallel to the X,Y and Z axes. A viewer standing on the Z-axis is looking at the cube. The perspective view so generated would have a vanishing point. This would be the point where all edges - "
strmuloption(17) = "of the cube parallel to X-axis appear to meet"
strmuloption(18) = "of the cube parallel to Y-axis appear to meet"
strmuloption(19) = "of the cube parallel to Z-axis appear to meet"
strmuloption(20) = "on top of the cube appear to meet"

strmultiple(6) = "The use of lookup tables in display monitor is to -"
strmuloption(21) = "increase the resolution of the picture on the screen"
strmuloption(22) = "increase the range of color shades"
strmuloption(23) = "provide a variety of color palettes but number of available colors remain same"
strmuloption(24) = "store the color information of the picture"

strmultiple(7) = "The back face removal algorithm is being applied to an object, one of whose sides is a quadrilateral. The vertices of the side, read clockwise are A,B,C,D,a. Pick out the correct statement from the choices given below:"
strmuloption(25) = "The method is not applicable because the side has to be rectangular or square."
strmuloption(26) = "Vectors AB and CD can be used to determine the inward normal"
strmuloption(27) = "Vectors BC and AB can be used to determine the inward normal"
strmuloption(28) = "Vectors DA and BC can be used to determine the inward normal"

strmultiple(8) = "The cross product of two vectors -"
strmuloption(29) = "lies in the same plane as the vectors"
strmuloption(30) = "has a smaller magnitude than either of the vectors"
strmuloption(31) = "has a large magnitude than either of the vectors."
strmuloption(32) = "is perpendicular to plane containing the vectors    "

strmultiple(9) = "The format for storing digital audio in multimedia applications is-"
strmuloption(33) = "JPEG"
strmuloption(34) = "TIFF"
strmuloption(35) = "WAV"
strmuloption(36) = "BMP"

strmultiple(10) = "Pick out the INCORRECT statement related to authoring tools for multimedia -"
strmuloption(37) = "Author ware is simple and icon driven"
strmuloption(38) = "Animation in Author ware are better than those in Director"
strmuloption(39) = "Compare to Director, Author ware lacks capacity to synchronize sound effectively "
strmuloption(40) = "Director has no built in capabilities for electronic spreadsheet and charting   "

strmulanswer(1) = "There is no such algorithm"
strmulanswer(2) = "rotation by 90 degrees about X - axis."
strmulanswer(3) = "44. KHz"
strmulanswer(4) = "The control points may lie anywhere on the screen."
strmulanswer(5) = "of the cube parallel to Z-axis appear to meet"
strmulanswer(6) = "provide a variety of color palettes but number of available colors remain same"
strmulanswer(7) = "Vectors BC and AB can be used to determine the inward normal"
strmulanswer(8) = "is perpendicular to plane containing the vectors    "
strmulanswer(9) = "WAV"
strmulanswer(10) = "Animation in Author ware are better than those in Director"
End Sub

Private Sub multiplechoice_Paper7()
strmultiple(1) = "In 3D graphics, one of the following matrices would carry out rotation about the Y axis by 45 degrees -"
strmuloption(1) = "1   0   0   0" & Chr(13) & _
                  "0   .707    707 0" & Chr(13) & _
                  "0   -.707   .707    0" & Chr(13) & _
                  "0   0   0   1"
strmuloption(2) = "707 0   -.707   0" & Chr(13) & _
                  "0   1   0   0" & Chr(13) & _
                  ".707    0   .707    0" & Chr(13) & _
                  "0   0   0   1"
strmuloption(3) = "707 -.707   0   0" & Chr(13) & _
                  "-.707   .707    0   0" & Chr(13) & _
                  "0   0   1   0" & Chr(13) & _
                  "0   0   0   1"
strmuloption(4) = "1   0   -1  0" & Chr(13) & _
                  "0   1   0   0" & Chr(13) & _
                  ".707    0   .707    ." & Chr(13) & _
                  "0   0   0   1   "

strmultiple(2) = "T1,T2,T3,T4 are consecutive corners of a square. A Bezier curve drawn with these control points -"
strmuloption(5) = "would be a circle"
strmuloption(6) = "would touch all sides of a square"
strmuloption(7) = "would enclose the square"
strmuloption(8) = "will have no oscillations."

strmultiple(3) = "Pick out the wrong statement regarding the floating Horizon techniques -"
strmuloption(9) = "is used to create out door scene on the screen"
strmuloption(10) = "is used to plot 3D mathematical function on the screen."
strmuloption(11) = "is used for hidden surface elimination."
strmuloption(12) = "can be used only for Cuboid, Cones and Spheres etc."

strmultiple(4) = "In 3d graphics, the homogenous coordinates(0,5,0,0) -"
strmuloption(13) = "are not allowed"
strmuloption(14) = "represents a point X = 0, Y = 5, Z = 0"
strmuloption(15) = "represents a point in Y axis at infinity"
strmuloption(16) = "represents a point on X-axis or Z- axis at infinity"

strmultiple(5) = "90 degrees clockwise rotation about X axis would transfer point -"
strmuloption(17) = "on X axis to Z axis"
strmuloption(18) = "on Y axis to Z axis"
strmuloption(19) = "on Y axis to X axis"
strmuloption(20) = "on Z axis to Y axis"

strmultiple(6) = "Anticlockwise rotation of 90 degrees about Z axis would transfer a point -"
strmuloption(21) = "on X axis to negative"
strmuloption(22) = "on Y axis to X axis"
strmuloption(23) = "on X axis to Y axis"
strmuloption(24) = "on X axis to negative Y  axis"

strmultiple(7) = "A video sequence is usually captured by a video recorder at the rate of -"
strmuloption(25) = "1 frame/sec"
strmuloption(26) = "2 frames/sec"
strmuloption(27) = "25 frames/sec"
strmuloption(28) = "100 frames/sec"

strmultiple(8) = "Translation operations can result in -"
strmuloption(29) = "panning "
strmuloption(30) = "zooming"
strmuloption(31) = "interlacing"
strmuloption(32) = "viewing"

strmultiple(9) = "A multimedia PC gets booted through its -"
strmuloption(33) = "ROM"
strmuloption(34) = "RAM"
strmuloption(35) = "Cache"
strmuloption(36) = "CD-ROM"

strmultiple(10) = "Back ground audio for a text material is prepared on a 133 MHz multi media PC. When this is used in a 200 MHz PC, the audio track will -"
strmuloption(37) = "sound the same"
strmuloption(38) = "run faster"
strmuloption(39) = "run slower"
strmuloption(40) = "use synchronism with text material"

strmulanswer(1) = "707 0   -.707   0" & Chr(13) & _
                  "0   1   0   0" & Chr(13) & _
                  ".707    0   .707    0" & Chr(13) & _
                  "0   0   0   1"
strmulanswer(2) = "will have no oscillations."
strmulanswer(3) = "can be used only for Cuboid, Cones and Spheres etc."
strmulanswer(4) = "represents a point on X-axis or Z- axis at infinity"
strmulanswer(5) = "on Z axis to Y axis"
strmulanswer(6) = "on X axis to Y axis"
strmulanswer(7) = "25 frames/sec"
strmulanswer(8) = "panning "
strmulanswer(9) = "ROM"
strmulanswer(10) = "sound the same"
 

End Sub

Private Sub multiplechoice_Paper8()
strmultiple(1) = "One of the following is not related to multimedia"
strmuloption(1) = "FoxPro"
strmuloption(2) = "Authorware"
strmuloption(3) = "Director"
strmuloption(4) = "Animatorpro"

strmultiple(2) = "The perspective projection of the point(2,3,1) on the Z =0 plane with the eye placed at (0,0, -2) is -"
strmuloption(5) = "(2/3, 1.5)"
strmuloption(6) = "(1 2/3, 2.5)"
strmuloption(7) = "(2 1/3, 3.5)"
strmuloption(8) = "None of the above"

strmultiple(3) = "The perspective projection of the point (1,2,2) on the Z = 0 plane with the eye being placed at (0,0 -2) is -"
strmuloption(9) = "(0.5,1)"
strmuloption(10) = "(1,1.5)"
strmuloption(11) = "(1.5,2)"
strmuloption(12) = "None of the above"

strmultiple(4) = "An object is viewed by using perspective transformation equations. Then the maximum number of principal vanishing points possible are -"
strmuloption(13) = "1"
strmuloption(14) = "2"
strmuloption(15) = "3"
strmuloption(16) = "Infinite"

strmultiple(5) = "Which one of the following  transmits digital information to the computer?"
strmuloption(17) = "joystick"
strmuloption(18) = "mouse"
strmuloption(19) = "graphics tablet"
strmuloption(20) = "None of the above"

strmultiple(6) = "A B-spline curve is called a uniform B-spline when-"
strmuloption(21) = "the spacing between the adjacent knot values is a constant"
strmuloption(22) = "the distance between the adjacent control points is a constant"
strmuloption(23) = "the curve is infinite times differentiable."
strmuloption(24) = "None of the above"

strmultiple(7) = "Reflection about the line y = -x transformation is equal to -"
strmuloption(25) = "the sequence of transformations:" & Chr(13) & _
                   "i.  clockwise rotation by 45 degrees" & Chr(13) & _
                   "ii. reflection about the X-axis" & Chr(13) & _
                   "iii counter clockwise rotation by 45 degrees"""
strmuloption(26) = "the matrix transformation" & Chr(13) & _
                   "0   1   0" & Chr(13) & _
                   "1   0   0" & Chr(13) & _
                   "0   0   1" & Chr(13) & _
                   "in homogenous coordinate system"
strmuloption(27) = "reflection about the X-axis followed by a counter clockwise rotation of 90 degrees."
strmuloption(28) = "None of the above"

strmultiple(8) = "Under a parallel projection the point(2,3,-1) has been viewed at (3,3,0) then the direction of projection should be the vector -"
strmuloption(29) = "(1,0,1)"
strmuloption(30) = "(1,0,-1)"
strmuloption(31) = "(0,1,1)"
strmuloption(32) = "(0,-1,1)"

strmultiple(9) = "Let p(.) and q(.) be two Bezier curves defined by the sets of control points {p0, p1, p2, p3} and {q0, q1, q2, q3} respectively. The two curves are joined with first order continuity at p(1) = q(0) provided -"
strmuloption(33) = "p3 = (p2 + q1) /2"
strmuloption(34) = "p3 = (p2 + q1) /3"
strmuloption(35) = "p3 = (p1 + q2)/2"
strmuloption(36) = "p3 = (p1 + q2) /3"

strmultiple(10) = "Display resolution of a typical monitor would be around -"
strmuloption(37) = "10 dpi"
strmuloption(38) = "60 dpi"
strmuloption(39) = "200 dpi"
strmuloption(40) = "300 dpi"

strmulanswer(1) = "FoxPro"
strmulanswer(2) = "(1 2/3, 2.5)"
strmulanswer(3) = "(0.5,1)"
strmulanswer(4) = "3"
strmulanswer(5) = "None of the above"
strmulanswer(6) = "the spacing between the adjacent knot values is a constant"
strmulanswer(7) = "None of the above   "
strmulanswer(8) = "(1,0,1)"
strmulanswer(9) = "p3 = (p2 + q1) /2"
strmulanswer(10) = "60 dpi"
End Sub

Private Sub multiplechoice_Paper9()
strmultiple(1) = "One of the following techniques is used to provide variety of color palettes on the graphics screen -"
strmuloption(1) = "using multi colored electron gun"
strmuloption(2) = "increasing no of bit planes for R, G and B"
strmuloption(3) = "using look up table"
strmuloption(4) = "increasing no, of resolution of screen"

strmultiple(2) = "A homogenous coordinate system corresponding to N-dimensional coordingate system provides -"
strmuloption(5) = "N dimensions"
strmuloption(6) = "N + 1 dimensions"
strmuloption(7) = "N + 2 dimensions"
strmuloption(8) = "None of the above"

strmultiple(3) = "The equation of the line passing through P0(1,-5,2) and P1(6,7,-3) is -"
strmuloption(9) = "-5I + 12J - 5K"
strmuloption(10) = "5I + 12J + 5K"
strmuloption(11) = "5I + 12J - 5K"
strmuloption(12) = "5I - 12J -5K"

strmultiple(4) = "The following matrix -" & Chr(13) & _
    "a 1 a 0" & Chr(13) & _
    "b 1 0" & Chr(13) & _
    "0 0 1" & Chr(13) & _
    "defines a transformation called -"
strmuloption(13) = "translation"
strmuloption(14) = "rotation"
strmuloption(15) = "scaling"
strmuloption(16) = "shearing"

strmultiple(5) = "A triangle has been drawn on the screen. If it were to have orientation, the process would require-"
strmuloption(17) = "only rotation"
strmuloption(18) = "translation and rotation"
strmuloption(19) = "translation, rotation and translation back"
strmuloption(20) = "rotation and translation"

strmultiple(6) = "If the window parameters are xw [min] = 1, xw [max] = 3, and the viewport parameters xv[min] = 0, xv[max] = 1 the scale S[x]x is -"
strmuloption(21) = "1"
strmuloption(22) = "1/2"
strmuloption(23) = "2"
strmuloption(24) = "1/4"

strmultiple(7) = "The principal vanishing points for the standard perspective transformations are -"
strmuloption(25) = "three"
strmuloption(26) = "two"
strmuloption(27) = "one"
strmuloption(28) = "none"

strmultiple(8) = "what are the end points codes of the line between P1(-4,7) and P2(-2,10) given the window coordinates x [min] = -3 , x[max] = 2, y [min] = 1 and y[max] = 6"
strmuloption(29) = "(1001),(1000)"
strmuloption(30) = "(1010),(1001)"
strmuloption(31) = "(0000),(0010)"
strmuloption(32) = "(0101), (0100)"

strmultiple(9) = "Interlacing in raster scanning eliminates -"
strmuloption(33) = "only the horizontal retrace"
strmuloption(34) = "only the vertical retrace"
strmuloption(35) = "both horizontal and vertical retraces"
strmuloption(36) = "none of the above"

strmultiple(10) = "Seed fill algorithms, using a stack use the principle of -"
strmuloption(37) = "LIFO"
strmuloption(38) = "FIFO"
strmuloption(39) = "FILO"
strmuloption(40) = "LIFO"

strmulanswer(1) = "using look up table"
strmulanswer(2) = "N + 1 dimensions"
strmulanswer(3) = "5I + 12J - 5K"
strmulanswer(4) = "shearing"
strmulanswer(5) = "translation, rotation and translation back"
strmulanswer(6) = "1/2"
strmulanswer(7) = "one"
strmulanswer(8) = "(0000),(0010)"
strmulanswer(9) = "both horizontal and vertical retraces"
strmulanswer(10) = "FILO"
End Sub

Private Sub multiplechoice_Paper10()
strmultiple(1) = "The result of midpoint subdivision algorithm is to perform "
strmuloption(1) = "breadth first search"
strmuloption(2) = "depth first search"
strmuloption(3) = "logarithmic search"
strmuloption(4) = "best first search for the intersection point of the line with the window edge"

strmultiple(2) = "In the generation of circle by the Bresenham's algorithm is simple to generate -"
strmuloption(5) = "all octants at one go"
strmuloption(6) = "one octant first and others by successive reflections"
strmuloption(7) = "one octant first and other octants by successive rotations"
strmuloption(8) = "one octant first and others by successive translations"

strmultiple(3) = "The end points a given line are (0,0) and (6,8) . The slope and y intercept are computes as -"
strmuloption(9) = "3,3"
strmuloption(10) = "0,3"
strmuloption(11) = "3,0"
strmuloption(12) = "3,-3"

strmultiple(4) = "B-spline curves allow -"
strmuloption(13) = "Only the first order derivative to be continuous at the end points of curve"
strmuloption(14) = "only the second derivative to be continuous at the end points of curve"
strmuloption(15) = "both first and second derivative to be continuous at the end points"
strmuloption(16) = "first derivative at one end and the second derivative at the other end of the curve"

strmultiple(5) = "To determine the visibility of a line, its intersecting with the clipping window is determined. The end point codes for some lines -"
strmuloption(17) = "0010,0001"
strmuloption(18) = "0010,0000"
strmuloption(19) = "0000,0000"
strmuloption(20) = "0100,0000"

strmultiple(6) = "The dot product of a vector I = 3J + 2K, with another vector is 3. The other vector is -"
strmuloption(21) = "2I - 3K"
strmuloption(22) = "I + J"
strmuloption(23) = "3(I -3J + 2K)"
strmuloption(24) = "none of the above"

strmultiple(7) = "A Bezier curve is described by following parameter equation -" & Chr(13) & _
    "x(t) = t3 - 3t2 + 2t + 5" & Chr(13) & _
    "y(t) = 2t3 + 6t + 3" & Chr(13) & _
    "which of the following statements is true ?"
strmuloption(25) = "the curve ends at (5,3)"
strmuloption(26) = "it represents a circle"
strmuloption(27) = "the curve starts from (5,3)"
strmuloption(28) = "the equation can not represent a curve, as one term is missing from polynomial for y(t)"

strmultiple(8) = "For two parallel lines in a plane (XY plane) one of the following is true -"
strmuloption(29) = "lines must make equal intercept on Y axis and -Y axis"
strmuloption(30) = "lines must make equal intercept on X axis and -X axis"
strmuloption(31) = "lines must make equal intercept on X axis and Y axis"
strmuloption(32) = "none of the above statements are true"

strmultiple(9) = "If vector P = 1 (unit vector along X axis) and vector Q = J (unit vector along y axis), the Ross product of the vector would be along -"
strmuloption(33) = "I + J"
strmuloption(34) = "I - J"
strmuloption(35) = "K"
strmuloption(36) = "I"

strmultiple(10) = "Pick the correct statement from the following regarding nature of multimedia"
strmuloption(37) = "it is same as word processing as it also handles text"
strmuloption(38) = "it is same as a television program as both involve sound and moving pictures"
strmuloption(39) = "it is same as a vides because both deal with selection of audio and video material"
strmuloption(40) = "none of the above statements is true"

strmulanswer(1) = "logarithmic search"
strmulanswer(2) = "one octant first and others by successive reflections"
strmulanswer(3) = "3,0"
strmulanswer(4) = "first derivative at one end and the second derivative at the other end of the curve"
strmulanswer(5) = "0000,0000"
strmulanswer(6) = "none of the above"
strmulanswer(7) = "the curve starts from (5,3)"
strmulanswer(8) = "none of the above statements are true"
strmulanswer(9) = "K"
strmulanswer(10) = "none of the above statements is true"

End Sub


