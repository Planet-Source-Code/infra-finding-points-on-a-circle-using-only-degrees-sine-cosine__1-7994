VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+-----------------------------------------------+
'|         ***Source Code Information***         |
'|                                               |
'|Author:    InfraRed                            |
'|                                               |
'|E-Mail:    InfraRed@flashmail.com              |
'|                                               |
'|ICQ UIN:   17948286                            |
'|                                               |
'|Comments:  I hope you enjoy my source code.  I |
'|worked very hard on this, and if you use       |
'|anything from here, I would like to get credit |
'|for it.  If it makes you feel any better, you  |
'|can e-mail/ICQ me and ask permission to use my |
'|source code...  BUT you don't have to!  If you |
'|have any complaints, compliments, comments,    |
'|threats, fan mail, junk mail, hate mail, or    |
'|anything else you can think of, go ahead and   |
'|send.                                          |
'|                                               |
'|              ***Enjoy my code!***             |
'+-----------------------------------------------+

Dim DDI As Integer, Rad As Single, RotS As Integer, RotC As Integer, CentX As Single, CentY As Single, DotX(359) As Single, DotY(359) As Single, DotC As Long
'DDI = Dot Degree Interval
'Rad = Radius
'RotS = Rotation Speed
'RotC = Current Rotation (Speed)
'CentX = Center (X)
'CentY = Center (Y)
'DotX = Dot Location (X)
'DotY = Dot Location (Y)
'DotC = Dot Color

'Here are all of the property lets and gets, self explanitory
'----------------------------------------------------
Property Let Dot_Degree_Interval(NewDDI As Integer)
DDI = NewDDI
End Property

Property Let Radius(NewRad As Single)
Rad = NewRad
End Property

Property Let Dot_Color(NewColor As Long)
DotC = NewColor
End Property

Property Let Rotation_Speed(Pixel_Amount As Integer)
RotS = Pixel_Amount
End Property

Property Get Dot_Degree_Interval() As Integer
Dot_Degree_Interval = DDI
End Property

Property Get Radius() As Single
Radius = Rad
End Property

Property Get Dot_Color() As Long
DotColor = DotC
End Property

Property Get Rotation_Speed() As Integer
Rotation_Speed = RotS
End Property
'----------------------------------------------------

Public Sub Set_Center_Position(NewX As Single, NewY As Single) 'Sets the new center postion, self explanitory
CentX = NewX
CentY = NewY
End Sub

Public Sub Draw_Circle(YI As Integer) 'This draws the circle
Dim DegP(359) As Double, XP(359) As Single, YP(359) As Single, i As Integer
'DepP = Degree Position
'XP = X Position
'YP = Y Position

  If DC(YI) = True Then
  Erase_All YI
  Exit Sub
  End If
For i = 0 To (360 / DDI) - 1 'Loop through every dot
EraseP DotX(i), DotY(i) 'Erase the old dot
DegP(i) = ((360 / DDI) * i) + RotC 'Find the degree (in the circle) of the current dot and add rotation
XP(i) = (Cos(DegP(i)) * Radius) + CentX 'Find the dot's true X position and center it
YP(i) = (Sin(DegP(i)) * Radius) + CentY 'Find the dot's true Y position and center it
  '--------------------------------------------------
  'Now, I understand that some people may not know
  'How to use Cosine and Sine to find the
  'coordinates of dots on a circle, so I will explain
  'it to the best of my ability.  Here is a quick
  'explanation.  Now, you know that coordinates are
  'shown in (X, Y), well, Cosine (Cos) finds the X
  'and Sine (Sin) finds the Y.  So really, you could
  'think of Sine and Cosine as (Cosine, Sine).  Don't
  'get confused yet, lol, I will explain this further.
  'Now, Cosine can be used to find the coordinates of
  'a certain point by using the degrees of that point
  'Here is a quick example:
  'Cosine(Point_Degree) * Radius_Length = The X
  'coordinate of that Point.  And:
  'Sine(Point_Degree) * Radius_Length = The Y
  'coordinate of that Point.  Here is an example of
  'finding the (X, Y) of a point with the degree
  'measurement of 100°, and the circle has a radius
  'of 5.  To find the X:
  'Cos(100) * 5
  'To find the Y:
  'Sin(100) * 5
  'Simple enough, right?  Now you may notice the
  'CentX and the CentY.  Those are used for me to
  'put the circle at the point I want.  So I just
  'add a certain amount to the X and Y like this:
  '(X + CentX, Y + CentY) = New (X, Y)
  'That is all that does, and if it confuses you,
  'just ignore it for now.  I hope this little
  'tutorial helps you understand the use of Sine
  'and Cosine in finding the coordinates of a
  'point on a circle.
  '--------------------------------------------------
DotX(i) = XP(i) 'Set the main dot position (X)
DotY(i) = YP(i) 'Set the main dot position (Y)
      If DC(YI) = False Then SetP XP(i), YP(i), DotC 'Draw dot
DoEvents
Next i 'Loop back
'Add to the rotation
'-------------------------
RotC = RotC + RotS
  Do While RotC > DDI
  RotC = RotC - DDI
  Loop
    If DC(YI) = True Then
    Erase_All YI
    Exit Sub
    End If
'-------------------------
End Sub

Private Sub EraseP(X As Single, Y As Single) 'This erases a point, self explanitory
MainFrm.PSet (X, Y), RGB(0, 0, 0)
End Sub

Private Sub SetP(X As Single, Y As Single, Color As Long) 'This sets the color of a certain point on the screen, self explanitory
MainFrm.PSet (X, Y), Color
End Sub

Public Sub Erase_All(MI As Integer) 'This erases all the points of the circle, self explanitory
Dim i As Integer
For i = 0 To 359
EraseP DotX(i), DotY(i)
Next i
DC(MI) = False
End Sub
