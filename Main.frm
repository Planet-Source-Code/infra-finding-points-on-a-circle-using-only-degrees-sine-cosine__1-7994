VERSION 5.00
Begin VB.Form MainFrm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "CSS"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer PrepVal 
      Interval        =   1
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer CRT 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer CACI 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Dim MCA As Integer, CEMin As Integer, CEMax As Integer, CAA As Integer, CAR As Boolean, NDDI As Integer, NRS As Integer, CGR(9) As Boolean, MPX As Single, MPY As Single, DirX As Integer, DirY As Integer, CCPX As Single, CCPY As Single

'Sorry for all of the abreviations, this will help:
  'MCA = Max Circle Amount
  'CEMin = Circle Expansion Minimum
  'CEMax = Circle Expansion Maximum
  'CAA = Circle Amount Activated
  'CAR = Circle Activation Reverse
  'NDDI = Next Dot Degree Interval
  'NRS = Next Rotation Speed
  'CGR = Circle Go Reverse
  'MPX = Mouse Position X
  'MPY = Mouse Position Y
  'DirX = Direction X (0 = left, 1 = right)
  'DirY = Direction Y (0 = down, 1 = up)
  'CCPX = Circle Center Position X
  'CCPY = Circle Center Position Y

Private Sub CACI_Timer() 'CACI = Circle Amount Change Interval
If CAR = False Then 'If the max amount of circles are here, don't add any more!
CAA = CAA + 1 'Add to circle amount
PrepareCM 'Get the new circle ready and loaded
  If CAA = MCA Then CAR = True 'If the max circles has been reached, don't allow any more to form
End If
End Sub

Private Sub CRT_Timer() 'Circle Rotation Timer
Dim i As Integer
  If CCPX + DA + CEMax > Screen.Width And DirX = 1 Then 'Don't let the circle go off the right side of the screen
  DirX = 0 'Change the X direction (to left)
  ElseIf CCPX - (DA + CEMax) < 0 And DirX = 0 Then 'Don't let the circle go off the left side of the screen
  DirX = 1 'Change the X direction (to right)
  End If
  If CCPY + DA + CEMax > Screen.Height And DirY = 1 Then 'Don't let the circle go off the bottom of the screen
  DirY = 0 'Change the X direction (to up)
  ElseIf CCPY - (DA + CEMax) < 0 And DirY = 0 Then 'Don't let the circle go off the top of the screen
  DirY = 1 'Change the X direction (to down)
  End If
    'Add to center offset (move)
    '--------------------------------
    If DirX = 1 Then CCPX = CCPX + DA
    If DirX = 0 Then CCPX = CCPX - DA
    If DirY = 1 Then CCPY = CCPY + DA
    If DirY = 0 Then CCPY = CCPY - DA
    '--------------------------------
For i = 0 To CAA - 1 'Loop through circles
  If CM(i).Radius + CA > TP(i) And CGR(i) = False Then 'When the circle reaches the target radius, make it shrink back to 0
  CM(i).Radius = TP(i)
  CGR(i) = True
  ElseIf CM(i).Radius - CA < 1 And CGR(i) = True Then 'When circle reaches 0 radius, select a new target radius and start expanding
  CM(i).Radius = 1
  CGR(i) = False
  TP(i) = RndTP
  ElseIf CGR(i) = False Then
  CM(i).Radius = CM(i).Radius + CA 'Expand circle
  ElseIf CGR(i) = True Then
  CM(i).Radius = CM(i).Radius - CA 'Contract circle
  End If
CM(i).Set_Center_Position CCPX, CCPY 'Set the new center position
CM(i).Draw_Circle i 'Draw the current circle
Next i 'Loop back
End Sub

Private Sub Form_Click()
Unload Me 'Unload form
End
End Sub

Private Sub Form_Load()
SetUpSystem 'Hide mouse
LoadRegValues 'Loads the options (saved in registry)
End Sub

Public Sub LoadRegValues() 'Loads all of the options in the registry, self explanitory
CACI.Interval = GetSetting("Cyclone", "OP", "0", "30") & "000"
MCA = GetSetting("Cyclone", "OP", "1", "5")
CEMin = GetSetting("Cyclone", "OP", "2", "500")
CEMax = GetSetting("Cyclone", "OP", "3", "3000")
NDDI = GetSetting("Cyclone", "OP", "4", "6")
NRS = GetSetting("Cyclone", "OP", "5", "3")
End Sub

Public Sub PrepareValues()
Randomize Timer 'Generate a new random number set
CAA = 1 'Startup with 1 circle loaded
DirX = Int((2 - 1 + 1) * Rnd + 1) - 1 'Select random direction (left or right)
DirY = Int((2 - 1 + 1) * Rnd + 1) - 1 'Select random direction (up or down)
CCPX = Int((Screen.Width - CEMax + 1) * Rnd + CEMax) 'Select a random place (X) for the circle to start at
CCPY = Int((Screen.Height - CEMax + 1) * Rnd + CEMax) 'Select a random place (Y) for the circle to start at
PrepareCM 'Prepare the first circle
End Sub

Public Sub PrepareCM() 'Prepares and creates circles
CM(CAA - 1).Dot_Color = SelRndCol 'Select a random color for the circle
CM(CAA - 1).Dot_Degree_Interval = NDDI 'Select the circle's DDI (Dot Degree Interval)
CM(CAA - 1).Radius = 1 'Set its radius to 1
CM(CAA - 1).Rotation_Speed = NRS 'Set it to the default rotation speed
CM(CAA - 1).Set_Center_Position CCPX, CCPY 'Put it in the same position as the other circles
TP(CAA - 1) = RndTP 'Select the circle's target radius
CGR(CAA - 1) = False 'The circle is NOT contracting
End Sub

Public Sub KillCM() 'Kill the circle
DC(CAA) = True
End Sub

Function SelRndCol() As Long 'Select a random color, self explanitory
Dim Col(2) As Integer, i As Integer
Randomize Timer
Start:
For i = 0 To 2
Col(i) = Int((255 - 1 + 1) * Rnd + 1)
Next i
  If Not Col(0) > 150 Or Not Col(1) > 150 Or Not Col(2) > 150 Then GoTo Start
SelRndCol = RGB(Col(0), Col(1), Col(2))
End Function

Function RndTP() As Integer 'Select a random target radius, self explanitory
Randomize Timer
RndTP = Int((CEMax - CEMin + 1) * Rnd + CEMin)
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'If the mouse moves, kill the screensaver
If Not MPX = Mouse.X Or Not MPY = Mouse.Y Then
Unload Me
End
End If
End Sub

Private Sub Form_Unload(Cancel As Integer) 'Show the cursor when unloaded (when running this in a compiler)
ShowCursor True
End Sub

Private Sub PrepVal_Timer() 'Prepare values after the form has maximized
PrepareValues
CRT.Enabled = True
PrepVal.Enabled = False
End Sub

Public Sub SetUpSystem() 'Get original mouse positions
MPX = Mouse.X
MPY = Mouse.Y
Ontop Me 'Set ontop
End Sub
