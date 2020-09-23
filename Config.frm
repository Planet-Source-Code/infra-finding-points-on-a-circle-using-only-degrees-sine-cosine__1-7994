VERSION 5.00
Begin VB.Form Config 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox OP 
      Height          =   315
      Index           =   5
      Left            =   2520
      TabIndex        =   12
      Text            =   "OP"
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox OP 
      Height          =   315
      Index           =   4
      Left            =   2520
      TabIndex        =   11
      Text            =   "OP"
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox OP 
      Height          =   315
      Index           =   3
      Left            =   2520
      TabIndex        =   10
      Text            =   "OP"
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox OP 
      Height          =   315
      Index           =   2
      Left            =   2520
      TabIndex        =   9
      Text            =   "OP"
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox OP 
      Height          =   315
      Index           =   1
      Left            =   2520
      TabIndex        =   8
      Text            =   "OP"
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox OP 
      Height          =   315
      Index           =   0
      Left            =   2520
      TabIndex        =   7
      Text            =   "OP"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      X1              =   120
      X2              =   3480
      Y1              =   2290
      Y2              =   2290
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   120
      X2              =   3480
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label6 
      Caption         =   "Rotation Speed:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Dot Degree Interval:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Circle Expansion Maximum:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Circle Expansion Minimum:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Max Circle Amount:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Circle Amount Change Interval:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Config"
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

Private Sub Form_Load() 'Prepare all of the values
LoadValues
LoadRegValues
End Sub

Public Sub LoadValues() 'Add all of the numbers to the combo boxes
Dim i As Integer
For i = 1 To 60
OP(0).AddItem i
Next i
For i = 1 To 10
OP(1).AddItem i
Next i
For i = 1 To 500
OP(2).AddItem i
Next i
For i = 500 To 5000
OP(3).AddItem i
Next i
For i = 1 To 6
OP(4).AddItem i
Next i
For i = 8 To 10
OP(4).AddItem i
Next i
OP(4).AddItem 12
OP(4).AddItem 18
OP(4).AddItem 36
For i = 0 To 5
OP(5).AddItem i
Next i
End Sub

Public Sub LoadRegValues() 'Load current registry values (options), self explanitory
OP(0).Text = GetSetting("Cyclone", "OP", "0", "30")
OP(1).Text = GetSetting("Cyclone", "OP", "1", "5")
OP(2).Text = GetSetting("Cyclone", "OP", "2", "500")
OP(3).Text = GetSetting("Cyclone", "OP", "3", "3000")
OP(4).Text = GetSetting("Cyclone", "OP", "4", "6")
OP(5).Text = GetSetting("Cyclone", "OP", "5", "3")
End Sub

Public Sub SaveRegValues() 'Save all registry values, self explanitory
Call SaveSetting("Cyclone", "OP", "0", OP(0).Text)
Call SaveSetting("Cyclone", "OP", "1", OP(1).Text)
Call SaveSetting("Cyclone", "OP", "2", OP(2).Text)
Call SaveSetting("Cyclone", "OP", "3", OP(3).Text)
Call SaveSetting("Cyclone", "OP", "4", OP(4).Text)
Call SaveSetting("Cyclone", "OP", "5", OP(5).Text)
End Sub

Private Sub OK_Click() 'Save registry values and unload configuration screen
SaveRegValues
Unload Me
End Sub

Private Sub OP_Change(Index As Integer) 'Checks if the typed value is one of the values in combo box, if not, set the combo box text to nothing
Dim i As Integer, Yes As Boolean
For i = 0 To OP(Index).ListCount
  If OP(Index).List(i) = OP(Index).Text Then
  Yes = True
  Exit For
  End If
Next i
    If Yes = False Then OP(Index).Text = ""
End Sub
