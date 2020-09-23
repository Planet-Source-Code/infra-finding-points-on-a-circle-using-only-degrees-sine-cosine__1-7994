Attribute VB_Name = "Mod"
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

Global CM(9) As New CS, TP(9) As Integer, DC(9) As Boolean
'CM = Circle Maker
'TP = Target Position
'DC = Dead Circle
Global Const CA As Integer = 75, DA As Integer = 50
'CA = Change Amount
'DA = Direction Addition

'Declare all API
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Set the Rect type
Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'More API and globals
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRECT As Rect) As Long

Global MouseX As Single
Global MouseY As Single
Global Mouse As New CMouse

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public Const HWND_TOP = 0

Public Const WS_CHILD = &H40000000
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)

Public Const rmConfigure = 1
Public Const rmScreenSaver = 2
Public Const rmPreview = 3
Public RunMode As Integer

Private Const APP_NAME = "Circle ScreenSaver (Cyclone)"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Sub NotOntop(FormName As Form) 'Make form not ontop
Call SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub

Public Sub Ontop(FormName As Form) 'Make form ontop
Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub

Private Sub CheckShouldRun() 'Check if there is already a previous instance of the screensaver, and if so, destroy this one
  If Not App.PrevInstance Then Exit Sub
  If FindWindow(vbNullString, APP_NAME) Then End
MainFrm.Caption = APP_NAME
End Sub

Public Sub Main() 'Read command line and see if this a configuration execution or a screensaver execution, self explanitory
Dim args As String
Dim preview_hwnd As Long
Dim preview_rect As Rect
Dim window_style As Long
  args = UCase$(Trim$(Command$))
    Select Case Mid$(args, 1, 2)
        Case "/C"
            RunMode = rmConfigure
        Case "", "/S"
            RunMode = rmScreenSaver
        Case "/P"
            RunMode = rmPreview
        Case Else
            RunMode = rmScreenSaver
    End Select

    Select Case RunMode
        Case rmConfigure
            Config.Show
            Exit Sub
        
        Case rmScreenSaver
            Load MainFrm
            CheckShouldRun
            MainFrm.Show
            ShowCursor False
            Exit Sub
    End Select
End Sub

Private Function GetHwndFromCommand(ByVal args As String) As Long 'Gets the Hwnd from Command$, self explanitory
Dim argslen As Integer
Dim i As Integer
Dim ch As String
    args = Trim$(args)
    argslen = Len(args)
    For i = argslen To 1 Step -1
        ch = Mid$(args, i, 1)
        If ch < "0" Or ch > "9" Then Exit For
    Next i

    GetHwndFromCommand = CLng(Mid$(args, i + 1))
End Function

Public Sub HideMouse() 'Hide mouse cursor, self explanitory
On Error GoTo error
ShowCursor (bShow = False)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub ShowMouse() 'Show mouse cursor, self explanitory
On Error GoTo error
ShowCursor (bShow = True)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub
