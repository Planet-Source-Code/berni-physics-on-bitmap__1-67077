VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Simple physics on bitmap"
   ClientHeight    =   15165
   ClientLeft      =   -105
   ClientTop       =   1440
   ClientWidth     =   23880
   LinkTopic       =   "Form1"
   ScaleHeight     =   1011
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1592
   Begin VB.PictureBox Pic3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   60060
      Left            =   18000
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   60000
      ScaleWidth      =   30000
      TabIndex        =   10
      Top             =   11760
      Visible         =   0   'False
      Width           =   30060
   End
   Begin VB.Timer DebugUpdate 
      Interval        =   30
      Left            =   8280
      Top             =   120
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      DrawWidth       =   20
      Height          =   60060
      Left            =   16320
      Picture         =   "Form1.frx":55236
      ScaleHeight     =   4000
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   2000
      TabIndex        =   9
      Top             =   9240
      Visible         =   0   'False
      Width           =   30060
   End
   Begin VB.Frame Frame2 
      Caption         =   "Speed"
      Height          =   2175
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1215
      Begin VB.CommandButton Command1 
         Caption         =   "Normal"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Slow"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Very Slow"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Stop"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Step"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.Timer ScrollScreen 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7680
      Top             =   120
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Debug"
      ForeColor       =   &H00FFFF00&
      Height          =   4335
      Left            =   20040
      TabIndex        =   1
      Top             =   0
      Width           =   3855
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Debug"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   3975
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Timer Timer1 
      Left            =   7080
      Top             =   120
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      DrawWidth       =   20
      Height          =   60000
      Left            =   1080
      Picture         =   "Form1.frx":5AD2B
      ScaleHeight     =   4000
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   2000
      TabIndex        =   0
      Top             =   2640
      Width           =   30000
      Begin VB.Shape obj 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   375
         Left            =   3360
         Shape           =   3  'Circle
         Top             =   3360
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeyLeft, KeyRight, KeyShift As Boolean
Dim OSpeed As Integer
Dim XKin, YKin As Double
Dim Xpos, Ypos As Double
Dim DrawOn As Boolean
Dim LastX, LastY As Single
Dim tempA, tempB
Dim Roling As Boolean
Dim FormWidth, FormHeight As Integer
Dim ScroolB As Boolean
Dim ScroolX, ScroolY As Integer
'Basic Parametters
'-------------------------
Const Bounce As Double = 3 'Boncynes of the ball. smaller the number the more it bounces
Const AirResistance As Double = 0.999 'How fast the ball decelerates. 1 is frictionles ,0 is total stop
Const GroundFriction As Double = 0.999 'How fast the ball decelerates when on the ground. 1 is frictionles,0 is total stop
Const GravityStrenth As Double = 1 'How fast the ball falls, the biger the numbaer the faster it falls
'-------------------------

'The Speed controls
Private Sub Command1_Click()
Timer1.Interval = 30
ScrollScreen.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Interval = 100
ScrollScreen.Enabled = True
End Sub

Private Sub Command3_Click()
Timer1.Interval = 200
ScrollScreen.Enabled = True
End Sub

Private Sub Command4_Click()
Timer1.Interval = 0
ScrollScreen.Enabled = False
End Sub

Private Sub Command5_Click()
Timer1_Timer
End Sub

Private Sub DebugUpdate_Timer()
'DEBUG
Label1.Caption = "X Kinetic :" & vbNewLine & XKin & vbNewLine & CharBarPN(XKin) & vbNewLine _
& "Y Kinetic :" & vbNewLine & YKin & vbNewLine & CharBarPN(YKin) & vbNewLine _
& "X Position :" & vbNewLine & Xpos & vbNewLine _
& "Y Position :" & vbNewLine & Ypos & vbNewLine _
& "Hill Right :" & vbNewLine & tempA & vbNewLine & CharBar(tempA) & vbNewLine _
& "Hill Left :" & vbNewLine & tempB & vbNewLine & CharBar(tempB) & vbNewLine
End Sub

Private Sub Form_Resize() ' Ajdust evrything when form resized
Frame1.Left = Form1.Width / 15 - Frame1.Width - 10
FormWidth = Form1.Width / 15
FormHeight = Form1.Height / 15
End Sub


'Left Right Buttns
Private Sub Pic2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 16 Then KeyShift = True
If KeyCode = 37 Then KeyLeft = True
If KeyCode = 39 Then KeyRight = True
End Sub

Private Sub Pic2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 16 Then KeyShift = False
If KeyCode = 37 Then KeyLeft = False
If KeyCode = 39 Then KeyRight = False
End Sub

Private Sub Form_Load() 'Start up this puppy
LoadBitmap
Xpos = obj.Left
Ypos = obj.Top
End Sub

Private Sub Pic2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'Stuff to make mouse interaction work
If Button = 1 Then 'Check wich button on mouse pushed
    DrawOn = True
ElseIf Button = 4 Then
    ScroolB = True
    ScroolX = X
    ScroolY = Y
Else
    Xpos = X
    Ypos = Y
    XKin = 0
    YKin = 0
End If
End Sub

Private Sub Pic2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If DrawOn = True Then 'DrawWalls
    If KeyShift = False Then
        Pic1.Line (LastX, LastY)-(X, Y), RGB(0, 0, 0)
        Pic2.Line (LastX, LastY)-(X, Y), RGB(100, 100, 0)
    Else
        Pic1.Line (LastX, LastY)-(X, Y), vbWhite
        Pic2.PaintPicture Pic3, X - 13, Y - 13, 25, 25, X - 13, Y - 13, 25, 25
    End If
End If

If ScroolB = True Then 'Scrool
Pic2.Left = Pic2.Left - ScroolX + X
Pic2.Top = Pic2.Top - ScroolY + Y
ScrollScreen.Enabled = False
End If
LastX = X 'Used for drawing walls
LastY = Y
End Sub

Private Sub Pic2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Timer1.Interval <> 0 Then ScrollScreen.Enabled = True
ScroolB = False
DrawOn = False
LoadBitmap
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub ScrollScreen_Timer()
On Error Resume Next
If obj.Left + Pic2.Left < FormWidth / 5 Then Pic2.Left = Pic2.Left - (obj.Left + Pic2.Left - FormWidth / 5) / (FormWidth \ 100)
If obj.Left + Pic2.Left > FormWidth / 1.2 Then Pic2.Left = Pic2.Left - (obj.Left + Pic2.Left - FormWidth / 1.2) / (FormWidth \ 100)
If obj.Top + Pic2.Top < FormHeight / 5 Then Pic2.Top = Pic2.Top - (obj.Top + Pic2.Top - FormHeight / 5) / (FormHeight \ 100)
If obj.Top + Pic2.Top > FormHeight / 1.5 Then Pic2.Top = Pic2.Top - (obj.Top + Pic2.Top - FormHeight / 1.5) / (FormHeight \ 100)
End Sub

Private Sub Timer1_Timer()
GoOverGround 'Make shure ball didnt go trugh the ground
'Hill roll
tempA = GroundHigh(Xpos - 5, Ypos + 27)
tempB = GroundHigh(Xpos + 30, Ypos + 27)
XKin = XKin + tempA / 10 * (YKin / 10)
XKin = XKin - tempB / 10 * (YKin / 10)

'X axsis colision detection
If GroundCol(CSng(Xpos), CSng(Ypos) - 10) = True And XKin < 0 Then
    XKin = Abs(XKin) / 5
End If
If GroundCol(Xpos + 25, Ypos - 10) = True And XKin > 0 Then
    XKin = 0 - Abs(XKin) / Bounce
End If

'Y axsis colision detection
If GroundCol(Xpos + 13, Ypos + 28) = True Then
    XKin = XKin - (tempB * YKin) / 50
    XKin = XKin + (tempA * YKin) / 50
    YKin = 0 - Abs(YKin) / Bounce
    GoOverGround
End If
If GroundCol(Xpos + 13, Ypos) = True Then
    XKin = XKin - (tempB * YKin) / 50
    XKin = XKin + (tempA * YKin) / 50
    YKin = Abs(YKin) / Bounce
    GoOverGround
End If
'Gravity
If GroundCol(Xpos + 13, Ypos + 23) = False Then
    YKin = YKin + GravityStrenth
End If






'Direction keys
If KeyLeft = True Then
    Xpos = Xpos - 10
    XKin = 0
End If
If KeyRight = True Then
    Xpos = Xpos + 10
    XKin = 0
End If




'AirResistance
XKin = XKin * AirResistance
YKin = YKin * AirResistance
'GroundRessitance
If GroundCol(Xpos + 13, Ypos + 28) Then XKin = XKin * GroundFriction


'Update position

Xpos = Xpos + XKin
Ypos = Ypos + YKin
'MoveObject

obj.Left = Xpos
obj.Top = Ypos


End Sub


Private Function CharBar(val) As String
For i = 0 To val
CharBar = CharBar & "I"
Next i
For i = 0 To 50 - val
CharBar = CharBar & "."
Next i
End Function


Private Function CharBarPN(val) As String
CharBarPN = CharBarPN & "["

If val > 25 Then val = 25
If val < -25 Then val = -25
If val < 0 Then
    For i = 0 To 25 - Abs(val)
        CharBarPN = CharBarPN & "."
    Next i
    For i = 0 To Abs(val)
        CharBarPN = CharBarPN & "I"
    Next i
    CharBarPN = CharBarPN & ".........................."
Else
    CharBarPN = CharBarPN & ".........................."
    For i = 0 To val
        CharBarPN = CharBarPN & "I"
    Next i
    For i = 0 To 25 - val
        CharBarPN = CharBarPN & "."
    Next i

End If
CharBarPN = CharBarPN & "]"
End Function


Private Function GroundHigh(X As Single, Y As Single) As Integer
Do Until GroundCol(X, Y - i) = False
If i = 100 Then Exit Do
    i = i + 1
Loop
If i < 100 Then GroundHigh = i
End Function

Private Sub GoOverGround()
Dim i As Integer
Do Until GroundCol(Xpos + 13, Ypos + 23 - i) = False
If i = 100 Then Exit Do
    i = i + 1
Loop
If i < 100 Then Ypos = Ypos - i

End Sub
