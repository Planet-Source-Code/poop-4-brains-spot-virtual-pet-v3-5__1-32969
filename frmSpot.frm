VERSION 5.00
Begin VB.Form frmSpot 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Spot"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   Icon            =   "frmSpot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1560
      Picture         =   "frmSpot.frx":0442
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox ps 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1200
      Picture         =   "frmSpot.frx":0934
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox apm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1560
      Picture         =   "frmSpot.frx":0E26
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox aps 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1200
      Picture         =   "frmSpot.frx":1318
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox spotm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   5160
      Picture         =   "frmSpot.frx":180A
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Timer tmrSpot 
      Interval        =   300
      Left            =   2760
      Top             =   3240
   End
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdPresent 
         BackColor       =   &H00C00000&
         Caption         =   "Give Present"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00C00000&
         Caption         =   "Reset Spot"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdHelp 
         BackColor       =   &H00C00000&
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdPlay 
         BackColor       =   &H00C00000&
         Caption         =   "Play"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C00000&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton cmdSleep 
         BackColor       =   &H00C00000&
         Caption         =   "Sleep"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdEat 
         BackColor       =   &H00C00000&
         Caption         =   "Eat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   1500
         Left            =   2040
         Picture         =   "frmSpot.frx":B7AC
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   2
         Top             =   360
         Width           =   3000
         Begin VB.Image present 
            Height          =   300
            Left            =   720
            Picture         =   "frmSpot.frx":1A24E
            Top             =   1080
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Image apple 
            Height          =   300
            Left            =   1560
            Picture         =   "frmSpot.frx":1A740
            Top             =   1080
            Visible         =   0   'False
            Width           =   300
         End
      End
   End
   Begin VB.Timer tmrRun 
      Interval        =   300
      Left            =   2760
      Top             =   2760
   End
   Begin VB.PictureBox spotsrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   3720
      Picture         =   "frmSpot.frx":1AC32
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   1350
   End
End
Attribute VB_Name = "frmSpot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEat_Click()
If Spot.Action <> 0 Then Exit Sub
Spot.Frame = 0
Spot.Action = 2
Spot.LpAction = 3
Spot.X = 77
A.Visible = True
End Sub

Sub cmdExit_Click()
Unload frmPopup
Unload frmHelp
Unload frmAbout
Unload Me
End Sub

Private Sub cmdHelp_Click()
Load frmHelp
frmHelp.Visible = True
End Sub

Private Sub cmdPlay_Click()
If Spot.Action <> 0 Then Exit Sub
Spot.Frame = 0
Spot.Action = 3
Spot.LpAction = 3
End Sub

Private Sub cmdPresent_Click()
If Spot.Action <> 0 Then Exit Sub
Spot.Frame = 0
Box.Visible = True
Spot.Action = 4
Spot.LpAction = 1
Spot.X = present.Left - 30
End Sub

Private Sub cmdReset_Click()
ResetSpot
End Sub

Private Sub cmdSleep_Click()
Select Case cmdSleep.Caption
Case "Sleep"
If Spot.Action <> 0 Then Exit Sub
Spot.Action = 1
Spot.LpAction = 999999999
cmdSleep.Caption = "Wake-up"
Case "Wake-up"
Spot.Action = 0
Spot.LpAction = 0
cmdSleep.Caption = "Sleep"
End Select
End Sub

Function FileExist(path As String) As Boolean
On Error GoTo oops
FileExist = True
Open path For Input As #1
Close #1
Exit Function
oops:
FileExist = False
End Function

Function ResetSpot()
cmdSleep.Caption = "Sleep"
Spot.Action = 0
Spot.Activity = 10
Spot.Alive = True
Spot.Brain = 50
Spot.DoLose = 0
Spot.Frame = 0
Spot.LoseActivity = 0
Spot.LpAction = 0
Spot.Sleep = 50
Spot.SleepTimer = 0
Spot.Stomach = 50
Spot.TimeHungry = 0
Spot.TimeTired = 0
Spot.WalkLength = 5
Spot.XS = 2
Spot.Happy = 50
Spot.X = 10
A.Visible = False
Box.Visible = False
End Function

Function LoadSound(file As String)
Dim B() As Byte
B() = LoadResData(file, "CUSTOM")
Open "C:\WINDOWS\SPOT\" & file For Binary As #1
Put #1, , B()
Close #1
End Function

Private Sub Form_Load()
On Error Resume Next
If UCase(Dir("C:\WINDOWS\SPOT", vbDirectory)) <> "SPOT" Then MkDir ("C:\WINDOWS\SPOT\")

If FileExist("C:\WINDOWS\SPOT\play.wav") = False Then LoadSound "play.wav"
If FileExist("C:\WINDOWS\SPOT\eat.wav") = False Then LoadSound "eat.wav"
If FileExist("C:\WINDOWS\SPOT\spot.dat") = False Then
ResetSpot
Open "C:\WINDOWS\SPOT\spot.dat" For Output As #1
Write #1, Spot.Action, Spot.Activity, Spot.Alive, Spot.Brain, Spot.DoLose, Spot.Frame, Spot.LoseActivity, Spot.LpAction, Spot.LpAction, Spot.Sleep, Spot.SleepTimer, Spot.Stomach, Spot.TimeHungry, cmdSleep.Caption, Spot.X, A.Visible, Box.Visible, Spot.Happy
Close #1
End If

Dim n As String
Open "C:\WINDOWS\SPOT\spot.dat" For Input As #1
Input #1, Spot.Action, Spot.Activity, Spot.Alive, Spot.Brain, Spot.DoLose, Spot.Frame, Spot.LoseActivity, Spot.LpAction, Spot.LpAction, Spot.Sleep, Spot.SleepTimer, Spot.Stomach, Spot.TimeHungry, n, Spot.X, A.Visible, Box.Visible, Spot.Happy
cmdSleep.Caption = n
Close #1

Load frmPopup

Form_MouseMove vbLeftButton, 0, 0, 0
tmrRun_Timer
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Ret&
ReleaseCapture
Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open "C:\WINDOWS\SPOT\spot.dat" For Output As #1
Write #1, Spot.Action, Spot.Activity, Spot.Alive, Spot.Brain, Spot.DoLose, Spot.Frame, Spot.LoseActivity, Spot.LpAction, Spot.LpAction, Spot.Sleep, Spot.SleepTimer, Spot.Stomach, Spot.TimeHungry, cmdSleep.Caption, Spot.X, A.Visible, Box.Visible, Spot.Happy
Close #1

RemoveFromTray
End Sub

Private Sub tmrRun_Timer()
board.Cls
board.ForeColor = vbBlack
If Spot.Alive = True Then board.ForeColor = vbBlue
board.FontBold = True
board.FontSize = 18
board.CurrentX = 10
board.CurrentY = 10
board.Print "Spot"

DrawBar "Sleep:", Spot.Sleep, 0
DrawBar "Stomach:", Spot.Stomach, 1
DrawBar "Brain:", Spot.Brain, 2
DrawBar "Activity:", Spot.Activity, 3
DrawBar "Happiness:", Spot.Happy, 4

pic.Cls
Select Case Spot.Alive
Case True

If A.Visible = True Then
BitBlt pic.hDC, apple.Left, apple.Top, 20, 20, apm.hDC, 0, 0, vbSrcAnd
BitBlt pic.hDC, apple.Left, apple.Top, 20, 20, aps.hDC, 0, 0, vbSrcInvert
End If

If Box.Visible = True Then
BitBlt pic.hDC, present.Left, present.Top, 20, 20, pm.hDC, 0, 0, vbSrcAnd
BitBlt pic.hDC, present.Left, present.Top, 20, 20, ps.hDC, 0, 0, vbSrcInvert
End If

BitBlt pic.hDC, Spot.X, pic.ScaleHeight - 30, 30, 30, spotm.hDC, Spot.Frame * 30, Spot.Action * 30, vbSrcAnd
BitBlt pic.hDC, Spot.X, pic.ScaleHeight - 30, 30, 30, spotsrc.hDC, Spot.Frame * 30, Spot.Action * 30, vbSrcInvert

Case False
pic.DrawWidth = 10
pic.Line (0, 0)-(pic.ScaleWidth, pic.ScaleHeight)
pic.Line (pic.ScaleWidth, 0)-(0, pic.ScaleHeight)
pic.DrawWidth = 1
End Select
End Sub

Function GetY(spt As Long) As Long
Dim bigheight
board.FontSize = 18
bigheight = board.TextHeight("|")
board.FontSize = 8

GetY = 10 + bigheight + 10 + ((5 + board.TextHeight("|")) * spt)
End Function

Function DrawBar(Text As String, Value As Long, spt As Long)
board.CurrentX = 10
board.CurrentY = GetY(spt)
board.Print Text
board.Line (10 + board.TextWidth(Text) + 5, GetY(spt))-(10 + board.TextWidth(Text) + 5 + 50, GetY(spt) + board.TextHeight("|")), vbRed, BF
board.Line (10 + board.TextWidth(Text) + 5, GetY(spt))-(10 + board.TextWidth(Text) + 5 + Value, GetY(spt) + board.TextHeight("|")), vbGreen, BF
End Function

Private Sub tmrSpot_Timer()
If Spot.Alive = True Then DoSpot
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Message As Long
Message = X / Screen.TwipsPerPixelX
 Select Case Message
'Your Choice:
Case WM_RBUTTONUP
'***
Case WM_RBUTTONDOWN
PopupMenu frmPopup.POP_UP
End Select
End Sub
