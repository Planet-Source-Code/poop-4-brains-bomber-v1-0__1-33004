VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bomber"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8985
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   599
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   4
      Left            =   9840
      Picture         =   "frmGame.frx":030A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   26
      Top             =   3840
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   4
      Left            =   9840
      Picture         =   "frmGame.frx":7E94
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   25
      Top             =   3960
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   3
      Left            =   11520
      Picture         =   "frmGame.frx":FA1E
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   3
      Left            =   11520
      Picture         =   "frmGame.frx":10050
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   9840
      Picture         =   "frmGame.frx":10682
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   9840
      Picture         =   "frmGame.frx":1820C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   9840
      Picture         =   "frmGame.frx":1FD96
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   9840
      Picture         =   "frmGame.frx":21548
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   0
      Left            =   9840
      Picture         =   "frmGame.frx":22CFA
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   0
      Left            =   9840
      Picture         =   "frmGame.frx":2C01C
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox eSubm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9960
      Picture         =   "frmGame.frx":3533E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eSub 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9840
      Picture         =   "frmGame.frx":376A8
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eShipm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9960
      Picture         =   "frmGame.frx":39A12
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eShip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9840
      Picture         =   "frmGame.frx":3BD7C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox bombm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   9960
      Picture         =   "frmGame.frx":3E0E6
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox bomb 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   9840
      Picture         =   "frmGame.frx":3E268
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox eplanem 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   10320
      Picture         =   "frmGame.frx":3E3EA
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eplane 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   10200
      Picture         =   "frmGame.frx":40754
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox bulletm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   9960
      Picture         =   "frmGame.frx":42ABE
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox bullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   9840
      Picture         =   "frmGame.frx":42B50
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox playm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9960
      Picture         =   "frmGame.frx":42BE2
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox play 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9840
      Picture         =   "frmGame.frx":44F4C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   0
      Picture         =   "frmGame.frx":472B6
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   0
      Width           =   9000
      Begin VB.CommandButton cmdAbout 
         BackColor       =   &H000080FF&
         Caption         =   "About"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdTScore 
         BackColor       =   &H000080FF&
         Caption         =   "Top Score"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H000080FF&
         Caption         =   "New Game"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnd 
         BackColor       =   &H000080FF&
         Caption         =   "Exit"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1200
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAbout_Click()
MsgBox ("Bomber v1.0 developed by Kevin Fleet" & vbCrLf & "Copyright(R) 2002 KevCom"), vbOKOnly, "About"
End Sub

Private Sub cmdEnd_Click()
End
End Sub

Private Sub cmdNew_Click()
Randomize
NewGame
End Sub

Function MainLoop()
Dim C As Long, AddS, AddSh, AddPl, I

Do
If C > 5000 Then
C = 0

If Running = True Then

If AddS > 15 Then
AddS = 0
MakeSub
Else
AddS = AddS + 1
End If

If AddPl > 10 Then
AddPl = 0
MakePlane
Else
AddPl = AddPl + 1
End If

If AddSh > 15 Then
AddSh = 0
MakeShip
Else
AddSh = AddSh + 1
End If

CheckPlayerInput
MoveShots
MoveEnemys
MovePlayer
RunExplo

board.Cls

For I = 1 To 20
If PShot(I).Act = True Then
E.DrawObj board.hdc, PShot(I).X, PShot(I).Y, 5, 5, bulletm.hdc, 0, 0, Mask
E.DrawObj board.hdc, PShot(I).X, PShot(I).Y, 5, 5, bullet.hdc, 0, 0, sprite
End If

If PlShot(I).Act = True Then
E.DrawObj board.hdc, PlShot(I).X, PlShot(I).Y, 5, 5, bulletm.hdc, 0, 0, Mask
E.DrawObj board.hdc, PlShot(I).X, PlShot(I).Y, 5, 5, bullet.hdc, 0, 0, sprite
End If

If ShShot(I).Act = True Then
E.DrawObj board.hdc, ShShot(I).X, ShShot(I).Y, 5, 5, bulletm.hdc, 0, 0, Mask
E.DrawObj board.hdc, ShShot(I).X, ShShot(I).Y, 5, 5, bullet.hdc, 0, 0, sprite
End If

If I < 11 Then
If B(I).Act = True Then
E.DrawObj board.hdc, B(I).X, B(I).Y, 10, 10, bombm.hdc, 0, 0, Mask
E.DrawObj board.hdc, B(I).X, B(I).Y, 10, 10, bomb.hdc, 0, 0, sprite
End If
End If
Next I

For I = 1 To 10
If S(I).Act = True Then
E.DrawObj board.hdc, S(I).X, S(I).Y, 50, 30, eSubm.hdc, S(I).Dir * 50, 0, Mask
E.DrawObj board.hdc, S(I).X, S(I).Y, 50, 30, eSub.hdc, S(I).Dir * 50, 0, sprite
End If

If P(I).Act = True Then
E.DrawObj board.hdc, P(I).X, P(I).Y, 50, 30, eplanem.hdc, P(I).Dir * 50, 0, Mask
E.DrawObj board.hdc, P(I).X, P(I).Y, 50, 30, eplane.hdc, P(I).Dir * 50, 0, sprite
End If

If Sh(I).Act = True Then
E.DrawObj board.hdc, Sh(I).X, Sh(I).Y, 50, 30, eShipm.hdc, Sh(I).Dir * 50, 0, Mask
E.DrawObj board.hdc, Sh(I).X, Sh(I).Y, 50, 30, eShip.hdc, Sh(I).Dir * 50, 0, sprite
End If
Next I

E.DrawObj board.hdc, Player.X, Player.Y, 50, 30, playm.hdc, Player.Dir * 50, 0, Mask
E.DrawObj board.hdc, Player.X, Player.Y, 50, 30, play.hdc, Player.Dir * 50, 0, sprite

For I = 1 To 20
If Ex(I).Act = True Then
E.DrawObj board.hdc, Ex(I).X, Ex(I).Y, Ex(I).W, Ex(I).H, explom(Ex(I).T).hdc, Ex(I).Frame * Ex(I).W, 0, Mask
E.DrawObj board.hdc, Ex(I).X, Ex(I).Y, Ex(I).W, Ex(I).H, explo(Ex(I).T).hdc, Ex(I).Frame * Ex(I).W, 0, sprite
End If
Next I

board.CurrentX = 10
board.CurrentY = 10
board.Print "Score: " & Player.Score
board.CurrentX = 10
board.CurrentY = 20 + board.TextHeight("|")
board.Print "HP: "
board.Line (15 + board.TextWidth("HP: "), 20 + board.TextHeight("|"))-((15 + board.TextWidth("HP: ")) + 50, (20 + board.TextHeight("|")) + board.TextHeight("|")), vbRed, BF
board.Line (15 + board.TextWidth("HP: "), 20 + board.TextHeight("|"))-((15 + board.TextWidth("HP: ")) + Player.HP, (20 + board.TextHeight("|")) + board.TextHeight("|")), vbGreen, BF

Else
board.Cls

board.FontSize = 46
board.CurrentX = board.ScaleWidth \ 2 - board.TextWidth(Message) \ 2
board.CurrentY = board.ScaleHeight \ 2 - board.TextHeight("|") \ 2
board.Print Message

board.FontSize = 10
End If
Else
C = C + 1
End If

DoEvents
Loop
End Function

Private Sub cmdTScore_Click()
Load frmTopScore
frmTopScore.Show
End Sub

Private Sub Form_Load()
Message = "Bomber"
Me.Visible = True
MainLoop
End Sub

Private Sub Form_Resize()
board.Left = 0
board.Top = 0
End Sub
