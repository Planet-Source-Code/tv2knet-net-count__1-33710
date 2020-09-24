VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tel LvL <8 V0.1 Free Version"
   ClientHeight    =   6210
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   566
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.PictureBox imgLogo 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1575
         Left            =   6420
         Picture         =   "frmMain.frx":0000
         ScaleHeight     =   1515
         ScaleWidth      =   1500
         TabIndex        =   25
         Top             =   300
         Width           =   1560
      End
      Begin VB.CommandButton cmdStuck 
         Caption         =   "Oh Oh"
         Height          =   495
         Left            =   6000
         TabIndex        =   23
         ToolTipText     =   "Als je vast zit // If you're stuck"
         Top             =   5280
         Width           =   2295
      End
      Begin VB.CommandButton cmdStop 
         Enabled         =   0   'False
         Height          =   495
         Left            =   6000
         TabIndex        =   22
         Top             =   4620
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame Frame4 
         Caption         =   "Status "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   6000
         TabIndex        =   17
         Top             =   2460
         Width           =   2295
         Begin VB.Label lbGN 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Got Number:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   420
            Width           =   900
         End
         Begin VB.Label lblNN 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Need Number:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label lblScore 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Score: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   600
         End
         Begin VB.Label lblLifes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Levens: 3"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   1140
            Width           =   705
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Numbers"
         Height          =   2595
         Left            =   180
         TabIndex        =   3
         Top             =   3360
         Visible         =   0   'False
         Width           =   2295
         Begin VB.PictureBox Numb 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   540
            Index           =   10
            Left            =   780
            Picture         =   "frmMain.frx":2307
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   13
            Top             =   1920
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.PictureBox Numb 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   540
            Index           =   9
            Left            =   180
            Picture         =   "frmMain.frx":2F49
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   12
            Top             =   1380
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.PictureBox Numb 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   540
            Index           =   8
            Left            =   780
            Picture         =   "frmMain.frx":3B8B
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   11
            Top             =   1380
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.PictureBox Numb 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   540
            Index           =   7
            Left            =   1380
            Picture         =   "frmMain.frx":47CD
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   10
            Top             =   1380
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.PictureBox Numb 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   540
            Index           =   1
            Left            =   180
            Picture         =   "frmMain.frx":540F
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   9
            Top             =   240
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.PictureBox Numb 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   540
            Index           =   2
            Left            =   780
            Picture         =   "frmMain.frx":6051
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   8
            Top             =   240
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.PictureBox Numb 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   540
            Index           =   3
            Left            =   1380
            Picture         =   "frmMain.frx":6C93
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   7
            Top             =   240
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.PictureBox Numb 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   540
            Index           =   4
            Left            =   1380
            Picture         =   "frmMain.frx":78D5
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   6
            Top             =   840
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.PictureBox Numb 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   540
            Index           =   5
            Left            =   780
            Picture         =   "frmMain.frx":8517
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   5
            Top             =   840
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.PictureBox Numb 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   540
            Index           =   6
            Left            =   180
            Picture         =   "frmMain.frx":9159
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   4
            Top             =   840
            Visible         =   0   'False
            Width           =   540
         End
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   3420
         Top             =   2640
      End
      Begin VB.CommandButton cmdStart 
         Default         =   -1  'True
         Height          =   495
         Left            =   6000
         TabIndex        =   1
         Top             =   3960
         Width           =   2295
      End
      Begin VB.PictureBox PlayF 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   5775
         Left            =   120
         ScaleHeight     =   381
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   385
         TabIndex        =   2
         Top             =   240
         Width           =   5835
         Begin VB.Timer TmK 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   3720
            Top             =   2400
         End
         Begin VB.Frame Frame3 
            Caption         =   "Fields and Items"
            Height          =   2595
            Left            =   3480
            TabIndex        =   15
            Top             =   3120
            Visible         =   0   'False
            Width           =   2295
            Begin VB.PictureBox P1 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               Height          =   540
               Left            =   1140
               Picture         =   "frmMain.frx":9D9B
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   34
               Top             =   1560
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.PictureBox P2 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               Height          =   540
               Left            =   600
               Picture         =   "frmMain.frx":A9DD
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   33
               Top             =   300
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.PictureBox P3 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               Height          =   540
               Left            =   1140
               Picture         =   "frmMain.frx":B61F
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   32
               Top             =   300
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.PictureBox P4 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               Height          =   540
               Left            =   60
               Picture         =   "frmMain.frx":C261
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   31
               Top             =   900
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.PictureBox P5 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               Height          =   540
               Left            =   60
               Picture         =   "frmMain.frx":CEA3
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   30
               Top             =   300
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.PictureBox P6 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               Height          =   540
               Left            =   600
               Picture         =   "frmMain.frx":DAE5
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   29
               Top             =   900
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.PictureBox P7 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               Height          =   540
               Left            =   1140
               Picture         =   "frmMain.frx":E727
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   28
               Top             =   900
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.PictureBox Numb 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               Height          =   555
               Index           =   0
               Left            =   60
               ScaleHeight     =   33
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   33
               TabIndex        =   27
               Top             =   1560
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.PictureBox W2 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00E78D58&
               Height          =   555
               Left            =   600
               ScaleHeight     =   33
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   33
               TabIndex        =   26
               Top             =   1560
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.PictureBox PB 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               Height          =   540
               Left            =   1740
               Picture         =   "frmMain.frx":F369
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   24
               Top             =   240
               Visible         =   0   'False
               Width           =   540
            End
         End
         Begin VB.PictureBox LoadB 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   0
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   383
            TabIndex        =   14
            Top             =   2340
            Visible         =   0   'False
            Width           =   5775
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.tvirusonline.be"
         Height          =   255
         Left            =   5940
         TabIndex        =   16
         Top             =   1980
         Width           =   2415
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuChoose 
         Caption         =   "Kies Level"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuLang 
      Caption         =   "#mnuLang"
      Begin VB.Menu mnuLangEng 
         Caption         =   "English"
      End
      Begin VB.Menu mnuLangDutch 
         Caption         =   "Nederland"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "Over.."
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Closing As Boolean




Private Sub cmdStart_Click()
PlayF.SetFocus
'cmdStop.Enabled = True
cmdStart.Enabled = False

LL = 3
NN = 0
GN = 0
LoadB.Visible = True

LoadMap "levels", "01.lvl"
InGame = True
Timer2.Enabled = True
CLev = "01"
TmK.Enabled = True
Closing = False

End Sub

Private Sub cmdStop_Click()
cmdStop.Enabled = False
cmdStart.Enabled = True
InGame = False
Timer2.Enabled = False
DoEvents
CallEnd

MsgBox "Score:" + Str(Score) + vbCrLf + LBar + Str(LL), vbInformation, Tel

End Sub

Private Sub cmdStuck_Click()

If InGame = True Then
LL = LL - 1
LoadB.Visible = True
Timer2.Enabled = False
TmK.Enabled = flase
LoadMap "levels", "01.lvl"
InGame = True
Timer2.Enabled = True
If LL = -1 Then
InGame = False
Timer2.Enabled = False
MsgBox EndG, vbInformation, Tel
Exit Sub
End If
End If
End Sub

Private Sub Form_Load()
Closing = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Closing <> True Then Cancel = 1 Else End

End Sub

Private Sub Image1_Click()

End Sub

Private Sub mnuAbout_Click()
MsgBox "Programma gemaakt door: T-Virus Creations" + vbCrLf + "Program Created by T-Virus Creations" + vbCrLf + vbCrLf + "http://www.tvirusonline.be" + vbCrLf + "tvirus4ever@yahoo.co.uk", vbInformation, "Over..."

End Sub

Private Sub mnuChoose_Click()
On Error GoTo 10
Dim X As Integer
X = InputBox(ChooseT, Tel, RC(Str(Val(CLev) + 1), " "))
CLev = X
NN = -1
GN = 0
Score = 0
LL = LL + 1
InGame = False
CLev = RC(Str(CLev), " ")
LoadB.Visible = True
LoadMap "levels", CLev + ".lvl"
InGame = True
Timer2.Enabled = True
TmK.Enabled = True
PlayF.SetFocus
Closing = False

Exit Sub
10

End Sub

Private Sub mnuExit_Click()
Closing = True

Unload Me
End
End Sub

Private Sub mnuLangDutch_Click()
LoadLang 0, 1
LoadLang 0, 1
End Sub

Private Sub mnuLangEng_Click()
LoadLang 1, 1
LoadLang 1, 1
End Sub

Private Sub PlayF_KeyUp(KeyCode As Integer, Shift As Integer) ' Set Player Char
Dim OPX As Long
Dim OPY As Long
Dim tl As Long

If InGame = True Then

If InSet = False Then

InSet = True
OPX = PX
OPY = PY
SetOld PX, PY

Select Case KeyCode ' What Key Was Pressed
Case vbKeyLeft
PX = PX - 32 ' Set Position

Case vbKeyRight
PX = PX + 32 ' Set Position

Case vbKeyUp
PY = PY - 32 ' Set Position

Case vbKeyDown
PY = PY + 32 ' Set Position
End Select


If CheckField(PX, PY) = 1 Then

SetOld PX, PY
BitBlitX P4, PlayF, OPX, OPY, vbBlack

If OPX = PX And OPY = PY Then
BitBlitX P3, PlayF, OPX, OPY, vbBlack ' Set Grass
End If

PX = OPX
PY = OPY

InSet = False
End If

PlayF.Refresh
If NN = GN Then
MsgBox WinT, vbInformation, Tel
NN = -1
GN = 0
Score = Score + 10 ' Add 10 Points
LL = LL + 1 ' Add One Life
InGame = False
InSet = False
CallEnd ' Calling End Effect
DoEvents ' Don't Crash!

tl = Val(CLev)
tl = tl + 1
CLev = RC(Str(tl), " ") ' Get Level Number In String

If Len(CLev) <> 1 Then CLev = "0" + CLev
LoadB.Visible = True ' Show LoadBar

If LL > 3 Then LL = LL Else LL = 3
LoadMap "levels", CLev + ".lvl" ' Load Map
End If

InGame = True
Timer2.Enabled = True
InSet = False
Exit Sub
End If

If GN > NN And NN <> -1 Then
MsgBox MuchT, vbInformation, Tel
NN = -1
GN = 0
LL = LL - 1
InGame = False
CallEnd
InSet = False

If LL = -1 Then
InGame = False
Timer2.Enabled = False
MsgBox EndG, vbInformation, Tel ' Game Over
cmdStart.Enabled = True
InSet = False
Exit Sub

End If
Score = Score - 50 ' You lost some points :(

LoadB.Visible = True
LoadMap "levels", CLev + ".lvl"
InGame = True
Timer2.Enabled = True
DoEvents
End If

End If
End Sub



Private Sub Timer2_Timer() ' Water Effect Timer
TmK.Enabled = True
SetWater ' Call SetWater Function
End Sub

Private Sub TmK_Timer() ' Check Status Timer
If GN > NN And NN <> -1 Then
MsgBox MuchT, vbInformation, Tel
NN = -1
GN = 0
LL = LL - 1
CallEnd

If LL = -1 Then
InGame = False
Timer2.Enabled = False
MsgBox EndG, vbInformation, Tel
cmdStart.Enabled = True

Exit Sub
End If
Score = Score - 100
InGame = False
LoadB.Visible = True

LoadMap "levels", CLev + ".lvl"
InGame = True
Timer2.Enabled = True


DoEvents

End If
TmK.Enabled = Timer2.Enabled

End Sub

Public Sub SetWater() ' Water Effect
InSet = True
If TWave = 0 Then TWave = 1 Else TWave = 0
On Error Resume Next
If InGame = True Then
Dim xx As Long
Dim yy As Long
Dim stat  As Integer
xx = 0
yy = 0
For i = 1 To Len(DD) Step 2
xx = xx + 32
stat = stat + 1
If Mid$(DD, i, 2) = "WW" Then
If TWave = 0 Then
BitBlitX P1, PlayF, xx - 32, yy, vbBlack
Else
BitBlitX W2, PlayF, xx - 32, yy, vbBlack
End If

End If


If stat = 12 Then
stat = 0
yy = yy + 32
xx = 0

End If

Next
End If
PlayF.Refresh
InSet = False
End Sub
