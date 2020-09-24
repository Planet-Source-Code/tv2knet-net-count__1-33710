Attribute VB_Name = "modGame"
Public xx As Long ' Holds Position
Public yy As Long ' Holds Position
Public PX As Long ' Holds Position
Public PY As Long ' Holds Position
Public Score As Long ' Holds the score
Public NN As Integer ' The number you need to collect
Public GN As Integer ' The number you collected
Public DD As String ' MAP Data Container
Public InGame As Boolean ' If playing = True
Public TWave As Integer ' Current Water Effect
Public Const NoMore As String = "No more levels... // Geen nieuwe levels meer..." ' If you ran out of levels
Public InSet As Boolean


Public Sub LoadGame(Data As String) ' This processes the Loaded Map
' Declerations -----
Dim LB As Long
Dim CB As Long
Dim stat As Integer
' ------------------
With frmMain
.LoadB.Cls ' Clear progressbar
.PlayF.Cls ' Clear playfield


xx = 0
yy = 0
Dim D As String
D = Replace$(Data, vbCrLf, "") ' To get one row
DD = D ' Set DD Container = D
For i = 1 To Len(D) Step 2 ' To get tile type (Water, Wall,...)
xx = xx + 32 ' To get X position
stat = stat + 1 ' Used later...

'For CB = 0 To 32 * 11
'
'Next
Select Case Mid$(D, i, 2) ' Select One Place
Case "XX" ' Case Wall
BitBlitX .P5, .PlayF, xx - 32, yy, vbRed ' Draw Wall


Case "HH" ' Case Heart
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass
BitBlitX .P6, .PlayF, xx - 32, yy, vbBlack ' Draw Heart (Using Black as Trans. Color)


Case "01" ' Case 01
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass
BitBlitX .Numb(1), .PlayF, xx - 32, yy, vbBlack ' Draw Number


Case "02" ' Case 02
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass
BitBlitX .Numb(2), .PlayF, xx - 32, yy, vbBlack ' Draw Number


Case "03" ' Case 03
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass
BitBlitX .Numb(3), .PlayF, xx - 32, yy, vbBlack ' Draw Number

Case "04" ' Case 04
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass
BitBlitX .Numb(4), .PlayF, xx - 32, yy, vbBlack ' Draw Number

Case "05" ' Case 05
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass
BitBlitX .Numb(5), .PlayF, xx - 32, yy, vbBlack ' Draw Number

Case "06" ' Case 06
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass
BitBlitX .Numb(6), .PlayF, xx - 32, yy, vbBlack ' Draw Number

Case "07" ' Case 07
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass
BitBlitX .Numb(7), .PlayF, xx - 32, yy, vbBlack ' Draw Number

Case "08" ' Case 08
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass
BitBlitX .Numb(8), .PlayF, xx - 32, yy, vbBlack ' Draw Number

Case "09" ' Case 09
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass
BitBlitX .Numb(9), .PlayF, xx - 32, yy, vbBlack ' Draw Number

Case "10" ' Case 10
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass
BitBlitX .Numb(10), .PlayF, xx - 32, yy, vbBlack ' Draw Number


Case "WW" ' Case Water
BitBlitX .P1, .PlayF, xx - 32, yy, vbRed ' Draw Water


Case "GG" ' Case Grass
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass


Case "DD" ' Case Dirt
BitBlitX .P2, .PlayF, xx - 32, yy, vbRed ' Draw Dirt Patch


Case "II" ' Case OTUOW (One Time Use Only Wall)
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass

Case "HI" ' Case Invisible
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass


Case "&&" ' Case Key (Not used in MAPs, yet!)
BitBlitX .P3, .PlayF, xx - 32, yy, vbRed ' Draw Grass
BitBlitX .P7, .PlayF, xx - 32, yy, vbBlack ' Draw Key


Case "PP" ' Case
BitBlitX .P3, .PlayF, xx - 32, yy, vbBlack ' Draw Grass
BitBlitX .P4, .PlayF, xx - 32, yy, vbBlack ' Draw Player
PX = xx - 32 ' Set Player X Position
PY = yy ' Set Player Y Position
Case "" ' Case Nothing :)
End Select

If stat = 12 Then ' Go to next row
stat = 0 ' Reset
yy = yy + 32 ' Add Y Position
xx = 0 ' Reset X position
End If
.PlayF.Refresh ' Refresh Playfield (For Loading Effect)


BitBlitX .PB, .LoadB, yy, 1, vbBlack ' Updates LoadBar
.LoadB.CurrentX = 32 ' Sets X For TextOutPut
.LoadB.CurrentY = 16 ' Sets Y For TextOutPut
.LoadB.ForeColor = RGB(Rnd(255) * 255, Rnd(255) * 255, Rnd(255) * 255) ' Sets Random ForeColor
.LoadB.Print "Loading... // Bezig met laden.." ' Loading Text (English/Dutch)
.LoadB.Refresh ' Shows Progress

Next ' Again :)

.LoadB.Visible = False ' Hide LoadBar
.PlayF.SetFocus ' Set Focus To PlayField, To Catch Keys :)
'.cmdStop.Enabled = True ' Not Used
.cmdStart.Enabled = False ' Disable Start Button In Game, Cause You're Playing


End With

End Sub




Public Function CheckField(X As Long, Y As Long) As Integer
On Error Resume Next
With frmMain
Dim K As Long
Dim p As Long
Dim r As Long
Dim D As String
Dim z As String
D = Replace$(DD, vbCrLf, "") ' Allready done.. For protection
' Calculating Current Position...
K = X / 32
p = Y / 32
r = 12 * (p) + K '+ 1
r = r * 2 + 1
' Calculating Done!

If Mid$(D, r, 2) <> "XX" And Mid$(D, r, 2) <> "WW" And Mid$(D, r, 2) <> "HI" Then
BitBlitX .P4, .PlayF, X, Y, vbBlack ' Draw Player
Select Case Mid$(D, r, 2) ' = "HH" Then
Case "HH" ' Heart = Add 50 To Score
Score = Score + 50 ' Add 50 To Score
SetStatus Score, NN, GN ' Update Score Status
DD = Left$(D, r - 1) + "GG" + Mid$(D, r + 2, Len(D)) ' Replace HH(Heart) with GG(Grass)



Case "01"
GN = GN + 1 ' Add 1 To Collected
SetStatus Score, NN, GN ' Update Score Status
DD = Left$(D, r - 1) + "GG" + Mid$(D, r + 2, Len(D)) ' Replace 01(1) with GG(Grass)

Case "02"
GN = GN + 2 ' Add 2 To Collected
SetStatus Score, NN, GN ' Update Score Status
DD = Left$(D, r - 1) + "GG" + Mid$(D, r + 2, Len(D)) ' Replace 02(2) with GG(Grass)


Case "03"
GN = GN + 3 ' Add 3 To Collected
SetStatus Score, NN, GN ' Update Score Status
DD = Left$(D, r - 1) + "GG" + Mid$(D, r + 2, Len(D)) ' Replace 03(3) with GG(Grass)

Case "04"
GN = GN + 4 ' Add 4 To Collected
SetStatus Score, NN, GN ' Update Score Status
DD = Left$(D, r - 1) + "GG" + Mid$(D, r + 2, Len(D)) ' Replace 04(4) with GG(Grass)

Case "05"
GN = GN + 5 ' Add 5 To Collected
SetStatus Score, NN, GN ' Update Score Status
DD = Left$(D, r - 1) + "GG" + Mid$(D, r + 2, Len(D)) ' Replace 05(5) with GG(Grass)

Case "06"
GN = GN + 6 ' Add 6 To Collected
SetStatus Score, NN, GN ' Update Score Status
DD = Left$(D, r - 1) + "GG" + Mid$(D, r + 2, Len(D)) ' Replace 06(6) with GG(Grass)

Case "07"
GN = GN + 7 ' Add 7 To Collected
SetStatus Score, NN, GN ' Update Score Status
DD = Left$(D, r - 1) + "GG" + Mid$(D, r + 2, Len(D)) ' Replace 07(7) with GG(Grass)

Case "08"
GN = GN + 8 ' Add 8 To Collected
SetStatus Score, NN, GN ' Update Score Status
DD = Left$(D, r - 1) + "GG" + Mid$(D, r + 2, Len(D)) ' Replace 08(8) with GG(Grass)

Case "09"
GN = GN + 9 ' Add 9 To Collected
SetStatus Score, NN, GN ' Update Score Status
DD = Left$(D, r - 1) + "GG" + Mid$(D, r + 2, Len(D)) ' Replace 09(9) with GG(Grass)

Case "10"
GN = GN + 10 ' Add 10 To Collected
SetStatus Score, NN, GN ' Update Score Status
DD = Left$(D, r - 1) + "GG" + Mid$(D, r + 2, Len(D)) ' Replace 10(10) with GG(Grass)

Case "II" ' Case OTIOW
DD = Left$(D, r - 1) + "XX" + Mid$(D, r + 2, Len(D)) ' Replace II(OTUOW) with XX(Wall)

End Select

Else

CheckField = 1 ' Can't Pass
End If

'z = Mid$(D, r, 1)
'z = z
End With
End Function

Public Sub SetOld(X As Long, Y As Long) ' Update Last Field (Else Would Take To Long)
With frmMain
Dim K As Long
Dim p As Long
Dim r As Long
Dim D As String
Dim z As String

D = Replace$(DD, vbCrLf, "") ' Allready done... To be sure!
' Calculating Current Position...

K = X / 32
p = Y / 32
r = 12 * (p) + K '+ 1
r = r * 2 + 1
' Calculating Done!

BitBlitX .P3, .PlayF, X, Y, vbBlack ' Always Draw Grass For A Start

If Mid$(D, r, 2) = "XX" Then
BitBlitX .P5, .PlayF, X, Y, vbRed
End If

If Mid$(D, r, 2) = "WW" Then ' If water

If TWave = 0 Then ' Get Water Wave Style (Effect)
BitBlitX .P1, .PlayF, X, Y, vbBlack ' Draw Style 1
Else
BitBlitX .W2, .PlayF, X, Y, vbBlack ' Draw Style 2
End If

End If

If Mid$(D, r, 2) = "GG" Then ' If Grass
BitBlitX .P3, .PlayF, X, Y, vbBlack ' Draw Grass
End If

If Mid$(D, r, 2) = "PP" Then  ' If Player
BitBlitX .P3, .PlayF, X, Y, vbBlack ' Draw Player
End If

If Mid$(D, r, 2) = "DD" Then ' If Dirt
BitBlitX .P2, .PlayF, X, Y, vbBlack ' Draw Dirt
End If

If Mid$(D, r, 2) = "01" Then ' If number 1
BitBlitX .P3, .PlayF, X, Y, vbRed ' Draw Grass
BitBlitX .Numb(1), .PlayF, X, Y, vbBlack ' Draw Number 1
End If

If Mid$(D, r, 2) = "02" Then ' If number 2
BitBlitX .P3, .PlayF, X, Y, vbRed ' Draw Grass
BitBlitX .Numb(2), .PlayF, X, Y, vbBlack ' Draw Number 2
End If

If Mid$(D, r, 2) = "03" Then ' If number 3
BitBlitX .P3, .PlayF, X, Y, vbRed ' Draw Grass
BitBlitX .Numb(3), .PlayF, X, Y, vbBlack ' Draw Number 3
End If

If Mid$(D, r, 2) = "04" Then ' If number 4
BitBlitX .P3, .PlayF, X, Y, vbRed ' Draw Grass
BitBlitX .Numb(4), .PlayF, X, Y, vbBlack ' Draw Number 4
End If

If Mid$(D, r, 2) = "05" Then ' If number 5
BitBlitX .P3, .PlayF, X, Y, vbRed ' Draw Grass
BitBlitX .Numb(5), .PlayF, X, Y, vbBlack ' Draw Number 5
End If

If Mid$(D, r, 2) = "06" Then ' If number 6
BitBlitX .P3, .PlayF, X, Y, vbRed ' Draw Grass
BitBlitX .Numb(6), .PlayF, X, Y, vbBlack ' Draw Number 6
End If

If Mid$(D, r, 2) = "07" Then ' If number 7
BitBlitX .P3, .PlayF, X, Y, vbRed ' Draw Grass
BitBlitX .Numb(7), .PlayF, X, Y, vbBlack ' Draw Number 7
End If

If Mid$(D, r, 2) = "08" Then ' If number 8
BitBlitX .P3, .PlayF, X, Y, vbRed ' Draw Grass
BitBlitX .Numb(8), .PlayF, X, Y, vbBlack ' Draw Number 8
End If

If Mid$(D, r, 2) = "09" Then ' If number 9
BitBlitX .P3, .PlayF, X, Y, vbRed ' Draw Grass
BitBlitX .Numb(9), .PlayF, X, Y, vbBlack ' Draw Number 9
End If

If Mid$(D, r, 2) = "10" Then ' If number 10
BitBlitX .P3, .PlayF, X, Y, vbRed ' Draw Grass
BitBlitX .Numb(10), .PlayF, X, Y, vbBlack ' Draw Number 10
End If


'z = Mid$(D, 2, 1)
'z = z
End With
End Sub



Public Sub CallEnd() ' End Of Game Effect
With frmMain
Dim X As Long
Dim Y As Long

For X = 0 To 12
For Y = 0 To 12
BitBlitX .P5, .PlayF, X * 32, Y * 32, vbRed ' Draw Wall
.PlayF.Refresh ' Refresh Field

Next
DoEvents
Next
'.cmdStop.Enabled = False ' Not Used
.cmdStart.Enabled = False ' Disable Start Button ' Not Used
End With
End Sub

Public Sub LoadMap(Dir As String, LevelFile As String) ' Open and load a Level
' Declerations -
Dim X As String
Dim K As String
Dim tt As String
' --------------
On Error GoTo 25
' Get Level Name
tt = Left$(LevelFile, Len(LevelFile) - 4)
If Len(tt) <> 2 Then tt = "0" + tt
LevelFile = tt + ".lvl"
' Got Level Name

Open PaT + Dir + "\" + LevelFile For Input As #1 ' Open File

While EOF(1) = False ' Not End Of File, Yet... Add data
Input #1, X ' Get data from file
If K = "" Then K = X Else K = K + vbCrLf + X
Wend

Close #1 ' Close File

' Little Check..
GoTo 20
25
K = ""
20
' Little Check.. Done!

If Len(K) < 1 Then
GoTo 10
End If

X = Left$(K, 4) ' Get Number To Collect
X = Replace(X, vbCrLf, "") ' Remove vbCrLf
NN = Val(X) ' Convert to number
K = Right$(K, Len(K) - 4) ' Get Other Data

LoadGame K ' Process The Game
SetStatus Score, NN, GN ' Update Score Status
Exit Sub
10
MsgBox NoLevel, vbInformation, Tel ' Level Not Found
LoadMap "levels", "01.lvl" ' Load Map 01.lvl
End Sub

Public Sub SetStatus(Scores As Long, NNN As Integer, GGN As Integer) ' Update Score Status
With frmMain
.lblScore.Caption = ScoreT + Str(Scores) ' Score
.lblNN.Caption = Verzamel + Str(NNN) ' Number To Collect
.lbGN = Totaal + Str(GGN) ' Number Collected
.lblLifes = LBar + Str(LL) ' Lifes Left
End With
End Sub


Public Function PaT() As String ' Get Current Path
If Right$(App.Path, 1) <> "\" Then PaT = App.Path + "\" Else PaT = App.Path ' In One Line :)
End Function


Public Function RC(Data As String, Char As String) As String ' Remove Char Function
RC = Replace(Data, Char, "")
End Function
