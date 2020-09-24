Attribute VB_Name = "modLang"
Public CLev As String
Public LL As Integer
'-- Settings Container
Public OK As String
Public Ja As String
Public Nee As String
Public KiesLevel As String
Public Tel As String
Public Spelletjes As String
Public Waarschuwing As String
Public Geregistreerd As String
Public Verzamel As String
Public Totaal As String
Public ScoreT As String
Public StopT As String
Public AboutT As String
Public OverT As String
Public MuchT As String
Public WinT As String
Public EndG As String
Public LBar As String
Public ChooseT As String
Public NoLevel As String
Public Languages As String
'-- End Settings Container

Public Function OpenTXT(Dir As String, File As String) ' Load A Text Data File
Dim X As String
Dim K As String

Open PaT + Dir + "\" + File For Input As #1 ' Open File

While EOF(1) = False ' While Not End Of File (EOF)
Input #1, X ' Get Data
If K = "" Then K = X Else K = K + vbCrLf + X ' Add Data
Wend

Close #1

OpenTXT = K ' Return Loaded Data
End Function
Public Sub SetLang(File As String) ' Update Language
Dim K As String
Dim D() As String
On Error Resume Next
K = OpenTXT("lang", File)
D() = Split(K, vbCrLf)
DoEvents
OK = D(0)
Ja = D(1)
Nee = D(2)
KiesLevel = D(3)
Tel = D(4)
Spelletjes = D(5)
Waarschuwing = D(6)
Geregistreerd = D(7)
Verzamel = D(8)
Totaal = D(9)
ScoreT = D(10)
StopT = D(11)
StartT = D(12)
OverT = D(13)
MuchT = D(14)
WinT = D(15)
EndG = D(16)
LBar = D(17)
ChooseT = D(18)
NoLevel = D(19)
Languages = D(20)
DoEvents
With frmMain
.Caption = Tel + " V0.1"
.cmdStart.Caption = StartT
.cmdStop.Caption = StopT
.mnuChoose.Caption = KiesLevel
.mnuAbout.Caption = OverT
.lblNN.Caption = Verzamel + " 0"
.lbGN.Caption = Totaal + " 0"
.lblScore.Caption = ScoreT + " 0"
.mnuLang.Caption = Languages
End With
With frmSplash
.lblLicenseTo = Geregistreerd
.lblWarning = Waarschuwing
.lblProductName = Tel
.lblCompanyProduct = Spelletjes

End With

End Sub
Public Sub LoadLang(LID As Integer, Optional IGame As Integer = 0) ' Load Language FrontEnd
' Checking Command$ ...
If IGame = 0 Then
If Command$ <> "" Then
LID = Val(Command$)

End If
End If
' Checking Command$ Done!
DoEvents
Select Case LID

Case 0
SetLang "dutch.lng"
Case 1
SetLang "english.lng"

End Select

End Sub
