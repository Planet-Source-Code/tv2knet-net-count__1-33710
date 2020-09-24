Attribute VB_Name = "modCBitBlit"
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Sub BitBlitX(Picture1 As PictureBox, Picture2 As PictureBox, xx As Long, yy As Long, Color As Long)
' A slow but good BitBlit function
On Error Resume Next
Dim c As Long
Dim K As Long
Dim X As Long
Dim Y As Long


For X = 0 To Picture1.ScaleWidth - 1 ' Excluding the border

For Y = 0 To Picture1.ScaleHeight - 1 ' Excluding the border

c = GetPixel(Picture1.hdc, X, Y) ' Gets color of pixel on Picture1
If c <> Color Then ' For making transparant
SetPixel Picture2.hdc, X + xx, Y + yy, c ' Set pixel on Picture2
End If

Next
'Picture2.Refresh
Next

'Picture2.Refresh
DoEvents



End Sub
