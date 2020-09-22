Attribute VB_Name = "cbuttons"
Dim i As Integer, j As Integer, add As Integer
Public Function cbutton()
add = 96
Form1.Picture8.PaintPicture Form1.Picture3.Picture, 0, 0, 114, 18, 16, 88, 114, 18
Form1.Picture8.PaintPicture Form1.Picture3.Picture, 0, 18, 114, 18, 16, 88, 114, 18
Form1.Picture8.PaintPicture Form1.Picture3.Picture, 114, 0, 22, 16, 136, 89, 22, 16
Form1.Picture8.PaintPicture Form1.Picture3.Picture, 114, 16, 22, 16, 136, 89, 22, 16
Form1.Picture8.Picture = Form1.Picture8.Image

For i = 0 To Form1.Picture7.ScaleWidth
For j = 0 To Form1.Picture7.ScaleHeight
    If Form1.Picture7.Point(i, j) = RGB(255, 255, 255) Then
        Form1.Picture8.PSet (i, j), shiftcolor(Form1.Picture8.Point(i, j), add)
    ElseIf Form1.Picture7.Point(i, j) = RGB(0, 0, 0) Then
        Form1.Picture8.PSet (i, j), shiftcolor(Form1.Picture8.Point(i, j), -add)
    End If
Next j
Next i

End Function
