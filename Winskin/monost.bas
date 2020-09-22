Attribute VB_Name = "monost"
Dim i As Integer, j As Integer, add As Integer
Public Function mono()
add = 96
Form1.Picture14.PaintPicture Form1.Picture3.Picture, 0, 0, 29, 12, 239, 41, 29, 12
Form1.Picture14.PaintPicture Form1.Picture3.Picture, 29, 0, 29, 12, 212, 41, 29, 12
Form1.Picture14.PaintPicture Form1.Picture3.Picture, 0, 12, 29, 12, 239, 41, 29, 12
Form1.Picture14.PaintPicture Form1.Picture3.Picture, 29, 12, 29, 12, 212, 41, 29, 12

For i = 0 To Form1.Picture13.ScaleWidth
For j = 0 To Form1.Picture13.ScaleHeight
    If Form1.Picture13.Point(i, j) = RGB(255, 255, 255) Then
        Form1.Picture14.PSet (i, j), shiftcolor(Form1.Picture14.Point(i, j), add)
    ElseIf Form1.Picture13.Point(i, j) = RGB(0, 0, 0) Then
        Form1.Picture14.PSet (i, j), shiftcolor(Form1.Picture14.Point(i, j), -add)
    End If
Next j
Next i

End Function

