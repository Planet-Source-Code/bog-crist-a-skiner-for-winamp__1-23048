Attribute VB_Name = "shufrep"
Dim i As Integer, j As Integer, add As Integer
Public Function shuf()
add = 96
Form1.Picture10.PaintPicture Form1.Picture3.Picture, 0, 61, 46, 12, 219, 58, 46, 12
Form1.Picture10.PaintPicture Form1.Picture3.Picture, 0, 73, 46, 12, 219, 58, 46, 12
Form1.Picture10.PaintPicture Form1.Picture3.Picture, 46, 61, 46, 12, 219, 58, 46, 12
Form1.Picture10.PaintPicture Form1.Picture3.Picture, 46, 73, 46, 12, 219, 58, 46, 12

Form1.Picture10.PaintPicture Form1.Picture3.Picture, 0, 0, 28, 15, 210, 89, 28, 15
Form1.Picture10.PaintPicture Form1.Picture3.Picture, 28, 0, 47, 15, 164, 89, 47, 15
Form1.Picture10.PaintPicture Form1.Picture3.Picture, 0, 15, 28, 15, 210, 89, 28, 15
Form1.Picture10.PaintPicture Form1.Picture3.Picture, 28, 15, 47, 15, 164, 89, 47, 15
Form1.Picture10.PaintPicture Form1.Picture3.Picture, 0, 30, 28, 15, 210, 89, 28, 15
Form1.Picture10.PaintPicture Form1.Picture3.Picture, 28, 30, 47, 15, 164, 89, 47, 15
Form1.Picture10.PaintPicture Form1.Picture3.Picture, 0, 45, 28, 15, 210, 89, 28, 15
Form1.Picture10.PaintPicture Form1.Picture3.Picture, 28, 45, 47, 15, 164, 89, 47, 15

For i = 0 To Form1.Picture9.ScaleWidth
For j = 0 To Form1.Picture9.ScaleHeight
    If Form1.Picture9.Point(i, j) = RGB(255, 255, 255) Then
        Form1.Picture10.PSet (i, j), shiftcolor(Form1.Picture10.Point(i, j), add)
    ElseIf Form1.Picture9.Point(i, j) = RGB(0, 0, 0) Then
        Form1.Picture10.PSet (i, j), shiftcolor(Form1.Picture10.Point(i, j), -add)
    End If
Next j
Next i

End Function
