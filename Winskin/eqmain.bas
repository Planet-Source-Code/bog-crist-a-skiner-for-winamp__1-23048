Attribute VB_Name = "eqmain"
Dim i As Single, j As Integer, add As Integer, k As Integer
Public Function eqm()

add = 64

d3d Form1.Picture4, 0, 0, 274, 115, add
d3d Form1.Picture4, 1, 1, 273, 114, add

Form1.Picture4.Picture = Form1.Picture4.Image
tit
Form1.Picture4.PaintPicture Form1.Picture4.Picture, 0, 134, 275, 14, 0, 0, 275, 14
Form1.Picture4.PaintPicture Form1.Picture4.Picture, 0, 149, 275, 14, 0, 0, 275, 14
Form1.Picture4.PaintPicture Form1.Picture4.Picture, 0, 116, 9, 9, 264, 3, 9, 9
Form1.Picture4.PaintPicture Form1.Picture4.Picture, 0, 125, 9, 9, 264, 3, 9, 9

Form1.Picture5.PaintPicture Form1.Picture3.Picture, 0, 0, 248, 10, 16, 72, 248, 10
Form1.Picture5.Picture = Form1.Picture5.Image
Form1.Picture5.PaintPicture Form1.Picture5.Picture, 248, 0, 29, 10, 0, 0, 29, 10
Form1.Picture5.PaintPicture Form1.Picture5.Picture, 278, 0, 29, 10, 0, 0, 29, 10
d3d Form1.Picture5, 248, 0, 276, 9, add
d3d Form1.Picture5, 278, 0, 306, 9, -add
Form1.Picture4.PaintPicture Form1.Picture4.Picture, 0, 294, 113, 20, 86, 17, 113, 20

Form1.Picture4.PaintPicture Form1.Picture4.Picture, 10, 119, 58, 12, 14, 18, 58, 12
Form1.Picture4.PaintPicture Form1.Picture4.Picture, 69, 119, 58, 12, 14, 18, 58, 12
Form1.Picture4.PaintPicture Form1.Picture4.Picture, 128, 119, 58, 12, 14, 18, 58, 12
Form1.Picture4.PaintPicture Form1.Picture4.Picture, 187, 119, 58, 12, 14, 18, 58, 12
Form1.Picture4.PaintPicture Form1.Picture4.Picture, 224, 164, 44, 12, 217, 18, 44, 12
Form1.Picture4.PaintPicture Form1.Picture4.Picture, 224, 176, 44, 12, 217, 18, 44, 12
Form1.Picture4.PaintPicture Form1.Picture4.Picture, 0, 164, 11, 11, 22, 62, 11, 11
Form1.Picture4.PaintPicture Form1.Picture4.Picture, 0, 176, 11, 11, 22, 62, 11, 11
Form1.Picture4.PaintPicture Form1.Picture4.Picture, 0, 314, 113, 1, 86, 26, 113, 1
Form3.ProgressBar1.Value = 20


For i = 0 To Form1.Picture18.ScaleWidth
For j = 0 To Form1.Picture18.ScaleHeight
    If Form1.Picture18.Point(i, j) = RGB(255, 255, 255) Then
        Form1.Picture4.PSet (i, j), shiftcolor(Form1.Picture4.Point(i, j), add)
    ElseIf Form1.Picture18.Point(i, j) = RGB(0, 0, 0) Then
        Form1.Picture4.PSet (i, j), shiftcolor(Form1.Picture4.Point(i, j), -add)
    End If
Next j
Next i

For j = 0 To 27
    Form1.Picture6.PaintPicture Form1.Picture3.Picture, 0, j * 15, 68, 14, 107, 57, 68, 14
Next j
i = 0

For j = 0 To 27
    d3d Form1.Picture6, CInt(i), j * 15, CInt(i) + 15, j * 15 + 12, add
    i = i + 1.89
Next j

Form1.Picture4.PaintPicture Form1.Picture24.Picture, 13, 164, 209, 129, 0, 0, 209, 129

End Function
