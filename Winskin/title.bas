Attribute VB_Name = "title"
Dim i As Integer, j As Integer, add As Integer, name As String
Public Function tit()
name = Form1.Text1.Text

Form1.Picture12.PaintPicture Form1.Picture3.Picture, 27, 29, 275, 14, 0, 0, 275, 14
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 27, 43, 275, 14, 0, 0, 275, 14
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 0, 27, 9, 9, 254, 4, 9, 9
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 9, 27, 9, 9, 254, 4, 9, 9
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 0, 18, 9, 9, 254, 4, 9, 9
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 9, 18, 9, 9, 254, 4, 9, 9
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 0, 0, 9, 9, 6, 3, 9, 9
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 0, 9, 9, 9, 6, 3, 9, 9
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 9, 0, 9, 9, 244, 3, 9, 9
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 9, 9, 9, 9, 244, 3, 9, 9
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 18, 0, 9, 9, 264, 3, 9, 9
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 18, 9, 9, 9, 264, 3, 9, 9

Form1.Picture23.ForeColor = RGB(0, 0, 0)
Form1.Picture23.CurrentX = 5
Form1.Picture23.CurrentY = 1
Form1.Picture23.Print name
Form1.Picture23.ForeColor = RGB(255, 255, 255)
Form1.Picture23.CurrentX = 4
Form1.Picture23.CurrentY = 0
Form1.Picture23.Print name
add = 96
For i = 0 To Form1.Picture23.ScaleWidth
For j = 0 To Form1.Picture23.ScaleHeight
    If Form1.Picture23.Point(i, j) = RGB(255, 255, 255) Then
        Form1.Picture3.PSet (i, j), shiftcolor(Form1.Picture3.Point(i, j), add)
        Form1.Picture4.PSet (i, j), shiftcolor(Form1.Picture4.Point(i, j), add)
    ElseIf Form1.Picture23.Point(i, j) = RGB(0, 0, 0) Then
        Form1.Picture3.PSet (i, j), shiftcolor(Form1.Picture3.Point(i, j), -add)
        Form1.Picture4.PSet (i, j), shiftcolor(Form1.Picture4.Point(i, j), -add)
    End If
Next j
Next i
  Form1.Picture3.Picture = Form1.Picture3.Image
  Form1.Picture4.Picture = Form1.Picture4.Image
  
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 27, 0, 275, 14, 0, 0, 275, 14
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 27, 15, 275, 14, 0, 0, 275, 14
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 27, 58, 275, 14, 0, 0, 275, 14
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 27, 73, 275, 14, 0, 0, 275, 14
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 304, 0, 8, 44, 10, 22, 8, 44
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 312, 0, 8, 44, 10, 22, 8, 44
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 304, 44, 8, 44, 10, 22, 8, 44
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 312, 44, 8, 44, 10, 22, 8, 44
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 320, 44, 8, 44, 10, 22, 8, 44
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 328, 44, 8, 44, 10, 22, 8, 44
Form1.Picture12.PaintPicture Form1.Picture3.Picture, 336, 44, 8, 44, 10, 22, 8, 44

For i = 0 To Form1.Picture11.ScaleWidth
For j = 0 To Form1.Picture11.ScaleHeight
    If Form1.Picture11.Point(i, j) = RGB(255, 255, 255) Then
        Form1.Picture12.PSet (i, j), shiftcolor(Form1.Picture12.Point(i, j), add)
    ElseIf Form1.Picture11.Point(i, j) = RGB(0, 0, 0) Then
        Form1.Picture12.PSet (i, j), shiftcolor(Form1.Picture12.Point(i, j), -add)
    End If
Next j
Next i

End Function
