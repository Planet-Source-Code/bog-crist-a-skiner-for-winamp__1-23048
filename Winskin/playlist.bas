Attribute VB_Name = "playlist"
Dim i As Integer, j As Integer, add As Integer, h As Integer, v As Integer
Public Function pledit()
h = Form1.HScroll1.Value
v = Form1.VScroll1.Value
add = 64
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 0, 72, 125, 38, h, v + 310, 125, 38
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 126, 72, 150, 38, h + 126, v + 310, 150, 38
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 26, 0, 100, 20, h + 87, v + 232, 100, 20
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 26, 21, 100, 20, h + 87, v + 232, 100, 20
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 153, 0, 25, 20, h + 250, v + 232, 25, 20
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 153, 21, 25, 20, h + 250, v + 232, 25, 20
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 0, 0, 25, 20, h, v + 232, 25, 20
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 0, 21, 25, 20, h, v + 232, 25, 20
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 127, 0, 25, 20, h + 224, v + 232, 25, 20
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 127, 21, 25, 20, h + 224, v + 232, 25, 20
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 31, 42, 20, 29, h + 255, v + 252, 20, 29
Form3.ProgressBar1.Value = 100

    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 0, 42, 12, 29, h, v + 252, 12, 29
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 52, 53, 8, 18, h + 260, v + 252, 8, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 61, 53, 8, 18, h + 260, v + 252, 8, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 179, 0, 25, 38, h + 126, v + 310, 25, 38
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 205, 0, 275, 38, h + 126, v + 310, 275, 38
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 99, 42, 50, 14, h + 225, v + 232, 50, 14
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 99, 57, 50, 14, h + 225, v + 232, 50, 14
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 72, 42, 25, 14, h, v + 232, 25, 14
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 72, 57, 25, 14, h + 25, v + 232, 25, 14
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 150, 42, 9, 9, h + 254, v + 235, 9, 9
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 62, 42, 9, 9, h + 254, v + 235, 9, 9
Form3.ProgressBar1.Value = 110

    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 0, 149, 22, 18, h + 14, v + 318, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 0, 130, 22, 18, h + 14, v + 300, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 0, 111, 22, 18, h + 14, v + 282, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 48, 111, 3, 54, h + 11, v + 282, 3, 54
    
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 54, 149, 22, 18, h + 43, v + 318, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 54, 130, 22, 18, h + 43, v + 300, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 54, 111, 22, 18, h + 43, v + 282, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 54, 168, 22, 18, h + 43, v + 264, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 100, 111, 3, 72, h + 40, v + 264, 3, 72
    
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 104, 149, 22, 18, h + 72, v + 318, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 104, 130, 22, 18, h + 72, v + 300, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 104, 111, 22, 18, h + 72, v + 282, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 150, 111, 3, 54, h + 69, v + 282, 3, 54
    
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 154, 149, 22, 18, h + 101, v + 318, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 154, 130, 22, 18, h + 101, v + 300, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 154, 111, 22, 18, h + 101, v + 282, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 200, 111, 3, 54, h + 98, v + 282, 3, 54
    
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 204, 149, 22, 18, h + 231, v + 318, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 204, 130, 22, 18, h + 231, v + 300, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 204, 111, 22, 18, h + 231, v + 282, 22, 18
    Form1.Picture17.PaintPicture Form1.Picture2.Picture, 250, 111, 3, 54, h + 228, v + 282, 3, 54
Form3.ProgressBar1.Value = 120


For i = 0 To Form1.Picture16.ScaleWidth
For j = 0 To Form1.Picture16.ScaleHeight
    If Form1.Picture16.Point(i, j) = RGB(255, 255, 255) Then
        Form1.Picture17.PSet (i, j), shiftcolor(Form1.Picture17.Point(i, j), add)
    ElseIf Form1.Picture16.Point(i, j) = RGB(0, 0, 0) Then
        Form1.Picture17.PSet (i, j), shiftcolor(Form1.Picture17.Point(i, j), -add)
    End If
Next j
Next i

 Form1.Picture17.Picture = Form1.Picture17.Image
End Function
