Attribute VB_Name = "eqmin"
Dim i As Integer, j As Integer, add As Integer, h As Integer, v As Integer
Public Function eq_ex()
h = Form1.HScroll1.Value
v = Form1.VScroll1.Value
add = 64
    Form1.Picture26.PaintPicture Form1.Picture2.Picture, 0, 0, 275, 14, h, v + 116, 275, 14
    Form1.Picture26.PaintPicture Form1.Picture2.Picture, 0, 15, 275, 14, h, v + 116, 275, 14
    Form1.Picture26.PaintPicture Form1.Picture2.Picture, 1, 38, 9, 9, h + 255, v + 119, 9, 9
    Form1.Picture26.PaintPicture Form1.Picture2.Picture, 1, 47, 9, 9, h + 255, v + 119, 9, 9
    Form1.Picture26.PaintPicture Form1.Picture2.Picture, 11, 38, 9, 9, h + 264, v + 119, 9, 9
    Form1.Picture26.PaintPicture Form1.Picture2.Picture, 11, 47, 9, 9, h + 264, v + 119, 9, 9
    
For i = 0 To Form1.Picture25.ScaleWidth
For j = 0 To Form1.Picture25.ScaleHeight
    If Form1.Picture25.Point(i, j) = RGB(255, 255, 255) Then
        Form1.Picture26.PSet (i, j), shiftcolor(Form1.Picture26.Point(i, j), add)
    ElseIf Form1.Picture25.Point(i, j) = RGB(0, 0, 0) Then
        Form1.Picture26.PSet (i, j), shiftcolor(Form1.Picture26.Point(i, j), -add)
    ElseIf Form1.Picture25.Point(i, j) = RGB(255, 0, 0) Then
        Form1.Picture26.PSet (i, j), vbBlack
    End If
Next j
Next i
Form1.Picture26.Picture = Form1.Picture26.Image
End Function
