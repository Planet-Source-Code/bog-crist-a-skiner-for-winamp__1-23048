Attribute VB_Name = "balance"
Dim i As Integer, j As Integer, add As Integer
Public Function bal()
add = 64
For i = 0 To 27
Form1.Picture15.PaintPicture Form1.Picture3.Picture, 9, i * 15, 38, 14, 177, 57, 38, 14

'Form1.Picture15.CurrentX = 10
'Form1.Picture15.CurrentY = i * 15
'Form1.Picture15.Print i
Next i
Form1.Picture15.PaintPicture Form1.Picture3.Picture, 0, 422, 14, 11, 189, 58, 14, 11
Form1.Picture15.PaintPicture Form1.Picture3.Picture, 15, 422, 14, 11, 189, 58, 14, 11
d3d Form1.Picture15, 0, 422, 13, 432, -add
d3d Form1.Picture15, 15, 422, 28, 432, add
End Function
