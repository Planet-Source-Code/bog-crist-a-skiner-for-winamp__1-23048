Attribute VB_Name = "mod3d"
Public Function shiftcolor(pixel As Long, add As Integer) As Long
Dim r As Integer, g As Integer, b As Integer
r = (pixel And RGB(255, 0, 0)) + add           'This three lines extract the red,green and blue components
g = ((pixel And RGB(0, 255, 0)) \ 256) + add   'from the long integer returned by RGB and add to each one a constant.
b = ((pixel And RGB(0, 0, 255)) \ 65536) + add 'A more accurate method to obtain a brighter or darker color than a given one
                                               'is to transform from RGB to HSV (Hue,Saturation,Value), then to increase the
                                               '(V)alue parameter,then to transform to RGB again.I don't think that BitBlt
                                               'techniques work in this case.

If add > 0 Then
    If r > 255 Then r = 255
    If g > 255 Then g = 255
    If b > 255 Then b = 255
ElseIf add < 0 Then
    If r < 0 Then r = 0
    If g < 0 Then g = 0
    If b < 0 Then b = 0
End If
shiftcolor = RGB(r, g, b)
End Function
Public Sub d3d(pic As PictureBox, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, add As Integer)
'This sub draws a rectangle which will appear as being 3-dimensional.
'If add is positive the rectangle looks like an unpushed button and if add is negative the rectangle looks like a pushed button .
Dim i As Integer
For i = x1 To x2
pic.PSet (i, y1), shiftcolor(pic.Point(i, y1), add)
Next i
For i = y1 + 1 To y2 - 1
pic.PSet (x1, i), shiftcolor(pic.Point(x1, i), add)
Next i
For i = x1 To x2
pic.PSet (i, y2), shiftcolor(pic.Point(i, y2), -add)
Next i
For i = y1 + 1 To y2 - 1
pic.PSet (x2, i), shiftcolor(pic.Point(x2, i), -add)
Next i

End Sub
