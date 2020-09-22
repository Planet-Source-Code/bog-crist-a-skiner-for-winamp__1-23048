Attribute VB_Name = "number"
Dim i As Integer, j As Integer, average As Long, r As Long, g As Long, b As Long, t As Single, opus As Long
Dim r1 As Integer, g1 As Integer, b1 As Integer
Public Function numplay()
'----------------------------------------------------------------------------------------------------
For i = 36 To 98                                                    'These lines calculate the average
For j = 26 To 38                                                    'color for the area in which will
r = r + (Form1.Picture3.Point(i, j) And RGB(255, 0, 0))             'be display the time.This color will
g = g + ((Form1.Picture3.Point(i, j) And RGB(0, 255, 0)) \ 256)     'be the background for the clock digits.
b = b + ((Form1.Picture3.Point(i, j) And RGB(0, 0, 255)) \ 65536)
Next j
Next i
r = r / 819: g = g / 819: b = b / 819     '819=(38-26)*(98-36)
average = RGB(r, g, b)
'-----------------------------------------------------------------------------------------------------
r1 = r: g1 = g: b1 = b: t = 1

Do
r = r1 + t * (127 - r1)
g = g1 + t * (127 - g1)
b = b1 + t * (127 - b1)
t = t + 0.05
Loop Until r - 20 < 0 Or r + 20 > 255 Or g - 20 < 0 Or g + 20 > 255 Or b - 20 < 0 Or b + 20 > 255
opus = RGB(r, g, b)

For i = 0 To Form1.Picture19.ScaleWidth
For j = 0 To Form1.Picture19.ScaleHeight
    If Form1.Picture19.Point(i, j) = RGB(0, 255, 0) Then
        Form1.Picture20.PSet (i, j), average
    Else
        Form1.Picture20.PSet (i, j), opus
    End If
Next j
Next i

For i = 73 To 74
Form1.Picture3.PSet (i, 30), opus
Form1.Picture3.PSet (i, 34), opus
Next i

Form3.ProgressBar1.Value = 130

r = 0: g = 0: b = 0
For i = 24 To 34
For j = 28 To 36
r = r + (Form1.Picture3.Point(i, j) And RGB(255, 0, 0))
g = g + ((Form1.Picture3.Point(i, j) And RGB(0, 255, 0)) \ 256)
b = b + ((Form1.Picture3.Point(i, j) And RGB(0, 0, 255)) \ 65536)
Next j
Next i
r = r / 99: g = g / 99: b = b / 99
average = RGB(r, g, b)

r1 = r: g1 = g: b1 = b: t = 1

Do
r = r1 + t * (127 - r1)
g = g1 + t * (127 - g1)
b = b1 + t * (127 - b1)
t = t + 0.05
Loop Until r - 20 < 0 Or r + 20 > 255 Or g - 20 < 0 Or g + 20 > 255 Or b - 20 < 0 Or b + 20 > 255
opus = RGB(r, g, b)

For i = 0 To Form1.Picture21.ScaleWidth
For j = 0 To Form1.Picture21.ScaleHeight
    If Form1.Picture21.Point(i, j) = RGB(0, 255, 0) Then
        Form1.Picture22.PSet (i, j), average
    Else
        Form1.Picture22.PSet (i, j), opus
    End If
Next j
Next i



End Function
