VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinSkinMaker"
   ClientHeight    =   5745
   ClientLeft      =   1860
   ClientTop       =   2460
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   684
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture26 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1290
      Left            =   5025
      ScaleHeight     =   82
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   39
      Top             =   7395
      Width           =   4185
   End
   Begin VB.PictureBox Picture25 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1290
      Left            =   5025
      ScaleHeight     =   82
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   38
      Top             =   6000
      Width           =   4185
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options for position bar"
      Height          =   1170
      Left            =   8295
      TabIndex        =   35
      Top             =   3810
      Width           =   1905
      Begin VB.OptionButton Option4 
         Caption         =   "3D"
         Height          =   225
         Left            =   390
         TabIndex        =   37
         Top             =   795
         Width           =   1245
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Flat"
         Height          =   210
         Left            =   390
         TabIndex        =   36
         Top             =   435
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options for shuffle indicator"
      Height          =   1170
      Left            =   5940
      TabIndex        =   30
      Top             =   3810
      Width           =   2205
      Begin VB.OptionButton Option2 
         Caption         =   "Style 2"
         Height          =   285
         Left            =   555
         TabIndex        =   32
         Top             =   840
         Width           =   1365
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Style 1"
         Height          =   345
         Left            =   555
         TabIndex        =   31
         Top             =   345
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture8 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   600
      Left            =   8220
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   136
      TabIndex        =   29
      Top             =   10095
      Width           =   2100
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   5970
      TabIndex        =   28
      Top             =   3165
      Width           =   4215
   End
   Begin VB.PictureBox Picture24 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1995
      Left            =   11175
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   27
      Top             =   6840
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   5970
      TabIndex        =   26
      Text            =   "                                        "
      Top             =   2325
      Width           =   4215
   End
   Begin VB.PictureBox Picture23 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Height          =   270
      Left            =   -3270
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   25
      Top             =   6915
      Width           =   4185
   End
   Begin VB.PictureBox Picture22 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3360
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   24
      Top             =   9105
      Width           =   690
   End
   Begin VB.PictureBox Picture21 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3405
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   23
      Top             =   8820
      Width           =   690
   End
   Begin VB.PictureBox Picture20 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   1710
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   22
      Top             =   9090
      Width           =   1545
   End
   Begin VB.PictureBox Picture19 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   1725
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   21
      Top             =   8790
      Width           =   1545
   End
   Begin VB.PictureBox Picture18 
      AutoRedraw      =   -1  'True
      Height          =   4785
      Left            =   10815
      ScaleHeight     =   315
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   20
      Top             =   6030
      Width           =   4185
   End
   Begin VB.PictureBox Picture17 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Height          =   2910
      Left            =   10425
      ScaleHeight     =   190
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   19
      Top             =   2985
      Width           =   4260
   End
   Begin VB.PictureBox Picture16 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2910
      Left            =   10425
      ScaleHeight     =   190
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   18
      Top             =   15
      Width           =   4260
   End
   Begin VB.PictureBox Picture15 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6555
      Left            =   4380
      ScaleHeight     =   433
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   47
      TabIndex        =   17
      Top             =   8910
      Width           =   765
   End
   Begin VB.PictureBox Picture14 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   420
      Left            =   5535
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   16
      Top             =   9810
      Width           =   930
   End
   Begin VB.PictureBox Picture13 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   420
      Left            =   5475
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   15
      Top             =   9285
      Width           =   930
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Height          =   1365
      Left            =   -2850
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   344
      TabIndex        =   14
      Top             =   8130
      Width           =   5220
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1365
      Left            =   -3000
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   344
      TabIndex        =   13
      Top             =   7905
      Width           =   5220
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1335
      Left            =   1605
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   12
      Top             =   9105
      Width           =   1440
   End
   Begin VB.PictureBox Picture9 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1335
      Left            =   90
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   11
      Top             =   8700
      Width           =   1440
   End
   Begin VB.PictureBox Picture7 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   600
      Left            =   8190
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   136
      TabIndex        =   10
      Top             =   9330
      Width           =   2100
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6330
      Left            =   3195
      ScaleHeight     =   418
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   9
      Top             =   8895
      Width           =   1080
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   210
      Left            =   -4170
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   307
      TabIndex        =   8
      Top             =   7575
      Visible         =   0   'False
      Width           =   4665
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   360
      Left            =   4665
      TabIndex        =   7
      Top             =   660
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      Height          =   4785
      Left            =   5775
      ScaleHeight     =   315
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   6
      Top             =   9780
      Width           =   4185
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   5145
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Pictures (*.bmp;*.jpg;*.jpeg)|*.bmp;*.jpg;*.jpeg"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Preview"
      Height          =   360
      Left            =   4665
      TabIndex        =   5
      Top             =   165
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   1800
      Left            =   5970
      ScaleHeight     =   116
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   4
      Top             =   0
      Width           =   4185
   End
   Begin VB.PictureBox Picture1 
      Height          =   5250
      Left            =   0
      ScaleHeight     =   346
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   277
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   11580
         Left            =   120
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   768
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1024
         TabIndex        =   3
         Top             =   720
         Width           =   15420
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   15
      Left            =   60
      TabIndex        =   1
      Top             =   5430
      Width           =   4080
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5175
      LargeChange     =   15
      Left            =   4350
      TabIndex        =   0
      Top             =   -15
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Type (if it's necessary) the name of the skin."
      Height          =   300
      Left            =   5985
      TabIndex        =   34
      Top             =   2805
      Width           =   4200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type below the text which will appear in the titlebar."
      Height          =   195
      Left            =   5940
      TabIndex        =   33
      Top             =   2040
      Width           =   3645
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuLin3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find Skin Directory"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public skinDrv As String, skinDir As String
Private Sub Command1_Click()
Dim i As Integer, j As Integer, r As Long, g As Long, b As Long, add As Integer

Picture3.PaintPicture Picture2.Picture, 0, 0, 275, 116, HScroll1.Value, VScroll1.Value, 275, 116
add = 64
d3d Picture3, 0, 0, 274, 115, add
d3d Picture3, 1, 1, 273, 114, add
d3d Picture3, 110, 25, 265, 34, add
d3d Picture3, 23, 42, 100, 59, -add
If Option4.Value = True Then
d3d Picture3, 15, 71, 264, 82, -add
End If
Picture3.Picture = Picture3.Image

End Sub
Private Sub Command3_Click()
Dim SkinPath As String
On Error GoTo err1
Form3.Show
DoEvents
ChDrive skinDrv
ChDir skinDir
SkinPath = skinDir & "\" & Trim(Text2.Text)
MkDir SkinPath
Picture3.Cls
Command1_Click
Picture23.Cls
Picture12.Cls
Picture4.Cls
Picture16.Cls
Picture17.Cls
Form3.ProgressBar1.Value = 10

Picture4.PaintPicture Picture2.Picture, 0, 0, 275, 116, HScroll1.Value, VScroll1.Value + 116, 275, 116
eqm
Form3.ProgressBar1.Value = 30

cbutton
Form3.ProgressBar1.Value = 40

shuf
Form3.ProgressBar1.Value = 60

mono
Form3.ProgressBar1.Value = 80

bal
Form3.ProgressBar1.Value = 90

pledit
Form3.ProgressBar1.Value = 140
numplay
Form3.ProgressBar1.Value = 150
eq_ex
ChDir SkinPath
SavePicture Picture3.Image, "main.bmp"
SavePicture Picture4.Image, "eqmain.bmp"
SavePicture Picture5.Image, "posbar.bmp"
SavePicture Picture6.Image, "volume.bmp"
SavePicture Picture8.Image, "cbuttons.bmp"
SavePicture Picture10.Image, "shufrep.bmp"
Form3.ProgressBar1.Value = 160
SavePicture Picture12.Image, "titlebar.bmp"
SavePicture Picture14.Image, "monoster.bmp"
SavePicture Picture15.Image, "balance.bmp"
SavePicture Picture17.Image, "pledit.bmp"
SavePicture Picture20.Image, "numbers.bmp"
SavePicture Picture22.Image, "playpaus.bmp"
SavePicture Picture26.Image, "eq_ex.bmp"
Form3.ProgressBar1.Value = 180
Form3.Label1.Caption = "That's all,"
Form3.Caption = "Job done."
Form3.Timer1.Enabled = True
ChDir skinDir
err1:
   If err.number > 0 Then Unload Form3
   If err.number = 76 Then msg = MsgBox("Find Winamp\Skin\ directory first !", vbCritical, "Message from WinSkin")
   If err.number = 75 And Trim(Text2.Text) = "" Then msg = MsgBox("You didn't choose a name for the skin !", vbCritical, "Message from WinSkin")
   If err.number = 75 And Trim(Text2.Text) <> "" Then msg = MsgBox("You already have a skin """ & Trim(Text2.Text) & """ !", vbInformation, "Message from WinSkin")
   If err.number <> 75 And err.number <> 0 And err.number <> 76 Then msg = MsgBox("Error number " & err.number & " !", vbCritical, "Message from WinSkin")
End Sub
Private Sub Form_Load()
start
LoadFromRes
skinLoad
Option1.Value = True
Option3.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
skinUnload
Unload Form2
End Sub

Private Sub HScroll1_Change()
   Picture2.Left = -HScroll1.Value
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuFind_Click()
Form2.Show
End Sub

Private Sub mnuLoad_Click()
Dialog1.ShowOpen
If Dialog1.FileName <> "" Then Picture2.Picture = LoadPicture(Dialog1.FileName)
start
End Sub

Private Sub Option1_Click()
Picture9.Picture = LoadResPicture(102, vbResBitmap)
End Sub

Private Sub Option2_Click()
Picture9.Picture = LoadResPicture(110, vbResBitmap)
End Sub

Private Sub Text1_Change()
Text2.Text = Text1.Text
End Sub

Private Sub VScroll1_Change()
   Picture2.Top = -VScroll1.Value
End Sub
Private Sub start()
   Picture1.Move 0, 0
   Picture2.Move 0, 0

   HScroll1.Top = Picture1.Height
   HScroll1.Left = 0
   HScroll1.Width = Picture1.Width

   VScroll1.Top = 0
   VScroll1.Left = Picture1.Width
   VScroll1.Height = Picture1.Height

   HScroll1.Max = Picture2.Width - Picture1.Width
   VScroll1.Max = Picture2.Height - Picture1.Height

   VScroll1.Visible = (Picture1.Height < Picture2.Height)
   HScroll1.Visible = (Picture1.Width < Picture2.Width)
   HScroll1.Value = 0
   VScroll1.Value = 0
End Sub
Private Sub skinLoad()
On Error Resume Next
ChDir App.Path
ChDrive Left(App.Path, 1)
Open "skin.win" For Input As #1
Line Input #1, skinDrv
Line Input #1, skinDir
Close #1
ChDrive Left(skinDrv, 1)
ChDir skinDir
End Sub
Private Sub skinUnload()
ChDir App.Path
ChDrive Left(App.Path, 1)
Open "skin.win" For Output As #1
Print #1, skinDrv
Print #1, skinDir
Close #1
End Sub
Private Sub LoadFromRes()
Picture7.Picture = LoadResPicture(101, vbResBitmap)
Picture9.Picture = LoadResPicture(102, vbResBitmap)
Picture11.Picture = LoadResPicture(103, vbResBitmap)
Picture13.Picture = LoadResPicture(104, vbResBitmap)
Picture16.Picture = LoadResPicture(105, vbResBitmap)
Picture18.Picture = LoadResPicture(106, vbResBitmap)
Picture19.Picture = LoadResPicture(107, vbResBitmap)
Picture21.Picture = LoadResPicture(108, vbResBitmap)
Picture24.Picture = LoadResPicture(109, vbResBitmap)
Picture25.Picture = LoadResPicture(111, vbResBitmap)
End Sub
