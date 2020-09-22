VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Viewer (Change Path)"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   4080
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form2
End Sub

Private Sub Drive1_Change()
On Error GoTo ab
    Dir1.Path = Drive1.Drive
    ChDrive Drive1.Drive
    Form1.skinDrv = Drive1.Drive
ab: Drive1.Drive = Dir1.Path
    Exit Sub
End Sub
Private Sub Dir1_Change()
    Form1.skinDir = Dir1.Path
    ChDir Dir1.Path
End Sub

Private Sub Form_Load()
On Error GoTo err
Drive1.Drive = Form1.skinDrv
Dir1.Path = Form1.skinDir
err:
    If err.number <> 0 Then
        Form1.skinDrv = Drive1.Drive
        Form1.skinDir = Dir1.Path
    End If
End Sub
