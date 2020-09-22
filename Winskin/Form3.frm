VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "I'm  just  doing  my  job . . ."
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   6555
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   345
      Top             =   2340
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   465
      Left            =   165
      TabIndex        =   0
      Top             =   1050
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   820
      _Version        =   393216
      Appearance      =   1
      Max             =   180
   End
   Begin VB.Label Label2 
      Caption         =   "folks !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3630
      TabIndex        =   2
      Top             =   150
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   " Please wait a little,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   885
      TabIndex        =   1
      Top             =   150
      Width           =   2685
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time As Byte

Private Sub Timer1_Timer()
time = time + 1
If time = 2 Then
    time = 0                     ' Don't remove this
    Timer1.Enabled = False       ' two lines .
    Unload Form3
End If
End Sub
