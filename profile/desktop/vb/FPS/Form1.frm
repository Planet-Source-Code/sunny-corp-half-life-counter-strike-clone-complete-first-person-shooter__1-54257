VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Introduction"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Proceed"
      Height          =   435
      Left            =   2520
      TabIndex        =   13
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "(c) 2004 A.A"
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Label Label13 
      Caption         =   "Open Doors: O"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Talk: T"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Reload: Right Click"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Shoot: Left Click"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Strafe Right: S"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Strafe Left: A"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Move Backward: S"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Move Forward: W"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Look Around: Mouse"
      Height          =   1335
      Left            =   600
      TabIndex        =   4
      Top             =   3000
      Width           =   6015
   End
   Begin VB.Label Label4 
      Caption         =   "                                                             Controls:                "
      Height          =   1335
      Left            =   600
      TabIndex        =   3
      Top             =   2760
      Width           =   6015
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":0000
      Height          =   1935
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":00A6
      Height          =   1455
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "WHAT??!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load frmMain
frmMain.Show
End Sub

Private Sub Form_Load()

End Sub
