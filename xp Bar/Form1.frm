VERSION 5.00
Object = "*\AOcx\ProjectBar1.vbp"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Progres Bar Sample form"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   495
      Left            =   2370
      TabIndex        =   4
      Top             =   1335
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3750
      TabIndex        =   3
      Top             =   1335
      Width           =   1305
   End
   Begin Project1.CarlosProgressBar CarlosProgressBar1 
      Height          =   255
      Left            =   285
      TabIndex        =   0
      Top             =   840
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   450
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7395
      Top             =   195
   End
   Begin VB.Label Label2 
      Caption         =   "Created by : Carl Harvey,   INFO =  carl.harvey@videotron.ca"
      Height          =   270
      Left            =   270
      TabIndex        =   2
      Top             =   510
      Width           =   6465
   End
   Begin VB.Label Label1 
      Caption         =   "Progress Bar sample"
      Height          =   270
      Left            =   300
      TabIndex        =   1
      Top             =   135
      Width           =   6330
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
CarlosProgressBar1.Max = 100
CarlosProgressBar1.Value = 0
End Sub

Private Sub Timer1_Timer()
CarlosProgressBar1.Value = CarlosProgressBar1.Value + 1
End Sub

