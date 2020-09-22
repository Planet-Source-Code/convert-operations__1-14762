VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Convert Operations..."
   ClientHeight    =   3960
   ClientLeft      =   2625
   ClientTop       =   2325
   ClientWidth     =   7485
   Icon            =   "SPC.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3960
   ScaleWidth      =   7485
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   480
      Top             =   3360
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6720
      Picture         =   "SPC.frx":030A
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   3720
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1785
      Left            =   240
      Picture         =   "SPC.frx":0614
      Top             =   2160
      Width           =   6930
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   7095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Pythagoras"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic
Private Sub Form_Load()


Label1.Caption = App.ProductName
Label2.Caption = "Made By Matrix Man"
End Sub

Private Sub Image1_Click()
MsgBox "See this no more!", vbExclamation, "See ya!"
Image1.Visible = False
End Sub

Private Sub Label3_Click()
MsgBox "See the Readme file for this!", vbInformation, "Help"
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 0
Me.Hide
Form2.Show

End Sub
