VERSION 5.00
Begin VB.Form form1 
   Caption         =   "Main Convert "
   ClientHeight    =   6600
   ClientLeft      =   2625
   ClientTop       =   1275
   ClientWidth     =   7155
   Icon            =   "convert.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   7155
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   38
      Left            =   4560
      ScrollBars      =   1  'Horizontal
      TabIndex        =   46
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   37
      Left            =   2640
      ScrollBars      =   1  'Horizontal
      TabIndex        =   45
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   36
      Left            =   720
      ScrollBars      =   1  'Horizontal
      TabIndex        =   44
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   35
      Left            =   4560
      ScrollBars      =   1  'Horizontal
      TabIndex        =   43
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   34
      Left            =   2640
      ScrollBars      =   1  'Horizontal
      TabIndex        =   42
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   33
      Left            =   720
      ScrollBars      =   1  'Horizontal
      TabIndex        =   41
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   32
      Left            =   4560
      ScrollBars      =   1  'Horizontal
      TabIndex        =   40
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   31
      Left            =   2640
      ScrollBars      =   1  'Horizontal
      TabIndex        =   39
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5760
      TabIndex        =   35
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "open and convert another"
      Height          =   375
      Left            =   4320
      TabIndex        =   34
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   30
      Left            =   720
      ScrollBars      =   1  'Horizontal
      TabIndex        =   33
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Convert 6 - 5"
      Height          =   375
      Left            =   2880
      TabIndex        =   32
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check"
      Height          =   375
      Left            =   1440
      TabIndex        =   31
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   29
      Left            =   4560
      ScrollBars      =   1  'Horizontal
      TabIndex        =   30
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   28
      Left            =   4560
      ScrollBars      =   1  'Horizontal
      TabIndex        =   29
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   27
      Left            =   4560
      ScrollBars      =   1  'Horizontal
      TabIndex        =   28
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   26
      Left            =   4560
      ScrollBars      =   1  'Horizontal
      TabIndex        =   27
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   25
      Left            =   2640
      ScrollBars      =   1  'Horizontal
      TabIndex        =   26
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   24
      Left            =   2640
      ScrollBars      =   1  'Horizontal
      TabIndex        =   25
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   23
      Left            =   2640
      ScrollBars      =   1  'Horizontal
      TabIndex        =   24
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   22
      Left            =   2640
      ScrollBars      =   1  'Horizontal
      TabIndex        =   23
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   21
      Left            =   720
      ScrollBars      =   1  'Horizontal
      TabIndex        =   22
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   20
      Left            =   720
      ScrollBars      =   1  'Horizontal
      TabIndex        =   21
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   19
      Left            =   720
      ScrollBars      =   1  'Horizontal
      TabIndex        =   20
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   18
      Left            =   720
      ScrollBars      =   1  'Horizontal
      TabIndex        =   19
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   17
      Left            =   4560
      ScrollBars      =   1  'Horizontal
      TabIndex        =   18
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   16
      Left            =   4560
      ScrollBars      =   1  'Horizontal
      TabIndex        =   17
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   15
      Left            =   4560
      ScrollBars      =   1  'Horizontal
      TabIndex        =   16
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   14
      Left            =   4560
      ScrollBars      =   1  'Horizontal
      TabIndex        =   15
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   13
      Left            =   4560
      ScrollBars      =   1  'Horizontal
      TabIndex        =   14
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   12
      Left            =   4560
      ScrollBars      =   1  'Horizontal
      TabIndex        =   13
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   11
      Left            =   2640
      ScrollBars      =   1  'Horizontal
      TabIndex        =   12
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   10
      Left            =   2640
      ScrollBars      =   1  'Horizontal
      TabIndex        =   11
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   9
      Left            =   2640
      ScrollBars      =   1  'Horizontal
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   2640
      ScrollBars      =   1  'Horizontal
      TabIndex        =   9
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   2640
      ScrollBars      =   1  'Horizontal
      TabIndex        =   8
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   2640
      ScrollBars      =   1  'Horizontal
      TabIndex        =   7
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   720
      ScrollBars      =   1  'Horizontal
      TabIndex        =   6
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   720
      ScrollBars      =   1  'Horizontal
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   720
      ScrollBars      =   1  'Horizontal
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   720
      ScrollBars      =   1  'Horizontal
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   720
      ScrollBars      =   1  'Horizontal
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   720
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open and read"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6600
      Picture         =   "convert.frx":030A
      Top             =   480
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "convert.frx":0614
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Done!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2880
      TabIndex        =   38
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "==>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1440
      TabIndex        =   37
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "==>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/-------------------------------------------------\
'* Visual Basic 6 to Visual Basic 5 file converter *
'\-------------------------------------------------/
'===================================================
'/-------------------------------------------------\
'* this is the main VB6 -5 convert part            *
'* there is only one catch 22, too little boxes,   *
'* and you dont get all lines, too maney, you make *
'* the vbp file massive. so I chose 38, so as to   *
'* make it easy.                                   *
'\-------------------------------------------------/


Dim amd As String
Dim k
Private Sub Command1_Click()
On Error GoTo lineread
Open Form2.filename For Input As #1
 
Do While Not EOF(1) ' Check for end of file.
For k = 0 To 38 ' there are 38 textboxes
'you wont see them upon running it
    Line Input #1, InputData    ' Read line of data.
     Text1(k).Text = InputData
Next k
Loop
lineread:

Do While Not EOF(1) ' Check for end of file.
For k = 0 To 38
    Line Input #1, InputData    ' Read line of data.
    Let Text1(k).Text = InputData
Next k
Loop
Command1.Visible = False
End Sub

Private Sub Command2_Click()
Let Form1.Height = 7005
Let Form1.Width = 7275

' this checks for retained & DebugStartUpOption
For k = 0 To 38
If Text1(k).Text = "Retained=0" Then ' Retained first
Text1(k).BackColor = vbRed
Else
Text1(k).BackColor = vbGreen
End If
Next k
For k = 0 To 38
If Text1(k).Text = "DebugStartUpOption=0" Then ' DebugStartUpOption next
Text1(k).BackColor = vbRed
End If
If Text1(k).Text = "" Then
Let Text1(k).BackColor = vbBlack
End If
Next k
Close #1
Command2.Visible = False
End Sub

Private Sub Command3_Click()
' since all info in this is in the textboxes, clear the vbp file
Open Form2.filename For Output As #1
Close #1 ' closing it to open in append

Open Form2.filename For Append As #1 ' reopen it
For k = 0 To 38 ' check the boxes for the failed lines & clear them
If Text1(k).Text = "Retained=0" Then
Let Text1(k).Text = ""
End If
If Text1(k).Text = "DebugStartUpOption=0" Then ' DebugStartUpOption next
Text1(k).Text = ""
End If
Next k
For k = 0 To 38
Let amd = Text1(k).Text ' just an extra, to make error handling
Print #1, amd 'print the lines back to the file
Next k 'repeat for all boxes
Command3.Visible = False
End Sub

Private Sub Command4_Click()
Me.Hide
Form2.Show
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Form_Activate()
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
For k = 1 To 38
Text1(k).Text = ""
Next k
End Sub

Private Sub Form_Initialize()
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
For k = 1 To 38
Text1(k).Text = ""
Next k
End Sub

Private Sub Form_Load()
Let Form1.Height = 840
Form1.Width = 7275
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
For k = 1 To 38
Text1(k).Text = ""
Next k
End Sub
