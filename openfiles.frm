VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "File open..."
   ClientHeight    =   4020
   ClientLeft      =   3840
   ClientTop       =   1800
   ClientWidth     =   5865
   Icon            =   "openfiles.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4020
   ScaleWidth      =   5865
   Begin VB.CommandButton Command1 
      Caption         =   "O.K. lets go!"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox full_file 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   5775
   End
   Begin VB.FileListBox file 
      Height          =   2040
      Left            =   3120
      Pattern         =   "*.vbp"
      ReadOnly        =   0   'False
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
   Begin VB.DirListBox folder 
      Height          =   2115
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   3015
   End
   Begin VB.DriveListBox drive 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5280
      Picture         =   "openfiles.frx":030A
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public filename

Private Sub Command1_Click()
Let filename = full_file.Text
'= folder.Path & "\" & file.filename
Me.Hide
Form1.Show
End Sub

Private Sub drive_Change()
Let folder.Path = drive.drive
End Sub

Private Sub file_Click()
Let full_file.Text = folder.Path & "\" & file.filename
If folder.Path = "c:\" Then
Let full_file.Text = folder.Path & file.filename
End If
If folder.Path = "d:\" Then
Let full_file.Text = folder.Path & file.filename
End If

End Sub

Private Sub folder_Change()
Let file.Path = folder.Path
End Sub

Private Sub Form_Load()
Let Label1.Caption = "Please find your vbp file to convert, then press ok." & Chr(13) & "If your file is not there, then please type in the location of the file." & Chr(13) & "The vbp file, for reasons described in the code,cannot be over 1.34kb, I.E less than 1024 bytes." & Chr(13) & "Thanks, and hope this helps you"
End Sub
