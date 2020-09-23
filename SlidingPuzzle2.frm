VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Select Picture"
   ClientHeight    =   2880
   ClientLeft      =   9270
   ClientTop       =   2880
   ClientWidth     =   4485
   Icon            =   "SlidingPuzzle2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4485
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Select Picture"
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Width           =   4455
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFFFF&
      Height          =   2040
      Left            =   2520
      Pattern         =   "*.bmp;*.gif;*.jpg"
      TabIndex        =   2
      Top             =   0
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

File = File1.Path
If Right(File, 1) <> "\" Then File = File & "\"
File = File & File1.FileName
Form2.Hide
SlidingPuzzle.Show
SlidingPuzzle.Enabled = True
SlidingPuzzle.LoadPic

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo err:
Dir1.Path = Drive1.Drive
Text1.Text = Dir1.Path
Exit Sub
err:
MsgBox "Device Unavailable.", , "Error"
End Sub

Private Sub File1_Click()
Text1.Text = File1.Path & "\" & File1.FileName
End Sub

Private Sub File1_DblClick()
Command1_Click
End Sub

Private Sub Form_Load()
SlidingPuzzle.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
SlidingPuzzle.Enabled = True
End Sub

