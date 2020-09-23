VERSION 5.00
Begin VB.Form SlidingPuzzle 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sliding Puzzle"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   Icon            =   "SlidingPuzzle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   618
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Picture"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   12
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Solve"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   11
      Top             =   6000
      Width           =   1455
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "SlidingPuzzle.frx":0442
      Left            =   7800
      List            =   "SlidingPuzzle.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3480
      Width           =   975
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "SlidingPuzzle.frx":0446
      Left            =   7800
      List            =   "SlidingPuzzle.frx":0448
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   6720
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "SlidingPuzzle.frx":044A
      Left            =   7800
      List            =   "SlidingPuzzle.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "SlidingPuzzle.frx":044E
      Left            =   7800
      List            =   "SlidingPuzzle.frx":0450
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      Height          =   7260
      Left            =   120
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   7260
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Congratulations!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1335
         Left            =   360
         TabIndex        =   10
         Top             =   2520
         Visible         =   0   'False
         Width           =   6375
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      Height          =   7260
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   7260
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Scramble"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   7560
      TabIndex        =   8
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Tile Sliding"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   7560
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Gap"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   7560
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Tiles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   7500
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "SlidingPuzzle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim Grid(25, 25) As Integer, Coord(25) As Integer
Dim A As Integer, B As Integer, C As Integer, D As Integer
Dim NumOfTiles As Integer, Size As Integer, Old As Integer
Dim Gap As Integer, Blocks As Long
Dim DestX As Integer, DestY As Integer
Dim TempX As Integer, TempY As Integer
Dim X As Integer, Y As Integer
Dim Smooth As Boolean, GameOn As Boolean
Dim RF As Boolean, MixMoves As Integer
Dim MoveRecord As String, T(3) As String
Dim Colr As ColorConstants, CheckingOn As Boolean
Dim JustWon As Boolean, Solving As Boolean, Scrambling As Boolean
Dim Tick As Long, ShiftOn As Boolean, Pause As Integer

''Constants for the GenerateDC function
''**LoadImage Constants**
'Const IMAGE_BITMAP As Long = 0
'Const LR_LOADFROMFILE As Long = &H10
'Const LR_CREATEDIBSECTION As Long = &H2000
'Const LR_DEFAULTSIZE As Long = &H40

Private Sub Combo1_Click()

If Combo1.ListIndex + 3 <> NumOfTiles Then
    NumOfTiles = Combo1.ListIndex + 3
    ResetScreen
End If

End Sub

Private Sub Combo2_Click()

If Combo2.ListIndex <> Gap Then
    Gap = Combo2.ListIndex
    RefreshScreen
End If

End Sub

Private Sub Combo3_Click()
If Combo3.ListIndex = 0 Then
    Smooth = False
Else
    Smooth = True
End If

End Sub

Private Sub Combo4_Click()

MixMoves = Val(Combo4.Text)

End Sub

Private Sub Command1_Click()

If GameOn Or JustWon Then
    GameOn = False
    JustWon = False
    Label5.Visible = False
    Command1.Caption = "Start"
    Command2.Enabled = False
    Command3.Enabled = True
    Combo1.Enabled = True
    Combo4.Enabled = True
    ResetScreen
Else
    Combo1.Enabled = False
    Combo4.Enabled = False
    Scramble
    Command1.Caption = "Reset"
    Command2.Enabled = True
    Command3.Enabled = False
    GameOn = True
End If
Picture1.SetFocus
End Sub

Private Sub Command2_Click()
Solving = True
GameOn = False
Command1.Enabled = False
Command2.Enabled = False

Pause = 5000 \ Len(MoveRecord)
If Smooth And Size ^ 2 * Len(MoveRecord) > 500000 Then
    Smooth = False
End If

For D = Len(MoveRecord) To 1 Step -1

DestX = X
DestY = Y

Select Case Mid$(MoveRecord, D, 1)

Case Is = "u"
    DestY = Y - 1
    B = 1
    
Case Is = "d"
    DestY = Y + 1
    B = 0
    
Case Is = "l"
    DestX = X - 1
    B = 3
    
Case Is = "r"
    DestX = X + 1
    B = 2
    
End Select

If Grid(DestX, DestY) = 0 Then MsgBox "What the hell?": Exit Sub

Grid(X, Y) = Grid(DestX, DestY)
Grid(DestX, DestY) = 0

If Smooth Then
    'Right
    If DestX < X Then
        For A = Coord(DestX) + 1 To Coord(X)
        Picture1.Line (Coord(DestX), Coord(DestY))-(A - 1, Coord(DestY) + Size - 1 - Gap), Colr, BF
        DrawBlock A, Coord(Y), Grid(X, Y)
        Next A
    End If
    'Down
    If DestY < Y Then
        For A = Coord(DestY) + 1 To Coord(Y)
        Picture1.Line (Coord(DestX), Coord(DestY))-(Coord(DestX) + Size - 1 - Gap, A - 1), Colr, BF
        DrawBlock Coord(X), A, Grid(X, Y)
        Next A
    End If
    'Left
    If DestX > X Then
        For A = Coord(DestX) To Coord(X) Step -1
        Picture1.Line (A + Size - 1, Coord(DestY))-(Coord(DestX) + Size - 1 - Gap, Coord(DestY) + Size - 1 - Gap), Colr, BF
        DrawBlock A, Coord(Y), Grid(X, Y)
        Next A
    End If
    'Up
    If DestY > Y Then
        For A = Coord(DestY) To Coord(Y) Step -1
        Picture1.Line (Coord(DestX), A + Size - 1)-(Coord(DestX) + Size - 1 - Gap, Coord(DestY) + Size - 1 - Gap), Colr, BF
        DrawBlock Coord(X), A, Grid(X, Y)
        Next A
    End If
Else
    Picture1.Line (Coord(DestX), Coord(DestY))-(Coord(DestX) + Size - 1 - Gap, Coord(DestY) + Size - 1 - Gap), Colr, BF
    DrawBlock Coord(X), Coord(Y), Grid(X, Y)
End If

X = DestX
Y = DestY

If Not Smooth Then
    Tick = GetTickCount
    Do While GetTickCount < Tick + Pause
    Loop
End If

Next D

DrawCorner
Label5.Visible = True
Command1.Caption = "Reset"
Command1.Enabled = True
GameOn = False
JustWon = True
Combo3_Click

End Sub

Private Sub Command3_Click()
Form2.Show
End Sub

Private Sub Form_Load()
'Blocks = GenerateDC(App.Path & "\default.jpg")
Picture2.Picture = LoadPicture(App.Path & "\default.jpg")
NumOfTiles = 3
Gap = 0
MixMoves = 15
Smooth = True
CheckingOn = True
Colr = Picture1.BackColor
T(0) = "u"
T(1) = "d"
T(2) = "l"
T(3) = "r"

Combo1.AddItem "3 x 3"
Combo1.AddItem "4 x 4"
Combo1.AddItem "5 x 5"
Combo1.AddItem "6 x 6"
Combo1.AddItem "7 x 7"
Combo1.AddItem "8 x 8"
Combo1.AddItem "9 x 9"
Combo1.AddItem "10 x 10"
Combo2.AddItem "No Gap"
Combo2.AddItem "Small"
Combo2.AddItem "Medium"
Combo2.AddItem "Large"
Combo3.AddItem "Jumpy"
Combo3.AddItem "Smooth"
Combo4.AddItem "5 Moves"
Combo4.AddItem "10 Moves"
Combo4.AddItem "15 Moves"
Combo4.AddItem "25 Moves"
Combo4.AddItem "50 Moves"
Combo4.AddItem "75 Moves"
Combo4.AddItem "100 Moves"
Combo4.AddItem "150 Moves"
Combo4.AddItem "200 Moves"
Combo4.AddItem "250 Moves"
Combo4.AddItem "500 Moves"
Combo4.AddItem "1000 Moves"
Combo1.ListIndex = 0
Combo2.ListIndex = 0
Combo3.ListIndex = 0
Combo4.ListIndex = 2

ResetScreen

End Sub

''IN: FileName: The file name of the graphics
''OUT: The Generated DC
'Public Function GenerateDC(FileName As String) As Long
'Dim DC As Long
'Dim hBitmap As Long
'
''Create a Device Context, compatible with the screen
'DC = CreateCompatibleDC(0)
'
'If DC < 1 Then
'    GenerateDC = 0
'    Exit Function
'End If
'
''Load the image....BIG NOTE: This function is not supported under NT, there you can not
''specify the LR_LOADFROMFILE flag
'hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
'
'If hBitmap = 0 Then 'Failure in loading bitmap
'    DeleteDC DC
'    GenerateDC = 0
'    Exit Function
'End If
'
''Throw the Bitmap into the Device Context
'SelectObject DC, hBitmap
'
''Return the device context
'GenerateDC = DC
'
''Delte the bitmap handle object
'DeleteObject hBitmap
'
'End Function
'
''Deletes a generated DC
'Private Function DeleteGeneratedDC(DC As Long) As Long
'
'If DC > 0 Then
'    DeleteGeneratedDC = DeleteDC(DC)
'Else
'   DeleteGeneratedDC = 0
'End If
'
'End Function

'********** DRAW BLOCK **********
Private Sub DrawBlock(dX, dY, dTileNum)
TempX = dTileNum Mod NumOfTiles
If TempX = 0 Then TempX = NumOfTiles
TempY = (dTileNum - 1) \ NumOfTiles + 1

BitBlt Picture1.hdc, dX, dY, Size - Gap, Size - Gap, Picture2.hdc, Coord(TempX), Coord(TempY), vbSrcCopy

If RF Then Picture1.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
'DeleteGeneratedDC Blocks
Unload Me
Unload Form2
End
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)

If GameOn And Shift = 0 And ShiftOn Then
    ShiftOn = False
    RefreshScreen
End If

End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)

If GameOn And Shift = 2 And ShiftOn Then
    Exit Sub
End If

If GameOn And Shift = 2 Then
    ShiftOn = True
    For B = 1 To NumOfTiles
    For A = 1 To NumOfTiles
    If Grid(A, B) > 0 And Grid(A, B) <> (B - 1) * NumOfTiles + A Then
        Picture1.Line (Coord(A), Coord(B))-(Coord(A) + Size - Gap - 1, Coord(B) + Size - Gap - 1), vbRed, B
        Picture1.Line (Coord(A), Coord(B))-(Coord(A) + Size - Gap - 1, Coord(B) + Size - Gap - 1), vbRed
        Picture1.Line (Coord(A), Coord(B) + Size - Gap - 1)-(Coord(A) + Size - Gap - 1, Coord(B)), vbRed
    End If
    Next A
    Next B
End If
KeyRoutine (KeyCode)

End Sub


Sub KeyRoutine(TheKey As Integer)


If Not GameOn Then Exit Sub

DestX = X
DestY = Y

Select Case TheKey

Case Is = vbKeyUp
    DestY = Y + 1
    B = 0

Case Is = vbKeyDown
    DestY = Y - 1
    B = 1

Case Is = vbKeyLeft
    DestX = X + 1
    B = 2

Case Is = vbKeyRight
    DestX = X - 1
    B = 3

End Select

If Grid(DestX, DestY) = 0 Then Exit Sub

MoveRecord = MoveRecord + T(B)
Grid(X, Y) = Grid(DestX, DestY)
Grid(DestX, DestY) = 0

If Smooth Then
    'Right
    If DestX < X Then
        For A = Coord(DestX) + 1 To Coord(X)
        Picture1.Line (Coord(DestX), Coord(DestY))-(A - 1, Coord(DestY) + Size - 1 - Gap), Colr, BF
        DrawBlock A, Coord(Y), Grid(X, Y)
        Next A
    End If
    'Down
    If DestY < Y Then
        For A = Coord(DestY) + 1 To Coord(Y)
        Picture1.Line (Coord(DestX), Coord(DestY))-(Coord(DestX) + Size - 1 - Gap, A - 1), Colr, BF
        DrawBlock Coord(X), A, Grid(X, Y)
        Next A
    End If
    'Left
    If DestX > X Then
        For A = Coord(DestX) To Coord(X) Step -1
        Picture1.Line (A + Size - 1, Coord(DestY))-(Coord(DestX) + Size - 1 - Gap, Coord(DestY) + Size - 1 - Gap), Colr, BF
        DrawBlock A, Coord(Y), Grid(X, Y)
        Next A
    End If
    'Up
    If DestY > Y Then
        For A = Coord(DestY) To Coord(Y) Step -1
        Picture1.Line (Coord(DestX), A + Size - 1)-(Coord(DestX) + Size - 1 - Gap, Coord(DestY) + Size - 1 - Gap), Colr, BF
        DrawBlock Coord(X), A, Grid(X, Y)
        Next A
    End If
Else
    Picture1.Line (Coord(DestX), Coord(DestY))-(Coord(DestX) + Size - 1 - Gap, Coord(DestY) + Size - 1 - Gap), Colr, BF
    DrawBlock Coord(X), Coord(Y), Grid(X, Y)
End If

X = DestX
Y = DestY

If CheckingOn Then CheckIfSolved

End Sub

Sub ResetScreen()

RF = False
Erase Grid
Picture1.Cls
Size = 480 \ NumOfTiles

For A = 1 To NumOfTiles
Coord(A) = Size * (A - 1)
Next A

C = 0
For B = 1 To NumOfTiles
For A = 1 To NumOfTiles
C = C + 1
Grid(A, B) = C
Next A
Next B

X = NumOfTiles
Y = NumOfTiles

Grid(X, Y) = 0

For B = 1 To Y
For A = 1 To X
If Not (A = X And B = Y) Then
    DrawBlock Coord(A), Coord(B), Grid(A, B)
End If
Next A
Next B

DrawCorner

RF = True

End Sub

Sub Scramble()

Scrambling = True
X = NumOfTiles
Y = NumOfTiles
MoveRecord = ""
Old = 100

Randomize Timer
For A = 1 To MixMoves


Do

DestX = X
DestY = Y

Jer:
B = Int(Rnd * 4)
If Old + B = 1 Or Old + B = 5 Then GoTo Jer

Select Case B

Case Is = 0
    DestY = Y + 1

Case Is = 1
    DestY = Y - 1

Case Is = 2
    DestX = X + 1

Case Is = 3
    DestX = X - 1

End Select

Loop While Grid(DestX, DestY) = 0

MoveRecord = MoveRecord & T(B)
Grid(X, Y) = Grid(DestX, DestY)
Grid(DestX, DestY) = 0

X = DestX
Y = DestY

Old = B

Next A

RefreshScreen

Scrambling = False

End Sub

Sub RefreshScreen()

Picture1.Cls
RF = False
For B = 1 To NumOfTiles
For A = 1 To NumOfTiles
If Grid(A, B) > 0 Then
    DrawBlock Coord(A), Coord(B), Grid(A, B)
End If
Next A
Next B

If Not GameOn And Not Scrambling Then DrawCorner

RF = True

Picture1.Refresh

End Sub

'(for debugging)
'Sub Fart()
'MsgBox Grid(1, 1) & Grid(2, 1) & Grid(3, 1) & vbCrLf & Grid(1, 2) & Grid(2, 2) _
'& Grid(3, 2) & vbCrLf & Grid(1, 3) & Grid(2, 3) & Grid(3, 3)
'End Sub

Sub CheckIfSolved()

C = 1
For B = 1 To NumOfTiles
For A = 1 To NumOfTiles

If Grid(A, B) = C Then
    C = C + 1
Else
    If Not (A = NumOfTiles And B = NumOfTiles) Then
        Exit Sub
    End If
End If

Next A
Next B

Command1.Caption = "Reset"
GameOn = False
JustWon = True
Label5.Visible = True
Command2.Enabled = False
DrawCorner

End Sub

Sub DrawCorner()
DrawBlock Coord(NumOfTiles), Coord(NumOfTiles), NumOfTiles ^ 2
End Sub

Sub LoadPic()
'Blocks = GenerateDC(File)
Picture2.Picture = LoadPicture(File)
ResetScreen
End Sub


