VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "TicTacToe"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3585
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3585
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Game"
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Game"
      Height          =   2295
      Left            =   2280
      TabIndex        =   10
      Top             =   240
      Width           =   1095
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Text            =   "text2"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Text            =   "text1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player2:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player1:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name's"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   525
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   1080
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   2280
      Width           =   45
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   240
      TabIndex        =   16
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   195
      Left            =   1440
      TabIndex        =   8
      Top             =   1200
      Width           =   105
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   195
      Left            =   1080
      TabIndex        =   7
      Top             =   1200
      Width           =   105
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   1200
      Width           =   105
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   105
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   105
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   105
   End
   Begin VB.Line Line8 
      X1              =   600
      X2              =   1680
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line7 
      X1              =   600
      X2              =   1680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line6 
      X1              =   1320
      X2              =   1320
      Y1              =   360
      Y2              =   1440
   End
   Begin VB.Line Line5 
      X1              =   960
      X2              =   960
      Y1              =   360
      Y2              =   1440
   End
   Begin VB.Line Line4 
      X1              =   3480
      X2              =   3480
      Y1              =   120
      Y2              =   2760
   End
   Begin VB.Line Line3 
      X1              =   3480
      X2              =   120
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line2 
      X1              =   3480
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   2760
      Y2              =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Save Names
SaveSetting App.EXEName, "Options", "Name1", Text1.Text
SaveSetting App.EXEName, "Options", "Name2", Text2.Text

'Set Label 1-9 to ""
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""

'Start Game
Label13.Caption = "It's your move " & Text1.Text
Label14.Caption = "Player1"

'Enable Timer
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Text1.Text = GetSetting(App.EXEName, "Options", "Name1", "Borg")
Text2.Text = GetSetting(App.EXEName, "Options", "Name2", "Roswell")
End Sub

Private Sub Label1_Click()
If Label1.Caption = "" Then
    If Label14.Caption = "Player1" Then
        Label1.Caption = "X"
        Label14.Caption = "Player2"
        Label13.Caption = "It's your move " & Text2.Text
    Else
        Label1.Caption = "O"
        Label14.Caption = "Player1"
        Label13.Caption = "It's your move " & Text1.Text
    End If
Else
    MsgBox "ilegal operation", vbCritical, "TicTacToe"
End If
End Sub

Private Sub Label2_Click()
If Label2.Caption = "" Then
    If Label14.Caption = "Player1" Then
        Label2.Caption = "X"
        Label14.Caption = "Player2"
        Label13.Caption = "It's your move " & Text2.Text
    Else
        Label2.Caption = "O"
        Label14.Caption = "Player1"
        Label13.Caption = "It's your move " & Text1.Text
    End If
Else
    MsgBox "ilegal operation", vbCritical, "TicTacToe"
End If
End Sub

Private Sub Label3_Click()
If Label3.Caption = "" Then
    If Label14.Caption = "Player1" Then
        Label3.Caption = "X"
        Label14.Caption = "Player2"
        Label13.Caption = "It's your move " & Text2.Text
    Else
        Label3.Caption = "O"
        Label14.Caption = "Player1"
        Label13.Caption = "It's your move " & Text1.Text
    End If
Else
    MsgBox "ilegal operation", vbCritical, "TicTacToe"
End If
End Sub

Private Sub Label4_Click()
If Label4.Caption = "" Then
    If Label14.Caption = "Player1" Then
        Label4.Caption = "X"
        Label14.Caption = "Player2"
        Label13.Caption = "It's your move " & Text2.Text
    Else
        Label4.Caption = "O"
        Label14.Caption = "Player1"
        Label13.Caption = "It's your move " & Text1.Text
    End If
Else
    MsgBox "ilegal operation", vbCritical, "TicTacToe"
End If
End Sub

Private Sub Label5_Click()
If Label5.Caption = "" Then
    If Label14.Caption = "Player1" Then
        Label5.Caption = "X"
        Label14.Caption = "Player2"
        Label13.Caption = "It's your move " & Text2.Text
    Else
        Label5.Caption = "O"
        Label14.Caption = "Player1"
        Label13.Caption = "It's your move " & Text1.Text
    End If
Else
    MsgBox "ilegal operation", vbCritical, "TicTacToe"
End If
End Sub

Private Sub Label6_Click()
If Label6.Caption = "" Then
    If Label14.Caption = "Player1" Then
        Label6.Caption = "X"
        Label14.Caption = "Player2"
        Label13.Caption = "It's your move " & Text2.Text
    Else
        Label6.Caption = "O"
        Label14.Caption = "Player1"
        Label13.Caption = "It's your move " & Text1.Text
    End If
Else
    MsgBox "ilegal operation", vbCritical, "TicTacToe"
End If
End Sub

Private Sub Label7_Click()
If Label7.Caption = "" Then
    If Label14.Caption = "Player1" Then
        Label7.Caption = "X"
        Label14.Caption = "Player2"
        Label13.Caption = "It's your move " & Text2.Text
    Else
        Label7.Caption = "O"
        Label14.Caption = "Player1"
        Label13.Caption = "It's your move " & Text1.Text
    End If
Else
    MsgBox "ilegal operation", vbCritical, "TicTacToe"
End If
End Sub

Private Sub Label8_Click()
If Label8.Caption = "" Then
    If Label14.Caption = "Player1" Then
        Label8.Caption = "X"
        Label14.Caption = "Player2"
        Label13.Caption = "It's your move " & Text2.Text
    Else
        Label8.Caption = "O"
        Label14.Caption = "Player1"
        Label13.Caption = "It's your move " & Text1.Text
    End If
Else
    MsgBox "ilegal operation", vbCritical, "TicTacToe"
End If
End Sub

Private Sub Label9_Click()
If Label9.Caption = "" Then
    If Label14.Caption = "Player1" Then
        Label9.Caption = "X"
        Label14.Caption = "Player2"
        Label13.Caption = "It's your move " & Text2.Text
    Else
        Label9.Caption = "O"
        Label14.Caption = "Player1"
        Label13.Caption = "It's your move " & Text1.Text
    End If
Else
    MsgBox "ilegal operation", vbCritical, "TicTacToe"
End If
End Sub

Private Sub Timer1_Timer()
If Label1.Caption = "X" And Label2.Caption = "X" And Label3.Caption = "X" Then
    Win (0)
End If
If Label1.Caption = "O" And Label2.Caption = "O" And Label3.Caption = "O" Then
    Win (1)
End If

If Label4.Caption = "X" And Label5.Caption = "X" And Label6.Caption = "X" Then
    Win (0)
End If
If Label4.Caption = "O" And Label5.Caption = "O" And Label6.Caption = "O" Then
    Win (1)
End If

If Label7.Caption = "X" And Label8.Caption = "X" And Label9.Caption = "X" Then
    Win (0)
End If
If Label7.Caption = "O" And Label8.Caption = "O" And Label9.Caption = "O" Then
    Win (1)
End If

If Label1.Caption = "X" And Label4.Caption = "X" And Label7.Caption = "X" Then
    Win (0)
End If
If Label1.Caption = "O" And Label4.Caption = "O" And Label7.Caption = "O" Then
    Win (1)
End If

If Label2.Caption = "X" And Label5.Caption = "X" And Label8.Caption = "X" Then
    Win (0)
End If
If Label2.Caption = "O" And Label5.Caption = "O" And Label8.Caption = "O" Then
    Win (1)
End If

If Label3.Caption = "X" And Label6.Caption = "X" And Label9.Caption = "X" Then
    Win (0)
End If
If Label3.Caption = "O" And Label6.Caption = "O" And Label9.Caption = "O" Then
    Win (1)
End If

If Label1.Caption = "X" And Label5.Caption = "X" And Label9.Caption = "X" Then
    Win (0)
End If
If Label1.Caption = "O" And Label5.Caption = "O" And Label9.Caption = "O" Then
    Win (1)
End If

If Label3.Caption = "X" And Label5.Caption = "X" And Label7.Caption = "X" Then
    Win (0)
End If
If Label3.Caption = "O" And Label5.Caption = "O" And Label7.Caption = "O" Then
    Win (1)
End If
End Sub
Sub Win(x)
If x = 0 Then
    MsgBox Text1.Text & " is the Winner!!", vbOKOnly, "TicTacToe"
Else
    MsgBox Text2.Text & " is the Winner!!", vbOKOnly, "TicTacToe"
End If
Timer1.Enabled = False
End Sub
