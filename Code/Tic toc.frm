VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6825
   Icon            =   "Tic toc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "ABOUT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   1680
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   120
      Picture         =   "Tic toc.frx":048A
      ScaleHeight     =   1395
      ScaleWidth      =   6555
      TabIndex        =   19
      Top             =   120
      Width           =   6615
   End
   Begin VB.CommandButton c1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   8
      Left            =   3120
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton c1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   1680
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton c1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   360
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton c1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   5
      Left            =   3120
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton c1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   4
      Left            =   1680
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton c1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   360
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton c1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   3120
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton c1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   1680
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton c1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   360
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "RESULT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4560
      TabIndex        =   7
      Top             =   3600
      Width           =   2055
      Begin VB.TextBox r2 
         BackColor       =   &H0000FF00&
         Height          =   375
         Left            =   600
         TabIndex        =   22
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox r1 
         BackColor       =   &H0000FF00&
         Height          =   375
         Left            =   600
         TabIndex        =   21
         Top             =   840
         Width           =   615
      End
      Begin VB.Label l6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label l5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1695
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   2040
         X2              =   0
         Y1              =   1320
         Y2              =   1320
      End
   End
   Begin VB.CommandButton restart 
      Caption         =   "RESTART"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox p2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox p1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Line Line8 
      BorderWidth     =   6
      X1              =   360
      X2              =   4440
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line7 
      BorderWidth     =   6
      X1              =   360
      X2              =   4440
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line6 
      BorderWidth     =   6
      X1              =   4440
      X2              =   4440
      Y1              =   6720
      Y2              =   3720
   End
   Begin VB.Line Line5 
      BorderWidth     =   6
      X1              =   360
      X2              =   360
      Y1              =   6720
      Y2              =   3720
   End
   Begin VB.Line Line4 
      BorderStyle     =   3  'Dot
      X1              =   360
      X2              =   4440
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  'Dot
      X1              =   360
      X2              =   4440
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   3120
      X2              =   3120
      Y1              =   3720
      Y2              =   6720
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   1680
      X2              =   1680
      Y1              =   3720
      Y2              =   6720
   End
   Begin VB.Label Label4 
      Caption         =   " O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   5
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   " X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "  Player 2 Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "  Player 1 Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ponewin As Integer
Dim ptwowin As Integer
Dim sign As Integer
Dim i As Integer
Dim j As Integer
Dim f As Integer
Dim ctr As Integer




Private Sub c1_Click(Index As Integer)
ctr = 0
If sign = 0 Then
c1(Index).Caption = "X"
c1(Index).Enabled = False
sign = 1

Else
c1(Index).Caption = "O"
c1(Index).Enabled = False

sign = 0
End If
For i = 0 To 6 Step 3
If c1(i).Caption = c1(i + 1).Caption And c1(i + 1).Caption = c1(i + 2).Caption And c1(i).Caption <> "" Then
f = 1
If c1(i).Caption = "X" Then
ponewin = ponewin + 1
r1.Text = ponewin
MsgBox "" & p1.Text & " Wins."
Else
ptwowin = ptwowin + 1
r2.Text = ptwowin
MsgBox "" & p2.Text & " Wins."
End If
For j = 0 To 8
c1(j).Caption = ""
c1(j).Enabled = True
sign = 0
Next j
End If
Next i



For i = 0 To 2
If c1(i).Caption = c1(i + 3).Caption And c1(i + 3).Caption = c1(i + 6).Caption And c1(i).Caption <> "" Then
f = 1
If c1(i).Caption = "X" Then
ponewin = ponewin + 1
r1.Text = ponewin
MsgBox "" & p1.Text & " Wins."
Else
ptwowin = ptwowin + 1
r2.Text = ptwowin
MsgBox "" & p2.Text & " Wins."
End If
For j = 0 To 8
c1(j).Caption = ""
sign = 0
c1(j).Enabled = True
'sign = 0
Next j
End If
Next i




If c1(0).Caption = c1(4).Caption And c1(4).Caption = c1(8).Caption And c1(0).Caption <> "" Then
f = 1
If c1(0).Caption = "X" Then
ponewin = ponewin + 1
r1.Text = ponewin
MsgBox "" & p1.Text & " Wins."
Else
ptwowin = ptwowin + 1
r2.Text = ptwowin
MsgBox "" & p2.Text & " Wins."
End If
'MsgBox "Sign - " & c1(0).Caption & " Wins."
'f = 1
For j = 0 To 8
c1(j).Caption = ""
sign = 0
c1(j).Enabled = True
'sign = 0
Next j
End If



If c1(2).Caption = c1(4).Caption And c1(4).Caption = c1(6).Caption And c1(2).Caption <> "" Then
f = 1
If c1(2).Caption = "X" Then
ponewin = ponewin + 1
r1.Text = ponewin
MsgBox "" & p1.Text & " Wins."
Else
ptwowin = ptwowin + 1
r2.Text = ptwowin
MsgBox "" & p2.Text & " Wins."
End If
'MsgBox "Sign - " & c1(2).Caption & " Wins."
'f = 1
For j = 0 To 8
c1(j).Caption = ""
sign = 0
c1(j).Enabled = True
'sign = 0
Next j
End If



For j = 0 To 8
If c1(j).Caption <> "" Then
ctr = ctr + 1
End If
Next j

If ctr = 9 Then
r1.Text = ponewin
r2.Text = ptwowin
MsgBox "Oops Match Drawn Try Again"

For j = 0 To 8
c1(j).Caption = ""
sign = 0
c1(j).Enabled = True
'sign = 0
Next j
End If

End Sub

Private Sub Command1_Click()
MsgBox "CREATED BY ----SHUBHAM SAGAR----"
MsgBox "For any help Mail us at ----shubhamth*****s@gmail.com---- or contact us on ----7079****77----"

End Sub

Private Sub Form_Load()
ponewin = 0
ptwowin = 0
sign = 0
f = 0
ctr = 0

End Sub



Private Sub p1_Change()
l5.Caption = p1.Text

End Sub

Private Sub p2_Change()
l6.Caption = p2.Text
End Sub



Private Sub r1_Change()
r1.Text = ponewin

End Sub

Private Sub restart_Click()
sign = 0
For i = 0 To 8 Step 1
c1(i).Caption = ""
c1(i).Enabled = True

Next i
sign = sign + 1
End Sub
