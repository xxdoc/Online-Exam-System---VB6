VERSION 5.00
Begin VB.Form stu_ans 
   Caption         =   "ANSWER"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frm_stu_ans.frx":0000
   ScaleHeight     =   15690
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame f_ans 
      Caption         =   "ANSWER :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11055
      Left            =   4800
      TabIndex        =   0
      Top             =   1800
      Width           =   16935
      Begin VB.CommandButton cmdanext 
         Caption         =   "PREV"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   720
         TabIndex        =   3
         Top             =   9120
         Width           =   5295
      End
      Begin VB.CommandButton cmdaprev 
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   6600
         TabIndex        =   2
         Top             =   9120
         Width           =   5175
      End
      Begin VB.CommandButton cmdmenu 
         Caption         =   "MENU"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   12480
         TabIndex        =   1
         Top             =   9120
         Width           =   4095
      End
      Begin VB.Label lbans 
         Caption         =   "ANS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1800
         TabIndex        =   9
         Top             =   7320
         Width           =   12855
      End
      Begin VB.Label lbaq 
         Caption         =   "QUESTION"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   1680
         TabIndex        =   8
         Top             =   600
         Width           =   14655
      End
      Begin VB.Label lbaopt1 
         Caption         =   "ANSWER 1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         TabIndex        =   7
         Top             =   2640
         Width           =   15135
      End
      Begin VB.Label lbaopt2 
         Caption         =   "ANSWER 2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1800
         TabIndex        =   6
         Top             =   3720
         Width           =   15135
      End
      Begin VB.Label lbaopt3 
         Caption         =   "ANSWER 3"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1800
         TabIndex        =   5
         Top             =   4800
         Width           =   14415
      End
      Begin VB.Label lbaopt4 
         Caption         =   "ANSWER 4"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1800
         TabIndex        =   4
         Top             =   5880
         Width           =   15135
      End
   End
End
Attribute VB_Name = "stu_ans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmdanext_Click()
If Not rs.BOF And Not rs.EOF Then
    rs.MoveNext
    If rs.EOF Then
        rs.MoveFirst
    End If
    lbaq = rs("questions")
    lbaopt1 = rs("opt1")
    lbaopt2 = rs("opt2")
    lbaopt3 = rs("opt3")
    lbaopt4 = rs("opt4")
    lbans = "Answer is " & rs("optans")
End If
End Sub

Private Sub cmdaprev_Click()
If Not rs.BOF And Not rs.EOF Then
    rs.MovePrevious
    If rs.BOF Then
        rs.MoveLast
    End If
    lbaq = rs("questions")
    lbaopt1 = rs("opt1")
    lbaopt2 = rs("opt2")
    lbaopt3 = rs("opt3")
    lbaopt4 = rs("opt4")
    lbans = "Answer is " & rs("optans")
End If

End Sub

Private Sub cmdmenu_Click()
rs.Close
Unload stu_ans
stu_menu.Show
End Sub

Private Sub Form_Load()
rs.Open "select * from test where testid='" & t_id & "'", con, adOpenDynamic, adLockPessimistic

rs.MoveFirst

lbaq = rs("questions")
    lbaopt1 = rs("opt1")
    lbaopt2 = rs("opt2")
    lbaopt3 = rs("opt3")
    lbaopt4 = rs("opt4")
    lbans = "Answer is " & rs("optans")
End Sub

