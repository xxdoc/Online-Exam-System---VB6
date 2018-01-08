VERSION 5.00
Begin VB.Form stu_results 
   Caption         =   "Results"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "stu_results.frx":0000
   ScaleHeight     =   15690
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "MARKS OBTAINED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12015
      Left            =   7080
      TabIndex        =   0
      Top             =   840
      Width           =   14415
      Begin VB.CommandButton cmdmenu 
         Caption         =   "MENU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10680
         TabIndex        =   1
         Top             =   10560
         Width           =   2895
      End
      Begin VB.Label lbmfeedback 
         Caption         =   "DIFICULT "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   6720
         TabIndex        =   12
         Top             =   5160
         Width           =   7815
      End
      Begin VB.Label Label15 
         Caption         =   "FEEDBACK :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Top             =   5160
         Width           =   2295
      End
      Begin VB.Label lbmpoints 
         Caption         =   "2.5 /  100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   10
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "POINTS :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         TabIndex        =   9
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label lbmmarks 
         Caption         =   "18 / 50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         TabIndex        =   8
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "MARKS :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label lbmmail 
         Caption         =   "Karthikjr@gmail.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6720
         TabIndex        =   6
         Top             =   3120
         Width           =   3495
      End
      Begin VB.Label Label9 
         Caption         =   "E-MAIL :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         TabIndex        =   5
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lbmstuname 
         Caption         =   "KARTHIK K R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         TabIndex        =   4
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Label Label7 
         Caption         =   "STUDENT NAME :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   2520
         Width           =   3375
      End
      Begin VB.Label lb_ttitle 
         Caption         =   "RESULTS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6120
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
   End
End
Attribute VB_Name = "stu_results"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmdmenu_Click()
Unload stu_results
stu_menu.Show
End Sub

Private Sub Form_Load()
rs.Open "select * from results where stuname='" & stu_name & "'", con, adOpenDynamic, adLockPessimistic
lbmstuname.Caption = rs("stuname")
lbmmail = rs("stumail")
lbmmarks = rs("marks")
lbmfeedback = rs("feedback")
lb_ttitle.Caption = rs("testtitle")
lbmpoints = rs("points")
rs.Close
End Sub
