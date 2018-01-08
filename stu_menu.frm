VERSION 5.00
Begin VB.Form stu_menu 
   Caption         =   "s"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "stu_menu.frx":0000
   ScaleHeight     =   15690
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame f_stumenu 
      Caption         =   "STUDENT MENU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   1320
      TabIndex        =   0
      Top             =   4680
      Width           =   8895
      Begin VB.CommandButton cmdlogout 
         Caption         =   "LOGOUT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4680
         TabIndex        =   8
         Top             =   4200
         Width           =   3015
      End
      Begin VB.CommandButton cmdtake 
         Caption         =   "TAKE TEST"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1200
         TabIndex        =   7
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CommandButton cmdans 
         Caption         =   "ANSWER"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4680
         TabIndex        =   6
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CommandButton cmdresult 
         Caption         =   "RESULTS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1200
         TabIndex        =   5
         Top             =   2880
         Width           =   3015
      End
      Begin VB.CommandButton cmdtop 
         Caption         =   "TOP RANKS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4680
         TabIndex        =   4
         Top             =   2880
         Width           =   3015
      End
      Begin VB.CommandButton cmdprofile 
         Caption         =   "PROFILE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1200
         TabIndex        =   3
         Top             =   4200
         Width           =   3015
      End
      Begin VB.CommandButton cmdabt 
         Caption         =   "ABOUT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1200
         TabIndex        =   2
         Top             =   5520
         Width           =   3015
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4680
         TabIndex        =   1
         Top             =   5520
         Width           =   3015
      End
   End
End
Attribute VB_Name = "stu_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim t_taken As String


Private Sub cmdans_Click()
If t_astatus = "Locked" Then
MsgBox "Answer Not Announced"
Else
Unload stu_menu
stu_ans.Show
End If
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdlogout_Click()
Unload stu_menu
frm_Login.Show
End Sub

Private Sub cmdprofile_Click()
Unload stu_menu
stu_prf.Show
End Sub

Private Sub cmdresult_Click()
rs.Open "select * from results where stuname='" & stu_name & "'", con, adOpenDynamic, adLockPessimistic
If rs.EOF Then
rs.Close
MsgBox "NO RESULTS FOUND"
Exit Sub
Else
rs.Close
Unload stu_menu
stu_results.Show
End If
End Sub

Private Sub cmdtake_Click()
rs.Open "select * from results where stuname='" & stu_name & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
t_taken = "NO"
rs.Close
Else
t_taken = rs("tstatus")
rs.Close
End If

If t_tstatus = "Locked" Then
MsgBox "NO TEST ANNOUNCED"
ElseIf t_taken = "TAKEN" Then
MsgBox "ALREADY TAKEN TEST...!!!"
Else
Unload Me
stu_test.Show
End If
End Sub

Private Sub cmdtop_Click()
Unload stu_menu
stu_topranks.Show
End Sub

Private Sub Form_Load()
rs.Open "select * from testcontrol where sec='" & stu_sec & "' and sem='" & stu_class & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
t_tstatus = "Locked"
t_astatus = "Locked"
rs.Close
Else
t_name = rs("ttitle")
t_class = rs("sem")
t_sec = rs("sec")
t_dur = rs("duration")
t_id = rs("testid")
t_marks = rs("marks")
t_totalq = rs("totalq")
t_tstatus = rs("tstatus")
t_astatus = rs("astatus")
t_subject = rs("subject")
rs.Close
End If
End Sub
