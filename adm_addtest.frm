VERSION 5.00
Begin VB.Form frm_addtest 
   Caption         =   "ADD TEST - Online EXAM"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "adm_addtest.frx":0000
   ScaleHeight     =   15690
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame f_submit 
      BackColor       =   &H80000010&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3735
      Left            =   8280
      TabIndex        =   50
      Top             =   10320
      Width           =   6375
      Begin VB.CommandButton cmdmenu 
         Caption         =   "MENU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1080
         TabIndex        =   52
         Top             =   720
         Width           =   4095
      End
      Begin VB.CommandButton cmdtsubmit 
         Caption         =   "SUMBIT TEST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   960
         TabIndex        =   51
         Top             =   2400
         Width           =   4215
      End
   End
   Begin VB.Frame f_tc 
      BackColor       =   &H80000010&
      Caption         =   "TIME CONTROL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3615
      Left            =   8280
      TabIndex        =   46
      Top             =   6480
      Width           =   6375
      Begin VB.CommandButton cmd_timec 
         Caption         =   "COMFRIM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3000
         TabIndex        =   48
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txttime 
         Height          =   855
         Left            =   1920
         TabIndex        =   47
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "TIME CONTROL :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   240
         TabIndex        =   49
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame f_astatus 
      BackColor       =   &H80000010&
      Caption         =   "ANSWER STATUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3255
      Left            =   8280
      TabIndex        =   42
      Top             =   2880
      Width           =   6375
      Begin VB.CommandButton cmd_astatus 
         Caption         =   "COMFRIM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3120
         TabIndex        =   44
         Top             =   2160
         Width           =   2775
      End
      Begin VB.ComboBox castatus 
         Height          =   315
         Left            =   2280
         TabIndex        =   43
         Text            =   "SELECT STATUS"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "ANSWER STATUS :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   735
         Left            =   360
         TabIndex        =   45
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame f_tstatus 
      BackColor       =   &H80000010&
      Caption         =   "TEST STATUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2535
      Left            =   8280
      TabIndex        =   38
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmd_tsatus 
         Caption         =   "CONFRIM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3240
         TabIndex        =   40
         Top             =   1440
         Width           =   2775
      End
      Begin VB.ComboBox ctstatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2400
         TabIndex        =   39
         Text            =   "SELECT STATUS"
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "TEST STATUS :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   360
         TabIndex        =   41
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame f_tq 
      BackColor       =   &H80000010&
      Caption         =   "NO OF QUESTIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4095
      Left            =   120
      TabIndex        =   32
      Top             =   5760
      Width           =   7815
      Begin VB.CommandButton cmdqcfm 
         BackColor       =   &H8000000A&
         Caption         =   "CONFIRM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3360
         Width           =   3255
      End
      Begin VB.TextBox txtnoq 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         TabIndex        =   34
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox txt_qmark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         TabIndex        =   33
         Top             =   2160
         Width           =   5055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NO OF QUESTIONS :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "MARKS PER QUESTION :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   120
         TabIndex        =   36
         Top             =   2280
         Width           =   2175
      End
   End
   Begin VB.Frame f_tforms 
      BackColor       =   &H80000010&
      Caption         =   "TEST FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   11535
      Left            =   14880
      TabIndex        =   15
      Top             =   1200
      Width           =   12855
      Begin VB.CommandButton cmdqupdate 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         TabIndex        =   24
         Top             =   10200
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.CommandButton cmdqnext 
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9720
         TabIndex        =   23
         Top             =   10200
         Width           =   2415
      End
      Begin VB.CommandButton cmdqprev 
         Caption         =   "PREV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   22
         Top             =   10200
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtopt4 
         Height          =   615
         Left            =   2400
         TabIndex        =   21
         Top             =   6840
         Width           =   10215
      End
      Begin VB.TextBox txtopt3 
         Height          =   615
         Left            =   2400
         TabIndex        =   20
         Top             =   5760
         Width           =   10215
      End
      Begin VB.TextBox txtopt2 
         Height          =   615
         Left            =   2400
         TabIndex        =   19
         Top             =   4680
         Width           =   10215
      End
      Begin VB.TextBox txtopt1 
         Height          =   615
         Left            =   2400
         TabIndex        =   18
         Top             =   3360
         Width           =   10215
      End
      Begin VB.TextBox txtques 
         Height          =   2295
         Left            =   2520
         TabIndex        =   17
         Top             =   480
         Width           =   10095
      End
      Begin VB.ComboBox cmb_answer 
         Height          =   315
         Left            =   2400
         TabIndex        =   16
         Top             =   8400
         Width           =   6855
      End
      Begin VB.Label lbqno 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   495
         Left            =   5040
         TabIndex        =   31
         Top             =   9240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "ANSWER :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   360
         TabIndex        =   30
         Top             =   8400
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "OPTION 4 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   360
         TabIndex        =   29
         Top             =   6960
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "OPTION 3 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "OPTION 2 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "OPTION 1 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "QUESTIONS :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame f_tdetails 
      BackColor       =   &H80000010&
      Caption         =   "TEST DESTAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   5175
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   7815
      Begin VB.TextBox txt_sub 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   9
         Top             =   1440
         Width           =   3495
      End
      Begin VB.ComboBox cmb_sec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         TabIndex        =   8
         Text            =   "SELECT SECTION"
         Top             =   3120
         Width           =   3495
      End
      Begin VB.ComboBox cmb_sem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         TabIndex        =   7
         Text            =   "SELECT SEM"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CommandButton cmd_submit 
         BackColor       =   &H8000000A&
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4080
         Width           =   3135
      End
      Begin VB.CommandButton cmd_cancel 
         BackColor       =   &H8000000A&
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4080
         Width           =   3015
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "FACULTY : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lb_fname 
         BackStyle       =   0  'Transparent
         Caption         =   "Test"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "SUBJECT :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "SECTION :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label sem 
         BackStyle       =   0  'Transparent
         Caption         =   "SEMINSTER :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2400
         Width           =   1815
      End
   End
   Begin VB.Frame f_ttitle 
      BackColor       =   &H80000010&
      Caption         =   "TEST TITLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   10080
      Width           =   7815
      Begin VB.TextBox txt_ttitle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   2
         Top             =   1320
         Width           =   5055
      End
      Begin VB.CommandButton cmd_ttitle 
         Caption         =   "CONFIRM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   1
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label lb_ttitle 
         BackStyle       =   0  'Transparent
         Caption         =   "TEST TITLE :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1440
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm_addtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ttitle, tsub As String
Dim tqno, qmarks As Integer
Dim i, testid As Integer
Dim rs As New ADODB.Recordset
Dim ttsub, tfname, tsem, tsec, timec As String
Dim astatus, tstatus As String

Private Sub cmd_astatus_Click()
' ------------------------------------ aststus is used to Enable answer control ------------------------------

If castatus.Text <> "SELECT STATUS" Then
f_tc.Visible = True
astatus = castatus.Text
Else
MsgBox "Select Answer Status"
End If
End Sub

Private Sub cmd_cancel_Click()
admin.Show
Me.Hide
End Sub

Private Sub cmd_submit_Click()
' ------------------------------------ cmd_submit is used to submit the Test details  ------------------------------

If txt_sub <> "" Then
ttsub = txt_sub
Else
MsgBox "Enter Subject Name"
Exit Sub
End If

If cmb_sem <> "SELECT SEM" Then
tsem = cmb_sem.Text
Else
MsgBox "Select Semister"
f_tq.Visible = False
Exit Sub
End If

If cmb_sec <> "SELECT SECTION" Then
tsec = cmb_sec.Text
Else
MsgBox "Select SECTION"
f_tq.Visible = False
Exit Sub
End If

If cmb_sec = a_sec And cmb_sem = a_sem Then
MsgBox "Delete Previous Test", vbInformation
f_tq.Visible = False
Exit Sub
End If


tfname = pusername
f_tq.Visible = True

End Sub


Private Sub cmd_timec_Click()
' ------------------------------------ cmd_timec is used to set test duration control  ------------------------------

If txttime <> "" Then
f_submit.Visible = True
f_tforms.Visible = True
timec = txttime
cmdmenu.Enabled = False
cmdtsubmit.Enabled = False
Else
MsgBox "enter Test Duration"
End If
End Sub

Private Sub cmd_tsatus_Click()
' ------------------------------------ cmd_tstatus is used to Enable to test or not ------------------------------

If ctstatus.Text <> "SELECT STATUS" Then
f_astatus.Visible = True
tstatus = ctstatus.Text
Else
MsgBox "Select Text Status"
End If
End Sub

Private Sub cmd_ttitle_Click()
' ------------------------------------ cmd_title is used to set test title ------------------------------
If txt_ttitle <> "" Then
ttitle = txt_ttitle
f_tstatus.Visible = True
Else
MsgBox "Enter Test Title"
End If
End Sub

Private Sub cmdtsubmit_Click()
' ------------------------------------ cmdsubmit is used to submit the Test ------------------------------
cmdmenu.Enabled = True
cmdtsubmit.Enabled = False

rs.Open "select * from testcontrol", con, adOpenDynamic, adLockOptimistic
rs.AddNew
rs("duration") = timec
rs("subject") = ttsub
rs("ttitle") = ttitle
rs("testid") = testid
rs("astatus") = astatus
rs("tstatus") = tstatus
rs("totalq") = tqno
rs("sem") = tsem
rs("sec") = tsec
rs("faculty") = tfname
rs("marks") = qmarks
rs.Update
rs.Close

End Sub

Private Sub Form_Load()

cmdqprev.Enabled = False

' ------------------------------------ Adding text to combo box  ------------------------------

ctstatus.AddItem "Locked"
ctstatus.AddItem "UnLocked"

castatus.AddItem "Locked"
castatus.AddItem "UnLocked"

cmb_answer.AddItem 1
cmb_answer.AddItem 2
cmb_answer.AddItem 3
cmb_answer.AddItem 4

cmb_sem.AddItem "I BCA"
cmb_sem.AddItem "II BCA"
cmb_sem.AddItem "III BCA"
cmb_sem.AddItem "IV BCA"
cmb_sem.AddItem "V BCA"
cmb_sem.AddItem "VI BCA"

cmb_sec.AddItem "A"
cmb_sec.AddItem "B"

lb_fname.Caption = fname
f_tq.Visible = False
f_tforms.Visible = False
f_tstatus.Visible = False
f_astatus.Visible = False
f_tc.Visible = False
f_ttitle.Visible = False
f_submit.Visible = False

i = 1

' ------------------------------------ generating random number for testid  ------------------------------
testid = Int((10000 * Rnd) + 1)
End Sub

Private Sub cmdmenu_Click()
clear
f_tq.Visible = False
f_ttitle.Visible = False
f_tstatus.Visible = False
f_tc.Visible = False
f_tforms.Visible = False
Unload Me
admin.Show
End Sub

Private Sub cmdqcfm_Click()
' ------------------------------------ This Frame is used for question control   ------------------------------

If txtnoq <> "" Then
tqno = Val(txtnoq)
lbqno.Caption = tqno & i
Else
MsgBox "Enter Total Questions"
Exit Sub
End If

If txt_qmark <> "" Then
qmarks = Val(txt_qmark)
Else
MsgBox "Enter Marks Per Question"
Exit Sub
End If
f_ttitle.Visible = True
End Sub

Private Sub cmdqnext_Click()
cmdqprev.Enabled = True
' ------------------------------------ cmdqnext is used to store the each to database  ------------------------------

lbqno.Caption = tqno & i
If i >= tqno + 1 Then
MsgBox "Entered All Questions"
cmdtsubmit.Enabled = True
Else
rs.Open "select * from test", con, adOpenDynamic, adLockPessimistic
rs.AddNew
rs("questions") = txtques
rs("opt1") = txtopt1
rs("opt2") = txtopt2
rs("opt3") = txtopt3
rs("opt4") = txtopt4
rs("optans") = cmb_answer.Text
rs("testid") = testid
rs.Update
rs.MoveNext
rs.Close
txtques = ""
txtopt1 = ""
txtopt2 = ""
txtopt3 = ""
txtopt4 = ""
txtans = ""
i = i + 1
End If
End Sub

Private Sub cmdqprev_Click()
' ------------------------------------ cmdqprev is used to go previous for updating the data  ------------------------------
rs.Open "select * from test where testid='" & testid & "'", con, adOpenDynamic, adLockPessimistic

rs.MovePrevious

If rs.EOF Then
rs.MoveFirst
End If

If rs.BOF Then
rs.MoveNext
End If

txtques = rs("questions")
txtopt1 = rs("opt1")
txtopt2 = rs("opt2")
txtopt3 = rs("opt3")
txtopt4 = rs("opt4")
cmb_answer.Text = rs("optans")

rs.Close
End Sub

Private Sub cmdttle_Click()
' ------------------------------------ set the test title  ------------------------------
ttitle = txtttitle
End Sub

Private Function clear()
' ------------------------------------ Function is used to clear the text box values  ------------------------------
txtques = ""
opt1 = ""
opt2 = ""
opt3 = ""
opt4 = ""
ans = ""
End Function
