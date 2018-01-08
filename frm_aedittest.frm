VERSION 5.00
Begin VB.Form frm_aedittest 
   Caption         =   "EDIT TEST"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20640
   LinkTopic       =   "Form1"
   Picture         =   "frm_aedittest.frx":0000
   ScaleHeight     =   15690
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame f_submit 
      BackColor       =   &H80000013&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   8280
      TabIndex        =   50
      Top             =   10680
      Width           =   6375
      Begin VB.CommandButton cmdmenu 
         Caption         =   "MENU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1080
         TabIndex        =   51
         Top             =   2400
         Width           =   4215
      End
   End
   Begin VB.Frame f_tc 
      BackColor       =   &H80000013&
      Caption         =   "TIME CONTROL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   8280
      TabIndex        =   46
      Top             =   6840
      Width           =   6375
      Begin VB.CommandButton cmd_timec 
         Caption         =   "COMFRIM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   48
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txttime 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3120
         TabIndex        =   47
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "TIME CONTROL :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   49
         Top             =   960
         Width           =   2415
      End
   End
   Begin VB.Frame f_astatus 
      BackColor       =   &H80000013&
      Caption         =   "ANSWER STATUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   8280
      TabIndex        =   42
      Top             =   3240
      Width           =   6375
      Begin VB.CommandButton cmd_astatus 
         Caption         =   "COMFRIM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   44
         Top             =   2040
         Width           =   2775
      End
      Begin VB.ComboBox castatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3120
         TabIndex        =   43
         Text            =   "SELECT STATUS"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "ANSWER STATUS :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   45
         Top             =   1080
         Width           =   2895
      End
   End
   Begin VB.Frame f_tstatus 
      BackColor       =   &H80000013&
      Caption         =   "TEST STATUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   8280
      TabIndex        =   38
      Top             =   480
      Width           =   6375
      Begin VB.CommandButton cmd_tsatus 
         Caption         =   "CONFRIM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         TabIndex        =   40
         Top             =   1440
         Width           =   3255
      End
      Begin VB.ComboBox ctstatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2400
         TabIndex        =   39
         Text            =   "SELECT STATUS"
         Top             =   600
         Width           =   3255
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
         Height          =   375
         Left            =   360
         TabIndex        =   41
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame f_tq 
      BackColor       =   &H80000013&
      Caption         =   "NO OF QUESTIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   32
      Top             =   6120
      Width           =   7815
      Begin VB.CommandButton cmdqcfm 
         Caption         =   "CONFIRM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         TabIndex        =   35
         Top             =   3000
         Width           =   3015
      End
      Begin VB.TextBox txtnoq 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   34
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txt_qmark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   33
         Top             =   2160
         Width           =   3855
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
         Height          =   375
         Left            =   600
         TabIndex        =   37
         Top             =   1080
         Width           =   2655
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
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   2280
         Width           =   3495
      End
   End
   Begin VB.Frame f_tforms 
      BackColor       =   &H80000013&
      Caption         =   "TEST FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11535
      Left            =   15120
      TabIndex        =   15
      Top             =   1800
      Width           =   12855
      Begin VB.CommandButton cmdqupdate 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         TabIndex        =   24
         Top             =   10200
         Width           =   4095
      End
      Begin VB.CommandButton cmdqnext 
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   22
         Top             =   10200
         Width           =   2175
      End
      Begin VB.TextBox txtopt4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2640
         TabIndex        =   21
         Top             =   6840
         Width           =   9495
      End
      Begin VB.TextBox txtopt3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   20
         Top             =   5760
         Width           =   9495
      End
      Begin VB.TextBox txtopt2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   19
         Top             =   4680
         Width           =   9495
      End
      Begin VB.TextBox txtopt1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   18
         Top             =   3360
         Width           =   9495
      End
      Begin VB.TextBox txtques 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   2640
         TabIndex        =   17
         Top             =   480
         Width           =   9495
      End
      Begin VB.ComboBox cmb_answer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2640
         TabIndex        =   16
         Top             =   8400
         Width           =   5895
      End
      Begin VB.Label lbqno 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         TabIndex        =   31
         Top             =   9240
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "ANSWER :"
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
         Left            =   240
         TabIndex        =   30
         Top             =   8400
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "OPTION 4 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   29
         Top             =   6960
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "OPTION 3 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   28
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "OPTION 2 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   27
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "OPTION 1 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   26
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "QUESTION :"
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
         Left            =   360
         TabIndex        =   25
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Frame f_tdetails 
      BackColor       =   &H80000013&
      Caption         =   "TEST DESTAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   7815
      Begin VB.TextBox txt_sub 
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
         Left            =   2760
         TabIndex        =   9
         Top             =   1440
         Width           =   3495
      End
      Begin VB.ComboBox cmb_sec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2760
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2760
         TabIndex        =   7
         Text            =   "SELECT SEM"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CommandButton cmd_submit 
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   6
         Top             =   4080
         Width           =   3015
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4080
         TabIndex        =   5
         Top             =   4080
         Width           =   2655
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
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lb_fname 
         BackStyle       =   0  'Transparent
         Caption         =   "Test"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
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
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   1560
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
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   3120
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
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   2400
         Width           =   1815
      End
   End
   Begin VB.Frame f_ttitle 
      BackColor       =   &H80000013&
      Caption         =   "TEST TITLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   10440
      Width           =   7815
      Begin VB.TextBox txt_ttitle 
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
         Left            =   2040
         TabIndex        =   2
         Top             =   1320
         Width           =   5175
      End
      Begin VB.CommandButton cmd_ttitle 
         Caption         =   "CONFIRM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         TabIndex        =   1
         Top             =   2880
         Width           =   3255
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
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
   End
   Begin VB.Label lb_ttid 
      BackStyle       =   0  'Transparent
      Caption         =   "7000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   18840
      TabIndex        =   54
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "TEST ID :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   17160
      TabIndex        =   53
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frm_aedittest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ttitle, tsub As String
Dim tqno, qmarks As Integer
Dim i, testid As Integer
Dim rs  As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
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

rs1.Open "select * from testcontrol where testid='" & t_id & "'", con, adOpenDynamic, adLockOptimistic
rs1("astatus") = castatus.Text
rs1.Update
MsgBox "Updated Successfully"
rs1.Close
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

tfname = pusername

rs1.Open "select * from testcontrol where testid='" & t_id & "'", con, adOpenDynamic, adLockOptimistic
rs1("subject") = txt_sub
rs1("sem") = cmb_sem.Text
rs1("sec") = cmb_sec.Text
rs1.Update
MsgBox "Updated Successfully"
rs1.Close

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

rs1.Open "select * from testcontrol where testid='" & t_id & "'", con, adOpenDynamic, adLockOptimistic
rs1("duration") = txttime
rs1.Update
MsgBox "Updated Successfully"
rs1.Close
End Sub

Private Sub cmd_tsatus_Click()
' ------------------------------------ cmd_tstatus is used to Enable to test or not ------------------------------

If ctstatus.Text <> "SELECT STATUS" Then
f_astatus.Visible = True
tstatus = ctstatus.Text
Else
MsgBox "Select Text Status"
End If

rs1.Open "select * from testcontrol where testid='" & t_id & "'", con, adOpenDynamic, adLockOptimistic
rs1("tstatus") = ctstatus.Text
rs1.Update
MsgBox "Updated Successfully"
rs1.Close

End Sub

Private Sub cmd_ttitle_Click()
' ------------------------------------ cmd_title is used to set test title ------------------------------
If txt_ttitle <> "" Then
ttitle = txt_ttitle
f_tstatus.Visible = True
Else
MsgBox "Enter Test Title"
End If

rs1.Open "select * from testcontrol where testid='" & t_id & "'", con, adOpenDynamic, adLockOptimistic
rs1("ttitle") = txt_ttitle
rs1.Update
MsgBox "Updated Successfully"
rs1.Close

End Sub

Private Sub cmdqupdate_Click()
rs("questions") = txtques
rs("opt1") = txtopt1
rs("opt2") = txtopt2
rs("opt3") = txtopt3
rs("opt4") = txtopt4
rs("optans") = cmb_answer.Text
rs.Update
MsgBox "Updated successfully"
End Sub

Private Sub cmdtsubmit_Click()
' ------------------------------------ cmdsubmit is used to submit the Test ------------------------------
rs.Close
MsgBox "Submited Successfully"
End Sub
Private Sub Form_Load()
rs.Open "select * from testcontrol where faculty='" & pusername & "'", con, adOpenDynamic, adLockOptimistic
t_dur = rs("duration")
t_subject = rs("subject")
t_name = rs("ttitle")
t_id = rs("testid")
t_astatus = rs("astatus")
t_tstatus = rs("tstatus")
t_totalq = rs("totalq")
t_class = rs("sem")
t_sec = rs("sec")
tfname = rs("faculty")
t_marks = rs("marks")
rs.Close

rs.Open "select * from test where testid='" & t_id & "'", con, adOpenDynamic, adLockPessimistic
rs.MoveFirst
txtques = rs("questions")
txtopt1 = rs("opt1")
txtopt2 = rs("opt2")
txtopt3 = rs("opt3")
txtopt4 = rs("opt4")
cmb_answer.Text = rs("optans")


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
f_tq.Visible = True
f_tforms.Visible = True
f_tstatus.Visible = True
f_astatus.Visible = True
f_tc.Visible = True
f_ttitle.Visible = True
f_submit.Visible = True

i = 1

castatus.Text = t_astatus
ctstatus.Text = t_tstatus
cmb_sem.Text = t_class
cmb_sec.Text = t_sec

txt_sub = t_subject
txtnoq = t_totalq
txt_qmark = t_marks
txt_ttitle = t_name
txttime = t_dur
lb_ttid = t_id

End Sub

Private Sub cmdmenu_Click()
rs.Close
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

rs1.Open "select * from testcontrol where testid='" & t_id & "'", con, adOpenDynamic, adLockOptimistic
rs1("totalq") = txtnoq
rs1("marks") = txt_qmark
rs1.Update
MsgBox "Updated Successfully"
rs1.Close

End Sub

Private Sub cmdqnext_Click()
' ------------------------------------ cmdqnext is used to store the each to database  ------------------------------
If Not rs.BOF And Not rs.EOF Then
    rs.MovePrevious
    If rs.BOF Then
        rs.MoveLast
    End If
    txtques = rs("questions")
txtopt1 = rs("opt1")
txtopt2 = rs("opt2")
txtopt3 = rs("opt3")
txtopt4 = rs("opt4")
cmb_answer.Text = rs("optans")
End If
End Sub

Private Sub cmdqprev_Click()
' ------------------------------------ cmdqprev is used to go previous for updating the data  ------------------------------
If Not rs.BOF And Not rs.EOF Then
    rs.MoveNext
    If rs.EOF Then
        rs.MoveFirst
    End If
    txtques = rs("questions")
txtopt1 = rs("opt1")
txtopt2 = rs("opt2")
txtopt3 = rs("opt3")
txtopt4 = rs("opt4")
cmb_answer.Text = rs("optans")
End If
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

