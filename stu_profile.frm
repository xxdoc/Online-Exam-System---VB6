VERSION 5.00
Begin VB.Form stu_profile 
   Caption         =   "STUDENT PROFILE - Madhu"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "stu_profile.frx":0000
   ScaleHeight     =   15690
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame stu_profile 
      Caption         =   "PROFILE CHANGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10095
      Left            =   7440
      TabIndex        =   0
      Top             =   3000
      Width           =   13455
      Begin VB.CommandButton cmdnext 
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
         Height          =   975
         Left            =   3960
         TabIndex        =   13
         Top             =   7560
         Width           =   5055
      End
      Begin VB.TextBox txtspwd 
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
         Left            =   6120
         TabIndex        =   10
         Top             =   4320
         Width           =   3015
      End
      Begin VB.TextBox txtsmail 
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
         Left            =   6120
         TabIndex        =   9
         Top             =   3480
         Width           =   3015
      End
      Begin VB.Label lbssec 
         Caption         =   "A SECTION"
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
         Left            =   6120
         TabIndex        =   12
         Top             =   6480
         Width           =   2175
      End
      Begin VB.Label lbssem 
         Caption         =   "IV BCA"
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
         Left            =   6120
         TabIndex        =   11
         Top             =   5400
         Width           =   3015
      End
      Begin VB.Label lbsreg 
         Caption         =   "15KXSB7051"
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
         Left            =   6000
         TabIndex        =   8
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label lbsname 
         Caption         =   "MADHU"
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
         Left            =   6120
         TabIndex        =   7
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label6 
         Caption         =   "SECTION :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   6480
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "SEMISTER :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "STUDENT PASSWORD :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   4440
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "STUDENT EMAIL :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   3480
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "REGISTER NUMBER :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   2640
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "STUDENT NAME :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   1
         Top             =   1680
         Width           =   3375
      End
   End
End
Attribute VB_Name = "stu_profile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmdnext_Click()
rs.Open "select * from profile where pname='" & stu_name & "'", con, adOpenDynamic, adLockOptimistic
rs("pemail") = txtsmail
rs("password") = txtspwd
rs("status") = "YES"
rs.Update
txtsmail = stu_email
MsgBox "Profile Updated Successfully"
txtsmail = ""
txtspwd = ""
rs.Close
Me.Hide
stu_menu.Show
End Sub

Private Sub Form_Load()
lbsname = stu_name
lbsreg = stu_reg
txtsmail = stu_email
lbssem = stu_class
lbssec = stu_sec
End Sub

