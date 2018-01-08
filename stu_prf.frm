VERSION 5.00
Begin VB.Form stu_prf 
   Caption         =   "Profile"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "stu_prf.frx":0000
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
      Left            =   6600
      TabIndex        =   0
      Top             =   2400
      Width           =   13455
      Begin VB.CommandButton cmd_menu 
         Caption         =   "MENU"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7320
         TabIndex        =   14
         Top             =   8760
         Width           =   4815
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
         Left            =   6000
         TabIndex        =   3
         Top             =   3240
         Width           =   3015
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
         Left            =   6000
         TabIndex        =   2
         Top             =   4080
         Width           =   3015
      End
      Begin VB.CommandButton cmd_update 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1320
         TabIndex        =   1
         Top             =   8760
         Width           =   5055
      End
      Begin VB.Label points 
         Caption         =   "POINTS :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   16
         Top             =   6840
         Width           =   1575
      End
      Begin VB.Label lb_points 
         Caption         =   "50 XP"
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
         Left            =   6000
         TabIndex        =   15
         Top             =   6840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "STUDENT NAME :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "REGISTER NUMBER :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   2280
         Width           =   3855
      End
      Begin VB.Label Label3 
         Caption         =   "STUDENT EMAIL :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   3120
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "STUDENT PASSWORD :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   4080
         Width           =   4095
      End
      Begin VB.Label Label5 
         Caption         =   "SEMISTER :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "SECTION :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   6000
         Width           =   1935
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
         Top             =   1440
         Width           =   3135
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
         TabIndex        =   6
         Top             =   2400
         Width           =   3255
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
         Left            =   6000
         TabIndex        =   5
         Top             =   5040
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
         Height          =   375
         Left            =   6000
         TabIndex        =   4
         Top             =   6000
         Width           =   1935
      End
   End
End
Attribute VB_Name = "stu_prf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset


Private Sub cmd_menu_Click()
Unload stu_prf
stu_menu.Show
End Sub

Private Sub cmd_update_Click()
rs.Open "select * from profile where pname='" & stu_name & "'", con, adOpenDynamic, adLockOptimistic
rs("pemail") = txtsmail
rs("password") = txtspwd
rs("status") = "YES"
rs.Update
MsgBox "Profile Updated Successfully"
txtsmail = ""
txtspwd = ""
rs.Close
Unload stu_prf
stu_menu.Show
End Sub

Private Sub Form_Load()
lbsname = stu_name
lbsreg = stu_reg

lbssem = stu_class
lbssec = stu_sec

rs.Open "select * from profile where pname='" & stu_name & "'", con, adOpenDynamic, adLockOptimistic
txtsmail = rs("pemail")
lb_points = rs("points") & "  XP"
rs.Close

txtspwd = ppassword
txtspwd.PasswordChar = "*"
End Sub

