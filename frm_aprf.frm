VERSION 5.00
Begin VB.Form frm_aprf 
   Caption         =   "ADMIN PROFILE"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame stu_profile 
      BackColor       =   &H80000009&
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
      Height          =   6975
      Left            =   10560
      TabIndex        =   0
      Top             =   6480
      Width           =   12855
      Begin VB.CommandButton cmd_update 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2520
         TabIndex        =   4
         Top             =   4560
         Width           =   4455
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
         Height          =   555
         Left            =   7440
         TabIndex        =   3
         Top             =   3480
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
         Left            =   7440
         TabIndex        =   2
         Top             =   2760
         Width           =   3015
      End
      Begin VB.CommandButton cmd_menu 
         Caption         =   "MENU"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7560
         TabIndex        =   1
         Top             =   4560
         Width           =   4335
      End
      Begin VB.Label lbsname 
         BackStyle       =   0  'Transparent
         Caption         =   "MADHU"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         TabIndex        =   8
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "TEACHER PASSWORD :"
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
         Left            =   3000
         TabIndex        =   7
         Top             =   3600
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "TEACHER EMAIL :"
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
         Left            =   3840
         TabIndex        =   6
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TEACHER NAME :"
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
         TabIndex        =   5
         Top             =   2160
         Width           =   3375
      End
   End
   Begin VB.Image Image1 
      Height          =   15450
      Left            =   0
      Picture         =   "frm_aprf.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   24000
   End
End
Attribute VB_Name = "frm_aprf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset


Private Sub cmd_menu_Click()
Unload frm_aprf
admin.Show
End Sub

Private Sub cmd_update_Click()
rs.Open "select * from profile where pname='" & fname & "'", con, adOpenDynamic, adLockOptimistic
rs("pemail") = txtsmail
rs("password") = txtspwd
rs("status") = "YES"
rs.Update
MsgBox "Profile Updated Successfully"
txtsmail = ""
txtspwd = ""
rs.Close
Unload frm_aprf
admin.Show
End Sub

Private Sub Form_Load()
lbsname = fname
txtsmail = femail
txtspwd = ppassword
txtspwd.PasswordChar = "*"
End Sub


