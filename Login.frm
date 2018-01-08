VERSION 5.00
Begin VB.Form frm_Login 
   Caption         =   "Online Exam - Login"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "Login.frx":0000
   ScaleHeight     =   15690
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton exit 
      Caption         =   "EXIT"
      DownPicture     =   "Login.frx":33D20
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   715
      Left            =   23640
      MousePointer    =   99  'Custom
      Picture         =   "Login.frx":44FC1
      TabIndex        =   8
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton login 
      Caption         =   "LOGIN"
      DownPicture     =   "Login.frx":56262
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   715
      Left            =   21000
      Picture         =   "Login.frx":67503
      TabIndex        =   7
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton clear 
      BackColor       =   &H8000000D&
      Caption         =   "CLEAR"
      DisabledPicture =   "Login.frx":787A4
      DownPicture     =   "Login.frx":89A45
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   715
      Left            =   18360
      Picture         =   "Login.frx":9ACE6
      TabIndex        =   6
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton cmd_pshow 
      Caption         =   "SHOW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   25560
      TabIndex        =   5
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox txtpass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      IMEMode         =   3  'DISABLE
      Left            =   19800
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3360
      Width           =   5055
   End
   Begin VB.TextBox txtuser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   19800
      TabIndex        =   3
      Top             =   2280
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "PASSSWORD :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17040
      TabIndex        =   2
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "USERNAME :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17280
      TabIndex        =   1
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "Online Exam"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21240
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   15720
      Left            =   0
      Picture         =   "Login.frx":ABF87
      Stretch         =   -1  'True
      Top             =   0
      Width           =   28620
   End
End
Attribute VB_Name = "frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim pass As String

Private Sub clear_Click()
txtuser = ""
txtpass = ""
End Sub
Private Sub cmd_pshow_GotFocus()
cmd_pshow.Caption = "HIDE"
txtpass.PasswordChar = ""
End Sub

Private Sub cmd_pshow_LostFocus()
cmd_pshow.Caption = "SHOW"
txtpass.PasswordChar = "*"
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
Label3.BackStyle = 0
Label1.BackStyle = 0
Label2.BackStyle = 0
End Sub

Private Sub login_Click()
' ------------------------------------ Login to the examination panel  ------------------------------

rs.Open "select * from profile where username='" & txtuser & "' and password='" & txtpass & "' and role", con, adOpenDynamic, adLockPessimistic
If rs.EOF Then
MsgBox "Login Failed !!"
txtuser = ""
txtpass = ""
rs.Close
Exit Sub
Else

' ------------------------------------ Retriving the profile details  ------------------------------

orole = rs("role")
pusername = rs("username")
ppassword = rs("password")
stu_name = rs("pname")
stu_class = rs("pclass")
stu_email = rs("pemail")
stu_sec = rs("psec")
stu_reg = rs("preg")
points = rs("points")
pstatus = rs("status")
fname = rs("pname")
femail = rs("pemail")
MsgBox "Login Sucessfull !!"
txtuser = ""
txtpass = ""
rs.Close
End If

' ------------------------------------ If orole = admin then admin is logged in or student panel logged in  ------------------------------

If pstatus = "NO" Then
stu_profile.Show
Unload frm_Login
ElseIf orole = "student" Then
stu_menu.Show
Unload frm_Login
Else
frm_otp.Show
Unload frm_Login
End If
End Sub

