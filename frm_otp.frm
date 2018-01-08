VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frm_otp 
   Caption         =   "OTP - LOGIN"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frm_otp.frx":0000
   ScaleHeight     =   15690
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6600
      Top             =   5760
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6120
      TabIndex        =   4
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton cmd_resend 
      Caption         =   "RESEND"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   3
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton cmd_login 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   6720
      Width           =   2655
   End
   Begin VB.TextBox txtotp 
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
      Left            =   4560
      TabIndex        =   1
      Top             =   4440
      Width           =   4935
   End
   Begin SHDocVwCtl.WebBrowser wotp 
      Height          =   150
      Left            =   19560
      TabIndex        =   0
      Top             =   4425
      Visible         =   0   'False
      Width           =   30
      ExtentX         =   53
      ExtentY         =   265
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label lbmin 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   9
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label lbcolen 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   5640
      Width           =   135
   End
   Begin VB.Label lbsec 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   7
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "REMAINING TIME :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   5640
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter The OTP :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   4560
      Width           =   2895
   End
End
Attribute VB_Name = "frm_otp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sec, min, i As Integer
Dim otp, token As Long
Dim rs As New ADODB.Recordset

Private Sub cmd_back_Click()
Unload Me
frm_Login.Show
End Sub

Private Sub cmd_login_Click()
rs.Open "select * from otp where lname='" & pusername & "' and otp=" & Val(txtotp) & "", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "Entered Wrong OTP"
rs.Close
wotp.Navigate "http://rcanddt.com/api/tele/tele.php?text=Login Failed FOR USER :" & pusername & "-Online Exam" & token
Else
wotp.Navigate "http://rcanddt.com/api/tele/tele.php?text=Login Successfull FOR USER :" & pusername & "-Online Exam" & token
MsgBox "Login Successfull"
rs.Delete
rs.Close
admin.Show
Unload Me
End If
End Sub
Private Sub cmd_resend_Click()
Timer1.Enabled = True
cmd_resend.Enabled = False

sec = 60
min = 0

otp_generate
otp_update
otp_send

End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Timer1.Interval = 1000
sec = 60
min = 0

cmd_resend.Enabled = False

otp_generate

rs.Open "select * from otp", con, adOpenDynamic, adLockOptimistic
rs.AddNew
rs("lname") = pusername
rs("otp") = otp
rs.Update
rs.Close

otp_send

wotp.Visible = True

End Sub

Private Sub Timer1_Timer()
sec = sec - 1
lbsec.Caption = sec
lbmin.Caption = min
If sec = 0 Then
sec = 60
min = min - 1
End If

If min = -1 Then
MsgBox "sorry Times Up...!!! "
Timer1.Enabled = False
cmd_resend.Enabled = True
otp_generate
otp_update
End If
End Sub
Public Function otp_generate()
Randomize
otp = Int((1000000 * Rnd) + 1)
End Function
Public Function otp_update()
rs.Open "select * from otp where lname='" & pusername & "'", con, adOpenDynamic, adLockOptimistic
rs("lname") = pusername
rs("otp") = otp
rs.Update
rs.Close
End Function
Public Function otp_send()
Randomize
token = Int((1000000 * Rnd) + 1)
wotp.Navigate "http://rcanddt.com/api/tele/tele.php?text=OTP : " & otp & "FOR USER :" & pusername & "-Online Exam" & token
End Function
Public Function login_reponse()
End Function


