VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frm_asearch 
   Caption         =   "Search Students"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   23130
   FillColor       =   &H00800000&
   LinkTopic       =   "Form1"
   Picture         =   "frm_asearch.frx":0000
   ScaleHeight     =   10980
   ScaleWidth      =   23130
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser rmail 
      Height          =   4695
      Left            =   240
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   3135
      ExtentX         =   5530
      ExtentY         =   8281
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
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
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
      ForeColor       =   &H8000000D&
      Height          =   12015
      Left            =   4320
      TabIndex        =   0
      Top             =   840
      Width           =   16455
      Begin VB.CommandButton cmd_clear 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9120
         TabIndex        =   20
         Top             =   10560
         Width           =   2655
      End
      Begin VB.CommandButton cmd_mail 
         Caption         =   "SEND MAIL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5520
         TabIndex        =   19
         Top             =   10560
         Width           =   2775
      End
      Begin VB.CommandButton cmd_search 
         BackColor       =   &H80000002&
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12240
         MaskColor       =   &H000000FF&
         TabIndex        =   16
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txt_search 
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
         Left            =   4200
         TabIndex        =   15
         Top             =   840
         Width           =   7695
      End
      Begin VB.CommandButton cmdmenu 
         Caption         =   "MENU"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   12360
         TabIndex        =   2
         Top             =   10560
         Width           =   2895
      End
      Begin VB.CommandButton cmd_smail 
         Caption         =   "MANAGE USER"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1920
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   10560
         Width           =   2775
      End
      Begin VB.Label lb_reg 
         BackStyle       =   0  'Transparent
         Caption         =   "15kxsb7043"
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
         Left            =   7080
         TabIndex        =   18
         Top             =   4200
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "REG NO :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   5280
         TabIndex        =   17
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH STUDENT :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label lbmfeedback 
         BackStyle       =   0  'Transparent
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
         Height          =   3255
         Left            =   6960
         TabIndex        =   13
         Top             =   6960
         Width           =   7815
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "FEEDBACK :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   4680
         TabIndex        =   12
         Top             =   6960
         Width           =   2175
      End
      Begin VB.Label lbmpoints 
         BackStyle       =   0  'Transparent
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
         Left            =   6960
         TabIndex        =   11
         Top             =   6360
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "POINTS :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   5280
         TabIndex        =   10
         Top             =   6360
         Width           =   1695
      End
      Begin VB.Label lbmmarks 
         BackStyle       =   0  'Transparent
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
         Left            =   6960
         TabIndex        =   9
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "MARKS :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   5400
         TabIndex        =   8
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label lbmmail 
         BackStyle       =   0  'Transparent
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
         Left            =   7080
         TabIndex        =   7
         Top             =   4920
         Width           =   3495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "E-MAIL :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   5520
         TabIndex        =   6
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label lbmstuname 
         BackStyle       =   0  'Transparent
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
         Left            =   7080
         TabIndex        =   5
         Top             =   3480
         Width           =   3855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "STUDENT NAME :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   3480
         Width           =   3135
      End
      Begin VB.Label lb_ttitle 
         BackStyle       =   0  'Transparent
         Caption         =   "RESULTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   6360
         TabIndex        =   3
         Top             =   2400
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm_asearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim search As String
Dim sname, regno, smarks, sfeed, sttitle, smail As String
Private Sub cmd_clear_Click()
lbmstuname.Caption = ""
lb_reg = ""
lbmmail = ""
lbmmarks = ""
lbmfeedback = ""
lb_ttitle.Caption = ""
lbmpoints = ""
End Sub

Private Sub cmd_mail_Click()
Randomize
token = Int((1000000 * Rnd) + 1)
rmail.Navigate "http://rcanddt.com/api/phpmail/vendor/pmail.php?text=" & smail & "&name=" & sname & "&subject=Online Exam Marks - " & smarks & "&msg=Hi " & sname & " This message Regarding performance in Exam.Test Title : " & sttitle & "You Have Scored :" & smarks & "marks.feedback is " & sfeed & "token = " & token
MsgBox "E-mail successfully Sent To " & sname, vbInformation
End Sub

Private Sub cmd_search_Click()
rs.Open "select * from results where stuname='" & txt_search & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "No Student Found"
rs.Close
Else
lbmstuname.Caption = rs("stuname")
lb_reg = rs("regno")
lbmmail = rs("stumail")
lbmmarks = rs("marks")
lbmfeedback = rs("feedback")
lb_ttitle.Caption = rs("testtitle")
sname = rs("stuname")
regno = rs("regno")
smarks = rs("marks")
sfeed = rs("feedback")
sttitle = rs("testtitle")
smail = rs("stumail")
rs.Close
End If
End Sub

Private Sub cmd_smail_Click()
If txt_search = "" Then
MsgBox "Enter Student Details"
Else
mng_user = txt_search
Unload frm_asearch
frm_manage.Show
End If
End Sub

Private Sub cmdmenu_Click()
Unload frm_asearch
admin.Show
End Sub

Private Sub Form_Load()
lbmstuname.Caption = ""
lb_reg = ""
lbmmail = ""
lbmmarks = ""
lbmfeedback = ""
lb_ttitle.Caption = ""
lbmpoints = ""
End Sub

Private Sub txt_search_Change()
search = txt_search.Text
rs.Open "select * from results where stuname='" & txt_search & "' or regno='" & search & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
rs.Close
Exit Sub
Else
lbmstuname.Caption = rs("stuname")
lb_reg = rs("regno")
lbmmail = rs("stumail")
lbmmarks = rs("marks")
lbmfeedback = rs("feedback")
lb_ttitle.Caption = rs("testtitle")
lbmpoints = rs("points")
sname = rs("stuname")
regno = rs("regno")
smarks = rs("marks")
sfeed = rs("feedback")
sttitle = rs("testtitle")
smail = rs("stumail")
rs.Close
End If

End Sub
