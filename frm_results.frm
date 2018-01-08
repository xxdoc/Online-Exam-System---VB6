VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frm_aresults 
   Caption         =   "Results"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frm_results.frx":0000
   ScaleHeight     =   15690
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser rmail 
      Height          =   11895
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   6015
      ExtentX         =   10610
      ExtentY         =   20981
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
      Left            =   6600
      TabIndex        =   0
      Top             =   1800
      Width           =   14415
      Begin VB.CommandButton cmdprev 
         Caption         =   "PREV"
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
         Left            =   360
         TabIndex        =   4
         Top             =   10560
         Width           =   3015
      End
      Begin VB.CommandButton cmd_smail 
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
         Left            =   3960
         TabIndex        =   3
         Top             =   10560
         Width           =   2775
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "NEXT"
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
         Left            =   7320
         TabIndex        =   2
         Top             =   10560
         Width           =   2775
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
         Left            =   10680
         TabIndex        =   1
         Top             =   10560
         Width           =   2895
      End
      Begin VB.Label lb_ttitle 
         Caption         =   "RESULTS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         TabIndex        =   15
         Top             =   480
         Width           =   2175
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
         Left            =   3120
         TabIndex        =   14
         Top             =   2520
         Width           =   3375
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
         TabIndex        =   13
         Top             =   2520
         Width           =   3855
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
         Left            =   4920
         TabIndex        =   12
         Top             =   3120
         Width           =   1695
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
         TabIndex        =   11
         Top             =   3120
         Width           =   3495
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
         Left            =   4800
         TabIndex        =   10
         Top             =   3840
         Width           =   1815
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
         TabIndex        =   9
         Top             =   3840
         Width           =   1695
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
         Left            =   4680
         TabIndex        =   8
         Top             =   4560
         Width           =   1695
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
         TabIndex        =   7
         Top             =   4560
         Width           =   1935
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
         Left            =   4080
         TabIndex        =   6
         Top             =   5160
         Width           =   2415
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
         TabIndex        =   5
         Top             =   5160
         Width           =   7815
      End
   End
End
Attribute VB_Name = "frm_aresults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Dim lb_title As String
Dim dbstuname As String
Dim sname, regno, smarks, sfeed, sttitle, smail As String


Private Sub cmd_smail_Click()
Randomize
token = Int((1000000 * Rnd) + 1)
rmail.Navigate "http://rcanddt.com/api/phpmail/vendor/pmail.php?text=" & smail & "&name=" & sname & "&subject=Online Exam Marks - " & smarks & "&msg=Hi " & sname & " This message Regarding performance in Exam.Test Title : " & sttitle & "You Have Scored :" & smarks & "marks.feedback is " & sfeed & "token = " & token
MsgBox "MAIL SENT"
End Sub

Private Sub cmdmenu_Click()
rs.Close
admin.Show
Unload frm_aresults
End Sub
Private Sub cmdnext_Click()
If Not rs.BOF And Not rs.EOF Then
    rs.MovePrevious
    If rs.BOF Then
        rs.MoveLast
        results_dis
        MsgBox "You Checked All Results"
        cmdnext.Enabled = False
    Else
    results_dis
    End If
End If
End Sub

Private Sub cmdprev_Click()
cmdnext.Enabled = True
If Not rs.BOF And Not rs.EOF Then
    rs.MoveNext
    If rs.EOF Then
        rs.MoveFirst
    End If
    results_dis
End If

End Sub

Private Sub Form_Load()
rs.Open "select * from results", con, adOpenDynamic, adLockPessimistic
rs.MoveFirst
results_dis
End Sub

Private Function results_dis()
lbmstuname.Caption = rs("stuname")
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
End Function

