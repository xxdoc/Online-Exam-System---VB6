VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form stu_test 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frm_stu_test.frx":0000
   ScaleHeight     =   15690
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser mmail 
      Height          =   1935
      Left            =   19200
      TabIndex        =   17
      Top             =   12720
      Visible         =   0   'False
      Width           =   3615
      ExtentX         =   6376
      ExtentY         =   3413
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
   Begin VB.OptionButton topt3 
      BackColor       =   &H80000013&
      Caption         =   "Option3"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   15
      Top             =   8520
      Width           =   7335
   End
   Begin VB.CommandButton cmdcalc 
      BackColor       =   &H8000000D&
      Caption         =   "SUBMIT"
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
      Left            =   5640
      TabIndex        =   13
      Top             =   12480
      Width           =   3495
   End
   Begin VB.OptionButton topt4 
      BackColor       =   &H80000013&
      Caption         =   "Option4"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   12
      Top             =   10080
      Width           =   7335
   End
   Begin VB.OptionButton topt2 
      BackColor       =   &H80000013&
      Caption         =   "Option2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   11
      Top             =   6960
      Width           =   7335
   End
   Begin VB.OptionButton topt1 
      BackColor       =   &H80000013&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   10
      Top             =   5520
      Width           =   7335
   End
   Begin VB.CommandButton cmdtnext 
      BackColor       =   &H8000000D&
      Caption         =   "NEXT"
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
      Left            =   9960
      TabIndex        =   9
      Top             =   12480
      Width           =   3495
   End
   Begin VB.CommandButton submit 
      BackColor       =   &H80000001&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   18240
      TabIndex        =   7
      Top             =   11160
      Width           =   5055
   End
   Begin VB.CommandButton cmdmenu 
      BackColor       =   &H80000003&
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   18240
      MaskColor       =   &H00C00000&
      TabIndex        =   6
      Top             =   9840
      Width           =   5055
   End
   Begin VB.TextBox txtfeed 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   17400
      TabIndex        =   4
      Top             =   5160
      Width           =   6375
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   23400
      Top             =   4440
   End
   Begin VB.Label fname 
      BackColor       =   &H80000013&
      Height          =   615
      Left            =   1320
      TabIndex        =   16
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label lbsques 
      BackColor       =   &H80000013&
      Caption         =   "QUESTION 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   1320
      TabIndex        =   14
      Top             =   2160
      Width           =   13095
   End
   Begin VB.Label lbsname 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "TEST TITLE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11400
      TabIndex        =   8
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "FEED BACK :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17520
      TabIndex        =   5
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label lbsec 
      BackColor       =   &H80000013&
      Caption         =   "00"
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
      Left            =   22080
      TabIndex        =   3
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lbcolen 
      BackColor       =   &H80000013&
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
      Left            =   21600
      TabIndex        =   2
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lbmin 
      BackColor       =   &H80000013&
      Caption         =   "00"
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
      Left            =   20760
      TabIndex        =   1
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "TIME REMAING :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   20400
      TabIndex        =   0
      Top             =   2640
      Width           =   2655
   End
End
Attribute VB_Name = "stu_test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim tstatus As String
Dim ttitle, rtitle As String
Dim astatus As String
Dim sans, dbans As String
Dim marks, cmarks As Integer
Dim optpos, otans As Integer
Dim min, sec As Integer
Dim token As Long

Private Sub cmdcalc_Click()
' ------------------------------------ To check for blank answer  ------------------------------

If topt1.Value = False And topt2.Value = False And topt3.Value = False And topt4.Value = False Then
MsgBox "Select answer"
Else
cmdtnext.Enabled = True
calc
cmdcalc.Enabled = False
End If
End Sub

Private Sub cmdmenu_Click()
stu_menu.Show
Me.Hide
End Sub

Private Sub cmdtnext_Click()
' ------------------------------------ cmdnext is used to display next question  ------------------------------

cmdcalc.Enabled = True

topt1.Value = False
topt2.Value = False
topt3.Value = False
topt4.Value = False

' ------------------------------------ if rs.eof then alert user that answered all qusetion ------------------------------

If Not rs.BOF And Not rs.EOF Then
    rs.MoveNext
    If rs.EOF Then
        MsgBox "ALL Question are answered...!!!"
        MsgBox "Total marks is " & marks
        rs.Close
        cmdtnext.Enabled = False
        cmdcalc.Enabled = False
        submit.Enabled = True
        Exit Sub
    End If
    lbsques.Caption = rs("questions")
    topt1.Caption = rs("opt1")
    topt2.Caption = rs("opt2")
    topt3.Caption = rs("opt3")
    topt4.Caption = rs("opt4")
    otans = rs("optans")
    ans = dbans
End If

cmdtnext.Enabled = False

End Sub

Private Sub Form_Load()
' ------------------------------------ Retriving the Question from databse ------------------------------

rs.Open "select * from test where testid='" & t_id & "'", con, adOpenDynamic, adLockPessimistic

rs.MoveFirst

lbsques.Caption = rs("questions")
    topt1.Caption = rs("opt1")
    topt2.Caption = rs("opt2")
   
   topt3.Caption = rs("opt3")
    topt4.Caption = rs("opt4")
    otans = rs("optans")
    ans = dbans
    
    cmdtnext.Enabled = False
    
    cmdmenu.Enabled = False
    
    submit.Enabled = False
    
' ------------------------------------ enabling the timer control and setting sec to 60   ------------------------------

    Timer1.Enabled = True
    Timer1.Interval = 1000
    
    sec = 60
min = t_dur

fname.Caption = t_subject
    lbsname.Caption = t_name
    
    cmarks = 0
    marks = 0
    upoints = points
End Sub

Private Function calc()
If optpos = otans And optpos = otans Then
marks = cmarks + t_marks
cmarks = marks
End If
End Function

Private Sub submit_Click()
' ------------------------------------ submit the test results to results table  ------------------------------

Dim rs2 As New ADODB.Recordset

rs1.Open "select * from results", con, adOpenDynamic, adLockPessimistic
If txtfeed <> "" Then
rs1.AddNew
rs1("stuname") = stu_name
rs1("marks") = marks
rs1("tstatus") = "TAKEN"
rs1("feedback") = txtfeed
rs1("testtitle") = t_name
rs1("stumail") = stu_email
rs1("regno") = stu_reg
rs1("points") = upoints + marks
rs1("testid") = t_id
rs1.Update
rs1.Close

rs2.Open "select * from profile where pname='" & stu_name & "'", con, adOpenDynamic, adLockOptimistic
rs2("points") = upoints + marks
rs2.Update
rs2.Close



' ------------------------------------ submit the attendance to attendance table  ------------------------------


MsgBox "submitted Successfully !!"

Randomize
token = Int((1000000 * Rnd) + 1)
mmail.Navigate "http://rcanddt.com/api/phpmail/vendor/pmail.php?text=" & stu_email & "&name=" & stu_name & "&subject=Online Exam Marks - " & t_name & "&msg=Hi " & t_name & " This message Regarding performance in Exam.Subject : " & t_subject & "Test Title : " & t_name & "You Have Scored :" & marks & "marks." & token = " & token"

submit.Enabled = False
cmdmenu.Enabled = True


Else
MsgBox "Feedback Cannot Be Empty"
rs1.Close
End If
End Sub

Private Sub Timer1_Timer()
' ------------------------------------ Timing control of test ------------------------------

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

submit.Enabled = True
cmdmenu.Enabled = False

cmdtnext.Enabled = False
cmdcalc.Enabled = False

End If
End Sub
Private Sub topt1_Click()
If topt1.Value = True Then
optpos = 1
uans = optpos
End If
End Sub
Private Sub topt2_Click()
If topt2.Value = True Then
optpos = 2
uans = optpos
End If
End Sub
Private Sub topt3_Click()
If topt3.Value = True Then
optpos = 3
uans = optpos
End If
End Sub
Private Sub topt4_Click()
If topt4.Value = True Then
optpos = 4
uans = optpos
End If
End Sub
