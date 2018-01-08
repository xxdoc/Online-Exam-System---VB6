VERSION 5.00
Begin VB.Form frm_del 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frm_del.frx":0000
   ScaleHeight     =   15690
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame f_del 
      BackColor       =   &H8000000A&
      Caption         =   "Select Test :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   7440
      TabIndex        =   0
      Top             =   3480
      Width           =   11175
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
         Height          =   735
         Left            =   7080
         TabIndex        =   13
         Top             =   5640
         Width           =   2415
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
         Height          =   735
         Left            =   3720
         TabIndex        =   12
         Top             =   5640
         Width           =   3015
      End
      Begin VB.CommandButton cmd_del 
         Caption         =   "DELETE"
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
         Left            =   1080
         TabIndex        =   11
         Top             =   5640
         Width           =   2415
      End
      Begin VB.Label lbsec 
         BackStyle       =   0  'Transparent
         Caption         =   "B SEC"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "SECTION :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label lbsem 
         BackStyle       =   0  'Transparent
         Caption         =   "IV BCA"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "semister :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label lbtestid 
         BackStyle       =   0  'Transparent
         Caption         =   "7050"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   6
         Top             =   2760
         Width           =   4215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "test id :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lbsname 
         BackStyle       =   0  'Transparent
         Caption         =   "assembly"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   2040
         Width           =   5415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "subject :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lbtname 
         BackStyle       =   0  'Transparent
         Caption         =   "assembly"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   1320
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Test Name :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   1
         Top             =   1320
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm_del"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim tid As Integer
Private Sub cmd_back_Click()
rs.Close
Unload Me
admin.Show
End Sub
Private Sub cmd_del_Click()
Dim rs1 As New ADODB.Connection
Dim rs2 As New ADODB.Connection
Dim rs3 As New ADODB.Connection

rs1.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Documents and Settings\Administrator\My Documents\OnlineExam.accdb;Persist Security Info=False"
rs1.Execute "delete * from testcontrol where testid ='" & tid & "'"
rs1.Close

rs1.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Documents and Settings\Administrator\My Documents\OnlineExam.accdb;Persist Security Info=False"
rs1.Execute "delete * from test where testid ='" & tid & "'"
rs1.Close

rs1.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Documents and Settings\Administrator\My Documents\OnlineExam.accdb;Persist Security Info=False"
rs1.Execute "delete * from results where testid ='" & tid & "'"
rs1.Close

MsgBox "Deleted Successfully", vbInformation
rs.Close
MsgBox "Restart Application", vbCritical
End

End Sub
Private Sub cmdnext_Click()
If Not rs.BOF And Not rs.EOF Then
    rs.MovePrevious
    If rs.BOF Then
        rs.MoveLast
        disp
        MsgBox "You Checked All Test", vbInformation
        cmdnext.Enabled = False
    Else
    disp
    End If
End If
End Sub
Private Sub Form_Load()
rs.Open "select * from testcontrol where faculty='" & pusername & "'", con, adOpenDynamic, adLockBatchOptimistic
disp
End Sub

Private Function disp()
lbtestid = rs("testid")
lbsem = rs("sem")
lbsec = rs("sec")
lbsname = rs("subject")
lbtname = rs("ttitle")
tid = rs("testid")
End Function

