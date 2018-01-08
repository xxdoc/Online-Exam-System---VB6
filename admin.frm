VERSION 5.00
Begin VB.Form admin 
   Caption         =   "Online Exam - Admin"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "admin.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton addtest 
      BackColor       =   &H80000013&
      Caption         =   "ADD TEST"
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
      Left            =   7920
      MaskColor       =   &H8000000E&
      TabIndex        =   8
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton profile 
      BackColor       =   &H80000013&
      Caption         =   "PROFILE"
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
      Left            =   13200
      MaskColor       =   &H8000000E&
      TabIndex        =   7
      Top             =   6840
      Width           =   3015
   End
   Begin VB.CommandButton cmd_del 
      BackColor       =   &H80000013&
      Caption         =   "DELETE TEST"
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
      Left            =   18360
      MaskColor       =   &H8000000E&
      TabIndex        =   6
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton logout 
      BackColor       =   &H80000013&
      Caption         =   "LOGOUT"
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
      Left            =   18360
      MaskColor       =   &H8000000E&
      TabIndex        =   5
      Top             =   6840
      Width           =   3255
   End
   Begin VB.CommandButton search 
      BackColor       =   &H80000013&
      Caption         =   "SEARCH"
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
      Left            =   13200
      MaskColor       =   &H8000000E&
      TabIndex        =   4
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton mng_user 
      BackColor       =   &H80000013&
      Caption         =   "MANAGE USER "
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
      Left            =   18360
      MaskColor       =   &H8000000E&
      TabIndex        =   3
      Top             =   4440
      Width           =   3135
   End
   Begin VB.CommandButton ed_test 
      BackColor       =   &H80000013&
      Caption         =   "EDIT TEST"
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
      Left            =   13200
      MaskColor       =   &H8000000E&
      TabIndex        =   2
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CommandButton results 
      BackColor       =   &H80000013&
      Caption         =   "RESULTS"
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
      Left            =   7920
      MaskColor       =   &H8000000E&
      TabIndex        =   1
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton cmd_mail 
      BackColor       =   &H80000013&
      Caption         =   "SEND MAIL"
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
      Left            =   7920
      MaskColor       =   &H8000000E&
      TabIndex        =   0
      Top             =   6840
      Width           =   3375
   End
End
Attribute VB_Name = "admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Private Sub addtest_Click()
rs.Open "select * from testcontrol where faculty='" & pusername & "'", con, adOpenDynamic, adLockBatchOptimistic
If rs.EOF Then
rs.Close
a_sec = "NOT"
a_sem = "NOT"
Else
a_sec = rs("sec")
a_sem = rs("sem")
rs.Close
End If
Unload admin
frm_addtest.Show
End Sub

Private Sub cmd_del_Click()
rs.Open "select * from testcontrol where faculty='" & pusername & "'", con, adOpenDynamic, adLockBatchOptimistic
If rs.EOF Then
rs.Close
MsgBox "No Test Found"
Exit Sub
Else
rs.Close
frm_del.Show
Unload admin
End If
End Sub

Private Sub cmd_mail_Click()
Unload Me
frm_mail.Show
End Sub

Private Sub ed_test_Click()
rs.Open "select * from testcontrol where faculty='" & pusername & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
rs.Close
MsgBox "No Test Found"
Exit Sub
Else
rs.Close
frm_aedittest.Show
Unload admin
End If
End Sub

Private Sub logout_Click()
Unload admin
frm_Login.Show
End Sub

Private Sub mng_user_Click()
Unload admin
frm_manage.Show
End Sub

Private Sub profile_Click()
Unload admin
frm_aprf.Show
End Sub

Private Sub results_Click()
rs.Open "select * from results", con, adOpenDynamic, adLockPessimistic
If rs.EOF Then
rs.Close
MsgBox "No Results Found"
Exit Sub
Else
rs.Close
frm_aresults.Show
Unload admin
End If
End Sub

Private Sub search_Click()
Unload admin
frm_asearch.Show
End Sub
