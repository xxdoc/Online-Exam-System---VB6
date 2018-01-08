VERSION 5.00
Begin VB.Form frm_manage 
   Caption         =   "Manage users"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frm_manage.frx":0000
   ScaleHeight     =   15690
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_search 
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
      Height          =   975
      Left            =   9000
      TabIndex        =   13
      Top             =   6240
      Width           =   5175
   End
   Begin VB.CommandButton cmd_smenu 
      Caption         =   "MAIN MENU"
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
      Left            =   9000
      TabIndex        =   11
      Top             =   11640
      Width           =   5175
   End
   Begin VB.CommandButton cmd_sdel 
      Caption         =   "DELETE STUDENT"
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
      Left            =   9000
      TabIndex        =   10
      Top             =   10320
      Width           =   5175
   End
   Begin VB.CommandButton cmd_supd 
      Caption         =   "UPDATE STUDENT"
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
      Left            =   9000
      TabIndex        =   9
      Top             =   9000
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "MANAGE USERS :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12255
      Left            =   6720
      TabIndex        =   0
      Top             =   1080
      Width           =   10455
      Begin VB.ComboBox cmb_sec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3840
         TabIndex        =   12
         Text            =   "Combo2"
         Top             =   3840
         Width           =   4575
      End
      Begin VB.CommandButton cmd_sadd 
         Caption         =   "ADD STUDENT"
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
         Left            =   2280
         TabIndex        =   8
         Top             =   6600
         Width           =   5175
      End
      Begin VB.ComboBox cmb_sem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3840
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   2880
         Width           =   4575
      End
      Begin VB.TextBox txtsreg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   6
         Top             =   2040
         Width           =   4575
      End
      Begin VB.TextBox txtsname 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   5
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
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
         Left            =   1800
         TabIndex        =   4
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         Left            =   1560
         TabIndex        =   3
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "REGNO :"
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
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME :"
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
         Left            =   2280
         TabIndex        =   1
         Top             =   1320
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frm_manage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim reg As String


Private Sub cmd_sadd_Click()
rs.Open "select * from profile", con, adOpenDynamic, adLockOptimistic
rs.AddNew
rs("username") = txtsreg
rs("password") = txtsname
rs("role") = "student"
rs("pname") = txtsname
rs("preg") = txtsreg
rs("pclass") = cmb_sem.Text
rs("psec") = cmb_sec.Text
rs("status") = "NO"
rs("points") = 0
rs("pemail") = "Enter Your Mail"
rs.Update
MsgBox "Student Added Succssfull"
txtsname = ""
txtsreg = ""
cmb_sem.Text = "select Semister"
cmb_sec.Text = "Select Section"
rs.Close
End Sub

Private Sub cmd_sdel_Click()
rs.Open "select * from profile where pname ='" & txtsname & "'", con, adOpenDynamic, adLockOptimistic
rs.Delete
rs.Update
MsgBox "Deleted Successfully"
rs.Close
End Sub

Private Sub cmd_search_Click()
rs.Open "select * from profile where pname ='" & txtsname & "'", con, adOpenDynamic, adLockOptimistic

If rs.EOF Then
MsgBox "student Not Found"
rs.Close
Exit Sub
End If

txtsname = rs("pname")
txtsreg = rs("preg")
cmb_sem.Text = rs("pclass")
cmb_sec.Text = rs("psec")
rs.Close
End Sub

Private Sub cmd_smenu_Click()
Unload frm_manage
admin.Show
End Sub

Private Sub cmd_supd_Click()
rs.Open "select * from profile where pname ='" & txtsname & "'", con, adOpenDynamic, adLockOptimistic

If rs.EOF Then
MsgBox "student Not Found"
rs.Close
Exit Sub
End If

If rs.BOF Then
MsgBox "student Not Found"
rs.Close
Exit Sub
End If

rs("pname") = txtsname
rs("preg") = txtsreg
rs("pclass") = cmb_sem.Text
rs("psec") = cmb_sec.Text
rs.Update
MsgBox "Student Updated Succssfull"
txtsname = ""
txtsreg = ""
cmb_sem.Text = "select Semister"
cmb_sec.Text = "Select Section"
rs.Close
End Sub

Private Sub Form_Load()
cmb_sem.AddItem "I BCA"
cmb_sem.AddItem "II BCA"
cmb_sem.AddItem "III BCA"
cmb_sem.AddItem "IV BCA"
cmb_sem.AddItem "V BCA"
cmb_sem.AddItem "VI BCA"

cmb_sec.AddItem "A"
cmb_sec.AddItem "B"

cmb_sem.Text = "select Semister"
cmb_sec.Text = "Select Section"

If mng_user = "" Then
Exit Sub
Else
rs.Open "select * from profile where pname ='" & mng_user & "' or preg ='" & mng_user & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "Student Not Found"
Else
txtsname = rs("pname")
txtsreg = rs("preg")
cmb_sem.Text = rs("pclass")
cmb_sec.Text = rs("psec")
rs.Close
End If
End If
End Sub

Private Sub txtsname_LostFocus()
rs.Open "select * from profile where pname ='" & txtsname & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
rs.Close
Exit Sub
End If

If rs.BOF Then
rs.Close
Exit Sub
End If

txtsname = rs("pname")
txtsreg = rs("preg")
cmb_sem.Text = rs("pclass")
cmb_sec.Text = rs("psec")
rs.Close
End Sub



Private Sub txtsreg_LostFocus()
rs.Open "select * from profile where preg ='" & txtsreg & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
rs.Close
Exit Sub
End If

If rs.BOF Then
rs.Close
Exit Sub
End If

txtsname = rs("pname")
txtsreg = rs("preg")
cmb_sem.Text = rs("pclass")
cmb_sec.Text = rs("psec")
rs.Close

End Sub
