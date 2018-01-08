VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frm_mail 
   Caption         =   "SEND MAIL"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frm_mail.frx":0000
   ScaleHeight     =   15690
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser wmail 
      Height          =   11535
      Left            =   19080
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   8295
      ExtentX         =   14631
      ExtentY         =   20346
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
      Caption         =   "SEND MAIL :"
      ForeColor       =   &H8000000D&
      Height          =   7455
      Left            =   4800
      TabIndex        =   0
      Top             =   3600
      Width           =   13935
      Begin VB.CommandButton cmdmenu 
         Caption         =   "Menu"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10920
         TabIndex        =   11
         Top             =   5760
         Width           =   1935
      End
      Begin VB.CommandButton cmdsend 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10920
         TabIndex        =   10
         Top             =   4320
         Width           =   1935
      End
      Begin VB.TextBox txtmsg 
         Height          =   2535
         Left            =   3120
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   4320
         Width           =   6975
      End
      Begin VB.TextBox txtsub 
         Height          =   495
         Left            =   3120
         TabIndex        =   7
         Top             =   3240
         Width           =   6975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11040
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtreg 
         Height          =   495
         Left            =   3120
         TabIndex        =   2
         Top             =   1320
         Width           =   6975
      End
      Begin VB.Label Label5 
         Caption         =   "Message :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   8
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Subject :"
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
         Left            =   1320
         TabIndex        =   6
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lbmail 
         Caption         =   "MADHUMANKATHA@LIVE.COM"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   5
         Top             =   2280
         Width           =   7215
      End
      Begin VB.Label Label2 
         Caption         =   "Mail - ID :"
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
         Left            =   1320
         TabIndex        =   4
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Reg No :"
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
         Left            =   1320
         TabIndex        =   1
         Top             =   1440
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_mail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim mail, mname As String
Dim token As Long

Private Sub cmdmenu_Click()
Unload Me
admin.Show
End Sub

Private Sub cmdsend_Click()
Randomize
token = Int((1000000 * Rnd) + 1)
MsgBox "Mail is" & txtreg & txtsub & txtmsg
wmail.Navigate "http://rcanddt.com/api/phpmail/vendor/pmail.php?text=" & mail & "&name=" & mname & "&subject=" & txtsub & "&msg=" & txtmsg & "&token=" & token
End Sub

Private Sub Form_Load()
cmdsend.Enabled = False
End Sub

Private Sub txtmsg_Change()
cmdsend.Enabled = True
End Sub

Private Sub txtreg_Change()
rs.Open "select * from profile where preg='" & txtreg & "' or pname='" & txtreg & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
rs.Close
Exit Sub
Else
mname = rs("pname")
mail = rs("pemail")
lbmail = mail
rs.Close
End If
End Sub

