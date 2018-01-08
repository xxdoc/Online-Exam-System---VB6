Attribute VB_Name = "Module1"
Public con As New ADODB.Connection

'------------------------------------  orole retrive and store user role from db ------------------------------------
Public orole As String

' ------------------------------------ This is used to store teacher details ------------------------------
Public fname, femail As String

' ------------------------------------ This is used to store Student details ------------------------------
Public stu_name, stu_sec, stu_class, stu_reg, mng_user As String
Public stu_email, points As String

' ------------------------------------ This is used to store User login details ------------------------------
Public pusername, ppassword, pstatus As String

' ------------------------------------ This is used to store Test details ------------------------------
Public t_name, t_class, t_sec, t_dur, t_id, t_marks, t_totalq, t_subject As String
Public t_tstatus, t_astatus As String
Public a_sec, a_sem, a_teacher As String

Public Sub Main()
' ------------------------------------ This is used to establish connection  ------------------------------
con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Documents and Settings\Administrator\My Documents\OnlineExam.accdb;Persist Security Info=False"
frm_Login.Show
End Sub
