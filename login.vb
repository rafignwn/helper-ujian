Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim username As String
    Dim password As String
    
    username = txtUserName.Text
    password = txtPassword.Text
    
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT id, nama, id_outlet, role FROM tb_user WHERE username = ? AND password = ?"
    
    cmd.Parameters.Append cmd.CreateParameter("username", adVarChar, adParamInput, 50, username)
    cmd.Parameters.Append cmd.CreateParameter("password", adVarChar, adParamInput, 100, password)
    
    rs.Open cmd
    
    If rs.EOF Then
        MsgBox "Invalid username or password.", vbExclamation, "Login Failed"
        txtUserName.Text = ""
        txtPassword.Text = ""
        txtUserName.SetFocus
        cmd.Parameters.Refresh 'membersihkan parameter pada objek Command
        rs.Close 'menutup Recordset
    Else
        MsgBox "Selamat Datang " & rs("nama").Value & "!", vbInformation, "Login Successful"
        
        cmd.Parameters.Refresh 'membersihkan parameter pada objek Command
        Home.nama = rs("nama").Value
        Home.IdOutlet = rs("id_outlet").Value
        Home.role = rs("role").Value
        frmDataUser.id_user_login = rs("id").Value
        rs.Close 'menutup Recordset
        
        Unload Me
        
        Home.Show
    End If
End Sub

Private Sub Form_Load()
    conn.Open "Provider=MSDASQL.1;Persist security info=false; Data Source=londi"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    If conn.State = adStateOpen Then
        conn.Close
    End If
End Sub

