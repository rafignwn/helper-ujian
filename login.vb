Dim Conn As New ADODB.Connection
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
    
    username = txtUsername.Text
    password = txtPassword.Text
    
    cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT id, nama, role FROM tb_user WHERE username = ? AND password = ?"
    
    cmd.Parameters.Append cmd.CreateParameter("username", adVarChar, adParamInput, 50, username)
    cmd.Parameters.Append cmd.CreateParameter("password", adVarChar, adParamInput, 100, password)
    
    rs.Open cmd
    
    If rs.EOF Then
        MsgBox "Invalid username or password.", vbExclamation, "Login Failed"
        txtUsername.Text = ""
        txtPassword.Text = ""
        txtUsername.SetFocus
        cmd.Parameters.Refresh 'membersihkan parameter pada objek Command
        rs.Close 'menutup Recordset
    Else
        MsgBox "Selamat Datang " & rs("nama").Value & "!", vbInformation, "Login Successful"
        
        cmd.Parameters.Refresh 'membersihkan parameter pada objek Command
        ' Home.nama = rs("nama").Value
        ' Home.IdOutlet = rs("id_outlet").Value
        ' Home.role = rs("role").Value
        ' frmDataUser.id_user_login = rs("id").Value
        main.IdUser = rs("id").Value
        rs.Close 'menutup Recordset
        
        Unload Me
        
        main.Show
    End If
End Sub

Private Sub Form_Load()
    Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Persist security info=false; Data Source=D:\vb-uye\Program_Prasmanan\Database1.mdb"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    If Conn.State = adStateOpen Then
        Conn.Close
    End If
End Sub

