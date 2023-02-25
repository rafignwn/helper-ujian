Dim con As New ADODB.Connection
Private id_outlet
Public id_user_login

Private Sub Clear_Form()
    txtNama.Text = ""
    txtId.Text = ""
    txtPw.Text = ""
    txtUsername.Text = ""
    Combo1.Text = "Pilih Outlet"
    txtRole.Text = ""
End Sub

Private Sub Show_Data()
    Adodc1.ConnectionString = con
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "SELECT tb_user.id AS id_user, tb_user.nama AS nama_pengguna, username, role, tb_outlet.id AS id_outlet, tb_outlet.nama AS nama_outlet FROM tb_user INNER JOIN tb_outlet ON tb_user.id_outlet = tb_outlet.id"
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Columns(0).Caption = "ID"
    DataGrid1.Columns(1).Caption = "Nama"
    DataGrid1.Columns(2).Caption = "Username"
    DataGrid1.Columns(3).Caption = "Role"
    DataGrid1.Columns(4).Caption = "ID Outlet"
    DataGrid1.Columns(5).Caption = "Nama Outlet"
    
End Sub

Private Sub Tambah_Data()
    con.Execute "INSERT INTO tb_user (nama, username, password, role, id_outlet) VALUES ('" & txtNama & "', '" & txtUsername & "', '" & txtPw & "', '" & txtRole & "', " & id_outlet & ")"
    
    ' refresh data di dalam data grid
    Adodc1.Refresh
    
    MsgBox "Pengguna Baru Berhasil Ditambahkan", vbInformation, "Info Bahagia"
    
    Call Clear_Form
End Sub

Private Sub Update_Data()
    If Len(Trim(txtId.Text)) <> 0 Then
        If Len(Trim(txtPw.Text)) <> 0 Then
            con.Execute "UPDATE tb_user SET nama = '" & txtNama & "', username = '" & txtUsername & "', password = '" & txtPw & "', role = '" & txtRole & "', id_outlet = " & id_outlet & " WHERE id = " & txtId & ";"
        Else
            con.Execute "UPDATE tb_user SET nama = '" & txtNama & "', username = '" & txtUsername & "', role = '" & txtRole & "', id_outlet = " & id_outlet & " WHERE id = " & txtId & ";"
        End If
        
        Adodc1.Refresh
        
        MsgBox "Data Berhasil Di Update", vbInformation, "Info"
        
        Call Clear_Form
    Else
        MsgBox "ID Tidak Boleh Kosong", vbExclamation, "Peringatan"
    End If
End Sub

Private Sub Hapus_Data()
    If txtId = id_user_login Then
        MsgBox "Anda Tidak bisa menghapus diri sendiri", vbInformation, "Info Penting"
    Else
        If Len(Trim(txtId.Text)) <> 0 Then
            con.Execute "DELETE FROM tb_user WHERE id = " & txtId & ";"
            
            Adodc1.Refresh
            
            MsgBox "Data Berhasil Dihapus", vbInformation, "Info"
            Call Clear_Form
        Else
            MsgBox "ID Tidak Boleh Kosong", vbExclamation, "Peringatan"
        End If
    End If
End Sub

Private Sub Combo1_Click()
    id_outlet = Combo1.ItemData(Combo1.ListIndex)
End Sub

Private Sub DataGrid1_Click()
    txtId.Text = CStr(DataGrid1.Columns(0).Value)
    txtNama.Text = CStr(DataGrid1.Columns(1).Value)
    txtUsername.Text = CStr(DataGrid1.Columns(3).Value)
    txtRole.Text = CStr(DataGrid1.Columns(3).Value)
    id_outlet = CInt(DataGrid1.Columns(4).Value)
    Combo1.Text = CStr(DataGrid1.Columns(5).Value)
End Sub

Private Sub Form_Load()
    con.Open "Provider=MSDASQL.1;Persist security info=false; Data Source=londi"
    
    Call Show_Data
    
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT * FROM tb_outlet", con
    
    While Not rs.EOF
        Combo1.AddItem rs("nama").Value
        Combo1.ItemData(Combo1.NewIndex) = rs("id").Value
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub hapusBtn_Click()
    Call Hapus_Data
End Sub

Private Sub newBtn_Click()
    Call Clear_Form
End Sub

Private Sub saveBtn_Click()
    Call Tambah_Data
End Sub

Private Sub updateBtn_Click()
    Call Update_Data
End Sub
