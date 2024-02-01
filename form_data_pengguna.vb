Dim IdUser As Integer
Dim Conn As New ADODB.Connection

Private Sub cmdGadulu_Click()
    ClearForm
End Sub

Private Sub ClearForm()
    txtNama.Text = ""
    txtUsername.Text = ""
    txtEmail.Text = ""
    txtPassword = ""
    cmbRole.Text = "Pilih Role"
    IdUser = 0
End Sub

Private Function adaYangKosong() As Boolean
    adaYangKosong = (txtNama.Text = "" Or txtUsername.Text = "" Or txtEmail.Text = "" Or cmbRole.Text = "Pilih Role")
End Function

Private Function passwordKosong() As Boolean
    passwordKosong = (txtPassword.Text = "")
End Function

Private Function pesanKosong()
    MsgBox "Tolong Disi Form yang Masih Kosong!", vbInformation, "Tidak Boleh"
End Function

Private Function pesanBerhasil(pesan As String)
    ' Kirim pesan berhasil
    MsgBox pesan, vbInformation, "BERHASIL"
    
    adodcPengguna.Refresh
    
    ClearForm
End Function

Private Sub cmdHapus_Click()
    If IdUser = 0 Then
        MsgBox "Silahkan Pilih Data Yang Ingin Dihapus pada Tabel Data!", vbInformation, "Informasi"
    Else
        Conn.Execute "DELETE FROM tb_user WHERE id = " & IdUser
        
        pesanBerhasil "Data Pengguna " & txtNama.Text & " Berhasil Dihapus!"
    End If
End Sub

Private Sub cmdTambah_Click()
    If IdUser = 0 Then
        If adaYangKosong And passwordKosong Then
            pesanKosong
        Else
            Conn.Execute "INSERT INTO tb_user (nama, username, [password], email, role) VALUES ('" & txtNama.Text & "', '" & txtUsername.Text & "', '" & txtPassword.Text & "', '" & txtEmail.Text & "', '" & cmbRole.Text & "');"
            
            pesanBerhasil "Data Pengguna Baru Berhasil Ditambahkan!"
        End If
    Else
        MsgBox "Pengguna dengan ID " & IdUser & " Sudah Ada! (Hanya Bisa Melakukan Update/Hapus Pengguna)", vbInformation, "Tidak Bisa Bro"
    End If
End Sub

Private Sub cmdUpdate_Click()
    If Not IdUser = 0 Then
        If adaYangKosong Then
            pesanKosong
        ElseIf passwordKosong Then
            Conn.Execute "UPDATE tb_user SET nama = '" & txtNama.Text & "', username = '" & txtUsername.Text & "', email = '" & txtEmail.Text & "', role = '" & cmbRole.Text & "' WHERE id = " & IdUser
            pesanBerhasil "Data Pengguna " & txtNama.Text & " Berhasil Diupdate"
        Else
            Conn.Execute "UPDATE tb_user SET nama = '" & txtNama.Text & "', username = '" & txtUsername.Text & "', email = '" & txtEmail.Text & "', role = '" & cmbRole.Text & "', password = '" & txtPassword.Text & "' WHERE id = " & IdUser
            pesanBerhasil "Data Pengguna " & txtNama.Text & " Berhasil Diupdate"
        End If
    Else
        MsgBox "Silahkan Pilih Data Pengguna yang Akan Diupdate di Table Data!", vbInformation, "Informasi"
    End If
End Sub

Private Sub DGPengguna_Click()
    IdUser = DGPengguna.Columns(0).Value
    txtNama.Text = CStr(DGPengguna.Columns(1).Value)
    txtUsername.Text = CStr(DGPengguna.Columns(2).Value)
    txtEmail.Text = CStr(DGPengguna.Columns(3).Value)
    cmbRole.Text = CStr(DGPengguna.Columns(4).Value)
End Sub

Public Sub ambilDataPengguna()
    adodcPengguna.RecordSource = "SELECT * FROM tb_user"
    adodcPengguna.Refresh
    
    Set DGPengguna.DataSource = adodcPengguna
End Sub

Private Sub Form_Load()
    ambilDataPengguna
    cmbRole.AddItem "Admin", 0
    cmbRole.AddItem "Kasir", 1
    cmbRole.AddItem "Owner", 2
    
    Conn.Open adodcPengguna.ConnectionString
    IdUser = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conn.Close
End Sub

Private Sub txtCari_Change()
    If Not txtCari.Text = "" Then
        cmdCari.Caption = "BATAL"
        adodcPengguna.RecordSource = "SELECT * FROM tb_user WHERE nama LIKE '%' + '" & txtCari.Text & "' + '%' OR username LIKE '%' + '" & txtCari.Text & "' + '%' OR email LIKE '%' + '" & txtCari.Text & "' + '%'"
        
        adodcPengguna.Refresh
    Else
        cmdCari.Caption = "CARI"
        ambilDataPengguna
    End If
End Sub
