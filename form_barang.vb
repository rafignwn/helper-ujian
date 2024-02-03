Dim IdBarang As Integer
Dim Conn As New ADODB.Connection

Private Sub cmdGadulu_Click(Index As Integer)
    ClearForm
End Sub

Private Sub ClearForm()
    txtNamaBarang.Text = ""
    txtStok.Text = ""
    txtHarga.Text = ""
    txtSuplier.Text = ""
    IdBarang = 0
End Sub

Private Function adaYangKosong() As Boolean
    adaYangKosong = (txtNamaBarang.Text = "" Or txtStok.Text = "" Or txtHarga.Text = "" Or txtSuplier.Text = "")
End Function

Private Function pesanKosong()
    MsgBox "Tolong Disi Form yang Masih Kosong!", vbInformation, "Tidak Boleh"
End Function

Private Function pesanBerhasil(pesan As String)
    ' Kirim pesan berhasil
    MsgBox pesan, vbInformation, "BERHASIL"
    
    adodcBarang.Refresh
    
    ClearForm
End Function

Private Sub cmdHapus_Click(Index As Integer)
    If IdBarang = 0 Then
        MsgBox "Silahkan Pilih Data Yang Ingin Dihapus pada Tabel Data!", vbInformation, "Informasi"
    Else
        Conn.Execute "DELETE FROM tb_barang WHERE id = " & IdBarang
        
        pesanBerhasil "Data Barang " & txtNamaBarang.Text & " Berhasil Dihapus!"
    End If
End Sub

Private Sub cmdTambah_Click(Index As Integer)
    If IdBarang = 0 Then
        If adaYangKosong Then
            pesanKosong
        Else
            Conn.Execute "INSERT INTO tb_barang (nama_barang, stok, harga, suplier) VALUES ('" & txtNamaBarang.Text & "', " & Val(txtStok.Text) & ", " & Val(txtHarga.Text) & ", '" & txtSuplier.Text & "');"
            
            pesanBerhasil "Data Barang Baru Berhasil Ditambahkan!"
        End If
    Else
        MsgBox "Barang dengan ID " & IdBarang & " Sudah Ada! (Hanya Bisa Melakukan Update/Hapus Pengguna)", vbInformation, "Tidak Bisa Bro"
    End If
End Sub

Private Sub cmdUpdate_Click(Index As Integer)
    If Not IdBarang = 0 Then
        If adaYangKosong Then
            pesanKosong
        Else
            Conn.Execute "UPDATE tb_barang SET nama_barang = '" & txtNamaBarang.Text & "', stok = '" & txtStok.Text & "', harga = '" & txtHarga.Text & "', suplier = '" & txtSuplier.Text & "' WHERE id = " & IdBarang
            pesanBerhasil "Data Barang " & txtNamaBarang.Text & " Berhasil Diupdate"
        End If
    Else
        MsgBox "Silahkan Pilih Data Barang yang Akan Diupdate di Table Data!", vbInformation, "Informasi"
    End If
End Sub

Private Sub DGBarang_Click()
    IdBarang = DGBarang.Columns(0).Value
    txtNamaBarang.Text = CStr(DGBarang.Columns(1).Value)
    txtStok.Text = CStr(DGBarang.Columns(2).Value)
    txtHarga.Text = CStr(DGBarang.Columns(3).Value)
    txtSuplier.Text = "" ' CStr(DGBarang.Columns(4).Value)
End Sub

Public Sub ambilDataPengguna()
    adodcBarang.RecordSource = "SELECT * FROM tb_barang"
    adodcBarang.Refresh
    
    Set DGBarang.DataSource = adodcBarang
End Sub

Private Sub Form_Load()
    ambilDataPengguna
    
    Conn.Open adodcBarang.ConnectionString
    IdBarang = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conn.Close
End Sub

Private Sub txtCari_Change()
    If Not txtCari.Text = "" Then
        cmdCari.Caption = "BATAL"
        adodcBarang.RecordSource = "SELECT * FROM tb_barang WHERE nama_barang LIKE '%' + '" & txtCari.Text & "' + '%' OR suplier LIKE '%' + '" & txtCari.Text & "' + '%'"
        
        adodcBarang.Refresh
    Else
        cmdCari.Caption = "CARI"
        ambilDataPengguna
    End If
End Sub
