Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Public IdOutlet
Public nama
Public role

Private Sub cmdhitung_Click()
    ' menghitung jumlah yang harus di bayarkan
    txtbayar = Val(txtharga) * CDbl(txtjumlah)
    txtubay.SetFocus
    
    ' mengecek isi text input bayar, jika tidak kosong maka lakukan perhitungan uang kembali
    If Len(Trim(txtubay.Text)) <> 0 Then
        txtsusuk = Val(txtubay) - Val(txtbayar)
    End If
End Sub

Private Sub cmdreset_Click()
    Combo1.Text = "Pilih Paket"
    txtbayar = ""
    txtharga = ""
    txtubay = ""
    txtjumlah = ""
    txtsusuk = ""
End Sub

Private Sub Combo1_Click()
    Combo1.Text = Combo1.List(Combo1.ListIndex)
    txtharga = Combo1.ItemData(Combo1.ListIndex)
    nama_paket = Combo1.Text
    txtjumlah.SetFocus
End Sub

Private Sub Data_Pengguna_Click()
    frmDataUser.Show
End Sub

Private Sub Form_Load()
    ' membuka koneksi ke database
    conn.Open "Provider=MSDASQL.1;Persist security info=false; Data Source=londi"
    
    If role = "owner" Then
        Laporan.Visible = False
        Data_Pengguna.Visible = False
    End If
    
    If role = "kasir" Then
        Data_Paket.Visible = False
        Data_Pengguna.Visible = False
        Laporan.Visible = False
    Else
        Laporan.Visible = True
    End If
    
    ' mengambil data outlet dari user yang login
    Set rs = New ADODB.Recordset
    ' query untuk mengambil data outlet
    rs.Open "SELECT * FROM tb_outlet WHERE id = " & IdOutlet, conn
    ' mengganti judul outlet di halaman ini
    Label1.Caption = rs("nama").Value & username
    ' menutup recordset
    rs.Close
    Set rs = Nothing
    
    ' mengambil data paket londri outlet
    Set rs = New ADODB.Recordset
    ' query untuk mengambil data dari tabel paket, by id outlet
    rs.Open "SELECT * FROM tb_paket WHERE id_outlet = " & IdOutlet & ";", conn
        
    ' mengisi item combo box dengan data paket yang sudah diambil
    Set Combo1.DataSource = rs
    While Not rs.EOF
        Combo1.AddItem rs("nama_paket").Value
        Combo1.ItemData(Combo1.NewIndex) = rs("harga").Value
        rs.MoveNext
    Wend
    
    rs.Close
End Sub

Private Sub Keluar_Click()
    Unload Me
    frmLogin.Show
End Sub

Private Sub txtubay_Change()
txtsusuk = Val(txtubay) - Val(txtbayar)
End Sub
