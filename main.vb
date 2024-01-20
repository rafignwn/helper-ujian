Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim countShowItem As Integer
Dim countItem As Integer

Dim itemTop As Integer
Dim harga As Double
Dim totalHarga As Double

Dim leftLblNama As Integer
Dim leftLblQty As Integer
Dim leftLblHarga As Integer
Dim leftLblTotalHarga As Integer

Dim lblTextTotal As Label
Dim lblTotal As Label
Dim lblTextTunai As Label
Dim lblTunai As Label
Dim lblKembalian As Label
Dim lblTextKembalian As Label

Dim koleksiLabel As Collection

Private Sub tambahDaftarBarang(capNama As String, capQty As String, capHarga As Double)
    If countShowItem < countItem Then
        With koleksiLabel("lblNama" & countShowItem)
            .Visible = True
            .Caption = capNama
        End With
        
        With koleksiLabel("lblQty" & countShowItem)
            .Visible = True
            .Caption = capQty
        End With
        
        With koleksiLabel("lblHarga" & countShowItem)
            .Visible = True
            .Caption = FormatCurrency(capHarga)
        End With
        
        With koleksiLabel("lblTotalHarga" & countShowItem)
            .Visible = True
            .Caption = FormatCurrency(capHarga * CDbl(capQty))
        End With
        
        totTop = koleksiLabel("lblNama" & countShowItem).Top + koleksiLabel("lblNama" & countShowItem).Height + 400
        lineBelanja1.Y1 = totTop
        lineBelanja1.Y2 = totTop
        lblTextTotal.Visible = True
        lblTextTotal.Top = totTop + 200
        
        lblTotal.Caption = FormatCurrency(totalHarga)
        lblTotal.Visible = True
        lblTotal.Top = totTop + 200
        
        countShowItem = countShowItem + 1
    Else
        Dim lblNama As Label
        Set lblNama = Controls.Add("VB.Label", "lblNama" & countItem)
        
        Dim lblQty As Label
        Set lblQty = Controls.Add("VB.Label", "lblQty" & countItem)
        
        Dim lblHarga As Label
        Set lblHarga = Controls.Add("VB.Label", "lblHarga" & countItem)
        
        Dim lblTotalHarga As Label
        Set lblTotalHarga = Controls.Add("VB.Label", "lblTotalHarga" & countItem)
        
        With lblNama
            .Caption = capNama
            .Top = itemTop
            .Left = leftLblNama
            .FontName = "OCR-A BT"
            .FontSize = 10
            .FontBold = False
            .Visible = True
            .Height = 220
            .Width = 2500
        End With
        
        With lblQty
            .Caption = capQty
            .Top = itemTop
            .Left = leftLblQty
            .FontName = "OCR-A BT"
            .FontBold = False
            .FontSize = 10
            .Visible = True
            .Height = 220
            .Width = 375
        End With
        
        With lblHarga
            .Caption = FormatCurrency(capHarga)
            .Top = itemTop
            .Left = leftLblHarga
            .FontName = "OCR-A BT"
            .FontBold = False
            .FontSize = 10
            .Visible = True
            .Height = 220
            .Width = 2000
        End With
        
        With lblTotalHarga
            .Caption = FormatCurrency(capHarga * CDbl(capQty))
            .Top = itemTop
            .Left = leftLblTotalHarga
            .FontName = "OCR-A BT"
            .FontBold = False
            .FontSize = 10
            .Visible = True
            .Height = 220
            .Width = 2200
        End With
        
        itemTotTop = itemTop + 900
        itemLineTop = itemTop + 700
        If countItem = 0 Then
            Set lblTextTotal = Controls.Add("VB.Label", "lblTextTotal")
            Set lblTotal = Controls.Add("VB.Label", "lblTotal")
            
            With lblTextTotal
                .Caption = "TOTAL BELANJA"
                .Top = itemTotTop
                .Left = leftLblNama
                .Height = 300
                .Width = 2500
                .Visible = True
                .FontBold = True
                .FontName = "OCR-A BT"
                .FontSize = 10
            End With
            
            With lblTotal
                .Caption = FormatCurrency(totalHarga)
                .Top = itemTotTop
                .Left = leftLblTotalHarga
                .Height = 300
                .Width = 2200
                .Visible = True
                .FontBold = True
                .FontName = "OCR-A BT"
                .FontSize = 10
            End With
        Else
            lblTextTotal.Top = itemTotTop
            lblTotal.Caption = FormatCurrency(totalHarga)
            lblTotal.Top = itemTotTop
        End If
        
        lineBelanja1.Y1 = itemLineTop
        lineBelanja1.Y2 = itemLineTop
        
        koleksiLabel.Add lblNama, "lblNama" & countItem
        koleksiLabel.Add lblQty, "lblQty" & countItem
        koleksiLabel.Add lblHarga, "lblHarga" & countItem
        koleksiLabel.Add lblTotalHarga, "lblTotalHarga" & countItem
        
        itemTop = itemTop + lblNama.Height + 100
        countItem = countItem + 1
        countShowItem = countShowItem + 1
    End If
    lineBelanja1.Visible = True
End Sub

Private Sub cmdHitung_Click()
    If lblTunai Is Nothing Then
        Set lblTextTunai = Controls.Add("VB.Label", "lblTextTunai")
        Set lblTunai = Controls.Add("VB.Label", "lblTunai")
        
        Set lblTextKembalian = Controls.Add("VB.Label", "lblTextKembalian")
        Set lblKembalian = Controls.Add("VB.Label", "lblKembalian")
        
        With lblTextTunai
            .Caption = "TUNAI"
            .FontBold = True
            .FontName = "OCR-A BT"
            .FontSize = 10
            .Left = 6240
            .Width = 3135
            .Height = 300
            .Visible = True
            .Top = lblTextTotal.Top + 400
        End With
        
        With lblTunai
            .Caption = FormatCurrency(txtTunai.Text)
            .FontBold = True
            .FontName = "OCR-A BT"
            .FontSize = 10
            .Left = leftLblTotalHarga
            .Width = 2200
            .Height = 300
            .Visible = True
            .Top = lblTextTotal.Top + 400
        End With
        
        With lblTextKembalian
            .Caption = "KEMBALIAN"
            .FontBold = True
            .FontName = "OCR-A BT"
            .FontSize = 10
            .Left = 6240
            .Width = 3135
            .Height = 300
            .Visible = True
            .Top = lblTextTunai.Top + 400
        End With
        
        kembalian = CDbl(txtTunai.Text) - totalHarga
        With lblKembalian
            .Caption = FormatCurrency(kembalian)
            .FontBold = True
            .FontName = "OCR-A BT"
            .FontSize = 10
            .Left = leftLblTotalHarga
            .Width = 2200
            .Height = 300
            .Visible = True
            .Top = lblTextTunai.Top + 400
        End With
    Else
        With lblTextTunai
            .Visible = True
            .Top = lblTextTotal.Top + 400
        End With
        
        With lblTunai
            .Caption = FormatCurrency(txtTunai.Text)
            .Visible = True
            .Top = lblTextTotal.Top + 400
        End With
        
        With lblTextKembalian
            .Visible = True
            .Top = lblTextTunai.Top + 400
        End With
        
        kembalian = CDbl(txtTunai.Text) - totalHarga
        With lblKembalian
            .Caption = FormatCurrency(kembalian)
            .Visible = True
            .Top = lblTextTunai.Top + 400
        End With
    End If
    txtTunai.Text = ""
End Sub

Private Sub cmdReset_Click()
    For i = countItem - 1 To 0 Step -1
        koleksiLabel("lblNama" & i).Visible = False
        koleksiLabel("lblQty" & i).Visible = False
        koleksiLabel("lblHarga" & i).Visible = False
        koleksiLabel("lblTotalHarga" & i).Visible = False
    Next i
    lblTextTotal.Visible = False
    lblTotal.Visible = False
    countShowItem = 0
    totalHarga = 0
    lineBelanja1.Visible = False
    txtTunai.Text = ""
    lblTextTunai.Visible = False
    lblTextKembalian.Visible = False
    lblTunai.Visible = False
    lblKembalian.Visible = False
    clear
End Sub

Private Sub clear()
    txtHarga = "Rp 0"
    comboBarang.Text = "Pilih Barang"
    lblBarang = ""
    txtQty.Text = ""
End Sub

Private Sub cmdTambah_Click()
    totalHarga = totalHarga + (harga * CDbl(txtQty.Text))
    tambahDaftarBarang comboBarang.Text, txtQty.Text, CDbl(harga)
    clear
    txtTunai.SetFocus
End Sub

Private Sub comboBarang_Click()
    txtQty.Text = ""
    comboBarang.Text = comboBarang.List(comboBarang.ListIndex)
    harga = CDbl(comboBarang.ItemData(comboBarang.ListIndex))
    txtHarga = FormatCurrency(harga)
    lblBarang = comboBarang.Text
    txtQty.SetFocus
End Sub

Private Sub Form_Load()
    ' membuka koneksi ke database
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Persist security info=false; Data Source=D:\vb-uye\Prasmanan\Database1.mdb"
    
    ' mengambil data paket londri outlet
    Set rs = New ADODB.Recordset
    ' query untuk mengambil data dari tabel paket, by id outlet
    rs.Open "SELECT * FROM tb_barang", conn
        
    ' mengisi item combo box dengan data paket yang sudah diambil
    Set comboBarang.DataSource = rs
    While Not rs.EOF
        comboBarang.AddItem rs("nama_barang").Value
        comboBarang.ItemData(comboBarang.NewIndex) = rs("harga").Value
        rs.MoveNext
    Wend
    
    rs.Close
    
    countItem = 0
    countShowItem = 0
    itemTop = 2590
    totalHarga = 0
    Set koleksiLabel = New Collection
    
    leftLblNama = 6240
    leftLblQty = 9000
    leftLblHarga = 9840
    leftLblTotalHarga = 12120
End Sub
