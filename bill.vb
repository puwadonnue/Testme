Imports System.Data.SqlClient
Public Class bill
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double

    Private Sub bill_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        auto_Number()
        auto_Number3()
        refresh_edit()
        refresh_edit1()
        refresh_edit3()
        FormatGridview()
        FormatGridview2()
        FormatGridview3()

        DataGridView2.Columns(3).Visible = False
        'DataGridView1.Columns(2).Visible = False
        sql = "select service_id from service"
        cmd_database_to_object(combo_service)
        sql = "select service_id from service"
        cmd_database_to_object(Combo_service3)
        sql = "select bill_id from bill"
        cmd_database_to_object(Combo_bill3)
        sql = "select clothestype_id from clothesType"
        cmd_database_to_object(Combo_type3)

        'sql = "select Cus_id from Customer"
        'cmd_database_to_object(ComboBox_cusid)
        refresh_sql()
        refresh_sql3()

    End Sub

    Private Sub TabPage2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage2.Click

    End Sub
    Public Sub FormatGridview()
        With DataGridView1

            .Columns(0).HeaderText = "รหัสใบเสร็จ"
            .Columns(1).HeaderText = "ชื่อลูกค้า"
            .Columns(2).HeaderText = "วันที่รับผ้า"
            .Columns(3).HeaderText = "วันที่ส่งผ้า"
            .Columns(4).HeaderText = "รหัสลูกค้า"
            .Columns(0).Width = 140
            .Columns(1).Width = 155
            .Columns(2).Width = 135
            .Columns(3).Width = 115
            .Columns(4).Width = 135
        End With
    End Sub


    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Me.Close()
    End Sub

    Private Sub refresh_edit()

        ' SELECT * FROM branch INNER JOIN member ON ( branch.branch_id = member.branch_id)
        sql = "select * from service"

        DataGridView2.DataSource = cmd_dataTable()
    End Sub

    Private Sub DataGridView2_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim i As Integer = DataGridView2.CurrentRow.Index
        combo_service.Text = DataGridView2.Item(0, i).Value
        txt_service.Text = DataGridView2.Item(1, i).Value
        price_service.Text = DataGridView2.Item(2, i).Value
        price_service2.Text = DataGridView2.Item(3, i).Value
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        sql = "update service set service_type='" & txt_service.Text & "',service_washing='" & price_service.Text & "',service_roning='" & price_service2.Text & "' where service_id='" & combo_service.Text & "'"
        If cmd_excuteNonquery() = 0 Then
            MessageBox.Show("คลิกที่ตารางก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MsgBox("แก้ไขสำเร็จ")
        End If
        combo_service.Text = "เลือกรหัสการบริการ"
        txt_service.Text = ""
        price_service.Text = ""
        price_service2.Text = ""
        refresh_edit()
    End Sub
    Public Sub FormatGridview2()
        With DataGridView2

            .Columns(0).HeaderText = "รหัสการบริการ"
            .Columns(1).HeaderText = "ประเภทการบริการ"
            .Columns(2).HeaderText = "ราคาค่าบริการ"
            .Columns(3).HeaderText = "รหัส"
            .Columns(0).Width = 140
            .Columns(1).Width = 155
            .Columns(2).Width = 135
            .Columns(3).Width = 115
        End With
    End Sub

    Private Sub combo_service_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles combo_service.SelectedIndexChanged
        sql = "select service_type from service where service_id like '" & combo_service.SelectedItem & "'"
        txt_service.Text = cmd_excuteScalar()
        sql = "select service_washing from service where service_id like '" & combo_service.SelectedItem & "'"
        price_service.Text = cmd_excuteScalar()
        sql = "select service_roning from service where service_id like '" & combo_service.SelectedItem & "'"
        price_service2.Text = cmd_excuteScalar()
        lbl_washing.Text = "บาท/ถัง"
        lbl_roning.Text = "บาท/ตัว"


    End Sub
    '''''''''''''''''tab 1 '''''''''''''''''''''''''
    Dim sDate, eDate As Date
    Public Sub time_d()
        sDate = Format(Me.DateTimePicker2.Value, "dd-MM-yyyy")
        eDate = Format(Me.DateTimePicker1.Value, "dd-MM-yyyy")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        time_d()
        sql = "select count(*) from bill where bill_id='" & txt_id1.Text & "'"
        If cmd_excuteScalar() > 0 Then
            ComboBox_cusid.Text = "เลือกรหัสลูกค้า"
            lblCus.Text = ""

            auto_Number()
            refresh_edit1()
            refresh_sql()
            Return
        End If
        sql = "insert into bill values(@bill_id,@Cus_id,@auto_id,@bill_add,@bill_exit)"
        cmd = New SqlCommand(sql, cn)
        cmd.Parameters.Clear()
        cmd.Parameters.AddWithValue("@bill_id", txt_id1.Text)
        cmd.Parameters.AddWithValue("@Cus_id", ComboBox_cusid.Text)
        cmd.Parameters.AddWithValue("@auto_id", txt_autoid1.Text)
        cmd.Parameters.AddWithValue("@bill_add", eDate)
        cmd.Parameters.AddWithValue("@bill_exit", sDate)
        If (ComboBox_cusid.Text = "เลือกรหัสลูกค้า") Then
            MessageBox.Show("กรุณาเลือกรหัสลูกค้าก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        ElseIf cmd.ExecuteNonQuery > 0 Then
            MsgBox("เพิ่มข้อมูลสำเร็จ")
            ComboBox_cusid.Text = "เลือกรหัสลูกค้า"
            DateTimePicker1.Value = Date.Now
            DateTimePicker2.Value = Date.Now
            lblCus.Text = ""

            time_d()
            auto_Number()
            refresh_edit1()
            refresh_sql()

        Else

            ComboBox_cusid.Text = "เลือกรหัสลูกค้า"
            DateTimePicker1.Value = Date.Now
            DateTimePicker2.Value = Date.Now
        End If
    End Sub
    Private Sub auto_Number()
        sql = "select max(auto_id) from bill"
        Try
            Dim numchar_id As String = "SB-" & (cmd_excuteScalar() + 1).ToString.PadLeft(3, "0")
            txt_id1.Text = numchar_id
            txt_autoid1.Text = cmd_excuteScalar() + 1
        Catch ex As Exception
            txt_id1.Text = "SB-001"
            txt_autoid1.Text = 1
        End Try
    End Sub
    Private Sub refresh_edit1()
        sql = "SELECT bill.bill_id,Customer.Cus_name,bill.bill_add,bill.bill_exit,Customer.Cus_id FROM bill INNER JOIN Customer ON(bill.Cus_id = Customer.Cus_id)"
        ' sql = "select * from bill"
        DataGridView1.DataSource = cmd_dataTable()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        time_d()
        sql = "UPDATE bill SET Cus_id = @Cus_id,auto_id = @auto_id,bill_add = @bill_add,bill_exit = @bill_exit WHERE bill_id = @bill_id "
        cmd = New SqlCommand(sql, cn)
        cmd.Parameters.Clear()
        cmd.Parameters.AddWithValue("@Cus_id", ComboBox_cusid.Text)
        cmd.Parameters.AddWithValue("@auto_id", txt_autoid1.Text)
        cmd.Parameters.AddWithValue("@bill_add", eDate)
        cmd.Parameters.AddWithValue("@bill_exit", sDate)
        cmd.Parameters.AddWithValue("@bill_id", txt_id1.Text)
        If cmd.ExecuteNonQuery > 0 Then
            MessageBox.Show("แก้ไขข้อมูลแล้ว")
            ComboBox_cusid.Text = "เลือกรหัสลูกค้า"
            lblCus.Text = ""
            DateTimePicker1.Value = Date.Now
            DateTimePicker2.Value = Date.Now
            auto_Number()
            refresh_edit1()
            refresh_sql()
        Else
            MessageBox.Show("คลิกที่ตารางก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim i As Integer = DataGridView1.CurrentRow.Index
        txt_id1.Text = DataGridView1.Item(0, i).Value
        lblCus.Text = DataGridView1.Item(1, i).Value
        ComboBox_cusid.Text = DataGridView1.Item(4, i).Value
        DateTimePicker1.Text = DataGridView1.Item(2, i).Value
        DateTimePicker2.Text = DataGridView1.Item(3, i).Value
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        sql = "delete from bill where bill_id='" & txt_id1.Text & "'"
        If cmd_excuteNonquery() = 0 Then
            MessageBox.Show("คลิกที่ตารางก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MsgBox("ลบข้อมูลสำเร็จ")
        End If
        lblCus.Text = ""
        ComboBox_cusid.Text = "เลือกรหัสลูกค้า"
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        auto_Number()
        refresh_edit1()
        refresh_sql()

    End Sub
    Private Sub refresh_sql()
        sql = "SELECT * FROM Customer WHERE Customer.Cus_id not in (SELECT Cus_id FROM bill)"
        cmd_database_to_object(ComboBox_cusid)
    End Sub

    Private Sub combo_service_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles combo_service.KeyPress
        e.Handled = True
    End Sub

    Private Sub ComboBox_cusid_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox_cusid.KeyPress
        e.Handled = True
    End Sub

    Private Sub TabPage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub ComboBox_cusid_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox_cusid.SelectedIndexChanged
        sql = "select Cus_name from Customer where Cus_id like '" & ComboBox_cusid.SelectedItem & "'"
        lblCus.Text = cmd_excuteScalar()
    End Sub
    ''' '''''''''''''''''''''''''3''''''''''''''''''''''''
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        sql = "select count(*) from bill_Service where billservice_id='" & txt_id3.Text & "'"
        If cmd_excuteScalar() > 0 Then
            Combo_service3.Text = "เลือกรหัสการบริการ"
            Combo_bill3.Text = "เลือกรหัสลูกค้า"
            Combo_type3.Text = "เลือกรหัสประเภทเสื้อผ้า"
            auto_Number3()
            num_nuam.Text = ""
            num3.Text = ""
            refresh_edit3()
            refresh_sql3()
            price3.Text = ""
            lbl_cusname3.Text = ""
            lbl_typeservice3.Text = ""
            lbl_typeclothes.Text = ""
            lblwashing3.Text = ""
            lblroning3.Text = ""
            Return
        End If
        sql = "insert into bill_Service values('" & txt_id3.Text & "','" & Combo_bill3.Text & "','" & Combo_service3.Text & "','" & Combo_type3.Text & "','" & num_nuam.Text & "','" & num3.Text & "','" & price3.Text & "','" & txt_autoid3.Text & "')"
        If (Combo_type3.Text = "เลือกรหัสประเภทเสื้อผ้า") Or Combo_bill3.Text = "เลือกรหัสลูกค้า" Or Combo_service3.Text = "เลือกรหัสการบริการ" Then
            MessageBox.Show("กรุณาป้อนข้อมูลให้ครบ !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        ElseIf cmd_excuteNonquery() >= 1 Then
            MsgBox("เพิ่มข้อมูลสำเร็จ")
            txt_id3.Text = ""
            Combo_service3.Text = "เลือกรหัสการบริการ"
            Combo_bill3.Text = "เลือกรหัสลูกค้า"
            Combo_type3.Text = "เลือกรหัสประเภทเสื้อผ้า"
            num_nuam.Text = ""
            num3.Text = ""
            auto_Number3()
            refresh_edit3()
            price3.Text = ""
            refresh_sql3()
            lbl_cusname3.Text = ""
            lbl_typeservice3.Text = ""
            lbl_typeclothes.Text = ""
            lblwashing3.Text = ""
            lblroning3.Text = ""
        End If
    End Sub
  
    Private Sub Combo_service3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_service3.SelectedIndexChanged
        sql = "select service_type from service where service_id like '" & Combo_service3.SelectedItem & "'"
        lbl_typeservice3.Text = cmd_excuteScalar()
        sql = "select service_washing from service where service_id like'" & Combo_service3.SelectedItem & "'"
        lblwashing3.Text = cmd_excuteScalar()
        sql = "select service_roning from service where service_id like'" & Combo_service3.SelectedItem & "'"
        lblroning3.Text = cmd_excuteScalar()
    End Sub

    Private Sub Combo_bill3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_bill3.SelectedIndexChanged
        'sql = "SELECT Customer.Cus_name FROM bill INNER JOIN Customer ON(bill.Cus_id = Customer.Cus_id) where bill.bill_id like '" & Combo_bill3.SelectedItem & "'"
        sql = "Select Cus_name from Customer where Cus_id like'" & Combo_bill3.SelectedItem & "'"
        lbl_cusname3.Text = cmd_excuteScalar()
    End Sub

    Private Sub Combo_type3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_type3.SelectedIndexChanged
        sql = "Select clothestype from clothesType where clothestype_id like '" & Combo_type3.SelectedItem & "'"
        lbl_typeclothes.Text = cmd_excuteScalar()
        If Combo_type3.Text = "CL-002" Then
            num3.ReadOnly = True
            num_nuam.ReadOnly = False
        ElseIf Combo_type3.Text = "CL-001" Then
            num_nuam.ReadOnly = True
        Else
            num_nuam.ReadOnly = False
            num3.ReadOnly = False

        End If

    End Sub

    Private Sub num3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles num3.TextChanged
        Dim price As Double
        Dim total As Double
        Dim num As Double
        Dim f As Double
        If Not IsNumeric(num3.Text) Then
            num3.Text = ""
        ElseIf lbl_typeservice3.Text = "ซักอบ" Or lbl_typeservice3.Text = "ซักตาก" Or Combo_type3.Text = "CL-001" Or Combo_type3.Text = "CL-003" Then
            f = nub.Text
            price = lblwashing3.Text
            num = num3.Text
            total = price * num
            price3.Text = total + f
        Else
            f = nub.Text
            'price = lblroning3.Text
            num = num3.Text
            total = price * num
            price3.Text = total + f
        End If

        If Combo_type3.Text = "CL-002" Then
        End If
    End Sub

    Private Sub num_nuam_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles num_nuam.TextChanged

        Dim k As Double
        If Combo_type3.Text = "CL-002" Then
            k = num_nuam.Text
            a = 3
            b = 50
            If k >= a Then
                a = num_nuam.Text - a
                c = a * 10
                b = b + c
                price3.Text = b
            Else
                MsgBox("ขนาดของผืนมี 3 ฟุตขึ้นไป")
            End If
        ElseIf Combo_type3.Text = "CL-003" Then
            k = num_nuam.Text
            a = 3
            b = 50
            If k >= a Then
                a = num_nuam.Text - a
                c = a * 10
                b = b + c
                nub.Text = b
            End If
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim fnt As Font
        Dim txt As String
        Dim black As New SolidBrush(Color.Black)
        Dim blue As New SolidBrush(Color.Blue)
        fnt = New Font("cordiaupc", 20, FontStyle.Bold)
        txt = "ใบเสร็จค่าบริการ"
        e.Graphics.DrawString(txt, fnt, black, 300, 50)
        txt = "รหัสใบเสร็จ"
        e.Graphics.DrawString(txt, fnt, black, 200, 90)
        txt = txt_id1.Text
        e.Graphics.DrawString(txt, fnt, black, 350, 90)
        txt = "วันที่รับผ้า"
        e.Graphics.DrawString(txt, fnt, black, 200, 130)
        txt = DateTimePicker1.Text
        e.Graphics.DrawString(txt, fnt, black, 350, 130)
        txt = "วันที่ส่งผ้า"
        e.Graphics.DrawString(txt, fnt, black, 200, 170)
        txt = DateTimePicker2.Text
        e.Graphics.DrawString(txt, fnt, black, 350, 170)
        txt = "รหัสลูกค้า"
        e.Graphics.DrawString(txt, fnt, black, 200, 210)
        txt = ComboBox_cusid.Text
        e.Graphics.DrawString(txt, fnt, black, 350, 210)
        txt = "ชื่อลูกค้า"
        e.Graphics.DrawString(txt, fnt, black, 200, 250)
        txt = lblCus.Text
        e.Graphics.DrawString(txt, fnt, black, 350, 250)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.Width = 850
        PrintPreviewDialog1.Height = 600
        PrintPreviewDialog1.ShowDialog()
    End Sub
    Private Sub auto_Number3()
        sql = "select max(auto_id) from bill_Service"
        Try
            Dim numchar_id As String = "BS-" & (cmd_excuteScalar() + 1).ToString.PadLeft(3, "0")
            txt_id3.Text = numchar_id
            txt_autoid3.Text = cmd_excuteScalar() + 1
        Catch ex As Exception
            txt_id3.Text = "BS-001"
            txt_autoid3.Text = 1
        End Try
    End Sub
    Private Sub refresh_edit3()
        '  sql = "select * from bill_Service"
        'sql = "SELECT * FROM bill_Service INNER JOIN Customer ON(bill.Cus_id = Customer.Cus_id)"
        ' sql = "select * from bill"
        sql = "SELECT bill_Service.billservice_id,Customer.Cus_name,bill_Service.nuam,bill_Service.num,bill_Service.price,service.service_id,clothesType.clothestype_id,Customer.Cus_id FROM bill_Service INNER JOIN Customer ON(bill_Service.Cus_id = Customer.Cus_id) INNER JOIN service ON(bill_Service.service_id = service.service_id) INNER JOIN clothesType ON(bill_Service.clothestype_id = clothesType.clothestype_id)"
        DataGridView3.DataSource = cmd_dataTable()
    End Sub
   
    Private Sub refresh_sql3()
        'sql = "SELECT * FROM bill WHERE bill.bill_id not in (SELECT bill_id FROM bill_Service)"
        ' sql = "SELECT * FROM bill_Service WHERE bill_Service.Cus_id not in (SELECT Cus_id FROM Customer)"
        sql = "SELECT * FROM Customer WHERE Customer.Cus_id not in (SELECT Cus_id FROM bill_Service)"
        cmd_database_to_object(Combo_bill3)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        sql = "update bill_Service set Cus_id='" & Combo_bill3.Text & "',service_id='" & Combo_service3.Text & "',clothestype_id='" & Combo_type3.Text & "',nuam='" & num_nuam.Text & "',num='" & num3.Text & "',price='" & price3.Text & "' where billservice_id='" & txt_id3.Text & "'"
        If cmd_excuteNonquery() = 0 Then
            MessageBox.Show("คลิกที่ตารางก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MsgBox("แก้ไขสำเร็จ")
        End If
        Combo_service3.Text = "เลือกรหัสการบริการ"
        Combo_bill3.Text = "เลือกรหัสลูกค้า"
        Combo_type3.Text = "เลือกรหัสประเภทเสื้อผ้า"
        auto_Number3()
        num_nuam.Text = ""
        num3.Text = ""
        refresh_edit3()
        refresh_sql3()
        price3.Text = ""
        lbl_cusname3.Text = ""
        lbl_typeservice3.Text = ""
        lbl_typeclothes.Text = ""
        lblwashing3.Text = ""
        lblroning3.Text = ""
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        sql = "delete from bill_Service where billservice_id='" & txt_id3.Text & "'"
        If cmd_excuteNonquery() = 0 Then
            MessageBox.Show("คลิกที่ตารางก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MsgBox("ลบข้อมูลสำเร็จ")
        End If
        lblCus.Text = ""
        Combo_service3.Text = "เลือกรหัสการบริการ"
        Combo_bill3.Text = "เลือกรหัสลูกค้า"
        Combo_type3.Text = "เลือกรหัสประเภทเสื้อผ้า"
        auto_Number3()
        num_nuam.Text = ""
        num3.Text = ""
        refresh_edit3()
        refresh_sql3()
        lbl_cusname3.Text = ""
        lbl_typeservice3.Text = ""
        lbl_typeclothes.Text = ""
        lblwashing3.Text = ""
        lblroning3.Text = ""
        price3.Text = ""
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Me.Close()
    End Sub
    Public Sub FormatGridview3()
        With DataGridView3

            .Columns(0).HeaderText = "รหัสใบเสร็จค่าบริการ"
            .Columns(1).HeaderText = "ชื้อลูกค้า"
            .Columns(2).HeaderText = "ขนาดผ้านวม"
            .Columns(3).HeaderText = "จำนวน"
            .Columns(4).HeaderText = "ราคาทั้งหมด"
            .Columns(5).HeaderText = "รหัสการบริการ"
            .Columns(6).HeaderText = "รหัสประเภทเสื้อผ้า"
            .Columns(7).HeaderText = "รหัสลูกค้า"
            .Columns(0).Width = 140
            .Columns(1).Width = 155
            .Columns(2).Width = 135
            .Columns(3).Width = 135
            .Columns(4).Width = 135
            .Columns(5).Width = 135
            .Columns(6).Width = 135
            .Columns(7).Width = 135
        End With
    End Sub

    Private Sub DataGridView3_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Dim i As Integer = DataGridView3.CurrentRow.Index
        txt_id3.Text = DataGridView3.Item(0, i).Value
        lbl_cusname3.Text = DataGridView3.Item(1, i).Value
        num_nuam.Text = DataGridView3.Item(2, i).Value
        num3.Text = DataGridView3.Item(3, i).Value
        price3.Text = DataGridView3.Item(4, i).Value
        Combo_service3.Text = DataGridView3.Item(5, i).Value
        Combo_type3.Text = DataGridView3.Item(6, i).Value
        Combo_bill3.Text = DataGridView3.Item(6, i).Value
    End Sub

    Private Sub PrintDocument2_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument2.PrintPage
        Dim fnt As Font
        Dim txt As String
        Dim black As New SolidBrush(Color.Black)
        Dim blue As New SolidBrush(Color.Blue)
        fnt = New Font("cordiaupc", 20, FontStyle.Bold)
        txt = "ใบเสร็จค่าบริการ"
        e.Graphics.DrawString(txt, fnt, black, 300, 50)
        txt = "รหัสใบเสร็จ"
        e.Graphics.DrawString(txt, fnt, black, 200, 90)
        txt = txt_id3.Text
        e.Graphics.DrawString(txt, fnt, black, 350, 90)
        txt = "ชื่อลูกค้า"
        e.Graphics.DrawString(txt, fnt, black, 200, 130)
        txt = lbl_cusname3.Text
        e.Graphics.DrawString(txt, fnt, black, 350, 130)
        txt = "ประเภทบริการ"
        e.Graphics.DrawString(txt, fnt, black, 200, 170)
        txt = lbl_typeservice3.Text
        e.Graphics.DrawString(txt, fnt, black, 350, 170)
        txt = "ประเภทเสื้อผ้า"
        e.Graphics.DrawString(txt, fnt, black, 200, 210)
        txt = lbl_typeclothes.Text
        e.Graphics.DrawString(txt, fnt, black, 350, 210)
        txt = "ราคาต่อถัง"
        e.Graphics.DrawString(txt, fnt, black, 200, 250)
        txt = lblwashing3.Text
        e.Graphics.DrawString(txt, fnt, black, 350, 250)
        txt = "ราคาต่อชิ้น"
        e.Graphics.DrawString(txt, fnt, black, 200, 290)
        txt = lblroning3.Text
        e.Graphics.DrawString(txt, fnt, black, 350, 290)
        txt = "ขนาดผ้านวม"
        e.Graphics.DrawString(txt, fnt, black, 200, 330)
        txt = num_nuam.Text
        e.Graphics.DrawString(txt, fnt, black, 350, 330)
        txt = "จำนวน"
        e.Graphics.DrawString(txt, fnt, black, 200, 370)
        txt = num3.Text
        e.Graphics.DrawString(txt, fnt, black, 350, 370)
        txt = "ราคารวม"
        e.Graphics.DrawString(txt, fnt, black, 200, 410)
        txt = price3.Text
        e.Graphics.DrawString(txt, fnt, black, 350, 410)
    End Sub

    Private Sub num_nuam_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles num_nuam.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8, 46, 48 To 57
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("กรอกได้เฉพาะตัวเลข")
        End Select
    End Sub

    Private Sub num3_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles num3.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8, 46, 48 To 57
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("กรอกได้เฉพาะตัวเลข")
        End Select
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        PrintPreviewDialog2.Document = PrintDocument2
        PrintPreviewDialog2.Width = 850
        PrintPreviewDialog2.Height = 600
        PrintPreviewDialog2.ShowDialog()
    End Sub
End Class