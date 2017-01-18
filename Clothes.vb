Public Class Clothes

    Private Sub TabPage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ประเภทเสื้อผ้า.Click

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.Close()

    End Sub

    Private Sub Clothes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        auto_Number()
        refresh_edit()
        FormatGridview()
        refresh_sql()
        auto_Number2()
        refresh_edit2()
        DataGridView1.Columns(2).Visible = False
        DataGridView2.Columns(4).Visible = False

    End Sub
    Private Sub auto_Number()
        sql = "select max(auto_id) from clothesType"
        Try

            Dim numchar_id As String = "CL-" & (cmd_excuteScalar() + 1).ToString.PadLeft(3, "0")
            txt_id.Text = numchar_id
            txt_autoid2.Text = cmd_excuteScalar() + 1
        Catch ex As Exception
            txt_id.Text = "CL-001"
            txt_autoid2.Text = 1
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        sql = "select count(*) from clothesType where clothestype_id='" & txt_id.Text & "'"
        If cmd_excuteScalar() > 0 Then
            ComboBox1.ValueMember = "เสื้อ"
            auto_Number()
            refresh_edit()
            Return
        End If
        sql = "insert into clothesType values('" & txt_id.Text & "','" & ComboBox1.Text & "','" & txt_autoid2.Text & "')"
        If cmd_excuteNonquery() >= 1 Then
            MsgBox("เพิ่มข้อมูลสำเร็จ")
            txt_autoid2.Text = ""
            auto_Number()
            refresh_edit()
        End If
    End Sub
    Private Sub refresh_edit()
        sql = "select * from clothesType"
        DataGridView1.DataSource = cmd_dataTable()
    End Sub
    Public Sub FormatGridview()

        With DataGridView1
            .Columns(0).HeaderText = "รหัสประเภทเสือผ้า"
            .Columns(1).HeaderText = "ประเภทเสื้อผ้า"
            .Columns(2).HeaderText = "รหัส"
            .Columns(0).Width = 140
            .Columns(1).Width = 115
            .Columns(2).Width = 115
        End With
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim i As Integer = DataGridView1.CurrentRow.Index
        txt_id.Text = DataGridView1.Item(0, i).Value
        ComboBox1.Text = DataGridView1.Item(1, i).Value
        txt_autoid2.Text = DataGridView1.Item(2, i).Value
    End Sub
    Private Sub DataGridView1_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.VisibleChanged

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        sql = "update clothesType set clothestype='" & ComboBox1.Text & "' where clothestype_id='" & txt_id.Text & "'"
        If cmd_excuteNonquery() = 0 Then
            MessageBox.Show("คลิกที่ตารางก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MsgBox("แก้ไขสำเร็จ")
        End If
        auto_Number()
        refresh_edit()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        sql = "delete from clothesType where clothestype_id='" & txt_id.Text & "'"
        If cmd_excuteNonquery() = 0 Then
            MessageBox.Show("คลิกที่ตารางก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MsgBox("ลบข้อมูลสำเร็จ")
        End If
        txt_autoid2.Text = ""
        auto_Number()
        refresh_edit()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Me.Close()
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''' tab2''''''''''''''''''''''''''''''''''''
    Private Sub refresh_sql()
        sql = "SELECT * FROM bill WHERE bill.bill_id not in (SELECT bill_id FROM clothesStatus)"
        cmd_database_to_object(ComboBox_billid)
    End Sub
    Private Sub auto_Number2()
        sql = "select max(auto_id) from clothesStatus"
        Try
            Dim numchar_id As String = "ST-" & (cmd_excuteScalar() + 1).ToString.PadLeft(3, "0")
            txt_id2.Text = numchar_id
            txt_autoid2.Text = cmd_excuteScalar() + 1
        Catch ex As Exception
            txt_id2.Text = "ST-001"
            txt_autoid2.Text = 1
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        sql = "select count(*) from clothesStatus where statuslaundry_id='" & txt_id2.Text & "'"
        If cmd_excuteScalar() > 0 Then
            auto_Number2()
            refresh_edit2()
            refresh_sql()
            ComboBox2.Text = "ว่าง"
            ComboBox3.Text = "ว่าง"
            ComboBox_billid.Text = "เลือกรหัสใบเสร็จ"
            Return
        End If
        sql = "insert into clothesStatus values('" & txt_id2.Text & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & ComboBox_billid.Text & "','" & txt_autoid2.Text & "')"
        If (ComboBox_billid.Text = "เลือกรหัสใบเสร็จ") Then
            MessageBox.Show("กรุณาเลือกรหัสใบเสร็จก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        ElseIf cmd_excuteNonquery() >= 1 Then
            MsgBox("เพิ่มข้อมูลสำเร็จ")
            txt_autoid2.Text = ""
            auto_Number2()
            refresh_edit2()
            refresh_sql()
            ComboBox2.Text = "ว่าง"
            ComboBox3.Text = "ว่าง"
            ComboBox_billid.Text = "เลือกรหัสใบเสร็จ"
        End If
    End Sub

    Private Sub refresh_edit2()
        sql = "select * from clothesStatus"
        DataGridView2.DataSource = cmd_dataTable()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        sql = "update clothesStatus set status_washing='" & ComboBox2.Text & "',status_roning='" & ComboBox3.Text & "',bill_id='" & ComboBox_billid.Text & "',auto_id='" & txt_autoid2.Text & "' where statuslaundry_id='" & txt_id2.Text & "'"
        If cmd_excuteNonquery() = 0 Then
            MessageBox.Show("คลิกที่ตารางก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MsgBox("แก้ไขสำเร็จ")
        End If
        auto_Number2()
        refresh_edit2()
        refresh_sql()
        ComboBox2.Text = "ว่าง"
        ComboBox3.Text = "ว่าง"
        ComboBox_billid.Text = "เลือกรหัสใบเสร็จ"
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        sql = "delete from clothesStatus where statuslaundry_id='" & txt_id2.Text & "'"
        If cmd_excuteNonquery() = 0 Then
            MessageBox.Show("คลิกที่ตารางก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MsgBox("ลบข้อมูลสำเร็จ")
        End If
        txt_autoid2.Text = ""
        auto_Number2()
        refresh_edit2()
        refresh_sql()
        ComboBox2.Text = "ว่าง"
        ComboBox3.Text = "ว่าง"
        ComboBox_billid.Text = "เลือกรหัสใบเสร็จ"
    End Sub
    Public Sub FormatGridview2()

        With DataGridView2
            .Columns(0).HeaderText = "รหัสสถานะการซัก "
            .Columns(1).HeaderText = "สถานะการซัก"
            .Columns(2).HeaderText = "สถานะการรีด"
            .Columns(3).HeaderText = "รหัสใบเสร็จ"
            .Columns(4).HeaderText = "รหัส"
            .Columns(0).Width = 140
            .Columns(1).Width = 125
            .Columns(2).Width = 125
            .Columns(3).Width = 125
            .Columns(4).Width = 125
        End With
    End Sub

    Private Sub DataGridView2_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim i As Integer = DataGridView2.CurrentRow.Index
        txt_id2.Text = DataGridView2.Item(0, i).Value
        ComboBox2.Text = DataGridView2.Item(1, i).Value
        ComboBox3.Text = DataGridView2.Item(2, i).Value
        ComboBox_billid.Text = DataGridView2.Item(3, i).Value
    End Sub

    Private Sub ComboBox_billid_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox_billid.KeyPress
        e.Handled = True
    End Sub

    Private Sub ComboBox3_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox3.KeyPress
        e.Handled = True
    End Sub

    Private Sub ComboBox2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox2.KeyPress
        e.Handled = True
    End Sub

    Private Sub ComboBox1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        e.Handled = True
    End Sub

    Private Sub txt_autoid2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_autoid2.Click

    End Sub
End Class