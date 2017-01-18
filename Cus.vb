Imports System.Data.SqlClient

Public Class Cus
    Dim sDate, eDate As Date
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()

    End Sub

    Private Sub Cus_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        auto_Number()
        refresh_edit()
        FormatGridview()
        txt_name.Select()

    End Sub

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        'time_d()
        'MsgBox(sDate)
        sql = "select count(*) from Customer where Cus_id='" & txt_id.Text & "'"
        If cmd_excuteScalar() > 0 Then
            txt_address.Text = ""
            txt_name.Text = ""
            txt_surname.Text = ""
            auto_Number()
            refresh_edit()
            Return
        End If
        'sql = "insert into Customer values('" & txt_id.Text & "','" & txt_name.Text & "','" & txt_surname.Text & "','" & txt_address.Text & "','" & txt_tel.Text & "','" & txt_autoid.Text & "')"
        'sql = "insert into Customer values(@Cus_id,@Cus_name,@Cus_Surname,@Cus_address,@Cus_add,@Cus_exit,@auto_id)"
        'cmd = New SqlCommand(sql, cn)
        'cmd.Parameters.Clear()
        'cmd.Parameters.AddWithValue("@Cus_id", txt_id.Text)
        'cmd.Parameters.AddWithValue("@Cus_name", txt_name.Text)
        'cmd.Parameters.AddWithValue("@Cus_Surname", txt_surname.Text)
        'cmd.Parameters.AddWithValue("@Cus_address", txt_address.Text)
        'cmd.Parameters.AddWithValue("@Cus_add", sDate)
        'cmd.Parameters.AddWithValue("@Cus_exit", eDate)
        'cmd.Parameters.AddWithValue("@auto_id", txt_autoid.Text)
        'If (txt_name.Text = "") Or (txt_surname.Text = "") Or (txt_address.Text = "") Or (txt_tel.Text = "") Then
        'MessageBox.Show("กรุณาป้อนข้อมูลให้ครบ !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        ' ElseIf cmd.ExecuteNonQuery > 0 Then
        'MsgBox("เพิ่มข้อมูลสำเร็จ")
        'txt_address.Text = ""
        'txt_name.Text = ""
        'txt_surname.Text = ""
        'txt_tel.Text = ""
        'time_d()
        'auto_Number()
        'refresh_edit()

        'Else
        'MsgBox("d")
        'txt_exit.Text = ""
        'txt_add.Text = ""
        'txt_add.Text = ""
        'txt_address.Text = ""
        'txt_name.Text = ""
        'txt_surname.Text = ""
        'txt_tel.Text = ""
        'End If
        sql = "insert into Customer values('" & txt_id.Text & "','" & txt_name.Text & "','" & txt_surname.Text & "','" & txt_address.Text & "','" & txt_tel.Text & "','" & txt_autoid.Text & "')"
        If (txt_name.Text = "") Or (txt_surname.Text = "") Or (txt_address.Text = "") Or (txt_tel.Text = "") Then
            MessageBox.Show("กรุณาป้อนข้อมูลให้ครบ !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            ' MsgBox("เพิ่มข้อมูลไม่สำเร็จ")
        ElseIf cmd_excuteNonquery() >= 1 Then
            MsgBox("เพิ่มข้อมูลสำเร็จ")
            txt_name.Text = ""
            txt_surname.Text = ""
            txt_address.Text = ""
            txt_tel.Text = ""
            txt_autoid.Text = ""
            auto_Number()
            refresh_edit()

        Else
            txt_name.Text = ""
            txt_surname.Text = ""
            txt_address.Text = ""
            txt_tel.Text = ""
        End If

    End Sub
    Private Sub auto_Number()
        sql = "select max(auto_id) from Customer"
        Try
            Dim numchar_id As String = "Cu-" & (cmd_excuteScalar() + 1).ToString.PadLeft(3, "0")
            txt_id.Text = numchar_id
            txt_autoid.Text = cmd_excuteScalar() + 1
        Catch ex As Exception
            txt_id.Text = "Cu-001"
            txt_autoid.Text = 1
        End Try
    End Sub
    Private Sub refresh_edit()
        sql = "select * from Customer"
        DataGridView1.DataSource = cmd_dataTable()
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim i As Integer = DataGridView1.CurrentRow.Index
        txt_id.Text = DataGridView1.Item(0, i).Value
        txt_name.Text = DataGridView1.Item(1, i).Value
        txt_surname.Text = DataGridView1.Item(2, i).Value
        txt_address.Text = DataGridView1.Item(3, i).Value
        txt_tel.Text = DataGridView1.Item(4, i).Value

    End Sub



    Private Sub btn_edit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_edit.Click
        sql = "update Customer set Cus_name='" & txt_name.Text & "',Cus_Surname='" & txt_surname.Text & "',Cus_address='" & txt_address.Text & "',Cus_Tel='" & txt_tel.Text & "' where Cus_id='" & txt_id.Text & "'"
      
        If cmd_excuteNonquery() = 0 Then
            MessageBox.Show("คลิกที่ตารางก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MsgBox("แก้ไขสำเร็จ")
        End If
        txt_id.Text = ""
        txt_name.Text = ""
        txt_surname.Text = ""
        txt_address.Text = ""
        txt_tel.Text = ""
        txt_autoid.Text = ""
        auto_Number()
        refresh_edit()
    End Sub

    Private Sub btn_del_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_del.Click
        sql = "delete from Customer where Cus_id='" & txt_id.Text & "'"
        If cmd_excuteNonquery() = 0 Then
            MessageBox.Show("คลิกที่ตารางก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MsgBox("ลบข้อมูลสำเร็จ")
        End If
        txt_id.Text = ""
        txt_name.Text = ""
        txt_surname.Text = ""
        txt_address.Text = ""
        txt_add.Text = ""
        txt_exit.Text = ""
        txt_autoid.Text = ""
        auto_Number()
        refresh_edit()
    End Sub
    Public Sub FormatGridview()

        With DataGridView1

            .Columns(0).HeaderText = "รหัสลูกค้า"
            .Columns(1).HeaderText = "ชื่อ"
            .Columns(2).HeaderText = "นามสกุล"
            .Columns(3).HeaderText = "ที่อยุ่"
            .Columns(4).HeaderText = "เบอร์โทรศัพท์"
            .Columns(0).Width = 140
            .Columns(1).Width = 115
            .Columns(2).Width = 135
            .Columns(3).Width = 115
            .Columns(4).Width = 145

        End With
    End Sub
   

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'sql = "select * from Customer where Cus_name='" & txt_search.Text & "'"

    End Sub

    'Public Sub time_d()
    ' sDate = Format(Me.DateTimePicker1.Value, "dd-MM-yyyy")
    '  eDate = Format(Me.DateTimePicker2.Value, "dd-MM-yyyy")
    '  End Sub
    Private Sub DataGridView1_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.VisibleChanged
        DataGridView1.Columns(5).Visible = False
    End Sub

    Private Sub txt_tel_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_tel.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8, 46, 48 To 57
                e.Handled = False
            Case 13
                Dim a As Integer = Len(txt_tel.Text)
                If a < 10 Then
                    MsgBox("กรอกเบอร์โทรศัพท์ให้ครบ 10 ตัว")

                Else
                    btn_add.Focus()
                End If
            Case Else
                e.Handled = True
                MessageBox.Show("กรอกได้เฉพาะตัวเลข")
        End Select
    End Sub

    Private Sub txt_name_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_name.KeyPress
        If Asc(e.KeyChar) = 13 Then
            txt_surname.Focus()
        End If
    End Sub

    Private Sub txt_surname_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_surname.KeyPress
        If Asc(e.KeyChar) = 13 Then
            txt_address.Focus()
        End If
    End Sub

    Private Sub txt_address_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_address.KeyPress
        If Asc(e.KeyChar) = 13 Then
            txt_tel.Focus()
        End If
    End Sub

    Private Sub txt_search_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_search.TextChanged
        sql = "select * from Customer Where Cus_name like '%" & txt_search.Text & "%'"
        DataGridView1.DataSource = cmd_dataTable()
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click

    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class