'Imports System.Data.OleDb
'Imports System.Data
'Imports System.Data.SqlClient
Public Class employee
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        sql = "select count(*) from Employee where Em_id='" & txt_id.Text & "'"
        If cmd_excuteScalar() > 0 Then
            txt_username.Text = ""
            txt_password.Text = ""
            txt_tel.Text = ""
            txt_address.Text = ""
            txt_name.Text = ""
            txt_surname.Text = ""
            txt_username.ReadOnly = False
            auto_Number()
            refresh_edit()

            Return
        End If
        sql = "insert into Employee values('" & txt_id.Text & "','" & txt_name.Text & "','" & txt_surname.Text & "','" & txt_address.Text & "','" & txt_tel.Text & "','" & txt_username.Text & "','" & txt_password.Text & "','" & txt_autoid.Text & "')"
        If (txt_name.Text = "") Or (txt_surname.Text = "") Or (txt_address.Text = "") Or (txt_tel.Text = "") Or (txt_username.Text = "") Or (txt_password.Text = "") Then
            MessageBox.Show("กรุณาป้อนข้อมูลให้ครบ !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            ' MsgBox("เพิ่มข้อมูลไม่สำเร็จ")
        ElseIf cmd_excuteNonquery() >= 1 Then
            MsgBox("เพิ่มข้อมูลสำเร็จ")
            txt_name.Text = ""
            txt_surname.Text = ""
            txt_address.Text = ""
            txt_tel.Text = ""
            txt_username.Text = ""
            txt_password.Text = ""
            txt_autoid.Text = ""
            auto_Number()
            refresh_edit()

        Else
            txt_name.Text = ""
            txt_surname.Text = ""
            txt_address.Text = ""
            txt_tel.Text = ""
            txt_username.Text = ""
            txt_password.Text = ""

        End If


    End Sub
    Private Sub auto_Number()
        sql = "select max(auto_id) from Employee"
        Try
            Dim numchar_id As String = "Em-" & (cmd_excuteScalar() + 1).ToString.PadLeft(3, "0")
            txt_id.Text = numchar_id
            txt_autoid.Text = cmd_excuteScalar() + 1
        Catch ex As Exception
            MsgBox("no")
            txt_id.Text = "Em-001"
            txt_autoid.Text = 1
        End Try
    End Sub

    Private Sub employee_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        auto_Number()
        refresh_edit()
        FormatGridview()

    End Sub
    Private Sub refresh_edit()
        sql = "select * from Employee"
        DataGridView1.DataSource = cmd_dataTable()
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim i As Integer = DataGridView1.CurrentRow.Index
        txt_id.Text = DataGridView1.Item(0, i).Value
        txt_name.Text = DataGridView1.Item(1, i).Value
        txt_surname.Text = DataGridView1.Item(2, i).Value
        txt_address.Text = DataGridView1.Item(3, i).Value
        txt_tel.Text = DataGridView1.Item(4, i).Value
        txt_username.Text = DataGridView1.Item(5, i).Value
        txt_password.Text = DataGridView1.Item(6, i).Value
        txt_username.ReadOnly = True


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        sql = "update Employee set Em_name='" & txt_name.Text & "',Em_surname='" & txt_surname.Text & "',Em_address='" & txt_address.Text & "',Em_tel='" & txt_tel.Text & "' where Em_id='" & txt_id.Text & "'"
        If cmd_excuteNonquery() = 0 Then
            MessageBox.Show("คลิกที่ตารางก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MsgBox("แก้ไขสำเร็จ")
            txt_username.ReadOnly = False

        End If
        txt_id.Text = ""
        txt_name.Text = ""
        txt_surname.Text = ""
        txt_address.Text = ""
        txt_tel.Text = ""
        txt_username.Text = ""
        txt_password.Text = ""
        txt_autoid.Text = ""
        auto_Number()
        refresh_edit()
    End Sub
    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        FormatGridview()
        If e.RowIndex > DataGridView1.Rows.Count Then
            e.Value = Nothing
        ElseIf e.ColumnIndex = 6 Then
            e.Value = "●●●●●●●●●"
        End If

    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        sql = "delete from Employee where Em_id='" & txt_id.Text & "'"
        If cmd_excuteNonquery() = 0 Then
            MessageBox.Show("คลิกที่ตารางก่อน !!!", "ผลการตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MsgBox("ลบข้อมูลสำเร็จ")
        End If
        txt_id.Text = ""
        txt_name.Text = ""
        txt_surname.Text = ""
        txt_address.Text = ""
        txt_tel.Text = ""
        txt_username.Text = ""
        txt_password.Text = ""
        txt_autoid.Text = ""
        txt_username.ReadOnly = False

        auto_Number()
        refresh_edit()
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.VisibleChanged
        DataGridView1.Columns(7).Visible = False
    End Sub
    Public Sub FormatGridview()

        With DataGridView1

            .Columns(0).HeaderText = "รหัสพนักงาน"
            .Columns(1).HeaderText = "ชื่อ"
            .Columns(2).HeaderText = "นามสกุล"
            .Columns(3).HeaderText = "ที่อยุ่"
            .Columns(4).HeaderText = "เบอร์โทรศัพท์"
            .Columns(5).HeaderText = "ชื่อผู้ใช้"
            .Columns(6).HeaderText = "รหัสผ่าน"
            .Columns(0).Width = 140
            .Columns(1).Width = 115
            .Columns(2).Width = 135
            .Columns(3).Width = 115
            .Columns(4).Width = 145
            .Columns(5).Width = 135
            .Columns(6).Width = 135
        End With
    End Sub

    Private Sub txt_tel_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_tel.KeyPress
        Select Case Asc(e.KeyChar)
            Case 8, 46, 48 To 57
                e.Handled = False
            Case 13
                Dim a As Integer = Len(txt_tel.Text)
                If a < 10 Then
                    MsgBox("กรอกเบอร์โทรศัพท์ให้ครบ 10 ตัว")

                Else
                    txt_username.Focus()
                End If

            Case Else
                e.Handled = True
                MessageBox.Show("กรอกได้เฉพาะตัวเลข")
        End Select
    End Sub
   

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub txt_name_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_name.KeyPress
        If Asc(e.KeyChar) = 13 Then
                txt_surname.Focus()
            End If
    End Sub

    Private Sub txt_surname_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_surname.KeyPress
        If Asc(e.KeyChar) = 13 Then
            txt_address.Focus()
        End If
    End Sub

    Private Sub txt_address_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_address.KeyPress
        If Asc(e.KeyChar) = 13 Then
            txt_tel.Focus()
        End If
    End Sub

    Private Sub txt_username_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_username.KeyPress
        If Asc(e.KeyChar) = 13 Then
            
            txt_password.Focus()
        End If
    End Sub

    Private Sub txt_password_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_password.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Button1.Focus()
        End If

    End Sub

 
    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub txt_tel_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_tel.TextChanged

    End Sub
End Class