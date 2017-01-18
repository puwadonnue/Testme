Public Class Login

    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btn_login_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_login.Click
        If txt_username.Text = "" Or txt_password.Text = "" Then
            MsgBox("กรุณากรอกข้อมูลให้ครบ")
            Return
        End If
        sql = "select count(*) from Employee where Username = '" & txt_username.Text & "' AND Password='" & txt_password.Text & "'"
        Dim Users As Integer = cmd_excuteScalar()
        If Users > 1 Then
            MsgBox("Login สำเร็จ")
            Form1.Show()
            Me.Hide()
        Else
            MsgBox("Username หรือ Password ผิด")
        End If

    End Sub

   
End Class