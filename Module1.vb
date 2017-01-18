Imports System.Data.SqlClient
Module Module1
    Friend cn As New SqlConnection("data source=.\SQLEXPRESS; Initial Catalog=laundry; Integrated Security=sspi;")
    Friend cmd As New SqlCommand
    Friend Da As New SqlDataAdapter
    Friend ds As DataSet
    Friend dr As SqlDataReader
    Friend sql As String = ""
    Friend Sub connect()
        If cn.State = ConnectionState.Closed Then cn.Open()
    End Sub
    'เพิ่ม/ลบ/แก้ไข ฐานข้อมูล'
    Friend Function cmd_excuteNonquery()
        connect()
        cmd = New SqlCommand(sql, cn)
        Return cmd.ExecuteNonQuery
    End Function
    'ดึงข้อมูลจาก DataGrid'
    Friend Function cmd_dataTable()
        Da = New SqlDataAdapter(sql, cn)
        ds = New DataSet
        Da.Fill(ds, "Table")
        Return ds.Tables("Table")
    End Function

    Friend Function cmd_excuteScalar()
        connect()
        cmd = New SqlCommand(sql, cn)
        Return cmd.ExecuteScalar()
    End Function
    Friend Sub cmd_database_to_object(ByVal obj As Object)
        connect()
        cmd = New SqlCommand(sql, cn)
        dr = cmd.ExecuteReader
        obj.Items.Clear()
        While dr.Read
            obj.Items.Add(dr(0))
        End While
        dr.Close()
    End Sub
    Friend Sub cmd_test(ByVal dd As Object)
       

        connect()
        cmd = New SqlCommand(sql, cn)
        dd.text = cmd_excuteScalar()




    End Sub
End Module
