Imports MySql.Data.MySqlClient
Public Class box_info
    Private Sub box_info_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub displayinfo(qrcode As String)
        Try
            con.Close()
            con.Open()
            Dim cmdrefreshgrid As New MySqlCommand("SELECT * FROM `painting_stock`
                                                    WHERE qrcode = '" & qrcode & "'", con)

            Dim da As New MySqlDataAdapter(cmdrefreshgrid)
            Dim dt As New DataTable
            da.Fill(dt)
            datagrid1.DataSource = dt
            'datagrid1.AutoResizeColumns()
            da.Dispose()


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally

            con.Close()
        End Try
    End Sub

    Private Sub txtqr_TextChanged(sender As Object, e As EventArgs) Handles txtqr.TextChanged


    End Sub

    Private Sub txtqr_KeyDown(sender As Object, e As KeyEventArgs) Handles txtqr.KeyDown
        If e.KeyCode = Keys.Enter Then

            displayinfo(txtqr.Text.Trim)
        End If
    End Sub
End Class