Imports MySql.Data.MySqlClient
Public Class scan_results
    Dim status As String
    Private Sub scan_results_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtpicker.Value = Date.Now
    End Sub



    Public Sub loaddata(text As String)
        status = text
        Select Case status
            Case "IN"
                btn_report.Visible = False
            Case "OUT"
                btn_report.Visible = True
        End Select
    End Sub

    Private Sub dtpicker_ValueChanged(sender As Object, e As EventArgs) Handles dtpicker.ValueChanged
        Try
            Select Case status
                Case "IN"
                    con.Close()
                    con.Open()
                    Dim cmdselect As New MySqlCommand("Select distinct  CONCAT(`firstname`,' ',`last`) as fullname FROM `painting_stock`
                                                INNER JOIN `trc_user` ON `userin` = `IDno`
                                                WHERE  `datein`='" & dtpicker.Value.ToString("yyyy-MM-dd") & "'", con)
                    dr = cmdselect.ExecuteReader
                    cmbuser.Items.Clear()
                    While (dr.Read())
                        cmbuser.Items.Add(dr.GetString("fullname"))
                    End While

                Case "OUT"
                    con.Close()
                    con.Open()
                    Dim cmdselect As New MySqlCommand("Select distinct CONCAT(`firstname`,' ',`last`) as fullname FROM `painting_stock`
                                                INNER JOIN `trc_user` ON `userout` = `IDno`
                                                WHERE  `dateout`='" & dtpicker.Value.ToString("yyyy-MM-dd") & "' and status='OUT'", con)
                    dr = cmdselect.ExecuteReader
                    cmbuser.Items.Clear()
                    While (dr.Read())
                        cmbuser.Items.Add(dr.GetString("fullname"))
                    End While

            End Select

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Guna2Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Guna2Panel1.Paint

    End Sub

    Private Sub cmbuser_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbuser.SelectedIndexChanged
        Try
            Select Case status
                Case "IN"
                    con.Close()
                    con.Open()
                    Dim cmdselect As New MySqlCommand("SELECT DISTINCT ts.`batch` 
                                               FROM `painting_stock` ts
                                               LEFT JOIN trc_user tsoout 
                                               ON ts.userin = tsoout.IDno
                                               WHERE `datein` = @datein 
                                               AND CONCAT(`firstname`, ' ', `last`) = @fullname", con)

                    cmdselect.Parameters.AddWithValue("@datein", dtpicker.Value.ToString("yyyy-MM-dd"))
                    cmdselect.Parameters.AddWithValue("@fullname", cmbuser.Text)

                    dr = cmdselect.ExecuteReader()
                    cmb_batch.Items.Clear()
                    While dr.Read()
                        cmb_batch.Items.Add(dr.GetString("batch"))
                    End While
                Case "OUT"
                    con.Close()
                    con.Open()
                    Dim cmdselect As New MySqlCommand("SELECT DISTINCT ts.`batchout` 
                                               FROM `painting_stock` ts
                                               LEFT JOIN trc_user tsoout 
                                               ON ts.userout = tsoout.IDno
                                               WHERE `dateout` = @dateout
                                               AND CONCAT(`firstname`, ' ', `last`) = @fullname and status='OUT'", con)

                    cmdselect.Parameters.AddWithValue("@dateout", dtpicker.Value.ToString("yyyy-MM-dd"))
                    cmdselect.Parameters.AddWithValue("@fullname", cmbuser.Text)

                    dr = cmdselect.ExecuteReader()
                    cmb_batch.Items.Clear()
                    While dr.Read()
                        cmb_batch.Items.Add(dr.GetString("batchout"))
                    End While



            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub refreshgrid_in()
        Try
            con.Close()
            con.Open()
            Dim cmdrefreshgrid As New MySqlCommand("SELECT `partname`,ts.partcode,  `lotnumber`, `remarks`, `qty` FROM `painting_stock` ts
                                                    JOIN painting_masterlist fm ON ts.partcode=fm.partcode
                                                    LEFT JOIN trc_user tsoout  ON ts.userin = tsoout.IDno
                                                    WHERE `datein`='" & dtpicker.Value.ToString("yyyy-MM-dd") & "' and CONCAT(`firstname`, ' ', `last`)='" & cmbuser.Text & "' and `batch`='" & cmb_batch.Text & "'", con)

            Dim da As New MySqlDataAdapter(cmdrefreshgrid)
            Dim dt As New DataTable
            da.Fill(dt)
            datagrid1.DataSource = dt
            'datagrid1.AutoResizeColumns()
            da.Dispose()

            con.Close()
            con.Open()
            Dim cmdrefreshgrid2 As New MySqlCommand("SELECT `partname`, SUM(`qty`) AS TOTAL FROM `painting_stock` ts
                                                    JOIN painting_masterlist fm ON ts.partcode=fm.partcode
                                                    LEFT JOIN trc_user tsoout  ON ts.userin = tsoout.IDno
                                                    WHERE `datein`='" & dtpicker.Value.ToString("yyyy-MM-dd") & "' and CONCAT(`firstname`, ' ', `last`)='" & cmbuser.Text & "' and `batch`='" & cmb_batch.Text & "' 
                                                    GROUP BY ts.partcode", con)

            Dim da2 As New MySqlDataAdapter(cmdrefreshgrid2)
            Dim dt2 As New DataTable
            da2.Fill(dt2)
            datagrid2.DataSource = dt2
            datagrid2.AutoResizeColumns()
            da2.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally

            con.Close()
        End Try
    End Sub

    Private Sub refreshgrid_out()
        Try
            con.Close()
            con.Open()
            Dim cmdrefreshgrid As New MySqlCommand("SELECT `partname`,ts.partcode,  `lotnumber`, `remarks`, `qty` FROM `painting_stock` ts
                                                    JOIN painting_masterlist fm ON ts.partcode=fm.partcode
                                                    LEFT JOIN trc_user tsoout  ON ts.userout = tsoout.IDno
                                                    WHERE `dateout`='" & dtpicker.Value.ToString("yyyy-MM-dd") & "' and CONCAT(`firstname`, ' ', `last`)='" & cmbuser.Text & "' and `batchout`='" & cmb_batch.Text & "'  and status='OUT'", con)

            Dim da As New MySqlDataAdapter(cmdrefreshgrid)
            Dim dt As New DataTable
            da.Fill(dt)
            datagrid1.DataSource = dt
            'datagrid1.AutoResizeColumns()
            da.Dispose()

            con.Close()
            con.Open()
            Dim cmdrefreshgrid2 As New MySqlCommand("SELECT `partname`,ts.partcode, SUM(`qty`) FROM `painting_stock` ts
                                                    JOIN painting_masterlist fm ON ts.partcode=fm.partcode
                                                    LEFT JOIN trc_user tsoout  ON ts.userout = tsoout.IDno
                                                    WHERE `dateout`='" & dtpicker.Value.ToString("yyyy-MM-dd") & "' and CONCAT(`firstname`, ' ', `last`)='" & cmbuser.Text & "' and `batchout`='" & cmb_batch.Text & "' and status='OUT'
                                                    GROUP BY ts.partcode", con)

            Dim da2 As New MySqlDataAdapter(cmdrefreshgrid2)
            Dim dt2 As New DataTable
            da2.Fill(dt2)
            datagrid2.DataSource = dt2
            datagrid2.AutoResizeColumns()
            da2.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally

            con.Close()
        End Try
    End Sub

    Private Sub cmb_batch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_batch.SelectedIndexChanged
        Select Case status
            Case "IN"
                refreshgrid_in()
            Case "OUT"
                refreshgrid_out()
        End Select
    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles btn_report.Click
        Dim print As New print_report
        print.viewdata(dtpicker.Value.ToString("yyyy-MM-dd"), cmbuser.Text, cmb_batch.Text)
        print.ShowDialog()
        print.BringToFront()

    End Sub
End Class