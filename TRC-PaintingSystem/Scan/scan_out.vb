
Imports MySql.Data.MySqlClient
Public Class scan_out
    Dim batch As String
    Dim supplier As String

    'duplicate info
    Dim status As String
    Dim located As String
    Dim datein As Date
    Dim partcode As String

    Dim lotnumber As String
    Dim remarks As String
    Dim qty As Integer

    Private Sub scan_in_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Guna2TextBox2_TextChanged(sender As Object, e As EventArgs) Handles batchcode.TextChanged
        Try
            batch = batchcode.Text
            If batchcode.Text = "" Then
                txtqr.Enabled = False

            Else
                txtqr.Enabled = True
                refreshgrid()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub
    Private Sub Txtqr_KeyDown(sender As Object, e As KeyEventArgs) Handles txtqr.KeyDown
        If e.KeyCode = Keys.Enter Then

            txt_box.Enabled = True
            txt_box.Focus()

        End If

    End Sub

    Private Sub ProcessQRcode(qrcode As String)
        Try

            Dim parts() As String = qrcode.Split("|")

            'CON 1 : QR SPLITING
            If parts.Length >= 5 AndAlso parts.Length <= 8 Then
                partcode = parts(0).Remove(0, 2).Trim
                lotnumber = parts(2).Remove(0, 2).Trim
                qty = parts(3).Remove(0, 2).Trim
                remarks = parts(4).Remove(0, 2).Trim
                supplier = parts(1).Remove(0, 2).Trim

                'CON 2 : IF SCANNED
                con.Close()
                con.Open()
                Dim cmdselect As New MySqlCommand("SELECT `qrcode`,`status`,`located`,`datein` FROM `inventory_fg_scan` WHERE `qrcode`='" & qrcode & "' LIMIT 1", con)
                dr = cmdselect.ExecuteReader
                If dr.Read = True Then
                    status = dr.GetString("status")
                    datein = dr.GetDateTime("datein")

                    Select Case status


                        Case "IN"
                            'update
                            update_inventory_fg_scan(qrcode)
                            hide_error()
                        Case "OUT"
                            'duplicate
                            display_error("Duplicate Entry", 1)
                    End Select

                Else 'CON 2 : IF NOT SCANNED
                    display_error("Record doesn't exist", 0)

                End If

            Else  'CON 1 : QR SPLITING
                display_error("INVALID QR FORMAT!", 0)

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            txtqr.Text = ""
            txt_box.Clear()
            txt_box.Enabled = False
            txtqr.Focus()
            refreshgrid()
        End Try

    End Sub

    Private Sub refreshgrid()
        Try
            con.Close()
            con.Open()
            Dim cmdrefreshgrid As New MySqlCommand("SELECT `qrcode`,`partcode`,  `lotnumber`, `remarks`, `qty` FROM `inventory_fg_scan`
                                                    WHERE `dateout`='" & datedb & "' and `userout`='" & idno & "' and `batchout`='" & batch & "' and `located`='U1-4' ", con)

            Dim da As New MySqlDataAdapter(cmdrefreshgrid)
            Dim dt As New DataTable
            da.Fill(dt)
            datagrid1.DataSource = dt
            'datagrid1.AutoResizeColumns()
            da.Dispose()

            con.Close()
            con.Open()
            Dim cmdrefreshgrid2 As New MySqlCommand("SELECT `partcode`, SUM(`qty`) FROM `inventory_fg_scan`
                                                    WHERE `dateout`='" & datedb & "' and `batchout`='" & batch & "' and `located`='U1-4' and `userout`='" & idno & "'
                                                    GROUP BY partcode", con)

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

    Private Sub update_inventory_fg_scan(qr As String)
        Try

            con.Close()
            con.Open()
            Dim cmdupdate As New MySqlCommand("UPDATE `inventory_fg_scan` SET `status`='OUT',`batchout`='" & batch & "',`dateout`='" & datedb & "',`userout`='" & idno & "',`boxno`='" & txt_box.Text & "',`pcout`='" & PCname & "' WHERE qrcode='" & qr & "'", con)
            cmdupdate.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub

    Private Sub Guna2Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Guna2Panel1.Paint

    End Sub

    Private Sub txt_box_TextChanged(sender As Object, e As EventArgs) Handles txt_box.TextChanged

    End Sub

    Private Sub txt_box_KeyDown(sender As Object, e As KeyEventArgs) Handles txt_box.KeyDown
        If e.KeyCode = Keys.Enter Then

            ProcessQRcode(txtqr.Text)
            txt_box.Enabled = False

        End If
    End Sub

    Private Sub txtqr_TextChanged(sender As Object, e As EventArgs) Handles txtqr.TextChanged

    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        Dim results As New scan_results
        results.loaddata("OUT")
        results.ShowDialog()
        results.BringToFront()
    End Sub
End Class