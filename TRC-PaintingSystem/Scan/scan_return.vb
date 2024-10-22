Imports MySql.Data.MySqlClient
Public Class scan_return
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
    Private Sub txtqr_TextChanged(sender As Object, e As EventArgs) Handles txtqr.TextChanged

    End Sub

    Private Sub txtqr_KeyDown(sender As Object, e As KeyEventArgs) Handles txtqr.KeyDown
        If e.KeyCode = Keys.Enter Then

            ProcessQRcode(txtqr.Text)

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
                Dim cmdselect As New MySqlCommand("SELECT `qrcode`,`status`,`located`,`datein` FROM `painting_stock` WHERE `qrcode`='" & qrcode & "' LIMIT 1", con)
                dr = cmdselect.ExecuteReader
                If dr.Read = True Then
                    status = dr.GetString("status")
                    datein = dr.GetDateTime("datein")

                    Select Case status


                        Case "IN"

                            display_error("Duplicate Entry", 1)

                        Case "OUT"
                            update_painting_stock(qrcode)


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

            txtqr.Focus()

        End Try

    End Sub

    Private Sub refreshgrid()
        Try
            datagrid1.Rows.Add(txtqr.Text, partcode, qty, lotnumber)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally

            con.Close()
        End Try
    End Sub

    Private Sub update_painting_stock(qr As String)
        Try

            con.Close()
            con.Open()
            Dim cmdupdate As New MySqlCommand("UPDATE `painting_stock` SET `status`='IN', returned='" & idno & "',pcin='" & PCname & "' WHERE qrcode='" & qr & "'", con)
            cmdupdate.ExecuteNonQuery()
            refreshgrid()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub
End Class