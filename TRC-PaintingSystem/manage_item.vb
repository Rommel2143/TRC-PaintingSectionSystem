Imports MySql.Data.MySqlClient
Public Class manage_item
    Private Sub manage_item_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles btn_saveitem.Click
        Try

            If txt_partcode.Text = "" Or txt_partname.Text = "" Then
                display_error("All fields are required!", 2)
                Exit Sub
            End If
            con.Close()
            con.Open()

            Dim checkQuery As String = "SELECT partcode FROM inventory_fg_masterlist WHERE partcode = @partcode"
            Using checkCmd As New MySqlCommand(checkQuery, con)
                checkCmd.Parameters.AddWithValue("@partcode", txt_partcode.Text)
                Using dr = checkCmd.ExecuteReader()
                    If dr.HasRows Then
                        display_error("Partcode already exists!", 2)
                        Exit Sub
                    End If
                End Using
            End Using

            ' Insert new record
            Dim insertQuery As String = "INSERT INTO inventory_fg_masterlist (partcode, partname, stockF1, stockU6, wipstock, section) " &
                                        "VALUES (@partcode, @partname, 0, 0, 0, 'PAINTING')"
            Using insertCmd As New MySqlCommand(insertQuery, con)
                insertCmd.Parameters.AddWithValue("@partcode", txt_partcode.Text)
                insertCmd.Parameters.AddWithValue("@partname", txt_partname.Text)
                insertCmd.ExecuteNonQuery()
            End Using

            txt_partcode.Clear()
            txt_partname.Clear()
            MessageBox.Show("Record Saved!")
            hide_error()

        Catch ex As Exception
            display_error("Error: " & ex.Message, 2)
        Finally
            con.Close()

        End Try
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles btn_saveuser.Click

        Try

            If txt_idno.Text = "" Or txt_first.Text = "" Or txt_last.Text = "" Then
                display_error("All fields are required!", 2)
                Exit Sub
            End If
            con.Close()
            con.Open()

            Dim checkQuery As String = "SELECT IDno FROM `trc_user` WHERE IDno = @idno"
            Using checkCmd As New MySqlCommand(checkQuery, con)
                checkCmd.Parameters.AddWithValue("@idno", txt_idno.Text)
                Using dr = checkCmd.ExecuteReader()
                    If dr.HasRows Then
                        display_error("User already exists!", 2)
                        Exit Sub
                    End If
                End Using
            End Using

            ' Insert new record
            Dim insertQuery As String = "INSERT INTO `trc_user`(`IDno`, `firstname`, `middle`, `last`, `password`,`section`, `level`) " &
                                        "VALUES (@IDno, @firstname, @middle,@last,NULL,'PAINTING' ,'1')"
            Using insertCmd As New MySqlCommand(insertQuery, con)
                insertCmd.Parameters.AddWithValue("@IDno", txt_idno.Text)
                insertCmd.Parameters.AddWithValue("@firstname", txt_first.Text)
                insertCmd.Parameters.AddWithValue("@middle", txt_middle.Text)
                insertCmd.Parameters.AddWithValue("@last", txt_last.Text)
                insertCmd.ExecuteNonQuery()
            End Using

            txt_idno.Clear()
            txt_first.Clear()
            txt_middle.Clear()
            txt_last.Clear()
            MessageBox.Show("User saved!")
            hide_error()

        Catch ex As Exception
            display_error("Error: " & ex.Message, 2)
        Finally
            con.Close()

        End Try




    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click

    End Sub
End Class