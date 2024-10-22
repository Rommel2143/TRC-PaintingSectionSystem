Imports MySql.Data.MySqlClient
Public Class device_info
    Private Sub device_info_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con.Close()
        con.Open()
        Dim selectpc As New MySqlCommand("Select * from trc_device", con)
        dr = selectpc.ExecuteReader
        If dr.Read = True Then
            lbl_devicename.Text = dr.GetString("PCname")
            lbl_mac.Text = dr.GetString("PCmac")
            lbl_location.Text = dr.GetString("location")
        End If
    End Sub
End Class