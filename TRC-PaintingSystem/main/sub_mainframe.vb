Imports System
Public Class sub_mainframe


    Private Sub display_formscan(form As Form, tittle As String)
        With form
            .Refresh()
            .TopLevel = False
            Panel1.Controls.Add(form)
            .BringToFront()
            .Show()
            lbl_tittle.Text = tittle
        End With
    End Sub

    Private Sub INToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles INToolStripMenuItem.Click
        display_formscan(scan_in, "Scan IN")
    End Sub

    Private Sub OUTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OUTToolStripMenuItem.Click
        display_formscan(scan_out, "Scan OUT")
    End Sub

    Private Sub AddItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddItemToolStripMenuItem.Click
        display_formscan(manage_item, "Manage Item")
    End Sub
End Class