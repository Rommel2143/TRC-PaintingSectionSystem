
Imports MySql.Data.MySqlClient
    Imports ClosedXML.Excel
Public Class filter_day
    Dim month As Integer


    Private Sub export_excel_Click(sender As Object, e As EventArgs) Handles export_excel.Click
        Try
            If datagrid1.Rows.Count > 0 Then
                Dim dt As New DataTable()

                ' Adding the Columns
                For Each column As DataGridViewColumn In datagrid1.Columns
                    dt.Columns.Add(column.HeaderText, column.ValueType)
                Next

                ' Adding the Rows
                For Each row As DataGridViewRow In datagrid1.Rows
                    If Not row.IsNewRow Then
                        dt.Rows.Add()
                        For Each cell As DataGridViewCell In row.Cells
                            dt.Rows(dt.Rows.Count - 1)(cell.ColumnIndex) = cell.Value.ToString()
                        Next
                    End If
                Next

                ' Save the data to an Excel file
                Using sfd As New SaveFileDialog()
                    sfd.Filter = "Excel Workbook|*.xlsx"
                    sfd.Title = "Save an Excel File"
                    sfd.ShowDialog()

                    If sfd.FileName <> "" Then
                        Using wb As New XLWorkbook()
                            wb.Worksheets.Add(dt, "Sheet1")
                            wb.SaveAs(sfd.FileName)
                        End Using
                        MessageBox.Show("Data successfully exported to Excel.", "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End Using
            Else
                MessageBox.Show("No data available to export.", "Export Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub filter_month_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub dtpicker1_ValueChanged(sender As Object, e As EventArgs) Handles dtpicker1.ValueChanged
        reload("SELECT 
    pm.partname, 
    ps.partcode, 
    
    SUM(CASE WHEN ps.datein IS NOT NULL THEN ps.qty ELSE 0 END) AS Produced_Parts,
SUM(CASE WHEN ps.dateout IS NOT NULL THEN ps.qty ELSE 0 END) AS Delivered
FROM 
    painting_stock ps
LEFT JOIN 
    painting_masterlist pm 
ON 
    pm.partcode = ps.partcode
WHERE 
    ps.dateout = '" & dtpicker1.Value.ToString("yyyy-MM-dd") & "'
    OR
  ps.datein = '" & dtpicker1.Value.ToString("yyyy-MM-dd") & "'
GROUP BY 
    ps.partcode, pm.partname
", datagrid1)
    End Sub
End Class