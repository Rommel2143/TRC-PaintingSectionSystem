Imports MySql.Data.MySqlClient

Public Class filter_month
    Private Sub txt_search_TextChanged(sender As Object, e As EventArgs)
        loaddata()
    End Sub

    Private Sub loaddata()
        If cmb_box.SelectedIndex = -1 Then
            MessageBox.Show("Please select a month.")
            Exit Sub
        End If

        Dim selectedMonth As Integer = cmb_box.SelectedIndex + 1
        Dim selectedYear As Integer = Guna2NumericUpDown1.Value

        ' Compute the last day of the selected month
        Dim endOfMonth As New DateTime(selectedYear, selectedMonth, DateTime.DaysInMonth(selectedYear, selectedMonth))
        Dim endOfMonthStr As String = endOfMonth.ToString("yyyy-MM-dd")

        Dim query As String = "
        SELECT 
            mm.partcode,
            mm.partname,
            IFNULL(painting.total_in, 0) AS Molding_IN,
            IFNULL(painting.total_out, 0) AS Molding_OUT,
            IFNULL(painting.box_count, 0) AS Molding_BoxCount,
            IFNULL(painting.total, 0) AS Molding_Total
        FROM painting_masterlist mm
        LEFT JOIN (
            SELECT partcode,
                SUM(CASE WHEN MONTH(dateIN) = " & selectedMonth & " THEN qty ELSE 0 END) AS total_in,
                SUM(CASE WHEN MONTH(dateOUT) = " & selectedMonth & " THEN qty ELSE 0 END) AS total_out,
                (SUM(CASE WHEN dateIN <= '" & endOfMonthStr & "' THEN 1 ELSE 0 END) -
                 SUM(CASE WHEN dateOUT <= '" & endOfMonthStr & "' THEN 1 ELSE 0 END)) AS box_count,
                (SUM(CASE WHEN dateIN <= '" & endOfMonthStr & "' THEN qty ELSE 0 END) -
                 SUM(CASE WHEN dateOUT <= '" & endOfMonthStr & "' THEN qty ELSE 0 END)) AS total
            FROM painting_stock
            GROUP BY partcode
        ) AS painting ON mm.partcode = painting.partcode
    "

        reload(query, datagrid1)
    End Sub





    Private Sub StyleDataGrid()
        Dim moldingColor As Color = Color.LightSkyBlue
        Dim unit56Color As Color = Color.LightGreen
        Dim sunboColor As Color = Color.LightSalmon

        With datagrid1
            ' Molding group header colors
            If .Columns.Contains("Molding_IN") Then
                .Columns("Molding_IN").HeaderText = "M_IN"
                .Columns("Molding_IN").HeaderCell.Style.BackColor = moldingColor
                .Columns("Molding_IN").HeaderCell.Style.ForeColor = Color.Black
            End If
            If .Columns.Contains("Molding_OUT") Then
                .Columns("Molding_OUT").HeaderText = "M_OUT"
                .Columns("Molding_OUT").HeaderCell.Style.BackColor = moldingColor
                .Columns("Molding_OUT").HeaderCell.Style.ForeColor = Color.Black
            End If
            If .Columns.Contains("Molding_Total") Then
                .Columns("Molding_Total").HeaderText = "Molding Total"
                .Columns("Molding_Total").HeaderCell.Style.BackColor = moldingColor
                .Columns("Molding_Total").HeaderCell.Style.ForeColor = Color.Black
            End If
            If .Columns.Contains("Molding_BoxCount") Then
                .Columns("Molding_BoxCount").HeaderText = "M_Boxes"
                .Columns("Molding_BoxCount").HeaderCell.Style.BackColor = moldingColor
                .Columns("Molding_BoxCount").HeaderCell.Style.ForeColor = Color.Black
            End If

            ' Unit56 group header colors
            If .Columns.Contains("Unit56_IN") Then
                .Columns("Unit56_IN").HeaderText = "U_IN"
                .Columns("Unit56_IN").HeaderCell.Style.BackColor = unit56Color
                .Columns("Unit56_IN").HeaderCell.Style.ForeColor = Color.Black
            End If
            If .Columns.Contains("Unit56_OUT") Then
                .Columns("Unit56_OUT").HeaderText = "U_OUT"
                .Columns("Unit56_OUT").HeaderCell.Style.BackColor = unit56Color
                .Columns("Unit56_OUT").HeaderCell.Style.ForeColor = Color.Black
            End If
            If .Columns.Contains("Unit56_Total") Then
                .Columns("Unit56_Total").HeaderText = "Unit 5-6 Total"
                .Columns("Unit56_Total").HeaderCell.Style.BackColor = unit56Color
                .Columns("Unit56_Total").HeaderCell.Style.ForeColor = Color.Black
            End If
            If .Columns.Contains("Unit56_BoxCount") Then
                .Columns("Unit56_BoxCount").HeaderText = "U_Boxes"
                .Columns("Unit56_BoxCount").HeaderCell.Style.BackColor = unit56Color
                .Columns("Unit56_BoxCount").HeaderCell.Style.ForeColor = Color.Black
            End If

            ' Sunbo group header colors
            If .Columns.Contains("Sunbo_IN") Then
                .Columns("Sunbo_IN").HeaderText = "S_IN"
                .Columns("Sunbo_IN").HeaderCell.Style.BackColor = sunboColor
                .Columns("Sunbo_IN").HeaderCell.Style.ForeColor = Color.Black
            End If
            If .Columns.Contains("Sunbo_OUT") Then
                .Columns("Sunbo_OUT").HeaderText = "S_OUT"
                .Columns("Sunbo_OUT").HeaderCell.Style.BackColor = sunboColor
                .Columns("Sunbo_OUT").HeaderCell.Style.ForeColor = Color.Black
            End If
            If .Columns.Contains("Sunbo_Total") Then
                .Columns("Sunbo_Total").HeaderText = "Sunbo Total"
                .Columns("Sunbo_Total").HeaderCell.Style.BackColor = sunboColor
                .Columns("Sunbo_Total").HeaderCell.Style.ForeColor = Color.Black
            End If
            If .Columns.Contains("Sunbo_BoxCount") Then
                .Columns("Sunbo_BoxCount").HeaderText = "S_Boxes"
                .Columns("Sunbo_BoxCount").HeaderCell.Style.BackColor = sunboColor
                .Columns("Sunbo_BoxCount").HeaderCell.Style.ForeColor = Color.Black
            End If


            .EnableHeadersVisualStyles = False ' Important to allow header style changes to show
            .AutoResizeColumns()
        End With
    End Sub

    Private Sub fg_monitoring_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cmb_box.SelectedItem = Date.Now.ToString("MMMM")

    End Sub



    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles Guna2Button2.Click
        exportExcel(datagrid1, "Molding FG Stocks", cmb_box.Text)
    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        loaddata()
    End Sub
End Class