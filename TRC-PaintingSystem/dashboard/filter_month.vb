Imports MySql.Data.MySqlClient

Public Class filter_month
    Private Sub txt_search_TextChanged(sender As Object, e As EventArgs)
        loaddata()
    End Sub


    Private Sub loaddata()
        Dim selectedMonthName As String = cmb_box.SelectedItem.ToString()
        Dim selectedYear As Integer = 2025

        ' Convert the selected month name to a month number (1-12)
        Dim monthNumber As Integer = DateTime.ParseExact(selectedMonthName, "MMMM", Globalization.CultureInfo.InvariantCulture).Month

        ' Create Date objects for the start and end of the month
        Dim startOfMonthDate As New DateTime(selectedYear, monthNumber, 1)
        Dim endOfMonthDate As DateTime = startOfMonthDate.AddMonths(1).AddDays(-1)

        ' Convert to string in format "yyyy-MM-dd"
        Dim startOfMonth As String = startOfMonthDate.ToString("yyyy-MM-dd")
        Dim endOfMonth As String = endOfMonthDate.ToString("yyyy-MM-dd")


        Dim query As String = "
    SELECT 
        mm.partcode,
        mm.partname,

        IFNULL(mold.total_in, 0) AS 'IN',
        IFNULL(mold.total_out, 0) AS 'OUT',
        IFNULL(mold.box_count, 0) AS 'Box',
        IFNULL(mold.total_in, 0) - IFNULL(mold.total_out, 0) AS Total

    FROM painting_masterlist mm

    LEFT JOIN (
        SELECT partcode,

            -- Sum of qty IN and OUT for the whole month
            SUM(CASE WHEN dateIN BETWEEN '" & startOfMonth & "' AND '" & endOfMonth & "' THEN qty ELSE 0 END) AS total_in,
            SUM(CASE WHEN dateOUT BETWEEN '" & startOfMonth & "' AND '" & endOfMonth & "' THEN qty ELSE 0 END) AS total_out,

            -- Box count only before selected date
            (SUM(CASE WHEN dateIN < '" & endOfMonth & "' THEN 1 ELSE 0 END) -
             SUM(CASE WHEN dateOUT < '" & endOfMonth & "' THEN 1 ELSE 0 END)) AS box_count

        FROM painting_stock
        GROUP BY partcode
    ) AS mold ON mm.partcode = mold.partcode
    "

        reload(query, datagrid1)
        StyleDataGrid()
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

    Private Sub dtpicker1_ValueChanged(sender As Object, e As EventArgs)
        loaddata()
    End Sub

    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles Guna2Button2.Click
        exportExcel(datagrid1, "Molding FG Stocks", cmb_box.Text)
    End Sub





End Class