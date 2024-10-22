Imports MySql.Data.MySqlClient
Public Class print_report
    Private Sub print_report_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Public Sub viewdata(dateout As String, name As String, batchout As String)
        Dim myrpt As New outgoing_delivery
        dt.Clear()
        con.Close()
        con.Open()
        Dim showreport As New MySqlCommand("SELECT
    ts.id,
ts.partcode,
ts.lotnumber,
ts.remarks,
    ts.qty,
    ts.batchout,
ts.userout,
 ts.dateout,
    ts.boxno,
    CONCAT(tsout.firstname, ' ', tsout.last) as Fullname,
    tm.partname
FROM
        painting_stock ts
    Left Join trc_user tsout ON ts.userout = tsout.IDno
    Left Join painting_masterlist tm ON ts.partcode = tm.partcode
WHERE
        ts.dateout = '" & dateout & "' and CONCAT(tsout.firstname, ' ', tsout.last) = '" & name & "' and ts.batchout='" & batchout & "' and ts.status='OUT'
ORDER BY
    tm.partname;", con)
        Dim da As New MySqlDataAdapter(showreport)
        da.Fill(dt)
        con.Close()


        myrpt.Database.Tables("outgoing_report").SetDataSource(dt)
        CrystalReportViewer1.ReportSource = Nothing
        CrystalReportViewer1.ReportSource = myrpt
    End Sub
End Class