Public Class frmCustomerReport

    Private Sub frmCustomerReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'comzDataSet.CustomerReport' table. You can move, or remove it, as needed.
        Me.CustomerReportTableAdapter.Fill(Me.comzDataSet.CustomerReport)

        Me.ReportViewer1.RefreshReport()
    End Sub

    Private Sub ReportViewer1_Load(sender As Object, e As EventArgs) Handles ReportViewer1.Load

    End Sub
End Class