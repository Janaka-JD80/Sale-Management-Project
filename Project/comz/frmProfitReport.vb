Public Class frmProfitReport

    Private Sub frmProfitReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'comzDataSet.BuyingReport' table. You can move, or remove it, as needed.
        Me.BuyingReportTableAdapter.Fill(Me.comzDataSet.BuyingReport)

        Me.ReportViewer1.RefreshReport()
    End Sub

    Private Sub ReportViewer1_Load(sender As Object, e As EventArgs) Handles ReportViewer1.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "MM/dd/yyyy"


        If Me.BuyingReportTableAdapter.Search(Me.comzDataSet.BuyingReport, DateTimePicker1.Text) = Nothing Then
            MessageBox.Show("There is no Transactins in this day")
        Else
            Me.BuyingReportTableAdapter.Search(Me.comzDataSet.BuyingReport, DateTimePicker1.Text)
            Me.ReportViewer1.RefreshReport()
        End If


    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.BuyingReportTableAdapter.Fill(Me.comzDataSet.BuyingReport)
        Me.ReportViewer1.RefreshReport()

    End Sub
End Class