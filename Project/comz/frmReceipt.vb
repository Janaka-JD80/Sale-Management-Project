Public Class frmReceipt

    Private Sub frmReceipt_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'comzDataSet.tblBuying' table. You can move, or remove it, as needed.
        Me.tblBuyingTableAdapter.Fill(Me.comzDataSet.tblBuying)

        Me.ReportViewer1.RefreshReport()
    End Sub
End Class