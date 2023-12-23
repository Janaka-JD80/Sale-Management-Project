Public Class frmReports

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frmProfitReport.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frmCustomerReport.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        frmMain.Show()
        Me.Close()

    End Sub
End Class