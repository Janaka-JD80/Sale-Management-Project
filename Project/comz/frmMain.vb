Imports System.Data.OleDb
Public Class frmMain


    Dim j As Integer = 0

    Private Sub Button1_Click(sender As Object, e As EventArgs)




    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub ItemToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles picCustomer.Click
        Customer.Show()
        Me.Hide()
    End Sub

    Private Sub PictureBox1_Click_1(sender As Object, e As EventArgs) Handles picItem.Click
        Items.Show()
        Me.Hide()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles picBuying.Click
        Buying.Refresh()
        Buying.Show()
        Me.Hide()
    End Sub

    Private Sub PictureBox1_Click_2(sender As Object, e As EventArgs) Handles PicReport.Click
        frmReports.Show()
    End Sub

    Private Sub picLogout_Click(sender As Object, e As EventArgs) Handles picLogout.Click
        Login.Show()
        Me.Close()
        Customer.Close()
        Items.Close()
        Buying.Close()
    End Sub
End Class