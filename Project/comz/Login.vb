Imports System.Data
Imports System.Data.OleDb

Public Class Login


    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        If (txtUser.Text = "Admin" And txtPasword.Text = "12345") Then
            frmMain.lbllog.Text = "Admin"
            frmMain.Show()
            txtUser.Clear()
            txtPasword.Clear()
            txtUser.Focus()
            Me.Hide()


        ElseIf (txtUser.Text = "Sub" And txtPasword.Text = "98765") Then
            frmMain.lbllog.Text = "SubAdmin"
            frmMain.Show()
            txtUser.Clear()
            txtPasword.Clear()
            txtUser.Focus()
            Me.Hide()
        Else
            MessageBox.Show("Login failed.....try again", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtUser.Clear()
            txtPasword.Clear()
            txtUser.Focus()
        End If

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub PictureBox1_Click_1(sender As Object, e As EventArgs)

    End Sub
End Class
