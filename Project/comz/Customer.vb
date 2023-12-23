Imports System.Data.OleDb
Public Class Customer
    Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\New folder\Vb\comz.mdb")
    Dim adCus As New OleDbDataAdapter
    Dim ds As New DataSet
    Dim n As Integer
    Dim chrDBcommand As Char
    Dim Search As String = "0"
    Dim c As String = "0"
    Dim p As Integer = 0


    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Customer_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed


    End Sub

    Private Sub Customer_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

    End Sub

    Private Sub Customer_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        con.Open()

        Dim cmCus As New OleDbCommand
        cmCus.Connection = con
        cmCus.CommandText = "SELECT * FROM tblCustomer"

        adCus.SelectCommand = cmCus

        adCus.Fill(ds, "Customer")
        n = ds.Tables("Customer").Rows.Count - 1
        con.Close()
        showRecord()
        Search = "0"
       

    End Sub
    Sub showRecord()

        Dim drCus As DataRow

        If n >= 0 Then
            drCus = ds.Tables("Customer").Rows(n)

            With drCus
                txtCustomerId.Text = .Item("cID")
                txtCustomerName.Text = .Item("cName")
                txtAddress.Text = .Item("cAddress")
                txtTelNo.Text = .Item("cTelNo")
            End With
        End If

    End Sub

    Private Sub btnLast_Click(sender As Object, e As EventArgs) Handles btnLast.Click
        Try
            n = ds.Tables("Customer").Rows.Count - 1
            showRecord()
            Search = "0"
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub btnFirst_Click(sender As Object, e As EventArgs) Handles btnFirst.Click
        Try
            n = 0
            showRecord()
            Search = "0"
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        Try
            If n < ds.Tables("Customer").Rows.Count - 1 Then
                n = n + 1
                showRecord()
                Search = "0"
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub btnPrevious_Click(sender As Object, e As EventArgs) Handles btnPrevious.Click

        Try
            If n > 0 Then
                n = n - 1
                showRecord()
                Search = "0"
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        chrDBcommand = "A"
        txtCustomerName.Focus()
        txtCustomerId.ReadOnly = True
        n = ds.Tables("Customer").Rows.Count - 1
        showRecord()
        Dim d As Integer = txtCustomerId.Text.Length - 1
        c = txtCustomerId.Text.Substring(1, d)
        p = CInt(c)
        txtCustomerId.Text = "C" & p + 1
        txtCustomerName.Clear()
        txtAddress.Clear()
        txtTelNo.Clear()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        If n >= 0 Then
            chrDBcommand = "E"
            txtCustomerId.Focus()
            txtCustomerId.ReadOnly = True


        End If
    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim cmBuilder As New OleDbCommandBuilder
        cmBuilder.DataAdapter = adCus
        Try

       
        If chrDBcommand = "A" Then

            If txtCustomerId.Text = "" Or txtCustomerName.Text = "" Or txtTelNo.Text = "" Then

                If txtCustomerId.Text = "" Then
                    MsgBox("Please Enter the Customer Id")
                ElseIf txtCustomerName.Text = "" Then
                    MsgBox("Plese Enter the Customer Name")
                ElseIf txtTelNo.Text = "" Then
                    MsgBox("Please Enter the Tel-No")

                End If


            ElseIf IsNumeric(txtCustomerName.Text) = True Then
                MsgBox("You cant Enter Numbers....")


            ElseIf (txtTelNo.Text).Length <> 10 Or (txtTelNo.Text).Substring(0, 1) <> "0" Or IsNumeric(txtTelNo.Text) = False Then

                If (txtTelNo.Text).Length <> 10 Then
                    MsgBox("Not a valid Phone Number")
                ElseIf (txtTelNo.Text).Substring(0, 1) <> "0" Then
                    MsgBox("First digit must be zero")
                ElseIf IsNumeric(txtTelNo.Text) = False Then
                    MsgBox("please enter only numbers")
                End If

            Else

                Dim drCus As DataRow
                drCus = ds.Tables("Customer").NewRow


                With drCus
                    .Item("cID") = txtCustomerId.Text
                    .Item("cName") = txtCustomerName.Text
                    .Item("cAddress") = txtAddress.Text
                    .Item("cTelNo") = txtTelNo.Text

                End With
                ds.Tables("Customer").Rows.Add(drCus)
                adCus.InsertCommand = cmBuilder.GetInsertCommand
                n = n + 1
                MessageBox.Show("Data added succesfully")


            End If


        ElseIf chrDBcommand = "E" Then



            If txtCustomerId.Text = "" Or txtCustomerName.Text = "" Or txtTelNo.Text = "" Then

                If txtCustomerId.Text = "" Then
                    MsgBox("Please Enter the Customer Id")
                ElseIf txtCustomerName.Text = "" Then
                    MsgBox("Plese Enter the Customer Name")
                ElseIf txtTelNo.Text = "" Then
                    MsgBox("Please Enter the Tel-No")

                End If


            ElseIf IsNumeric(txtCustomerName.Text) = True Then
                MsgBox("You cant Enter Numbers....")


            ElseIf (txtTelNo.Text).Length <> 10 Or (txtTelNo.Text).Substring(0, 1) <> "0" Or IsNumeric(txtTelNo.Text) = False Then

                If (txtTelNo.Text).Length <> 10 Then
                    MsgBox("Not a valid Phone Number")
                ElseIf (txtTelNo.Text).Substring(0, 1) <> "0" Then
                    MsgBox("First digit must be zero")
                ElseIf IsNumeric(txtTelNo.Text) = False Then
                    MsgBox("please enter only numbers")
                End If


            Else

                Dim tbCus As DataTable
                Dim dcPrimaryKey(0) As DataColumn
                tbCus = ds.Tables("Customer")
                dcPrimaryKey(0) = tbCus.Columns("cID")
                tbCus.PrimaryKey = dcPrimaryKey
                Dim drCus As DataRow = tbCus.Rows.Find(txtCustomerId.Text)

                With drCus
                    '.Item("cID") = txtCID.Text
                    .Item("cName") = txtCustomerName.Text
                    .Item("cAddress") = txtAddress.Text
                    .Item("cTelNo") = txtTelNo.Text
                End With

                adCus.UpdateCommand = cmBuilder.GetUpdateCommand
                MessageBox.Show("Data updated successfully")

            End If



        ElseIf chrDBcommand = "D" Then
            ds.Tables("Customer").Rows(n).Delete()
            adCus.DeleteCommand = cmBuilder.GetDeleteCommand
            n = n - 1
            MessageBox.Show("Data deleted successfully")

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try

        con.Open()
        Try
            adCus.Update(ds, "Customer")
            clearControls()
            showRecord()
        Catch ex As Exception
            MessageBox.Show("You are trying to save data incorrectly")

        End Try
        con.Close()
       
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try

            If Search = "1" Then
                Dim y As String = MessageBox.Show("Do you want to delete this data", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If y = vbYes Then

                    Dim cmDelete As New OleDbCommand("Delete from tblCustomer where cID=@cID", con)
                    cmDelete.Parameters.AddWithValue("@cID", txtSearch.Text)
                    n = n - 1
                    con.Open()
                    cmDelete.ExecuteNonQuery()
                    con.Close()
                    Search = "0"
                    n = 0
                    ds.Clear()
                    con.Open()

                    Dim cmCus As New OleDbCommand
                    cmCus.Connection = con
                    cmCus.CommandText = "SELECT * FROM tblCustomer"
                    adCus.SelectCommand = cmCus
                    adCus.Fill(ds, "Customer")
                    n = ds.Tables("Customer").Rows.Count - 1
                    con.Close()
                    showRecord()
                    Search = "0"

                    showRecord()
                    MessageBox.Show("Data deleted sucessfully")

                End If


            Else

                If n >= 0 Then
                    chrDBcommand = "D"
                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        showRecord()
    End Sub
    Sub clearControls()
        txtCustomerId.Clear()
        txtCustomerName.Clear()
        txtAddress.Clear()
        txtTelNo.Clear()

    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        Try
            Dim tbCustomer As DataTable
            Dim dcPrimaryKey(0) As DataColumn
            tbCustomer = ds.Tables("Customer")
            dcPrimaryKey(0) = tbCustomer.Columns("cID")
            tbCustomer.PrimaryKey = dcPrimaryKey

            Dim strCusID As String

            strCusID = txtSearch.Text


            If Not strCusID Is Nothing Then
                Dim drCustomer As DataRow = tbCustomer.Rows.Find(strCusID)
                If Not drCustomer Is Nothing Then
                    With drCustomer
                        txtCustomerId.Text = .Item("cID")
                        txtCustomerName.Text = .Item("cName")
                        txtAddress.Text = .Item("cAddress")
                        txtTelNo.Text = .Item("cTelNo")
                    End With
                    Search = "1"
                Else
                    MessageBox.Show("Customer Not Found")
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        frmMain.Show()
        Me.Hide()
    End Sub

    Private Sub txtTelNo_TextChanged(sender As Object, e As EventArgs) Handles txtTelNo.TextChanged
        If txtTelNo.Text <> "" Then

            If (txtTelNo.Text).Length > 10 Then
                MsgBox("Not a valid Phone Number")

            ElseIf (txtTelNo.Text).Substring(0, 1) <> "0" Then
                MsgBox("First digit must be zero")

            ElseIf IsNumeric(txtTelNo.Text) = False Then
                MsgBox("please enter only numbers")
            End If
        End If
    End Sub
End Class