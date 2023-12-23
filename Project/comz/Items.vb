Imports System.Data.OleDb

Public Class Items
    Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\New folder\Vb\comz.mdb")
    Dim dsitem As New DataSet
    Dim aditem As New OleDbDataAdapter
    Dim n As Integer
    Dim chrDBcommand As Char
    Dim iSearch As String = "0"
    Dim c As String = "0"
    Dim p As Integer = 0

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        chrDBcommand = "A"
        txtItemid.ReadOnly = True
        txtItemName.Focus()
        n = dsitem.Tables("Item").Rows.Count - 1
        showRecords()
        Dim d As Integer = txtItemid.Text.Length - 1
        c = txtItemid.Text.Substring(1, d)
        p = CInt(c)
        txtItemid.Text = "I" & p + 1
        txtItemName.Clear()
        txtStock.Clear()
        txtCost.Clear()
        txtPrice.Clear()
        txtWarranty.Clear()
        

    End Sub

    Private Sub Items_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        con.Open()
        Dim cmitem As New OleDbCommand
        cmitem.Connection = con
        cmitem.CommandText = "SELECT * FROM tblItem"
        aditem.SelectCommand = cmitem
        aditem.Fill(dsitem, "Item")
        n = dsitem.Tables("Item").Rows.Count - 1
        con.Close()
        showRecords()
        iSearch = "0"
    End Sub

    Sub showRecords()
        Dim drCus As DataRow

        If n >= 0 Then
            drCus = dsitem.Tables("Item").Rows(n)

            With drCus
                txtItemid.Text = .Item("iId")
                txtItemName.Text = .Item("iName")
                txtStock.Text = .Item("iAmount")
                txtCost.Text = .Item("iCost")
                txtPrice.Text = .Item("iPrice")
                txtWarranty.Text = .Item("iWarranty")

            End With
        End If

    End Sub



    Private Sub btnLast_Click(sender As Object, e As EventArgs) Handles btnLast.Click
        Try
            n = dsitem.Tables("Item").Rows.Count - 1
            showRecords()
            iSearch = "0"

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnFirst_Click(sender As Object, e As EventArgs) Handles btnFirst.Click
        Try
            n = 0
            showRecords()
            iSearch = "0"

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        Try
            If n < dsitem.Tables("Item").Rows.Count - 1 Then
                n = n + 1
                showRecords()
                iSearch = "0"
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnPrevious_Click(sender As Object, e As EventArgs) Handles btnPrevious.Click
        Try
            If n > 0 Then
                n = n - 1
                showRecords()
                iSearch = "0"
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        If n >= 0 Then
            chrDBcommand = "E"
            txtItemid.Focus()
            txtItemid.ReadOnly = True
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            If iSearch = "1" Then
                Dim q As String = MessageBox.Show("Do you want to delete this data", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If q = vbYes Then

                    Dim cmDelete As New OleDbCommand("Delete from tblItem where iId=@iId", con)
                    cmDelete.Parameters.AddWithValue("@iId", txtSearch.Text)
                    n = n - 1
                    con.Open()
                    cmDelete.ExecuteNonQuery()
                    con.Close()
                    iSearch = "0"
                    n = 0
                    showRecords()
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
        showRecords()
        txtSearch.Clear()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim cmBuilder As New OleDbCommandBuilder
        cmBuilder.DataAdapter = aditem

        Try

            If chrDBcommand = "A" Then
                If txtItemid.Text = "" Or txtItemName.Text = "" Then
                    If txtItemid.Text = "" Then
                        MessageBox.Show("Please enter the item ID")
                    ElseIf txtItemName.Text = "" Then
                        MessageBox.Show("Please Enter the Item Name")
                    End If

                ElseIf IsNumeric(txtStock.Text) = False Or IsNumeric(txtPrice.Text) = False Or IsNumeric(txtCost.Text) = False Then
                    If IsNumeric(txtStock.Text) = False Then
                        MessageBox.Show("please enter numeric values to stock")
                    ElseIf IsNumeric(txtPrice.Text) = False Then
                        MessageBox.Show("please enter numbers to price")
                    ElseIf IsNumeric(txtCost.Text) = False Then
                        MessageBox.Show("please enter numeric values to cost")
                    End If

                Else
                    Dim drCus As DataRow
                    drCus = dsitem.Tables("Item").NewRow

                    With drCus

                        .Item("iId") = txtItemid.Text
                        .Item("iName") = txtItemName.Text
                        .Item("iAmount") = txtStock.Text
                        .Item("iCost") = txtCost.Text
                        .Item("iPrice") = txtPrice.Text
                        .Item("iWarranty") = txtWarranty.Text

                    End With
                    dsitem.Tables("Item").Rows.Add(drCus)
                    aditem.InsertCommand = cmBuilder.GetInsertCommand
                    n = n + 1
                    MessageBox.Show("Item added sucessfully")

                End If




            ElseIf chrDBcommand = "E" Then

                If txtItemid.Text = "" Or txtItemName.Text = "" Then
                    If txtItemid.Text = "" Then
                        MessageBox.Show("Please enter the item ID")
                    ElseIf txtItemName.Text = "" Then
                        MessageBox.Show("Please Enter the Item Name")
                    End If

                ElseIf IsNumeric(txtStock.Text) = False Or IsNumeric(txtPrice.Text) = False Or IsNumeric(txtCost.Text) = False Then
                    If IsNumeric(txtStock.Text) = False Then
                        MessageBox.Show("please enter numeric values to stock")
                    ElseIf IsNumeric(txtPrice.Text) = False Then
                        MessageBox.Show("please enter numbers to price")
                    ElseIf IsNumeric(txtCost.Text) = False Then
                        MessageBox.Show("please enter numeric values to cost")
                    End If

                Else

                    Dim tbCus As DataTable
                    Dim dcPrimaryKey(0) As DataColumn
                    tbCus = dsitem.Tables("Item")
                    dcPrimaryKey(0) = tbCus.Columns("iId")
                    tbCus.PrimaryKey = dcPrimaryKey
                    Dim drCus As DataRow = tbCus.Rows.Find(txtItemid.Text)

                    With drCus


                        '.Item("iId") = txtItemid.Text
                        .Item("iName") = txtItemName.Text
                        .Item("iAmount") = txtStock.Text
                        .Item("iCost") = txtCost.Text
                        .Item("iPrice") = txtPrice.Text
                        .Item("iWarranty") = txtWarranty.Text

                    End With

                    aditem.UpdateCommand = cmBuilder.GetUpdateCommand
                    MessageBox.Show("Item updated sucessfully")

                End If
            ElseIf chrDBcommand = "D" Then
                dsitem.Tables("Item").Rows(n).Delete()
                aditem.DeleteCommand = cmBuilder.GetDeleteCommand
                n = n - 1
                MessageBox.Show("Item deleted sucessfully")

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        con.Open()
        Try

            aditem.Update(dsitem, "Item")
            clearControls()
            showRecords()
        Catch ex As Exception
            MessageBox.Show("You are trying to save data incorrectly")

        End Try
        con.Close()

    End Sub

    Sub clearControls()

        txtItemid.Clear()
        txtItemName.Clear()
        txtStock.Clear()
        txtCost.Clear()
        txtPrice.Clear()
        txtWarranty.Clear()

    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Try
            Dim tbiTem As DataTable
            Dim dcPrimaryKey(0) As DataColumn
            tbiTem = dsitem.Tables("Item")
            dcPrimaryKey(0) = tbiTem.Columns("iId")
            tbiTem.PrimaryKey = dcPrimaryKey

            Dim strCusID As String

            strCusID = txtSearch.Text

            If Not strCusID Is Nothing Then
                Dim drItem As DataRow = tbiTem.Rows.Find(strCusID)
                If Not drItem Is Nothing Then
                    With drItem
                        txtItemid.Text = .Item("iId")
                        txtItemName.Text = .Item("iName")
                        txtStock.Text = .Item("iAmount")
                        txtCost.Text = .Item("iCost")
                        txtPrice.Text = .Item("iPrice")
                        txtWarranty.Text = .Item("iWarranty")
                    End With
                    iSearch = "1"
                Else
                    MessageBox.Show("Item Not Found")
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

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub txtCost_TextChanged(sender As Object, e As EventArgs) Handles txtCost.TextChanged
        If txtCost.Text <> "" Then
            If IsNumeric(txtCost.Text) = False Then
                MsgBox("Please enter numbers")
            End If
        End If
    End Sub

    Private Sub txtPrice_TextChanged(sender As Object, e As EventArgs) Handles txtPrice.TextChanged
        If txtPrice.Text <> "" Then
            If IsNumeric(txtPrice.Text) = False Then
                MsgBox("Please enter numbers")
            End If
        End If
    End Sub

    Private Sub txtStock_TextChanged(sender As Object, e As EventArgs) Handles txtStock.TextChanged
        If txtStock.Text <> "" Then
            If IsNumeric(txtStock.Text) = False Then
                MsgBox("Please enter numbers")
            End If
        End If
    End Sub
End Class