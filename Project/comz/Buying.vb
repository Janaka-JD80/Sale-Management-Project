Imports System.Data.OleDb

Public Class Buying

    Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\New folder\Vb\comz.mdb")
    Dim dsitem As New DataSet
    Dim aditem As New OleDbDataAdapter
    Dim n As Integer
    Dim adBuy As New OleDbDataAdapter
    Dim dsBuy As New DataSet
    Dim b As Integer
    Dim grdtotal As Double
    Dim total As Integer
    Dim index As Integer

    Dim adCus As New OleDbDataAdapter
    Dim ds As New DataSet
    Dim m As Integer


   
    Private Sub Buying_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        btnRemove.Hide()
        btnSave.Hide()
        btnedit.Hide()

        con.Open()

        Dim cmitem As New OleDbCommand
        cmitem.Connection = con
        cmitem.CommandText = "SELECT * FROM tblItem"
        aditem.SelectCommand = cmitem
        aditem.Fill(dsitem, "Item")
        n = dsitem.Tables("Item").Rows.Count - 1
        con.Close()

        DataGridView2.ColumnCount = 4
        DataGridView2.Columns(0).Name = "Item Id"
        DataGridView2.Columns(1).Name = "Item Name"
        DataGridView2.Columns(2).Name = "Amount"
        DataGridView2.Columns(3).Name = "Price"

        DataGridView2.Columns(0).Width = 70
        DataGridView2.Columns(1).Width = 190
        DataGridView2.Columns(2).Width = 60
        DataGridView2.Columns(3).Width = 90
        
       
        Dim tbItem As DataTable
        Dim dcPrimaryKey(0) As DataColumn
        tbItem = dsitem.Tables("Item")
        dcPrimaryKey(0) = tbItem.Columns("iId")
        tbItem.PrimaryKey = dcPrimaryKey


        Dim drItem As DataRow

        If n >= 0 Then

            Dim i As Integer = 0

            While i <= n
                drItem = dsitem.Tables("Item").Rows(i)

                With drItem
                    DataGridView2.Rows.Add(.Item("iId"), .Item("iName"), .Item("iAmount"), .Item("iPrice"))
                    
                End With
                i = i + 1
            End While
        End If


        Try

            con.Open()

            Dim cmCus As New OleDbCommand
            cmCus.Connection = con
            cmCus.CommandText = "SELECT * FROM tblCustomer"

            adCus.SelectCommand = cmCus

            adCus.Fill(ds, "Customer")
            m = ds.Tables("Customer").Rows.Count - 1
            con.Close()

            Dim tbCustomer As DataTable
            Dim dcPrimaryKe(0) As DataColumn
            tbCustomer = ds.Tables("Customer")

            dcPrimaryKe(0) = tbCustomer.Columns("cID")
            tbCustomer.PrimaryKey = dcPrimaryKe


            Dim drCustomer As DataRow

            If m >= 0 Then

                Dim l As Integer = 0

                While l <= m
                    drCustomer = ds.Tables("Customer").Rows(l)

                    With drCustomer

                        cboCus.Items.Add(.Item("cID"))


                    End With
                    l = l + 1
                End While
            End If




            Dim drCus As DataRow

            If m >= 0 Then
                drCus = ds.Tables("Customer").Rows(m)

                With drCus
                    cboCus.Text = .Item("cID")
                End With
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try



    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub txtItemid_MouseDoubleClick(sender As Object, e As MouseEventArgs)



    End Sub

    Private Sub txtItemid_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridView2_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs)



    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnAddBill.Click
       
        DataGridView1.AllowUserToAddRows = False



        Dim count As Integer
       
        If txtQuantity.Text = "" Or IsNumeric(txtQuantity.Text) = False Then
            If txtQuantity.Text = "" Then
                MessageBox.Show("Please Enter the Quantity")

            ElseIf IsNumeric(txtQuantity.Text) = False Then
                MessageBox.Show("Please enter only only numbers")

            End If

        Else



            For nn As Integer = 0 To DataGridView2.Rows.Count - 1

                If txtItemid.Text = DataGridView2.Item(0, nn).Value Then

                    If Val(txtQuantity.Text) > Val(DataGridView2.Item(2, nn).Value) Then
                        MessageBox.Show("This Item is out of stock")

                    Else

                        Try

                            Dim cmBuild As New OleDbCommandBuilder
                            cmBuild.DataAdapter = aditem
                            Dim tbItem As DataTable
                            Dim dcPrimaryKey(0) As DataColumn
                            tbItem = dsitem.Tables("Item")
                            dcPrimaryKey(0) = tbItem.Columns("iId")
                            tbItem.PrimaryKey = dcPrimaryKey
                            Dim dritem As DataRow = tbItem.Rows.Find(txtItemid.Text)
                            With dritem

                                .Item("iAmount") = .Item("iAmount") - Val(txtQuantity.Text)

                            End With

                            aditem.UpdateCommand = cmBuild.GetUpdateCommand



                            For g As Integer = 0 To DataGridView1.Rows.Count - 1

                                If txtItemid.Text = DataGridView1.Rows(g).Cells(0).Value Then


                                    DataGridView1.Rows(g).Cells(2).Value = DataGridView1.Rows(g).Cells(2).Value + Val(txtQuantity.Text)


                                    total = Val(txtQuantity.Text) * Val(txtPrice.Text)
                                    DataGridView1.Rows(g).Cells(3).Value = DataGridView1.Rows(g).Cells(3).Value + total
                                    grdtotal = grdtotal + total
                                    txtgrd.Text = grdtotal


                                    con.Open()
                                    Try
                                        aditem.Update(dsitem, ("Item"))


                                    Catch ex As Exception
                                        MessageBox.Show("You are trying to save data incorrectly...")
                                    End Try
                                    con.Close()



                                    DataGridView2.Rows.Clear()

                                    con.Open()

                                    Dim cm As New OleDbCommand
                                    cm.Connection = con
                                    cm.CommandText = "SELECT * FROM tblItem"
                                    aditem.SelectCommand = cm
                                    aditem.Fill(dsitem, "Item")
                                    n = dsitem.Tables("Item").Rows.Count - 1
                                    con.Close()

                                    Dim drI As DataRow

                                    If n >= 0 Then

                                        Dim f As Integer = 0

                                        While f <= n
                                            drI = dsitem.Tables("Item").Rows(f)

                                            With drI
                                                DataGridView2.Rows.Add(.Item("iId"), .Item("iName"), .Item("iAmount"), .Item("iPrice"))

                                            End With
                                            f = f + 1
                                        End While
                                    End If

                                    Return

                                End If

                            Next



                            Dim rnum As Integer = DataGridView1.Rows.Add()
                            total = Val(txtQuantity.Text) * Val(txtPrice.Text)

                            DataGridView1.Rows.Item(rnum).Cells("Column1").Value = txtItemid.Text
                            DataGridView1.Rows.Item(rnum).Cells("Column2").Value = txtPrice.Text
                            DataGridView1.Rows.Item(rnum).Cells("Column3").Value = txtQuantity.Text
                            DataGridView1.Rows.Item(rnum).Cells("Column4").Value = total

                            grdtotal = grdtotal + total
                            txtgrd.Text = grdtotal


                        Catch ex As Exception
                            MsgBox(ex.Message)

                        End Try

                    End If

                    count = 1
                End If
            Next

            If count <> 1 Then
                MessageBox.Show("Item not found")
            End If

            con.Open()
            Try
                aditem.Update(dsitem, ("Item"))


            Catch ex As Exception
                MessageBox.Show("You are trying to save data incorrectly...")
            End Try
            con.Close()



        End If


        DataGridView2.Rows.Clear()

        con.Open()

        Dim cmitem As New OleDbCommand
        cmitem.Connection = con
        cmitem.CommandText = "SELECT * FROM tblItem"
        aditem.SelectCommand = cmitem
        aditem.Fill(dsitem, "Item")
        n = dsitem.Tables("Item").Rows.Count - 1
        con.Close()

        Dim drIte As DataRow

        If n >= 0 Then

            Dim i As Integer = 0

            While i <= n
                drIte = dsitem.Tables("Item").Rows(i)

                With drIte
                    DataGridView2.Rows.Add(.Item("iId"), .Item("iName"), .Item("iAmount"), .Item("iPrice"))

                End With
                i = i + 1
            End While
        End If



    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles txtgrd.TextChanged

    End Sub
    Sub Cancel()
        ' txtItemid.Text = ""
        txtItemName.Text = ""
        'cboCus.Text = ""
        txtPrice.Text = ""
        txtQuantity.Text = ""
        'txtgrd.Text = ""
        txtItemid.Clear()
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Cancel()

        DataGridView2.Rows.Clear()

        con.Open()

        Dim cmitem As New OleDbCommand
        cmitem.Connection = con
        cmitem.CommandText = "SELECT * FROM tblItem"
        aditem.SelectCommand = cmitem
        aditem.Fill(dsitem, "Item")
        n = dsitem.Tables("Item").Rows.Count - 1
        con.Close()

        Dim drIte As DataRow

        If n >= 0 Then

            Dim i As Integer = 0

            While i <= n
                drIte = dsitem.Tables("Item").Rows(i)

                With drIte
                    DataGridView2.Rows.Add(.Item("iId"), .Item("iName"), .Item("iAmount"), .Item("iPrice"))

                End With
                i = i + 1
            End While
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)
       
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub btnLast_Click(sender As Object, e As EventArgs)
       
    End Sub

    Private Sub btnFirst_Click(sender As Object, e As EventArgs)
      
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs)
      
    End Sub

    Private Sub btnPrevious_Click(sender As Object, e As EventArgs)
       
    End Sub

    Private Sub DataGridView2_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
       

    End Sub

    Private Sub DataGridView2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        Try
            Dim sele As DataGridViewRow
            sele = DataGridView2.Rows(e.RowIndex)
            txtItemid.Text = sele.Cells(0).Value
            txtItemName.Text = sele.Cells(1).Value
            txtPrice.Text = sele.Cells(3).Value
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

   

    Private Sub btnAddBill_Enter(sender As Object, e As EventArgs) Handles btnAddBill.Enter

      
    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs)
       

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Dim k As Integer
        k = MessageBox.Show("Are you sure want to save", " Saving", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If k = vbYes Then

            DateTimePicker1.Format = DateTimePickerFormat.Custom
            DateTimePicker1.CustomFormat = "MM/dd/yyyy"

            con.Open()

            Dim cmCus As New OleDbCommand
            cmCus.Connection = con
            cmCus.CommandText = "SELECT * FROM tblBuying"

            adBuy.SelectCommand = cmCus

            adBuy.Fill(dsBuy, "Buying")
            b = dsBuy.Tables("Buying").Rows.Count - 1


            con.Close()


            Dim cmBuilder As New OleDbCommandBuilder
            cmBuilder.DataAdapter = adBuy

            If cboCus.Text = "" Then
                MessageBox.Show("You didn't select any Customer")

            ElseIf DataGridView1.Rows.Count = 0 Then
                MessageBox.Show("There are no items to save")

            Else


                Try

                    For s As Integer = 0 To DataGridView1.Rows.Count - 1

                        Dim drCus As DataRow
                        drCus = dsBuy.Tables("Buying").NewRow

                        With drCus


                            .Item("cID") = cboCus.Text
                            .Item("iID") = DataGridView1.Item(0, s).Value
                            .Item("bDate") = DateTimePicker1.Text
                            .Item("bQty") = DataGridView1.Item(2, s).Value
                            .Item("Amount") = DataGridView1.Item(3, s).Value


                        End With
                        dsBuy.Tables("Buying").Rows.Add(drCus)
                        adBuy.InsertCommand = cmBuilder.GetInsertCommand
                        b = b + 1
                    Next
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                con.Open()
                Try

                    adBuy.Update(dsBuy, "Buying")
                    MessageBox.Show("Data saved successfully")

                Catch ex As Exception
                    MessageBox.Show("You are trying to save data incorrectly")

                End Try
                con.Close()

                'DataGridView1.Rows.Clear()
                Cancel()
            End If

        End If


    End Sub

    Private Sub txtDate_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub txtDate_ValueChanged(sender As Object, e As EventArgs)


    End Sub

    Private Sub Button1_Click_3(sender As Object, e As EventArgs) Handles btnRemove.Click

        If DataGridView1.Rows.Count <> 0 Then


            con.Open()
            Dim cmitem As New OleDbCommand
            cmitem.Connection = con
            cmitem.CommandText = "SELECT * FROM tblItem"
            aditem.SelectCommand = cmitem
            aditem.Fill(dsitem, "Item")
            n = dsitem.Tables("Item").Rows.Count - 1
            con.Close()

            Try
                Dim p As Integer = DataGridView1.CurrentRow.Index

                Dim cmBuild As New OleDbCommandBuilder
                cmBuild.DataAdapter = aditem
                Dim tbItem As DataTable
                Dim dcPrimaryKey(0) As DataColumn
                tbItem = dsitem.Tables("Item")
                dcPrimaryKey(0) = tbItem.Columns("iId")
                tbItem.PrimaryKey = dcPrimaryKey
                Dim dritem As DataRow = tbItem.Rows.Find(DataGridView1.Item(0, p).Value)
                With dritem
                    .Item("iAmount") = Val(DataGridView1.Item(2, p).Value) + .Item("iAmount")
                End With
                aditem.UpdateCommand = cmBuild.GetUpdateCommand
                grdtotal = grdtotal - DataGridView1.Item(3, p).Value
                txtgrd.Text = grdtotal
            Catch ex As Exception
                MsgBox(ex.Message)

            End Try

            Try

                Dim v As Integer = DataGridView1.CurrentRow.Index
                DataGridView1.Rows.RemoveAt(v)


            Catch ex As Exception

                MsgBox(ex.Message)

            End Try

            con.Open()

            Try
                aditem.Update(dsitem, ("Item"))

            Catch ex As Exception
                MessageBox.Show("You are trying to save data incorrectly...")
            End Try
            con.Close()



            DataGridView2.Rows.Clear()


            Dim drIte As DataRow

            If n >= 0 Then

                Dim i As Integer = 0

                While i <= n
                    drIte = dsitem.Tables("Item").Rows(i)

                    With drIte
                        DataGridView2.Rows.Add(.Item("iId"), .Item("iName"), .Item("iAmount"), .Item("iPrice"))

                    End With
                    i = i + 1
                End While
            End If

        Else
            MessageBox.Show("There is no rows to remove")
        End If

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick


        index = e.RowIndex

        con.Open()

        Dim cmitem As New OleDbCommand
        cmitem.Connection = con
        cmitem.CommandText = "SELECT * FROM tblItem"
        aditem.SelectCommand = cmitem
        aditem.Fill(dsitem, "Item")
        n = dsitem.Tables("Item").Rows.Count - 1
        con.Close()

        Try
            Dim p As Integer = DataGridView1.CurrentRow.Index

            Dim cmBuild As New OleDbCommandBuilder
            cmBuild.DataAdapter = aditem
            Dim tbItem As DataTable
            Dim dcPrimaryKey(0) As DataColumn
            tbItem = dsitem.Tables("Item")
            dcPrimaryKey(0) = tbItem.Columns("iId")
            tbItem.PrimaryKey = dcPrimaryKey
            Dim dritem As DataRow = tbItem.Rows.Find(DataGridView1.Item(0, p).Value)
            With dritem
                .Item("iAmount") = Val(DataGridView1.Item(2, p).Value) + .Item("iAmount")
            End With
            aditem.UpdateCommand = cmBuild.GetUpdateCommand

            grdtotal = grdtotal - DataGridView1.Item(3, p).Value
            txtgrd.Text = grdtotal


        Catch ex As Exception
            MsgBox(ex.Message)

        End Try

        Try
            Dim selectraw As DataGridViewRow
            selectraw = DataGridView1.Rows(index)
            DataGridView1.Rows.RemoveAt(index)
        Catch ex As Exception
            MsgBox(ex.Message)

        End Try

        con.Open()

        Try
            aditem.Update(dsitem, ("Item"))

        Catch ex As Exception
            MessageBox.Show("You are trying to save data incorrectly...")
        End Try
        con.Close()

        DataGridView2.Rows.Clear()


        Dim drIte As DataRow

        If n >= 0 Then

            Dim i As Integer = 0

            While i <= n
                drIte = dsitem.Tables("Item").Rows(i)

                With drIte
                    DataGridView2.Rows.Add(.Item("iId"), .Item("iName"), .Item("iAmount"), .Item("iPrice"))

                End With
                i = i + 1
            End While
        End If


    End Sub

    Private Sub Button1_Click_4(sender As Object, e As EventArgs) Handles Button1.Click
        Dim k As Integer
        k = MessageBox.Show("Are you sure want to exit", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If k = vbYes Then
            DataGridView1.Rows.Clear()
            frmMain.Show()
            Me.Hide()
        End If

    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim tbiTem As DataTable
        Dim dcPrimaryKey(0) As DataColumn
        tbiTem = dsitem.Tables("Item")
        dcPrimaryKey(0) = tbiTem.Columns("iId")
        tbiTem.PrimaryKey = dcPrimaryKey

        Dim strCusID As String

        strCusID = txtSearch.Text
        txtSearch.Clear()

        If Not strCusID Is Nothing Then
            Dim drItem As DataRow = tbiTem.Rows.Find(strCusID)
            DataGridView2.Rows.Clear()
            Dim num As Integer = DataGridView2.Rows.Add()
            If Not drItem Is Nothing Then
                With drItem

                    DataGridView2.Rows.Item(num).Cells("Item Id").Value = .Item("iId")
                    DataGridView2.Rows.Item(num).Cells("Item Name").Value = .Item("iName")
                    DataGridView2.Rows.Item(num).Cells("Amount").Value = .Item("iAmount")
                    DataGridView2.Rows.Item(num).Cells("Price").Value = .Item("iPrice")

                End With
              

            Else
                DataGridView2.Rows.Clear()



                con.Open()

                Dim cmitem As New OleDbCommand
                cmitem.Connection = con
                cmitem.CommandText = "SELECT * FROM tblItem"
                aditem.SelectCommand = cmitem
                aditem.Fill(dsitem, "Item")
                n = dsitem.Tables("Item").Rows.Count - 1
                con.Close()

                Dim drIte As DataRow

                If n >= 0 Then

                    Dim i As Integer = 0

                    While i <= n
                        drIte = dsitem.Tables("Item").Rows(i)

                        With drIte
                            DataGridView2.Rows.Add(.Item("iId"), .Item("iName"), .Item("iAmount"), .Item("iPrice"))

                        End With
                        i = i + 1
                    End While
                End If
                MessageBox.Show("Item Not Found")
            End If
        End If

    End Sub

    Private Sub btnshow_Click(sender As Object, e As EventArgs)

        DataGridView2.Rows.Clear()

        con.Open()

        Dim cmitem As New OleDbCommand
        cmitem.Connection = con
        cmitem.CommandText = "SELECT * FROM tblItem"
        aditem.SelectCommand = cmitem
        aditem.Fill(dsitem, "Item")
        n = dsitem.Tables("Item").Rows.Count - 1
        con.Close()

        DataGridView2.ColumnCount = 4
        DataGridView2.Columns(0).Name = "Item Id"
        DataGridView2.Columns(1).Name = "Item Name"
        DataGridView2.Columns(2).Name = "Amount"
        DataGridView2.Columns(3).Name = "Price"

        DataGridView2.Columns(0).Width = 70
        DataGridView2.Columns(1).Width = 190
        DataGridView2.Columns(2).Width = 60
        DataGridView2.Columns(3).Width = 90


        Dim tbItem As DataTable
        Dim dcPrimaryKey(0) As DataColumn
        tbItem = dsitem.Tables("Item")
        dcPrimaryKey(0) = tbItem.Columns("iId")
        tbItem.PrimaryKey = dcPrimaryKey


        Dim drItem As DataRow

        If n >= 0 Then

            Dim i As Integer = 0

            While i <= n
                drItem = dsitem.Tables("Item").Rows(i)

                With drItem
                    DataGridView2.Rows.Add(.Item("iId"), .Item("iName"), .Item("iAmount"), .Item("iPrice"))

                End With
                i = i + 1
            End While
        End If
    End Sub

    Private Sub cboCus_MouseClick(sender As Object, e As MouseEventArgs) Handles cboCus.MouseClick
     
    End Sub

    Private Sub cboCus_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles cboCus.MouseDoubleClick
       
    End Sub

    Private Sub cboCus_MouseHover(sender As Object, e As EventArgs) Handles cboCus.MouseHover

    End Sub

    Private Sub cboCus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCus.SelectedIndexChanged

    End Sub

    Private Sub txtCId_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "MM/dd/yyyy"
    End Sub

    Private Sub txtd_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnedit_Click(sender As Object, e As EventArgs) Handles btnedit.Click

        If DataGridView1.Rows.Count <> 0 Then


            Dim va As Integer
            Dim dd As Integer

            Dim cmitem As New OleDbCommand
            cmitem.Connection = con
            cmitem.CommandText = "SELECT * FROM tblItem"
            aditem.SelectCommand = cmitem
            aditem.Fill(dsitem, "Item")
            n = dsitem.Tables("Item").Rows.Count - 1
            con.Close()


            Try
                Dim se As Integer = DataGridView1.CurrentRow.Index

                Dim cmBuild As New OleDbCommandBuilder
                cmBuild.DataAdapter = aditem
                Dim tbItem As DataTable
                Dim dcPrimaryKey(0) As DataColumn
                tbItem = dsitem.Tables("Item")
                dcPrimaryKey(0) = tbItem.Columns("iId")
                tbItem.PrimaryKey = dcPrimaryKey
                Dim dritem As DataRow = tbItem.Rows.Find(DataGridView1.Item(0, se).Value)
                With dritem
                    dd = .Item("iAmount")

                End With

                va = InputBox("Enter the new Qty")

                If va > dd + DataGridView1.Item(2, se).Value Then

                    MessageBox.Show("Out of stock")

                ElseIf va <= dd + DataGridView1.Item(2, se).Value Then
                    Dim qq As Integer = DataGridView1.Item(2, se).Value
                    Dim rr As Double = DataGridView1.Item(3, se).Value
                    total = CDbl(va * DataGridView1.Item(1, se).Value)
                    DataGridView1.Item(3, se).Value = total
                    grdtotal = (grdtotal + total) - rr
                    DataGridView1.Item(2, se).Value = va
                    txtgrd.Text = grdtotal
                    With dritem
                        .Item("iAmount") = (dd + qq) - va

                    End With
                    aditem.UpdateCommand = cmBuild.GetUpdateCommand

                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            con.Open()
            Try
                aditem.Update(dsitem, ("Item"))


            Catch ex As Exception
                MessageBox.Show("You are trying to save data incorrectly...")
            End Try
            con.Close()


            DataGridView2.Rows.Clear()

            con.Open()

            Dim cm As New OleDbCommand
            cm.Connection = con
            cm.CommandText = "SELECT * FROM tblItem"
            aditem.SelectCommand = cm
            aditem.Fill(dsitem, "Item")
            n = dsitem.Tables("Item").Rows.Count - 1
            con.Close()

            Dim drI As DataRow

            If n >= 0 Then

                Dim f As Integer = 0

                While f <= n
                    drI = dsitem.Tables("Item").Rows(f)

                    With drI
                        DataGridView2.Rows.Add(.Item("iId"), .Item("iName"), .Item("iAmount"), .Item("iPrice"))

                    End With
                    f = f + 1
                End While
            End If
        Else
            MessageBox.Show("There is no rows to edit")
        End If
    End Sub

    Private Sub txtItemid_TextChanged_1(sender As Object, e As EventArgs) Handles txtItemid.TextChanged
        con.Open()

        Dim cm As New OleDbCommand
        cm.Connection = con
        cm.CommandText = "SELECT * FROM tblItem"
        aditem.SelectCommand = cm
        aditem.Fill(dsitem, "Item")
        n = dsitem.Tables("Item").Rows.Count - 1
        con.Close()

        For kk As Integer = 0 To DataGridView2.Rows.Count - 1

            If txtItemid.Text = DataGridView2.Item(0, kk).Value Then

                txtItemName.Text = DataGridView2.Item(1, kk).Value
                txtPrice.Text = DataGridView2.Item(3, kk).Value


            End If

        Next

    End Sub

    Private Sub DataGridView1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DataGridView1.RowsAdded

       
    End Sub

    Private Sub DataGridView1_RowStateChanged(sender As Object, e As DataGridViewRowStateChangedEventArgs) Handles DataGridView1.RowStateChanged
        If DataGridView1.Rows.Count = 0 Then
            btnRemove.Hide()
            btnedit.Hide()
            btnSave.Hide()
        ElseIf DataGridView1.Rows.Count > 0 Then
            btnRemove.Show()
            btnedit.Show()
            btnSave.Show()
        End If
    End Sub

    Private Sub txtQuantity_TextChanged(sender As Object, e As EventArgs) Handles txtQuantity.TextChanged
        If txtQuantity.Text <> "" Then
            If IsNumeric(txtQuantity.Text) = False Then
                MsgBox("Please enter numbers")
            End If
        End If
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        frmReceipt.Show()
    End Sub
End Class