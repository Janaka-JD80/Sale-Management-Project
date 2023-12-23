Imports System.Data.OleDb
Module Module1
    Public con As New OleDbConnection

    Sub Main()

        con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\New folder\nvq\Y Project\Com.mdb"
        Dim Items As New Items
        Items.ShowDialog()
    End Sub


End Module
