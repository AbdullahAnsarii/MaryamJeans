Imports System.Data.OleDb

Public Class Form1
    Dim pro As String
    Dim connstring As String
    Dim command As String
    Dim myconnection As OleDbConnection = New OleDbConnection

    Dim ds As New OleDbDataAdapter
    Dim con As New OleDbConnection
    Dim dt As New DataTable

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        pro = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Abdullah\Documents\shopdb.accdb"
        connstring = pro
        myconnection.ConnectionString = connstring
        myconnection.Open()
        command = "insert into Table1 ([ID],[Item_Name],[Qty],[Price]) values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
        Dim cmd As OleDbCommand = New OleDbCommand(command, myconnection)
        cmd.Parameters.Add(New OleDbParameter("ID", CType(TextBox1.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Item_Name", CType(TextBox2.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Qty", CType(TextBox3.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Price", CType(TextBox4.Text, String)))
        MessageBox.Show("Item Saved!!!", "Notification")
        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myconnection.Close()
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        pro = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Abdullah\Documents\shopdb.accdb"
        connstring = pro
        myconnection.ConnectionString = connstring
        myconnection.Open()
        command = "Delete from [Table1] where [ID]=" & TextBox1.Text & ""
        Dim cmd As OleDbCommand = New OleDbCommand(command, myconnection)

        MessageBox.Show("Item Deleted!!!", "Notification")
        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myconnection.Close()
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        pro = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Abdullah\Documents\shopdb.accdb"
        connstring = pro
        myconnection.ConnectionString = connstring
        myconnection.Open()
        command = "update Table1 set [Item_Name] = '" & TextBox2.Text & "', [Qty] = '" & TextBox3.Text & "', [Price] = '" & TextBox4.Text & "' where [ID] = " & TextBox1.Text & ""
        Dim cmd As OleDbCommand = New OleDbCommand(command, myconnection)

        MessageBox.Show("Item Updated!!!", "Notification")
        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myconnection.Close()
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Abdullah\Documents\shopdb.accdb"
        con.Open()
        ds = New OleDbDataAdapter("Select * from Table1", con)
        ds.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
        con.Close()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
    End Sub
End Class
